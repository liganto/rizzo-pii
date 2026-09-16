# -*- coding: utf-8 -*-
"""Anonimizzazione dei fogli Excel (src/app/xlsx_export.py).

Il modulo non dipende dal modello: le span che in produzione arrivano da
analyze() qui sono costruite a mano cercando i valori nel testo steso, cosi' il
test verifica il PERCORSO workbook -> testo -> workbook e non cosa predice il
modello oggi.

Serve openpyxl. La CI non installa dipendenze (vedi .github/workflows/tests.yml):
se manca, il test si salta invece di fallire.

Nomi, codici fiscali, email e telefoni sono SINTETICI.
"""

import base64
import io
import json
import sys
import unittest
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src" / "app"))

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.comments import Comment
    from openpyxl.worksheet.datavalidation import DataValidation

    import xlsx_export as X
except ImportError:                     # openpyxl assente (CI): si salta
    X = None


def _fixture():
    wb = Workbook()
    ws = wb.active
    ws.title = "Anagrafica"
    ws.append(["Cognome", "Nome", "Codice fiscale", "Telefono", "Nascita", "Reparto", "Matricola"])
    ws.append(["ROSSI", "GIACOMO", "RSSGCM80A01H501U", 3331234567, datetime(1980, 1, 1), "Paghe", 542])
    ws.append(["CORDELLA", "DE", "VRDPLA75B02F205X", 3479876543, datetime(1975, 2, 2), "Paghe", 773])
    ws.append(["ROSSI", "SIMONE", "RSSSMN90C03L219K", None, None, "Magazzino", 45])
    ws["F2"].comment = Comment("Referente: Giacomo Rossi", "Mario Bianchi")
    ws["H2"] = '="Sig. " & "ROSSI"'
    ws["I2"] = "mario.rossi@example.it"
    ws["I2"].hyperlink = "mailto:mario.rossi@example.it"
    ws["J2"] = "riga\tcon\nTAB ROSSI"
    ws["K2"] = "Mario"
    ws["L2"] = "Rossi"
    ws.oddHeader.center.text = "Elenco di ROSSI"

    hidden = wb.create_sheet("ROSSI privato")
    hidden.sheet_state = "hidden"
    hidden.append(["Nota", "ROSSI deve firmare"])

    wb.properties.creator = "Mario Bianchi"
    wb.properties.lastModifiedBy = "Mario Bianchi"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# cosa "trova il modello": solo la PRIMA occorrenza di ogni valore, di proposito,
# per verificare che le altre arrivino dalla propagazione del dizionario
MODEL_FINDS = {
    "ROSSI": "[FULLNAME_1]",
    "GIACOMO": "[FULLNAME_2]",
    "RSSGCM80A01H501U": "[CF_1]",
    "VRDPLA75B02F205X": "[CF_2]",
    "RSSSMN90C03L219K": "[CF_3]",
    "3331234567": "[PHONE_1]",
    "01/01/1980": "[DATE_1]",
    "mario.rossi@example.it": "[EMAIL_1]",
    "Mario Rossi": "[FULLNAME_3]",
    "45": "[BUILDINGNUM_1]",
    "Bianchi": "[FULLNAME_4]",
}


def _anonimizza(data=None, finds=MODEL_FINDS):
    """-> (bytes dell'xlsx anonimizzato, report, testo steso)"""
    lay = X.read_workbook(data or _fixture())
    text = lay.text
    spans, taken = [], []
    for val, ph in sorted(finds.items(), key=lambda kv: -len(kv[0])):
        i = text.find(val)
        while i >= 0 and any(a < i + len(val) and i < b for a, b in taken):
            i = text.find(val, i + 1)
        if i >= 0:
            spans.append((i, i + len(val), ph))
            taken.append((i, i + len(val)))
    mapping = {ph: val for val, ph in finds.items()}
    out, report = X.write_workbook(lay, spans, mapping)
    return out, report, text


def _run(data=None, finds=MODEL_FINDS):
    out, report, text = _anonimizza(data, finds)
    return load_workbook(io.BytesIO(out)), report, text


@unittest.skipIf(X is None, "serve openpyxl")
class TestoSteso(unittest.TestCase):

    def test_una_riga_per_riga_e_tab_tra_le_celle(self):
        text = X.read_workbook(_fixture()).text
        self.assertIn("ROSSI\tGIACOMO\tRSSGCM80A01H501U\t3331234567\t01/01/1980\tPaghe\t542", text)
        self.assertIn("Referente: Giacomo Rossi", text)       # commento
        self.assertIn("Elenco di ROSSI", text)                # intestazione
        self.assertIn("ROSSI privato", text)                  # nome del foglio nascosto

    def test_tab_e_acapo_dentro_una_cella_non_spostano_gli_offset(self):
        text = X.read_workbook(_fixture()).text
        self.assertIn("riga con TAB ROSSI", text)

    def test_file_non_excel(self):
        with self.assertRaises(X.XlsxError):
            X.read_workbook(b"%PDF-1.7 non sono un foglio")

    def test_segmenti_senza_testo(self):
        with self.assertRaises(X.XlsxError):
            X.spans_from_segments([{"t": "ciao "}, {"label": "FULLNAME", "ph": "[FULLNAME_1]"}])

    def test_segmenti_ad_offset(self):
        segs = [{"t": "Sono "}, {"label": "FULLNAME", "ph": "[FULLNAME_1]", "t": "Mario"},
                {"t": "\t"}, {"label": "CF", "ph": "[CF_1]", "t": "RSSGCM80A01H501U"}]
        self.assertEqual(X.spans_from_segments(segs),
                         [(5, 10, "[FULLNAME_1]"), (11, 27, "[CF_1]")])


@unittest.skipIf(X is None, "serve openpyxl")
class Workbook_Anonimizzato(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.wb, cls.report, cls.text = _run()
        cls.ws = cls.wb["Anagrafica"]

    def test_celle_pii_sostituite_anche_dove_il_modello_non_le_ha_viste(self):
        ws = self.ws
        self.assertEqual(ws["A2"].value, "[FULLNAME_1]")
        self.assertEqual(ws["A4"].value, "[FULLNAME_1]")     # propagazione
        self.assertEqual(ws["B2"].value, "[FULLNAME_2]")
        self.assertEqual(ws["C3"].value, "[CF_2]")

    def test_celle_non_pii_intatte_e_col_loro_tipo(self):
        ws = self.ws
        self.assertEqual(ws["F2"].value, "Paghe")
        self.assertEqual(ws["G2"].value, 542)
        self.assertIsInstance(ws["E3"].value, datetime)       # data non toccata resta data
        self.assertEqual(ws["D3"].value, 3479876543)          # telefono NON trovato: resta numero
        self.assertEqual(ws["A1"].value, "Cognome")

    def test_numero_e_data_con_pii_diventano_testo(self):
        self.assertEqual(self.ws["D2"].value, "[PHONE_1]")
        self.assertEqual(self.ws["E2"].value, "[DATE_1]")

    def test_confini_di_parola(self):
        """"DE" non e' nel dizionario, ma nemmeno "ROSSI" deve toccare "CORDELLA"."""
        self.assertEqual(self.ws["A3"].value, "CORDELLA")
        self.assertEqual(self.ws["B3"].value, "DE")

    def test_un_valore_non_scavalca_le_celle(self):
        """"Mario Rossi" e' nel dizionario: K2="Mario" e L2="Rossi" sono due celle."""
        self.assertEqual(self.ws["K2"].value, "Mario")
        self.assertEqual(self.ws["L2"].value, "[FULLNAME_1]")  # "Rossi" == "ROSSI"

    def test_formula_resta_formula(self):
        self.assertEqual(self.ws["H2"].value, '="Sig. " & "[FULLNAME_1]"')

    def test_cella_con_tab_interni(self):
        self.assertEqual(self.ws["J2"].value, "riga\tcon\nTAB [FULLNAME_1]")

    def test_commento_e_autore(self):
        c = self.ws["F2"].comment
        # "Giacomo" arriva dalla propagazione: il modello aveva visto "GIACOMO" in B2
        self.assertEqual(c.text, "Referente: [FULLNAME_2] [FULLNAME_1]")
        self.assertEqual(c.author, X.COMMENT_AUTHOR)

    def test_intestazione(self):
        self.assertEqual(self.ws.oddHeader.center.text, "Elenco di [FULLNAME_1]")

    def test_collegamento_rimosso(self):
        self.assertEqual(self.ws["I2"].value, "[EMAIL_1]")
        self.assertIsNone(self.ws["I2"].hyperlink)
        self.assertEqual(self.report["unlinked"], 1)

    def test_nome_del_foglio_nascosto(self):
        self.assertIn("FULLNAME_1 privato", self.wb.sheetnames)
        ws = self.wb["FULLNAME_1 privato"]
        self.assertEqual(ws.sheet_state, "hidden")
        self.assertEqual(ws["B1"].value, "[FULLNAME_1] deve firmare")
        self.assertEqual(self.report["renamed_sheets"], 1)

    def test_proprieta_del_documento(self):
        # openpyxl rimette creator="openpyxl" su un campo vuoto: conta che l'autore non ci sia
        for prop in ("creator", "lastModifiedBy"):
            self.assertNotIn("Bianchi", getattr(self.wb.properties, prop) or "")

    def test_valori_corti_segnalati_e_non_propagati(self):
        self.assertIn("[BUILDINGNUM_1]", self.report["skipped"])

    def test_nessun_residuo(self):
        self.assertEqual(self.report["residual"], [])
        for val in ("ROSSI", "GIACOMO", "RSSGCM80A01H501U", "mario.rossi@example.it", "Bianchi"):
            for ws in self.wb.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        with self.subTest(val=val, cell=c.coordinate):
                            self.assertNotIn(val.casefold(), str(c.value or "").casefold())


@unittest.skipIf(X is None, "serve openpyxl")
class Residui(unittest.TestCase):

    def test_pii_in_un_posto_non_riscritto_viene_segnalata(self):
        """Le convalide dati non vengono riscritte: la verifica finale deve dirlo."""
        wb = load_workbook(io.BytesIO(_fixture()))
        dv = DataValidation(type="list", formula1='"GIACOMO,SIMONE"')
        wb["Anagrafica"].add_data_validation(dv)
        dv.add("B10")
        buf = io.BytesIO()
        wb.save(buf)
        _, report, _ = _run(buf.getvalue())
        self.assertIn("[FULLNAME_2]", report["residual"])


@unittest.skipIf(X is None, "serve openpyxl")
class Ripristino(unittest.TestCase):
    """Andata e ritorno: anonimizza -> ripristina -> confronta con l'originale."""

    MAPPING = {ph: val for val, ph in MODEL_FINDS.items()}

    @classmethod
    def setUpClass(cls):
        cls.orig = load_workbook(io.BytesIO(_fixture()))
        cls.anon, cls.anon_report, _ = _anonimizza()
        out, cls.report = X.restore_workbook(cls.anon, cls.MAPPING, cls.anon_report["types"])
        cls.wb = load_workbook(io.BytesIO(out))

    def test_tipi_registrati_solo_per_celle_intere(self):
        types = self.anon_report["types"]
        self.assertEqual(types.get("[PHONE_1]"), "int")
        self.assertEqual(types.get("[DATE_1]"), "datetime")
        self.assertNotIn("[FULLNAME_1]", types)

    def test_ogni_cella_torna_come_prima(self):
        """A meno della grafia: il dizionario tiene una forma sola ("Rossi" -> "ROSSI")."""
        self.assertEqual(self.wb.sheetnames, self.orig.sheetnames)
        for wo, wr in zip(self.orig.worksheets, self.wb.worksheets):
            for row in wo.iter_rows():
                for co in row:
                    cr = wr[co.coordinate]
                    with self.subTest(foglio=wo.title, cella=co.coordinate):
                        if isinstance(co.value, str):
                            self.assertEqual(co.value.casefold(), str(cr.value).casefold())
                        else:
                            self.assertEqual(co.value, cr.value)     # stesso valore E tipo
                        if co.comment is not None:
                            self.assertEqual(co.comment.text.casefold(),
                                             cr.comment.text.casefold())
        self.assertEqual(self.wb["Anagrafica"].oddHeader.center.text, "Elenco di ROSSI")

    def test_formula_ripristinata(self):
        self.assertEqual(self.wb["Anagrafica"]["H2"].value, '="Sig. " & "ROSSI"')

    def test_file_modificato_da_un_llm(self):
        wb = load_workbook(io.BytesIO(self.anon))
        ws = wb["Anagrafica"]
        ws.append(["[ FULLNAME_1 ]", "vedi FULLNAME_2", "[FULLNAME_99]", "ART_2 del contratto"])
        ws["Z1"] = "Esito per [CF_1]"
        buf = io.BytesIO()
        wb.save(buf)
        out, report = X.restore_workbook(buf.getvalue(), self.MAPPING)
        ws = load_workbook(io.BytesIO(out))["Anagrafica"]
        self.assertEqual(ws["A5"].value, "ROSSI")
        self.assertEqual(ws["B5"].value, "vedi GIACOMO")
        self.assertEqual(ws["C5"].value, "[FULLNAME_99]")         # sconosciuto: resta
        self.assertEqual(ws["D5"].value, "ART_2 del contratto")   # non e' nel dizionario
        self.assertEqual(ws["Z1"].value, "Esito per RSSGCM80A01H501U")
        self.assertEqual(report["unknown"], ["[FULLNAME_99]"])

    def test_senza_tipi_resta_testo(self):
        out, report = X.restore_workbook(self.anon, self.MAPPING)
        ws = load_workbook(io.BytesIO(out))["Anagrafica"]
        self.assertEqual(ws["D2"].value, "3331234567")
        self.assertEqual(report["typed"], 0)

    def test_parse_dictionary(self):
        mapping, types = X.parse_dictionary(
            {"[A_1]": "x", "_types": {"[A_1]": "int"}, "commento": "ignorato", "[B_2]": 3})
        self.assertEqual(mapping, {"[A_1]": "x"})
        self.assertEqual(types, {"[A_1]": "int"})
        for bad in ([], {}, {"_types": {}}):
            with self.assertRaises(X.XlsxError):
                X.parse_dictionary(bad)


def _carica_app():
    """Importa app col pipeline HF finto (come test_tsv_anonymize): i pesi non servono,
    ma torch/transformers/flask/fitz si'."""
    if "app" in sys.modules:
        return sys.modules["app"]
    from unittest.mock import MagicMock, patch
    with patch("transformers.pipeline", return_value=MagicMock(return_value=[])):
        import app as pii_app
    return pii_app


try:
    APP = _carica_app() if X is not None else None
except (Exception, SystemExit):         # dipendenze o modello assenti: si salta
    APP = None


@unittest.skipIf(APP is None, "serve l'ambiente dell'app (torch/transformers/flask/fitz)")
class Endpoint(unittest.TestCase):
    """POST /xlsx end-to-end, con il modello sostituito da entita' iniettate a mano."""

    def _finto_modello(self, text):
        ents = []
        for val, ph in MODEL_FINDS.items():
            i = text.find(val)
            if i >= 0:
                ents.append({"label": ph[1:ph.rindex("_")], "start": i, "end": i + len(val),
                             "score": 0.99, "validated": False, "source": "modello"})
        return ents, 1

    def _post(self, url, filename, data, fields=None):
        from unittest.mock import patch
        with patch.object(APP, "detect_model", side_effect=self._finto_modello):
            return APP.app.test_client().post(
                url, data={"file": (io.BytesIO(data), filename), **(fields or {})},
                content_type="multipart/form-data")

    def test_dizionario_e_ripristino_via_api(self):
        resp = self._post("/xlsx", "dipendenti.xlsx", _fixture(), {"include_mapping": "1"})
        self.assertEqual(resp.status_code, 200)
        d = resp.get_json()
        self.assertEqual(d["filename"], "dipendenti_anonimizzato.xlsx")
        self.assertIn("[FULLNAME_1]", d["mapping"])
        anon = base64.b64decode(d["xlsx_base64"])
        dizionario = json.dumps({**d["mapping"], "_types": d["types"]})

        resp = self._post("/xlsx/restore", d["filename"], anon, {"mapping": dizionario})
        self.assertEqual(resp.status_code, 200, resp.data[:300])
        self.assertIn("dipendenti_ripristinato.xlsx", resp.headers["Content-Disposition"])
        ws = load_workbook(io.BytesIO(resp.data))["Anagrafica"]
        self.assertEqual(ws["A2"].value.casefold(), "rossi")
        self.assertEqual(ws["D2"].value, 3331234567)

    def test_ripristino_senza_dizionario(self):
        resp = self._post("/xlsx/restore", "dipendenti.xlsx", _fixture())
        self.assertEqual(resp.status_code, 400)
        resp = self._post("/xlsx/restore", "dipendenti.xlsx", _fixture(), {"mapping": "{rotto"})
        self.assertEqual(resp.status_code, 400)

    def test_scarica_excel_anonimizzato(self):
        resp = self._post("/xlsx", "dipendenti.xlsx", _fixture())
        self.assertEqual(resp.status_code, 200, resp.data[:300])   # bytes: e' un binario
        self.assertEqual(resp.mimetype, APP.XLSX_MIME)
        self.assertIn("dipendenti_anonimizzato.xlsx", resp.headers["Content-Disposition"])
        self.assertEqual(resp.headers["X-PII-Residual"], "0")
        wb = load_workbook(io.BytesIO(resp.data))
        ws = wb["Anagrafica"]
        self.assertEqual(ws["F2"].value, "Paghe")
        for row in ws.iter_rows():
            for c in row:
                self.assertNotIn("rossi", str(c.value or "").casefold(), c.coordinate)

    def test_analyze_accetta_xlsx(self):
        resp = self._post("/analyze", "dipendenti.xlsx", _fixture())
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Paghe", resp.get_json()["anonymized_text"])

    def test_xls_rifiutato_con_istruzioni(self):
        resp = self._post("/analyze", "vecchio.xls", b"\xd0\xcf\x11\xe0")
        self.assertEqual(resp.status_code, 400)
        self.assertIn(".xlsx", resp.get_json()["error"])

    def test_xlsx_non_valido(self):
        resp = self._post("/xlsx", "rotto.xlsx", b"non sono uno zip")
        self.assertEqual(resp.status_code, 400)


if __name__ == "__main__":
    unittest.main()
