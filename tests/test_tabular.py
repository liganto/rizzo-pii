# -*- coding: utf-8 -*-
"""Anonimizzazione per colonna di un'estrazione tabellare (src/app/tabular.py).

Il modulo non usa il modello: qui gira tutto senza pesi. Serve openpyxl; se manca
(la CI non installa dipendenze) i test si saltano.

I codici fiscali e le partite IVA del fixture sono SINTETICI: le cifre di controllo
vengono calcolate qui con i validatori del prodotto, cosi' il test verifica davvero
il riconoscimento per checksum senza incollare dati di nessuno.
"""

import io
import json
import string
import sys
import unittest
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src" / "app"))

try:
    from openpyxl import Workbook, load_workbook

    import detectors as D
    import tabular as T
    import xlsx_export as X
except ImportError:                     # openpyxl assente (CI): si salta
    T = None


def _cf(base15):
    """Completa un CF sintetico con il carattere di controllo giusto."""
    return next(base15 + ch for ch in string.ascii_uppercase if D.cf_ok(base15 + ch))


def _piva(base10):
    return next(base10 + ch for ch in string.digits if D.piva_ok(base10 + ch))


def _fixture(righe=6):
    wb = Workbook()
    ws = wb.active
    ws.title = "Estrazione"
    ws.append(["Albo", "ds_cognome", "ds_nome", "cd_fiscale", "Partita iva",
               "dt_nascita", "Matricola", "Reparto", "RAL", "Note"])
    for i in range(righe):
        ws.append([
            ("MILANO", "ROMA")[i % 2],                # poche categorie = quasi-identificatore
            "ROSSI%d" % i + " " * 10,                 # campo a lunghezza fissa
            "MARIO%d" % i,
            _cf("RSSMRA85H12F20%d" % (i % 10)),        # 15 caratteri + controllo
            _piva("123456789%d" % (i % 10)),
            datetime(1980 + i, 3, 14),
            10000 + i,
            "Paghe",
            30000 + i * 1000,
            "nessuna nota",
        ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _col(prof, header):
    return next(c for c in prof["columns"] if c["header"] == header)


@unittest.skipIf(T is None, "serve openpyxl")
class Profilo(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.prof = T.profile(_fixture())

    def test_intestazioni_e_righe(self):
        self.assertEqual(self.prof["sheet"], "Estrazione")
        self.assertEqual(self.prof["rows"], 6)
        self.assertEqual(len(self.prof["columns"]), 10)

    def test_riconosce_dai_valori_col_checksum(self):
        cf = _col(self.prof, "cd_fiscale")
        self.assertEqual(cf["tag"], "CF")
        self.assertEqual(cf["suggested_mode"], "mask")
        self.assertIn("valori", cf["reason"])
        piva = _col(self.prof, "Partita iva")
        self.assertEqual(piva["tag"], "PIVA")

    def test_riconosce_dal_nome_della_colonna(self):
        for header, tag in (("ds_cognome", "FULLNAME"), ("ds_nome", "FULLNAME"),
                            ("dt_nascita", "DATE"), ("Matricola", "ID_DOC")):
            c = _col(self.prof, header)
            with self.subTest(colonna=header):
                self.assertEqual(c["tag"], tag)
                self.assertEqual(c["suggested_mode"], "mask")

    def test_nomi_che_non_sono_persone(self):
        """Regressione (file reale): 'Table Name' non e' un nome di persona e
        'dt_prima_iscrizione' non e' un documento. Il match e' per token, non per
        sottostringa, altrimenti si mascherano colonne tecniche e date di evento."""
        for header in ("Table Name", "Nome file", "dt_prima_iscrizione",
                       "Data presentazione dich", "Anno di competenza", "sconto",
                       "regime_fiscale", "anno fiscale", "IVA dovuta", "aliquota IVA"):
            with self.subTest(colonna=header):
                self.assertIsNone(T._tag_from_name(header))
        # ma le colonne vere continuano a essere riconosciute
        for header, tag in (("cd_fisc_soc", "CF"), ("cd_iva_stp", "PIVA"),
                            ("ds_cognome", "FULLNAME"), ("dt_nascita", "DATE"),
                            ("Matricola", "ID_DOC"), ("Codice Fiscale", "CF"),
                            ("dtNascita", "DATE")):
            with self.subTest(colonna=header):
                self.assertEqual(T._tag_from_name(header), tag)

    def test_le_misure_restano_fuori(self):
        for header in ("RAL", "Reparto", "Note"):
            c = _col(self.prof, header)
            with self.subTest(colonna=header):
                self.assertEqual(c["suggested_mode"], "keep")
                self.assertIsNone(c["tag"])

    def test_forma_e_niente_valori(self):
        cf = _col(self.prof, "cd_fiscale")
        self.assertEqual(cf["shape"], "AAAAAA99A99A999A")
        piatto = json.dumps(self.prof)
        self.assertNotIn("MILANO", piatto)          # nel profilo non esce nessun valore
        self.assertNotIn("Paghe", piatto)

    def test_quasi_identificatori_segnalati(self):
        self.assertTrue(_col(self.prof, "Albo")["quasi_identifier"])
        self.assertFalse(_col(self.prof, "cd_fiscale")["quasi_identifier"])

    def test_foglio_inesistente(self):
        with self.assertRaises(X.XlsxError):
            T.profile(_fixture(), sheet="Manca")

    def test_file_non_excel(self):
        with self.assertRaises(X.XlsxError):
            T.profile(b"non sono un foglio")


@unittest.skipIf(T is None, "serve openpyxl")
class Anonimizzazione(unittest.TestCase):

    SCELTE = {"cd_fiscale": "mask", "ds_cognome": "mask", "ds_nome": "mask",
              "Partita iva": "mask", "Matricola": "mask", "dt_nascita": "mask",
              "RAL": "keep"}

    @classmethod
    def setUpClass(cls):
        cls.data = _fixture()
        cls.out, cls.report = T.anonymize(cls.data, cls.SCELTE)
        cls.ws = load_workbook(io.BytesIO(cls.out))["Estrazione"]

    def test_colonne_mascherate(self):
        for col in ("D", "B", "C", "E", "G", "F"):
            v = self.ws["%s2" % col].value
            with self.subTest(colonna=col):
                self.assertTrue(str(v).startswith("["), v)

    def test_prefissi_dal_tag_riconosciuto(self):
        self.assertTrue(self.ws["D2"].value.startswith("[CF_"))
        self.assertTrue(self.ws["E2"].value.startswith("[PIVA_"))
        # nome/cognome hanno lo stesso tag ma restano colonne distinte
        self.assertTrue(self.ws["B2"].value.startswith("[FULLNAME_"))
        self.assertNotEqual(self.ws["B2"].value, self.ws["C2"].value)

    def test_misure_e_colonne_non_nominate_intatte(self):
        self.assertEqual(self.ws["I2"].value, 30000)       # RAL: numero, non testo
        self.assertEqual(self.ws["H2"].value, "Paghe")
        self.assertEqual(self.ws["A2"].value, "MILANO")    # non nominata = keep
        self.assertEqual(self.ws["J2"].value, "nessuna nota")

    def test_report(self):
        self.assertEqual(self.report["left"], 0)
        self.assertEqual(self.report["ignored"], [])
        cf = next(c for c in self.report["columns"] if c["header"] == "cd_fiscale")
        self.assertEqual((cf["mode"], cf["cells"], cf["distinct"], cf["prefix"]),
                         ("mask", 6, 6, "CF"))
        self.assertEqual(len(self.report["mapping"]), 36)   # 6 colonne x 6 righe
        self.assertEqual(self.report["types"]["[ID_DOC_1]"], "int")
        self.assertEqual(self.report["types"]["[DATE_1]"], "datetime")

    def test_ritorno_identico_all_originale(self):
        """Il dizionario e' quello di xlsx_export: il ripristino chiude il giro."""
        rip, rep = X.restore_workbook(self.out, self.report["mapping"], self.report["types"])
        self.assertEqual(rep["unknown"], [])
        a = load_workbook(io.BytesIO(self.data))["Estrazione"]
        b = load_workbook(io.BytesIO(rip))["Estrazione"]
        for row in a.iter_rows():
            for ca in row:
                cb = b[ca.coordinate]
                with self.subTest(cella=ca.coordinate):
                    self.assertEqual(ca.value, cb.value)        # spazi di riempimento inclusi
                    self.assertEqual(type(ca.value), type(cb.value))

    def test_modalita_year_ed_empty(self):
        out, report = T.anonymize(_fixture(), {"dt_nascita": "year", "ds_cognome": "empty"})
        ws = load_workbook(io.BytesIO(out))["Estrazione"]
        self.assertEqual(ws["F2"].value, 1980)                  # solo l'anno, come numero
        self.assertIsNone(ws["B2"].value)
        self.assertEqual(report["mapping"], {})                 # nessuna delle due e' reversibile
        self.assertEqual(report["left"], 0)

    def test_colonne_per_indice_e_sconosciute(self):
        out, report = T.anonymize(_fixture(), {4: "mask", "colonna_che_non_esiste": "mask"})
        self.assertEqual(report["ignored"], ["colonna_che_non_esiste"])
        ws = load_workbook(io.BytesIO(out))["Estrazione"]
        self.assertTrue(ws["D2"].value.startswith("[CF_"))

    def test_celle_di_soli_spazi_non_diventano_placeholder(self):
        wb = load_workbook(io.BytesIO(_fixture()))
        wb["Estrazione"]["E2"] = "   "
        buf = io.BytesIO()
        wb.save(buf)
        out, report = T.anonymize(buf.getvalue(), {"Partita iva": "mask"})
        ws = load_workbook(io.BytesIO(out))["Estrazione"]
        self.assertEqual(ws["E2"].value, "   ")
        self.assertEqual(len(report["mapping"]), 5)

    def test_modalita_sconosciuta(self):
        with self.assertRaises(X.XlsxError):
            T.anonymize(_fixture(), {"cd_fiscale": "cancella"})

    def test_prefisso_sempre_valido_per_il_ripristino(self):
        """Un'intestazione con emoji/parentesi deve dare un placeholder ripristinabile."""
        wb = load_workbook(io.BytesIO(_fixture()))
        wb["Estrazione"]["J1"] = "Società 👥 (Affari)"
        buf = io.BytesIO()
        wb.save(buf)
        out, report = T.anonymize(buf.getvalue(), {"Società 👥 (Affari)": "mask"})
        ph = load_workbook(io.BytesIO(out))["Estrazione"]["J2"].value
        self.assertRegex(ph, r"^\[[A-Z][A-Z0-9_]*_\d+\]$")
        self.assertIsNotNone(X._PH_RE.fullmatch(ph))
        rip, _ = X.restore_workbook(out, report["mapping"], report["types"])
        self.assertEqual(load_workbook(io.BytesIO(rip))["Estrazione"]["J2"].value, "nessuna nota")


def _carica_app():
    if "app" in sys.modules:
        return sys.modules["app"]
    from unittest.mock import MagicMock, patch
    with patch("transformers.pipeline", return_value=MagicMock(return_value=[])):
        import app as pii_app
    return pii_app


try:
    APP = _carica_app() if T is not None else None
except (Exception, SystemExit):         # dipendenze o modello assenti: si salta
    APP = None


@unittest.skipIf(APP is None, "serve l'ambiente dell'app (torch/transformers/flask/fitz)")
class Endpoint(unittest.TestCase):
    """/xlsx/columns e /xlsx/columns/anonymize: nessuna inferenza, modello mai chiamato."""

    def _post(self, url, data, filename="estrazione.xlsx", fields=None):
        return APP.app.test_client().post(
            url, data={"file": (io.BytesIO(data), filename), **(fields or {})},
            content_type="multipart/form-data")

    def test_profilo(self):
        r = self._post("/xlsx/columns", _fixture())
        self.assertEqual(r.status_code, 200, r.data[:300])
        d = r.get_json()
        self.assertEqual(d["sheet"], "Estrazione")
        self.assertEqual(d["rows"], 6)
        cf = next(c for c in d["columns"] if c["header"] == "cd_fiscale")
        self.assertEqual((cf["tag"], cf["suggested_mode"]), ("CF", "mask"))
        self.assertNotIn("MILANO", r.get_data(as_text=True))

    def test_anonimizza_binario(self):
        r = self._post("/xlsx/columns/anonymize", _fixture(),
                       fields={"columns": json.dumps({"cd_fiscale": "mask"})})
        self.assertEqual(r.status_code, 200, r.data[:300])
        self.assertEqual(r.mimetype, APP.XLSX_MIME)
        self.assertIn("estrazione_anonimizzato.xlsx", r.headers["Content-Disposition"])
        self.assertEqual(r.headers["X-PII-Left"], "0")
        ws = load_workbook(io.BytesIO(r.data))["Estrazione"]
        self.assertTrue(ws["D2"].value.startswith("[CF_"))

    def test_anonimizza_con_dizionario_e_ripristino(self):
        r = self._post("/xlsx/columns/anonymize", _fixture(),
                       fields={"columns": json.dumps({"cd_fiscale": "mask", "ds_nome": "mask"}),
                               "include_mapping": "1"})
        self.assertEqual(r.status_code, 200, r.data[:300])
        d = r.get_json()
        self.assertEqual(len(d["mapping"]), 12)
        import base64
        anon = base64.b64decode(d["xlsx_base64"])
        r2 = self._post("/xlsx/restore", anon, d["filename"],
                        {"mapping": json.dumps({**d["mapping"], "_types": d["types"]})})
        self.assertEqual(r2.status_code, 200, r2.data[:300])
        ws = load_workbook(io.BytesIO(r2.data))["Estrazione"]
        self.assertEqual(ws["C2"].value, "MARIO0")

    def test_columns_mancante_o_rotto(self):
        self.assertEqual(self._post("/xlsx/columns/anonymize", _fixture()).status_code, 400)
        r = self._post("/xlsx/columns/anonymize", _fixture(), fields={"columns": "{rotto"})
        self.assertEqual(r.status_code, 400)
        r = self._post("/xlsx/columns/anonymize", _fixture(),
                       fields={"columns": json.dumps({"cd_fiscale": "cancella"})})
        self.assertEqual(r.status_code, 400)

    def test_solo_xlsx(self):
        self.assertEqual(self._post("/xlsx/columns", b"%PDF-", "atto.pdf").status_code, 400)


if __name__ == "__main__":
    unittest.main()
