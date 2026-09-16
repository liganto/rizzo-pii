# -*- coding: utf-8 -*-
"""
Anonimizzazione PER COLONNA di un'estrazione tabellare (.xlsx). Niente modello.

Perche' esiste, accanto a `xlsx_export`: il modello e' addestrato su PROSA (atti,
contratti, sentenze). Su un'estrazione gestionale - decine di colonne, ogni cella
un numero senza contesto - un codice a cinque cifre somiglia a un documento, a un
importo o a un civico, e il risultato e' un file mascherato ovunque e inutilizzabile
per l'analisi. In una tabella pero' l'informazione che conta e' gia' nello schema:
la colonna DICE cos'e' il valore. Quindi si decide per colonna, una volta, invece
di indovinare cella per cella: nessun falso positivo sulle misure, secondi invece
di ore, e il conto non dipende dal numero di righe.

Due passi:

  prof = profile(xlsx_bytes)
      Per ogni colonna: tipo prevalente, valori distinti, celle vuote, FORMA
      prevalente (lettere -> A, cifre -> 9) e una PROPOSTA di trattamento, decisa
      da due fonti:
        1. il NOME dell'intestazione ("cd_fiscale", "cognome", "dt_nascita"...);
        2. i VALORI, passati alla rete regex+checksum di `detectors` su un campione:
           se quasi tutte le celle sono un CF/IBAN/PIVA/email/carta valida, la
           colonna e' identificativa anche se si chiama `avg(0)`.
      Il profilo NON contiene valori reali, solo forme: si puo' mostrare a video
      o loggare senza esporre il contenuto del file.

  out, report = anonymize(xlsx_bytes, {"cd_fiscale": "mask", "dt_nascita": "year"})
      Modalita' per colonna:
        "mask"  -> placeholder coerente e REVERSIBILE ([CF_1], [COGNOME_7]): stesso
                   valore = stesso placeholder in tutta la colonna;
        "year"  -> di una data resta il solo ANNO (non reversibile): l'analisi per
                   eta' o per anzianita' sopravvive, la data esatta no;
        "empty" -> cella svuotata (non reversibile);
        "keep"  -> intatta (default per tutto cio' che non viene nominato).

Il dizionario prodotto e' quello di `xlsx_export` ({placeholder: valore} + "_types"),
quindi il ripristino e' gia' fatto: `restore_workbook()` / `POST /xlsx/restore`.
I valori vanno nel dizionario COSI' COME SONO, riempimento di spazi compreso (queste
estrazioni hanno campi a lunghezza fissa): il ripristino torna identico all'originale.
La deduplica invece ignora spazi e maiuscole, cosi' "ROSSI" e "Rossi " sono lo stesso
soggetto.

Limite dichiarato: si lavora su UN foglio per volta (`sheet`), con le intestazioni
su una riga (`header_row`). Un file con piu' tabelle si tratta una tabella alla volta.

ATTENZIONE, e' pseudonimizzazione: con il dizionario si risale a tutto, e le colonne
lasciate in chiaro restano dei QUASI-IDENTIFICATORI (provincia + sesso + anzianita' +
importi possono bastare a riconoscere un caso raro). `profile()` segnala le colonne
a bassa cardinalita' proprio per questo.
"""

import io
import re
import unicodedata
from datetime import date, datetime, time

from openpyxl import load_workbook

from detectors import detect_regex
from xlsx_export import XlsxError, cell_text

MODES = ("mask", "keep", "year", "empty")
SAMPLE = 200          # celle per colonna date in pasto alla rete regex+checksum
_MAX_DISTINCT = 50000  # oltre, si smette di contare (file enormi)

# Intestazione -> (tag, e' identificativa?). Confronto su nome normalizzato
# (minuscole, senza accenti, senza separatori): "cd_fiscale", "Codice Fiscale",
# "CODICEFISCALE" sono la stessa cosa.
_BY_NAME = [
    # NB: niente "fiscale" e niente "iva" da soli: prenderebbero "regime_fiscale",
    # "anno fiscale", "IVA dovuta", "aliquota IVA" - misure, non identificativi.
    (("codicefiscale", "cdfiscale", "codfiscale", "cf", "codfisc", "cdfisc"), "CF"),
    (("partitaiva", "piva", "cdiva", "codiva", "vat"), "PIVA"),
    (("iban", "conto", "coordinatebancarie"), "IBAN"),
    (("email", "mail", "pec", "postaelettronica"), "EMAIL"),
    (("telefono", "cellulare", "tel", "cell", "phone", "mobile", "fax"), "TELEPHONENUM"),
    (("cognome", "surname", "lastname"), "FULLNAME"),
    (("nome", "firstname", "fullname", "nominativo", "denominazione", "ragionesociale"), "FULLNAME"),
    (("datanascita", "dtnascita", "nascita", "birthdate", "dateofbirth"), "DATE"),
    (("matricola", "tessera", "badge", "idutente", "userid"), "ID_DOC"),
    (("documento", "cartaidentita", "passaporto", "patente"), "ID_DOC"),
    (("indirizzo", "via", "address", "residenza", "domicilio"), "STREET"),
    (("cap", "zip", "postalcode"), "ZIPCODE"),
    (("comune", "citta", "city", "luogonascita"), "CITY"),
    (("targa", "plate"), "TARGA"),
    (("cartacredito", "creditcard", "pan"), "CREDITCARDNUMBER"),
]
# Se una di queste parole e' nel nome, la colonna NON parla di una persona: "Table Name"
# non e' un nome, "Nome file" nemmeno. Senza, `name`/`nome` catturano mezza tabella.
_STOP_TOKENS = {"table", "tabella", "file", "foglio", "report", "campo", "colonna"}
# Tag che da soli identificano una persona: la proposta e' "mask".
IDENTIFYING = {"CF", "PIVA", "IBAN", "EMAIL", "TELEPHONENUM", "FULLNAME", "ID_DOC",
               "CREDITCARDNUMBER", "STREET", "TARGA", "DATE"}


def _norm_header(h):
    """'Società 👥 (Affari)' -> 'societaaffari'; 'cd_fiscale' -> 'cdfiscale'."""
    s = unicodedata.normalize("NFKD", str(h or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def placeholder_prefix(header, tag=None, index=None):
    """Prefisso del placeholder: valido per il ripristino, cioe' [A-Z][A-Z0-9_]*.
    Preferisce il tag riconosciuto, altrimenti l'intestazione ripulita."""
    if tag:
        return tag
    s = unicodedata.normalize("NFKD", str(header or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").upper()[:24].strip("_")
    if not s or not s[0].isalpha():
        s = "COL%s" % (index if index is not None else "")
    return s


def _shape(text):
    """Forma del valore: lettere -> A, cifre -> 9. Non rivela il contenuto."""
    return re.sub(r"[0-9]", "9", re.sub(r"[^\W\d_]", "A", text))[:32]


def _tokens(header):
    """'cd_fisc_soc' -> ['cd','fisc','soc']; 'Table Name' -> ['table','name'];
    'dtNascita' -> ['dt','nascita']."""
    s = unicodedata.normalize("NFKD", str(header or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", s)          # camelCase -> due parole
    return [t for t in re.split(r"[^A-Za-z0-9]+", s.lower()) if t]


def _tag_from_name(header):
    """Tag dedotto dall'intestazione. Il confronto e' per TOKEN, non per sottostringa:
    cercare 'name' dentro 'tablename' (o 'conto' dentro 'sconto') produce falsi positivi
    che maschererebbero colonne tecniche o di misura."""
    toks = _tokens(header)
    if not toks or _STOP_TOKENS.intersection(toks):
        return None
    compatto = "".join(toks)
    for chiavi, tag in _BY_NAME:
        for k in chiavi:
            # nome intero ('cdfiscale'), una parola esatta ('nascita' in dt_nascita),
            # o l'inizio del nome compatto ('cdfisc' in cd_fisc_soc, 'cdiva' in cd_iva_soc)
            if compatto == k or k in toks or (len(k) >= 5 and compatto.startswith(k)):
                return tag
    return None


# Solo queste famiglie decidono da sole guardando i VALORI: hanno un checksum o una
# forma inequivocabile. Le altre (importi, civici, CAP...) sono forme generiche: un
# numero di matricola le soddisfa per caso e scavalcherebbe il nome della colonna.
_STRONG = {"CF", "PIVA", "IBAN", "CREDITCARDNUMBER", "EMAIL", "TELEPHONENUM", "TARGA"}


def _tag_from_values(values):
    """Tag se QUASI TUTTE le celle del campione sono per intero un identificativo
    riconosciuto dalla rete regex+checksum. E' la rete del prodotto, non una copia."""
    if len(values) < 3:
        return None
    conteggi, coperte = {}, 0
    for v in values:
        ents = [e for e in detect_regex(v)
                if e["start"] == 0 and e["end"] == len(v) and e["label"] in _STRONG]
        if not ents:
            continue
        e = max(ents, key=lambda x: (x["validated"], x["end"] - x["start"]))
        conteggi[e["label"]] = conteggi.get(e["label"], 0) + 1
        coperte += 1
    if coperte < 0.8 * len(values):
        return None
    label = max(conteggi, key=conteggi.get)
    return label if conteggi[label] >= 0.8 * len(values) else None


def _sheet(wb, sheet):
    fogli = [ws for ws in wb.worksheets if hasattr(ws, "iter_rows")]
    if not fogli:
        raise XlsxError("Il file non contiene fogli di calcolo.")
    if sheet is None:
        return fogli[0]
    for ws in fogli:
        if ws.title == sheet:
            return ws
    if isinstance(sheet, int) and 0 <= sheet < len(fogli):
        return fogli[sheet]
    raise XlsxError("Foglio non trovato: %r" % (sheet,))


def _load(data, **kw):
    try:
        return load_workbook(io.BytesIO(data), keep_links=False, **kw)
    except Exception as e:
        raise XlsxError("Impossibile leggere il file Excel: serve un .xlsx valido e "
                        f"non protetto da password ({type(e).__name__}).") from e


def profile(data, sheet=None, header_row=1, sample=SAMPLE):
    """Profilo delle colonne + proposta di trattamento. Nessun valore reale esce da qui."""
    wb = _load(data, read_only=True)
    ws = _sheet(wb, sheet)
    righe = ws.iter_rows(values_only=True)
    head = None
    for _ in range(header_row):
        head = next(righe, None)
    if head is None:
        raise XlsxError("Foglio vuoto: nessuna riga di intestazione.")
    head = list(head)
    n = len(head)

    tipi = [{} for _ in range(n)]
    distinti = [set() for _ in range(n)]
    vuote, forme, campioni, lunghezze = [0] * n, [{} for _ in range(n)], [[] for _ in range(n)], [[] for _ in range(n)]
    righe_dati = 0
    for row in righe:
        righe_dati += 1
        for i in range(n):
            v = row[i] if i < len(row) else None
            txt = cell_text(v)
            if txt is None or not txt.strip():
                vuote[i] += 1
                continue
            tipi[i][type(v).__name__] = tipi[i].get(type(v).__name__, 0) + 1
            if len(distinti[i]) < _MAX_DISTINCT:
                distinti[i].add(txt.strip())
            f = _shape(txt.strip())
            forme[i][f] = forme[i].get(f, 0) + 1
            lunghezze[i].append(len(txt.strip()))
            if len(campioni[i]) < sample:
                campioni[i].append(txt.strip())

    colonne = []
    for i in range(n):
        piene = righe_dati - vuote[i]
        tag_nome = _tag_from_name(head[i])
        tag_val = _tag_from_values(campioni[i]) if piene else None
        tag = tag_val or tag_nome
        if tag_val and tag_nome and tag_val != tag_nome:
            motivo = "valori (%s) e nome colonna (%s)" % (tag_val, tag_nome)
        elif tag_val:
            motivo = "valori riconosciuti dalla rete regex/checksum"
        elif tag_nome:
            motivo = "nome della colonna"
        else:
            motivo = ""
        mode = "mask" if (tag in IDENTIFYING and piene) else "keep"
        forma = max(forme[i], key=forme[i].get) if forme[i] else ""
        colonne.append({
            "index": i + 1,
            "header": head[i] if head[i] is not None else "",
            "type": max(tipi[i], key=tipi[i].get) if tipi[i] else "",
            "filled": piene,
            "empty": vuote[i],
            "distinct": len(distinti[i]),
            "distinct_capped": len(distinti[i]) >= _MAX_DISTINCT,
            "length": [min(lunghezze[i]), max(lunghezze[i])] if lunghezze[i] else [0, 0],
            "shape": forma,
            "tag": tag,
            "suggested_mode": mode,
            "reason": motivo,
            # POCHE CATEGORIE IN PROPORZIONE alle righe (~2%, minimo 2, tetto 200):
            # non identifica da sola, ma incrociata con le altre puo' farlo (provincia +
            # sesso + anzianita'). Fuori restano i valori quasi unici (CF, matricola),
            # che sono identificatori diretti e li segnala gia' il tag. L'UI avvisa.
            "quasi_identifier": bool(piene and 1 < len(distinti[i]) <= min(200, max(2, piene // 50))),
        })
    return {
        "sheet": ws.title,
        "sheets": [w.title for w in wb.worksheets if hasattr(w, "iter_rows")],
        "rows": righe_dati,
        "header_row": header_row,
        "columns": colonne,
    }


def _kind(v):
    for typ, name in ((datetime, "datetime"), (date, "date"), (time, "time"),
                      (float, "float"), (int, "int")):
        if isinstance(v, typ):
            return name
    return None


def _year(v):
    if isinstance(v, (datetime, date)):
        return v.year
    m = re.search(r"(19|20)\d{2}", str(v))
    return int(m.group()) if m else None


def anonymize(data, columns, sheet=None, header_row=1):
    """columns: {intestazione (o indice 1-based): "mask"|"year"|"empty"|"keep"}.
    Ritorna (bytes, report). report: mapping, types, columns, cells, left."""
    for h, m in columns.items():
        if m not in MODES:
            raise XlsxError("Modalita' sconosciuta per %r: %r (ammesse: %s)"
                            % (h, m, ", ".join(MODES)))
    wb = _load(data)
    ws = _sheet(wb, sheet)
    head = [c.value for c in next(ws.iter_rows(min_row=header_row, max_row=header_row))]

    # risoluzione delle scelte: per intestazione esatta, normalizzata, o indice
    per_indice = {}
    normalizzate = {_norm_header(h): i + 1 for i, h in enumerate(head)}
    esatte = {str(h): i + 1 for i, h in enumerate(head)}
    ignorate = []
    for chiave, mode in columns.items():
        if mode == "keep":
            continue
        if isinstance(chiave, int) or (isinstance(chiave, str) and chiave.isdigit()):
            idx = int(chiave)
            if not (1 <= idx <= len(head)):
                ignorate.append(chiave)
                continue
        else:
            idx = esatte.get(str(chiave)) or normalizzate.get(_norm_header(chiave))
            if idx is None:
                ignorate.append(chiave)
                continue
        per_indice[idx] = mode

    prefissi = {}
    for idx, mode in per_indice.items():
        if mode == "mask":
            h = head[idx - 1]
            prefissi[idx] = placeholder_prefix(h, _tag_from_name(h), idx)

    mapping, types, visti, contatori, celle = {}, {}, {}, {}, {}
    for row in ws.iter_rows(min_row=header_row + 1):
        for c in row:
            mode = per_indice.get(c.column)
            if mode is None or c.value is None:
                continue
            raw = cell_text(c.value)
            if raw is None or not raw.strip():
                continue                      # celle vuote o di soli spazi: niente da nascondere
            if mode == "empty":
                c.value = None
            elif mode == "year":
                anno = _year(c.value)
                if anno is None:
                    continue
                c.value = anno
                c.number_format = "General"
            else:                             # mask
                pref = prefissi[c.column]
                chiave = (pref, raw.strip().casefold())
                ph = visti.get(chiave)
                if ph is None:
                    contatori[pref] = contatori.get(pref, 0) + 1
                    ph = "[%s_%d]" % (pref, contatori[pref])
                    visti[chiave] = ph
                    mapping[ph] = raw         # con il riempimento: il ripristino torna identico
                    k = _kind(c.value)
                    if k:
                        types[ph] = k
                c.value = ph
            celle[c.column] = celle.get(c.column, 0) + 1

    for prop in ("creator", "lastModifiedBy", "title", "subject", "description",
                 "keywords", "category"):
        if getattr(wb.properties, prop, None):
            setattr(wb.properties, prop, None)

    buf = io.BytesIO()
    wb.save(buf)
    out = buf.getvalue()

    colonne = []
    for idx, mode in sorted(per_indice.items()):
        colonne.append({
            "index": idx,
            "header": head[idx - 1] if head[idx - 1] is not None else "",
            "mode": mode,
            "cells": celle.get(idx, 0),
            "distinct": contatori.get(prefissi.get(idx, ""), 0) if mode == "mask" else 0,
            "prefix": prefissi.get(idx),
        })
    report = {
        "sheet": ws.title,
        "mapping": mapping,
        "types": types,
        "columns": colonne,
        "ignored": ignorate,
        "cells": sum(celle.values()),
        "left": _left(out, ws.title, per_indice, header_row),
    }
    return out, report


def _left(data, sheet, per_indice, header_row):
    """Rete di sicurezza: nelle colonne trattate non deve restare NESSUN valore
    originale. Si contano le celle che non sono un placeholder (per "mask"), non
    sono vuote (per "empty") e non sono un anno (per "year")."""
    ws = _sheet(_load(data, read_only=True), sheet)
    righe = ws.iter_rows(values_only=True)
    for _ in range(header_row):
        next(righe, None)
    left = 0
    for row in righe:
        for idx, mode in per_indice.items():
            v = row[idx - 1] if idx - 1 < len(row) else None
            txt = cell_text(v)
            if txt is None or not txt.strip():
                continue
            if mode == "empty":
                left += 1
            elif mode == "year":
                left += 0 if re.fullmatch(r"(19|20)\d{2}", txt.strip()) else 1
            elif not txt.strip().startswith("["):
                left += 1
    return left
