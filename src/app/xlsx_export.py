# -*- coding: utf-8 -*-
"""
Anonimizzazione dei fogli Excel (.xlsx).

Stesso patto di pdf_export.py: nessuna dipendenza dal modello. Il modulo lavora su
bytes + span gia' decise da analyze(), quindi e' testabile in isolamento (serve
solo openpyxl, niente torch/transformers/fitz).

Il flusso e' in due tempi, con analyze() in mezzo:

  layout = read_workbook(xlsx_bytes)
      Legge il workbook e lo STENDE in un unico testo (layout.text): un blocco per
      foglio, una riga di testo per riga del foglio, celle separate da TAB. E' la
      stessa forma di un incolla da Excel, che l'app gestisce gia' (issue #54).
      Ogni pezzo di testo ricorda da dove viene (una UNITA': cella, stringa
      letterale di una formula, commento, intestazione/pie' di pagina, nome del
      foglio). TAB e a-capo DENTRO una cella diventano spazi, carattere per
      carattere: gli offset restano 1:1 con il valore originale.

  res = analyze(layout.text, excluded, mapping_enabled=True)     # in app.py

  out, report = write_workbook(layout, spans_from_segments(res["segments"]),
                               res["mapping"])
      Riporta ogni placeholder nella sua unita' e salva un nuovo .xlsx.

Perche' un testo solo e non un'inferenza per cella:
  - il modello vede la RIGA come contesto: "ROSSI" da solo e' ambiguo, accanto a
    "GIACOMO" e a un codice fiscale no;
  - la numerazione dei placeholder e' coerente su TUTTO il file: lo stesso
    soggetto e' [FULLNAME_3] in ogni foglio;
  - un'inferenza a chunk invece di migliaia di chiamate.

Oltre alle span del modello, ogni valore del dizionario viene cercato OVUNQUE nel
workbook (come fa redact_pdf): se il modello riconosce "ROSSI" alla riga 2 ma non
alla 40, la riga 40 viene anonimizzata lo stesso. Ricerca per token, a confini di
parola, senza scavalcare una cella: "DE" non viene toccato dentro "CORDELLA".

Resta com'era: formattazione, larghezze, celle unite, e il TIPO delle celle non
toccate (un numero resta un numero). Una cella numerica o data in cui cade una PII
diventa testo: il placeholder non e' un numero.

Ripulito oltre alle celle - tutti posti dove una PII sopravvive senza vedersi:
  - stringhe letterali dentro le FORMULE (="Mario " & A2);
  - testo dei COMMENTI, e l'autore di ogni commento;
  - INTESTAZIONI e PIE' DI PAGINA di stampa;
  - NOMI DEI FOGLI (i placeholder perdono le quadre, vietate da Excel). NB: le
    formule che citano il foglio per nome non vengono riscritte;
  - COLLEGAMENTI IPERTESTUALI il cui indirizzo contiene un valore trovato
    (tipicamente mailto:), rimossi;
  - proprieta' del documento (autore, ultimo autore, titolo...) e proprieta'
    personalizzate;
  - TABELLE PIVOT, rimosse: la loro cache e' una copia integrale dei dati sorgente,
    invisibile a video ma dentro il file;
  - collegamenti a cartelle esterne (keep_links=False).
openpyxl non conserva grafici e immagini di un file letto: nell'output non ci sono.

Rete di sicurezza finale, come per il PDF: il file prodotto viene RILETTO e ogni
valore del dizionario cercato ovunque, compresi i posti che non riscriviamo (nomi
definiti, convalide dati). `report["residual"]` = placeholder il cui valore e'
ancora leggibile: deve essere vuota, l'UI avvisa se non lo e'.
`report["skipped"]` = valori troppo corti per essere cercati nel resto del file
(< 2 caratteri alfanumerici o 2 sole cifre): sono sostituiti dove il modello li ha
trovati, ma altrove possono restare in chiaro.

RIPRISTINO - restore_workbook(xlsx_bytes, mapping, types)
  Il percorso inverso: rimette i valori veri al posto dei placeholder, anche in un
  file modificato nel frattempo (righe/colonne aggiunte da un LLM). Stessa
  tolleranza della scheda "Ripristina" dell'UI: quadre facoltative e spazi interni
  ("[ FULLNAME_1 ]", "FULLNAME_1" - senza quadre e' anche la forma dei nomi dei
  fogli). `types` ({placeholder: "int"|"float"|"datetime"|"date"|"time"}, prodotto
  da write_workbook in report["types"]) ridà il tipo alle celle che erano INTERAMENTE
  un numero o una data. Limiti onesti: il dizionario tiene una sola grafia per
  valore ("ROSSI" e "Rossi" tornano uguali), e cio' che l'anonimizzazione ha
  rimosso (autore, collegamenti, pivot) non torna.
"""

import bisect
import io
import re
from datetime import date, datetime, time

from openpyxl import load_workbook
from openpyxl.comments import Comment


class XlsxError(ValueError):
    """Errore d'uso (file non valido, protetto da password, input incoerente...)."""


_CTRL = re.compile(r"[\t\r\n]")
# stringhe letterali di una formula: "..." con "" come virgoletta escapata
_FORMULA_LIT = re.compile(r'"((?:[^"]|"")*)"')
_HF_ATTRS = ("oddHeader", "oddFooter", "evenHeader", "evenFooter",
             "firstHeader", "firstFooter")
_HF_PARTS = ("left", "center", "right")
_DOC_PROPS = ("creator", "lastModifiedBy", "title", "subject", "description",
              "keywords", "category", "identifier", "contentStatus")
_TITLE_BAD = re.compile(r"[\\/?*:\[\]]")
_TOK = re.compile(r"\w+|[^\w\s]")
# autore dei commenti nell'output. Non vuoto: con "" openpyxl scrive la stringa "None"
# e Excel la mostra in testa al commento.
COMMENT_AUTHOR = "anonimo"


# --------------------------------------------------------------------------- #
# Lettura: workbook -> testo + unita'
# --------------------------------------------------------------------------- #
def cell_text(value):
    """Come appare il valore di una cella nel testo steso. None = niente testo
    (celle vuote, booleani, durate: nessuna PII possibile)."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, datetime):
        fmt = "%d/%m/%Y" if value.time() == time(0) else "%d/%m/%Y %H:%M"
        return value.strftime(fmt)
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    if isinstance(value, int):
        return str(value)
    return None


def _kind(value):
    """Tipo di una cella non testuale, per ridarglielo al ripristino."""
    for typ, name in ((datetime, "datetime"), (date, "date"), (time, "time"),
                      (float, "float"), (int, "int")):
        if isinstance(value, typ):
            return name
    return None


class _Unit:
    """Un pezzo di layout.text che viene da `target`, a partire da src[a]."""

    __slots__ = ("start", "end", "target", "a")

    def __init__(self, start, end, target, a):
        self.start, self.end, self.target, self.a = start, end, target, a


class XlsxLayout:
    """Il workbook caricato + il suo testo steso. Monouso: write_workbook() lo
    modifica sul posto."""

    def __init__(self, wb):
        self.wb = wb
        self.sheets = list(wb.worksheets)
        self.units = []
        self.targets = {}          # target -> stringa sorgente originale
        self.kinds = {}            # target di celle numero/data -> _kind()
        self._parts = []
        self._pos = 0

    @property
    def text(self):
        return "".join(self._parts)

    def _lit(self, s):
        self._parts.append(s)
        self._pos += len(s)

    def _unit(self, target, src, a=0, b=None):
        b = len(src) if b is None else b
        piece = _CTRL.sub(" ", src[a:b])
        if not piece.strip():
            return
        self.targets[target] = src
        self.units.append(_Unit(self._pos, self._pos + len(piece), target, a))
        self._lit(piece)


def read_workbook(data):
    """bytes .xlsx -> XlsxLayout. Solleva XlsxError se il file non si legge."""
    try:
        wb = load_workbook(io.BytesIO(data), keep_links=False)
    except Exception as e:                      # zip rotto, cifrato, .xls travestito...
        raise XlsxError("Impossibile leggere il file Excel: serve un .xlsx valido e "
                        f"non protetto da password ({type(e).__name__}).") from e
    lay = XlsxLayout(wb)
    for si, ws in enumerate(lay.sheets):
        lay._unit(("title", si), ws.title)
        lay._lit("\n")
        if not hasattr(ws, "iter_rows"):        # chartsheet: solo il nome
            lay._lit("\n")
            continue
        comments = []
        for row in ws.iter_rows():
            last = -1
            for i, c in enumerate(row):
                if c.comment is not None and c.comment.text:
                    comments.append(c)
                v = c.value
                if (_FORMULA_LIT.search(v) if isinstance(v, str) and c.data_type == "f"
                        else cell_text(v)):
                    last = i                        # una formula senza letterali non e' testo
            if last < 0:
                continue
            for i, c in enumerate(row[:last + 1]):
                if i:
                    lay._lit("\t")
                v = c.value
                if isinstance(v, str) and c.data_type == "f":
                    lits = list(_FORMULA_LIT.finditer(v))
                    for k, m in enumerate(lits):
                        if k:
                            lay._lit(" ")
                        lay._unit(("formula", si, c.row, c.column), v, m.start(1), m.end(1))
                    continue
                txt = cell_text(v)
                if txt:
                    lay._unit(("cell", si, c.row, c.column), txt)
                    if not isinstance(v, str):
                        lay.kinds[("cell", si, c.row, c.column)] = _kind(v)
            lay._lit("\n")
        for c in comments:
            lay._unit(("comment", si, c.row, c.column), c.comment.text)
            lay._lit("\n")
        for attr in _HF_ATTRS:
            hf = getattr(ws, attr, None)
            for part in _HF_PARTS:
                p = getattr(hf, part, None)
                if p is not None and p.text:
                    lay._unit(("hf", si, attr, part), p.text)
                    lay._lit("\n")
        lay._lit("\n")
    return lay


def workbook_text(data):
    """Solo il testo steso (per /analyze e l'anteprima testuale)."""
    return read_workbook(data).text


# --------------------------------------------------------------------------- #
# Span: dal risultato di analyze() + propagazione dei valori del dizionario
# --------------------------------------------------------------------------- #
def spans_from_segments(segments):
    """[(start, end, placeholder)] sul testo analizzato, dai segmenti di analyze().
    Richiede mapping_enabled=True: senza il campo `t` gli offset non si ricostruiscono."""
    spans, pos = [], 0
    for s in segments:
        if "t" not in s:
            raise XlsxError("Segmenti senza testo: analyze() va chiamata con "
                            "mapping_enabled=True.")
        if "label" in s:
            spans.append((pos, pos + len(s["t"]), s["ph"]))
        pos += len(s["t"])
    return spans


def _too_short(value):
    """Valori non cercabili nel resto del file senza devastarlo (stessa soglia del PDF)."""
    alnum = re.sub(r"[\W_]+", "", (value or "").casefold())
    return len(alnum) < 2 or (len(alnum) == 2 and alnum.isdigit())


class _ValueIndex:
    """Ricerca dei valori del dizionario per TOKEN (\\w+ o singola punteggiatura),
    case-insensitive. Tra due token di un valore sono ammessi solo spazi: una
    corrispondenza non scavalca mai TAB o a-capo, cioe' non passa da una cella
    all'altra. Indicizzato per primo token: niente regex con migliaia di alternative."""

    def __init__(self, mapping):
        self.by_first, self.skipped = {}, []
        seen = set()
        for ph, val in (mapping or {}).items():
            toks = tuple(t.casefold() for t in _TOK.findall(val or ""))
            if not toks or toks in seen:
                continue
            if _too_short(val):
                self.skipped.append(ph)
                continue
            seen.add(toks)
            self.by_first.setdefault(toks[0], []).append((toks, ph))
        for cands in self.by_first.values():
            cands.sort(key=lambda x: -len(x[0]))      # prima il valore piu' lungo

    def find(self, text):
        if not text or not self.by_first:
            return []
        toks = list(_TOK.finditer(text))
        low = [m.group().casefold() for m in toks]
        out, i = [], 0
        while i < len(toks):
            step = 1
            for vt, ph in self.by_first.get(low[i], ()):
                n = len(vt)
                if tuple(low[i:i + n]) != vt:
                    continue
                if any(text[toks[j].end():toks[j + 1].start()].strip("  ")
                       for j in range(i, i + n - 1)):
                    continue
                out.append((toks[i].start(), toks[i + n - 1].end(), ph))
                step = n
                break
            i += step
        return out


def _with_propagation(text, spans, index):
    """Span del modello + occorrenze dei valori del dizionario non gia' coperte."""
    kept = sorted(spans)
    starts = [s for s, _, _ in kept]
    for s, e, ph in index.find(text):
        i = bisect.bisect_right(starts, s)
        if (i and kept[i - 1][1] > s) or (i < len(kept) and kept[i][0] < e):
            continue
        kept.insert(i, (s, e, ph))
        starts.insert(i, s)
    return kept


def _distribute(layout, spans):
    """target -> [(a, b, placeholder)] sulla stringa sorgente del target. Una span a
    cavallo di due unita' mette il placeholder in entrambe: nel dubbio si maschera."""
    starts = [u.start for u in layout.units]
    edits = {}
    for s, e, ph in spans:
        i = max(0, bisect.bisect_right(starts, s) - 1)
        while i < len(layout.units) and layout.units[i].start < e:
            u = layout.units[i]
            if u.end > s:
                a = u.a + max(s, u.start) - u.start
                b = u.a + min(e, u.end) - u.start
                edits.setdefault(u.target, []).append((a, b, ph))
            i += 1
    return edits


def _apply(src, edits):
    out, pos = [], 0
    for a, b, ph in sorted(edits):
        if a < pos:                                   # sovrapposizione: gia' coperta
            continue
        out.append(src[pos:a])
        out.append(ph)
        pos = b
    out.append(src[pos:])
    return "".join(out)


# --------------------------------------------------------------------------- #
# Scrittura
# --------------------------------------------------------------------------- #
def _safe_title(title, taken):
    t = _TITLE_BAD.sub(lambda m: "" if m.group() in "[]" else "_", title).strip("' ")
    t = (t or "Foglio")[:31]
    base, n = t, 2
    while t.casefold() in taken:
        suffix = f" ({n})"
        t = base[:31 - len(suffix)] + suffix
        n += 1
    return t


def write_workbook(layout, spans, mapping):
    """Applica le span al workbook e lo salva. Ritorna (bytes, report)."""
    wb = layout.wb
    index = _ValueIndex(mapping)
    edits = _distribute(layout, _with_propagation(layout.text, spans, index))
    report = {"occurrences": 0, "cells": 0, "renamed_sheets": 0, "removed_pivots": 0,
              "unlinked": 0, "skipped": sorted(index.skipped), "residual": [],
              "types": {}}

    titles = {}
    for target, lst in edits.items():
        new = _apply(layout.targets[target], lst)
        report["occurrences"] += len(lst)
        kind, si = target[0], target[1]
        ws = layout.sheets[si]
        if kind in ("cell", "formula"):
            cell = ws.cell(row=target[2], column=target[3])
            cell.value = new
            if kind == "cell" and new.startswith("="):
                cell.data_type = "s"                  # era testo, non deve diventare formula
            src = layout.targets[target]
            if (target in layout.kinds and len(lst) == 1
                    and lst[0][0] == 0 and lst[0][1] == len(src)):
                report["types"][lst[0][2]] = layout.kinds[target]   # la cella INTERA era un numero/data
            report["cells"] += 1
        elif kind == "comment":
            ws.cell(row=target[2], column=target[3]).comment = Comment(new, COMMENT_AUTHOR)
        elif kind == "hf":
            getattr(getattr(ws, target[2]), target[3]).text = new
        elif kind == "title":
            titles[si] = new

    if titles:
        taken = {ws.title.casefold() for si, ws in enumerate(layout.sheets) if si not in titles}
        for si, new in titles.items():
            safe = _safe_title(new, taken)
            taken.add(safe.casefold())
            layout.sheets[si].title = safe
            report["renamed_sheets"] += 1

    for ws in layout.sheets:
        if hasattr(ws, "iter_rows"):
            for row in ws.iter_rows():
                for c in row:
                    if c.comment is not None and c.comment.author != COMMENT_AUTHOR:
                        c.comment = Comment(c.comment.text, COMMENT_AUTHOR)
                    h = c.hyperlink
                    if h is not None and any(index.find(x) for x in
                                             (h.target, h.location, h.tooltip, h.display) if x):
                        c.hyperlink = None
                        report["unlinked"] += 1
        pivots = getattr(ws, "_pivots", None)
        if pivots:
            report["removed_pivots"] += len(pivots)
            ws._pivots = []

    for prop in _DOC_PROPS:
        if getattr(wb.properties, prop, None):
            setattr(wb.properties, prop, None)
    try:
        from openpyxl.packaging.custom import CustomPropertyList
        wb.custom_doc_props = CustomPropertyList()
    except ImportError:                               # openpyxl < 3.1: non le legge nemmeno
        pass

    buf = io.BytesIO()
    wb.save(buf)
    out = buf.getvalue()
    report["residual"] = _residual(out, index)
    return out, report


# --------------------------------------------------------------------------- #
# Verifica finale
# --------------------------------------------------------------------------- #
def _all_strings(wb):
    """Ogni stringa del workbook in cui una PII potrebbe sopravvivere."""
    for ws in wb.worksheets:
        yield ws.title
        if hasattr(ws, "iter_rows"):
            for row in ws.iter_rows():
                for c in row:
                    yield cell_text(c.value)
                    if c.comment is not None:
                        yield c.comment.text
                        yield c.comment.author
                    if c.hyperlink is not None:
                        h = c.hyperlink
                        yield from (h.target, h.location, h.tooltip, h.display)
            for attr in _HF_ATTRS:
                hf = getattr(ws, attr, None)
                for part in _HF_PARTS:
                    yield getattr(getattr(hf, part, None), "text", None)
            dvs = getattr(getattr(ws, "data_validations", None), "dataValidation", None) or []
            for dv in dvs:
                yield dv.formula1
                yield dv.formula2
            yield from _defined_names(getattr(ws, "defined_names", None))
    yield from _defined_names(wb.defined_names)
    for prop in _DOC_PROPS:
        yield getattr(wb.properties, prop, None)


def _defined_names(names):
    if names is None:
        return
    items = names.values() if hasattr(names, "values") else getattr(names, "definedName", [])
    for dn in items:
        yield getattr(dn, "attr_text", None)


def _residual(xlsx_bytes, index):
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    found = set()
    for s in _all_strings(wb):
        if isinstance(s, str):
            found.update(ph for _, _, ph in index.find(s))
    return sorted(found)


# --------------------------------------------------------------------------- #
# Ripristino: placeholder -> valori veri
# --------------------------------------------------------------------------- #
TYPES_KEY = "_types"       # chiave del dizionario scaricato dall'UI con i tipi delle celle
# "[FULLNAME_1]", "[ ID_DOC_2 ]" oppure senza quadre, a confini di parola
_PH_RE = re.compile(r"\[\s*([A-Z][A-Z0-9_]*_\d+)\s*\]|\b([A-Z][A-Z0-9_]*_\d+)\b")
_PARSERS = {
    "int": int,
    "float": float,
    "datetime": lambda s: datetime.strptime(s, "%d/%m/%Y %H:%M" if " " in s else "%d/%m/%Y"),
    "date": lambda s: datetime.strptime(s, "%d/%m/%Y").date(),
    "time": lambda s: datetime.strptime(s, "%H:%M").time(),
}


def parse_dictionary(obj):
    """Dizionario come lo scarica l'UI: {placeholder: valore, ..., "_types": {...}}.
    Ritorna (mapping, types). Le chiavi che non sono placeholder vengono ignorate."""
    if not isinstance(obj, dict):
        raise XlsxError("Il dizionario deve essere un oggetto JSON {placeholder: valore}.")
    mapping = {k: v for k, v in obj.items()
               if isinstance(v, str) and _PH_RE.fullmatch(k) and k.startswith("[")}
    if not mapping:
        raise XlsxError("Dizionario vuoto: nessun placeholder da ripristinare.")
    types = obj.get(TYPES_KEY)
    return mapping, (types if isinstance(types, dict) else {})


def _typed(text, kind):
    parse = _PARSERS.get(kind)
    if parse is None:
        return None
    try:
        return parse(text.strip())
    except ValueError:
        return None


def restore_workbook(data, mapping, types=None):
    """xlsx anonimizzato + dizionario -> (bytes, report). report: "restored"
    (placeholder sostituiti), "cells", "typed" (celle tornate numero/data), "unknown"
    (placeholder tra quadre che il dizionario non conosce: restano come sono)."""
    types = types or {}
    try:
        wb = load_workbook(io.BytesIO(data), keep_links=False)
    except Exception as e:
        raise XlsxError("Impossibile leggere il file Excel: serve un .xlsx valido e "
                        f"non protetto da password ({type(e).__name__}).") from e
    report = {"restored": 0, "cells": 0, "typed": 0, "unknown": set()}

    def sub(text, formula=False):
        n = 0

        def rep(m):
            nonlocal n
            name = m.group(1) or m.group(2)
            val = mapping.get(f"[{name}]")
            if val is None:
                if m.group(1):
                    report["unknown"].add(f"[{name}]")
                return m.group(0)
            n += 1
            # nelle formule i placeholder stanno dentro un letterale "...": virgolette raddoppiate
            return val.replace('"', '""') if formula else val

        new = _PH_RE.sub(rep, text)
        report["restored"] += n
        return new, n

    titles = {}
    for si, ws in enumerate(wb.worksheets):
        new, n = sub(ws.title)
        if n:
            titles[si] = new
        if not hasattr(ws, "iter_rows"):
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str):
                    formula = c.data_type == "f"
                    new, n = sub(v, formula)
                    if n:
                        report["cells"] += 1
                        whole = None if formula else _PH_RE.fullmatch(v.strip())
                        kind = types.get(f"[{whole.group(1) or whole.group(2)}]") if whole else None
                        typed = _typed(new, kind)
                        if typed is not None:
                            c.value = typed
                            report["typed"] += 1
                        else:
                            c.value = new
                            if not formula and new.startswith("="):
                                c.data_type = "s"
                if c.comment is not None and c.comment.text:
                    new, n = sub(c.comment.text)
                    if n:
                        c.comment = Comment(new, c.comment.author or COMMENT_AUTHOR)
        for attr in _HF_ATTRS:
            hf = getattr(ws, attr, None)
            for part in _HF_PARTS:
                p = getattr(hf, part, None)
                if p is not None and p.text:
                    new, n = sub(p.text)
                    if n:
                        p.text = new

    if titles:
        sheets = wb.worksheets
        taken = {ws.title.casefold() for si, ws in enumerate(sheets) if si not in titles}
        for si, new in titles.items():
            safe = _safe_title(new, taken)
            taken.add(safe.casefold())
            sheets[si].title = safe

    buf = io.BytesIO()
    wb.save(buf)
    report["unknown"] = sorted(report["unknown"])
    return buf.getvalue(), report
