r"""Hosted-page speed helpers shared by publish_daily.py and publish_refresh.py.

Two transforms applied ONLY to the HOSTED page (the self-contained offline deck in the
dated folders is never touched):

1. extract_fonts(html, repo) — the template inlines ~2.0 MB of base64 woff2 fonts inside
   its <style> block (85% of the shell's wire cost; base64 of already-compressed woff2
   barely gzips). Decode each @font-face data: URI to repo/fonts/<n>.woff2, point the
   CSS at the file, and add font-display:swap so text paints before fonts arrive.
   DEDUPLICATED BY CONTENT: the 28 @font-face rules carry only 8 distinct payloads (each
   subset file is repeated across every weight that uses it), so 8 files are written, not
   28. File names are content-addressed (family + a hash of the bytes) so unchanged fonts
   keep their name and never re-diff in git.

2. pack_data(data) / pack_residency(res) / pack_standing(std) / pack_sip(block) /
   pack_returns(block) — the per-scheme quartile/score month-arrays are single-digit ints
   (or null); packed as strings ('3312-4…', '-'=null) they are ~4x smaller raw and much
   faster to JSON-parse. Residency rows pack as 'a,b,c,d;;…' (';' between months, '' = a
   null month). Standing rank series can be MULTI-digit (rank 12 of 24), so unlike the
   quartile digit-strings they need a separator: comma-joined with '' = null ('8,12,,3').
   The return-profile series are FLOAT percents, so they use that same comma form with a
   float element ('12.34,,-0.5'). A sip convention block reuses the digit/residency/standing
   forms on its sub-blocks (sipret ships untouched — tiny plain floats).
   The page's boot unpacks when it sees _packed=1 (the offline inline path stays plain
   arrays). Floats are rounded (AUM-share 6dp, crores 2dp, returns 2dp) — far inside
   display precision (1dp % / whole cr).
"""
import base64
import hashlib
import re
from pathlib import Path

VR_PACK_FIELDS = ["q1m", "q3m", "q6m", "q9m", "q24m", "q60m", "q1y", "q3y", "y", "t", "b"]
AP_PACK_FIELDS = ["aq1m", "aq3m", "aq6m", "aq9m", "aq24m", "aq60m", "aq1y", "aq3y"]
# Must equal peer_monitor.SIPGAIN_DP — the page re-solves an XIRR from these gains, so the
# wire precision is load-bearing, not cosmetic (see pack_returns).
_SIPGAIN_DP = 4


_FACE_RE = r"@font-face\s*\{[^}]*\}"
_URI_RE = r"url\(\s*['\"]?(data:[^;]+;base64),([A-Za-z0-9+/=]+)['\"]?\s*\)"


def extract_fonts(html: str, repo: Path) -> str:
    """Pull base64 @font-face payloads out to repo/fonts/*.woff2; return rewritten html.

    ONE FILE PER DISTINCT PAYLOAD (deduplicated by content, 2026-08-11). A webfont family is
    shipped as a handful of unicode-range subset files, and the SAME subset bytes are repeated
    in every @font-face rule that uses them — the template's 28 rules (3 families x 4 weights
    x 2 subsets, plus 2 italics x 2 subsets) carry only 8 distinct payloads, each appearing
    2-4 times. Naming a file per RULE, the way this used to (family+weight, with a numeric
    suffix when that name repeated), wrote 28 files holding 1.52 MB where 8 files holding
    0.46 MB say the same thing, and made a returning visitor re-download identical bytes under
    several names because each name is its own cache entry. So: hash the DECODED payload, keep
    one file per distinct hash, and point every rule that carries those bytes at that one file.

    NAMES ARE CONTENT-ADDRESSED — `<family>-<first 8 hex of sha256(bytes)>.woff2`. The name is
    a pure function of the bytes (plus the family label), which is what makes it stable: a font
    that did not change keeps its name and never re-diffs in git, and the name cannot shuffle
    just because a rule moved in the template. The weight is deliberately NOT in the name any
    more — the same file legitimately serves weights 400 through 700, so a weight in the name
    would be a lie. When one payload is shared across families (does not happen today) the
    alphabetically first family labels it, so the name stays order-independent.

    The `fonts/` directory belongs to this function alone (one page per repo, one call per
    publish), so after writing, any leftover .woff/.woff2 that this run did NOT produce is
    removed — otherwise the 28 old per-rule files would sit in the published repo forever and
    the saving would be on the wire only, not in git. The prune is skipped when the html
    carried no @font-face payload at all, so a page that was already extracted can never
    delete the fonts it is still pointing at.

    The offline deck is NOT touched by any of this — it keeps its fonts inlined so it stays a
    single self-contained file.
    """
    fonts_dir = repo / "fonts"
    fonts_dir.mkdir(exist_ok=True)

    # ---- pass 1: decode each payload ONCE, and name it by its content -------------------
    dig_of = {}      # the exact base64 text -> content digest (so pass 2 never re-decodes)
    blob = {}        # digest -> decoded bytes
    ext = {}         # digest -> ".woff2" / ".woff" (from the data: URI's mime)
    fams = {}        # digest -> the family labels whose rules carry these bytes
    for face in re.findall(_FACE_RE, html):
        fam = re.search(r"font-family\s*:\s*['\"]?([^'\";}]+)", face)
        label = (fam.group(1).strip().replace(" ", "") if fam else "font").lower()
        for mime, b64 in re.findall(_URI_RE, face):
            dig = dig_of.get(b64)
            if dig is None:
                raw = base64.b64decode(b64)
                dig = hashlib.sha256(raw).hexdigest()[:8]
                dig_of[b64] = dig
                blob[dig] = raw
            ext.setdefault(dig, ".woff2" if "woff2" in mime else ".woff")
            fams.setdefault(dig, set()).add(label)
    name_of = {d: f"{min(f)}-{d}{ext[d]}" for d, f in fams.items()}

    # ---- write one file per distinct payload, then drop anything stale ------------------
    for dig, name in name_of.items():
        (fonts_dir / name).write_bytes(blob[dig])
    if name_of:                                  # never prune on a page with no fonts inlined
        keep = set(name_of.values())
        for old in fonts_dir.iterdir():
            if old.is_file() and old.suffix in (".woff", ".woff2") and old.name not in keep:
                old.unlink()

    # ---- pass 2: rewrite every rule to point at its file + paint text early -------------
    def _font_face(m):
        block = re.sub(_URI_RE, lambda dm: f"url(fonts/{name_of[dig_of[dm.group(2)]]})", m.group(0))
        if "font-display" not in block:
            block = block.replace("@font-face{", "@font-face{font-display:swap;", 1) \
                if "@font-face{" in block else re.sub(r"(@font-face\s*\{)", r"\1font-display:swap;", block, count=1)
        return block

    return re.sub(_FACE_RE, _font_face, html)


def _pack_series(arr):
    """[1,3,None,4,...] (single-digit ints) -> '13-4...'; None -> '-'. Non-list passes through."""
    if not isinstance(arr, list):
        return arr
    out = []
    for v in arr:
        if v is None:
            out.append("-")
        else:
            iv = int(v)
            if iv < 0 or iv > 9 or iv != v:          # not single-digit-safe: leave array as-is
                return arr
            out.append(str(iv))
    return "".join(out)


def _round_series(arr, nd):
    return [None if v is None else round(v, nd) for v in arr] if isinstance(arr, list) else arr


def pack_data(data: dict) -> dict:
    """In-place pack of the vr/allpeer month-series; sets _packed=1."""
    for r in data.get("vr", []):
        for f in VR_PACK_FIELDS:
            if f in r:
                r[f] = _pack_series(r[f])
        r["a"] = _round_series(r.get("a"), 6)
        r["cr"] = _round_series(r.get("cr"), 2)
    for r in data.get("allpeer", []):
        for f in AP_PACK_FIELDS:
            if f in r:
                r[f] = _pack_series(r[f])
        r["a"] = _round_series(r.get("a"), 6)
        r["cr"] = _round_series(r.get("cr"), 2)
    data["_packed"] = 1
    return data


# ---- the three {f,v}-record wire transforms, factored so every caller shares ONE form ----
# residency and standing have the SAME nesting ({uni:{key:{win:{'f':i,'v':[...]}}}}); only
# the per-element string form differs. Each transform lives in exactly one helper, used by
# BOTH the top-level file packer (pack_residency / pack_standing) and pack_sip's sub-block
# path — so the hosted file and the sip file can never drift onto different wire formats.
# The returns block (added 2026-08-11) is the third member: same {f,v} record, same
# comma-joined string, but FLOAT elements and mixed nesting depth — see _returns_v_str /
# _pack_returns_inner below for why it needs its own element helper and its own walker.

def _residency_v_str(v):
    """One residency 'v' list of [q1..q4] rows -> its wire string: each row's 4 values joined
    with ',', rows joined with ';', a null row becomes the empty string ('a,b,c,d;;e,f,g,h')."""
    return ";".join("" if row is None else ",".join(map(str, row)) for row in v)


def _pack_residency_inner(res: dict) -> dict:
    """Walk {uni:{scheme:{win:{'f':i,'v':[...]}}}} and stringify each 'v' IN PLACE (no envelope)."""
    for uni in res.values():
        for sch in uni.values():
            for rec in sch.values():
                v = rec.get("v")
                if isinstance(v, list):
                    rec["v"] = _residency_v_str(v)
    return res


def pack_residency(res: dict) -> dict:
    """{uni:{scheme:{win:{'f':i,'v':[[q1..q4]|None,..]}}}} -> v as 'a,b,c,d;;...' ('' = null)."""
    if not res:
        return res
    return {"_packed": 1, "data": _pack_residency_inner(res)}


def _standing_v_str(v):
    """One standing 'v' list of ranks -> comma-joined wire string: [8,12,None,3] -> '8,12,,3'.
    Ranks are 1-based positions so they can be MULTI-digit — hence the ',' separator instead
    of the quartile digit-string form; an empty field between commas = null (not rated that
    month). A non-integer value means the record is not wire-safe: leave the plain array
    (same bail rule as _pack_series; the page treats a list 'v' as already unpacked)."""
    out = []
    for r in v:
        if r is None:
            out.append("")
        else:
            ir = int(r)
            if ir != r:                              # non-integer rank: keep the array as-is
                return v
            out.append(str(ir))
    return ",".join(out)


def _pack_standing_inner(std: dict) -> dict:
    """Walk {uni:{key:{win:{'f':i,'v':[rank|None,..]}}}} and stringify each 'v' IN PLACE
    (no envelope). Same nesting as residency: uni ('all'/'vr') -> scheme or 'cat|scheme'
    membership key -> window label -> {f,v} record."""
    for uni in std.values():
        for key in uni.values():
            for rec in key.values():
                v = rec.get("v")
                if isinstance(v, list):
                    rec["v"] = _standing_v_str(v)
    return std


def pack_standing(std: dict) -> dict:
    """Category-standing rank block -> v as '8,12,,3' ('' = null); {_packed:1, data:...}
    envelope exactly like pack_residency, so the page's fetch path handles both alike."""
    if not std:
        return std
    return {"_packed": 1, "data": _pack_standing_inner(std)}


def _returns_v_str(v, nd=2):
    """One return-profile 'v' list of percent floats -> comma-joined wire string:
    [12.34, None, -0.5, 100.0] -> '12.34,,-0.5,100'.

    `nd` = decimals to keep. It is a PARAMETER because the engine deliberately ships two
    precisions: displayed returns (`lump`, `bench`) at 2 dp, but `sipgain` at 4 dp because
    the page RE-SOLVES an XIRR from it — a 2 dp gain on a days-old single instalment
    annualises with an exponent near 365/4, turning half a basis point of rounding into
    ~0.5 percentage points of annual rate. Rounding sipgain to 2 dp here would also make the
    hosted page disagree with the self-contained offline deck (which carries the raw 4 dp
    JSON), which is the one failure mode this deck's contract forbids outright.

    Same wire SHAPE as the standing form (',' separator, an empty field between commas =
    null) but a SIBLING helper because these values are FLOATS, not ranks. _standing_v_str
    deliberately bails back to the plain array the moment `int(x) != x` — which every
    non-whole return (12.34) trips — so reusing it would silently ship the whole block
    un-packed. The float rules, all chosen so the page's `x === '' ? null : +x` decoder
    reads back the same number it was given:
      * round to 2 dp defensively. The engine already rounds, but float arithmetic can hand
        us 12.340000000000001 and that would put 15 useless characters on the wire.
      * print the SHORTEST exact decimal — '100.00' -> '100', '5.50' -> '5.5'. Trailing
        zeros are pure wire cost; `+x` parses either form to the identical double. (Fixed
        '.2f' formatting first, never '%g', so a big value like 123456.78 keeps every digit
        instead of being cut to 6 significant figures.)
      * exact zero prints '0', never '-0'.
      * a NON-FINITE value (NaN / inf) is written as an empty field, i.e. null. Bare NaN
        survives Python's json.dumps but KILLS the browser's JSON.parse, so it must never
        reach the wire — and a NaN return means "no value" here anyway, which is exactly
        what null means.
    A value that is not a number at all means the record is not wire-safe: leave the plain
    array (same bail rule as _pack_series / _standing_v_str — the page treats a list 'v' as
    already unpacked, so an un-packed record still renders)."""
    out = []
    for x in v:
        if x is None:
            out.append("")
            continue
        try:
            f = round(float(x), nd)
        except (TypeError, ValueError):      # not a number: keep the array as-is
            return v
        if f != f or f in (float("inf"), float("-inf")):
            out.append("")                   # NaN/inf -> null (browser-JSON safety)
        elif f == 0:
            out.append("0")                  # kills a '-0' from a tiny negative rounding
        else:
            s = f"{f:.{nd}f}"
            out.append(s.rstrip("0").rstrip("."))
    return ",".join(out)


def _pack_returns_inner(node, nd=2):
    """Recursively stringify every {'f':i,'v':[...]} record's 'v' IN PLACE, at any depth.

    Residency and standing are always exactly three levels (uni -> key -> window -> record),
    so they can use hand-counted loops. The returns block is NOT one shape: `lump` is
    uni -> scheme -> window -> record (3), `bench` is category -> window -> record (2), and
    `sipgain` is convention -> uni -> key -> window -> record (4). The design contract also
    allows `bench` to be shipped once PER AXIS if the all-peer as-of dates turn out not to be
    a subset of the month axis, which adds a level again. One depth-agnostic walk packs every
    shape the engine may legitimately emit, instead of three hand-counted loops that silently
    stop packing the day a level moves. A record is recognised by carrying a LIST 'v' — a
    scheme or category that happened to be named "v" holds a dict, so it just recurses."""
    if not isinstance(node, dict):
        return node
    if isinstance(node.get("v"), list):
        node["v"] = _returns_v_str(node["v"], nd)
        return node
    for child in node.values():
        _pack_returns_inner(child, nd)
    return node


def pack_returns(block: dict) -> dict:
    """Return-profile block -> {_packed:1, data:...}, the same envelope pack_standing uses so
    the page's lazy-fetch path handles all three files alike. Every {f,v} float series inside
    `lump` / `bench` / `sipgain` is comma-joined ('' = null); `sipdays` is left ALONE.

    Why sipdays rides plain: `sipdays[conv][win][asofIdx]` is a list of integer day offsets
    `(asof - installment_date).days` that the page feeds straight into its XIRR bisection. It
    is a nested list of small ints — already about as compact as JSON gets — and, unlike a
    {f,v} series, it is not a trimmed axis-aligned array, so there is nothing to gain from
    stringifying it and a real risk in reshaping the one input the solver depends on. It is
    skipped BY NAME rather than by shape, so nothing that ever gets nested inside it can be
    mistaken for a packable record.
    PRECISION IS PER-SUBTREE, and it must match what the engine emitted. `lump`/`bench` are
    DISPLAYED numbers and ship at 2 dp; `sipgain` is RE-SOLVED into an XIRR by the page, so
    the engine emits it at 4 dp (peer_monitor.SIPGAIN_DP) and it must be packed at 4 dp too.
    Packing it at 2 dp would (a) reintroduce the ~0.5 pp annual-rate error on short windows
    that the 4 dp decision was taken to remove, and (b) make the hosted page disagree with
    the offline deck, which carries the unpacked 4 dp JSON. If the engine's SIPGAIN_DP ever
    changes, change _SIPGAIN_DP here in the same commit.
    None passes through unchanged, so a missing block still serializes as 'null'."""
    if not block:
        return block
    for key, sub in block.items():
        if key == "sipdays":                 # plain nested ints — see the docstring
            continue
        _pack_returns_inner(sub, _SIPGAIN_DP if key == "sipgain" else 2)
    return {"_packed": 1, "data": block}


def pack_sip(block: dict) -> dict:
    """In-place pack of ONE sip convention block ({vr, allpeer, residency, standing, sipret}).

    vr/allpeer here are dicts keyed by membership ('cat|scheme') / scheme name — unlike the
    top-level vr/allpeer row LISTS — but their month-series carry the same single-digit
    values (quartiles 1-4, y/t/b score digits), so each field goes through the same
    _pack_series digit-string form pack_data uses. residency/standing sub-blocks reuse the
    shared inner transforms above (identical wire form to residency.json / standing.json,
    just without their own envelope — ONE _packed=1 on the whole file tells the page every
    sub-block is packed). sipret is a tiny {scheme:{win:[gain,xirr]}} float map and ships
    untouched. Returns None unchanged so a missing convention still serializes ('null')."""
    if not block:
        return block
    for row in (block.get("vr") or {}).values():
        for f in VR_PACK_FIELDS:
            if f in row:
                row[f] = _pack_series(row[f])
    for row in (block.get("allpeer") or {}).values():
        for f in AP_PACK_FIELDS:
            if f in row:
                row[f] = _pack_series(row[f])
    res = block.get("residency")
    if isinstance(res, dict):
        _pack_residency_inner(res)
    std = block.get("standing")
    if isinstance(std, dict):
        _pack_standing_inner(std)
    block["_packed"] = 1
    return block
