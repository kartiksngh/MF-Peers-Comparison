r"""Hosted-page speed helpers shared by publish_daily.py and publish_refresh.py.

Two transforms applied ONLY to the HOSTED page (the self-contained offline deck in the
dated folders is never touched):

1. extract_fonts(html, repo) — the template inlines ~2.0 MB of base64 woff2 fonts inside
   its <style> block (85% of the shell's wire cost; base64 of already-compressed woff2
   barely gzips). Decode each @font-face data: URI to repo/fonts/<n>.woff2, point the
   CSS at the file, and add font-display:swap so text paints before fonts arrive.
   File names are stable (family-weight based) so unchanged fonts never re-diff in git.

2. pack_data(data) / pack_residency(res) — the per-scheme quartile/score month-arrays are
   single-digit ints (or null); packed as strings ('3312-4…', '-'=null) they are ~4x
   smaller raw and much faster to JSON-parse. The page's boot unpacks when it sees
   _packed=1 (the offline inline path stays plain arrays). Floats are rounded
   (AUM-share 6dp, crores 2dp) — far inside display precision (1dp % / whole cr).
"""
import base64
import re
from pathlib import Path

VR_PACK_FIELDS = ["q1m", "q3m", "q6m", "q9m", "q24m", "q60m", "q1y", "q3y", "y", "t", "b"]
AP_PACK_FIELDS = ["aq1m", "aq3m", "aq6m", "aq9m", "aq24m", "aq60m", "aq1y", "aq3y"]


def extract_fonts(html: str, repo: Path) -> str:
    """Pull base64 @font-face payloads out to repo/fonts/*.woff2; return rewritten html."""
    fonts_dir = repo / "fonts"
    fonts_dir.mkdir(exist_ok=True)
    counters = {}

    def _font_face(m):
        block = m.group(0)
        fam = re.search(r"font-family\s*:\s*['\"]?([^'\";}]+)", block)
        wgt = re.search(r"font-weight\s*:\s*(\d+)", block)
        base = (fam.group(1).strip().replace(" ", "") if fam else "font").lower()
        name = f"{base}-{wgt.group(1) if wgt else '400'}"
        counters[name] = counters.get(name, 0) + 1
        if counters[name] > 1:                       # e.g. italic variant, same weight
            name = f"{name}-{counters[name]}"

        def _data_uri(dm):
            payload = dm.group(2)
            ext = ".woff2" if "woff2" in dm.group(1) else ".woff"
            fp = fonts_dir / f"{name}{ext}"
            fp.write_bytes(base64.b64decode(payload))
            return f"url(fonts/{fp.name})"

        block = re.sub(r"url\(\s*['\"]?(data:[^;]+;base64),([A-Za-z0-9+/=]+)['\"]?\s*\)",
                       _data_uri, block)
        if "font-display" not in block:
            block = block.replace("@font-face{", "@font-face{font-display:swap;", 1) \
                if "@font-face{" in block else re.sub(r"(@font-face\s*\{)", r"\1font-display:swap;", block, count=1)
        return block

    return re.sub(r"@font-face\s*\{[^}]*\}", _font_face, html)


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


def pack_residency(res: dict) -> dict:
    """{uni:{scheme:{win:{'f':i,'v':[[q1..q4]|None,..]}}}} -> v as 'a,b,c,d;;...' ('' = null)."""
    if not res:
        return res
    for uni in res.values():
        for sch in uni.values():
            for rec in sch.values():
                v = rec.get("v")
                if isinstance(v, list):
                    rec["v"] = ";".join("" if row is None else ",".join(map(str, row)) for row in v)
    return {"_packed": 1, "data": res}
