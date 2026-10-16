"""Decode Power BI querydata (DSR) responses into plain tables."""
import json, re
from datetime import datetime, timezone

NULL_KEY = "\u00d8"  # 'Ø'

def _val(v):
    if isinstance(v, str):
        if v == "'-'":
            return None
        if re.fullmatch(r"-?\d+L", v):
            return int(v[:-1])
        if re.fullmatch(r"-?\d+(\.\d+)?M", v):
            return float(v[:-1])
    return v

def to_date(ms):
    if isinstance(ms, (int, float)) and ms > 10**11:
        return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).strftime("%d-%m-%Y")
    return ms

def parse(res):
    if isinstance(res, (bytes, bytearray)):
        res = res.decode("utf-8")
    if isinstance(res, str):
        res = json.loads(res)
    return res

def select_names(res):
    """Ordered display names from the descriptor (hidden/null selects dropped)."""
    data = parse(res)["results"][0]["result"]["data"]
    return [s for s in data["descriptor"]["Select"] if s]

def decode(res):
    """Return (columns, rows[list of lists], complete[bool])."""
    res = parse(res)
    data = res["results"][0]["result"]["data"]
    sel = [s for s in data["descriptor"]["Select"] if s]
    by_n = {}      # schema N (G0/M0) -> name
    for s in sel:
        by_n[s["Value"]] = s["Name"]
        for a in s.get("Subtotal", []) or []:
            by_n[a] = s["Name"]
    cols = [s["Name"] for s in sel]
    rows = []
    complete = True
    for ds in data["dsr"]["DS"]:
        vd = ds.get("ValueDicts", {})
        if ds.get("IC") is False or ("RT" in ds):
            complete = False
        for ph in ds.get("PH", []):
            for key, dm in ph.items():
                for row in dm:
                    pass
                schema = None
                prev = None
                for row in dm:
                    if "S" in row:
                        schema = row["S"]
                    if schema is None:
                        continue
                    n = len(schema)
                    c = row.get("C", [])
                    rm, nm = row.get("R", 0), row.get(NULL_KEY, 0)
                    vals, ci = [], 0
                    for i in range(n):
                        if nm >> i & 1:
                            v = None
                        elif rm >> i & 1 and prev is not None:
                            v = prev[i]
                        else:
                            v = c[ci] if ci < len(c) else None
                            ci += 1
                            dn = schema[i].get("DN")
                            if dn and isinstance(v, int) and dn in vd:
                                v = vd[dn][v]
                        vals.append(v)
                    prev = vals
                    rec = {}
                    for i, sc in enumerate(schema):
                        rec[by_n.get(sc["N"], sc["N"])] = _val(vals[i])
                    rec["_total"] = key == "DM0" and sc["N"].startswith("A")
                    rows.append(rec)
    return cols, rows, complete
