import json, glob, sys, os
KEYS = ["name","what_they_do","stage","check_size","sectors","geography","twitter","linkedin",
        "email_or_form","other_channels","best_channel","outreach_rule","notes","sources","verified_date"]
ARR = {"sectors","other_channels","sources"}
out, seen, problems = [], set(), []
for f in sorted(glob.glob("/home/kreshnith/investors-preseed/raw/_part_*.json")):
    try:
        data = json.load(open(f))
    except Exception as e:
        problems.append(f"{f}: invalid JSON: {e}"); continue
    for e in data:
        rec = {}
        for k in KEYS:
            v = e.get(k, "unknown")
            if k in ARR:
                if isinstance(v, str): v = [v] if v and v != "unknown" else []
                rec[k] = [str(x) for x in v]
            else:
                rec[k] = "unknown" if v in (None, "") else str(v)
        rec["verified_date"] = "2026-09-02"
        extra = set(e) - set(KEYS)
        if extra: problems.append(f"{rec['name']}: extra keys dropped {sorted(extra)}")
        key = rec["name"].lower().strip()
        if key in seen: problems.append(f"duplicate: {rec['name']}"); continue
        seen.add(key); out.append(rec)
dst = "/home/kreshnith/investors-preseed/raw/angels_twitter.json"
json.dump(out, open(dst, "w"), indent=2, ensure_ascii=False)
print("entries:", len(out))
for p in problems: print("PROBLEM:", p)
unk = lambda k: [r["name"] for r in out if r[k] == "unknown" or "(unverified)" in r[k]]
for k in ["twitter","email_or_form","check_size","linkedin"]:
    print(f"{k} unknown/unverified ({len(unk(k))}):", "; ".join(unk(k)))
