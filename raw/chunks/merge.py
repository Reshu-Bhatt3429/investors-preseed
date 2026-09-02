import json, glob, sys
KEYS = ["name","type","stage","check_size_terms","sectors","geography","website",
        "application_url","application_timing","email","twitter","partner_twitters",
        "linkedin","best_channel","notes","sources","verified_date"]
out, seen, problems = [], set(), []
for f in sorted(glob.glob("/home/kreshnith/investors-preseed/raw/chunks/chunk*.json")):
    try:
        data = json.load(open(f))
    except Exception as e:
        problems.append(f"{f}: JSON error {e}"); continue
    for e in data:
        rec = {}
        for k in KEYS:
            v = e.get(k, "unknown")
            if k in ("partner_twitters","sources"):
                if isinstance(v, str): v = [v] if v and v!="unknown" else []
                if v is None: v = []
            else:
                if v is None or (isinstance(v,str) and not v.strip()): v = "unknown"
                if isinstance(v, list): v = "; ".join(map(str,v))
            rec[k] = v
        rec["verified_date"] = "2026-09-02"
        extra = set(e) - set(KEYS)
        if extra: problems.append(f"{rec['name']}: extra keys {extra}")
        key = rec["name"].strip().lower()
        if key in seen: problems.append(f"duplicate: {rec['name']}"); continue
        seen.add(key); out.append(rec)
json.dump(out, open("/home/kreshnith/investors-preseed/raw/accelerators_and_funds.json","w"), indent=2, ensure_ascii=False)
print("entries:", len(out))
for p in problems: print("PROBLEM:", p)
unk = {k: sum(1 for r in out if r[k]=="unknown" or r[k]==[]) for k in KEYS}
print("unknown counts:", unk)
