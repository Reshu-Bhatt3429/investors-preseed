import json, glob
KEYS = ["name","type","sectors","stage","check_size","geography","website","pitch_email_or_form",
        "twitter","partner_twitters","linkedin","best_channel","cold_pitch_policy","notes","sources","verified_date"]
files = sorted(glob.glob("/home/kreshnith/investors-preseed/raw/chunks/[A-G]_*.json"))
out, seen, problems = [], set(), []
for f in files:
    try:
        data = json.load(open(f))
    except Exception as e:
        problems.append(f"{f}: JSON error {e}"); continue
    for e in data:
        rec = {}
        for k in KEYS:
            v = e.get(k, "unknown")
            if v is None or (isinstance(v, str) and not v.strip()): v = "unknown"
            if isinstance(v, list): v = ", ".join(map(str, v))
            rec[k] = str(v)
        rec["verified_date"] = "2026-09-02"
        extra = set(e) - set(KEYS)
        missing = set(KEYS) - set(e)
        if extra: problems.append(f"{f}:{rec['name']}: extra keys {extra}")
        if missing: problems.append(f"{f}:{rec['name']}: missing keys {missing}")
        key = rec["name"].strip().lower()
        if key in seen: problems.append(f"duplicate: {rec['name']}"); continue
        seen.add(key); out.append(rec)
    print(f, len(data))
json.dump(out, open("/home/kreshnith/investors-preseed/raw/fintech_crypto_ai_vcs.json","w"), indent=2, ensure_ascii=False)
print("entries:", len(out))
for p in problems: print("PROBLEM:", p)
unk = {k: sum(1 for r in out if r[k]=="unknown") for k in KEYS}
print("unknown counts:", unk)
