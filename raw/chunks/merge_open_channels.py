import json, os, sys
KEYS=["name","channel_type","description","reachable_investors","cost","url","how_to_apply_or_post","contacts","geography","typical_outcome","notes","sources","verified_date"]
ORDER=["chunk_reddit.json","chunk_directories.json","chunk_angel_groups.json","chunk_grants.json","chunk_communities.json"]
DROP={"Hustle Fund Angel Squad (founder access via Hustle Fund pitch form)"}  # duplicate of directories entry
out=[]; seen=set(); missing=[]
for f in ORDER:
    if not os.path.exists(f): missing.append(f); continue
    for e in json.load(open(f)):
        if e["name"] in DROP: continue
        key=e["name"].strip().lower()
        if key in seen: print("dup skipped:",e["name"]); continue
        seen.add(key)
        rec={k:e.get(k,"unknown") for k in KEYS}
        if not isinstance(rec["sources"],list): rec["sources"]=[str(rec["sources"])]
        for k in KEYS:
            if k!="sources" and not isinstance(rec[k],str): rec[k]=str(rec[k])
            if k!="sources" and rec[k].strip()=="": rec[k]="unknown"
        rec["verified_date"]="2026-09-02"
        out.append(rec)
dst="/home/kreshnith/investors-preseed/raw/open_channels_reddit_directories.json"
json.dump(out,open(dst,"w"),indent=2,ensure_ascii=False)
print("wrote",len(out),"entries to",dst,"| missing chunks:",missing)
