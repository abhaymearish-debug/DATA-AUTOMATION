"""Re-parse each rebuilt PDF and diff it against its source, block by block."""
import sys, json
sys.path.insert(0,".")
from extract_target import parse
U="/sessions/gifted-adoring-hypatia/mnt/uploads/"
pairs=[("achieved_target_cluster_-_1-4.pdf","TARGET vs ACHIEVED - CLUSTER 1 - MOBILE.pdf"),
       ("achieved_target_cluster_-_2-4.pdf","TARGET vs ACHIEVED - CLUSTER 2 - MOBILE.pdf"),
       ("achieved_target_cluster_-_3-4.pdf","TARGET vs ACHIEVED - CLUSTER 3 - MOBILE.pdf"),
       ("achieved_target_clusters_summary-4.pdf","TARGET vs ACHIEVED - SUMMARY - MOBILE.pdf")]
allfail=[]; cells=0
for src,new in pairs:
    a=parse(U+src); b=parse(new)
    import importlib.util as _il
    _s=_il.spec_from_file_location("bt","build_target_mobile.py")
    _m=_il.module_from_spec(_s); _s.loader.exec_module(_m)
    f=[]
    if a["title"]!=b["title"]: f.append(f"title {a['title']!r} vs {b['title']!r}")
    if _m.band_title(a["band_left"])!=b["band_left"]:
        f.append(f"band L {_m.band_title(a['band_left'])!r} vs {b['band_left']!r}")
    if a["band_right"]!=b["band_right"]: f.append(f"band R {a['band_right']!r} vs {b['band_right']!r}")
    # headers are deliberately shortened in the rebuild, so compare the SOURCE
    # label through the same map rather than expecting it verbatim
    import importlib.util as _il
    _s=_il.spec_from_file_location("bt","build_target_mobile.py")
    _m=_il.module_from_spec(_s); _s.loader.exec_module(_m)
    ha=[_m.header_text(x) for x in a["headers"]]; hb=[" ".join(x) for x in b["headers"]]
    if ha!=hb: f.append(f"headers differ:\n      {ha}\n      {hb}")
    if len(a["blocks"])!=len(b["blocks"]): f.append(f"block count {len(a['blocks'])} vs {len(b['blocks'])}")
    summary = any(z.get("kind")=="grand" for z in a["blocks"])
    rm = {"total":"bond","grand":"total"} if summary else {}
    for x,y in zip(a["blocks"],b["blocks"]):
        want_name=" ".join(x["name"]).replace("CLUSTER - ", "CLUSTER ")
        if rm.get(x.get("kind"),x.get("kind"))=="bond" and want_name.upper().endswith(" TOTAL"):
            want_name=want_name[:-6].rstrip()
        if want_name.replace(" ","")!=" ".join(y["name"]).replace(" ",""):
            f.append(f"name {want_name!r} vs {' '.join(y['name'])!r}")
        if x["tgt"]!=y["tgt"]: f.append(f"{x['name']} TGT {x['tgt']} vs {y['tgt']}")
        if x["ach"]!=y["ach"]: f.append(f"{x['name']} ACH {x['ach']} vs {y['ach']}")
        if x["pct"]!=y["pct"]: f.append(f"{x['name']} pct {x['pct']} vs {y['pct']}")
        # the summary sheet deliberately remaps its bands, so compare the SOURCE
        # kind through that same mapping rather than expecting equality
        want = rm.get(x.get("kind"), x.get("kind"))
        if want != y.get("kind"):
            f.append(f"{x['name']} band {x.get('kind')}->{want} but got {y.get('kind')}")
        cells += len(x["tgt"])+len(x["ach"])+1
    print(f"{src:<42} {'OK' if not f else str(len(f))+' FAIL'}   blocks {len(a['blocks'])}")
    for line in f[:4]: print("      ", line)
    allfail += f
print(f"\ncells compared: {cells}")
print("ALL FOUR MATCH" if not allfail else f"{len(allfail)} FAILURES")
