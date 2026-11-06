"""
build_cube.py — Rebuild the `const DATA={...}` compare-tab cube (index.html)
FROM SCRATCH out of data/lookup.json + data/manifest.json + data/rows-2569-N.json.

Field meanings confirmed empirically from the live cube + JS usage (index.html,
xpJS module, grep for "DATA.reg"/"DATA.dld"):
  - DATA.reg[i] : index into regList/regName — the VESSEL (boat) for row i.
    (946 distinct values in the live cube, matches len(regList)=1023 range —
     this is the array that was left truncated at 37,643 after June was added
     to col/wt/rev but never to reg.)
  - DATA.dld[i] : index into an internal, unlisted set of DELIVERY-DOCUMENT
    (trip) ids — used only via Set() cardinality (DATA.dld[i] is never
    displayed), so a fresh renumbering on each rebuild is safe as long as
    rows sharing one DLD string get the same index.

Why rebuild instead of append: the cube had drifted from the real data in two
independent ways (June: `reg` array never extended; May: still held the
pre-fee-fix numbers — confirmed: cube's May ปัตตานี/ปลาโอดำ rate came out to
2.00 บาท/กก., the known-wrong rate, where data/rows-2569-5.json has the
corrected 0.60). A full rebuild from the same source-of-truth files that
drive the main dashboard removes both bug classes at once and cannot drift
again, since the cube keeps no state of its own between rebuilds.

Usage: python3 build_cube.py <path-to-index.html>
Writes the file in place (only the `const DATA=...;` blob is touched).
"""
import json, sys, pathlib

REPO = pathlib.Path(sys.argv[1]).parent if len(sys.argv) > 1 else pathlib.Path(".")
INDEX = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "index.html"

TH_MONTHS = ["", "มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
             "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

def port_short(full, cube_cats_so_far):
    """cube's short port name is always a substring of lookup.json's full name."""
    for c in cube_cats_so_far:
        if c in full:
            return c
    return full.replace("ท่าเทียบเรือประมง", "").replace("สะพานปลา", "")

BSIZE_MAP = {
    "ขนาด L (60–149GT)": "L (60–149GT)",
    "ขนาด S  (10–29 GT)": "S (10–29GT)",
    "ขนาด M ( 30–59 GT)": "M (30–59GT)",
    "ขนาด XL (มากกว่า 150GT)": "XL (มากกว่า 150GT)",
}

def main():
    lookup = json.loads((REPO / "data" / "lookup.json").read_text(encoding="utf-8"))
    manifest = json.loads((REPO / "data" / "manifest.json").read_text(encoding="utf-8"))
    months = manifest["months"]

    html = INDEX.read_text(encoding="utf-8")
    i = html.find("const DATA=")
    if i == -1:
        raise SystemExit("const DATA= not found in " + str(INDEX))
    start = i + len("const DATA=")
    decoder = json.JSONDecoder()
    _, end_offset = decoder.raw_decode(html[start:])
    end = start + end_offset
    old_cube = json.loads(html[start:end])
    old_regname_by_reg = dict(zip(old_cube["regList"], old_cube["regName"]))

    dims = {k: {"cats": list(v["cats"])} for k, v in old_cube["dims"].items()}
    dim_labels = old_cube["dimLabels"]

    def idx_append(cats, v):
        if v not in cats:
            cats.append(v)
        return cats.index(v)

    # seed port -> region map from the OLD cube so region assignment for
    # already-known ports never changes on rebuild
    port_region = {}
    for i_row in range(old_cube["n"]):
        pr = dims["port"]["cats"][old_cube["col"]["port"][i_row]]
        rg = dims["region"]["cats"][old_cube["col"]["region"][i_row]]
        port_region.setdefault(pr, rg)

    col = {k: [] for k in dims}
    wt, rev = [], []
    reg_arr, dld_arr = [], []
    regList, regName = [], []
    reg_index = {}
    dld_index = {}

    n = 0
    for ym in months:
        y, m = ym.split("-")
        mon_label = f"{TH_MONTHS[int(m)]} {y}"
        mon_idx = idx_append(dims["mon"]["cats"], mon_label)

        rows = json.loads((REPO / "data" / f"rows-{ym}.json").read_text(encoding="utf-8"))
        for r in rows:
            dld_s, ymd_i, port_i, regnum, btype_i, bsize_i, gear_i, fish_i, w, rv = r

            port_full = lookup["ports"][port_i]
            port_s = port_short(port_full, dims["port"]["cats"])
            port_idx = idx_append(dims["port"]["cats"], port_s)
            region = port_region.get(port_s, "อ่าวไทย")
            region_idx = idx_append(dims["region"]["cats"], region)
            bsize_full = lookup["bsizes"][bsize_i]
            bsize_s = BSIZE_MAP.get(bsize_full, bsize_full)
            bsize_idx = idx_append(dims["bsize"]["cats"], bsize_s)
            btype_idx = idx_append(dims["btype"]["cats"], lookup["btypes"][btype_i])
            gear_idx = idx_append(dims["gear"]["cats"], lookup["gears"][gear_i])
            fish_idx = idx_append(dims["fish"]["cats"], lookup["fishes"][fish_i])

            col["mon"].append(mon_idx)
            col["region"].append(region_idx)
            col["port"].append(port_idx)
            col["bsize"].append(bsize_idx)
            col["btype"].append(btype_idx)
            col["gear"].append(gear_idx)
            col["fish"].append(fish_idx)

            reg_s = str(regnum)
            ridx = reg_index.get(reg_s)
            if ridx is None:
                name = old_regname_by_reg.get(reg_s, reg_s)
                regList.append(reg_s); regName.append(name)
                ridx = len(regList) - 1
                reg_index[reg_s] = ridx
            reg_arr.append(ridx)

            didx = dld_index.get(dld_s)
            if didx is None:
                didx = len(dld_index)
                dld_index[dld_s] = didx
            dld_arr.append(didx)

            wt.append(float(w)); rev.append(float(rv))
            n += 1

    new_cube = {
        "n": n,
        "dimLabels": dim_labels,
        "dims": dims,
        "col": col,
        "reg": reg_arr,
        "regList": regList,
        "regName": regName,
        "dld": dld_arr,
        "wt": wt,
        "rev": rev,
    }

    newblob = json.dumps(new_cube, ensure_ascii=False, separators=(",", ":"))
    tail = html[end+1:] if html[end:end+1] == ";" else html[end:]
    html2 = html[:i] + "const DATA=" + newblob + ";" + tail
    INDEX.write_text(html2, encoding="utf-8")
    print(f"OK: wrote n={n} rows, {len(regList)} boats, {len(dld_index)} trips, months={months}")
    return new_cube, old_cube

if __name__ == "__main__":
    main()
