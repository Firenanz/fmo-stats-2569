"""
add_month.py — เพิ่มเดือนใหม่เข้าแดชบอร์ด TFCC แบบสคริปต์เดียว
(แทนการแก้ data/*.json + const DATA ด้วยมือทุกเดือน)

เหตุผลที่ต้องมีสคริปต์นี้ (บั๊กจริงที่เคยเกิดจากการแก้มือ):
  1. พ.ค. 2569 — ไฟล์ fee ที่ยังไม่ผ่าน 02_TFCC_ALL_PORT_FEE.py หลุดเข้าระบบ
     ทำให้รายได้เพี้ยนไปทุกทิศทาง รวมเกินจริง +1,987,872 บาท (10/14 ท่า)
  2. มิ.ย. 2569 — เพิ่มเดือนเข้า window.__EMBED__/data/*.json สำเร็จ แต่ลืม
     อัปเดต const DATA cube ของแท็บเปรียบเทียบ (คนละก้อนข้อมูล) — ผู้ใช้ต้องทัก
     มาเองถึงรู้
  3. (พบเพิ่มระหว่างทำสคริปต์นี้) const DATA cube เดือนพ.ค. ยังค้างข้อมูลชุด
     "ก่อนแก้" จากข้อ 1 อยู่ — ไม่เคยถูกแทนที่ด้วยชุดที่แก้แล้ว

สคริปต์นี้แก้ทั้ง 3 จุดด้วยสถาปัตยกรรม 2 ชั้น:
  (a) ชั้นนี้ (add_month.py): xlsx fee + xlsx บัญชี -> data/rows-2569-N.json
      + data/lookup.json(acc) + data/manifest.json  — มี assert_fee_processed
      guard และ golden-number regression check ก่อนเขียนไฟล์ใดๆ ทั้งสิ้น
  (b) build_cube.py: data/rows-*.json (ทั้งหมด) -> const DATA cube ใหม่ทั้งก้อน
      ใน index.html — รันทุกครั้งหลัง (a) เสมอ ไม่มีทาง "ลืม" เพราะเรียกอัตโนมัติ
      และเพราะ rebuild จากต้นทางทุกครั้ง ไม่มี state ค้างให้เพี้ยน

ใช้งาน:
    python3 add_month.py --fee <ไฟล์fee_Mk2.xlsx> --acc <ไฟล์เทียบบัญชี.xlsx> \
        --acc-sheet "เทียบ กรกฎาคม 2569" --ym 2569-7 --month-name กรกฎาคม --repo .

ถ้าไม่มีไฟล์บัญชีพร้อมตอนนี้ ใส่ --acc skip ได้ (จะเขียน rows.json/cube ไปก่อน
แล้วรอเติม acc ทีหลังด้วย --acc-only)
"""
import argparse, json, pathlib, sys, subprocess

HERE = pathlib.Path(__file__).parent

# ---------------------------------------------------------------------------
# 0. กฎเหล็ก — assert_fee_processed (คัดลอกจากคู่มือ คำต่อคำ ห้ามแก้ให้หลวมลง)
# ---------------------------------------------------------------------------
def assert_fee_processed(ws, filename=""):
    header = [c.value for c in ws[1]]
    header_str = " | ".join(str(h) for h in header if h is not None)

    if len(header) < 14 or not any("รายได้" in str(h) for h in header if h):
        raise ValueError(
            f"❌ ไฟล์ '{filename}' ไม่มีคอลัมน์ 'รายได้(ประมาณการ)' — "
            f"นี่คือไฟล์ดิบที่ยังไม่ผ่าน 02_TFCC_ALL_PORT_FEE.py!\n"
            f"Header ที่เจอ: {header_str}\n"
            f"ต้องรัน: python TFCC_ALL_PORT_FEE.py \"{filename}\" ก่อน แล้วเอาไฟล์ _fee_Mk2.xlsx ที่ได้มาฝังแทน"
        )

    checked, empty = 0, 0
    for r in ws.iter_rows(min_row=2, max_row=200, values_only=True):
        if r[9] is None:
            continue
        checked += 1
        if r[12] is None or str(r[12]).strip() == "":
            empty += 1
        if checked >= 20:
            break

    if checked > 0 and empty > 0:
        raise ValueError(
            f"❌ ไฟล์ '{filename}' คอลัมน์ '12.อัตราค่าธรรมเนียม' ว่าง {empty}/{checked} แถวที่สุ่มตรวจ — "
            f"นี่คือไฟล์ดิบที่ยังไม่ผ่าน 02_TFCC_ALL_PORT_FEE.py!\n"
            f"ต้องรัน: python TFCC_ALL_PORT_FEE.py \"{filename}\" ก่อน แล้วเอาไฟล์ _fee_Mk2.xlsx ที่ได้มาฝังแทน"
        )
    print(f"  ✅ Guard ผ่าน: '{filename}' เป็นไฟล์ที่ผ่าน fee script มาแล้ว ({checked} แถวสุ่มตรวจ ไม่มีช่องว่าง)")


# ---------------------------------------------------------------------------
# 1. ตัวเลขยืนยันแล้ว (golden) + เลขที่เคยใช้แล้วผิด (blocklist)
#    ที่มา: claude/master hand off stats (TFCC+Dashboard) 2570.md §18 และ CLAUDE.md
#    *** ถ้าแก้ไขค่าพวกนี้ ต้องมีเหตุผลที่ตรวจสอบได้ — ห้ามแก้เพื่อให้เช็คผ่านง่ายขึ้น ***
# ---------------------------------------------------------------------------
GOLDEN = {
    # ym_key: {"announced": ..., "acc": ..., "ports": n}
    "2569-3": {"announced": 5377658.04, "acc": 5956508.01, "ports": 13},
    "2569-4": {"announced": 5017260.24, "acc": 9221241.10, "ports": 12},
    "2569-5": {"announced": 5007184.86, "acc": None, "ports": 13},  # acc พ.ค. ไม่มีใน CLAUDE.md แยกชัด (มีแค่ตัวเลขรวม 5,024,246 ซึ่งต่างจาก rows file เล็กน้อย — ไม่ใช้เป็น golden บังคับ)
}

BLOCKLIST = {
    11428823.53: "เม.ย. — อ่านชีตบัญชีผิด (income-เดือนก่อน แทน รับ-ต่อเดือน)",
    11317076.83: "เม.ย. — ชีตผิด + ตกตราด/สุราษฎร์",
    9707777.21: "เม.ย. — ชีตถูกแต่รวม 14 ท่า (ปนตราด) ฐานไม่ตรงกับมีนา",
    9238663.00: "เม.ย. — ปนยอดสุราษฎร์ธานีของมีนาเข้ามาผิดเดือน",
    9238661.95: "เม.ย. — เดียวกับด้านบน (ทศนิยมต่าง)",
    5961518.01: "มี.ค. — รวมตราดผิด (ที่ถูกคือ 5,956,508.01 ตัดตราด)",
}


def check_against_blocklist(value, context=""):
    for bad, reason in BLOCKLIST.items():
        if abs(value - bad) < 0.01:
            raise ValueError(
                f"❌ หยุด! ค่าที่คำนวณได้ ({value:,.2f}) ตรงกับเลขที่เคยใช้แล้วผิด ({context})\n"
                f"   สาเหตุเดิม: {reason}\n"
                f"   ห้ามใช้ค่านี้ — ตรวจสอบว่าใช้ชีต/คอลัมน์ถูกหรือไม่ ก่อนทำต่อ"
            )


def check_against_golden(ym, announced_total, acc_total):
    g = GOLDEN.get(ym)
    if not g:
        return  # เดือนใหม่ (ยังไม่มี golden ให้เทียบ) — ข้ามได้ แต่ยัง check blocklist เสมอ
    if g["announced"] is not None and abs(announced_total - g["announced"]) > 0.01:
        raise ValueError(
            f"❌ หยุด! ตามประกาศ {ym} ที่คำนวณได้ ({announced_total:,.2f}) ไม่ตรงกับค่ายืนยันแล้ว ({g['announced']:,.2f})\n"
            f"   นี่คือเดือนที่เคยยืนยันค่าถูกต้องไว้แล้ว — ถ้าไม่ตรง แปลว่าสคริปต์มีจุดผิดพลาด ต้องหาสาเหตุก่อนไปต่อ"
        )
    if g["acc"] is not None and acc_total is not None and abs(acc_total - g["acc"]) > 0.01:
        raise ValueError(
            f"❌ หยุด! acc {ym} ที่คำนวณได้ ({acc_total:,.2f}) ไม่ตรงกับค่ายืนยันแล้ว ({g['acc']:,.2f})"
        )
    print(f"  ✅ Golden-number check ผ่าน: {ym} ตรงกับค่ายืนยันแล้วทุกตัว")


# ---------------------------------------------------------------------------
# 2. อ่านไฟล์ fee -> แถว rows-2569-N.json (ใช้ index ใน lookup.json เดิม)
# ---------------------------------------------------------------------------
def build_rows_from_fee(ws, lookup):
    def idx(arr, v):
        v = "" if v is None else str(v)
        if v not in arr:
            arr.append(v)
        return arr.index(v)

    rows = []
    skipped_subtotal = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[8] is None:  # คอลัมน์เครื่องมือว่าง = แถว subtotal
            skipped_subtotal += 1
            continue
        dt = r[1]
        ymd = dt.year * 10000 + dt.month * 100 + dt.day
        rows.append([
            str(r[0]), ymd,
            idx(lookup["ports"], r[2]), str(r[4]),
            idx(lookup["btypes"], r[6]), idx(lookup["bsizes"], r[7]),
            idx(lookup["gears"], r[8]), idx(lookup["fishes"], r[9]),
            float(r[11] or 0), float(r[13] or 0),
        ])
    print(f"  ข้ามแถว subtotal {skipped_subtotal} แถว, ได้ {len(rows)} แถวข้อมูลจริง")
    return rows


def sum_fee_columns(ws):
    """ผลรวมคอลัมน์ 11 (น้ำหนัก) / 13 (รายได้) ตรงจากไฟล์ fee — ไว้เทียบกับ rows ที่เขียนออกมา"""
    wt = rev = 0.0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[8] is None:
            continue
        wt += float(r[11] or 0)
        rev += float(r[13] or 0)
    return wt, rev


# ---------------------------------------------------------------------------
# 3. อ่านไฟล์บัญชี -> acc ต่อท่า (ชื่อเต็ม) + ผลรวมเทียบกับคอลัมน์รวมในชีตเดียวกัน
# ---------------------------------------------------------------------------
def build_acc_from_file(path, sheet_name):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    acc = {}
    grand_total_declared = None
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row or row[0] is None:
            continue
        label = str(row[0]).strip()
        if "รวมทั้งสิ้น" in label or "รวมส่วนกลาง" in label:
            if isinstance(row[2], (int, float)):
                grand_total_declared = float(row[2])
            continue
        if "รวม" in label:
            continue
        if not isinstance(row[2], (int, float)):
            continue
        nm = label.replace("ทร.", "ท่าเทียบเรือประมง").replace("สป.", "สะพานปลา")
        # ตัด "(1)"/"(2)" ต่อท้ายชื่อซ้ำ (เช่น สงขลา/ภูเก็ตที่มี 2 คอลัมน์บัญชี) แล้วบวกรวมเป็นท่าเดียว
        nm_clean = nm.split("(")[0].strip()
        acc[nm_clean] = acc.get(nm_clean, 0.0) + float(row[2])
    computed_total = sum(acc.values())
    if grand_total_declared is not None and abs(computed_total - grand_total_declared) > 0.01:
        raise ValueError(
            f"❌ หยุด! ผลรวม acc ที่บวกเองจากรายท่า ({computed_total:,.2f}) "
            f"ไม่ตรงกับคอลัมน์ 'รวมทั้งสิ้น/รวมส่วนกลาง' ในชีตเดียวกัน ({grand_total_declared:,.2f})\n"
            f"   ต้องหาสาเหตุก่อน (คอลัมน์ไม่ตรง/ชื่อท่าซ้ำไม่ถูก map เป็นตัวเดียวกัน ฯลฯ)"
        )
    return acc, computed_total


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fee", required=True, help="ไฟล์ fee xlsx (ต้องผ่าน 02_TFCC_ALL_PORT_FEE.py แล้ว)")
    ap.add_argument("--acc", help="ไฟล์เทียบบัญชี xlsx (ใส่ 'skip' ถ้ายังไม่มี)")
    ap.add_argument("--acc-sheet", help="ชื่อชีตในไฟล์บัญชี เช่น 'เทียบ กรกฎาคม 2569'")
    ap.add_argument("--ym", required=True, help="เช่น 2569-7")
    ap.add_argument("--month-name", required=True, help="เช่น กรกฎาคม")
    ap.add_argument("--repo", default=".", help="path ของ repo (มี data/, index.html)")
    ap.add_argument("--dry-run", action="store_true", help="เช็คทุกอย่างแต่ไม่เขียนไฟล์จริง")
    args = ap.parse_args()

    import openpyxl
    repo = pathlib.Path(args.repo)

    lookup_path = repo / "data" / "lookup.json"
    manifest_path = repo / "data" / "manifest.json"
    rows_path = repo / "data" / f"rows-{args.ym}.json"

    lookup = json.loads(lookup_path.read_text(encoding="utf-8"))
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

    if args.ym in manifest["months"]:
        raise SystemExit(f"❌ เดือน {args.ym} มีอยู่แล้วใน manifest.json — ถ้าต้องการแก้เดือนเดิม ต้องลบของเก่าออกก่อนด้วยมือ (ไม่ใช่ flow ปกติของสคริปต์นี้)")

    print(f"=== ขั้นที่ 1: ตรวจไฟล์ fee ({args.fee}) ===")
    wb = openpyxl.load_workbook(args.fee, data_only=True)
    ws = wb[wb.sheetnames[0]]
    assert_fee_processed(ws, args.fee)

    print(f"=== ขั้นที่ 2: แปลงเป็น rows + ตรวจผลรวม ===")
    rows = build_rows_from_fee(ws, lookup)
    fee_wt, fee_rev = sum_fee_columns(ws)
    rows_wt = sum(r[8] for r in rows)
    rows_rev = sum(r[9] for r in rows)
    if abs(rows_wt - fee_wt) > 0.01 or abs(rows_rev - fee_rev) > 0.01:
        raise ValueError(
            f"❌ หยุด! ผลรวมใน rows ที่สร้าง (wt={rows_wt:,.2f}, rev={rows_rev:,.2f}) "
            f"ไม่ตรงกับผลรวมในไฟล์ fee ต้นทาง (wt={fee_wt:,.2f}, rev={fee_rev:,.2f}) เป๊ะ"
        )
    print(f"  ✅ น้ำหนักรวม {rows_wt:,.2f} กก. / รายได้ตามประกาศรวม {rows_rev:,.2f} บาท — ตรงกับไฟล์ fee เป๊ะ")

    check_against_blocklist(rows_rev, f"{args.ym} ตามประกาศ")
    check_against_golden(args.ym, rows_rev, None)

    acc = None
    acc_total = None
    if args.acc and args.acc != "skip":
        print(f"=== ขั้นที่ 3: อ่านไฟล์บัญชี ({args.acc}, ชีต '{args.acc_sheet}') ===")
        if not args.acc_sheet:
            raise SystemExit("❌ ต้องระบุ --acc-sheet ด้วยถ้าจะอ่านไฟล์บัญชี")
        acc, acc_total = build_acc_from_file(args.acc, args.acc_sheet)
        check_against_blocklist(acc_total, f"{args.ym} acc")
        check_against_golden(args.ym, rows_wt and rows_rev, acc_total)
        print(f"  ✅ acc รวม {acc_total:,.2f} บาท ({len(acc)} ท่า) — ตรงกับคอลัมน์รวมในชีตเดียวกัน")
    else:
        print("=== ขั้นที่ 3: ข้าม (ไม่มีไฟล์บัญชีตอนนี้ — acc จะว่างไว้ก่อน ต้องรันเติมทีหลัง) ===")

    if args.dry_run:
        print("\n--dry-run: ไม่เขียนไฟล์จริง จบการตรวจที่นี่")
        return

    print(f"=== ขั้นที่ 4: เขียนไฟล์ ===")
    rows_path.write_text(json.dumps(rows, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"  เขียน {rows_path}")

    if acc is not None:
        lookup.setdefault("acc", {})[args.ym] = acc
    lookup_path.write_text(json.dumps(lookup, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"  อัปเดต {lookup_path}")

    manifest["months"].append(args.ym)
    manifest["months"].sort(key=lambda s: tuple(int(x) for x in s.split("-")))
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"  อัปเดต {manifest_path}")

    print(f"=== ขั้นที่ 5: rebuild DATA cube (compare tab) จาก rows ทั้งหมด ===")
    r = subprocess.run([sys.executable, str(HERE / "build_cube.py"), str(repo / "index.html")],
                        capture_output=True, text=True)
    print("  " + r.stdout.strip())
    if r.returncode != 0:
        raise SystemExit("❌ build_cube.py ล้มเหลว:\n" + r.stderr)

    print(f"\n✅ เสร็จสิ้น — เพิ่มเดือน {args.month_name} {args.ym} แล้ว "
          f"({len(rows)} แถว, น้ำหนัก {rows_wt:,.2f} กก., รายได้ตามประกาศ {rows_rev:,.2f} บาท"
          + (f", acc {acc_total:,.2f} บาท" if acc_total else ", ยังไม่มี acc") + ")")
    print("ขั้นต่อไป: เปิด index.html ด้วย Playwright เช็ค JS error + เทียบตัวเลขบนหน้าเว็บจริงก่อน commit/push")


if __name__ == "__main__":
    main()
