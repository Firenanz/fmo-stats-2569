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
    python3 add_month.py --fee <ไฟล์fee_Mk2.xlsx> --acc <ไฟล์บัญชีรายเดือน.xlsx> \
        --ym 2569-7 --month-name กรกฎาคม --repo .
    (--acc-sheet ไม่ต้องใส่ก็ได้ ดีฟอลต์ "รับ-ต่อเดือน" ซึ่งท่าเป็นคอลัมน์ ไม่ใช่แถว)

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
import datetime as _dt


def _parse_ymd(v):
    """รองรับทั้ง datetime object (เดือนก่อนๆ) และ string 'dd/mm/พ.ศ.' (เจอในไฟล์ ก.ค. 2569)"""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.year * 10000 + v.month * 100 + v.day
    s = str(v).strip()
    d, m, y = s.split("/")
    return int(y) * 10000 + int(m) * 100 + int(d)


def _to_float(v):
    """รองรับทั้งตัวเลข native และ string มี comma คั่นหลักพัน (เจอในคอลัมน์น้ำหนักไฟล์ ก.ค. 2569)"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    return float(s) if s else 0.0


# เรือเข้าท่าปลายเดือน แต่ขนถ่ายเสร็จ (วันที่ที่ TFCC บันทึก) ข้ามไปต้นเดือนถัดไป 1-7 วัน
# ถือเป็นเรื่องปกติ/ตั้งใจของระบบ TFCC เอง (ดู 01_TFCC_ALL_PORTS_Mk34.py TOLERANCE_DAYS=7,
# ยืนยันซ้ำใน master handoff doc 5 พ.ย. 2569) — ไฟล์เดือนไหนก็นับเป็นเดือนนั้น ไม่ใช่นับตาม
# วันที่ขนถ่ายเสร็จที่หลุดออกไปเดือนถัดไป ***ห้ามแก้ตรรกะนี้โดยไม่ถามผู้ใช้ก่อน***
_TOLERANCE_DAYS = 7


def _clamp_to_target_month(ymd_val, target_y, target_m):
    """ถ้าวันที่หลุดไปเดือนอื่น (แต่ไม่เกิน _TOLERANCE_DAYS วันจากขอบเดือนเป้าหมาย) ให้ปัดเข้า
    เดือนเป้าหมาย (ตัวสุดท้าย/แรกของเดือน) — คืนค่า (ymd_ใหม่, ถูกปรับหรือไม่)"""
    y, m = ymd_val // 10000, (ymd_val // 100) % 100
    if (y, m) == (target_y, target_m):
        return ymd_val, False
    d_actual = _dt.date(y - 543, m, ymd_val % 100)
    last_day = _dt.date(target_y - 543, target_m % 12 + 1, 1) - _dt.timedelta(days=1) \
        if target_m < 12 else _dt.date(target_y - 543 + 1, 1, 1) - _dt.timedelta(days=1)
    first_day = _dt.date(target_y - 543, target_m, 1)
    if d_actual > last_day and (d_actual - last_day).days <= _TOLERANCE_DAYS:
        return target_y * 10000 + target_m * 100 + last_day.day, True
    if d_actual < first_day and (first_day - d_actual).days <= _TOLERANCE_DAYS:
        return target_y * 10000 + target_m * 100 + 1, True
    raise ValueError(
        f"❌ หยุด! พบวันที่ {d_actual} ซึ่งไม่ใช่เดือนเป้าหมาย ({target_y}-{target_m:02d}) "
        f"และห่างเกิน {_TOLERANCE_DAYS} วัน (tolerance ของ TFCC เอง) — ไม่ใช่เคสเรือขนถ่ายข้ามเดือนปกติ "
        f"ต้องตรวจสอบไฟล์ก่อน (อาจมีแถวปีก่อน/ไฟล์ปนเดือนอื่นหลุดเข้ามา)"
    )


def build_rows_from_fee(ws, lookup, target_ym=None):
    def idx(arr, v):
        v = "" if v is None else str(v)
        if v not in arr:
            arr.append(v)
        return arr.index(v)

    target_y = target_m = None
    if target_ym:
        target_y, target_m = (int(x) for x in target_ym.split("-"))

    rows = []
    skipped_subtotal = 0
    clamped = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[8] is None:  # คอลัมน์เครื่องมือว่าง = แถว subtotal
            skipped_subtotal += 1
            continue
        ymd = _parse_ymd(r[1])
        if target_y is not None:
            ymd, was_clamped = _clamp_to_target_month(ymd, target_y, target_m)
            if was_clamped:
                clamped += 1
        rows.append([
            str(r[0]), ymd,
            idx(lookup["ports"], r[2]), str(r[4]),
            idx(lookup["btypes"], r[6]), idx(lookup["bsizes"], r[7]),
            idx(lookup["gears"], r[8]), idx(lookup["fishes"], r[9]),
            _to_float(r[11]), _to_float(r[13]),
        ])
    print(f"  ข้ามแถว subtotal {skipped_subtotal} แถว, ได้ {len(rows)} แถวข้อมูลจริง")
    if clamped:
        print(f"  ℹ️  ปัดวันที่เข้าเดือน {target_ym} ให้ {clamped} แถว "
              f"(เรือเข้าปลายเดือนแต่ขนถ่ายเสร็จข้ามเดือน ≤{_TOLERANCE_DAYS} วัน — ตรงกับ tolerance ของ 01_TFCC_ALL_PORTS)")
    return rows


def sum_fee_columns(ws):
    """ผลรวมคอลัมน์ 11 (น้ำหนัก) / 13 (รายได้) ตรงจากไฟล์ fee — ไว้เทียบกับ rows ที่เขียนออกมา"""
    wt = rev = 0.0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[8] is None:
            continue
        wt += _to_float(r[11])
        rev += _to_float(r[13])
    return wt, rev


# ---------------------------------------------------------------------------
# 3. อ่านไฟล์บัญชี -> acc ต่อท่า (ชื่อเต็ม) + ผลรวมเทียบกับคอลัมน์รวมในชีตเดียวกัน
#
# *** โครงสร้างจริงของชีต "รับ-ต่อเดือน" (ยืนยันจากไฟล์จริง พ.ย. 2569) ***
# ท่าเป็น "คอลัมน์" ไม่ใช่แถว! แถว header (มีคำว่า "รวมส่วนกลาง") อยู่แถว ~4
# แถวข้อมูลคือแถวที่คอลัมน์ B == "ค่าธรรมเนียมสัตว์น้ำผ่านท่า"
# คอลัมน์ไม่ใช่ท่า TFCC ทั้งหมด (ส่วนกลาง/สป.กรุงเทพ/ทร.หัวหิน/ทร.ปราณบุรี/
# ทร.ฉะเชิงเทรา ฯลฯ) ถูกตัดทิ้งโดยตั้งใจเพราะไม่อยู่ใน ACC_PORT_MAP ด้านล่าง
# ***เดิมฟังก์ชันนี้เขียนโดยสมมติว่าท่าเป็น "แถว" (ผิด) — แก้ พ.ย. 2569 ก่อนใช้จริงครั้งแรก***
# ---------------------------------------------------------------------------
import re as _re

ACC_PORT_MAP = {
    "สป.สป.": "สะพานปลาสมุทรปราการ",
    "สป.สมุทรสาคร": "สะพานปลาสมุทรสาคร",
    "ทร.ตราด": "ท่าเทียบเรือประมงตราด",
    "ทร.อ่างศิลา": "ท่าเทียบเรือประมงอ่างศิลา",
    "ทร.ชุมพร": "ท่าเทียบเรือประมงชุมพร",
    "ทร.หลังสวน": "ท่าเทียบเรือประมงหลังสวน",            # ตัดเลขท้าย (เช่น "13") ก่อน lookup
    "ทร.สุราษฎร์ฯ": "ท่าเทียบเรือประมงสุราษฏร์ธานี",      # ตัดเลขท้ายก่อน lookup (สะกด ฎ แต่ของจริงใช้ ฏ)
    "ทร.นครฯ": "สะพานปลานครศรีธรรมราช",                  # หัวตารางขึ้น "ทร." แต่ชื่อจริงเป็น "สะพานปลา"
    "ทร.สงขลา": "ท่าเทียบเรือประมงสงขลา",                # มี (1)/(2) สองคอลัมน์ บวกรวมเป็นท่าเดียว
    "ทรสงขลา": "ท่าเทียบเรือประมงสงขลา",                 # สะกดไม่มีจุดก็เจอในไฟล์จริง เผื่อไว้
    "ทร.ปัตตานี": "ท่าเทียบเรือประมงปัตตานี",
    "ทร.นราธิวาส": "ท่าเทียบเรือประมงนราธิวาส",
    "ทร.ระนอง": "ท่าเทียบเรือประมงระนอง",
    "ทร.ภูเก็ต": "ท่าเทียบเรือประมงภูเก็ต",              # ขึ้นซ้ำ 2 คอลัมน์ในไฟล์จริง บวกรวมเป็นท่าเดียว
    "ทร.สตูล": "ท่าเทียบเรือประมงสตูล",
}


def _clean_acc_header(h):
    h = str(h).strip()
    h = _re.sub(r"\(\d+\)$", "", h)   # ตัด "(1)"/"(2)" ท้าย
    h = _re.sub(r"\d+$", "", h)       # ตัดเลขท้าย เช่น "13"/"14"/" 15"
    return h.strip()


def build_acc_from_file(path, sheet_name="รับ-ต่อเดือน",
                         row_label="ค่าธรรมเนียมสัตว์น้ำผ่านท่า",
                         total_col_label="รวมส่วนกลาง"):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]

    header_row = None
    for r in range(1, 10):
        if any(ws.cell(row=r, column=c).value == total_col_label
               for c in range(1, ws.max_column + 1)):
            header_row = r
            break
    if header_row is None:
        raise ValueError(f"❌ ไม่เจอ header row ที่มีคอลัมน์ '{total_col_label}' ในชีต '{sheet_name}'")
    headers = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]

    data_row = None
    for r in range(header_row + 1, ws.max_row + 1):
        if str(ws.cell(row=r, column=2).value or "").strip() == row_label:
            data_row = r
            break
    if data_row is None:
        raise ValueError(f"❌ ไม่เจอแถว '{row_label}' ในชีต '{sheet_name}'")

    acc = {}
    declared_total = None
    for c, h in enumerate(headers, 1):
        if h is None:
            continue
        v = ws.cell(row=data_row, column=c).value
        if str(h).strip() == total_col_label:
            declared_total = float(v or 0)
            continue
        key = _clean_acc_header(h)
        if key not in ACC_PORT_MAP:
            continue  # ไม่ใช่ท่า TFCC — ตัดทิ้งโดยตั้งใจ (ส่วนกลาง/กรุงเทพ/หัวหิน/ปราณบุรี/ฉะเชิงเทรา ฯลฯ)
        port = ACC_PORT_MAP[key]
        acc[port] = acc.get(port, 0.0) + float(v or 0)

    computed_total = sum(acc.values())
    if declared_total is not None and abs(computed_total - declared_total) > 0.01:
        raise ValueError(
            f"❌ หยุด! ผลรวม acc ที่บวกเองจากรายท่า ({computed_total:,.2f}) "
            f"ไม่ตรงกับคอลัมน์ '{total_col_label}' ในชีตเดียวกัน ({declared_total:,.2f})\n"
            f"   ต้องหาสาเหตุก่อน (ท่าใหม่ที่ยังไม่อยู่ใน ACC_PORT_MAP / คอลัมน์ขยับ ฯลฯ)"
        )
    return acc, computed_total


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fee", required=True, help="ไฟล์ fee xlsx (ต้องผ่าน 02_TFCC_ALL_PORT_FEE.py แล้ว)")
    ap.add_argument("--acc", help="ไฟล์บัญชีรายเดือน xlsx (ใส่ 'skip' ถ้ายังไม่มี)")
    ap.add_argument("--acc-sheet", default="รับ-ต่อเดือน",
                     help="ชื่อชีตในไฟล์บัญชี (ปกติคือ 'รับ-ต่อเดือน' — ท่าเป็นคอลัมน์ ไม่ใช่แถว)")
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
    rows = build_rows_from_fee(ws, lookup, target_ym=args.ym)
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
