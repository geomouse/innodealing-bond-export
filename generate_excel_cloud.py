#!/usr/bin/env python3
# 债立方信用债数据汇总 - 云端版
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
import glob
import os
import re
import sys
import warnings
warnings.filterwarnings('ignore')

TARGET_HEADERS = [
    '截标时间', '债券简称', '债券代码', '发行期限', '计划发行(亿)', '投标区间(%)', '票面利率(%)',
    '发行人', '区域', '相似二级券', '相似二级剩余期限', '相似二级估值', 'YY评分',
    '主体评级', '债项评级', '主承销商', '担保人', '债券类型', '发行人类型', '交易市场',
    '债券全称', '募集方式', '实际到期日', '到期日休市情况', '上市日', '公告日', '缴款日',
    '板块分类', 'DM评分', 'DM一级行业', 'DM二级行业', 'DM三级行业', '申万行业',
    '行权日', '是否有担保', '条款', '偿付顺序', '发行截止日', '团费(%)'
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
TODAY = date.today().strftime('%Y-%m-%d')
OUTPUT_PATH = os.path.join(DATA_DIR, f"信用债发行汇总_{TODAY}.xlsx")

_files = glob.glob(os.path.join(DATA_DIR, "credit_bond_*.xlsx"))
DATES = sorted([
    re.search(r'credit_bond_(\d{4}-\d{2}-\d{2})\.xlsx', f).group(1)
    for f in _files
    if re.search(r'credit_bond_(\d{4}-\d{2}-\d{2})\.xlsx', f)
])

if not DATES:
    print("未找到任何 credit_bond_*.xlsx 文件")
    sys.exit(1)

print(f"检测到 {len(DATES)} 个日期文件: {DATES[0]} ~ {DATES[-1]}")

# ===== 读取所有日期的数据 =====
all_data = {}
for dt in DATES:
    filepath = os.path.join(DATA_DIR, f"credit_bond_{dt}.xlsx")
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    source_headers = [cell.value for cell in ws[2]]
    header_map = {}
    for i, h in enumerate(source_headers):
        if h and h not in header_map:
            header_map[h] = i
    rows = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if not row[0]:
            continue
        row_dict = {}
        for col_name, idx in header_map.items():
            row_dict[col_name] = row[idx] if idx < len(row) else None
        rows.append(row_dict)
    all_data[dt] = rows
    print(f"  {dt}: {len(rows)} 条")

# ===== 今天数据 =====
today_rows = all_data.get(TODAY, [])
print(f"\n今天({TODAY})数据: {len(today_rows)} 条")

# ===== 汇总 =====
summary_bonds = {}
for dt in DATES:
    for row in all_data[dt]:
        bond_name = row.get('债券简称')
        if not bond_name:
            continue
        if bond_name not in summary_bonds:
            summary_bonds[bond_name] = {
                'data': {}, 'first_seen': dt,
                'dates_seen': [dt], 'latest_date': dt
            }
        else:
            if dt not in summary_bonds[bond_name]['dates_seen']:
                summary_bonds[bond_name]['dates_seen'].append(dt)
            if dt > summary_bonds[bond_name]['latest_date']:
                summary_bonds[bond_name]['latest_date'] = dt
        for key, value in row.items():
            if value is not None and value != '' and value != 0:
                if key in ('债券代码', '票面利率(%)'):
                    summary_bonds[bond_name]['data'][key] = value
                else:
                    if dt >= summary_bonds[bond_name]['latest_date'] or key not in summary_bonds[bond_name]['data']:
                        summary_bonds[bond_name]['data'][key] = value
            elif key not in summary_bonds[bond_name]['data']:
                summary_bonds[bond_name]['data'][key] = value

total_bonds = len(summary_bonds)
rate_filled = sum(1 for v in summary_bonds.values() if v['data'].get('票面利率(%)'))
print(f"\n汇总唯一债券数: {total_bonds} (票面利率补全: {rate_filled}/{total_bonds})")

# ===== 创建输出Excel =====
wb_out = openpyxl.Workbook()
header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
data_font = Font(name='Arial', size=10)
data_alignment = Alignment(horizontal='left', vertical='center')
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
today_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

# Sheet1: 当天提取
ws1 = wb_out.active
ws1.title = f"当天提取_{TODAY}"
for col_idx, header in enumerate(TARGET_HEADERS, 1):
    cell = ws1.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = thin_border
for row_idx, row_data in enumerate(today_rows, 2):
    for col_idx, header in enumerate(TARGET_HEADERS, 1):
        value = row_data.get(header)
        cell = ws1.cell(row=row_idx, column=col_idx, value=value if value is not None else '')
        cell.font = data_font
        cell.alignment = data_alignment
        cell.border = thin_border

col_widths = {
    '截标时间': 12, '债券简称': 18, '债券代码': 16, '发行期限': 8, '计划发行(亿)': 10,
    '投标区间(%)': 14, '票面利率(%)': 10, '发行人': 28, '区域': 10, '相似二级券': 16,
    '相似二级剩余期限': 12, '相似二级估值': 12, 'YY评分': 8, '主体评级': 8, '债项评级': 8,
    '主承销商': 30, '担保人': 12, '债券类型': 10, '发行人类型': 12, '交易市场': 8,
    '债券全称': 40, '募集方式': 8, '实际到期日': 12, '到期日休市情况': 10, '上市日': 12,
    '公告日': 12, '缴款日': 12, '板块分类': 8, 'DM评分': 8, 'DM一级行业': 10,
    'DM二级行业': 10, 'DM三级行业': 10, '申万行业': 10, '行权日': 12, '是否有担保': 8,
    '条款': 8, '偿付顺序': 8, '发行截止日': 12, '团费(%)': 8
}
for col_idx, header in enumerate(TARGET_HEADERS, 1):
    ws1.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(header, 12)
ws1.row_dimensions[1].height = 30
ws1.freeze_panes = 'A2'

# Sheet2: 汇总
ws2 = wb_out.create_sheet(title="汇总")
SUMMARY_HEADERS = ['首次提取日期'] + TARGET_HEADERS
for col_idx, header in enumerate(SUMMARY_HEADERS, 1):
    cell = ws2.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = thin_border

sorted_bonds = sorted(summary_bonds.items(), key=lambda x: (x[1]['data'].get('截标时间') or ''), reverse=True)
for row_idx, (bond_name, info) in enumerate(sorted_bonds, 2):
    cell = ws2.cell(row=row_idx, column=1, value=info['first_seen'])
    cell.font = data_font
    cell.alignment = data_alignment
    cell.border = thin_border
    for col_idx, header in enumerate(TARGET_HEADERS, 2):
        value = info['data'].get(header)
        cell = ws2.cell(row=row_idx, column=col_idx, value=value if value is not None else '')
        cell.font = data_font
        cell.alignment = data_alignment
        cell.border = thin_border
        if header in ('债券代码', '票面利率(%)') and value:
            cell.fill = today_fill

ws2.column_dimensions['A'].width = 12
for col_idx, header in enumerate(TARGET_HEADERS, 2):
    ws2.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(header, 12)
ws2.row_dimensions[1].height = 30
ws2.freeze_panes = 'B2'

wb_out.save(OUTPUT_PATH)
print(f"\n[SUCCESS] 输出文件: {OUTPUT_PATH}")
print(f"  Sheet1: {len(today_rows)} 条")
print(f"  Sheet2: {len(sorted_bonds)} 条 (票面利率: {rate_filled}/{total_bonds})")

# 输出统计信息供 Node.js 读取
print(f"__STATS__:{len(today_rows)}:{len(sorted_bonds)}:{rate_filled}:{total_bonds}")
