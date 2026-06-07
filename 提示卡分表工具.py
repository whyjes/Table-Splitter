#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提示卡分表工具
================
将总表中的每行数据拆分为独立的 sheet，方便阅读。
- 每个车组生成一个 sheet，以"姓名+车次"命名
- 新建"查询页"，带超链接可跳转到对应 sheet
- 原文件不变，生成新文件（分表）

用法：
  1. 拖拽 Excel 文件到本程序图标上
  2. 或命令行：提示卡分表工具.exe <文件路径>
"""

import sys
import os
import re
import traceback

# ----- 只在打包时用到的导入声明（供 PyInstaller 识别）-----
# noinspection PyUnresolvedReferences
_hook_openpyxl = (
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.workbook",
    "openpyxl.worksheet",
)


def split_prompt_card(input_path: str) -> str:
    """处理单个 Excel 文件，返回输出文件路径"""
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    # --- 校验 ---
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"文件不存在：{input_path}")

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in (".xlsx", ".xlsm"):
        raise ValueError(f"不支持的文件格式：{ext}，请使用 .xlsx 文件")

    # --- 打开 ---
    wb = openpyxl.load_workbook(input_path)
    ws_src = wb.active

    # 读取表头（第2行）
    headers = []
    for col in range(1, ws_src.max_column + 1):
        headers.append(ws_src.cell(row=2, column=col).value)

    # --- 收集数据行 ---
    rows_data = []
    for r in range(3, ws_src.max_row + 1):
        chezu = ws_src.cell(row=r, column=3).value
        if not chezu or not str(chezu).strip():
            continue
        values = []
        for c in range(1, ws_src.max_column + 1):
            values.append(ws_src.cell(row=r, column=c).value)
        rows_data.append(values)

    if not rows_data:
        raise ValueError("未找到任何有效数据（车组号列为空）")

    # --- 样式定义 ---
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    label_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    label_fill = PatternFill(
        start_color="4472C4", end_color="4472C4", fill_type="solid"
    )
    value_font = Font(name="微软雅黑", size=11, color="333333")
    title_font = Font(name="微软雅黑", bold=True, size=14, color="1F4E79")
    wrap_align = Alignment(wrap_text=True, vertical="top")

    # --- 删除所有旧 sheet（只保留个人页和查询页）---
    for s in list(wb.sheetnames):
        del wb[s]

    # --- 逐个创建个人 sheet ---
    created = []
    seen_names = {}

    for values in rows_data:
        chezu = str(values[2]).strip() if values[2] else ""
        jixieshi = str(values[4]).strip() if values[4] else ""
        checi = str(values[3]).strip() if values[3] else ""

        # Sheet 名 = 姓名+车次（清理非法字符）
        raw = f"{jixieshi}{checi}"
        safe = re.sub(r'[\\/?*\[\]]', "·", raw)
        base = safe[:31]

        if base in seen_names:
            seen_names[base] += 1
            sheet_name = f"{base[:28]}_{seen_names[base]}"
        else:
            seen_names[base] = 1
            sheet_name = base

        ws_new = wb.create_sheet(title=sheet_name)

        # 标题
        ws_new.merge_cells("A1:B1")
        c = ws_new["A1"]
        c.value = f"提示卡 — {chezu}（{jixieshi}）"
        c.font = title_font
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_new.row_dimensions[1].height = 30

        ws_new["A2"] = None  # 空行

        cur = 3
        for idx, h in enumerate(headers):
            h_name = h.strip() if h else ""
            val = values[idx]

            if hasattr(val, "strftime"):
                val = val.strftime("%Y-%m-%d")
            elif val is None:
                val = ""
            else:
                val = re.sub(r"\n{3,}", "\n\n", str(val).strip())

            # 标签
            cl = ws_new.cell(row=cur, column=1)
            cl.value = h_name
            cl.font = label_font
            cl.fill = label_fill
            cl.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cl.border = thin_border

            # 值
            cv = ws_new.cell(row=cur, column=2)
            cv.value = val
            cv.font = value_font
            cv.alignment = wrap_align
            cv.border = thin_border

            lines = max(1, len(val) // 80 + val.count("\n") + 1)
            ws_new.row_dimensions[cur].height = max(25, lines * 18)
            cur += 1

        ws_new.column_dimensions["A"].width = 28
        ws_new.column_dimensions["B"].width = 100
        created.append((sheet_name, chezu, jixieshi, checi))

    # --- 创建查询页（放到最前面）---
    if "查询页" in wb.sheetnames:
        del wb["查询页"]

    ws_q = wb.create_sheet(title="查询页", index=0)

    ws_q.merge_cells("A1:D1")
    t = ws_q["A1"]
    t.value = "提示卡查询 — 点击跳转"
    t.font = Font(name="微软雅黑", bold=True, size=16, color="1F4E79")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_q.row_dimensions[1].height = 36

    hdr_fill = PatternFill(
        start_color="4472C4", end_color="4472C4", fill_type="solid"
    )
    hdr_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    hdr_align = Alignment(horizontal="center", vertical="center")

    for ci, txt in enumerate(["序号", "车组号", "姓名+车次", "跳转"], 1):
        c = ws_q.cell(row=2, column=ci)
        c.value = txt
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = hdr_align
        c.border = thin_border

    link_font = Font(
        name="微软雅黑", size=11, color="0563C1", underline="single"
    )
    normal_font = Font(name="微软雅黑", size=11)

    for i, (sn, cz, js, cc) in enumerate(created, 1):
        rn = i + 2
        c1 = ws_q.cell(row=rn, column=1)
        c1.value = i
        c1.font = normal_font
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = thin_border

        c2 = ws_q.cell(row=rn, column=2)
        c2.value = cz
        c2.font = normal_font
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = thin_border

        c3 = ws_q.cell(row=rn, column=3)
        c3.value = f"{js}{cc}"
        c3.font = normal_font
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = thin_border

        c4 = ws_q.cell(row=rn, column=4)
        c4.value = f'=HYPERLINK("#''{sn}''!A1","点击查看")'
        c4.font = link_font
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border = thin_border
        ws_q.row_dimensions[rn].height = 24

    ws_q.column_dimensions["A"].width = 8
    ws_q.column_dimensions["B"].width = 12
    ws_q.column_dimensions["C"].width = 22
    ws_q.column_dimensions["D"].width = 14

    # --- 保存 ---
    base, ext = os.path.splitext(input_path)
    out_path = f"{base}（分表）{ext}"
    wb.save(out_path)
    return out_path


def main():
    """入口：支持拖拽 / 命令行参数 / 交互输入"""
    if len(sys.argv) >= 2:
        # 命令行参数或拖拽
        inputs = sys.argv[1:]
    else:
        # 交互模式
        print("=" * 50)
        print("  提示卡分表工具 v1.0")
        print("=" * 50)
        inp = input("\n请拖入或输入 Excel 文件路径：\n> ").strip().strip('"')
        if not inp:
            print("未输入文件，退出。")
            input("\n按回车键退出...")
            return
        inputs = [inp]

    for fp in inputs:
        path = fp.strip().strip('"')
        if not os.path.isfile(path):
            print(f"  ✗ 跳过（文件不存在）：{path}")
            continue

        try:
            out = split_prompt_card(path)
            print(f"\n  ✓ 生成成功：{out}")
            print(f"    文件大小：{os.path.getsize(out):,} 字节")
        except Exception as e:
            print(f"\n  ✗ 处理失败：{path}")
            print(f"    错误：{e}")

    input("\n完成！按回车键退出...")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("\n发生未预期错误，按回车键退出...")
