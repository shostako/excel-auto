#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF全ページ詳細解析スクリプト
各ページの偶数番号テーブルの構造を詳細比較
"""

import pdfplumber

# ファイルパス
pdf_ok = "/mnt/z/全社共有/生産管理課/生産管理/受注/【内示表】ハモコ・ジャパン_251022.pdf"
pdf_ng = "/mnt/z/全社共有/生産管理課/生産管理/受注/2025年/【内示表】ハモコ・ジャパン_250620.pdf"

def analyze_table_structure(table, table_name):
    """テーブル構造を詳細解析"""
    if not table or len(table) == 0:
        return {
            "name": table_name,
            "rows": 0,
            "cols": 0,
            "row1": None,
            "row2": None,
            "error": "空テーブル"
        }

    row1 = table[0] if len(table) >= 1 else None
    row2 = table[1] if len(table) >= 2 else None

    return {
        "name": table_name,
        "rows": len(table),
        "cols": len(table[0]),
        "row1": row1,
        "row2": row2,
    }

def compare_pdfs(pdf_path_ok, pdf_path_ng):
    """2つのPDFの全ページを比較"""

    print(f"{'='*80}")
    print("PDF全ページ比較解析")
    print(f"{'='*80}\n")

    with pdfplumber.open(pdf_path_ok) as pdf_ok_obj:
        with pdfplumber.open(pdf_path_ng) as pdf_ng_obj:

            max_pages = min(len(pdf_ok_obj.pages), len(pdf_ng_obj.pages))

            for page_num in range(max_pages):
                print(f"\n{'='*80}")
                print(f"ページ {page_num + 1}")
                print(f"{'='*80}")

                page_ok = pdf_ok_obj.pages[page_num]
                page_ng = pdf_ng_obj.pages[page_num]

                tables_ok = page_ok.extract_tables()
                tables_ng = page_ng.extract_tables()

                print(f"正常PDF: テーブル数 = {len(tables_ok)}")
                print(f"エラーPDF: テーブル数 = {len(tables_ng)}")

                if len(tables_ok) != len(tables_ng):
                    print(f"⚠️ テーブル数が一致しません！")

                # 偶数番号のテーブル（Table002相当 = index 1, Table004相当 = index 3...）
                # Power QueryのTable番号は1始まりだが、Pythonは0始まり
                # Table001 = tables[0], Table002 = tables[1]...

                for table_idx in range(1, len(tables_ok), 2):  # 1, 3, 5, 7...
                    if table_idx >= len(tables_ok) or table_idx >= len(tables_ng):
                        break

                    table_name = f"Table{str(table_idx + 1).zfill(3)}"  # Table002, Table004...

                    print(f"\n--- {table_name} (Python index {table_idx}) ---")

                    struct_ok = analyze_table_structure(tables_ok[table_idx], table_name)
                    struct_ng = analyze_table_structure(tables_ng[table_idx], table_name)

                    # 基本構造の比較
                    if struct_ok["rows"] != struct_ng["rows"]:
                        print(f"⚠️ 行数が違います: 正常={struct_ok['rows']}, エラー={struct_ng['rows']}")

                    if struct_ok["cols"] != struct_ng["cols"]:
                        print(f"⚠️ 列数が違います: 正常={struct_ok['cols']}, エラー={struct_ng['cols']}")

                    # 1行目の比較（ヘッダー行1）
                    if struct_ok["row1"] != struct_ng["row1"]:
                        print(f"⚠️ 1行目が違います:")
                        print(f"  正常PDF: {struct_ok['row1'][:10]}")  # 最初の10列
                        print(f"  エラーPDF: {struct_ng['row1'][:10]}")

                        # 「品番」の位置を確認
                        try:
                            idx_ok = struct_ok["row1"].index("品番") if "品番" in struct_ok["row1"] else -1
                            idx_ng = struct_ng["row1"].index("品番") if "品番" in struct_ng["row1"] else -1
                            print(f"  「品番」の位置: 正常=Column{idx_ok}, エラー=Column{idx_ng}")
                        except:
                            pass

                    # 2行目の比較（ヘッダー行2）
                    if struct_ok["row2"] != struct_ng["row2"]:
                        print(f"⚠️ 2行目が違います:")
                        print(f"  正常PDF: {struct_ok['row2'][:10]}")
                        print(f"  エラーPDF: {struct_ng['row2'][:10]}")

                    if struct_ok["rows"] == struct_ng["rows"] and \
                       struct_ok["cols"] == struct_ng["cols"] and \
                       struct_ok["row1"] == struct_ng["row1"] and \
                       struct_ok["row2"] == struct_ng["row2"]:
                        print(f"✓ 同一構造")

# 実行
try:
    compare_pdfs(pdf_ok, pdf_ng)
except Exception as e:
    print(f"\nエラー発生: {e}")
    import traceback
    traceback.print_exc()

print(f"\n{'='*80}")
print("解析完了")
print(f"{'='*80}")
