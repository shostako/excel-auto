#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF構造比較デバッグスクリプト
正常なPDFとエラーが出るPDFのテーブル構造を比較する
"""

import pdfplumber

# ファイルパス
pdf_ok = "/mnt/z/全社共有/生産管理課/生産管理/受注/【内示表】ハモコ・ジャパン_251022.pdf"
pdf_ng = "/mnt/z/全社共有/生産管理課/生産管理/受注/2025年/【内示表】ハモコ・ジャパン_250620.pdf"

def analyze_pdf(pdf_path, label):
    """PDFのテーブル構造を解析"""
    print(f"\n{'='*60}")
    print(f"{label}: {pdf_path}")
    print(f"{'='*60}")

    with pdfplumber.open(pdf_path) as pdf:
        print(f"総ページ数: {len(pdf.pages)}\n")

        # 最初のページを詳細解析
        page1 = pdf.pages[0]
        tables = page1.extract_tables()

        print(f"ページ1のテーブル数: {len(tables)}\n")

        # 各テーブルの構造を表示
        for idx, table in enumerate(tables):
            print(f"--- Table {idx+1} ---")
            print(f"行数: {len(table)}")
            if len(table) > 0:
                print(f"列数: {len(table[0])}")
                print(f"最初の5行（各行の最初の5列のみ）:")
                for row_idx, row in enumerate(table[:5]):
                    # 各セルの内容を表示（Noneは空文字に変換）
                    row_display = [str(cell) if cell else "" for cell in row[:5]]
                    print(f"  行{row_idx+1}: {row_display}")
            print()

# 正常なPDFを解析
try:
    analyze_pdf(pdf_ok, "正常PDF (251022)")
except Exception as e:
    print(f"エラー: {e}")

# エラーが出るPDFを解析
try:
    analyze_pdf(pdf_ng, "エラーPDF (250620)")
except Exception as e:
    print(f"エラー: {e}")

print(f"\n{'='*60}")
print("解析完了")
print(f"{'='*60}")
