# GUIモジュール
"""
tkinterによるGUIインターフェース
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, timedelta
from pathlib import Path
import threading
from typing import List, Optional

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

from .aggregator import get_aggregated_data
from .excel_writer import generate_excel
from .database import test_connection
from .config import DB_PATH


# 選択肢
HASSEI_OPTIONS = ["成形", "塗装", "モール", "加工"]
HAKKEN2_OPTIONS = ["成形", "塗装", "モール", "加工"]


class RyuushutsuApp:
    """流出不良集計ツール GUI"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("流出不良集計ツール")
        self.root.geometry("450x480")
        self.root.resizable(False, False)

        self._create_widgets()
        self._layout_widgets()

    def _create_widgets(self):
        """ウィジェット作成"""
        # フレーム
        self.main_frame = ttk.Frame(self.root, padding=20)

        # 日付範囲
        self.date_frame = ttk.LabelFrame(self.main_frame, text="日付範囲", padding=10)

        ttk.Label(self.date_frame, text="開始日:").grid(row=0, column=0, sticky="w")
        ttk.Label(self.date_frame, text="終了日:").grid(row=1, column=0, sticky="w")

        # デフォルト: 過去1週間
        today = date.today()
        week_ago = today - timedelta(days=7)

        if HAS_TKCALENDAR:
            self.start_date = DateEntry(
                self.date_frame,
                width=15,
                date_pattern="yyyy/mm/dd",
                year=week_ago.year,
                month=week_ago.month,
                day=week_ago.day,
            )
            self.end_date = DateEntry(
                self.date_frame,
                width=15,
                date_pattern="yyyy/mm/dd",
                year=today.year,
                month=today.month,
                day=today.day,
            )
        else:
            self.start_date_var = tk.StringVar(value=week_ago.strftime("%Y/%m/%d"))
            self.end_date_var = tk.StringVar(value=today.strftime("%Y/%m/%d"))
            self.start_date = ttk.Entry(self.date_frame, textvariable=self.start_date_var, width=15)
            self.end_date = ttk.Entry(self.date_frame, textvariable=self.end_date_var, width=15)

        self.start_date.grid(row=0, column=1, padx=5, pady=2)
        self.end_date.grid(row=1, column=1, padx=5, pady=2)

        # フィルタ
        self.filter_frame = ttk.LabelFrame(self.main_frame, text="フィルタ", padding=10)

        # 発生（単一選択）
        ttk.Label(self.filter_frame, text="発生:").grid(row=0, column=0, sticky="w")
        self.hassei_var = tk.StringVar(value=HASSEI_OPTIONS[0])
        self.hassei_combo = ttk.Combobox(
            self.filter_frame,
            textvariable=self.hassei_var,
            values=HASSEI_OPTIONS,
            state="readonly",
            width=12,
        )
        self.hassei_combo.grid(row=0, column=1, columnspan=4, padx=5, pady=2, sticky="w")

        # 発見2（複数選択チェックボックス）
        ttk.Label(self.filter_frame, text="発見2:").grid(row=1, column=0, sticky="w")
        self.hakken2_vars = {}
        for i, opt in enumerate(HAKKEN2_OPTIONS):
            var = tk.BooleanVar(value=False)
            self.hakken2_vars[opt] = var
            cb = ttk.Checkbutton(self.filter_frame, text=opt, variable=var)
            cb.grid(row=1, column=1+i, padx=2, pady=2, sticky="w")

        ttk.Label(self.filter_frame, text="(未選択=全て)").grid(row=2, column=1, columnspan=4, sticky="w")

        # モード2（テキスト入力）
        ttk.Label(self.filter_frame, text="モード2:").grid(row=3, column=0, sticky="w")
        self.mode2_var = tk.StringVar(value="")
        self.mode2_entry = ttk.Entry(
            self.filter_frame,
            textvariable=self.mode2_var,
            width=20,
        )
        self.mode2_entry.grid(row=3, column=1, columnspan=4, padx=5, pady=2, sticky="w")
        ttk.Label(self.filter_frame, text="(空=全て)").grid(row=4, column=1, columnspan=4, sticky="w")

        # ボタン
        self.button_frame = ttk.Frame(self.main_frame)

        self.exec_button = ttk.Button(
            self.button_frame,
            text="実行",
            command=self._on_execute,
            width=15,
        )
        self.close_button = ttk.Button(
            self.button_frame,
            text="閉じる",
            command=self.root.quit,
            width=15,
        )

        # ステータス
        self.status_var = tk.StringVar(value="準備完了")
        self.status_label = ttk.Label(
            self.main_frame,
            textvariable=self.status_var,
            foreground="gray",
        )

        # プログレスバー
        self.progress = ttk.Progressbar(
            self.main_frame,
            mode="indeterminate",
            length=350,
        )

    def _layout_widgets(self):
        """レイアウト"""
        self.main_frame.pack(fill="both", expand=True)

        self.date_frame.pack(fill="x", pady=5)
        self.filter_frame.pack(fill="x", pady=5)

        self.button_frame.pack(pady=20)
        self.exec_button.pack(side="left", padx=10)
        self.close_button.pack(side="left", padx=10)

        self.progress.pack(fill="x", pady=5)
        self.status_label.pack()

    def _get_start_date(self) -> date:
        """開始日を取得"""
        if HAS_TKCALENDAR:
            return self.start_date.get_date()
        else:
            from datetime import datetime
            return datetime.strptime(self.start_date_var.get(), "%Y/%m/%d").date()

    def _get_end_date(self) -> date:
        """終了日を取得"""
        if HAS_TKCALENDAR:
            return self.end_date.get_date()
        else:
            from datetime import datetime
            return datetime.strptime(self.end_date_var.get(), "%Y/%m/%d").date()

    def _get_hakken2_list(self) -> Optional[List[str]]:
        """選択された発見2のリストを取得（未選択ならNone）"""
        selected = [opt for opt, var in self.hakken2_vars.items() if var.get()]
        return selected if selected else None

    def _on_execute(self):
        """実行ボタン押下"""
        # 入力値取得
        try:
            start = self._get_start_date()
            end = self._get_end_date()
        except ValueError as e:
            messagebox.showerror("入力エラー", f"日付の形式が不正です: {e}")
            return

        if start > end:
            messagebox.showerror("入力エラー", "開始日は終了日以前にしてください")
            return

        hassei = self.hassei_var.get()
        hakken2_list = self._get_hakken2_list()
        mode2 = self.mode2_var.get().strip() or None

        # 非同期実行
        self.exec_button.config(state="disabled")
        self.progress.start()
        self.status_var.set("処理中...")

        thread = threading.Thread(
            target=self._execute_task,
            args=(start, end, hassei, hakken2_list, mode2),
            daemon=True,
        )
        thread.start()

    def _execute_task(self, start: date, end: date, hassei: str,
                      hakken2_list: Optional[List[str]], mode2: Optional[str]):
        """バックグラウンドタスク"""
        try:
            # DB接続テスト
            self._update_status("データベース接続中...")
            if not test_connection():
                raise ConnectionError("データベースに接続できません")

            # データ取得＆集計
            self._update_status("データ取得・集計中...")
            df = get_aggregated_data(start, end, hassei, hakken2_list, mode2)

            # Excel出力
            self._update_status("Excel出力中...")
            output_path = generate_excel(df, hassei, start, end)

            # 完了
            self._update_status(f"完了: {output_path.name}")
            self.root.after(0, lambda: messagebox.showinfo(
                "完了",
                f"Excelファイルを出力しました:\n{output_path}"
            ))

        except Exception as e:
            import traceback
            from pathlib import Path
            error_detail = traceback.format_exc()
            # エラーログをデスクトップに保存
            error_log = Path.home() / "Desktop" / "ryuushutsu_error.txt"
            with open(error_log, "w", encoding="utf-8") as f:
                f.write(error_detail)
            self._update_status(f"エラー: {e}")
            self.root.after(0, lambda: messagebox.showerror("エラー", f"{e}\n\nエラー詳細: {error_log}"))

        finally:
            self.root.after(0, self._finish_task)

    def _update_status(self, message: str):
        """ステータス更新（スレッドセーフ）"""
        self.root.after(0, lambda: self.status_var.set(message))

    def _finish_task(self):
        """タスク完了処理"""
        self.progress.stop()
        self.exec_button.config(state="normal")


def run_gui():
    """GUIを起動"""
    root = tk.Tk()
    app = RyuushutsuApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()
