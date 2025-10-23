#!/bin/bash
# ワークスペースクリーンアップスクリプト
# セッション終了時にsrc/、macros/、参考マクロ/のファイルを削除してコミット

set -e

cd /home/shostako/ClaudeCode/excel-auto

# --autoオプションの処理
AUTO_MODE=false
if [ "$1" = "--auto" ]; then
    AUTO_MODE=true
fi

echo "=========================================="
echo "ワークスペースクリーンアップ"
echo "=========================================="
echo

# 削除対象ファイルの確認
echo "【削除対象ファイル】"
found=0

if ls src/*.bas 2>/dev/null | grep -q .; then
    echo
    echo "src/:"
    ls -1 src/*.bas 2>/dev/null | sed 's|src/|  - |'
    found=1
fi

if ls macros/*.bas 2>/dev/null | grep -q .; then
    echo
    echo "macros/:"
    ls -1 macros/*.bas 2>/dev/null | sed 's|macros/|  - |'
    found=1
fi

if ls 参考マクロ/*.bas 2>/dev/null | grep -q .; then
    echo
    echo "参考マクロ/:"
    ls -1 参考マクロ/*.bas 2>/dev/null | sed 's|参考マクロ/|  - |'
    found=1
fi

if [ $found -eq 0 ]; then
    echo "  （削除対象ファイルはありません）"
    echo
    echo "ワークスペースは既にクリーンです。"
    exit 0
fi

echo
echo "=========================================="

# --autoモードの場合は確認をスキップ
if [ "$AUTO_MODE" = true ]; then
    confirm="y"
    echo "自動モード: 削除とコミットを実行します"
else
    read -p "削除してGitコミットしますか？ (y/n): " confirm
fi

echo

if [ "$confirm" = "y" ] || [ "$confirm" = "Y" ]; then
    # ファイル削除
    rm -f src/*.bas macros/*.bas 参考マクロ/*.bas 2>/dev/null || true

    # Git操作
    git add -A

    # 変更があるか確認
    if git diff --cached --quiet; then
        echo "削除するファイルがありませんでした。"
    else
        git commit -m "clean: ワークスペースクリーンアップ - セッション終了"
        echo
        echo "✓ クリーンアップ完了"
        echo "✓ Gitコミット完了"
    fi
else
    echo "キャンセルしました。"
    exit 1
fi

echo
echo "=========================================="
echo "ワークスペース状態："
echo "  src/      : $(ls src/*.bas 2>/dev/null | wc -l) ファイル"
echo "  macros/   : $(ls macros/*.bas 2>/dev/null | wc -l) ファイル"
echo "  参考マクロ/: $(ls 参考マクロ/*.bas 2>/dev/null | wc -l) ファイル"
echo "=========================================="
