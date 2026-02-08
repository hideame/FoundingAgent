#!/usr/bin/env python3
"""
PDF記入例ファイルからテキストを抽出してデータベースに保存するスクリプト

使い方:
    python scripts/extract_examples_from_pdf.py
"""

import re
import sys
from pathlib import Path
from typing import Dict, Tuple

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pdfplumber

# 業種とPDFファイルのマッピング
INDUSTRY_FILES = {
    "software": "software_example.pdf",  # ソフトウェア開発
    "restaurant": "restaurant_example.pdf",  # 飲食店（居酒屋）
    "beauty": "beauty_example.pdf",  # 美容室
    "apparel": "apparel_example.pdf",  # アパレル（子供服）
    "construction": "construction_example.pdf",  # 建設業（内装工事・リフォーム）
    "cram_school": "cram_school_example.pdf",  # 学習塾（英語塾）
    "care_service": "care_service_example.pdf",  # 介護サービス（デイサービス）
    "car_sales": "car_sales_example.pdf",  # 中古車販売
    "dentist": "dentist_example.pdf",  # 歯科医院
}

# セクションキーと対応する見出しパターン
SECTION_PATTERNS = {
    "motivation": r"１\s+創業の動機",
    "background": r"２\s+経営者の略歴等",
    "service": r"３\s+取扱商品・サービス",
    "employees": r"４\s+従業員",
    "partners": r"５\s+取引先・取引関係等",
    "related_companies": r"６\s+関連企業",
    "loans": r"７\s+お借入の状況",
    "funds": r"８\s+必要な資金と調達方法",
    "outlook": r"９\s+事業の見通し",
}


def extract_text_from_pdf(pdf_path: Path) -> str:
    """PDFファイルから全テキストを抽出"""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def find_section_positions(text: str) -> Dict[str, Tuple[int, int]]:
    """
    テキスト内の各セクションの開始位置と終了位置を検出

    Returns:
        {section_key: (start_pos, end_pos)} の辞書
    """
    positions = {}

    for section_key, pattern in SECTION_PATTERNS.items():
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            positions[section_key] = match.start()

    # セクションの開始位置でソート
    sorted_sections = sorted(positions.items(), key=lambda x: x[1])

    # 各セクションの終了位置を次のセクションの開始位置とする
    result = {}
    for i, (section_key, start_pos) in enumerate(sorted_sections):
        if i < len(sorted_sections) - 1:
            end_pos = sorted_sections[i + 1][1]
        else:
            end_pos = len(text)
        result[section_key] = (start_pos, end_pos)

    return result


def clean_extracted_text(text: str) -> str:
    """
    抽出したテキストから不要なノイズ（右カラムのフォームヘッダーなど）を除去

    PDFの2カラムレイアウトで混入する以下のようなノイズを除去:
    - 「企 業 名」「関 連代表者名」「① 所 在 地」などのフォームラベル
    - 極端に短い行（スペースが多い行）
    """
    # 不要なパターンを削除
    noise_patterns = [
        r"企\s*業\s*名",
        r"関\s*連代表者名",
        r"[①②]\s*所\s*在\s*地",
        r"業\s*種",
        r"お借入先名",
        r"お使いみち",
        r"お借入残高",
        r"年間返済額",
        r"見積先",
        r"金\s*額",
        r"調達の方法",
        r"取\s*引\s*先\s*名",
        r"所在地等（市区町村）",
        r"回収・支払の条件",
        r"\s{10,}",  # 連続する10個以上のスペース
    ]

    for pattern in noise_patterns:
        text = re.sub(pattern, " ", text)

    # 複数の改行を2つまでに制限
    text = re.sub(r"\n{3,}", "\n\n", text)

    # 余分なスペースを削除
    text = re.sub(r" {2,}", " ", text)

    return text.strip()


def extract_sections(pdf_path: Path) -> Dict[str, str]:
    """
    PDFファイルから各セクションのテキストを抽出

    PDFは2カラムレイアウトで、pdfplumberは以下の順序でテキストを抽出:
    - 見出し(左) + 見出し(右) → 記入例(左) + フォームヘッダー(右) → ...

    各セクションの見出し行を除去し、次の数字見出しまでのテキストを記入例として抽出。

    Returns:
        {section_key: extracted_text} の辞書
    """
    full_text = extract_text_from_pdf(pdf_path)
    sections = {}

    for section_key, pattern in SECTION_PATTERNS.items():
        # セクション見出しを検索
        match = re.search(pattern + r"[^\n]*", full_text, re.MULTILINE)
        if not match:
            continue

        # 見出し行の終わりから開始
        start_pos = match.end()

        # 次の数字見出し（１-９のいずれか）までを抽出
        # 全角数字の見出しパターン: [１-９]\s+
        remaining_text = full_text[start_pos:]
        next_heading_match = re.search(r"[１-９]\s+\S", remaining_text, re.MULTILINE)

        if next_heading_match:
            end_pos = next_heading_match.start()
            section_text = remaining_text[:end_pos]
        else:
            section_text = remaining_text

        # テキストをクリーンアップ
        section_text = clean_extracted_text(section_text)

        # 空白行や短すぎるテキストをスキップ
        if len(section_text) < 10:
            continue

        sections[section_key] = section_text

    return sections


def main():
    """メイン処理: 各業種のPDFからセクションを抽出して表示"""
    pdf_dir = project_root / "app" / "static" / "templates" / "examples"

    print("=== PDF記入例の抽出 ===\n")

    all_data = {}  # {industry_type: {section_key: text}}

    for industry_type, filename in INDUSTRY_FILES.items():
        pdf_path = pdf_dir / filename

        if not pdf_path.exists():
            print(f"⚠️  {industry_type}: {filename} が見つかりません")
            continue

        print(f"📄 {industry_type} ({filename})")
        print("-" * 60)

        try:
            sections = extract_sections(pdf_path)
            all_data[industry_type] = sections

            for section_key in SECTION_PATTERNS.keys():
                if section_key in sections:
                    section_text = sections[section_key]
                    # 最初の100文字だけ表示
                    preview = section_text[:100].replace("\n", " ")
                    print(f"  ✅ {section_key}: {preview}...")
                else:
                    print(f"  ❌ {section_key}: 見つかりません")

            print(f"\n合計: {len(sections)}/10 セクション抽出\n")

        except Exception as e:
            print(f"❌ エラー: {e}\n")

    # 統計情報
    print("\n" + "=" * 80)
    print("=== 抽出結果サマリー ===")
    total_sections = sum(len(sections) for sections in all_data.values())
    expected_sections = len(INDUSTRY_FILES) * len(SECTION_PATTERNS)
    print(f"抽出セクション数: {total_sections} / {expected_sections}")
    print(f"抽出率: {total_sections / expected_sections * 100:.1f}%")

    return all_data


if __name__ == "__main__":
    main()
