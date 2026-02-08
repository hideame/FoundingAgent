#!/usr/bin/env python3
"""
PDFから抽出した記入例をデータベースに保存するスクリプト

使い方:
    python scripts/seed_examples.py
"""

import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import asyncio

# PDF抽出スクリプトをインポート
from extract_examples_from_pdf import main as extract_main
from sqlalchemy import delete, select

from app.database import AsyncSessionLocal
from app.models import ExampleContent


async def seed_database(all_data: dict):
    """
    抽出したデータをデータベースに保存

    Args:
        all_data: {industry_type: {section_key: text}} の辞書
    """
    async with AsyncSessionLocal() as session:
        # 既存のデータを削除
        print("\n既存のexample_contentsデータを削除中...")
        await session.execute(delete(ExampleContent))
        await session.commit()
        print("✅ 削除完了")

        # 新しいデータを挿入
        print("\n新しい記入例を挿入中...")
        inserted_count = 0

        for industry_type, sections in all_data.items():
            for section_key, example_text in sections.items():
                # テキストをクリーンアップ
                cleaned_text = clean_text(example_text)

                # 空のテキストはスキップ
                if not cleaned_text or len(cleaned_text) < 10:
                    print(
                        f"  ⚠️  {industry_type}.{section_key}: テキストが短すぎるためスキップ"
                    )
                    continue

                example = ExampleContent(
                    industry_type=industry_type,
                    section_key=section_key,
                    example_text=cleaned_text,
                )
                session.add(example)
                inserted_count += 1

                # プレビュー表示
                preview = cleaned_text[:80].replace("\n", " ")
                print(f"  ✅ {industry_type}.{section_key}: {preview}...")

        await session.commit()
        print(f"\n✅ 合計 {inserted_count} 件のデータを挿入しました")


def clean_text(text: str) -> str:
    """
    抽出したテキストをクリーンアップ
    - 不要な空白・改行を整理
    - 表形式のデータを読みやすくフォーマット
    """
    import re

    # 連続する空白を1つに
    text = re.sub(r" {2,}", " ", text)

    # 3つ以上の改行を2つに
    text = re.sub(r"\n{3,}", "\n\n", text)

    # 前後の空白を削除
    text = text.strip()

    return text


async def verify_data():
    """データベースに保存されたデータを検証"""
    async with AsyncSessionLocal() as session:
        result = await session.execute(select(ExampleContent))
        examples = result.scalars().all()

        print(f"\n=== データベース検証 ===")
        print(f"保存されているレコード数: {len(examples)}")

        # 業種ごとの集計
        by_industry = {}
        for example in examples:
            if example.industry_type not in by_industry:
                by_industry[example.industry_type] = []
            by_industry[example.industry_type].append(example.section_key)

        print("\n業種別セクション数:")
        for industry_type, section_keys in sorted(by_industry.items()):
            print(f"  {industry_type}: {len(section_keys)} セクション")
            missing = []
            from extract_examples_from_pdf import SECTION_PATTERNS

            for section_key in SECTION_PATTERNS.keys():
                if section_key not in section_keys:
                    missing.append(section_key)
            if missing:
                print(f"    未登録: {', '.join(missing)}")


async def main_async():
    """メイン処理"""
    # PDFから記入例を抽出
    print("=== ステップ1: PDFから記入例を抽出 ===")
    all_data = extract_main()

    # データベースに保存
    print("\n=== ステップ2: データベースに保存 ===")
    await seed_database(all_data)

    # 検証
    print("\n=== ステップ3: データの検証 ===")
    await verify_data()

    print("\n✅ 全ての処理が完了しました!")


if __name__ == "__main__":
    asyncio.run(main_async())
