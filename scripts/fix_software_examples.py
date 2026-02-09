#!/usr/bin/env python3
"""
software業種のexample_contentsを手動で正しい内容に更新するスクリプト

PDFの2カラムレイアウトにより自動抽出で混入したノイズを修正し、
正確なデータを登録する。
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import asyncio

from sqlalchemy import delete, select

from app.database import AsyncSessionLocal
from app.models import ExampleContent

# PDFから手動で正確に読み取った各セクションの内容
SOFTWARE_SECTIONS = {
    "motivation": (
        "・ソフトウェア会社に約20年勤めるなかで、ソフトウェアの企画開発・製作・販売・運用・管理に一貫して携わってきた。\n"
        "○○データ（株）では、プロジェクトリーダーとしてチームを纏め上げた経験から自分の技術に自信がつき、\n"
        "独立を前向きに検討していたところ、同社からの受注を得られる運びとなり、事業の見通しが立ったため。"
    ),
    "background": (
        "Ｈ○年○月：○○工科学院卒\n"
        "Ｈ○年○月～：（株）○○システムズ（ソフトウェア開発業）７年勤務\n"
        "Ｈ○年○月～：○○データ（株）（ソフトウェア開発業）12年勤務（医療関連事業部プロジェクトリーダーを務める）（当時の月給40万円）\n"
        "Ｒ○年○月：退職（退職金200万円）\n"
        "現在：創業準備中\n"
        "過去の事業経験：事業を経営していたことはない。\n"
        "取得資格：有（応用情報技術者資格（Ｈ○年○月）番号等）\n"
        "許認可：特になし\n"
        "知的財産権等：特になし、申請中、登録済"
    ),
    "service": (
        "事業内容：介護・医療施設用の顧客・財務管理システム開発及びファームウェア開発（開発期間３ヵ月～\n"
        "半年ほど）○○データ（株）からの業務請負だけでなく、介護・医療施設からの受注も拡大していく。\n"
        "取扱商品・サービスの内容：\n"
        "① 顧客・財務管理システム開発（売上シェア 70％）\n"
        "② 医療関連機器のファームウェア開発（売上シェア 20％）\n"
        "③ コンサルティング（売上シェア 10％）\n"
        "受注（販売）単価：300万円～1,000万円\n"
        "セールスポイント（自社の強み）：介護・医療関連のシステム開発の知識を生かし、システム開発の提供だけでなく、運用に関するコンサルティングが可能。\n"
        "販売ターゲット・販売戦略（集客方法）：元勤務先（○○データ（株））からの業務請負（Ｒ○年○月○日契約締結済）を軸に、並行して営業を行い受注の幅を広げる。\n"
        "競合・市場など自社を取り巻く状況：介護・医療関係はシステム化の需要が大きい。\n"
        "当社のように企画開発・製作・販売・運用・管理に一貫して対応できる企業は少ないため、成長が見込める。"
    ),
    "employees": (
        "常勤役員の人数（法人の方のみ）：2人\n"
        "従業員数（３ヵ月以上継続雇用者）：1人\n"
        "（うち家族従業員）：0人\n"
        "（うちパート従業員）：0人"
    ),
    "partners": (
        "販売先：\n"
        "○○データ（株）（元勤務先） シェア70％ 掛取引の割合100％ ○○区○○ 末日〆 翌月末日回収\n"
        "医療法人○○会（元勤務先の販売先） シェア30％ 掛取引の割合100％ ○○区○○ 末日〆 翌月末日回収\n"
        "仕入先：なし\n"
        "外注先：\n"
        "○○ソフト（株）（元勤務先の外注先） シェア100％ 掛取引の割合100％ ○○区○○ 末日〆 翌月末日支払\n"
        "人件費の支払：末日〆 翌月25日支払（ボーナスの支給月 6月、12月）"
    ),
    # related_companies: PDFでは空欄（記入なし）のため登録しない
    "loans": (
        "◎◎銀行△△支店 その他 お借入残高2,554万円 年間返済額132万円"
    ),
    "funds": (
        "必要な資金：\n"
        "設備資金：店舗、工場、機械、車両など 690万円\n"
        "（内訳）\n"
        "・パソコン・サーバー等一式 ○○社 500万円\n"
        "・事務機器 ○×社 70万円\n"
        "・備品類 △△社 20万円\n"
        "・保証金 100万円\n"
        "運転資金：商品仕入、経費支払資金など 920万円\n"
        "（内訳）\n"
        "・外注費支払 270万円\n"
        "・諸経費支払（家賃等含む） 650万円\n"
        "（システム開発に、最短でも３ヵ月かかるため、つなぎ資金が必要）\n"
        "合計 1,610万円\n\n"
        "調達方法：\n"
        "自己資金 610万円\n"
        "日本政策金融公庫 国民生活事業からの借入 500万円（内訳・返済方法）元金６万円×84回（年○.○％）\n"
        "他の金融機関等からの借入 500万円（内訳・返済方法）○○銀行 元金６万円×84回（年○.○％）\n"
        "合計 1,610万円"
    ),
    "outlook": (
        "創業当初：\n"
        "売上高① 300万円\n"
        "売上原価②（仕入高） 90万円\n"
        "経費：\n"
        "  人件費 100万円\n"
        "  家賃 20万円\n"
        "  支払利息 3万円\n"
        "  その他 75万円\n"
        "  合計③ 198万円\n"
        "利益（①－②－③） 12万円\n\n"
        "１年後又は軌道に乗った後（○年○月頃）：\n"
        "売上高① 390万円\n"
        "売上原価②（仕入高） 117万円\n"
        "経費：\n"
        "  人件費 140万円\n"
        "  家賃 20万円\n"
        "  支払利息 3万円\n"
        "  その他 95万円\n"
        "  合計③ 258万円\n"
        "利益（①－②－③） 15万円\n\n"
        "＜創業当初＞\n"
        "①売上高 300万円／件×１件／月＝300万円（受注契約書あり）\n"
        "②原価率（外注費）30％（勤務時の経験から）\n"
        "③人件費 代表者１人、役員１人、従業員１人\n"
        " （代）45万円＋（役）30万円＋（従）25万円＝100万円\n"
        "  家賃 20万円\n"
        "  支払利息（内訳）500万円×年○.○％÷12ヵ月＝○万円\n"
        "            500万円×年○.○％÷12ヵ月＝○万円 計３万円\n"
        "  その他光熱費、消耗品費等 75万円\n"
        "＜軌道に乗った後＞\n"
        "①創業当初の1.3倍（勤務時の経験から）\n"
        "②当初の原価率を採用\n"
        "③人件費 従業員1人増、役員報酬・従業員給与増額 計40万円増\n"
        "  その他諸経費 20万円増"
    ),
}


async def update_software_examples():
    """software業種のexample_contentsを更新"""
    async with AsyncSessionLocal() as session:
        # software業種の既存データを削除
        print("software業種の既存データを削除中...")
        await session.execute(
            delete(ExampleContent).where(
                ExampleContent.industry_type == "software"
            )
        )
        await session.commit()
        print("削除完了\n")

        # 正しいデータを挿入
        print("正しいデータを挿入中...")
        inserted_count = 0
        for section_key, example_text in SOFTWARE_SECTIONS.items():
            example = ExampleContent(
                industry_type="software",
                section_key=section_key,
                example_text=example_text,
            )
            session.add(example)
            inserted_count += 1
            preview = example_text[:80].replace("\n", " ")
            print(f"  {section_key}: {preview}...")

        await session.commit()
        print(f"\n合計 {inserted_count} 件挿入完了")

        # 検証
        print("\n=== 検証 ===")
        result = await session.execute(
            select(ExampleContent).where(
                ExampleContent.industry_type == "software"
            )
        )
        rows = result.scalars().all()
        print(f"登録セクション数: {len(rows)}")
        for row in rows:
            print(f"  {row.section_key}: {len(row.example_text)}文字")


if __name__ == "__main__":
    asyncio.run(update_software_examples())
