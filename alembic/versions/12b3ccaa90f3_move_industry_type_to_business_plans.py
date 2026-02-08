"""move_industry_type_to_business_plans

Revision ID: 12b3ccaa90f3
Revises: 9b5be5058aa3
Create Date: 2026-02-08 15:15:10.207406

"""

from typing import Sequence, Union

import sqlalchemy as sa

from alembic import op

# revision identifiers, used by Alembic.
revision: str = "12b3ccaa90f3"
down_revision: Union[str, Sequence[str], None] = "9b5be5058aa3"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    # 1. business_plansテーブルにindustry_typeカラムを追加
    op.add_column(
        "business_plans", sa.Column("industry_type", sa.String(50), nullable=True)
    )

    # 2. 既存データをsessionsからbusiness_plansに移行
    # まず、business_plansレコードが存在するセッションのデータを移行
    op.execute("""
        UPDATE business_plans bp
        INNER JOIN sessions s ON bp.session_id = s.id
        SET bp.industry_type = s.industry_type
        WHERE s.industry_type IS NOT NULL
    """)

    # 3. sessionsテーブルからindustry_typeカラムを削除
    op.drop_column("sessions", "industry_type")


def downgrade() -> None:
    """Downgrade schema."""
    # 1. sessionsテーブルにindustry_typeカラムを再追加
    op.add_column("sessions", sa.Column("industry_type", sa.String(50), nullable=True))

    # 2. データをbusiness_plansからsessionsに戻す
    op.execute("""
        UPDATE sessions s
        INNER JOIN business_plans bp ON s.id = bp.session_id
        SET s.industry_type = bp.industry_type
        WHERE bp.industry_type IS NOT NULL
    """)

    # 3. business_plansテーブルからindustry_typeカラムを削除
    op.drop_column("business_plans", "industry_type")
