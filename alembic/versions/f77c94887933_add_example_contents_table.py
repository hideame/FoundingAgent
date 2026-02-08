"""add_example_contents_table

Revision ID: f77c94887933
Revises: 12b3ccaa90f3
Create Date: 2026-02-08 16:33:24.182588

"""

from typing import Sequence, Union

import sqlalchemy as sa

from alembic import op

# revision identifiers, used by Alembic.
revision: str = "f77c94887933"
down_revision: Union[str, Sequence[str], None] = "12b3ccaa90f3"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.create_table(
        "example_contents",
        sa.Column(
            "id", sa.Integer(), autoincrement=True, nullable=False, comment="記入例ID"
        ),
        sa.Column(
            "industry_type",
            sa.String(length=50),
            nullable=False,
            comment="業種タイプ（software, restaurant, beauty等）",
        ),
        sa.Column(
            "section_key",
            sa.String(length=50),
            nullable=False,
            comment="セクションキー（motivation, background, service等）",
        ),
        sa.Column("example_text", sa.Text(), nullable=False, comment="記入例の内容"),
        sa.Column("created_at", sa.DateTime(), nullable=False, comment="作成日時"),
        sa.Column("updated_at", sa.DateTime(), nullable=False, comment="更新日時"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("industry_type", "section_key", name="uq_industry_section"),
        comment="業種別・項目別の記入例マスターテーブル",
    )


def downgrade() -> None:
    """Downgrade schema."""
    op.drop_table("example_contents")
