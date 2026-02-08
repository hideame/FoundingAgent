"""add_industry_type_to_sessions

Revision ID: 9b5be5058aa3
Revises: 8eb452de0f0b
Create Date: 2026-02-08 13:42:34.127504

"""

from typing import Sequence, Union

import sqlalchemy as sa

from alembic import op

# revision identifiers, used by Alembic.
revision: str = "9b5be5058aa3"
down_revision: Union[str, Sequence[str], None] = "8eb452de0f0b"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    # Add industry_type column to sessions table
    op.add_column("sessions", sa.Column("industry_type", sa.String(50), nullable=True))


def downgrade() -> None:
    """Downgrade schema."""
    # Remove industry_type column from sessions table
    op.drop_column("sessions", "industry_type")
