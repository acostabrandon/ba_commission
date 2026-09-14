"""Persistence for the commission calculator.

Uses a shared Postgres database when DATABASE_URL is provided (Streamlit Cloud +
a free hosted Postgres such as Neon/Supabase), otherwise a local SQLite file for
development. A payroll "run" is stored as one row with a JSON payload, so the whole
team opens the same in-progress run and picks up where others left off.
"""
from __future__ import annotations
import os, json, datetime, uuid
from sqlalchemy import create_engine, String, Text, DateTime, LargeBinary, select, delete
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, Session


def _database_url() -> str:
    # Prefer Streamlit secrets, then env var, then local sqlite.
    url = None
    try:
        import streamlit as st
        url = st.secrets.get("DATABASE_URL")  # type: ignore[attr-defined]
    except Exception:
        url = None
    url = url or os.environ.get("DATABASE_URL")
    if not url:
        return "sqlite:///commission.db"
    # SQLAlchemy needs the psycopg driver spelled out for some URLs.
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    return url


class Base(DeclarativeBase):
    pass


class Run(Base):
    __tablename__ = "commission_runs"
    id: Mapped[str] = mapped_column(String(40), primary_key=True)
    name: Mapped[str] = mapped_column(String(200))
    period_start: Mapped[str] = mapped_column(String(20))
    period_end: Mapped[str] = mapped_column(String(20))
    status: Mapped[str] = mapped_column(String(20), default="draft")
    updated_at: Mapped[datetime.datetime] = mapped_column(DateTime, default=datetime.datetime.utcnow)
    updated_by: Mapped[str] = mapped_column(String(120), default="")
    payload: Mapped[str] = mapped_column(Text, default="{}")


class Attachment(Base):
    __tablename__ = "commission_attachments"
    id: Mapped[str] = mapped_column(String(40), primary_key=True)
    run_id: Mapped[str] = mapped_column(String(40), index=True)
    kind: Mapped[str] = mapped_column(String(20))   # 'lps' | 'order' | 'other'
    name: Mapped[str] = mapped_column(String(300))
    content: Mapped[bytes] = mapped_column(LargeBinary)


_engine = None


def engine():
    global _engine
    if _engine is None:
        _engine = create_engine(_database_url(), future=True)
        Base.metadata.create_all(_engine)
    return _engine


def list_runs() -> list[dict]:
    with Session(engine()) as s:
        rows = s.execute(select(Run).order_by(Run.updated_at.desc())).scalars().all()
        return [{"id": r.id, "name": r.name, "period_start": r.period_start,
                 "period_end": r.period_end, "status": r.status,
                 "updated_at": r.updated_at, "updated_by": r.updated_by} for r in rows]


def create_run(name, period_start, period_end, user="") -> str:
    rid = uuid.uuid4().hex[:12]
    with Session(engine()) as s:
        s.add(Run(id=rid, name=name, period_start=period_start, period_end=period_end,
                  status="draft", updated_by=user, payload=json.dumps({})))
        s.commit()
    return rid


def load_run(rid) -> dict | None:
    with Session(engine()) as s:
        r = s.get(Run, rid)
        if not r:
            return None
        return {"id": r.id, "name": r.name, "period_start": r.period_start,
                "period_end": r.period_end, "status": r.status,
                "updated_at": r.updated_at, "updated_by": r.updated_by,
                "payload": json.loads(r.payload or "{}")}


def save_run(rid, payload: dict, status=None, user="") -> None:
    with Session(engine()) as s:
        r = s.get(Run, rid)
        if not r:
            raise KeyError(rid)
        r.payload = json.dumps(payload, default=str)
        if status:
            r.status = status
        r.updated_by = user or r.updated_by
        r.updated_at = datetime.datetime.utcnow()
        s.commit()


def delete_run(rid) -> None:
    with Session(engine()) as s:
        r = s.get(Run, rid)
        if r:
            s.delete(r)
        s.execute(delete(Attachment).where(Attachment.run_id == rid))
        s.commit()


def save_attachment(run_id, kind, name, content: bytes) -> None:
    """Store or replace an attachment (keyed by run_id + kind + name)."""
    with Session(engine()) as s:
        s.execute(delete(Attachment).where(Attachment.run_id == run_id,
                                           Attachment.kind == kind, Attachment.name == name))
        s.add(Attachment(id=uuid.uuid4().hex[:16], run_id=run_id, kind=kind, name=name, content=content))
        s.commit()


def list_attachments(run_id, kind=None) -> list[dict]:
    with Session(engine()) as s:
        q = select(Attachment).where(Attachment.run_id == run_id)
        if kind:
            q = q.where(Attachment.kind == kind)
        return [{"id": a.id, "name": a.name, "kind": a.kind, "size": len(a.content)}
                for a in s.execute(q).scalars().all()]


def get_attachment(run_id, name, kind=None) -> bytes | None:
    with Session(engine()) as s:
        q = select(Attachment).where(Attachment.run_id == run_id, Attachment.name == name)
        if kind:
            q = q.where(Attachment.kind == kind)
        a = s.execute(q).scalars().first()
        return a.content if a else None
