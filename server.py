from contextlib import contextmanager
from datetime import date
from io import BytesIO
import json
import os
from pathlib import Path
import sqlite3

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, Response
from openpyxl import Workbook
from pydantic import BaseModel, Field

ROOT = Path(__file__).resolve().parent
DATABASE = Path(os.environ.get("PAYMENTS_DB", ROOT / "payments.sqlite3"))
app = FastAPI(title="Платёжный контур")

STATUSES = {
    "review": "Новые", "approval": "На согласовании", "approved": "К оплате",
    "queued": "В списке на оплату", "paid": "Оплачено",
}
DEPARTMENTS = ("Администрация", "Операции", "ИТ")
SEED = [
    (1, 0, "Городские сети", 18750, "Интернет и телефония за ноябрь", 24, "review", None),
    (2, 1, "Линия офиса", 36400, "Бумага и картриджи", 25, "review", None),
    (3, 0, "Квартал", 128000, "Аренда офиса за декабрь", 27, "approval", "Аренда"),
    (4, 1, "Счёт и право", 52000, "Сопровождение договора", 23, "approval", "Услуги"),
    (5, 0, "Папирус", 9450, "Расходные материалы", 20, "approved", "Материалы"),
]


class PaymentRequest(BaseModel):
    id: int = Field(gt=0)
    company: str
    recipient: str
    amount: float = Field(gt=0, le=999999999999)
    purpose: str
    due: date
    stage: str
    department: str = "Администрация"
    article: str | None = None
    step: int = Field(default=0, ge=0)
    document: str | None = None
    owner: str | None = None
    returned: bool = False
    batch: str | None = None
    bankOperation: str | None = None
    paidDate: date | None = None
    history: list[str] = Field(default_factory=list)


class Snapshot(BaseModel):
    revision: int
    requests: list[PaymentRequest]


class ExportSelection(BaseModel):
    revision: int
    ids: list[int]


@contextmanager
def database(write=False):
    conn = sqlite3.connect(DATABASE, timeout=10)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("BEGIN IMMEDIATE" if write else "BEGIN")
        yield conn
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        conn.close()


def initialize():
    with database(write=True) as conn:
        conn.execute("CREATE TABLE IF NOT EXISTS state (id INTEGER PRIMARY KEY CHECK(id=1), revision INTEGER NOT NULL)")
        conn.execute("CREATE TABLE IF NOT EXISTS requests (id INTEGER PRIMARY KEY, payload TEXT NOT NULL)")
        conn.execute("""CREATE TABLE IF NOT EXISTS history (
            request_id INTEGER REFERENCES requests(id), position INTEGER,
            description TEXT NOT NULL, recorded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY(request_id, position))""")
        conn.execute("""CREATE TABLE IF NOT EXISTS decisions (
            request_id INTEGER, position INTEGER, decision TEXT NOT NULL, approver TEXT,
            PRIMARY KEY(request_id, position),
            FOREIGN KEY(request_id, position) REFERENCES history(request_id, position))""")
        conn.execute("""CREATE TABLE IF NOT EXISTS payments (
            request_id INTEGER PRIMARY KEY REFERENCES requests(id), amount_cents INTEGER NOT NULL,
            bank_operation TEXT UNIQUE, paid_date TEXT)""")
        if conn.execute("INSERT OR IGNORE INTO state VALUES (1, 0)").rowcount:
            companies = ("ООО «Северный контур»", "ООО «Практика Плюс»")
            for rid, company, recipient, amount, purpose, day, stage, article in SEED:
                row = PaymentRequest(
                    id=rid, company=companies[company], recipient=f"ООО «{recipient}»",
                    amount=amount, purpose=purpose, due=date(2026, 11, day),
                    stage=stage, article=article, step=int(stage == "approved"),
                    department=DEPARTMENTS[(rid - 1) % len(DEPARTMENTS)],
                )
                persist(conn, row, [])


def persist(conn, row, old_history):
    payload = row.model_dump(mode="json", exclude={"history"})
    conn.execute("""INSERT INTO requests VALUES (?, ?)
        ON CONFLICT(id) DO UPDATE SET payload=excluded.payload""",
        (row.id, json.dumps(payload, ensure_ascii=False)))
    for position, entry in enumerate(row.history[len(old_history):], len(old_history)):
        conn.execute("INSERT INTO history(request_id,position,description) VALUES(?,?,?)",
                     (row.id, position, entry))
        if entry.startswith(("Согласовано: ", "Возврат: ")):
            decision, _, detail = entry.partition(": ")
            conn.execute("INSERT INTO decisions VALUES(?,?,?,?)",
                         (row.id, position, decision, detail if decision == "Согласовано" else None))
    if row.stage == "paid":
        conn.execute("""INSERT INTO payments VALUES(?,?,?,?)
            ON CONFLICT(request_id) DO UPDATE SET amount_cents=excluded.amount_cents,
            bank_operation=excluded.bank_operation, paid_date=excluded.paid_date""",
            (row.id, round(row.amount * 100), row.bankOperation, row.paidDate))


def read_snapshot(conn):
    revision = conn.execute("SELECT revision FROM state WHERE id=1").fetchone()[0]
    by_id = {}
    for rid, payload in conn.execute("SELECT id,payload FROM requests ORDER BY id DESC"):
        by_id[rid] = {**json.loads(payload), "history": []}
    for rid, entry in conn.execute("SELECT request_id,description FROM history ORDER BY position"):
        by_id[rid]["history"].append(entry)
    return {"revision": revision, "requests": list(by_id.values())}


initialize()


@app.get("/")
def index():
    return FileResponse(ROOT / "index.html")


@app.get("/api/requests")
def get_requests():
    with database() as conn:
        return read_snapshot(conn)


@app.put("/api/requests")
def update_requests(snapshot: Snapshot):
    try:
        with database(write=True) as conn:
            current = read_snapshot(conn)
            if snapshot.revision != current["revision"]:
                raise HTTPException(409, "Данные изменились в другом окне")
            by_id = {row["id"]: row for row in current["requests"]}
            ids = {row.id for row in snapshot.requests}
            if len(ids) != len(snapshot.requests) or not by_id.keys() <= ids:
                raise HTTPException(422, "Список содержит пропущенные или повторные номера")
            for row in snapshot.requests:
                old = by_id.get(row.id, {})
                history = old.get("history", [])
                if (row.stage not in STATUSES or row.department not in DEPARTMENTS or
                    row.history[:len(history)] != history or
                    (old.get("stage") == "paid" and row.model_dump(mode="json") != old)):
                    raise HTTPException(422, "Недопустимое изменение заявки")
                if old != row.model_dump(mode="json"):
                    persist(conn, row, history)
            conn.execute("UPDATE state SET revision=revision+1 WHERE id=1")
            return read_snapshot(conn)
    except sqlite3.IntegrityError as exc:
        raise HTTPException(409, "Операция уже связана с другой заявкой") from exc


@app.post("/api/registry.xlsx")
def export_registry(selection: ExportSelection):
    with database() as conn:
        snapshot = read_snapshot(conn)
    if selection.revision != snapshot["revision"]:
        raise HTTPException(409, "Обновите реестр")
    ids = set(selection.ids)
    book = Workbook()
    sheet = book.active
    sheet.title = "Реестр заявок"
    sheet.append(["Номер", "Заявка", "Компания", "Подразделение", "Получатель", "Сумма, ₽", "Срок", "Статус"])
    for row in snapshot["requests"]:
        if row["id"] not in ids:
            continue
        status = "На исправлении" if row["returned"] else STATUSES[row["stage"]]
        sheet.append([row["id"], row["purpose"], row["company"], row["department"],
                      row["recipient"], row["amount"], date.fromisoformat(row["due"]), status])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = "s"
        sheet.cell(sheet.max_row, 6).number_format = "#,##0.00"
        sheet.cell(sheet.max_row, 7).number_format = "dd.mm.yyyy"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in zip("ABCDEFGH", (10, 36, 30, 20, 30, 16, 16, 24)):
        sheet.column_dimensions[column].width = width
    output = BytesIO()
    book.save(output)
    return Response(output.getvalue(),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": 'attachment; filename="payment-registry.xlsx"'})


def import_legacy(path):
    rows = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(rows, list) or not rows:
        raise ValueError("Ожидается непустой список заявок")
    with database() as conn:
        snapshot = read_snapshot(conn)
    if snapshot["revision"] != 0:
        raise ValueError("Импорт доступен только до первого изменения новой базы")
    merged = {row["id"]: row for row in snapshot["requests"]}
    imported = set()
    for raw in rows:
        row = PaymentRequest.model_validate(raw)
        if row.id in imported:
            raise ValueError("Повторный номер заявки")
        imported.add(row.id)
        if "department" not in raw and row.id in merged:
            row.department = merged[row.id]["department"]
        merged[row.id] = row.model_dump(mode="json")
    update_requests(Snapshot(revision=0, requests=list(merged.values())))
    return len(imported)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--import-legacy", required=True)
    args = parser.parse_args()
    print(f"Импортировано заявок: {import_legacy(args.import_legacy)}")
