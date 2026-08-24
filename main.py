from __future__ import annotations
import base64
import csv
import hashlib
import hmac
import os
import shutil
import sqlite3
import re

import psycopg
from psycopg.rows import dict_row
import tempfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import flet as ft
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak


APP_DIR = Path(__file__).resolve().parent

APP_ENV = os.getenv("APP_ENV", "desktop").strip().lower()
WEB_MODE = APP_ENV in {"web", "azure", "production"}
APP_HOST = os.getenv("APP_HOST", "0.0.0.0")
APP_PORT = int(
    os.getenv(
        "SERVER_PORT",
        os.getenv("PORT", os.getenv("APP_PORT", "8000"))
    )
)

# Azure App Service persists /home across application restarts/deployments.
# Desktop mode continues using the original folders next to main.py.
if WEB_MODE:
    PERSIST_ROOT = Path(os.getenv("CRM_PERSIST_ROOT", "/home/crm"))
    DATA_DIR = Path(os.getenv("DATA_DIR", str(PERSIST_ROOT / "data")))
    EXPORT_DIR = Path(os.getenv("EXPORT_DIR", str(PERSIST_ROOT / "exports")))
    MIGRATION_DIR = Path(os.getenv("MIGRATION_DIR", str(PERSIST_ROOT / "migration_inbox")))
else:
    DATA_DIR = APP_DIR / "data"
    EXPORT_DIR = APP_DIR / "exports"
    MIGRATION_DIR = APP_DIR / "migration_inbox"

DB_PATH = DATA_DIR / "saudi_sensing_crm.db"
SEED_DB_PATH = APP_DIR / "seed_data" / "saudi_sensing_crm.db"
LOGO_PATH = APP_DIR / "assets" / "saudi_sensing_logo.png"
LOGO_BASE64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")
REPORT_LOGO_PATH = APP_DIR / "assets" / "saudi_sensing_report_logo.png"
VAT_RATE = 15.0


DATABASE_BACKEND = os.getenv("DATABASE_BACKEND", "sqlite").strip().lower()
USE_POSTGRES = DATABASE_BACKEND in {"postgres", "postgresql"}


def _postgres_sql(sql_text: str) -> str:
    """Translate the small SQLite SQL subset used by this CRM to PostgreSQL."""
    converted = sql_text.replace("?", "%s")
    if re.search(r"\bINSERT\s+OR\s+IGNORE\b", converted, flags=re.IGNORECASE):
        converted = re.sub(
            r"\bINSERT\s+OR\s+IGNORE\b",
            "INSERT",
            converted,
            flags=re.IGNORECASE,
        )
        stripped = converted.rstrip()
        suffix = ";" if stripped.endswith(";") else ""
        if suffix:
            stripped = stripped[:-1].rstrip()
        converted = stripped + " ON CONFLICT DO NOTHING" + suffix
    return converted


class PostgresCompatConnection:
    """Tiny compatibility layer so the existing SQLite-style calls keep working."""
    def __init__(self):
        self._con = psycopg.connect(
            host=os.environ["PGHOST"],
            port=int(os.getenv("PGPORT", "5432")),
            dbname=os.getenv("PGDATABASE", "postgres"),
            user=os.environ["PGUSER"],
            password=os.environ["PGPASSWORD"],
            sslmode="require",
            row_factory=dict_row,
        )

    def execute(self, statement, params=None):
        return self._con.execute(
            _postgres_sql(statement),
            tuple(params) if params is not None else None,
        )

    def commit(self):
        return self._con.commit()

    def rollback(self):
        return self._con.rollback()

    def close(self):
        return self._con.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None:
            self._con.commit()
        else:
            self._con.rollback()
        self._con.close()
        return False


def prepare_runtime_storage():
    """Create persistent runtime folders and seed the database on first web launch."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    MIGRATION_DIR.mkdir(parents=True, exist_ok=True)

    if WEB_MODE and not DB_PATH.exists() and SEED_DB_PATH.exists():
        shutil.copy2(SEED_DB_PATH, DB_PATH)

    if WEB_MODE:
        packaged_migration = APP_DIR / "migration_inbox"
        if packaged_migration.exists():
            for src in packaged_migration.glob("*.csv"):
                dst = MIGRATION_DIR / src.name
                if not dst.exists():
                    shutil.copy2(src, dst)


def safe_open_path(path):
    """Desktop helper. On a web server there is no local Explorer to open."""
    if WEB_MODE:
        return False
    try:
        os.startfile(str(path))
        return True
    except Exception:
        return False

ACCOUNT_TYPES = ["End User", "EPC", "OEM", "System Integrator", "Consultant", "Government", "Distributor", "Partner", "Other"]
ACCOUNT_STATUSES = ["Active", "Prospect", "Dormant", "Blocked"]
INDUSTRIES = ["Oil & Gas", "Petrochemical", "Power", "Water", "Mining", "Food & Beverage", "Pharmaceutical", "Cement", "Aviation", "Smart Buildings", "Manufacturing", "Other"]
INTEREST_AREAS = [
    "Instrumentation", "Flow Measurement", "Analyzers / Metering", "Automation / PLC / SCADA",
    "DCS / ESD", "Cybersecurity", "Vibration Monitoring / VMS", "Industrial IoT",
    "Water / RO / Chemical Dosing", "Energy Management", "PoC / Pilot",
    "Service Agreement", "Framework Agreement", "Localization / Manufacturing", "Other"
]
LEAD_STATUSES = ["New", "Contacted", "Qualified", "Nurturing", "Converted", "Disqualified"]
LEAD_SOURCES = ["Referral", "Existing Customer", "Tender", "Event", "Website", "LinkedIn", "Partner", "Cold Outreach", "Saudi Aramco Portal", "Etimad", "Other"]
OPPORTUNITY_STAGES = ["Prospecting", "Qualification", "Technical Evaluation", "PoC / Trial", "Proposal", "Commercial Negotiation", "Customer Approval", "Awarded", "Lost", "On Hold"]
OPPORTUNITY_STATUSES = ["Open", "Won", "Lost", "On Hold"]
QUOTATION_STATUSES = ["NEW REQUEST", "ASSIGNED", "IN PROGRESS", "READY FOR SUBMISSION", "PENDING PROPOSAL MANAGER APPROVAL", "PENDING GM APPROVAL", "APPROVED FOR SUBMISSION", "SUBMITTED", "CUSTOMER REVIEW", "REVISED", "WON", "LOST", "REJECTED", "CANCELLED"]
POC_STATUSES = ["Planned", "Active", "On Hold", "Completed - Successful", "Completed - Unsuccessful", "Cancelled"]
ACTIVITY_TYPES = ["Call", "Email", "Meeting", "Site Visit", "Proposal Follow-up", "Technical Follow-up", "Collection Follow-up", "Partner Follow-up", "Tender Action", "Other"]
ACTIVITY_STATUSES = ["Open", "In Progress", "Completed", "Cancelled"]
PRIORITIES = ["Low", "Medium", "High", "Critical"]
OWNER_ROLES = ["Proposal Engineer", "Sales Engineer", "Sales Manager", "Technical Engineer", "Project Manager", "Finance", "Management", "Other"]
OWNER_STATUSES = ["Active", "Inactive"]

MEETING_TYPES = ["Introductory", "Technical", "Commercial", "Executive", "Site Visit", "PoC Review", "Partner Meeting", "Internal Review", "Other"]

STAGE_PROBABILITY = {
    "Prospecting": 10,
    "Qualification": 25,
    "Technical Evaluation": 40,
    "PoC / Trial": 50,
    "Proposal": 60,
    "Commercial Negotiation": 75,
    "Customer Approval": 90,
    "Awarded": 100,
    "Lost": 0,
    "On Hold": 20,
}


def money(value: float) -> str:
    return f"SAR {float(value):,.2f}"


def safe_float(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    return float(str(value).replace(",", "").strip())


def safe_int(value: Any) -> int:
    return int(round(safe_float(value)))


def valid_date(value: str | None) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    datetime.strptime(value, "%Y-%m-%d")
    return value


def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return f"{salt.hex()}:{digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        salt_hex, digest_hex = stored.split(":", 1)
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(digest_hex)
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
        return hmac.compare_digest(actual, expected)
    except Exception:
        return False


class Database:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def connect(self):
        if USE_POSTGRES:
            return PostgresCompatConnection()

        con = sqlite3.connect(self.path)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA foreign_keys = ON")
        return con

    def initialize(self):
        if USE_POSTGRES:
            # PostgreSQL schema/data are created by the one-time migration step.
            with self.connect() as con:
                row = con.execute(
                    "SELECT COUNT(*) AS count FROM information_schema.tables "
                    "WHERE table_schema='public' AND table_name='users'"
                ).fetchone()
                if not row or int(row["count"]) == 0:
                    raise RuntimeError(
                        "PostgreSQL schema is missing. Run migrate_sqlite_to_postgres.py first."
                    )
            return

        with self.connect() as con:
            con.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'admin',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS owners (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                owner_code TEXT UNIQUE NOT NULL,
                full_name TEXT UNIQUE NOT NULL,
                role TEXT,
                department TEXT,
                email TEXT,
                mobile TEXT,
                can_receive_quotation INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'Active',
                notes TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS id_sequences (
                entity TEXT PRIMARY KEY,
                prefix TEXT NOT NULL,
                next_number INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_reference TEXT UNIQUE,
                account_name TEXT UNIQUE NOT NULL,
                account_type TEXT,
                industry TEXT,
                city TEXT,
                country TEXT,
                website TEXT,
                main_phone TEXT,
                owner TEXT,
                status TEXT,
                notes TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contact_reference TEXT UNIQUE,
                account_id INTEGER,
                first_name TEXT NOT NULL,
                last_name TEXT,
                job_title TEXT,
                department TEXT,
                email TEXT,
                mobile TEXT,
                phone TEXT,
                influence_level TEXT,
                relationship_status TEXT,
                owner TEXT,
                cost_price REAL NOT NULL DEFAULT 0,
                gross_margin_percent REAL NOT NULL DEFAULT 0,
                gross_margin_value REAL NOT NULL DEFAULT 0,
                requested_by TEXT,
                concern_owner TEXT,
                proposal_manager_approval TEXT,
                proposal_manager_approved_by TEXT,
                proposal_manager_approved_at TEXT,
                gm_approval TEXT,
                gm_approved_by TEXT,
                gm_approved_at TEXT,
                approval_comments TEXT,
                assigned_date TEXT,
                work_started_date TEXT,
                submitted_date TEXT,
                completed_date TEXT,
                assignment_notes TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS leads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lead_reference TEXT UNIQUE NOT NULL,
                account_id INTEGER,
                company_name TEXT,
                contact_name TEXT,
                job_title TEXT,
                email TEXT,
                mobile TEXT,
                source TEXT,
                interest_area TEXT,
                lead_status TEXT,
                lead_score REAL NOT NULL DEFAULT 0,
                owner TEXT,
                next_action_date TEXT,
                estimated_value REAL NOT NULL DEFAULT 0,
                notes TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS opportunities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                opportunity_reference TEXT UNIQUE NOT NULL,
                account_id INTEGER,
                opportunity_name TEXT NOT NULL,
                project_type TEXT,
                stage TEXT,
                probability REAL NOT NULL DEFAULT 0,
                estimated_value REAL NOT NULL DEFAULT 0,
                expected_close_date TEXT,
                sales_owner TEXT,
                technical_owner TEXT,
                customer_budget REAL NOT NULL DEFAULT 0,
                competitors TEXT,
                next_step TEXT,
                next_action_date TEXT,
                status TEXT,
                lost_reason TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS quotations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quotation_number TEXT UNIQUE NOT NULL,
                opportunity_id INTEGER,
                account_id INTEGER,
                quotation_date TEXT,
                valid_until TEXT,
                base_value REAL NOT NULL DEFAULT 0,
                discount REAL NOT NULL DEFAULT 0,
                vat_rate REAL NOT NULL DEFAULT 15,
                status TEXT,
                revision INTEGER NOT NULL DEFAULT 0,
                owner TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE SET NULL,
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS pocs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                poc_reference TEXT UNIQUE NOT NULL,
                opportunity_id INTEGER,
                account_id INTEGER,
                poc_title TEXT NOT NULL,
                solution TEXT,
                start_date TEXT,
                planned_end_date TEXT,
                status TEXT,
                success_criteria TEXT,
                estimated_cost REAL NOT NULL DEFAULT 0,
                commercial_value REAL NOT NULL DEFAULT 0,
                owner TEXT,
                outcome TEXT,
                next_step TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE SET NULL,
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS meetings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                meeting_reference TEXT UNIQUE NOT NULL,
                account_id INTEGER,
                opportunity_id INTEGER,
                meeting_date TEXT,
                meeting_type TEXT,
                subject TEXT,
                location TEXT,
                attendees TEXT,
                owner TEXT,
                outcome TEXT,
                next_action TEXT,
                next_action_date TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE SET NULL,
                FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE SET NULL
            );

            CREATE TABLE IF NOT EXISTS quotation_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quotation_id INTEGER NOT NULL,
                document_type TEXT NOT NULL,
                original_file_name TEXT NOT NULL,
                stored_file_path TEXT NOT NULL,
                uploaded_by TEXT,
                uploaded_at TEXT NOT NULL,
                notes TEXT,
                FOREIGN KEY(quotation_id) REFERENCES quotations(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS inbox_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                owner_name TEXT NOT NULL,
                item_type TEXT NOT NULL,
                reference_id TEXT NOT NULL,
                title TEXT NOT NULL,
                message TEXT,
                action_type TEXT,
                related_record_id INTEGER,
                status TEXT NOT NULL DEFAULT 'OPEN',
                created_at TEXT NOT NULL,
                completed_at TEXT
            );

            CREATE TABLE IF NOT EXISTS notifications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                owner_name TEXT NOT NULL,
                title TEXT NOT NULL,
                message TEXT,
                related_type TEXT,
                related_record_id INTEGER,
                is_read INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS activities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                activity_reference TEXT UNIQUE NOT NULL,
                account_id INTEGER,
                opportunity_id INTEGER,
                activity_type TEXT,
                subject TEXT,
                due_date TEXT,
                priority TEXT,
                status TEXT,
                owner TEXT,
                completed_date TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(account_id) REFERENCES accounts(id) ON DELETE SET NULL,
                FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE SET NULL
            );


            CREATE TABLE IF NOT EXISTS legacy_opportunity_details (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                opportunity_id INTEGER UNIQUE,
                source_row INTEGER,
                created_date TEXT,
                customer_name_source TEXT,
                account_basis TEXT,
                business_unit TEXT,
                source_project_type TEXT,
                end_user TEXT,
                industry TEXT,
                competitive INTEGER,
                probability_band TEXT,
                source_currency TEXT,
                forecast_gm_percent REAL,
                forecast_gm_value REAL,
                gm_value_basis TEXT,
                expected_po_year INTEGER,
                expected_po_month TEXT,
                quarter TEXT,
                delivery_date TEXT,
                created_by TEXT,
                assigned_to TEXT,
                include_in_forecast INTEGER,
                source_stage TEXT,
                opportunity_update TEXT,
                must_win INTEGER,
                suspended INTEGER,
                quality_flags TEXT,
                FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_opportunities_stage ON opportunities(stage,status);
            CREATE INDEX IF NOT EXISTS idx_activities_due ON activities(due_date,status);
            CREATE INDEX IF NOT EXISTS idx_leads_status ON leads(lead_status);
            """)

            self.ensure_schema_upgrades(con)
            self.initialize_sequences(con)
            self.ensure_default_owners(con)

            if con.execute("SELECT COUNT(*) AS count FROM users").fetchone()["count"] == 0:
                con.execute(
                    """INSERT INTO users
                    (username,password_hash,full_name,role,active,created_at)
                    VALUES (?,?,?,?,1,?)""",
                    ("admin", hash_password("admin123"), "System Administrator", "admin",
                     datetime.now().isoformat(timespec="seconds"))
                )

    ALLOWED_EDIT_TABLES = {
        "accounts", "contacts", "leads", "opportunities", "quotations",
        "pocs", "meetings", "activities", "owners"
    }

    def raw_record(self, table, record_id):
        if table not in self.ALLOWED_EDIT_TABLES:
            raise ValueError("This table is not editable.")
        return self.one(f"SELECT * FROM {table} WHERE id=?", (record_id,))

    def table_columns(self, table):
        if table not in self.ALLOWED_EDIT_TABLES:
            raise ValueError("This table is not editable.")
        with self.connect() as con:
            if USE_POSTGRES:
                rows = con.execute(
                    """
                    SELECT column_name AS name
                    FROM information_schema.columns
                    WHERE table_schema='public' AND table_name=?
                    ORDER BY ordinal_position
                    """,
                    (table,),
                ).fetchall()
                return [dict(row) for row in rows]
            return [dict(row) for row in con.execute(f"PRAGMA table_info({table})").fetchall()]

    def update_record(self, table, record_id, values):
        if table not in self.ALLOWED_EDIT_TABLES:
            raise ValueError("This table is not editable.")
        allowed_columns = {
            row["name"] for row in self.table_columns(table)
            if row["name"] not in {"id", "created_at"}
        }
        clean_values = {
            key: value for key, value in values.items()
            if key in allowed_columns
        }
        if not clean_values:
            return
        assignments = ", ".join(f"{column}=?" for column in clean_values)
        params = list(clean_values.values()) + [record_id]
        with self.connect() as con:
            con.execute(
                f"UPDATE {table} SET {assignments} WHERE id=?",
                params,
            )

    def delete_record(self, table, record_id):
        if table not in self.ALLOWED_EDIT_TABLES:
            raise ValueError("This table is not editable.")
        with self.connect() as con:
            con.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))

    SEQUENCE_CONFIG = {
        "accounts": ("ACC", "account_reference"),
        "contacts": ("CON", "contact_reference"),
        "leads": ("LD", "lead_reference"),
        "opportunities": ("OPP", "opportunity_reference"),
        "quotations": ("QTN", "quotation_number"),
        "meetings": ("MTG", "meeting_reference"),
        "activities": ("ACT", "activity_reference"),
        "owners": ("OWN", "owner_code"),
    }

    def ensure_schema_upgrades(self, con):
        def columns(table):
            return {
                row["name"]
                for row in con.execute(f"PRAGMA table_info({table})").fetchall()
            }

        if "account_reference" not in columns("accounts"):
            con.execute("ALTER TABLE accounts ADD COLUMN account_reference TEXT")
            con.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_accounts_reference "
                "ON accounts(account_reference)"
            )

        if "contact_reference" not in columns("contacts"):
            con.execute("ALTER TABLE contacts ADD COLUMN contact_reference TEXT")
            con.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_contacts_reference "
                "ON contacts(contact_reference)"
            )

        legacy_columns = columns("legacy_opportunity_details")
        if "original_opportunity_reference" not in legacy_columns:
            con.execute(
                "ALTER TABLE legacy_opportunity_details "
                "ADD COLUMN original_opportunity_reference TEXT"
            )

        quotation_columns = columns("quotations")
        quotation_upgrades = {
            "assigned_date": "TEXT",
            "work_started_date": "TEXT",
            "submitted_date": "TEXT",
            "completed_date": "TEXT",
            "assignment_notes": "TEXT",
        }
        for column_name, column_type in quotation_upgrades.items():
            if column_name not in quotation_columns:
                con.execute(
                    f"ALTER TABLE quotations ADD COLUMN {column_name} {column_type}"
                )

        quotation_columns = columns("quotations")
        commercial_upgrades = {
            "cost_price": "REAL NOT NULL DEFAULT 0",
            "gross_margin_percent": "REAL NOT NULL DEFAULT 0",
            "gross_margin_value": "REAL NOT NULL DEFAULT 0",
            "requested_by": "TEXT",
            "concern_owner": "TEXT",
            "proposal_manager_approval": "TEXT",
            "proposal_manager_approved_by": "TEXT",
            "proposal_manager_approved_at": "TEXT",
            "gm_approval": "TEXT",
            "gm_approved_by": "TEXT",
            "gm_approved_at": "TEXT",
            "approval_comments": "TEXT",
        }
        for column_name, column_type in commercial_upgrades.items():
            if column_name not in quotation_columns:
                con.execute(
                    f"ALTER TABLE quotations ADD COLUMN {column_name} {column_type}"
                )

        lead_columns = columns("leads")
        if "account_id" not in lead_columns:
            con.execute("ALTER TABLE leads ADD COLUMN account_id INTEGER")


    def initialize_sequences(self, con):
        for entity, (prefix, _) in self.SEQUENCE_CONFIG.items():
            con.execute(
                """INSERT OR IGNORE INTO id_sequences(entity,prefix,next_number)
                   VALUES (?,?,1)""",
                (entity, prefix),
            )

        # Assign references to existing accounts and contacts.
        for entity in ("accounts", "contacts"):
            prefix, column = self.SEQUENCE_CONFIG[entity]
            rows = con.execute(
                f"SELECT id,{column} FROM {entity} ORDER BY id"
            ).fetchall()
            next_number = 1
            for row in rows:
                if not row[column]:
                    reference = f"{prefix}-{next_number:05d}"
                    con.execute(
                        f"UPDATE {entity} SET {column}=? WHERE id=?",
                        (reference, row["id"]),
                    )
                next_number += 1
            con.execute(
                "UPDATE id_sequences SET next_number=? WHERE entity=?",
                (max(next_number, 1), entity),
            )

        # Preserve and normalize existing opportunity references.
        opportunity_rows = con.execute(
            """SELECT o.id,o.opportunity_reference,l.original_opportunity_reference
               FROM opportunities o
               LEFT JOIN legacy_opportunity_details l ON l.opportunity_id=o.id
               ORDER BY o.id"""
        ).fetchall()
        for sequence_number, row in enumerate(opportunity_rows, start=1):
            old_reference = row["opportunity_reference"] or ""
            if row["original_opportunity_reference"] in (None, ""):
                con.execute(
                    """UPDATE legacy_opportunity_details
                       SET original_opportunity_reference=?
                       WHERE opportunity_id=?""",
                    (old_reference, row["id"]),
                )
            con.execute(
                "UPDATE opportunities SET opportunity_reference=? WHERE id=?",
                (f"OPP-{sequence_number:05d}", row["id"]),
            )
        con.execute(
            "UPDATE id_sequences SET next_number=? WHERE entity='opportunities'",
            (len(opportunity_rows) + 1,),
        )

        # Normalize existing references for the remaining sequenced tables.
        for entity in ("leads", "quotations", "meetings", "activities"):
            prefix, column = self.SEQUENCE_CONFIG[entity]
            rows = con.execute(
                f"SELECT id,{column} FROM {entity} ORDER BY id"
            ).fetchall()
            for sequence_number, row in enumerate(rows, start=1):
                con.execute(
                    f"UPDATE {entity} SET {column}=? WHERE id=?",
                    (f"{prefix}-{sequence_number:05d}", row["id"]),
                )
            con.execute(
                "UPDATE id_sequences SET next_number=? WHERE entity=?",
                (len(rows) + 1, entity),
            )

    def ensure_default_owners(self, con):
        predefined_owners = [
            ("OWN-00001", "Abdelaziz Mohamed", "Proposal Engineer", "Proposals", 1),
            ("OWN-00002", "Bahaa Abusaqer", "Project Manager", "Projects", 0),
            ("OWN-00003", "Basith Ahmed", "Proposal Manager", "Proposals", 1),
            ("OWN-00004", "Hala Wutayd", "Sales Admin", "Sales", 0),
            ("OWN-00005", "Hussain AlHaydar", "Sales Engineer", "Sales", 0),
            ("OWN-00006", "Khalid Al Sairy", "Proposal Engineer", "Proposals", 1),
            ("OWN-00007", "Rania Alamim", "Proposal Engineer", "Proposals", 1),
            ("OWN-00008", "Saud Al Shammari", "General Manager", "Management", 1),
            ("OWN-00009", "Shady Moussa", "Sales Engineer", "Sales", 0),
            ("OWN-00010", "Syed Imran", "Operation Head, Valves and Services", "Operations", 0),
        ]

        existing_count = con.execute("SELECT COUNT(*) AS count FROM owners").fetchone()["count"]
        if existing_count == 0:
            for owner_code, full_name, role, department, quotation_flag in predefined_owners:
                con.execute(
                    """INSERT INTO owners
                    (owner_code,full_name,role,department,email,mobile,
                     can_receive_quotation,status,notes,created_at)
                    VALUES (?,?,?,?,?,?,?,'Active','',?)""",
                    (
                        owner_code, full_name, role, department, "", "",
                        quotation_flag,
                        datetime.now().isoformat(timespec="seconds"),
                    ),
                )
            con.execute(
                "UPDATE id_sequences SET next_number=11 WHERE entity='owners'"
            )


    def next_reference(self, entity, con=None):
        if entity not in self.SEQUENCE_CONFIG:
            raise ValueError(f"Unknown sequence entity: {entity}")

        owns_connection = con is None
        connection = con or self.connect()
        try:
            row = connection.execute(
                "SELECT prefix,next_number FROM id_sequences WHERE entity=?",
                (entity,),
            ).fetchone()
            if not row:
                prefix = self.SEQUENCE_CONFIG[entity][0]
                connection.execute(
                    """INSERT INTO id_sequences(entity,prefix,next_number)
                       VALUES (?,?,1)""",
                    (entity, prefix),
                )
                number = 1
            else:
                prefix = row["prefix"]
                number = int(row["next_number"])

            reference = f"{prefix}-{number:05d}"
            connection.execute(
                "UPDATE id_sequences SET next_number=? WHERE entity=?",
                (number + 1, entity),
            )
            if owns_connection:
                connection.commit()
            return reference
        finally:
            if owns_connection:
                connection.close()

    def consume_explicit_reference(self, entity, reference, con):
        current = con.execute(
            "SELECT prefix,next_number FROM id_sequences WHERE entity=?",
            (entity,),
        ).fetchone()
        if not current:
            return
        expected = f"{current['prefix']}-{int(current['next_number']):05d}"
        if reference == expected:
            con.execute(
                "UPDATE id_sequences SET next_number=? WHERE entity=?",
                (int(current["next_number"]) + 1, entity),
            )

    def peek_reference(self, entity):
        row = self.one(
            "SELECT prefix,next_number FROM id_sequences WHERE entity=?",
            (entity,),
        )
        if not row:
            prefix = self.SEQUENCE_CONFIG[entity][0]
            return f"{prefix}-00001"
        return f"{row['prefix']}-{int(row['next_number']):05d}"

    def owners(self, active_only=False, quotation_only=False):
        query = "SELECT * FROM owners WHERE 1=1"
        params = []
        if active_only:
            query += " AND status='Active'"
        if quotation_only:
            query += " AND can_receive_quotation=1 AND status='Active'"
        query += " ORDER BY owner_code"
        return self.rows(query, params)

    def insert_owner(self, data):
        with self.connect() as con:
            owner_code = self.next_reference("owners", con=con)
            con.execute(
                """INSERT INTO owners
                (owner_code,full_name,role,department,email,mobile,
                 can_receive_quotation,status,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (
                    owner_code, data["full_name"], data["role"], data["department"],
                    data["email"], data["mobile"],
                    1 if data["can_receive_quotation"] else 0,
                    data["status"], data["notes"],
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )

    def add_notification(
        self,
        owner_name,
        title,
        message,
        related_type="",
        related_record_id=None,
    ):
        if not owner_name:
            return
        with self.connect() as con:
            con.execute(
                """INSERT INTO notifications
                (owner_name,title,message,related_type,related_record_id,is_read,created_at)
                VALUES (?,?,?,?,?,0,?)""",
                (
                    owner_name, title, message, related_type,
                    related_record_id,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )

    def notifications_for_owner(self, owner_name, unread_only=False):
        query = "SELECT * FROM notifications WHERE owner_name=?"
        params = [owner_name]
        if unread_only:
            query += " AND is_read=0"
        query += " ORDER BY id DESC"
        return self.rows(query, params)

    def unread_notification_count(self, owner_name):
        if not owner_name:
            return 0
        row = self.one(
            "SELECT COUNT(*) AS count FROM notifications WHERE owner_name=? AND is_read=0",
            (owner_name,),
        )
        return int(row["count"] if row else 0)

    def mark_notification_read(self, notification_id):
        with self.connect() as con:
            con.execute(
                "UPDATE notifications SET is_read=1 WHERE id=?",
                (notification_id,),
            )

    def add_inbox_item(
        self,
        owner_name,
        item_type,
        reference_id,
        title,
        message,
        action_type,
        related_record_id,
    ):
        if not owner_name:
            return
        with self.connect() as con:
            con.execute(
                """INSERT INTO inbox_items
                (owner_name,item_type,reference_id,title,message,action_type,
                 related_record_id,status,created_at)
                VALUES (?,?,?,?,?,?,?,'OPEN',?)""",
                (
                    owner_name, item_type, reference_id, title, message,
                    action_type, related_record_id,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )

    def inbox_for_owner(self, owner_name):
        return self.rows(
            """SELECT * FROM inbox_items
               WHERE owner_name=? AND status='OPEN'
               ORDER BY id DESC""",
            (owner_name,),
        )

    def close_inbox_items(self, related_record_id, action_type=None):
        query = """UPDATE inbox_items
                   SET status='COMPLETED',completed_at=?
                   WHERE related_record_id=? AND status='OPEN'"""
        params = [datetime.now().isoformat(timespec="seconds"), related_record_id]
        if action_type:
            query += " AND action_type=?"
            params.append(action_type)
        with self.connect() as con:
            con.execute(query, params)

    def add_quotation_document(
        self,
        quotation_id,
        document_type,
        original_file_name,
        stored_file_path,
        uploaded_by,
        notes="",
    ):
        with self.connect() as con:
            con.execute(
                """INSERT INTO quotation_documents
                (quotation_id,document_type,original_file_name,stored_file_path,
                 uploaded_by,uploaded_at,notes)
                VALUES (?,?,?,?,?,?,?)""",
                (
                    quotation_id, document_type, original_file_name,
                    stored_file_path, uploaded_by,
                    datetime.now().isoformat(timespec="seconds"),
                    notes,
                ),
            )

    def quotation_documents(self, quotation_id):
        return self.rows(
            """SELECT * FROM quotation_documents
               WHERE quotation_id=?
               ORDER BY id DESC""",
            (quotation_id,),
        )

    def authenticate(self, username, password):
        with self.connect() as con:
            row = con.execute(
                "SELECT * FROM users WHERE username=? AND active=1",
                (username.strip(),)
            ).fetchone()
        return dict(row) if row and verify_password(password, row["password_hash"]) else None

    def rows(self, query, params=()):
        with self.connect() as con:
            return [dict(r) for r in con.execute(query, params).fetchall()]

    def one(self, query, params=()):
        with self.connect() as con:
            row = con.execute(query, params).fetchone()
        return dict(row) if row else None

    def account_id(self, name):
        if not name:
            return None
        row = self.one("SELECT id FROM accounts WHERE lower(account_name)=lower(?)", (name.strip(),))
        return row["id"] if row else None

    def opportunity_id(self, reference):
        if not reference:
            return None
        row = self.one("SELECT id FROM opportunities WHERE lower(opportunity_reference)=lower(?)", (reference.strip(),))
        return row["id"] if row else None

    def insert_account(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            con.execute(
                f"""{sql} INTO accounts
                (account_reference,account_name,account_type,industry,city,country,website,main_phone,owner,status,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (self.next_reference("accounts", con=con),
                 data["account_name"],data["account_type"],data["industry"],data["city"],data["country"],
                 data["website"],data["main_phone"],data["owner"],data["status"],data["notes"],
                 datetime.now().isoformat(timespec="seconds"))
            )

    def insert_contact(self, data):
        with self.connect() as con:
            con.execute(
                """INSERT INTO contacts
                (contact_reference,account_id,first_name,last_name,job_title,department,email,mobile,phone,
                 influence_level,relationship_status,owner,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (self.next_reference("contacts", con=con),
                 data["account_id"],data["first_name"],data["last_name"],data["job_title"],
                 data["department"],data["email"],data["mobile"],data["phone"],data["influence_level"],
                 data["relationship_status"],data["owner"],data["notes"],
                 datetime.now().isoformat(timespec="seconds"))
            )

    def insert_lead(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            reference = data.get("lead_reference") or self.next_reference("leads", con=con)
            self.consume_explicit_reference("leads", reference, con)
            con.execute(
                f"""{sql} INTO leads
                (lead_reference,account_id,company_name,contact_name,job_title,email,mobile,
                 source,interest_area,lead_status,lead_score,owner,next_action_date,
                 estimated_value,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    reference,data.get("account_id"),data["company_name"],
                    data["contact_name"],data["job_title"],data["email"],
                    data["mobile"],data["source"],data["interest_area"],
                    data["lead_status"],data["lead_score"],data["owner"],
                    data["next_action_date"],data["estimated_value"],
                    data["notes"],datetime.now().isoformat(timespec="seconds")
                )
            )

    def insert_opportunity(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            reference = data.get("opportunity_reference") or self.next_reference("opportunities", con=con)
            self.consume_explicit_reference("opportunities", reference, con)
            con.execute(
                f"""{sql} INTO opportunities
                (opportunity_reference,account_id,opportunity_name,project_type,stage,probability,
                 estimated_value,expected_close_date,sales_owner,technical_owner,customer_budget,
                 competitors,next_step,next_action_date,status,lost_reason,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (reference,data["account_id"],data["opportunity_name"],
                 data["project_type"],data["stage"],data["probability"],data["estimated_value"],
                 data["expected_close_date"],data["sales_owner"],data["technical_owner"],
                 data["customer_budget"],data["competitors"],data["next_step"],data["next_action_date"],
                 data["status"],data["lost_reason"],data["notes"],
                 datetime.now().isoformat(timespec="seconds"))
            )

    def insert_quotation(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            reference = data.get("quotation_number") or self.next_reference("quotations", con=con)
            self.consume_explicit_reference("quotations", reference, con)
            con.execute(
                f"""{sql} INTO quotations
                (quotation_number,opportunity_id,account_id,quotation_date,valid_until,base_value,
                 discount,vat_rate,status,revision,owner,cost_price,gross_margin_percent,
                 gross_margin_value,requested_by,concern_owner,proposal_manager_approval,
                 proposal_manager_approved_by,proposal_manager_approved_at,gm_approval,
                 gm_approved_by,gm_approved_at,approval_comments,assigned_date,
                 work_started_date,submitted_date,completed_date,assignment_notes,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    reference,data["opportunity_id"],data["account_id"],data["quotation_date"],
                    data["valid_until"],data["base_value"],data["discount"],data["vat_rate"],
                    data["status"],data["revision"],data["owner"],
                    data.get("cost_price",0),data.get("gross_margin_percent",0),
                    data.get("gross_margin_value",0),data.get("requested_by",""),
                    data.get("concern_owner",""),data.get("proposal_manager_approval",""),
                    data.get("proposal_manager_approved_by",""),
                    data.get("proposal_manager_approved_at",""),data.get("gm_approval",""),
                    data.get("gm_approved_by",""),data.get("gm_approved_at",""),
                    data.get("approval_comments",""),data.get("assigned_date",""),
                    data.get("work_started_date",""),data.get("submitted_date",""),
                    data.get("completed_date",""),data.get("assignment_notes",""),
                    data["notes"],datetime.now().isoformat(timespec="seconds")
                )
            )

    def insert_poc(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            con.execute(
                f"""{sql} INTO pocs
                (poc_reference,opportunity_id,account_id,poc_title,solution,start_date,planned_end_date,
                 status,success_criteria,estimated_cost,commercial_value,owner,outcome,next_step,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (data["poc_reference"],data["opportunity_id"],data["account_id"],data["poc_title"],
                 data["solution"],data["start_date"],data["planned_end_date"],data["status"],
                 data["success_criteria"],data["estimated_cost"],data["commercial_value"],data["owner"],
                 data["outcome"],data["next_step"],data["notes"],
                 datetime.now().isoformat(timespec="seconds"))
            )

    def insert_meeting(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            reference = data.get("meeting_reference") or self.next_reference("meetings", con=con)
            self.consume_explicit_reference("meetings", reference, con)
            con.execute(
                f"""{sql} INTO meetings
                (meeting_reference,account_id,opportunity_id,meeting_date,meeting_type,subject,location,
                 attendees,owner,outcome,next_action,next_action_date,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (reference,data["account_id"],data["opportunity_id"],data["meeting_date"],
                 data["meeting_type"],data["subject"],data["location"],data["attendees"],data["owner"],
                 data["outcome"],data["next_action"],data["next_action_date"],data["notes"],
                 datetime.now().isoformat(timespec="seconds"))
            )

    def insert_activity(self, data, ignore=False):
        sql = "INSERT OR IGNORE" if ignore else "INSERT"
        with self.connect() as con:
            reference = data.get("activity_reference") or self.next_reference("activities", con=con)
            self.consume_explicit_reference("activities", reference, con)
            con.execute(
                f"""{sql} INTO activities
                (activity_reference,account_id,opportunity_id,activity_type,subject,due_date,priority,
                 status,owner,completed_date,notes,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (reference,data["account_id"],data["opportunity_id"],data["activity_type"],
                 data["subject"],data["due_date"],data["priority"],data["status"],data["owner"],
                 data["completed_date"],data["notes"],datetime.now().isoformat(timespec="seconds"))
            )

    def accounts(self):
        return self.rows("SELECT * FROM accounts ORDER BY account_reference")

    def contacts(self):
        return self.rows("""SELECT c.*,a.account_name FROM contacts c
                            LEFT JOIN accounts a ON a.id=c.account_id
                            ORDER BY c.contact_reference""")

    def lead_full_details(self, lead_id):
        return self.one(
            """SELECT l.*,a.account_name,a.account_reference,a.account_type,
                      a.industry AS account_industry,a.city AS account_city,
                      a.country AS account_country
               FROM leads l
               LEFT JOIN accounts a ON a.id=l.account_id
               WHERE l.id=?""",
            (lead_id,),
        )

    def leads(self):
        return self.rows(
            """SELECT l.*,a.account_name
               FROM leads l
               LEFT JOIN accounts a ON a.id=l.account_id
               ORDER BY l.lead_reference"""
        )

    def opportunities(self):
        return self.rows("""SELECT o.*,a.account_name FROM opportunities o
                            LEFT JOIN accounts a ON a.id=o.account_id
                            ORDER BY o.opportunity_reference""")

    def opportunity_full_details(self, opportunity_id):
        return self.one(
            """SELECT
                   o.*,
                   a.account_name,
                   a.account_reference,
                   a.account_type,
                   a.industry AS account_industry,
                   a.city AS account_city,
                   a.country AS account_country,
                   l.source_row,
                   l.created_date AS source_created_date,
                   l.customer_name_source,
                   l.account_basis,
                   l.business_unit,
                   l.source_project_type,
                   l.end_user,
                   l.industry AS source_industry,
                   l.competitive,
                   l.probability_band,
                   l.source_currency,
                   l.forecast_gm_percent,
                   l.forecast_gm_value,
                   l.gm_value_basis,
                   l.expected_po_year,
                   l.expected_po_month,
                   l.quarter,
                   l.delivery_date,
                   l.created_by,
                   l.assigned_to,
                   l.include_in_forecast,
                   l.source_stage,
                   l.opportunity_update,
                   l.must_win,
                   l.suspended,
                   l.quality_flags,
                   l.original_opportunity_reference
               FROM opportunities o
               LEFT JOIN accounts a ON a.id=o.account_id
               LEFT JOIN legacy_opportunity_details l ON l.opportunity_id=o.id
               WHERE o.id=?""",
            (opportunity_id,),
        )

    def legacy_opportunities(self):
        return self.rows("""SELECT l.*,o.opportunity_reference,o.opportunity_name,o.estimated_value,
                            o.probability,o.stage,o.status,o.sales_owner,a.account_name
                            FROM legacy_opportunity_details l
                            JOIN opportunities o ON o.id=l.opportunity_id
                            LEFT JOIN accounts a ON a.id=o.account_id
                            ORDER BY l.source_row""")

    def quotations(self):
        return self.rows("""SELECT q.*,a.account_name,o.opportunity_reference,o.opportunity_name
                            FROM quotations q
                            LEFT JOIN accounts a ON a.id=q.account_id
                            LEFT JOIN opportunities o ON o.id=q.opportunity_id
                            ORDER BY q.quotation_number""")

    def quotation_for_opportunity(self, opportunity_id, status=None):
        query = "SELECT * FROM quotations WHERE opportunity_id=?"
        params = [opportunity_id]
        if status:
            query += " AND status=?"
            params.append(status)
        query += " ORDER BY id DESC LIMIT 1"
        return self.one(query, params)

    def pocs(self):
        return self.rows("""SELECT p.*,a.account_name,o.opportunity_reference
                            FROM pocs p
                            LEFT JOIN accounts a ON a.id=p.account_id
                            LEFT JOIN opportunities o ON o.id=p.opportunity_id
                            ORDER BY p.id DESC""")

    def meetings(self):
        return self.rows("""SELECT m.*,a.account_name,o.opportunity_reference
                            FROM meetings m
                            LEFT JOIN accounts a ON a.id=m.account_id
                            LEFT JOIN opportunities o ON o.id=m.opportunity_id
                            ORDER BY m.meeting_date DESC,m.id DESC""")

    def activities(self):
        return self.rows("""SELECT x.*,a.account_name,o.opportunity_reference
                            FROM activities x
                            LEFT JOIN accounts a ON a.id=x.account_id
                            LEFT JOIN opportunities o ON o.id=x.opportunity_id
                            ORDER BY x.due_date,x.id DESC""")

    def dashboard(self):
        accounts = self.accounts()
        leads = self.leads()
        opps = self.opportunities()
        quotes = self.quotations()
        pocs = self.pocs()
        activities = self.activities()

        open_opps = [o for o in opps if o["status"] == "Open"]
        pipeline = sum(float(o["estimated_value"]) for o in open_opps)
        weighted = sum(float(o["estimated_value"]) * float(o["probability"]) / 100 for o in open_opps)
        won = [o for o in opps if o["status"] == "Won" or o["stage"] == "Awarded"]
        lost = [o for o in opps if o["status"] == "Lost" or o["stage"] == "Lost"]
        win_rate = len(won) / (len(won) + len(lost)) * 100 if won or lost else 0
        converted = len([l for l in leads if l["lead_status"] == "Converted"])
        lead_conversion = converted / len(leads) * 100 if leads else 0
        quote_value = sum(max(0, float(q["base_value"]) - float(q["discount"])) for q in quotes if q["status"] not in ("Rejected","Expired","Cancelled"))
        active_pocs = len([p for p in pocs if p["status"] in ("Planned","Active","On Hold")])
        today = date.today()
        overdue = []
        due_7 = []
        for a in activities:
            if not a["due_date"] or a["status"] in ("Completed","Cancelled"):
                continue
            try:
                d = datetime.strptime(a["due_date"], "%Y-%m-%d").date()
            except ValueError:
                continue
            if d < today:
                overdue.append(a)
            elif d <= today + timedelta(days=7):
                due_7.append(a)

        return {
            "accounts": len(accounts), "leads": len(leads), "open_opportunities": len(open_opps),
            "pipeline": pipeline, "weighted_pipeline": weighted, "win_rate": win_rate,
            "lead_conversion": lead_conversion, "quotation_value": quote_value,
            "active_pocs": active_pocs, "overdue_activities": len(overdue),
            "due_7_days": len(due_7), "overdue_rows": overdue[:10],
        }


class CRMApp:
    def __init__(self, page: ft.Page):
        self.page = page
        prepare_runtime_storage()
        self.db = Database(DB_PATH)
        self.db.initialize()
        self.user = None
        self.current_view = "dashboard"
        self.dashboard_filters = {
            "year": "All",
            "month": "All",
            "account": "All",
            "sales_owner": "All",
            "stage": "All",
            "business_unit": "All",
            "status": "Open",
        }

        self.page.title = "Saudi Sensing CRM | A Tamimi Energy Company"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.theme = ft.Theme(font_family="Calibri")
        self.page.bgcolor = "#F3F7F6"
        self.page.padding = 0
        if not WEB_MODE:
            self.page.window.width = 1500
            self.page.window.height = 900
            self.page.window.min_width = 1150
            self.page.window.min_height = 720
        self.show_login()

    def notify(self, message, error=False):
        snack = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE),
            bgcolor=ft.Colors.RED_700 if error else ft.Colors.GREEN_700,
        )
        try:
            self.page.open(snack)
        except Exception:
            self.page.overlay.append(snack)
            snack.open = True
            self.page.update()

    def show_login(self):
        username = ft.TextField(label="Username", value="admin", width=390)
        password = ft.TextField(label="Password", password=True, can_reveal_password=True, width=390)

        def login(_):
            user = self.db.authenticate(username.value or "", password.value or "")
            if not user:
                self.notify("Invalid username or password.", True)
                return
            self.user = user
            self.show_main()

        card = ft.Container(
            width=520, padding=40, bgcolor=ft.Colors.WHITE, border_radius=24,
            shadow=ft.BoxShadow(blur_radius=35, color=ft.Colors.with_opacity(0.12, ft.Colors.BLACK)),
            content=ft.Column(
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Container(
                        width=300,
                        height=145,
                        alignment=ft.Alignment(0, 0),
                        content=ft.Image(
                          src_base64=LOGO_BASE64,
                            fit=ft.ImageFit.CONTAIN,
                        ),
                    ),
                    ft.Text("Saudi Sensing CRM", size=30, weight=ft.FontWeight.BOLD, color="#103F37"),
                    ft.Text("Customers, leads, pipeline, PoCs and sales execution", color="#657A80"),
                    ft.Container(height=10), username, password,
                    ft.FilledButton("Sign In", icon=ft.Icons.LOGIN, width=390, height=50,
                                    on_click=login, style=ft.ButtonStyle(bgcolor="#0E5A47")),
                    ft.Text("Default login: admin / admin123", size=12, color="#7B8D91"),
                ]
            )
        )
        self.page.clean()
        self.page.add(ft.Container(expand=True, alignment=ft.Alignment(0,0), content=card))
        self.page.update()

    def show_main(self):
        try:
            self.content = ft.Container(expand=True, padding=24)
            sidebar_control = self.sidebar()
            self.page.clean()
            unread_count = self.db.unread_notification_count(
                self.current_owner_name()
            )
            inbox_count = len(
                self.db.inbox_for_owner(self.current_owner_name())
            )
            top_area = ft.Column(
                expand=True,
                spacing=0,
                controls=[
                    ft.Container(
                        height=62,
                        padding=ft.Padding(
                            left=18,
                            right=18,
                            top=7,
                            bottom=7,
                        ),
                        bgcolor=ft.Colors.WHITE,
                        border=ft.Border(
                            bottom=ft.BorderSide(1, "#D7E3DF")
                        ),
                        content=ft.Row(
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            controls=[
                                ft.Text(
                                    "Saudi Sensing CRM",
                                    size=18,
                                    weight=ft.FontWeight.BOLD,
                                    color="#103F37",
                                ),
                                ft.Row(
                                    spacing=8,
                                    controls=[
                                        ft.TextButton(
                                            content=ft.Row(
                                                spacing=5,
                                                controls=[
                                                    ft.Icon(
                                                        ft.Icons.INBOX,
                                                        color="#103F37",
                                                    ),
                                                    ft.Text(
                                                        "Inbox",
                                                        color="#103F37",
                                                        weight=ft.FontWeight.BOLD,
                                                    ),
                                                    ft.Container(
                                                        padding=5,
                                                        bgcolor=(
                                                            "#1667A8"
                                                            if inbox_count
                                                            else "#71827E"
                                                        ),
                                                        border_radius=10,
                                                        content=ft.Text(
                                                            str(inbox_count),
                                                            color=ft.Colors.WHITE,
                                                            size=11,
                                                            weight=ft.FontWeight.BOLD,
                                                        ),
                                                    ),
                                                ],
                                            ),
                                            on_click=lambda _: self.navigate("inbox"),
                                        ),
                                        ft.TextButton(
                                            content=ft.Row(
                                                spacing=5,
                                                controls=[
                                                    ft.Icon(
                                                        ft.Icons.NOTIFICATIONS,
                                                        color="#103F37",
                                                    ),
                                                    ft.Text(
                                                        "Notifications",
                                                        color="#103F37",
                                                        weight=ft.FontWeight.BOLD,
                                                    ),
                                                    ft.Container(
                                                        padding=5,
                                                        bgcolor=(
                                                            "#C0392B"
                                                            if unread_count
                                                            else "#71827E"
                                                        ),
                                                        border_radius=10,
                                                        content=ft.Text(
                                                            str(unread_count),
                                                            color=ft.Colors.WHITE,
                                                            size=11,
                                                            weight=ft.FontWeight.BOLD,
                                                        ),
                                                    ),
                                                ],
                                            ),
                                            on_click=lambda _: self.navigate("notifications"),
                                        ),
                                        ft.TextButton(
                                            content=ft.Row(
                                                spacing=5,
                                                controls=[
                                                    ft.Icon(
                                                        ft.Icons.BADGE,
                                                        color="#103F37",
                                                    ),
                                                    ft.Text(
                                                        "Owners",
                                                        color="#103F37",
                                                        weight=ft.FontWeight.BOLD,
                                                    ),
                                                ],
                                            ),
                                            on_click=lambda _: self.navigate("owners"),
                                        ),
                                    ],
                                ),
                            ],
                        ),
                    ),
                    self.content,
                ],
            )
            self.page.add(
                ft.Row(
                    expand=True,
                    spacing=0,
                    controls=[sidebar_control, top_area],
                )
            )
            self.navigate("dashboard")
        except Exception as error:
            self.show_startup_error(error)

    def show_startup_error(self, error):
        error_path = APP_DIR / "crm_startup_error.log"
        try:
            error_path.write_text(
                f"{datetime.now().isoformat(timespec='seconds')}\n{type(error).__name__}: {error}",
                encoding="utf-8",
            )
        except Exception:
            pass

        self.page.clean()
        self.page.add(
            ft.Container(
                expand=True,
                alignment=ft.Alignment(0, 0),
                padding=30,
                content=ft.Container(
                    width=720,
                    padding=28,
                    bgcolor=ft.Colors.WHITE,
                    border_radius=20,
                    content=ft.Column(
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Icon(
                                ft.Icons.ERROR_OUTLINE,
                                size=58,
                                color=ft.Colors.RED_700,
                            ),
                            ft.Text(
                                "The CRM could not load after sign-in.",
                                size=25,
                                weight=ft.FontWeight.BOLD,
                                color=ft.Colors.RED_700,
                            ),
                            ft.Text(
                                str(error),
                                selectable=True,
                                text_align=ft.TextAlign.CENTER,
                            ),
                            ft.Text(
                                f"Error details were saved in:\n{error_path}",
                                selectable=True,
                                size=12,
                                color="#667B80",
                                text_align=ft.TextAlign.CENTER,
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.CENTER,
                                controls=[
                                    ft.FilledButton(
                                        "Retry",
                                        icon=ft.Icons.REFRESH,
                                        on_click=lambda _: self.show_main(),
                                    ),
                                    ft.OutlinedButton(
                                        "Return to Login",
                                        icon=ft.Icons.LOGIN,
                                        on_click=lambda _: self.show_login(),
                                    ),
                                ],
                            ),
                        ],
                    ),
                ),
            )
        )
        self.page.update()

    def current_owner_name(self):
        if not self.user:
            return ""
        username = (self.user.get("username") or "").strip()
        full_name = (self.user.get("full_name") or "").strip()
        owners = self.db.owners(active_only=True)
        for owner in owners:
            if owner["full_name"].lower() == full_name.lower():
                return owner["full_name"]
        if username == "admin":
            for owner in owners:
                if owner["full_name"] == "Saud Al Shammari":
                    return owner["full_name"]
        return full_name

    def notification_badge(self):
        owner_name = self.current_owner_name()
        count = self.db.unread_notification_count(owner_name)
        return ft.TextButton(
            content=ft.Row(
                spacing=5,
                controls=[
                    ft.Icon(ft.Icons.NOTIFICATIONS, color="#103F37"),
                    ft.Container(
                        padding=5,
                        bgcolor="#C0392B" if count else "#71827E",
                        border_radius=10,
                        content=ft.Text(
                            str(count),
                            color=ft.Colors.WHITE,
                            size=11,
                            weight=ft.FontWeight.BOLD,
                        ),
                    ),
                ],
            ),
            tooltip="Notifications",
            on_click=lambda _: self.navigate("notifications"),
        )

    def sidebar(self):
        def nav(label, icon, key):
            return ft.TextButton(
                width=240,
                height=44,
                content=ft.Row([
                    ft.Icon(icon, color=ft.Colors.WHITE, size=20),
                    ft.Text(label, color=ft.Colors.WHITE, size=14),
                ]),
                on_click=lambda _: self.navigate(key),
            )

        return ft.Container(
            width=270,
            bgcolor="#103F37",
            padding=18,
            content=ft.Column([
                ft.Container(
                    height=95,
                    bgcolor=ft.Colors.WHITE,
                    border_radius=12,
                    padding=8,
                    content=ft.Image(
                        src_base64=LOGO_BASE64,
                        fit=ft.ImageFit.CONTAIN,
                    ),
                ),
                ft.Text(
                    "Customer Relationship Management",
                    color="#BCD5CE",
                    size=13,
                    text_align=ft.TextAlign.CENTER,
                ),
                ft.Text(
                    self.user["full_name"],
                    color="#90B4AA",
                    size=12,
                ),
                ft.Divider(color="#2B6659"),
                nav("Dashboard", ft.Icons.DASHBOARD, "dashboard"),
                nav("Accounts", ft.Icons.BUSINESS, "accounts"),
                nav("Customer Contacts", ft.Icons.CONTACTS, "contacts"),
                nav("Leads", ft.Icons.PERSON_SEARCH, "leads"),
                nav("Opportunities", ft.Icons.TRENDING_UP, "opportunities"),
                nav("Quotations", ft.Icons.REQUEST_QUOTE, "quotations"),
                nav("Pipeline Analytics", ft.Icons.BAR_CHART, "pipeline"),
                nav("Reports", ft.Icons.ASSESSMENT, "reports"),
                ft.Container(expand=True),
                ft.TextButton(
                    "Backup Database",
                    icon=ft.Icons.BACKUP,
                    style=ft.ButtonStyle(color=ft.Colors.WHITE),
                    on_click=lambda _: self.backup(),
                ),
                ft.TextButton(
                    "Sign Out",
                    icon=ft.Icons.LOGOUT,
                    style=ft.ButtonStyle(color=ft.Colors.WHITE),
                    on_click=lambda _: self.show_login(),
                ),
            ]),
        )

    def navigate(self, view):
        if view in {
            "meetings",
            "activities",
            "migration",
            "legacy_pipeline",
        }:
            view = "dashboard"
        self.current_view = view
        try:
            methods = {
                "dashboard": self.dashboard_view, "accounts": self.accounts_view,
                "contacts": self.contacts_view, "leads": self.leads_view,
                "opportunities": self.opportunities_view, "legacy_pipeline": self.legacy_pipeline_view,
                "quotations": self.quotations_view,
                "pocs": self.pocs_view, "inbox": self.inbox_view, "notifications": self.notifications_view, "owners": self.owners_view, "meetings": self.meetings_view,
                "activities": self.activities_view, "pipeline": self.pipeline_view,
                "reports": self.reports_view, "migration": self.migration_view,
            }
            control = methods[view]()
            if view not in ("reports", "migration"):
                if isinstance(control, ft.Column):
                    insert_at = 1 if len(control.controls) >= 1 else 0
                    control.controls.insert(insert_at, self.section_export_toolbar(view))
        except Exception as e:
            control = ft.Container(
                bgcolor=ft.Colors.WHITE, padding=24, border_radius=18,
                content=ft.Column([
                    ft.Icon(ft.Icons.ERROR_OUTLINE, size=48, color=ft.Colors.RED_700),
                    ft.Text("This page could not be loaded.", size=24, weight=ft.FontWeight.BOLD),
                    ft.Text(str(e), selectable=True),
                    ft.FilledButton("Retry", on_click=lambda _: self.navigate(view))
                ])
            )
        self.content.content = ft.Column(expand=True, scroll=ft.ScrollMode.AUTO, controls=[control])
        self.page.update()

    def title(self, text, subtitle=""):
        return ft.Column([
            ft.Text(text, size=29, weight=ft.FontWeight.BOLD, color="#103F37"),
            ft.Text(subtitle, color="#657A80"),
        ])

    def card(self, title, value, icon, color, subtitle=""):
        return ft.Container(
            expand=True, bgcolor=ft.Colors.WHITE, padding=19, border_radius=18,
            shadow=ft.BoxShadow(blur_radius=14, color=ft.Colors.with_opacity(0.065, ft.Colors.BLACK)),
            content=ft.Column([
                ft.Row([ft.Icon(icon, color=color), ft.Text(title, color="#6C7F84")]),
                ft.Text(value, size=24, weight=ft.FontWeight.BOLD, color=color),
                ft.Text(subtitle, size=11, color="#819196") if subtitle else ft.Container(),
            ])
        )

    def filtered_opportunities(self):
        opportunities = self.db.opportunities()
        legacy_rows = {
            row["opportunity_reference"]: row
            for row in self.db.legacy_opportunities()
        }

        filters = self.dashboard_filters
        result = []

        for opportunity in opportunities:
            legacy = legacy_rows.get(opportunity["opportunity_reference"], {})
            expected_close = opportunity["expected_close_date"] or ""
            year = expected_close[:4] if len(expected_close) >= 4 else ""
            month = expected_close[5:7] if len(expected_close) >= 7 else ""
            business_unit = legacy.get("business_unit") or ""

            if filters["year"] != "All" and year != filters["year"]:
                continue
            if filters["month"] != "All" and month != filters["month"]:
                continue
            if filters["account"] != "All" and (opportunity["account_name"] or "") != filters["account"]:
                continue
            if filters["sales_owner"] != "All" and (opportunity["sales_owner"] or "Unassigned") != filters["sales_owner"]:
                continue
            if filters["stage"] != "All" and (opportunity["stage"] or "") != filters["stage"]:
                continue
            if filters["business_unit"] != "All" and business_unit != filters["business_unit"]:
                continue
            if filters["status"] != "All" and (opportunity["status"] or "") != filters["status"]:
                continue

            item = dict(opportunity)
            item["business_unit"] = business_unit
            item["source_stage"] = legacy.get("source_stage") or ""
            item["include_in_forecast"] = legacy.get("include_in_forecast", 0)
            item["forecast_gm_percent"] = legacy.get("forecast_gm_percent", 0)
            item["forecast_gm_value"] = legacy.get("forecast_gm_value", 0)
            result.append(item)

        return result

    def dashboard_filter_summary(self):
        labels = {
            "year": "Year",
            "month": "Month",
            "account": "Account",
            "sales_owner": "Sales Owner",
            "stage": "Stage",
            "business_unit": "Business Unit",
            "status": "Status",
        }
        active = [
            f"{labels[key]}: {value}"
            for key, value in self.dashboard_filters.items()
            if value != "All"
        ]
        return " | ".join(active) if active else "All records"

    def reset_dashboard_filters(self):
        self.dashboard_filters = {
            "year": "All",
            "month": "All",
            "account": "All",
            "sales_owner": "All",
            "stage": "All",
            "business_unit": "All",
            "status": "Open",
        }
        self.navigate("dashboard")

    def dashboard_filter_bar(self):
        opportunities = self.db.opportunities()
        legacy = self.db.legacy_opportunities()

        years = sorted({
            (row["expected_close_date"] or "")[:4]
            for row in opportunities
            if len(row["expected_close_date"] or "") >= 4
        })
        accounts = sorted({
            row["account_name"] for row in opportunities if row["account_name"]
        })
        owners = sorted({
            row["sales_owner"] or "Unassigned" for row in opportunities
        })
        stages = sorted({
            row["stage"] for row in opportunities if row["stage"]
        })
        business_units = sorted({
            row["business_unit"] for row in legacy if row["business_unit"]
        })
        statuses = sorted({
            row["status"] for row in opportunities if row["status"]
        })

        month_names = {
            "01": "January", "02": "February", "03": "March",
            "04": "April", "05": "May", "06": "June",
            "07": "July", "08": "August", "09": "September",
            "10": "October", "11": "November", "12": "December",
        }

        def dropdown(label, key, values, width):
            options = [ft.dropdown.Option("All")]
            for value in values:
                display = month_names.get(value, value) if key == "month" else value
                options.append(ft.dropdown.Option(key=value, text=display))
            control = ft.Dropdown(
                label=label,
                value=self.dashboard_filters[key],
                width=width,
                options=options,
                dense=True,
            )
            def changed(event):
                self.dashboard_filters[key] = event.control.value or "All"
                self.navigate("dashboard")
            control.on_change = changed
            return control

        return ft.Container(
            bgcolor=ft.Colors.WHITE,
            padding=15,
            border_radius=16,
            content=ft.Column(
                controls=[
                    ft.Row(
                        wrap=True,
                        spacing=10,
                        run_spacing=10,
                        controls=[
                            dropdown("Year", "year", years, 130),
                            dropdown("Month", "month", list(month_names), 150),
                            dropdown("Account", "account", accounts, 250),
                            dropdown("Sales Owner", "sales_owner", owners, 210),
                            dropdown("Stage", "stage", stages, 190),
                            dropdown("Business Unit", "business_unit", business_units, 210),
                            dropdown("Status", "status", statuses, 150),
                            ft.OutlinedButton(
                                "Clear Filters",
                                icon=ft.Icons.FILTER_ALT_OFF,
                                on_click=lambda _: self.reset_dashboard_filters(),
                            ),
                        ],
                    ),
                    ft.Text(
                        f"Applied filters: {self.dashboard_filter_summary()}",
                        size=11,
                        color="#667B80",
                        selectable=True,
                    ),
                ]
            ),
        )

    def calculate_filtered_dashboard(self):
        opportunities = self.filtered_opportunities()
        open_items = [row for row in opportunities if row["status"] == "Open"]
        won_items = [
            row for row in opportunities
            if row["status"] == "Won" or row["stage"] == "Awarded"
        ]
        lost_items = [
            row for row in opportunities
            if row["status"] == "Lost" or row["stage"] == "Lost"
        ]

        pipeline = sum(float(row["estimated_value"] or 0) for row in open_items)
        weighted = sum(
            float(row["estimated_value"] or 0)
            * float(row["probability"] or 0) / 100
            for row in open_items
        )
        win_rate = (
            len(won_items) / (len(won_items) + len(lost_items)) * 100
            if won_items or lost_items else 0
        )

        account_count = len({
            row["account_name"] for row in opportunities if row["account_name"]
        })

        return {
            "accounts": account_count,
            "leads": 0,
            "open_opportunities": len(open_items),
            "pipeline": pipeline,
            "weighted_pipeline": weighted,
            "quotation_value": 0,
            "win_rate": win_rate,
            "lead_conversion": 0,
            "active_pocs": 0,
            "assigned_quotations": len([
                row for row in self.db.quotations()
                if row["status"] in ("ASSIGNED", "IN PROGRESS", "INTERNAL REVIEW")
            ]),
            "overdue_quotations": len([
                row for row in self.db.quotations()
                if row["status"] not in ("WON", "LOST", "CANCELLED", "SUBMITTED")
                and row["valid_until"]
                and row["valid_until"] < date.today().isoformat()
            ]),
            "open_inbox_items": len(self.db.inbox_for_owner(self.current_owner_name())),
            "unread_notifications": self.db.unread_notification_count(self.current_owner_name()),
            "overdue_activities": 0,
            "due_7_days": 0,
            "overdue_rows": [],
            "all_filtered_opportunities": opportunities,
        }

    def section_export_toolbar(self, section):
        section_labels = {
            "dashboard": "Executive Dashboard",
            "accounts": "Accounts",
            "contacts": "Contacts",
            "leads": "Leads",
            "opportunities": "Opportunities",
            "legacy_pipeline": "Migrated Pipeline",
            "quotations": "Quotations",
            "pocs": "PoCs",
            "meetings": "Meetings",
            "activities": "Activities",
            "owners": "Owners",
            "pipeline": "Pipeline Analytics",
        }
        label = section_labels.get(section, section.replace("_", " ").title())
        return ft.Container(
            bgcolor="#EAF2EF",
            padding=10,
            border_radius=12,
            content=ft.Row(
                alignment=ft.MainAxisAlignment.END,
                controls=[
                    ft.Text(
                        f"Export {label}:",
                        weight=ft.FontWeight.BOLD,
                        color="#103F37",
                    ),
                    ft.OutlinedButton(
                        "PDF",
                        icon=ft.Icons.PICTURE_AS_PDF,
                        on_click=lambda _: self.export_section_pdf(section),
                    ),
                    ft.OutlinedButton(
                        "Excel",
                        icon=ft.Icons.TABLE_VIEW,
                        on_click=lambda _: self.export_section_excel(section),
                    ),
                ],
            ),
        )

    def section_dataset(self, section):
        if section == "dashboard":
            rows = self.filtered_opportunities()
            headers = [
                "Reference", "Account", "Opportunity", "Stage", "Probability",
                "Value", "Weighted Value", "Expected Close", "Sales Owner",
                "Business Unit", "Status",
            ]
            data = [[
                row["opportunity_reference"], row["account_name"], row["opportunity_name"],
                row["stage"], f'{float(row["probability"] or 0):.0f}%',
                float(row["estimated_value"] or 0),
                float(row["estimated_value"] or 0) * float(row["probability"] or 0) / 100,
                row["expected_close_date"], row["sales_owner"], row["business_unit"],
                row["status"],
            ] for row in rows]
            return "Executive Dashboard", headers, data, {5, 6}

        if section == "accounts":
            rows = self.db.accounts()
            return "Accounts", [
                "Account", "Type", "Industry", "City", "Country", "Owner", "Status", "Phone"
            ], [[
                row["account_name"], row["account_type"], row["industry"], row["city"],
                row["country"], row["owner"], row["status"], row["main_phone"],
            ] for row in rows], set()

        if section == "contacts":
            rows = self.db.contacts()
            return "Contacts", [
                "Account", "Name", "Job Title", "Department", "Email", "Mobile",
                "Influence", "Relationship", "Owner"
            ], [[
                row["account_name"], f'{row["first_name"]} {row["last_name"] or ""}'.strip(),
                row["job_title"], row["department"], row["email"], row["mobile"],
                row["influence_level"], row["relationship_status"], row["owner"],
            ] for row in rows], set()

        if section == "leads":
            rows = self.db.leads()
            return "Leads", [
                "Reference", "Company", "Contact", "Source", "Interest", "Status",
                "Score", "Estimated Value", "Owner", "Next Action"
            ], [[
                row["lead_reference"], row["company_name"], row["contact_name"],
                row["source"], row["interest_area"], row["lead_status"], row["lead_score"],
                float(row["estimated_value"] or 0), row["owner"], row["next_action_date"],
            ] for row in rows], {7}

        if section == "opportunities":
            rows = self.db.opportunities()
            return "Opportunities", [
                "Reference", "Account", "Opportunity", "Type", "Stage", "Probability",
                "Value", "Weighted Value", "Close Date", "Sales Owner", "Technical Owner",
                "Next Step", "Status"
            ], [[
                row["opportunity_reference"], row["account_name"], row["opportunity_name"],
                row["project_type"], row["stage"], f'{float(row["probability"] or 0):.0f}%',
                float(row["estimated_value"] or 0),
                float(row["estimated_value"] or 0) * float(row["probability"] or 0) / 100,
                row["expected_close_date"], row["sales_owner"], row["technical_owner"],
                row["next_step"], row["status"],
            ] for row in rows], {6, 7}

        if section == "legacy_pipeline":
            rows = self.db.legacy_opportunities()
            return "Migrated Pipeline", [
                "Reference", "Account", "End User", "Deal Name", "Business Unit",
                "Source Type", "Industry", "Source Stage", "CRM Stage", "Probability",
                "Gross", "GM %", "GM Value", "Expected PO", "Delivery", "Sales Owner",
                "Assigned To", "Forecast", "Must Win", "Suspended", "Quality Flags"
            ], [[
                row["opportunity_reference"], row["account_name"], row["end_user"],
                row["opportunity_name"], row["business_unit"], row["source_project_type"],
                row["industry"], row["source_stage"], row["stage"],
                f'{float(row["probability"] or 0):.0f}%', float(row["estimated_value"] or 0),
                float(row["forecast_gm_percent"] or 0), float(row["forecast_gm_value"] or 0),
                f'{row["expected_po_month"]} {row["expected_po_year"]}', row["delivery_date"],
                row["sales_owner"], row["assigned_to"],
                "Yes" if row["include_in_forecast"] else "No",
                "Yes" if row["must_win"] else "No",
                "Yes" if row["suspended"] else "No", row["quality_flags"],
            ] for row in rows], {10, 12}

        if section == "quotations":
            rows = self.db.quotations()
            return "Quotations", [
                "Quotation", "Opportunity", "Account", "Date", "Valid Until",
                "Selling Price", "Cost Price", "Discount", "Net Before VAT",
                 "VAT Amount", "Total Including VAT", "Gross Margin %", "Gross Margin Value",
                "Revision", "Status", "Owner"
            ], [[
                row["quotation_number"], row["opportunity_reference"], row["account_name"],
                row["quotation_date"], row["valid_until"], float(row["base_value"] or 0),
                float(row["discount"] or 0),
                float(row["base_value"] or 0) - float(row["discount"] or 0),
                (float(row["base_value"] or 0) - float(row["discount"] or 0))
                * (1 + float(row["vat_rate"] or 0) / 100),
                row["revision"], row["status"], row["owner"],
            ] for row in rows], {5, 6, 7, 8}

        if section == "pocs":
            rows = self.db.pocs()
            return "Proof of Concepts", [
                "Reference", "Account", "Opportunity", "PoC Title", "Solution",
                "Start", "End", "Status", "Estimated Cost", "Commercial Value",
                "Owner", "Next Step"
            ], [[
                row["poc_reference"], row["account_name"], row["opportunity_reference"],
                row["poc_title"], row["solution"], row["start_date"], row["planned_end_date"],
                row["status"], float(row["estimated_cost"] or 0),
                float(row["commercial_value"] or 0), row["owner"], row["next_step"],
            ] for row in rows], {8, 9}

        if section == "meetings":
            rows = self.db.meetings()
            return "Meetings", [
                "Reference", "Date", "Account", "Opportunity", "Type", "Subject",
                "Location", "Owner", "Outcome", "Next Action", "Next Action Date"
            ], [[
                row["meeting_reference"], row["meeting_date"], row["account_name"],
                row["opportunity_reference"], row["meeting_type"], row["subject"],
                row["location"], row["owner"], row["outcome"], row["next_action"],
                row["next_action_date"],
            ] for row in rows], set()

        if section == "owners":
            rows = self.db.owners()
            return "Owners", [
                "Owner ID", "Owner Name", "Role", "Department", "Email",
                "Mobile", "Eligible for Quotation Assignment", "Status"
            ], [[
                row["owner_code"], row["full_name"], row["role"], row["department"],
                row["email"], row["mobile"],
                "Yes" if row["can_receive_quotation"] else "No", row["status"],
            ] for row in rows], set()

        if section == "activities":
            rows = self.db.activities()
            return "Activities", [
                "Reference", "Due Date", "Account", "Opportunity", "Type", "Subject",
                "Priority", "Status", "Owner", "Completed Date"
            ], [[
                row["activity_reference"], row["due_date"], row["account_name"],
                row["opportunity_reference"], row["activity_type"], row["subject"],
                row["priority"], row["status"], row["owner"], row["completed_date"],
            ] for row in rows], set()

        if section == "pipeline":
            rows = self.filtered_opportunities()
            return "Pipeline Analytics", [
                "Reference", "Account", "Opportunity", "Stage", "Probability",
                "Gross Value", "Weighted Value", "Expected Close", "Sales Owner",
                "Business Unit", "Status"
            ], [[
                row["opportunity_reference"], row["account_name"], row["opportunity_name"],
                row["stage"], f'{float(row["probability"] or 0):.0f}%',
                float(row["estimated_value"] or 0),
                float(row["estimated_value"] or 0) * float(row["probability"] or 0) / 100,
                row["expected_close_date"], row["sales_owner"], row["business_unit"],
                row["status"],
            ] for row in rows], {5, 6}

        return section.title(), [], [], set()

    def export_section_excel(self, section):
        try:
            title, headers, rows, money_columns = self.section_dataset(section)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = section.replace(" ", "_")
            path = EXPORT_DIR / f"Saudi_Sensing_CRM_{safe_name}_{stamp}.xlsx"

            workbook = xlsxwriter.Workbook(str(path))
            sheet = workbook.add_worksheet(title[:31])
            title_format = workbook.add_format({
                "bold": True, "font_size": 18, "font_name": "Calibri",
                "font_color": "#103F37",
            })
            subtitle_format = workbook.add_format({
                "italic": True, "font_name": "Calibri", "font_color": "#667B80",
            })
            header_format = workbook.add_format({
                "bold": True, "bg_color": "#DDEBE6", "border": 1,
                "font_name": "Calibri", "text_wrap": True,
            })
            text_format = workbook.add_format({
                "border": 1, "font_name": "Calibri", "text_wrap": True,
                "valign": "top",
            })
            money_format = workbook.add_format({
                "border": 1, "font_name": "Calibri",
                "num_format": '#,##0.00 "SAR"', "valign": "top",
            })

            sheet.write("A1", f"Saudi Sensing CRM - {title}", title_format)
            if section in ("dashboard", "pipeline"):
                sheet.write("A2", f"Applied filters: {self.dashboard_filter_summary()}", subtitle_format)

            start_row = 3
            for column, header in enumerate(headers):
                sheet.write(start_row, column, header, header_format)

            for row_index, row in enumerate(rows, start=start_row + 1):
                for column, value in enumerate(row):
                    if column in money_columns:
                        sheet.write_number(row_index, column, float(value or 0), money_format)
                    else:
                        sheet.write(row_index, column, "" if value is None else value, text_format)

            if headers:
                sheet.autofilter(start_row, 0, max(start_row + 1, start_row + len(rows)), len(headers) - 1)
                sheet.freeze_panes(start_row + 1, 0)
                for column, header in enumerate(headers):
                    width = 18
                    if any(word in header.lower() for word in ["opportunity", "subject", "notes", "quality", "next step"]):
                        width = 34
                    elif any(word in header.lower() for word in ["account", "owner", "business unit"]):
                        width = 24
                    sheet.set_column(column, column, width)

            workbook.close()
            self.notify(f"{title} Excel exported: {path.name}")
            os.startfile(path)
        except Exception as error:
            self.notify(f"Section Excel export failed: {error}", True)

    def export_section_pdf(self, section):
        try:
            title, headers, rows, money_columns = self.section_dataset(section)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = section.replace(" ", "_")
            path = EXPORT_DIR / f"Saudi_Sensing_CRM_{safe_name}_{stamp}.pdf"

            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                "SectionTitle",
                parent=styles["Title"],
                fontName="Helvetica-Bold",
                fontSize=18,
                textColor=colors.HexColor("#103F37"),
            )
            subtitle_style = ParagraphStyle(
                "SectionSubtitle",
                parent=styles["BodyText"],
                fontSize=8,
                textColor=colors.HexColor("#667B80"),
            )
            cell_style = ParagraphStyle(
                "SectionCell",
                parent=styles["BodyText"],
                fontSize=6.4,
                leading=8,
                wordWrap="CJK",
            )
            header_style = ParagraphStyle(
                "SectionHeader",
                parent=cell_style,
                fontName="Helvetica-Bold",
                textColor=colors.HexColor("#103F37"),
                alignment=1,
            )

            def paragraph(value, style=cell_style):
                value = "" if value is None else str(value)
                value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                return Paragraph(value, style)

            document = SimpleDocTemplate(
                str(path),
                pagesize=landscape(A4),
                leftMargin=8 * mm,
                rightMargin=8 * mm,
                topMargin=9 * mm,
                bottomMargin=9 * mm,
            )

            story = [
                Paragraph(f"Saudi Sensing CRM - {title}", title_style),
            ]
            if section in ("dashboard", "pipeline"):
                story.extend([
                    Paragraph(f"Applied filters: {self.dashboard_filter_summary()}", subtitle_style),
                    Spacer(1, 2 * mm),
                ])

            if not headers:
                story.append(Paragraph("No data available.", subtitle_style))
            else:
                table_data = [[paragraph(header, header_style) for header in headers]]
                for row in rows:
                    formatted_row = []
                    for index, value in enumerate(row):
                        if index in money_columns:
                            formatted_row.append(paragraph(money(float(value or 0))))
                        else:
                            formatted_row.append(paragraph(value))
                    table_data.append(formatted_row)

                available_width = landscape(A4)[0] - 16 * mm
                base_width = available_width / len(headers)
                column_widths = []
                for header in headers:
                    if any(word in header.lower() for word in ["opportunity", "subject", "quality", "notes", "next step"]):
                        column_widths.append(base_width * 1.65)
                    elif any(word in header.lower() for word in ["probability", "status", "date", "reference"]):
                        column_widths.append(base_width * 0.75)
                    else:
                        column_widths.append(base_width)

                scale = available_width / sum(column_widths)
                column_widths = [width * scale for width in column_widths]

                table = Table(
                    table_data,
                    repeatRows=1,
                    colWidths=column_widths,
                )
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBE6")),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#AFC3BD")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                        colors.white, colors.HexColor("#F7FAF9")
                    ]),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]))
                story.append(table)

            document.build(story)
            self.notify(f"{title} PDF exported: {path.name}")
            if WEB_MODE:
    download_path = APP_DIR / "assets" / path.name
    shutil.copy2(path, download_path)
    self.page.launch_url(f"/{path.name}")
else:
    os.startfile(path)
        except Exception as error:
            self.notify(f"Section PDF export failed: {error}", True)

    def dashboard_view(self):
        d = self.calculate_filtered_dashboard()
        overdue_rows = []

        return ft.Column([
            self.title("Dashboard", "Real-time commercial intelligence and sales execution priorities"),
            self.dashboard_filter_bar(),
            ft.Row([
                self.card("Total Accounts", str(d["accounts"]), ft.Icons.BUSINESS, "#0E5A47"),
                self.card("Leads", str(d["leads"]), ft.Icons.PERSON_SEARCH, "#2667A8"),
                self.card("Open Opportunities", str(d["open_opportunities"]), ft.Icons.TRENDING_UP, "#7A4FA3"),
                self.card("Assigned Quotations", str(d.get("assigned_quotations", 0)), ft.Icons.REQUEST_QUOTE, "#D27A00"),
            ]),
            ft.Row([
                self.card("Total Pipeline", money(d["pipeline"]), ft.Icons.ACCOUNT_BALANCE_WALLET, "#0E5A47"),
                self.card("Weighted Pipeline", money(d["weighted_pipeline"]), ft.Icons.SHOW_CHART, "#2E8B57"),
                self.card("Open Quotation Value", money(d["quotation_value"]), ft.Icons.REQUEST_QUOTE, "#1667A8"),
                self.card("Win Rate", f'{float(d.get("win_rate", 0) or 0):.1f}%', ft.Icons.STAR, "#B47A00"),
            ]),
            ft.Row([
                self.card("Lead Conversion", f'{float(d.get("lead_conversion", 0) or 0):.1f}%', ft.Icons.SWAP_HORIZ, "#247A73"),
                self.card("Overdue Activities", str(d["overdue_activities"]), ft.Icons.ERROR_OUTLINE,
                          "#C0392B" if d["overdue_activities"] else "#2E8B57"),
                self.card("Actions Due in 7 Days", str(d["due_7_days"]), ft.Icons.EVENT, "#D35400"),
                self.card("Overdue Quotations", str(d.get("overdue_quotations", 0)), ft.Icons.WARNING_AMBER, "#C0392B"),
                self.card("My Inbox", str(d.get("open_inbox_items", 0)), ft.Icons.INBOX, "#1667A8"),
                self.card("Notifications", str(d.get("unread_notifications", 0)), ft.Icons.NOTIFICATIONS, "#7A4FA3"),
            ]),
            ft.Container(
                bgcolor=ft.Colors.WHITE, padding=18, border_radius=18,
                content=ft.Column([
                    ft.Text("Overdue Follow-ups", size=19, weight=ft.FontWeight.BOLD),
                    ft.DataTable(
                        columns=[ft.DataColumn(ft.Text(x)) for x in ["Activity", "Account", "Subject", "Due Date", "Owner", "Priority"]],
                        rows=overdue_rows,
                    ) if overdue_rows else ft.Text("No overdue activities. Sales follow-ups are under control.", color="#2E8B57"),
                ])
            )
        ])

    def owner_options(self, quotation_only=False):
        return [
            ft.dropdown.Option(
                key=row["full_name"],
                text=f'{row["owner_code"]} - {row["full_name"]} | {row["role"]}',
            )
            for row in self.db.owners(
                active_only=True,
                quotation_only=False,
            )
        ]

    def account_options(self):
        return [ft.dropdown.Option(key=str(a["id"]), text=a["account_name"]) for a in self.db.accounts()]

    def opportunity_options(self):
        return [ft.dropdown.Option(key=str(o["id"]), text=f'{o["opportunity_reference"]} - {o["opportunity_name"]}') for o in self.db.opportunities()]

    def accounts_view(self):
        all_records = self.db.accounts()
        table_holder = ft.Column()

        def render(records):
            table_holder.controls = [
                self.crud_table_container(
                    ['Account ID', 'Account', 'Type', 'Industry', 'City', 'Country', 'Owner', 'Status', 'Phone'],
                    records,
                    lambda r: [r["account_reference"],r["account_name"],r["account_type"],r["industry"],r["city"],r["country"],r["owner"],r["status"],r["main_phone"]],
                    "accounts",
                    "No accounts created yet.",
                )
            ]
            try:
                table_holder.update()
            except Exception:
                pass

        render(all_records)

        def search_changed(event):
            query = (event.control.value or "").strip().lower()
            filtered = [
                row for row in all_records
                if not query or query in " ".join([
                    str(row[field] or "") for field in ['account_reference', 'account_name', 'account_type', 'industry', 'city', 'country', 'owner', 'status', 'main_phone']
                ]).lower()
            ]
            render(filtered)

        return ft.Column([
            ft.Row([
                self.title("Accounts", "Customers, partners, EPCs, OEMs, consultants and strategic relationships"),
                ft.FilledButton(
                    "New Account",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.account_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.search_with_clear("Search Accounts", search_changed),
            table_holder,
        ])

    def contacts_view(self):
        all_records = self.db.contacts()
        table_holder = ft.Column()

        def render(records):
            table_holder.controls = [
                self.crud_table_container(
                    ['Contact ID', 'Account', 'Name', 'Title', 'Department', 'Email', 'Mobile', 'Influence', 'Relationship', 'Owner'],
                    records,
                    lambda r: [r["contact_reference"],r["account_name"],f'{r["first_name"]} {r["last_name"] or ""}'.strip(),r["job_title"],r["department"],r["email"],r["mobile"],r["influence_level"],r["relationship_status"],r["owner"]],
                    "contacts",
                    "No contacts created yet.",
                )
            ]
            try:
                table_holder.update()
            except Exception:
                pass

        render(all_records)

        def search_changed(event):
            query = (event.control.value or "").strip().lower()
            filtered = [
                row for row in all_records
                if not query or query in " ".join([
                    str(row[field] or "") for field in ['contact_reference', 'account_name', 'first_name', 'last_name', 'job_title', 'department', 'email', 'mobile', 'owner']
                ]).lower()
            ]
            render(filtered)

        return ft.Column([
            ft.Row([
                self.title("Customer Contacts", "Decision makers, influencers, technical users and commercial stakeholders"),
                ft.FilledButton(
                    "New Contact",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.contact_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.search_with_clear("Search Customer Contacts", search_changed),
            table_holder,
        ])

    def lead_workspace(self, lead_id):
        try:
            record = self.db.lead_full_details(lead_id)
            if not record:
                self.notify("The lead no longer exists.", True)
                return

            def value_text(value):
                if value is None or str(value).strip() == "":
                    return "Not available"
                return str(value)

            def card(label, value, width=250):
                return ft.Container(
                    width=width,
                    padding=13,
                    bgcolor="#F5F8F7",
                    border_radius=11,
                    content=ft.Column(
                        spacing=4,
                        controls=[
                            ft.Text(label, size=11, color="#667B80"),
                            ft.Text(
                                value_text(value),
                                size=14,
                                weight=ft.FontWeight.BOLD,
                                color="#103F37",
                                selectable=True,
                            ),
                        ],
                    ),
                )

            workspace = ft.Column(
                expand=True,
                scroll=ft.ScrollMode.AUTO,
                controls=[
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                controls=[
                                    ft.Text(
                                        f"Lead Workspace | {record['lead_reference']}",
                                        size=28,
                                        weight=ft.FontWeight.BOLD,
                                        color="#103F37",
                                    ),
                                    ft.Text(
                                        "Lead details, qualification and conversion to opportunity",
                                        color="#667B80",
                                    ),
                                ],
                            ),
                            ft.Row(
                                controls=[
                                    ft.OutlinedButton(
                                        "Back to Leads",
                                        icon=ft.Icons.ARROW_BACK,
                                        on_click=lambda _: self.navigate("leads"),
                                    ),
                                    ft.FilledButton(
                                        "Convert to Opportunity",
                                        icon=ft.Icons.TRENDING_UP,
                                        on_click=lambda _: self.convert_lead_to_opportunity(
                                            lead_id
                                        ),
                                        style=ft.ButtonStyle(bgcolor="#0E5A47"),
                                    ),
                                    ft.IconButton(
                                        icon=ft.Icons.EDIT,
                                        tooltip="Edit Lead",
                                        on_click=lambda _: self.edit_record_dialog(
                                            "leads",
                                            lead_id,
                                        ),
                                    ),
                                ],
                            ),
                        ],
                    ),
                    ft.Container(
                        padding=16,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column(
                            spacing=12,
                            controls=[
                                ft.Text(
                                    "Lead Information",
                                    size=19,
                                    weight=ft.FontWeight.BOLD,
                                    color="#103F37",
                                ),
                                ft.Row(
                                    wrap=True,
                                    controls=[
                                        card("Lead ID", record["lead_reference"]),
                                        card("Account ID", record["account_reference"]),
                                        card("Account", record["account_name"], 330),
                                        card("Company", record["company_name"], 330),
                                        card("Contact Name", record["contact_name"]),
                                        card("Job Title", record["job_title"]),
                                        card("Email", record["email"], 330),
                                        card("Mobile", record["mobile"]),
                                        card("Lead Source", record["source"]),
                                        card("Interest Area", record["interest_area"]),
                                        card("Lead Status", record["lead_status"]),
                                        card("Lead Score", record["lead_score"]),
                                        card("Estimated Value", money(record["estimated_value"] or 0)),
                                        card("Owner", record["owner"]),
                                        card("Next Action Date", record["next_action_date"]),
                                    ],
                                ),
                                ft.Divider(),
                                ft.Text(
                                    f"Notes: {value_text(record['notes'])}",
                                    selectable=True,
                                    color="#304B45",
                                ),
                            ],
                        ),
                    ),
                    ft.Container(
                        padding=16,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column(
                            spacing=12,
                            controls=[
                                ft.Text(
                                    "Linked Account",
                                    size=19,
                                    weight=ft.FontWeight.BOLD,
                                    color="#103F37",
                                ),
                                ft.Row(
                                    wrap=True,
                                    controls=[
                                        card("Account Type", record["account_type"]),
                                        card("Industry", record["account_industry"]),
                                        card("City", record["account_city"]),
                                        card("Country", record["account_country"]),
                                    ],
                                ),
                            ],
                        ),
                    ),
                ],
            )

            self.current_view = "lead_workspace"
            self.content.content = workspace
            self.page.update()

        except Exception as error:
            self.notify(f"Could not open lead workspace: {error}", True)

    def convert_lead_to_opportunity(self, lead_id):
        lead = self.db.lead_full_details(lead_id)
        if not lead:
            self.notify("The lead no longer exists.", True)
            return

        existing = self.db.one(
            """SELECT * FROM opportunities
               WHERE account_id=? AND opportunity_name=?
               ORDER BY id DESC LIMIT 1""",
            (
                lead["account_id"],
                lead["company_name"] or lead["interest_area"],
            ),
        )

        opportunity_reference = self.db.peek_reference("opportunities")
        opportunity_name = ft.TextField(
            label="Opportunity Name",
            width=620,
            value=(
                f"{lead['company_name']} - {lead['interest_area']}"
                if lead["company_name"]
                else lead["interest_area"]
            ),
        )
        project_type = ft.Dropdown(
            label="Project Type",
            width=300,
            value=lead["interest_area"] or INTEREST_AREAS[0],
            options=[
                ft.dropdown.Option(item)
                for item in INTEREST_AREAS
            ],
        )
        stage = ft.Dropdown(
            label="Stage",
            width=240,
            value="Qualification",
            options=[
                ft.dropdown.Option(item)
                for item in OPPORTUNITY_STAGES
            ],
        )
        probability = ft.TextField(
            label="Probability %",
            width=180,
            value="25",
            keyboard_type=ft.KeyboardType.NUMBER,
        )
        estimated_value = ft.TextField(
            label="Estimated Value",
            width=240,
            value=str(float(lead["estimated_value"] or 0)),
            keyboard_type=ft.KeyboardType.NUMBER,
        )
        close_control, expected_close = self.calendar_field(
            "Expected Close Date",
            (date.today() + timedelta(days=90)).isoformat(),
            width=280,
        )
        sales_owner = ft.TextField(
            label="Sales Owner",
            width=300,
            value=lead["owner"] or self.current_owner_name(),
            read_only=True,
        )
        next_step = ft.TextField(
            label="Next Step",
            width=620,
            value="Qualify the opportunity and confirm customer requirements.",
        )
        notes = ft.TextField(
            label="Notes",
            width=620,
            multiline=True,
            min_lines=3,
            value=(
                f"Converted from Lead {lead['lead_reference']}. "
                f"{lead['notes'] or ''}"
            ),
        )

        def save_conversion(_):
            try:
                if existing:
                    raise ValueError(
                        f"A similar opportunity already exists: "
                        f"{existing['opportunity_reference']}"
                    )
                if not opportunity_name.value.strip():
                    raise ValueError("Opportunity Name is required.")
                valid_date(expected_close.value)

                self.db.insert_opportunity({
                    "opportunity_reference": opportunity_reference,
                    "account_id": lead["account_id"],
                    "opportunity_name": opportunity_name.value.strip(),
                    "project_type": project_type.value,
                    "stage": stage.value,
                    "probability": safe_float(probability.value),
                    "estimated_value": safe_float(estimated_value.value),
                    "customer_budget": 0,
                    "currency": "SAR",
                    "expected_close_date": expected_close.value,
                    "sales_owner": sales_owner.value,
                    "technical_owner": "",
                    "next_step": next_step.value,
                    "next_action_date": lead["next_action_date"] or expected_close.value,
                    "competitors": "",
                    "lost_reason": "",
                    "status": "Open",
                    "notes": notes.value,
                })

                self.db.update_record(
                    "leads",
                    lead_id,
                    {"lead_status": "Converted"},
                )

                self.page.close(dialog)
                self.notify(
                    f"Lead {lead['lead_reference']} converted to "
                    f"{opportunity_reference}."
                )
                opportunity = self.db.one(
                    "SELECT id FROM opportunities WHERE opportunity_reference=?",
                    (opportunity_reference,),
                )
                if opportunity:
                    self.opportunity_workspace(opportunity["id"])
                else:
                    self.navigate("opportunities")
            except Exception as error:
                self.notify(f"Could not convert lead: {error}", True)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text(
                f"Convert Lead to Opportunity | {opportunity_reference}",
                weight=ft.FontWeight.BOLD,
            ),
            content=ft.Container(
                width=760,
                content=ft.Column(
                    height=600,
                    scroll=ft.ScrollMode.AUTO,
                    controls=[
                        ft.Container(
                            padding=12,
                            bgcolor="#EAF2EF",
                            border_radius=10,
                            content=ft.Text(
                                f"Account: {lead['account_name']} | "
                                f"Lead: {lead['lead_reference']}"
                            ),
                        ),
                        opportunity_name,
                        ft.Row([project_type, stage, probability], wrap=True),
                        ft.Row([estimated_value, close_control], wrap=True),
                        sales_owner,
                        next_step,
                        notes,
                    ],
                ),
            ),
            actions=[
                ft.TextButton(
                    "Cancel",
                    on_click=lambda _: self.page.close(dialog),
                ),
                ft.FilledButton(
                    "Create Opportunity",
                    icon=ft.Icons.TRENDING_UP,
                    on_click=save_conversion,
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ],
        )
        self.page.open(dialog)

    def leads_view(self):
        all_records = self.db.leads()

        columns = [
            ft.DataColumn(ft.Container(width=105, content=ft.Text("Lead ID", weight=ft.FontWeight.BOLD))),
            ft.DataColumn(ft.Text("Account", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Company", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Contact", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Source", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Interest", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Status", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Score", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Estimated Value", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Owner", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Next Action", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Actions", weight=ft.FontWeight.BOLD)),
        ]
        table = ft.DataTable(
            columns=columns,
            rows=[],
            column_spacing=16,
            heading_row_color="#E5F0EC",
        )

        def build_rows(records):
            output = []
            for record in records:
                converted = record["lead_status"] == "Converted"
                output.append(ft.DataRow(cells=[
                    ft.DataCell(self.id_badge(
                        record["lead_reference"],
                        width=105,
                        clickable=lambda _, lid=record["id"]: self.lead_workspace(lid),
                        tooltip="Open Lead Workspace",
                    )),
                    ft.DataCell(ft.Text(record["account_name"] or "")),
                    ft.DataCell(ft.Text(record["company_name"] or "")),
                    ft.DataCell(ft.Text(record["contact_name"] or "")),
                    ft.DataCell(ft.Text(record["source"] or "")),
                    ft.DataCell(ft.Text(record["interest_area"] or "")),
                    ft.DataCell(ft.Text(record["lead_status"] or "")),
                    ft.DataCell(ft.Text(f'{float(record["lead_score"] or 0):.0f}')),
                    ft.DataCell(ft.Text(money(record["estimated_value"]))),
                    ft.DataCell(ft.Text(record["owner"] or "")),
                    ft.DataCell(ft.Text(record["next_action_date"] or "")),
                    ft.DataCell(ft.Row(spacing=2, controls=[
                        ft.FilledButton(
                            "Converted" if converted else "Convert",
                            icon=ft.Icons.CHECK if converted else ft.Icons.TRENDING_UP,
                            disabled=converted,
                            on_click=(
                                None if converted
                                else lambda _, lid=record["id"]: self.convert_lead_to_opportunity(lid)
                            ),
                            style=ft.ButtonStyle(bgcolor="#8A9A96" if converted else "#0E5A47"),
                        ),
                        ft.IconButton(
                            icon=ft.Icons.EDIT,
                            icon_color="#1667A8",
                            on_click=lambda _, rid=record["id"]: self.edit_record_dialog("leads", rid),
                        ),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_OUTLINE,
                            icon_color="#C0392B",
                            on_click=lambda _, rid=record["id"]: self.confirm_delete_record("leads", rid),
                        ),
                    ])),
                ]))
            return output

        table.rows = build_rows(all_records)

        def search_changed(event):
            query = (event.control.value or "").strip().lower()
            filtered = [
                row for row in all_records
                if not query or query in " ".join([
                    str(row["lead_reference"] or ""),
                    str(row["account_name"] or ""),
                    str(row["company_name"] or ""),
                    str(row["contact_name"] or ""),
                    str(row["source"] or ""),
                    str(row["interest_area"] or ""),
                    str(row["owner"] or ""),
                ]).lower()
            ]
            table.rows = build_rows(filtered)
            table.update()

        return ft.Column([
            ft.Row([
                self.title(
                    "Leads",
                    "Search Leads, open the Lead Workspace, or convert a Lead to an Opportunity."
                ),
                ft.FilledButton(
                    "New Lead",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.lead_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.search_with_clear("Search Leads", search_changed, 420),
            ft.Container(
                bgcolor=ft.Colors.WHITE,
                padding=14,
                border_radius=18,
                content=(self.horizontal_table(table) if all_records else ft.Text("No leads created yet.")),
            ),
        ])

    def send_opportunity_to_quotation(self, opportunity_id):
        try:
            opportunity = self.db.one(
                """SELECT o.*,a.account_name
                   FROM opportunities o
                   LEFT JOIN accounts a ON a.id=o.account_id
                   WHERE o.id=?""",
                (opportunity_id,),
            )
            if not opportunity:
                raise ValueError("The selected opportunity no longer exists.")

            existing = self.db.one(
                """SELECT * FROM quotations
                   WHERE opportunity_id=?
                     AND status NOT IN ('WON','LOST','REJECTED','CANCELLED')
                   ORDER BY id DESC LIMIT 1""",
                (opportunity_id,),
            )
            if existing:
                self.notify(
                    f"A quotation request already exists: {existing['quotation_number']}.",
                    True,
                )
                return

            owner_options = self.owner_options()
            if not owner_options:
                raise ValueError("No active Owners are available.")

            request_number = self.db.peek_reference("quotations")
            assigned_owner = ft.Dropdown(
                label="Assign To",
                width=600,
                options=owner_options,
                value=owner_options[0].key,
            )
            target_control, target_date = self.calendar_field(
                "Target Submission Date",
                (date.today() + timedelta(days=7)).isoformat(),
                width=300,
            )
            notes = ft.TextField(
                label="Assignment Notes",
                width=600,
                multiline=True,
                min_lines=3,
            )

            def create_request(_):
                try:
                    if not assigned_owner.value:
                        raise ValueError("Select an assigned owner.")
                    valid_date(target_date.value)

                    self.db.insert_quotation({
                        "quotation_number": request_number,
                        "opportunity_id": opportunity_id,
                        "account_id": opportunity["account_id"],
                        "quotation_date": date.today().isoformat(),
                        "valid_until": target_date.value,
                        "base_value": 0,
                        "discount": 0,
                        "vat_rate": VAT_RATE,
                        "status": "ASSIGNED",
                        "revision": 0,
                        "owner": assigned_owner.value,
                        "cost_price": 0,
                        "gross_margin_percent": 0,
                        "gross_margin_value": 0,
                        "requested_by": self.current_owner_name(),
                        "concern_owner": opportunity["sales_owner"] or self.current_owner_name(),
                        "assigned_date": datetime.now().isoformat(timespec="seconds"),
                        "work_started_date": "",
                        "submitted_date": "",
                        "completed_date": "",
                        "assignment_notes": notes.value or "",
                        "notes": f"Created from Opportunity {opportunity['opportunity_reference']}.",
                    })

                    quotation = self.db.one(
                        "SELECT id FROM quotations WHERE quotation_number=?",
                        (request_number,),
                    )
                    related_id = quotation["id"] if quotation else opportunity_id

                    self.db.add_inbox_item(
                        assigned_owner.value,
                        "QUOTATION",
                        request_number,
                        "New Quotation Assignment",
                        f"Prepare a quotation for {opportunity['opportunity_name']}.",
                        "PREPARE QUOTATION",
                        related_id,
                    )
                    self.db.add_notification(
                        assigned_owner.value,
                        "New Quotation Assigned",
                        f"{request_number} has been assigned to you.",
                        "QUOTATION",
                        related_id,
                    )

                    self.page.close(dialog)
                    self.notify(
                        f"{request_number} created and assigned to {assigned_owner.value}."
                    )
                    self.navigate("quotations")
                except Exception as error:
                    self.notify(f"Could not create quotation request: {error}", True)

            dialog = ft.AlertDialog(
                modal=True,
                title=ft.Text("Send Opportunity to Quotation", weight=ft.FontWeight.BOLD),
                content=ft.Container(
                    width=680,
                    content=ft.Column(
                        tight=True,
                        controls=[
                            ft.Text(
                                f"Opportunity ID: {opportunity['opportunity_reference']}",
                                weight=ft.FontWeight.BOLD,
                            ),
                            ft.Text(f"Account: {opportunity['account_name'] or ''}"),
                            ft.Text(f"Opportunity: {opportunity['opportunity_name'] or ''}"),
                            ft.Text(f"Quotation ID: {request_number}"),
                            ft.Divider(),
                            assigned_owner,
                            target_control,
                            notes,
                        ],
                    ),
                ),
                actions=[
                    ft.TextButton("Cancel", on_click=lambda _: self.page.close(dialog)),
                    ft.FilledButton(
                        "Create & Assign",
                        icon=ft.Icons.REQUEST_QUOTE,
                        on_click=create_request,
                        style=ft.ButtonStyle(bgcolor="#0E5A47"),
                    ),
                ],
            )
            self.page.open(dialog)
        except Exception as error:
            self.notify(f"Could not open quotation assignment: {error}", True)

    def open_exported_file(self, path):
        path = Path(path)
        try:
            os.startfile(str(path))
            return
        except Exception:
            pass
        try:
            os.startfile(str(path.parent))
        except Exception:
            pass

    def show_export_success(self, title, path):
        path = Path(path)

        def open_file(_):
            self.page.close(dialog)
            self.open_exported_file(path)

        def open_folder(_):
            self.page.close(dialog)
            try:
                os.startfile(str(path.parent))
            except Exception as error:
                self.notify(f"Could not open export folder: {error}", True)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text(title, weight=ft.FontWeight.BOLD),
            content=ft.Container(
                width=650,
                content=ft.Column(
                    tight=True,
                    controls=[
                        ft.Text("The PDF was created successfully."),
                        ft.Text(
                            str(path),
                            selectable=True,
                            color="#1667A8",
                        ),
                    ],
                ),
            ),
            actions=[
                ft.TextButton(
                    "Close",
                    on_click=lambda _: self.page.close(dialog),
                ),
                ft.OutlinedButton(
                    "Open Folder",
                    icon=ft.Icons.FOLDER_OPEN,
                    on_click=open_folder,
                ),
                ft.FilledButton(
                    "Open PDF",
                    icon=ft.Icons.PICTURE_AS_PDF,
                    on_click=open_file,
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ],
        )
        self.page.open(dialog)

    def pdf_text(self, value):
        if value is None or str(value).strip() == "":
            return "Not available"
        return (
            str(value)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    def pdf_information_table(self, items, label_style, value_style):
        rows = []
        for index in range(0, len(items), 2):
            row = []
            for label, value in items[index:index + 2]:
                row.extend([
                    Paragraph(self.pdf_text(label), label_style),
                    Paragraph(self.pdf_text(value), value_style),
                ])
            while len(row) < 4:
                row.extend([
                    Paragraph("", label_style),
                    Paragraph("", value_style),
                ])
            rows.append(row)

        table = Table(
            rows,
            colWidths=[37 * mm, 91 * mm, 37 * mm, 91 * mm],
            repeatRows=0,
        )
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B9C9C4")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E7F0ED")),
            ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#E7F0ED")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return table

    def report_header(self, title, reference, styles):
        controls = []
        logo_file = (
            REPORT_LOGO_PATH
            if REPORT_LOGO_PATH.exists()
            else LOGO_PATH
        )
        if logo_file.exists():
            controls.append(
                Table(
                    [[Image(
                        str(logo_file),
                        width=55 * mm,
                        height=25 * mm,
                        kind="proportional",
                    )]],
                    colWidths=[landscape(A4)[0] - 24 * mm],
                    style=TableStyle([
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ]),
                )
            )
            controls.append(Spacer(1, 2 * mm))

        controls.extend([
            Paragraph(
                self.pdf_text(title),
                ParagraphStyle(
                    "ExportTitle",
                    parent=styles["Title"],
                    fontName="Helvetica-Bold",
                    fontSize=20,
                    leading=24,
                    textColor=colors.HexColor("#103F37"),
                    alignment=1,
                    spaceAfter=2 * mm,
                ),
            ),
            Paragraph(
                self.pdf_text(reference),
                ParagraphStyle(
                    "ExportReference",
                    parent=styles["BodyText"],
                    fontName="Helvetica-Bold",
                    fontSize=11,
                    textColor=colors.HexColor("#1667A8"),
                    alignment=1,
                    spaceAfter=5 * mm,
                ),
            ),
        ])
        return controls

    def export_opportunity_pdf(self, opportunity_id):
        try:
            EXPORT_DIR.mkdir(parents=True, exist_ok=True)
            record = self.db.opportunity_full_details(opportunity_id)
            if not record:
                raise ValueError("Opportunity not found.")

            path = EXPORT_DIR / (
                f"{record['opportunity_reference']}_Opportunity_Details.pdf"
            )
            styles = getSampleStyleSheet()
            section_style = ParagraphStyle(
                "OpportunitySection",
                parent=styles["Heading2"],
                fontName="Helvetica-Bold",
                fontSize=13,
                leading=16,
                textColor=colors.HexColor("#103F37"),
                spaceBefore=4 * mm,
                spaceAfter=2 * mm,
            )
            label_style = ParagraphStyle(
                "OpportunityLabel",
                parent=styles["BodyText"],
                fontName="Helvetica-Bold",
                fontSize=8,
                leading=10,
                textColor=colors.HexColor("#397064"),
            )
            value_style = ParagraphStyle(
                "OpportunityValue",
                parent=styles["BodyText"],
                fontSize=8,
                leading=10,
                wordWrap="CJK",
            )
            notes_style = ParagraphStyle(
                "OpportunityNotes",
                parent=styles["BodyText"],
                fontSize=9,
                leading=12,
                borderColor=colors.HexColor("#B9C9C4"),
                borderWidth=0.5,
                borderPadding=7,
                backColor=colors.HexColor("#F6F9F8"),
            )

            estimated_value = float(record["estimated_value"] or 0)
            probability = float(record["probability"] or 0)
            weighted_value = estimated_value * probability / 100

            story = self.report_header(
                "Opportunity Details",
                record["opportunity_reference"],
                styles,
            )
            story.extend([
                Paragraph("CRM Opportunity Overview", section_style),
                self.pdf_information_table([
                    ("Opportunity Name", record["opportunity_name"]),
                    ("Account", record["account_name"]),
                    ("Account ID", record["account_reference"]),
                    ("Original Reference", record["original_opportunity_reference"]),
                    ("Project Type", record["project_type"]),
                    ("Stage", record["stage"]),
                    ("Status", record["status"]),
                    ("Probability", f"{probability:.0f}%"),
                    ("Estimated Value", money(estimated_value)),
                    ("Weighted Value", money(weighted_value)),
                    ("Expected Close Date", record["expected_close_date"]),
                    ("Sales Owner", record["sales_owner"]),
                    ("Technical Owner", record["technical_owner"]),
                    ("Next Action Date", record["next_action_date"]),
                ], label_style, value_style),
                Paragraph("Original Excel Information", section_style),
                self.pdf_information_table([
                    ("Source Row", record["source_row"]),
                    ("Source Created Date", record["source_created_date"]),
                    ("Customer Name in Source", record["customer_name_source"]),
                    ("Account Basis", record["account_basis"]),
                    ("Business Unit", record["business_unit"]),
                    ("Source Project Type", record["source_project_type"]),
                    ("End User", record["end_user"]),
                    ("Industry", record["source_industry"]),
                    ("Competitive", "Yes" if record["competitive"] else "No"),
                    ("Probability Band", record["probability_band"]),
                    ("Source Currency", record["source_currency"]),
                    ("Original Stage", record["source_stage"]),
                    ("Forecast GM %", record["forecast_gm_percent"]),
                    ("Forecast GM Value", money(record["forecast_gm_value"] or 0)),
                    ("GM Value Basis", record["gm_value_basis"]),
                    ("Expected PO Year", record["expected_po_year"]),
                    ("Expected PO Month", record["expected_po_month"]),
                    ("Quarter", record["quarter"]),
                    ("Delivery Date", record["delivery_date"]),
                    ("Created By", record["created_by"]),
                    ("Assigned To", record["assigned_to"]),
                    (
                        "Include in Forecast",
                        "Yes" if record["include_in_forecast"] else "No",
                    ),
                    ("Must Win", "Yes" if record["must_win"] else "No"),
                    ("Suspended", "Yes" if record["suspended"] else "No"),
                ], label_style, value_style),
                Paragraph("Opportunity Update", section_style),
                Paragraph(
                    self.pdf_text(record["opportunity_update"]),
                    notes_style,
                ),
                Paragraph("CRM Notes", section_style),
                Paragraph(
                    self.pdf_text(record["notes"]),
                    notes_style,
                ),
                Paragraph("Data Quality Flags", section_style),
                Paragraph(
                    self.pdf_text(record["quality_flags"]),
                    notes_style,
                ),
            ])

            document = SimpleDocTemplate(
                str(path),
                pagesize=landscape(A4),
                leftMargin=10 * mm,
                rightMargin=10 * mm,
                topMargin=8 * mm,
                bottomMargin=10 * mm,
                title=f"Opportunity {record['opportunity_reference']}",
                author="Saudi Sensing Solution Company Ltd",
            )
            document.build(story)
            self.show_export_success(
                "Opportunity PDF Exported",
                path,
            )
        except Exception as error:
            error_path = APP_DIR / "opportunity_pdf_export_error.log"
            try:
                error_path.write_text(
                    f"{datetime.now().isoformat(timespec='seconds')}\n"
                    f"{type(error).__name__}: {error}",
                    encoding="utf-8",
                )
            except Exception:
                pass
            self.notify(
                f"Opportunity PDF export failed: {error}",
                True,
            )

    def export_quotation_pdf(self, quotation_id):
        try:
            EXPORT_DIR.mkdir(parents=True, exist_ok=True)
            record = self.db.one(
                """SELECT q.*,a.account_name,a.account_reference,
                          o.opportunity_reference,o.opportunity_name,
                          o.project_type,o.sales_owner,o.technical_owner
                   FROM quotations q
                   LEFT JOIN accounts a ON a.id=q.account_id
                   LEFT JOIN opportunities o ON o.id=q.opportunity_id
                   WHERE q.id=?""",
                (quotation_id,),
            )
            if not record:
                raise ValueError("Quotation not found.")

            documents = self.db.quotation_documents(quotation_id)
            path = EXPORT_DIR / (
                f"{record['quotation_number']}_Quotation_Summary.pdf"
            )
            styles = getSampleStyleSheet()
            section_style = ParagraphStyle(
                "QuotationSection",
                parent=styles["Heading2"],
                fontName="Helvetica-Bold",
                fontSize=13,
                leading=16,
                textColor=colors.HexColor("#103F37"),
                spaceBefore=4 * mm,
                spaceAfter=2 * mm,
            )
            label_style = ParagraphStyle(
                "QuotationLabel",
                parent=styles["BodyText"],
                fontName="Helvetica-Bold",
                fontSize=8,
                leading=10,
                textColor=colors.HexColor("#397064"),
            )
            value_style = ParagraphStyle(
                "QuotationValue",
                parent=styles["BodyText"],
                fontSize=8,
                leading=10,
                wordWrap="CJK",
            )
            notes_style = ParagraphStyle(
                "QuotationNotes",
                parent=styles["BodyText"],
                fontSize=9,
                leading=12,
                borderColor=colors.HexColor("#B9C9C4"),
                borderWidth=0.5,
                borderPadding=7,
                backColor=colors.HexColor("#F6F9F8"),
            )

            selling_price = float(record["base_value"] or 0)
            cost_price = float(record["cost_price"] or 0)
            discount = float(record["discount"] or 0)
            vat_rate = float(record["vat_rate"] or 15)
            net_before_vat = max(0, selling_price - discount)
            vat_amount = net_before_vat * vat_rate / 100
            total_including_vat = net_before_vat + vat_amount
            gross_margin_value = net_before_vat - cost_price
            gross_margin_percent = (
                gross_margin_value / net_before_vat * 100
                if net_before_vat
                else 0
            )

            story = self.report_header(
                "Quotation Summary",
                record["quotation_number"],
                styles,
            )
            story.extend([
                Paragraph("Quotation and Opportunity", section_style),
                self.pdf_information_table([
                    ("Quotation ID", record["quotation_number"]),
                    ("Opportunity ID", record["opportunity_reference"]),
                    ("Opportunity", record["opportunity_name"]),
                    ("Project Type", record["project_type"]),
                    ("Account ID", record["account_reference"]),
                    ("Account", record["account_name"]),
                    ("Request Date", record["quotation_date"]),
                    ("Target Submission Date", record["valid_until"]),
                    ("Assigned Owner", record["owner"]),
                    ("Concern Owner", record["concern_owner"]),
                    ("Workflow Status", record["status"]),
                    ("Revision", record["revision"]),
                ], label_style, value_style),
                Paragraph("Commercial Summary", section_style),
                self.pdf_information_table([
                    ("Selling Price Before VAT", money(selling_price)),
                    ("Cost Price", money(cost_price)),
                    ("Discount", money(discount)),
                    ("Net Before VAT", money(net_before_vat)),
                    ("VAT Rate", f"{vat_rate:.2f}%"),
                    ("VAT Amount", money(vat_amount)),
                    ("Total Including VAT", money(total_including_vat)),
                    ("Gross Margin %", f"{gross_margin_percent:.2f}%"),
                    ("Gross Margin Value", money(gross_margin_value)),
                    ("Sales Owner", record["sales_owner"]),
                    ("Technical Owner", record["technical_owner"]),
                    ("Assigned Date", record["assigned_date"]),
                    ("Work Started Date", record["work_started_date"]),
                    ("Submitted Date", record["submitted_date"]),
                    ("Completed Date", record["completed_date"]),
                ], label_style, value_style),
                Paragraph("Approval Status", section_style),
                self.pdf_information_table([
                    (
                        "Proposal Manager Approval",
                        record["proposal_manager_approval"],
                    ),
                    (
                        "Proposal Manager Approved By",
                        record["proposal_manager_approved_by"],
                    ),
                    (
                        "Proposal Manager Approved At",
                        record["proposal_manager_approved_at"],
                    ),
                    ("GM Approval", record["gm_approval"]),
                    ("GM Approved By", record["gm_approved_by"]),
                    ("GM Approved At", record["gm_approved_at"]),
                ], label_style, value_style),
                Paragraph("Assignment Notes", section_style),
                Paragraph(
                    self.pdf_text(record["assignment_notes"]),
                    notes_style,
                ),
                Paragraph("Commercial / Proposal Notes", section_style),
                Paragraph(
                    self.pdf_text(record["notes"]),
                    notes_style,
                ),
                Paragraph("Approval Comments", section_style),
                Paragraph(
                    self.pdf_text(record["approval_comments"]),
                    notes_style,
                ),
                Paragraph("Uploaded Documents", section_style),
            ])

            if documents:
                document_rows = [["Type", "File Name", "Uploaded By", "Uploaded At"]]
                for item in documents:
                    document_rows.append([
                        Paragraph(
                            self.pdf_text(item["document_type"]),
                            value_style,
                        ),
                        Paragraph(
                            self.pdf_text(item["original_file_name"]),
                            value_style,
                        ),
                        Paragraph(
                            self.pdf_text(item["uploaded_by"]),
                            value_style,
                        ),
                        Paragraph(
                            self.pdf_text(item["uploaded_at"]),
                            value_style,
                        ),
                    ])
                documents_table = Table(
                    document_rows,
                    colWidths=[45 * mm, 110 * mm, 50 * mm, 55 * mm],
                    repeatRows=1,
                )
                documents_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#103F37")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B9C9C4")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]))
                story.append(documents_table)
            else:
                story.append(
                    Paragraph(
                        "No documents have been uploaded.",
                        notes_style,
                    )
                )

            document = SimpleDocTemplate(
                str(path),
                pagesize=landscape(A4),
                leftMargin=10 * mm,
                rightMargin=10 * mm,
                topMargin=8 * mm,
                bottomMargin=10 * mm,
                title=f"Quotation {record['quotation_number']}",
                author="Saudi Sensing Solution Company Ltd",
            )
            document.build(story)
            self.show_export_success(
                "Quotation PDF Exported",
                path,
            )
        except Exception as error:
            error_path = APP_DIR / "quotation_pdf_export_error.log"
            try:
                error_path.write_text(
                    f"{datetime.now().isoformat(timespec='seconds')}\n"
                    f"{type(error).__name__}: {error}",
                    encoding="utf-8",
                )
            except Exception:
                pass
            self.notify(
                f"Quotation PDF export failed: {error}",
                True,
            )

    def opportunity_workspace(self, opportunity_id):
        try:
            record = self.db.opportunity_full_details(opportunity_id)
            if not record:
                self.notify("The opportunity no longer exists.", True)
                return

            def display(value, empty="Not available"):
                if value is None or str(value).strip() == "":
                    return empty
                return str(value)

            def yes_no(value):
                return "Yes" if value else "No"

            def detail_card(label, value, width=245, accent="#103F37"):
                return ft.Container(
                    width=width,
                    padding=13,
                    bgcolor="#F5F8F7",
                    border_radius=11,
                    content=ft.Column(
                        spacing=4,
                        controls=[
                            ft.Text(label, size=11, color="#667B80"),
                            ft.Text(
                                display(value),
                                size=14,
                                weight=ft.FontWeight.BOLD,
                                color=accent,
                                selectable=True,
                            ),
                        ],
                    ),
                )

            def section(title, controls):
                return ft.Container(
                    padding=16,
                    bgcolor=ft.Colors.WHITE,
                    border_radius=16,
                    content=ft.Column(
                        spacing=12,
                        controls=[
                            ft.Text(
                                title,
                                size=19,
                                weight=ft.FontWeight.BOLD,
                                color="#103F37",
                            ),
                            *controls,
                        ],
                    ),
                )

            gross_value = float(record["estimated_value"] or 0)
            probability = float(record["probability"] or 0)
            weighted_value = gross_value * probability / 100
            gm_percent = float(record["forecast_gm_percent"] or 0)
            gm_value = float(record["forecast_gm_value"] or 0)

            overview = section(
                "CRM Opportunity Overview",
                [
                    ft.Row(
                        wrap=True,
                        controls=[
                            detail_card("Opportunity ID", record["opportunity_reference"]),
                            detail_card(
                                "Original Excel Reference",
                                record["original_opportunity_reference"],
                            ),
                            detail_card("Account ID", record["account_reference"]),
                            detail_card("Account", record["account_name"], width=330),
                        ],
                    ),
                    ft.Text(
                        display(record["opportunity_name"]),
                        size=22,
                        weight=ft.FontWeight.BOLD,
                        color="#103F37",
                        selectable=True,
                    ),
                    ft.Row(
                        wrap=True,
                        controls=[
                            detail_card("Project Type", record["project_type"]),
                            detail_card("CRM Stage", record["stage"]),
                            detail_card("CRM Status", record["status"]),
                            detail_card("Probability", f"{probability:.0f}%"),
                            detail_card("Gross Value", money(gross_value)),
                            detail_card("Weighted Value", money(weighted_value)),
                            detail_card(
                                "Expected Close Date",
                                record["expected_close_date"],
                            ),
                            detail_card("Sales Owner", record["sales_owner"]),
                            detail_card("Technical Owner", record["technical_owner"]),
                        ],
                    ),
                    ft.Text(
                        f"Next Step: {display(record['next_step'])}",
                        color="#304B45",
                        selectable=True,
                    ),
                    ft.Text(
                        f"Notes: {display(record['notes'])}",
                        color="#304B45",
                        selectable=True,
                    ),
                ],
            )

            original_excel = section(
                "Original Excel Data",
                [
                    ft.Row(
                        wrap=True,
                        controls=[
                            detail_card("Source Row", record["source_row"]),
                            detail_card("Source Created Date", record["source_created_date"]),
                            detail_card("Customer Name in Source", record["customer_name_source"], 330),
                            detail_card("Account Basis", record["account_basis"]),
                            detail_card("Business Unit", record["business_unit"]),
                            detail_card("Source Project Type", record["source_project_type"]),
                            detail_card("End User", record["end_user"], 330),
                            detail_card("Industry", record["source_industry"]),
                            detail_card("Competitive", yes_no(record["competitive"])),
                            detail_card("Probability Band", record["probability_band"]),
                            detail_card("Source Currency", record["source_currency"]),
                            detail_card("Original Stage", record["source_stage"]),
                        ],
                    ),
                    ft.Divider(),
                    ft.Text(
                        "Forecast and Delivery",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color="#397064",
                    ),
                    ft.Row(
                        wrap=True,
                        controls=[
                            detail_card("Forecast GM %", f"{gm_percent:.2f}%"),
                            detail_card("Forecast GM Value", money(gm_value)),
                            detail_card("GM Value Basis", record["gm_value_basis"]),
                            detail_card("Expected PO Year", record["expected_po_year"]),
                            detail_card("Expected PO Month", record["expected_po_month"]),
                            detail_card("Quarter", record["quarter"]),
                            detail_card("Delivery Date", record["delivery_date"]),
                            detail_card(
                                "Include in Forecast",
                                yes_no(record["include_in_forecast"]),
                            ),
                            detail_card("Must Win", yes_no(record["must_win"])),
                            detail_card("Suspended", yes_no(record["suspended"])),
                        ],
                    ),
                    ft.Divider(),
                    ft.Text(
                        "Ownership and Updates from Source",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color="#397064",
                    ),
                    ft.Row(
                        wrap=True,
                        controls=[
                            detail_card("Created By", record["created_by"]),
                            detail_card("Assigned To", record["assigned_to"]),
                        ],
                    ),
                    ft.Container(
                        padding=13,
                        bgcolor="#F5F8F7",
                        border_radius=11,
                        content=ft.Column(
                            controls=[
                                ft.Text(
                                    "Opportunity Update",
                                    size=11,
                                    color="#667B80",
                                ),
                                ft.Text(
                                    display(record["opportunity_update"]),
                                    selectable=True,
                                    color="#103F37",
                                ),
                            ],
                        ),
                    ),
                    ft.Container(
                        padding=13,
                        bgcolor=(
                            "#FFF4D6"
                            if record["quality_flags"]
                            else "#EAF2EF"
                        ),
                        border_radius=11,
                        content=ft.Column(
                            controls=[
                                ft.Text(
                                    "Data Quality Flags",
                                    size=11,
                                    color="#667B80",
                                ),
                                ft.Text(
                                    display(record["quality_flags"], "No quality flags"),
                                    selectable=True,
                                    color="#8A4B08" if record["quality_flags"] else "#2E8B57",
                                    weight=ft.FontWeight.BOLD,
                                ),
                            ],
                        ),
                    ),
                ],
            )

            account_section = section(
                "Linked Account Information",
                [
                    ft.Row(
                        wrap=True,
                        controls=[
                            detail_card("Account Type", record["account_type"]),
                            detail_card("Account Industry", record["account_industry"]),
                            detail_card("City", record["account_city"]),
                            detail_card("Country", record["account_country"]),
                        ],
                    ),
                ],
            )

            workspace = ft.Column(
                expand=True,
                scroll=ft.ScrollMode.AUTO,
                controls=[
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                controls=[
                                    ft.Text(
                                        f"Opportunity Workspace | {record['opportunity_reference']}",
                                        size=28,
                                        weight=ft.FontWeight.BOLD,
                                        color="#103F37",
                                    ),
                                    ft.Text(
                                        "Complete CRM and original Excel opportunity information",
                                        color="#667B80",
                                    ),
                                ],
                            ),
                            ft.Row(
                                controls=[
                                    ft.OutlinedButton(
                                        "Back to Opportunities",
                                        icon=ft.Icons.ARROW_BACK,
                                        on_click=lambda _: self.navigate("opportunities"),
                                    ),
                                    ft.FilledButton(
                                        "Export PDF",
                                        icon=ft.Icons.PICTURE_AS_PDF,
                                        on_click=lambda _: self.export_opportunity_pdf(
                                            opportunity_id
                                        ),
                                        style=ft.ButtonStyle(bgcolor="#1667A8"),
                                    ),
                                    ft.FilledButton(
                                        "Send to Quotation",
                                        icon=ft.Icons.REQUEST_QUOTE,
                                        on_click=lambda _: self.send_opportunity_to_quotation(
                                            opportunity_id
                                        ),
                                        style=ft.ButtonStyle(bgcolor="#0E5A47"),
                                    ),
                                    ft.IconButton(
                                        icon=ft.Icons.EDIT,
                                        tooltip="Edit Opportunity",
                                        on_click=lambda _: self.edit_record_dialog(
                                            "opportunities",
                                            opportunity_id,
                                        ),
                                    ),
                                ],
                            ),
                        ],
                    ),
                    overview,
                    original_excel,
                    account_section,
                ],
            )

            self.current_view = "opportunity_workspace"
            self.content.content = workspace
            self.page.update()

        except Exception as error:
            error_path = APP_DIR / "opportunity_workspace_error.log"
            try:
                error_path.write_text(
                    f"{datetime.now().isoformat(timespec='seconds')}\n"
                    f"{type(error).__name__}: {error}",
                    encoding="utf-8",
                )
            except Exception:
                pass
            self.notify(
                f"Could not open opportunity workspace: {error}. "
                "Details saved in opportunity_workspace_error.log",
                True,
            )

    def opportunities_view(self):
        all_records = self.db.opportunities()

        columns = [
            ft.DataColumn(ft.Container(width=125, content=ft.Text("Opportunity ID", weight=ft.FontWeight.BOLD))),
            ft.DataColumn(ft.Container(width=190, padding=ft.Padding(left=14,right=6,top=0,bottom=0), content=ft.Text("Account", weight=ft.FontWeight.BOLD))),
            ft.DataColumn(ft.Container(width=285, content=ft.Text("Opportunity", weight=ft.FontWeight.BOLD))),
            ft.DataColumn(ft.Text("Type", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Stage", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Probability", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Value", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Weighted Value", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Close Date", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Sales Owner", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Next Step", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Actions", weight=ft.FontWeight.BOLD)),
        ]

        table = ft.DataTable(
            columns=columns,
            rows=[],
            column_spacing=14,
            heading_row_color="#E5F0EC",
        )

        active_request_rows = self.db.rows(
            """SELECT DISTINCT opportunity_id
               FROM quotations
               WHERE opportunity_id IS NOT NULL
                 AND status NOT IN ('WON','LOST','REJECTED','CANCELLED')"""
        )
        active_request_ids = {
            row["opportunity_id"]
            for row in active_request_rows
            if row["opportunity_id"] is not None
        }

        def build_rows(records):
            result = []
            for record in records:
                has_request = record["id"] in active_request_ids

                opportunity_id = self.id_badge(
                    record["opportunity_reference"],
                    width=125,
                    clickable=lambda _, oid=record["id"]: self.opportunity_workspace(oid),
                    tooltip="Open Opportunity Workspace",
                )

                send_button = ft.FilledButton(
                    "Sent" if has_request else "Send to Quotation",
                    icon=ft.Icons.CHECK if has_request else ft.Icons.REQUEST_QUOTE,
                    disabled=has_request,
                    on_click=(
                        None
                        if has_request
                        else lambda _, oid=record["id"]: self.send_opportunity_to_quotation(oid)
                    ),
                    style=ft.ButtonStyle(
                        bgcolor="#8A9A96" if has_request else "#0E5A47"
                    ),
                )

                result.append(
                    ft.DataRow(cells=[
                        ft.DataCell(opportunity_id),
                        ft.DataCell(ft.Container(
                            width=190,
                            padding=ft.Padding(left=14,right=6,top=4,bottom=4),
                            content=ft.Text(record["account_name"] or "", selectable=True),
                        )),
                        ft.DataCell(ft.Container(width=285, content=ft.Text(record["opportunity_name"] or "", selectable=True))),
                        ft.DataCell(ft.Text(record["project_type"] or "")),
                        ft.DataCell(ft.Text(record["stage"] or "")),
                        ft.DataCell(ft.Text(f'{float(record["probability"] or 0):.0f}%')),
                        ft.DataCell(ft.Text(money(record["estimated_value"]))),
                        ft.DataCell(ft.Text(money(
                            float(record["estimated_value"] or 0)
                            * float(record["probability"] or 0) / 100
                        ))),
                        ft.DataCell(ft.Text(record["expected_close_date"] or "")),
                        ft.DataCell(ft.Text(record["sales_owner"] or "")),
                        ft.DataCell(ft.Container(width=180, content=ft.Text(record["next_step"] or ""))),
                        ft.DataCell(ft.Row(
                            spacing=2,
                            controls=[
                                send_button,
                                ft.IconButton(
                                    icon=ft.Icons.EDIT,
                                    tooltip="Edit",
                                    icon_color="#1667A8",
                                    on_click=lambda _, rid=record["id"]: self.edit_record_dialog("opportunities", rid),
                                ),
                                ft.IconButton(
                                    icon=ft.Icons.DELETE_OUTLINE,
                                    tooltip="Delete",
                                    icon_color="#C0392B",
                                    on_click=lambda _, rid=record["id"]: self.confirm_delete_record("opportunities", rid),
                                ),
                            ],
                        )),
                    ])
                )
            return result

        table.rows = build_rows(all_records)

        def search_changed(event):
            query = (event.control.value or "").strip().lower()
            filtered = [
                row for row in all_records
                if not query or query in " ".join([
                    str(row["opportunity_reference"] or ""),
                    str(row["account_name"] or ""),
                    str(row["opportunity_name"] or ""),
                    str(row["project_type"] or ""),
                    str(row["stage"] or ""),
                    str(row["sales_owner"] or ""),
                ]).lower()
            ]
            table.rows = build_rows(filtered)
            table.update()

        return ft.Column([
            ft.Row([
                self.title(
                    "Opportunities",
                    "Search opportunities, open full details, or send an opportunity for quotation preparation."
                ),
                ft.FilledButton(
                    "New Opportunity",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.opportunity_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.search_with_clear("Search Opportunities", search_changed, 440),
            ft.Container(
                padding=12,
                bgcolor="#EAF2EF",
                border_radius=12,
                content=ft.Text(
                    "Click the blue Opportunity ID to open its workspace. "
                    "Send to Quotation remains available in the Actions column."
                ),
            ),
            ft.Container(
                bgcolor=ft.Colors.WHITE,
                padding=14,
                border_radius=18,
                content=(self.horizontal_table(table) if all_records else ft.Text("No opportunities created yet.")),
            ),
        ])

    def legacy_pipeline_view(self):
        rows = self.db.legacy_opportunities()
        quality_issues = sum(1 for r in rows if r["quality_flags"])
        missing_accounts = sum(1 for r in rows if not r["account_name"])
        gm_issues = sum(
            1 for r in rows
            if r["quality_flags"] and "GM value differs" in r["quality_flags"]
        )
        return ft.Column([
            self.title(
                "Migrated Opportunity Pipeline",
                "The uploaded opportunity file is preserved, cleaned and mapped into the CRM."
            ),
            ft.Row([
                self.card("Imported Records", str(len(rows)), ft.Icons.TABLE_CHART, "#0E5A47"),
                self.card("Records with Quality Flags", str(quality_issues), ft.Icons.WARNING_AMBER, "#D35400"),
                self.card("Missing Account", str(missing_accounts), ft.Icons.ERROR_OUTLINE, "#C0392B"),
                self.card("GM Discrepancies", str(gm_issues), ft.Icons.CALCULATE, "#8E44AD"),
            ]),
            self.table_container(
                [
                    "Reference","Account","End User","Deal Name","Business Unit","Source Type",
                    "Industry","Source Stage","CRM Stage","Probability","Gross","GM %","GM Value",
                    "Expected PO","Delivery","Sales Owner","Assigned To","Forecast","Must Win",
                    "Suspended","Quality Flags"
                ],
                [[
                    r["opportunity_reference"],r["account_name"],r["end_user"],r["opportunity_name"],
                    r["business_unit"],r["source_project_type"],r["industry"],r["source_stage"],
                    r["stage"],f'{r["probability"]:.0f}%',money(r["estimated_value"]),
                    f'{r["forecast_gm_percent"]:.1f}%',money(r["forecast_gm_value"]),
                    f'{r["expected_po_month"]} {r["expected_po_year"]}',r["delivery_date"],
                    r["sales_owner"],r["assigned_to"],"Yes" if r["include_in_forecast"] else "No",
                    "Yes" if r["must_win"] else "No","Yes" if r["suspended"] else "No",
                    r["quality_flags"]
                ] for r in rows],
                "No migrated pipeline records are available."
            )
        ])

    def quotation_age_days(self, record):
        source = record.get("assigned_date") or record.get("created_at") or ""
        if not source:
            return 0
        try:
            started = datetime.fromisoformat(source).date()
            return (date.today() - started).days
        except Exception:
            return 0

    def update_quotation_workflow_dates(self, record_id, status):
        values = {}
        now = datetime.now().isoformat(timespec="seconds")
        if status == "IN PROGRESS":
            values["work_started_date"] = now
        elif status == "SUBMITTED":
            values["submitted_date"] = now
        elif status in ("WON", "LOST", "CANCELLED"):
            values["completed_date"] = now
        if values:
            self.db.update_record("quotations", record_id, values)

    def owner_by_role(self, role_name):
        for owner in self.db.owners(active_only=True):
            if (owner["role"] or "").strip().lower() == role_name.strip().lower():
                return owner["full_name"]
        return ""

    def send_quotation_to_concern_owner(self, quotation_id, record):
        concern_owner = record.get("concern_owner") or record.get("requested_by") or record.get("sales_owner") or ""
        if not concern_owner:
            concern_owner = record.get("sales_owner") or ""
        self.db.close_inbox_items(quotation_id)
        self.db.add_inbox_item(
            concern_owner,
            "QUOTATION",
            record["quotation_number"],
            "Quotation Ready for Submission",
            f"{record['quotation_number']} is ready for your review and submission.",
            "READY FOR SUBMISSION",
            quotation_id,
        )
        self.db.add_notification(
            concern_owner,
            "Quotation Ready",
            f"{record['quotation_number']} is ready for submission.",
            "QUOTATION",
            quotation_id,
        )

    def submit_for_proposal_manager_approval(self, quotation_id, record):
        manager = self.owner_by_role("Proposal Manager")
        if not manager:
            raise ValueError("No active Proposal Manager is configured in Owners.")
        self.db.close_inbox_items(quotation_id)
        self.db.update_record(
            "quotations",
            quotation_id,
            {
                "status": "PENDING PROPOSAL MANAGER APPROVAL",
                "proposal_manager_approval": "PENDING",
            },
        )
        self.db.add_inbox_item(
            manager,
            "QUOTATION",
            record["quotation_number"],
            "Proposal Manager Approval Required",
            f"Please review and approve {record['quotation_number']}.",
            "PROPOSAL MANAGER APPROVAL",
            quotation_id,
        )
        self.db.add_notification(
            manager,
            "Approval Required",
            f"{record['quotation_number']} requires your approval.",
            "QUOTATION",
            quotation_id,
        )

    def proposal_manager_approve(self, quotation_id, record, comments=""):
        gm = self.owner_by_role("General Manager")
        if not gm:
            raise ValueError("No active General Manager is configured in Owners.")
        now = datetime.now().isoformat(timespec="seconds")
        self.db.close_inbox_items(quotation_id)
        self.db.update_record(
            "quotations",
            quotation_id,
            {
                "status": "PENDING GM APPROVAL",
                "proposal_manager_approval": "APPROVED",
                "proposal_manager_approved_by": self.current_owner_name(),
                "proposal_manager_approved_at": now,
                "approval_comments": comments,
                "gm_approval": "PENDING",
            },
        )
        self.db.add_inbox_item(
            gm,
            "QUOTATION",
            record["quotation_number"],
            "GM Final Approval Required",
            f"{record['quotation_number']} was approved by the Proposal Manager and requires final approval.",
            "GM APPROVAL",
            quotation_id,
        )
        self.db.add_notification(
            gm,
            "Final Approval Required",
            f"{record['quotation_number']} requires your final approval.",
            "QUOTATION",
            quotation_id,
        )

    def gm_approve(self, quotation_id, record, comments=""):
        now = datetime.now().isoformat(timespec="seconds")
        concern_owner = record.get("concern_owner") or record.get("requested_by") or record.get("sales_owner") or ""
        self.db.close_inbox_items(quotation_id)
        self.db.update_record(
            "quotations",
            quotation_id,
            {
                "status": "APPROVED FOR SUBMISSION",
                "gm_approval": "APPROVED",
                "gm_approved_by": self.current_owner_name(),
                "gm_approved_at": now,
                "approval_comments": comments,
            },
        )
        self.db.add_inbox_item(
            concern_owner,
            "QUOTATION",
            record["quotation_number"],
            "Quotation Approved for Submission",
            f"{record['quotation_number']} has final approval and is ready to submit to the customer.",
            "SUBMIT TO CUSTOMER",
            quotation_id,
        )
        self.db.add_notification(
            concern_owner,
            "Approved for Submission",
            f"{record['quotation_number']} received final GM approval.",
            "QUOTATION",
            quotation_id,
        )

    def quotation_workspace(self, quotation_id):
        try:
            record = self.db.one(
                """SELECT q.*,a.account_name,o.opportunity_reference,
                          o.opportunity_name,o.project_type,
                          o.stage AS opportunity_stage,o.probability,
                          o.estimated_value AS opportunity_value,
                          o.expected_close_date,o.sales_owner,
                          o.technical_owner,
                          o.next_step AS opportunity_next_step
                   FROM quotations q
                   LEFT JOIN accounts a ON a.id=q.account_id
                   LEFT JOIN opportunities o ON o.id=q.opportunity_id
                   WHERE q.id=?""",
                (quotation_id,),
            )
            if not record:
                self.notify("The quotation no longer exists.", True)
                return

            current_owner = self.current_owner_name()
            owner_options = self.owner_options()

            assigned_owner = ft.Dropdown(
                label="Assigned Owner",
                width=360,
                value=record["owner"] or None,
                options=owner_options,
            )
            concern_owner = ft.Dropdown(
                label="Concern Owner / Requester",
                width=360,
                value=record["concern_owner"] or record["sales_owner"] or None,
                options=owner_options,
            )
            workflow_status = ft.Dropdown(
                label="Workflow Status",
                width=350,
                value=record["status"] or "ASSIGNED",
                options=[
                    ft.dropdown.Option(item)
                    for item in QUOTATION_STATUSES
                ],
            )
            target_date_row, target_date = self.calendar_field(
                "Target Submission Date",
                record["valid_until"]
                or (date.today() + timedelta(days=7)).isoformat(),
                width=300,
            )

            selling_price = ft.TextField(
                label="Total Price Before VAT",
                value=str(float(record["base_value"] or 0)),
                width=260,
                keyboard_type=ft.KeyboardType.NUMBER,
            )
            cost_price = ft.TextField(
                label="Cost Price",
                value=str(float(record["cost_price"] or 0)),
                width=240,
                keyboard_type=ft.KeyboardType.NUMBER,
            )
            discount = ft.TextField(
                label="Discount",
                value=str(float(record["discount"] or 0)),
                width=210,
                keyboard_type=ft.KeyboardType.NUMBER,
            )
            vat_rate = ft.TextField(
                label="VAT %",
                value=str(float(record["vat_rate"] or 15)),
                width=150,
                keyboard_type=ft.KeyboardType.NUMBER,
            )
            revision = ft.TextField(
                label="Revision",
                value=str(int(record["revision"] or 0)),
                width=140,
                keyboard_type=ft.KeyboardType.NUMBER,
            )
            assignment_notes = ft.TextField(
                label="Assignment Notes",
                value=record["assignment_notes"] or "",
                width=800,
                multiline=True,
                min_lines=3,
            )
            commercial_notes = ft.TextField(
                label="Commercial / Proposal Notes",
                value=record["notes"] or "",
                width=800,
                multiline=True,
                min_lines=4,
            )
            approval_comments = ft.TextField(
                label="Approval Comments",
                value=record["approval_comments"] or "",
                width=800,
                multiline=True,
                min_lines=3,
            )

            net_text = ft.Text("", size=18, weight=ft.FontWeight.BOLD, color="#103F37")
            vat_text = ft.Text("", size=18, weight=ft.FontWeight.BOLD, color="#1667A8")
            total_text = ft.Text("", size=20, weight=ft.FontWeight.BOLD, color="#0E5A47")
            gm_percent_text = ft.Text("", size=18, weight=ft.FontWeight.BOLD, color="#7A4FA3")
            gm_value_text = ft.Text("", size=18, weight=ft.FontWeight.BOLD, color="#B47A00")

            def commercial_values():
                sell = safe_float(selling_price.value)
                cost = safe_float(cost_price.value)
                disc = safe_float(discount.value)
                vat = safe_float(vat_rate.value)
                net = max(0, sell - disc)
                vat_amount = net * vat / 100
                total = net + vat_amount
                gm_value = net - cost
                gm_percent = (gm_value / net * 100) if net else 0
                return sell, cost, disc, vat, net, vat_amount, total, gm_percent, gm_value

            def refresh_totals(_=None):
                try:
                    _, _, _, _, net, vat_amount, total, gm_percent, gm_value = commercial_values()
                    net_text.value = money(net)
                    vat_text.value = money(vat_amount)
                    total_text.value = money(total)
                    gm_percent_text.value = f"{gm_percent:.2f}%"
                    gm_value_text.value = money(gm_value)
                    for control in [net_text, vat_text, total_text, gm_percent_text, gm_value_text]:
                        control.update()
                except Exception:
                    pass

            for control in [selling_price, cost_price, discount, vat_rate]:
                control.on_change = refresh_totals
            refresh_totals()

            docs_column = ft.Column(spacing=8)

            def refresh_documents():
                docs = self.db.quotation_documents(quotation_id)
                docs_column.controls.clear()
                if not docs:
                    docs_column.controls.append(
                        ft.Text("No documents uploaded yet.", color="#71827E")
                    )
                else:
                    for doc in docs:
                        docs_column.controls.append(
                            ft.Container(
                                padding=10,
                                bgcolor="#F5F8F7",
                                border_radius=10,
                                content=ft.Row(
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    controls=[
                                        ft.Column(
                                            spacing=2,
                                            controls=[
                                                ft.Text(
                                                    f"{doc['document_type']} | {doc['original_file_name']}",
                                                    weight=ft.FontWeight.BOLD,
                                                    color="#103F37",
                                                ),
                                                ft.Text(
                                                    f"Uploaded by {doc['uploaded_by']} on {doc['uploaded_at']}",
                                                    size=10,
                                                    color="#667B80",
                                                ),
                                            ],
                                        ),
                                        ft.OutlinedButton(
                                            "Open",
                                            icon=ft.Icons.FOLDER_OPEN,
                                            on_click=lambda _, path=doc["stored_file_path"]: os.startfile(path),
                                        ),
                                    ],
                                ),
                            )
                        )
                try:
                    docs_column.update()
                except Exception:
                    pass

            upload_type = ft.Dropdown(
                label="Document Type",
                width=260,
                value="Technical Proposal",
                options=[
                    ft.dropdown.Option("Technical Proposal"),
                    ft.dropdown.Option("Commercial Proposal"),
                    ft.dropdown.Option("Cost Sheet"),
                    ft.dropdown.Option("Vendor Quotation"),
                    ft.dropdown.Option("Customer RFQ"),
                    ft.dropdown.Option("Supporting Document"),
                    ft.dropdown.Option("Other"),
                ],
            )

            file_picker = ft.FilePicker()

            def file_selected(event):
                if not event.files:
                    return
                try:
                    quotation_folder = APP_DIR / "documents" / record["quotation_number"]
                    quotation_folder.mkdir(parents=True, exist_ok=True)
                    for selected in event.files:
                        source_path = Path(selected.path)
                        destination = quotation_folder / source_path.name
                        shutil.copy2(source_path, destination)
                        self.db.add_quotation_document(
                            quotation_id,
                            upload_type.value or "Other",
                            source_path.name,
                            str(destination),
                            current_owner,
                        )
                    self.notify("Document uploaded.")
                    refresh_documents()
                except Exception as error:
                    self.notify(f"Document upload failed: {error}", True)

            file_picker.on_result = file_selected
            self.page.overlay.append(file_picker)

            def select_documents(_):
                file_picker.pick_files(
                    allow_multiple=True,
                    dialog_title="Select quotation documents",
                )

            def info_box(label, value):
                return ft.Container(
                    expand=True,
                    padding=12,
                    bgcolor="#F5F8F7",
                    border_radius=10,
                    content=ft.Column(
                        spacing=3,
                        controls=[
                            ft.Text(label, size=11, color="#667B80"),
                            ft.Text(
                                "" if value is None else str(value),
                                size=14,
                                weight=ft.FontWeight.BOLD,
                                color="#103F37",
                                selectable=True,
                            ),
                        ],
                    ),
                )

            def save_values():
                sell, cost, disc, vat, _, _, _, gm_percent, gm_value = commercial_values()
                self.db.update_record(
                    "quotations",
                    quotation_id,
                    {
                        "owner": assigned_owner.value or "",
                        "concern_owner": concern_owner.value or "",
                        "status": workflow_status.value or "ASSIGNED",
                        "valid_until": target_date.value,
                        "base_value": sell,
                        "cost_price": cost,
                        "discount": disc,
                        "vat_rate": vat,
                        "gross_margin_percent": gm_percent,
                        "gross_margin_value": gm_value,
                        "revision": safe_int(revision.value),
                        "assignment_notes": assignment_notes.value or "",
                        "notes": commercial_notes.value or "",
                        "approval_comments": approval_comments.value or "",
                    },
                )

            def save_workspace(_):
                try:
                    save_values()
                    self.notify(f"Quotation {record['quotation_number']} saved.")
                    self.navigate("quotations")
                except Exception as error:
                    self.notify(f"Could not save quotation: {error}", True)

            def ready_for_submission(_):
                try:
                    save_values()
                    fresh = self.db.one(
                        """SELECT q.*,o.sales_owner
                           FROM quotations q
                           LEFT JOIN opportunities o ON o.id=q.opportunity_id
                           WHERE q.id=?""",
                        (quotation_id,),
                    )
                    self.db.update_record(
                        "quotations",
                        quotation_id,
                        {"status": "READY FOR SUBMISSION"},
                    )
                    self.send_quotation_to_concern_owner(quotation_id, fresh)
                    self.notify("Quotation sent back to the concern owner.")
                    self.navigate("quotations")
                except Exception as error:
                    self.notify(f"Could not send quotation: {error}", True)

            def send_for_approval(_):
                try:
                    save_values()
                    fresh = self.db.one(
                        "SELECT * FROM quotations WHERE id=?",
                        (quotation_id,),
                    )
                    self.submit_for_proposal_manager_approval(quotation_id, fresh)
                    self.notify("Sent to Proposal Manager for approval.")
                    self.navigate("quotations")
                except Exception as error:
                    self.notify(f"Could not start approval: {error}", True)

            def proposal_manager_approve_click(_):
                try:
                    fresh = self.db.one(
                        "SELECT * FROM quotations WHERE id=?",
                        (quotation_id,),
                    )
                    self.proposal_manager_approve(
                        quotation_id,
                        fresh,
                        approval_comments.value or "",
                    )
                    self.notify("Approved and sent to GM.")
                    self.navigate("quotations")
                except Exception as error:
                    self.notify(f"Approval failed: {error}", True)

            def gm_approve_click(_):
                try:
                    fresh = self.db.one(
                        """SELECT q.*,o.sales_owner
                           FROM quotations q
                           LEFT JOIN opportunities o ON o.id=q.opportunity_id
                           WHERE q.id=?""",
                        (quotation_id,),
                    )
                    self.gm_approve(
                        quotation_id,
                        fresh,
                        approval_comments.value or "",
                    )
                    self.notify("Final approval completed.")
                    self.navigate("quotations")
                except Exception as error:
                    self.notify(f"Final approval failed: {error}", True)

            is_proposal_manager = (
                current_owner == self.owner_by_role("Proposal Manager")
            )
            is_gm = (
                current_owner == self.owner_by_role("General Manager")
            )

            action_buttons = [
                ft.OutlinedButton(
                    "Back to Quotations",
                    icon=ft.Icons.ARROW_BACK,
                    on_click=lambda _: self.navigate("quotations"),
                ),
                ft.FilledButton(
                    "Export PDF",
                    icon=ft.Icons.PICTURE_AS_PDF,
                    on_click=lambda _: self.export_quotation_pdf(quotation_id),
                    style=ft.ButtonStyle(bgcolor="#1667A8"),
                ),
                ft.FilledButton(
                    "Save",
                    icon=ft.Icons.SAVE,
                    on_click=save_workspace,
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
                ft.FilledButton(
                    "Ready for Submission",
                    icon=ft.Icons.SEND,
                    on_click=ready_for_submission,
                    style=ft.ButtonStyle(bgcolor="#1667A8"),
                ),
                ft.FilledButton(
                    "Send for Approval",
                    icon=ft.Icons.APPROVAL,
                    on_click=send_for_approval,
                    style=ft.ButtonStyle(bgcolor="#7A4FA3"),
                ),
            ]
            if is_proposal_manager and record["status"] == "PENDING PROPOSAL MANAGER APPROVAL":
                action_buttons.append(
                    ft.FilledButton(
                        "Proposal Manager Approve",
                        icon=ft.Icons.CHECK_CIRCLE,
                        on_click=proposal_manager_approve_click,
                        style=ft.ButtonStyle(bgcolor="#B47A00"),
                    )
                )
            if is_gm and record["status"] == "PENDING GM APPROVAL":
                action_buttons.append(
                    ft.FilledButton(
                        "GM Final Approve",
                        icon=ft.Icons.VERIFIED,
                        on_click=gm_approve_click,
                        style=ft.ButtonStyle(bgcolor="#2E8B57"),
                    )
                )

            refresh_documents()

            workspace = ft.Column(
                expand=True,
                scroll=ft.ScrollMode.AUTO,
                controls=[
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                controls=[
                                    ft.Text(
                                        f"Quotation Workspace | {record['quotation_number']}",
                                        size=28,
                                        weight=ft.FontWeight.BOLD,
                                        color="#103F37",
                                    ),
                                    ft.Text(
                                        "Proposal preparation, commercial calculation, documents, inbox and approvals",
                                        color="#667B80",
                                    ),
                                ],
                            ),
                            ft.Row(wrap=True, controls=action_buttons),
                        ],
                    ),
                    ft.Container(
                        padding=15,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column([
                            ft.Row([
                                info_box("Opportunity ID", record["opportunity_reference"]),
                                info_box("Account", record["account_name"]),
                                info_box("Opportunity Value", money(record["opportunity_value"] or 0)),
                                info_box("Probability", f'{float(record["probability"] or 0):.0f}%'),
                            ]),
                            ft.Divider(),
                            ft.Text(
                                record["opportunity_name"] or "",
                                size=20,
                                weight=ft.FontWeight.BOLD,
                                color="#103F37",
                            ),
                            ft.Text(
                                f"Type: {record['project_type'] or ''} | "
                                f"Stage: {record['opportunity_stage'] or ''} | "
                                f"Expected Close: {record['expected_close_date'] or ''}",
                                color="#667B80",
                            ),
                        ]),
                    ),
                    ft.Container(
                        padding=15,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column([
                            ft.Text("Assignment & Workflow", size=19, weight=ft.FontWeight.BOLD, color="#103F37"),
                            ft.Row(wrap=True, controls=[
                                assigned_owner,
                                concern_owner,
                                workflow_status,
                                target_date_row,
                            ]),
                            assignment_notes,
                        ]),
                    ),
                    ft.Container(
                        padding=15,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column([
                            ft.Text("Commercial Calculation", size=19, weight=ft.FontWeight.BOLD, color="#103F37"),
                            ft.Row(wrap=True, controls=[
                                selling_price, cost_price, discount, vat_rate, revision
                            ]),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_AROUND,
                                controls=[
                                    ft.Column([ft.Text("Net Before VAT", color="#667B80"), net_text]),
                                    ft.Column([ft.Text("VAT Amount", color="#667B80"), vat_text]),
                                    ft.Column([ft.Text("Total Including VAT", color="#667B80"), total_text]),
                                    ft.Column([ft.Text("Gross Margin %", color="#667B80"), gm_percent_text]),
                                    ft.Column([ft.Text("Gross Margin Value", color="#667B80"), gm_value_text]),
                                ],
                            ),
                            commercial_notes,
                        ]),
                    ),
                    ft.Container(
                        padding=15,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column([
                            ft.Text("Documents", size=19, weight=ft.FontWeight.BOLD, color="#103F37"),
                            ft.Row(wrap=True, controls=[
                                upload_type,
                                ft.FilledButton(
                                    "Upload Documents",
                                    icon=ft.Icons.UPLOAD_FILE,
                                    on_click=select_documents,
                                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                                ),
                            ]),
                            docs_column,
                        ]),
                    ),
                    ft.Container(
                        padding=15,
                        bgcolor=ft.Colors.WHITE,
                        border_radius=16,
                        content=ft.Column([
                            ft.Text("Approval Process", size=19, weight=ft.FontWeight.BOLD, color="#103F37"),
                            ft.Row([
                                info_box(
                                    "Proposal Manager Approval",
                                    f"{record['proposal_manager_approval'] or 'Not Started'} "
                                    f"{record['proposal_manager_approved_by'] or ''}"
                                ),
                                info_box(
                                    "GM Approval",
                                    f"{record['gm_approval'] or 'Not Started'} "
                                    f"{record['gm_approved_by'] or ''}"
                                ),
                                info_box("Current Status", record["status"]),
                            ]),
                            approval_comments,
                        ]),
                    ),
                ],
            )

            self.current_view = "quotation_workspace"
            self.content.content = workspace
            self.page.update()

        except Exception as error:
            error_path = APP_DIR / "quotation_workspace_error.log"
            try:
                error_path.write_text(
                    f"{datetime.now().isoformat(timespec='seconds')}\n"
                    f"{type(error).__name__}: {error}",
                    encoding="utf-8",
                )
            except Exception:
                pass
            self.notify(
                f"Could not open quotation workspace: {error}. "
                "Details saved in quotation_workspace_error.log",
                True,
            )

    def quotations_view(self):
        all_records = self.db.quotations()

        columns = [
            ft.DataColumn(ft.Text("New", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Container(width=110, content=ft.Text("Quotation ID", weight=ft.FontWeight.BOLD))),
            ft.DataColumn(ft.Text("Opportunity ID", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Opportunity", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Account", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Request Date", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Target Submission", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Assigned Owner", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Workflow Status", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Age (Days)", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Base Value", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Net Before VAT", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Actions", weight=ft.FontWeight.BOLD)),
        ]
        table = ft.DataTable(columns=columns, rows=[], column_spacing=16, heading_row_color="#E5F0EC")

        def build_rows(records):
            output = []
            for record in records:
                output.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text("NEW" if record["status"] in ("NEW REQUEST","ASSIGNED") else "")),
                    ft.DataCell(self.id_badge(
                        record["quotation_number"],
                        width=110,
                        clickable=lambda _, qid=record["id"]: self.quotation_workspace(qid),
                        tooltip="Open Quotation Workspace",
                    )),
                    ft.DataCell(ft.Text(record["opportunity_reference"] or "", font_family="Consolas")),
                    ft.DataCell(ft.Text(record["opportunity_name"] or "")),
                    ft.DataCell(ft.Text(record["account_name"] or "")),
                    ft.DataCell(ft.Text(record["quotation_date"] or "")),
                    ft.DataCell(ft.Text(record["valid_until"] or "")),
                    ft.DataCell(ft.Text(record["owner"] or "")),
                    ft.DataCell(ft.Text(record["status"] or "")),
                    ft.DataCell(ft.Text(str(self.quotation_age_days(record)))),
                    ft.DataCell(ft.Text(money(record["base_value"]))),
                    ft.DataCell(ft.Text(money(
                        float(record["base_value"] or 0) - float(record["discount"] or 0)
                    ))),
                    ft.DataCell(ft.Row(spacing=2, controls=[
                        ft.FilledButton(
                            "Open",
                            icon=ft.Icons.OPEN_IN_NEW,
                            on_click=lambda _, qid=record["id"]: self.quotation_workspace(qid),
                            style=ft.ButtonStyle(bgcolor="#0E5A47"),
                        ),
                        ft.IconButton(
                            icon=ft.Icons.EDIT,
                            icon_color="#1667A8",
                            on_click=lambda _, rid=record["id"]: self.edit_record_dialog("quotations", rid),
                        ),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_OUTLINE,
                            icon_color="#C0392B",
                            on_click=lambda _, rid=record["id"]: self.confirm_delete_record("quotations", rid),
                        ),
                    ])),
                ]))
            return output

        table.rows = build_rows(all_records)

        def search_changed(event):
            query = (event.control.value or "").strip().lower()
            filtered = [
                row for row in all_records
                if not query or query in " ".join([
                    str(row["quotation_number"] or ""),
                    str(row["opportunity_reference"] or ""),
                    str(row["opportunity_name"] or ""),
                    str(row["account_name"] or ""),
                    str(row["owner"] or ""),
                    str(row["status"] or ""),
                ]).lower()
            ]
            table.rows = build_rows(filtered)
            table.update()

        return ft.Column([
            ft.Row([
                self.title("Quotations", "Search and manage quotation requests, assignments, and approvals."),
                ft.FilledButton(
                    "New Quotation",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.quotation_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.search_with_clear("Search Quotations", search_changed, 440),
            ft.Container(
                bgcolor=ft.Colors.WHITE,
                padding=14,
                border_radius=18,
                content=(self.horizontal_table(table) if all_records else ft.Text("No quotations created yet.")),
            ),
        ])

    def pocs_view(self):
        rows = self.db.pocs()
        return ft.Column([
            ft.Row([
                self.title("Proof of Concepts", "Pilot execution, success criteria, cost, commercial potential and outcomes"),
                ft.FilledButton(
                    "New PoC",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.poc_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.crud_table_container(
                ["PoC ID", "Account", "Opportunity", "PoC Title", "Solution", "Start", "End", "Status", "Estimated Cost", "Commercial Value", "Owner", "Next Step"],
                rows,
                lambda r: [
                    r["poc_reference"], r["account_name"], r["opportunity_reference"],
                    r["poc_title"], r["solution"], r["start_date"], r["planned_end_date"],
                    r["status"], money(r["estimated_cost"]), money(r["commercial_value"]),
                    r["owner"], r["next_step"],
                ],
                "pocs",
                "No PoCs created yet.",
            ),
        ])

    def inbox_view(self):
        owner_name = self.current_owner_name()
        rows = self.db.inbox_for_owner(owner_name)

        def action_button(row):
            if row["item_type"] == "QUOTATION":
                return ft.FilledButton(
                    "Open",
                    icon=ft.Icons.OPEN_IN_NEW,
                    on_click=lambda _, qid=row["related_record_id"]: self.quotation_workspace(qid),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                )
            return ft.Container()

        table_rows = []
        for row in rows:
            table_rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(row["created_at"])),
                        ft.DataCell(ft.Text(row["reference_id"], weight=ft.FontWeight.BOLD)),
                        ft.DataCell(ft.Text(row["title"])),
                        ft.DataCell(ft.Text(row["message"] or "")),
                        ft.DataCell(ft.Text(row["action_type"] or "")),
                        ft.DataCell(action_button(row)),
                    ]
                )
            )

        return ft.Column([
            self.title(
                "Inbox",
                f"Open actions assigned to {owner_name or 'the current user'}"
            ),
            ft.Container(
                bgcolor=ft.Colors.WHITE,
                padding=14,
                border_radius=18,
                content=(
                    ft.DataTable(
                        columns=[
                            ft.DataColumn(ft.Text("Date")),
                            ft.DataColumn(ft.Text("Reference")),
                            ft.DataColumn(ft.Text("Title")),
                            ft.DataColumn(ft.Text("Message")),
                            ft.DataColumn(ft.Text("Action")),
                            ft.DataColumn(ft.Text("Open")),
                        ],
                        rows=table_rows,
                        heading_row_color="#E5F0EC",
                    )
                    if table_rows
                    else ft.Text("Your inbox is clear.", color="#2E8B57")
                ),
            ),
        ])

    def notifications_view(self):
        owner_name = self.current_owner_name()
        rows = self.db.notifications_for_owner(owner_name)

        def mark_read(notification_id):
            self.db.mark_notification_read(notification_id)
            self.navigate("notifications")

        table_rows = []
        for row in rows:
            table_rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text("Unread" if not row["is_read"] else "Read")),
                        ft.DataCell(ft.Text(row["created_at"])),
                        ft.DataCell(ft.Text(row["title"], weight=ft.FontWeight.BOLD)),
                        ft.DataCell(ft.Text(row["message"] or "")),
                        ft.DataCell(
                            ft.TextButton(
                                "Mark Read",
                                on_click=lambda _, nid=row["id"]: mark_read(nid),
                                disabled=bool(row["is_read"]),
                            )
                        ),
                    ]
                )
            )

        return ft.Column([
            self.title("Notifications", f"Notifications for {owner_name or 'current user'}"),
            ft.Container(
                bgcolor=ft.Colors.WHITE,
                padding=14,
                border_radius=18,
                content=(
                    ft.DataTable(
                        columns=[
                            ft.DataColumn(ft.Text("Status")),
                            ft.DataColumn(ft.Text("Date")),
                            ft.DataColumn(ft.Text("Title")),
                            ft.DataColumn(ft.Text("Message")),
                            ft.DataColumn(ft.Text("Action")),
                        ],
                        rows=table_rows,
                        heading_row_color="#E5F0EC",
                    )
                    if table_rows
                    else ft.Text("No notifications.", color="#71827E")
                ),
            ),
        ])

    def owners_view(self):
        all_records = self.db.owners()
        holder = ft.Column()

        def render(records):
            holder.controls = [
                self.crud_table_container(
                    ["Owner ID","Owner Name","Role","Department","Email","Mobile","Eligible for Quotation Assignment","Status"],
                    records,
                    lambda r: [
                        r["owner_code"],r["full_name"],r["role"],r["department"],
                        r["email"],r["mobile"],
                        "Yes" if r["can_receive_quotation"] else "No",
                        r["status"],
                    ],
                    "owners",
                    "No owners created yet.",
                )
            ]
            try:
                holder.update()
            except Exception:
                pass

        render(all_records)

        def search_changed(event):
            query = (event.control.value or "").strip().lower()
            filtered = [
                row for row in all_records
                if not query or query in " ".join([
                    str(row["owner_code"] or ""),
                    str(row["full_name"] or ""),
                    str(row["role"] or ""),
                    str(row["department"] or ""),
                    str(row["email"] or ""),
                    str(row["status"] or ""),
                ]).lower()
            ]
            render(filtered)

        return ft.Column([
            ft.Row([
                self.title("Owners", "Central owner directory used for CRM assignments."),
                ft.FilledButton(
                    "New Owner",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.owner_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.search_with_clear("Search Owners", search_changed),
            holder,
        ])

    def owner_dialog(self):
        owner_code = self.labelled_readonly_field(
            "Auto Owner ID",
            self.db.peek_reference("owners"),
            width=320,
        )
        full_name = self.text("Owner Name", 420)
        role = self.dropdown("Role", OWNER_ROLES, "Proposal Engineer", 280)
        department = self.text("Department", 300)
        email = self.text("Email", 360)
        mobile = self.text("Mobile", 240)
        quotation_owner = ft.Checkbox(
            label="Can receive quotation assignments",
            value=False,
        )
        status = self.dropdown("Status", OWNER_STATUSES, "Active", 220)
        notes = self.text("Notes", 620, multiline=True)

        # The reference was reserved above. Avoid skipping it by inserting explicitly.
        reserved_code = owner_code.controls[1].value

        def save(_):
            try:
                if not full_name.value.strip():
                    raise ValueError("Owner Name is required.")
                self.db.insert_owner({
                    "owner_code": reserved_code,
                    "full_name": full_name.value.strip(),
                    "role": role.value,
                    "department": department.value,
                    "email": email.value,
                    "mobile": mobile.value,
                    "can_receive_quotation": quotation_owner.value,
                    "status": status.value,
                    "notes": notes.value,
                })
                self.page.close(dialog)
                self.notify("Owner created.")
                self.navigate("owners")
            except Exception as error:
                self.notify(f"Could not create owner: {error}", True)

        dialog = self.dialog(
            "New Owner",
            [
                owner_code, full_name, role, department, email, mobile,
                quotation_owner, status, notes
            ],
            save,
        )
        self.page.open(dialog)

    def meetings_view(self):
        rows = self.db.meetings()
        return ft.Column([
            ft.Row([
                self.title("Meetings", "Account-level and opportunity-level meetings with clear follow-up ownership"),
                ft.FilledButton(
                    "New Meeting",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.meeting_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.crud_table_container(
                ["Meeting ID", "Date", "Account", "Opportunity", "Type", "Subject", "Location", "Owner", "Outcome", "Next Action", "Next Action Date"],
                rows,
                lambda r: [
                    r["meeting_reference"], r["meeting_date"], r["account_name"],
                    r["opportunity_reference"], r["meeting_type"], r["subject"],
                    r["location"], r["owner"], r["outcome"], r["next_action"],
                    r["next_action_date"],
                ],
                "meetings",
                "No meetings recorded yet.",
            ),
        ])

    def activities_view(self):
        rows = self.db.activities()
        today = date.today()

        def values(r):
            display_status = r["status"]
            try:
                if (
                    r["due_date"]
                    and datetime.strptime(r["due_date"], "%Y-%m-%d").date() < today
                    and r["status"] not in ("Completed", "Cancelled")
                ):
                    display_status = "OVERDUE"
            except ValueError:
                pass
            return [
                r["activity_reference"], r["due_date"], r["account_name"],
                r["opportunity_reference"], r["activity_type"], r["subject"],
                r["priority"], display_status, r["owner"], r["completed_date"],
            ]

        return ft.Column([
            ft.Row([
                self.title("Activities & Follow-ups", "Calls, emails, site visits, proposal follow-ups and accountable next actions"),
                ft.FilledButton(
                    "New Activity",
                    icon=ft.Icons.ADD,
                    on_click=lambda _: self.activity_dialog(),
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            self.crud_table_container(
                ["Activity ID", "Due Date", "Account", "Opportunity", "Type", "Subject", "Priority", "Status", "Owner", "Completed Date"],
                rows,
                values,
                "activities",
                "No activities created yet.",
            ),
        ])

    def pipeline_view(self):
        opportunities = self.db.opportunities()
        stage_values = defaultdict(float)
        stage_counts = Counter()
        owner_values = defaultdict(float)
        forecast_months = defaultdict(float)
        for o in opportunities:
            if o["status"] != "Open":
                continue
            stage_values[o["stage"]] += float(o["estimated_value"])
            stage_counts[o["stage"]] += 1
            owner_values[o["sales_owner"] or "Unassigned"] += float(o["estimated_value"])
            if o["expected_close_date"]:
                forecast_months[o["expected_close_date"][:7]] += float(o["estimated_value"]) * float(o["probability"]) / 100

        stage_rows = [[stage, stage_counts[stage], money(stage_values[stage]),
                       f'{STAGE_PROBABILITY.get(stage,0)}%',
                       money(stage_values[stage]*STAGE_PROBABILITY.get(stage,0)/100)]
                      for stage in OPPORTUNITY_STAGES if stage_counts[stage]]

        owner_rows = sorted([[owner, money(value)] for owner,value in owner_values.items()],
                            key=lambda x: safe_float(x[1].replace("SAR","")), reverse=True)

        forecast_rows = [[month, money(value)] for month,value in sorted(forecast_months.items())]

        return ft.Column([
            self.title("Pipeline Analytics", "Stage funnel, owner concentration and weighted closing forecast"),
            ft.Row([
                ft.Container(expand=True,bgcolor=ft.Colors.WHITE,padding=18,border_radius=18,
                             content=ft.Column([ft.Text("Pipeline by Stage",size=19,weight=ft.FontWeight.BOLD),
                                                self.data_table(["Stage","Count","Value","Default Probability","Weighted Value"],stage_rows)])),
                ft.Container(expand=True,bgcolor=ft.Colors.WHITE,padding=18,border_radius=18,
                             content=ft.Column([ft.Text("Pipeline by Sales Owner",size=19,weight=ft.FontWeight.BOLD),
                                                self.data_table(["Owner","Pipeline Value"],owner_rows)])),
            ]),
            ft.Container(bgcolor=ft.Colors.WHITE,padding=18,border_radius=18,
                         content=ft.Column([ft.Text("Weighted Forecast by Closing Month",size=19,weight=ft.FontWeight.BOLD),
                                            self.data_table(["Month","Weighted Forecast"],forecast_rows)]))
        ])

    def reports_view(self):
        return ft.Column([
            self.title("Reports", "Executive reporting for management reviews and sales governance"),
            ft.Row([
                ft.FilledButton("Export Executive PDF", icon=ft.Icons.PICTURE_AS_PDF,
                                on_click=lambda _: self.export_pdf(), style=ft.ButtonStyle(bgcolor="#B52A24")),
                ft.FilledButton("Export Full CRM Excel", icon=ft.Icons.TABLE_VIEW,
                                on_click=lambda _: self.export_excel(), style=ft.ButtonStyle(bgcolor="#1F6F43")),
                ft.OutlinedButton("Open Reports Folder", icon=ft.Icons.FOLDER_OPEN,
                                  on_click=lambda _: os.startfile(EXPORT_DIR)),
            ]),
            ft.Container(
                bgcolor=ft.Colors.WHITE,padding=22,border_radius=18,
                content=ft.Column([
                    ft.Text("Included Insights",size=20,weight=ft.FontWeight.BOLD),
                    ft.Text("• Pipeline value and weighted pipeline"),
                    ft.Text("• Win rate and lead conversion"),
                    ft.Text("• Opportunity stage funnel"),
                    ft.Text("• Monthly weighted sales forecast"),
                    ft.Text("• Top accounts and sales owners"),
                    ft.Text("• Quotations, PoCs, meetings and activities"),
                    ft.Text("• Overdue follow-ups and immediate actions"),
                ])
            )
        ])

    def migration_view(self):
        files = sorted(p.name for p in MIGRATION_DIR.glob("*.csv"))
        return ft.Column([
            self.title("Data Migration", "Import your current CRM information using the prepared CSV templates"),
            ft.Row([
                ft.FilledButton("Import Migration Files", icon=ft.Icons.UPLOAD_FILE,
                                on_click=lambda _: self.import_migration(), style=ft.ButtonStyle(bgcolor="#0E5A47")),
                ft.OutlinedButton("Open Migration Folder", icon=ft.Icons.FOLDER_OPEN,
                                  on_click=lambda _: os.startfile(MIGRATION_DIR)),
            ]),
            ft.Container(
                bgcolor=ft.Colors.WHITE,padding=22,border_radius=18,
                content=ft.Column([
                    ft.Text("Migration Process",size=20,weight=ft.FontWeight.BOLD),
                    ft.Text("1. Open the migration folder."),
                    ft.Text("2. Replace the sample rows with your current data."),
                    ft.Text("3. Keep the headers and file names unchanged."),
                    ft.Text("4. Save each file as UTF-8 CSV."),
                    ft.Text("5. Click Import Migration Files."),
                    ft.Divider(),
                    ft.Text("Available Files",size=18,weight=ft.FontWeight.BOLD),
                    *[ft.Text(f"• {name}") for name in files],
                    ft.Divider(),
                    ft.Text("Import Order",size=18,weight=ft.FontWeight.BOLD),
                    ft.Text("Accounts are imported first, followed by Contacts, Leads, Opportunities, Quotations, PoCs, Meetings and Activities."),
                ])
            )
        ])

    TABLE_VIEW_MAP = {
        "accounts": "accounts",
        "contacts": "contacts",
        "leads": "leads",
        "opportunities": "opportunities",
        "quotations": "quotations",
        "pocs": "pocs",
        "meetings": "meetings",
        "activities": "activities",
        "owners": "owners",
    }

    FIELD_LABELS = {
        "account_id": "Account",
        "opportunity_id": "Related Opportunity",
        "lead_reference": "Lead ID",
        "opportunity_reference": "Opportunity ID",
        "quotation_number": "Quotation ID",
        "poc_reference": "PoC ID",
        "meeting_reference": "Meeting ID",
        "activity_reference": "Activity ID",
        "account_reference": "Account ID",
        "contact_reference": "Contact ID",
        "owner_code": "Owner ID",
        "full_name": "Owner Name",
        "can_receive_quotation": "Quotation Assignment Eligible",
        "assigned_date": "Assigned Date",
        "work_started_date": "Work Started Date",
        "submitted_date": "Submitted Date",
        "completed_date": "Completed Date",
        "assignment_notes": "Assignment Notes",
        "first_name": "First Name",
        "last_name": "Last Name",
        "job_title": "Job Title",
        "main_phone": "Main Phone",
        "lead_status": "Lead Status",
        "lead_score": "Lead Score",
        "interest_area": "Interest Area",
        "expected_close_date": "Expected Close Date",
        "sales_owner": "Sales Owner",
        "technical_owner": "Technical Owner",
        "next_action_date": "Next Action Date",
        "base_value": "Base Value",
        "valid_until": "Valid Until",
        "vat_rate": "VAT Rate %",
        "planned_end_date": "Planned End Date",
        "success_criteria": "Success Criteria",
        "meeting_date": "Meeting Date",
        "meeting_type": "Meeting Type",
        "activity_type": "Activity Type",
        "completed_date": "Completed Date",
    }

    DROPDOWN_FIELDS = {
        "account_type": ACCOUNT_TYPES,
        "status_accounts": ACCOUNT_STATUSES,
        "industry": INDUSTRIES,
        "lead_status": LEAD_STATUSES,
        "source": LEAD_SOURCES,
        "interest_area": INTEREST_AREAS,
        "project_type": INTEREST_AREAS,
        "stage": OPPORTUNITY_STAGES,
        "status_opportunities": OPPORTUNITY_STATUSES,
        "status_quotations": QUOTATION_STATUSES,
        "status_pocs": POC_STATUSES,
        "meeting_type": MEETING_TYPES,
        "activity_type": ACTIVITY_TYPES,
        "priority": PRIORITIES,
        "status_activities": ACTIVITY_STATUSES,
        "role": OWNER_ROLES,
        "status_owners": OWNER_STATUSES,
        "influence_level": ["Low", "Medium", "High", "Decision Maker"],
        "relationship_status": ["New", "Developing", "Established", "Strategic", "Dormant"],
    }

    def record_display_name(self, table, record):
        keys = {
            "accounts": "account_name",
            "contacts": "first_name",
            "leads": "lead_reference",
            "opportunities": "opportunity_reference",
            "quotations": "quotation_number",
            "pocs": "poc_reference",
            "meetings": "meeting_reference",
            "activities": "activity_reference",
            "owners": "owner_code",
        }
        return str(record.get(keys.get(table, "id"), record.get("id", "")))

    def crud_table_container(self, headers, records, value_builder, table_name, empty):
        if not records:
            return ft.Container(
                bgcolor=ft.Colors.WHITE,
                padding=14,
                border_radius=18,
                content=ft.Text(empty, color="#73868B"),
            )

        columns = [ft.DataColumn(ft.Text(header, weight=ft.FontWeight.BOLD)) for header in headers]
        columns.append(ft.DataColumn(ft.Text("Actions", weight=ft.FontWeight.BOLD)))

        rows = []
        for record in records:
            values = value_builder(record)
            action_cell = ft.Row(
                spacing=2,
                controls=[
                    ft.IconButton(
                        icon=ft.Icons.EDIT,
                        tooltip="Edit",
                        icon_color="#1667A8",
                        on_click=lambda _, table=table_name, rid=record["id"]: self.edit_record_dialog(table, rid),
                    ),
                    ft.IconButton(
                        icon=ft.Icons.DELETE_OUTLINE,
                        tooltip="Delete",
                        icon_color="#C0392B",
                        on_click=lambda _, table=table_name, rid=record["id"]: self.confirm_delete_record(table, rid),
                    ),
                ],
            )
            cells = [
                ft.DataCell(
                    ft.Text(
                        "" if value is None else str(value),
                        selectable=True,
                    )
                )
                for value in values
            ]
            cells.append(ft.DataCell(action_cell))
            rows.append(ft.DataRow(cells=cells))

        data_table = ft.DataTable(
            columns=columns,
            rows=rows,
            column_spacing=20,
            heading_row_color="#E5F0EC",
        )
        return ft.Container(
            bgcolor=ft.Colors.WHITE,
            padding=14,
            border_radius=18,
            content=self.horizontal_table(data_table),
        )

    def confirm_delete_record(self, table, record_id):
        record = self.db.raw_record(table, record_id)
        if not record:
            self.notify("The record no longer exists.", True)
            return

        display_name = self.record_display_name(table, record)

        def delete(_):
            try:
                self.db.delete_record(table, record_id)
                self.page.close(dialog)
                self.notify(f"{display_name} deleted.")
                self.navigate(self.TABLE_VIEW_MAP[table])
            except Exception as error:
                self.notify(f"Could not delete the record: {error}", True)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Confirm Deletion", weight=ft.FontWeight.BOLD),
            content=ft.Text(
                f"Delete {display_name}?\n\n"
                "Related records may also be affected. This action cannot be undone."
            ),
            actions=[
                ft.TextButton("Cancel", on_click=lambda _: self.page.close(dialog)),
                ft.FilledButton(
                    "Delete",
                    icon=ft.Icons.DELETE,
                    on_click=delete,
                    style=ft.ButtonStyle(bgcolor="#C0392B"),
                ),
            ],
        )
        self.page.open(dialog)

    def edit_record_dialog(self, table, record_id):
        record = self.db.raw_record(table, record_id)
        if not record:
            self.notify("The record no longer exists.", True)
            return

        columns = self.db.table_columns(table)
        controls = {}
        layout_controls = []

        long_text_fields = {
            "notes", "outcome", "next_step", "success_criteria", "attendees",
            "competitors", "lost_reason", "mitigation", "assignment_notes"
        }
        numeric_fields = {
            "lead_score", "estimated_value", "probability", "customer_budget",
            "base_value", "discount", "vat_rate", "revision", "estimated_cost",
            "commercial_value"
        }
        reference_fields = {
            "account_reference", "contact_reference", "owner_code",
            "lead_reference", "opportunity_reference", "quotation_number",
            "poc_reference", "meeting_reference", "activity_reference"
        }

        for column in columns:
            field = column["name"]
            if field in {"id", "created_at"}:
                continue
            if table == "quotations" and field == "account_id":
                continue

            label = self.FIELD_LABELS.get(
                field,
                field.replace("_", " ").title(),
            )
            value = record.get(field)

            if table == "quotations" and field == "owner":
                options = self.owner_options(quotation_only=False)
                control = ft.Dropdown(
                    label="Eligible for Quotation Assignment",
                    width=650,
                    value=str(value) if value else None,
                    options=options,
                )
            elif table == "owners" and field == "can_receive_quotation":
                control = ft.Dropdown(
                    label=label,
                    width=420,
                    value="Yes" if value else "No",
                    options=[ft.dropdown.Option("Yes"), ft.dropdown.Option("No")],
                )
            elif field == "account_id":
                options = self.account_options()
                control = ft.Dropdown(
                    label=label,
                    width=650,
                    value=str(value) if value else None,
                    options=options,
                )
            elif field == "opportunity_id":
                options = self.opportunity_options()
                control = ft.Dropdown(
                    label=label,
                    width=650,
                    value=str(value) if value else None,
                    options=options,
                )
            else:
                dropdown_key = field
                if field == "status":
                    dropdown_key = f"status_{table}"
                dropdown_values = self.DROPDOWN_FIELDS.get(dropdown_key)

                if dropdown_values:
                    control = ft.Dropdown(
                        label=label,
                        width=420,
                        value=str(value) if value not in (None, "") else dropdown_values[0],
                        options=[ft.dropdown.Option(item) for item in dropdown_values],
                    )
                else:
                    control = ft.TextField(
                        label=label,
                        width=650 if field in long_text_fields else 320,
                        value="" if value is None else str(value),
                        multiline=field in long_text_fields,
                        min_lines=3 if field in long_text_fields else 1,
                        read_only=field in reference_fields,
                        keyboard_type=ft.KeyboardType.NUMBER if field in numeric_fields else ft.KeyboardType.TEXT,
                    )

            controls[field] = control
            layout_controls.append(control)

        def save(_):
            try:
                values = {}
                for field, control in controls.items():
                    raw_value = control.value
                    if table == "owners" and field == "can_receive_quotation":
                        values[field] = 1 if raw_value == "Yes" else 0
                    elif field in numeric_fields:
                        values[field] = safe_float(raw_value)
                    elif field in {"account_id", "opportunity_id"}:
                        values[field] = int(raw_value) if raw_value else None
                    elif field == "revision":
                        values[field] = safe_int(raw_value)
                    else:
                        values[field] = raw_value or ""

                if table == "quotations":
                    opportunity_id = values.get("opportunity_id")
                    if not opportunity_id:
                        raise ValueError("A quotation must be linked to an opportunity.")
                    opportunity = self.db.one(
                        "SELECT account_id FROM opportunities WHERE id=?",
                        (opportunity_id,),
                    )
                    if not opportunity:
                        raise ValueError("The selected opportunity no longer exists.")
                    values["account_id"] = opportunity["account_id"]

                self.db.update_record(table, record_id, values)
                if table == "quotations" and "status" in values:
                    self.update_quotation_workflow_dates(record_id, values["status"])
                self.page.close(dialog)
                self.notify("Record updated successfully.")
                self.navigate(self.TABLE_VIEW_MAP[table])
            except Exception as error:
                self.notify(f"Could not update the record: {error}", True)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text(
                f"Edit {self.record_display_name(table, record)}",
                weight=ft.FontWeight.BOLD,
            ),
            content=ft.Container(
                width=760,
                content=ft.Column(
                    height=620,
                    scroll=ft.ScrollMode.AUTO,
                    controls=layout_controls,
                ),
            ),
            actions=[
                ft.TextButton("Cancel", on_click=lambda _: self.page.close(dialog)),
                ft.FilledButton(
                    "Save Changes",
                    icon=ft.Icons.SAVE,
                    on_click=save,
                    style=ft.ButtonStyle(bgcolor="#0E5A47"),
                ),
            ],
        )
        self.page.open(dialog)

    def id_badge(self, value, width=112, clickable=None, tooltip="Open Details"):
        label = ft.Text(
            "" if value is None else str(value),
            weight=ft.FontWeight.BOLD,
            color="#1667A8" if clickable else "#103F37",
            font_family="Consolas",
            size=13,
            no_wrap=True,
        )
        if clickable:
            return ft.TextButton(
                content=ft.Container(
                    width=width,
                    alignment=ft.Alignment(-1, 0),
                    content=label,
                ),
                tooltip=tooltip,
                on_click=clickable,
                style=ft.ButtonStyle(
                    padding=ft.Padding(left=0, right=0, top=0, bottom=0),
                    overlay_color=ft.Colors.with_opacity(0.06, "#1667A8"),
                ),
            )
        return ft.Container(
            width=width,
            alignment=ft.Alignment(-1, 0),
            content=label,
        )


    def search_field(self, label, on_change, width=380):
        return ft.TextField(
            label=label,
            prefix_icon=ft.Icons.SEARCH,
            width=width,
            dense=True,
            on_change=on_change,
        )

    def search_with_clear(self, label, on_change, width=380):
        field = self.search_field(label, on_change, width)

        def clear_search(_):
            field.value = ""
            field.update()
            on_change(type("SearchEvent", (), {"control": field})())

        return ft.Row(
            spacing=8,
            controls=[
                field,
                ft.OutlinedButton(
                    "Clear",
                    icon=ft.Icons.CLEAR,
                    on_click=clear_search,
                ),
            ],
        )

    def horizontal_table(self, content):
        return ft.Row(
            scroll=ft.ScrollMode.ALWAYS,
            controls=[content],
        )


    def table_container(self, headers, rows, empty):
        return ft.Container(
            bgcolor=ft.Colors.WHITE,padding=14,border_radius=18,
            content=self.data_table(headers,rows) if rows else ft.Text(empty,color="#73868B")
        )

    def data_table(self, headers, rows):
        return ft.DataTable(
            columns=[ft.DataColumn(ft.Text(h,weight=ft.FontWeight.BOLD)) for h in headers],
            rows=[ft.DataRow(cells=[ft.DataCell(ft.Text("" if v is None else str(v))) for v in row]) for row in rows],
            column_spacing=22,
            heading_row_color="#E5F0EC",
        )

    def calendar_field(self, label, initial_value, width=320):
        date_text = ft.TextField(
            label=label,
            width=width,
            value=initial_value,
            read_only=True,
        )

        def date_changed(event):
            if event.control.value:
                date_text.value = event.control.value.date().isoformat()
                date_text.update()

        date_picker = ft.DatePicker(
            value=datetime.strptime(initial_value, "%Y-%m-%d")
            if initial_value else datetime.now(),
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2040, 12, 31),
            on_change=date_changed,
        )

        def open_calendar(_):
            try:
                self.page.open(date_picker)
            except Exception:
                if date_picker not in self.page.overlay:
                    self.page.overlay.append(date_picker)
                date_picker.open = True
                self.page.update()

        return ft.Row(
            spacing=6,
            vertical_alignment=ft.CrossAxisAlignment.END,
            controls=[
                date_text,
                ft.IconButton(
                    icon=ft.Icons.CALENDAR_MONTH,
                    tooltip="Select date",
                    on_click=open_calendar,
                ),
            ],
        ), date_text

    def labelled_readonly_field(self, label, value, width=320):
        return ft.Column(
            spacing=4,
            controls=[
                ft.Text(
                    label,
                    size=13,
                    weight=ft.FontWeight.BOLD,
                    color="#103F37",
                ),
                ft.TextField(
                    value=value,
                    width=width,
                    read_only=True,
                    filled=True,
                    bgcolor="#F5F8F7",
                ),
            ],
        )

    def text(self, label, width=300, value="", multiline=False, read_only=False):
        return ft.TextField(
            label=label,
            width=width,
            value=value,
            multiline=multiline,
            min_lines=3 if multiline else 1,
            read_only=read_only,
        )

    def dropdown(self, label, items, value=None, width=300):
        return ft.Dropdown(label=label,width=width,value=value or (items[0] if items else None),
                           options=[ft.dropdown.Option(x) for x in items])

    def account_dialog(self):
        name=self.text("Account Name",430); atype=self.dropdown("Account Type",ACCOUNT_TYPES,width=260)
        industry=self.dropdown("Industry",INDUSTRIES,width=260); city=self.text("City",220)
        country=self.text("Country",220,"Saudi Arabia"); website=self.text("Website",320)
        phone=self.text("Main Phone",220); owner=self.text("Owner",260)
        status=self.dropdown("Status",ACCOUNT_STATUSES,"Active",220); notes=self.text("Notes",650,multiline=True)
        def save(_):
            try:
                if not name.value.strip(): raise ValueError("Account Name is required.")
                self.db.insert_account({"account_name":name.value.strip(),"account_type":atype.value,"industry":industry.value,
                    "city":city.value,"country":country.value,"website":website.value,"main_phone":phone.value,
                    "owner":owner.value,"status":status.value,"notes":notes.value})
                self.page.close(d); self.notify("Account created."); self.navigate("accounts")
            except Exception as e: self.notify(f"Could not create account: {e}",True)
        d=self.dialog("New Account",[ft.Row([name,atype]),ft.Row([industry,city,country]),ft.Row([website,phone]),
                                    ft.Row([owner,status]),notes],save); self.page.open(d)

    def contact_dialog(self):
        contact_reference = self.labelled_readonly_field(
            "Auto Contact ID",
            self.db.peek_reference("contacts"),
            width=320,
        )
        accounts=self.db.accounts()
        if not accounts: self.notify("Create an account first.",True); return
        account=ft.Dropdown(label="Account",width=520,value=str(accounts[0]["id"]),options=self.account_options())
        first=self.text("First Name",240); last=self.text("Last Name",240); title=self.text("Job Title",300)
        dept=self.text("Department",260); email=self.text("Email",320); mobile=self.text("Mobile",220)
        phone=self.text("Phone",220); influence=self.dropdown("Influence Level",["Low","Medium","High","Decision Maker"],"Medium",220)
        relationship=self.dropdown("Relationship",["New","Developing","Established","Strategic","Dormant"],"New",220)
        owner=self.text("Owner",240); notes=self.text("Notes",620,multiline=True)
        def save(_):
            try:
                if not first.value.strip(): raise ValueError("First Name is required.")
                self.db.insert_contact({"account_id":int(account.value),"first_name":first.value.strip(),"last_name":last.value,
                    "job_title":title.value,"department":dept.value,"email":email.value,"mobile":mobile.value,
                    "phone":phone.value,"influence_level":influence.value,"relationship_status":relationship.value,
                    "owner":owner.value,"notes":notes.value})
                self.page.close(d); self.notify("Contact created."); self.navigate("contacts")
            except Exception as e: self.notify(f"Could not create contact: {e}",True)
        d=self.dialog("New Contact",[account,ft.Row([first,last]),ft.Row([title,dept]),ft.Row([email,mobile]),
                                     ft.Row([phone,influence,relationship]),owner,notes],save); self.page.open(d)

    def lead_dialog(self):
        accounts = self.db.accounts()
        if not accounts:
            self.notify("Create an Account before creating a Lead.", True)
            return

        reference = self.text(
            "Auto Lead ID",
            220,
            self.db.peek_reference("leads"),
            read_only=True,
        )
        account = ft.Dropdown(
            label="Account",
            width=420,
            options=[
                ft.dropdown.Option(
                    key=str(row["id"]),
                    text=row["account_name"],
                )
                for row in accounts
            ],
            value=None,
            hint_text="Select an Account",
        )
        company_name = self.text("Company Name", 360)
        contact_name = self.text("Contact Name", 300)
        job_title = self.text("Job Title", 300)
        email = self.text("Email", 300)
        mobile = self.text("Mobile", 250)
        source = self.dropdown("Lead Source", LEAD_SOURCES, "Referral", 250)
        interest = self.dropdown(
            "Interest Area",
            INTEREST_AREAS,
            "Instrumentation",
            320,
        )
        status = self.dropdown("Lead Status", LEAD_STATUSES, "New", 220)
        score = self.text("Lead Score (0-100)", 180, "50")

        creator_name = self.current_owner_name() or (
            self.user["full_name"] if self.user else "System User"
        )
        owner = ft.TextField(
            label="Owner",
            width=320,
            value=creator_name,
            read_only=True,
            filled=True,
            bgcolor="#F5F8F7",
        )

        next_date_control, next_date = self.calendar_field(
            "Next Action Date",
            (date.today() + timedelta(days=7)).isoformat(),
            width=280,
        )
        estimated_value = self.text("Estimated Value", 240, "0")
        notes = self.text("Notes", 650, multiline=True)

        account_lookup = {str(row["id"]): row for row in accounts}

        def account_changed(_):
            selected = account_lookup.get(str(account.value))
            if selected:
                company_name.value = selected["account_name"] or ""
            else:
                company_name.value = ""
            try:
                company_name.update()
            except Exception:
                pass

        account.on_change = account_changed

        def save(_):
            try:
                if not account.value:
                    raise ValueError("Please select an Account.")
                if not company_name.value.strip():
                    raise ValueError("Company Name is required.")
                valid_date(next_date.value)

                self.db.insert_lead({
                    "lead_reference": reference.value.strip(),
                    "account_id": int(account.value),
                    "company_name": company_name.value.strip(),
                    "contact_name": contact_name.value,
                    "job_title": job_title.value,
                    "email": email.value,
                    "mobile": mobile.value,
                    "source": source.value,
                    "interest_area": interest.value,
                    "lead_status": status.value,
                    "lead_score": safe_float(score.value),
                    "owner": creator_name,
                    "next_action_date": next_date.value,
                    "estimated_value": safe_float(estimated_value.value),
                    "notes": notes.value,
                })
                self.page.close(dialog)
                self.notify(
                    f"Lead {reference.value} created and assigned to {creator_name}."
                )
                self.navigate("leads")
            except Exception as error:
                self.notify(f"Could not create lead: {error}", True)

        dialog = self.dialog(
            "New Lead",
            [
                ft.Row([reference, account], wrap=True),
                ft.Row([company_name, contact_name], wrap=True),
                ft.Row([job_title, email, mobile], wrap=True),
                ft.Row([source, interest], wrap=True),
                ft.Row([status, score, owner], wrap=True),
                ft.Row([next_date_control, estimated_value], wrap=True),
                notes,
            ],
            save,
        )
        self.page.open(dialog)

    def opportunity_dialog(self):
        accounts=self.db.accounts()
        if not accounts: self.notify("Create an account first.",True); return
        ref=self.text("Auto Opportunity ID",230,f"OPP-{datetime.now().strftime('%y%m%d%H%M%S')}",read_only=True)
        account=ft.Dropdown(label="Account",width=420,value=str(accounts[0]["id"]),options=self.account_options())
        name=self.text("Opportunity Name",520); ptype=self.dropdown("Project / Solution Type",INTEREST_AREAS,width=330)
        stage=self.dropdown("Stage",OPPORTUNITY_STAGES,"Qualification",260)
        prob=self.text("Probability %",180,str(STAGE_PROBABILITY["Qualification"]))
        value=self.text("Estimated Value",240,"0"); close=self.text("Expected Close Date YYYY-MM-DD",280,(date.today()+timedelta(days=90)).isoformat())
        sales=self.text("Sales Owner",240); technical=self.text("Technical Owner",240); budget=self.text("Customer Budget",220,"0")
        competitors=self.text("Competitors",420); next_step=self.text("Next Step",500)
        next_date=self.text("Next Action Date YYYY-MM-DD",280,(date.today()+timedelta(days=7)).isoformat())
        status=self.dropdown("Status",OPPORTUNITY_STATUSES,"Open",200); lost=self.text("Lost Reason",360)
        notes=self.text("Notes",680,multiline=True)
        def stage_change(_):
            prob.value=str(STAGE_PROBABILITY.get(stage.value,0)); prob.update()
        stage.on_change=stage_change
        def save(_):
            try:
                valid_date(close.value); valid_date(next_date.value)
                self.db.insert_opportunity({"opportunity_reference":ref.value.strip(),"account_id":int(account.value),
                    "opportunity_name":name.value.strip(),"project_type":ptype.value,"stage":stage.value,
                    "probability":safe_float(prob.value),"estimated_value":safe_float(value.value),
                    "expected_close_date":close.value,"sales_owner":sales.value,"technical_owner":technical.value,
                    "customer_budget":safe_float(budget.value),"competitors":competitors.value,"next_step":next_step.value,
                    "next_action_date":next_date.value,"status":status.value,"lost_reason":lost.value,"notes":notes.value})
                self.page.close(d); self.notify("Opportunity created."); self.navigate("opportunities")
            except Exception as e: self.notify(f"Could not create opportunity: {e}",True)
        d=self.dialog("New Opportunity",[ft.Row([ref,account]),name,ft.Row([ptype,stage,prob]),ft.Row([value,close]),
                                         ft.Row([sales,technical,budget]),competitors,next_step,ft.Row([next_date,status]),
                                         lost,notes],save); self.page.open(d)

    def quotation_dialog(self):
        opportunities = self.db.opportunities()
        if not opportunities:
            self.notify("Create an opportunity before creating a quotation.", True)
            return

        number = self.text(
            "Auto Quotation ID",
            260,
            f"Q-{datetime.now().strftime('%y%m%d%H%M%S')}",
            read_only=True,
        )
        opportunity = ft.Dropdown(
            label="Related Opportunity (required)",
            width=680,
            options=self.opportunity_options(),
            value=str(opportunities[0]["id"]),
        )
        linked_account = ft.TextField(
            label="Account inherited from opportunity",
            width=500,
            read_only=True,
        )

        opportunity_lookup = {
            str(row["id"]): row for row in opportunities
        }

        def update_linked_account(_=None):
            selected = opportunity_lookup.get(str(opportunity.value))
            linked_account.value = (
                selected["account_name"] if selected else ""
            )
            linked_account.update()

        opportunity.on_change = update_linked_account
        update_linked_account()

        qdate_control, qdate = self.calendar_field(
            "Quotation Date",
            date.today().isoformat(),
            width=260,
        )
        valid_control, valid = self.calendar_field(
            "Valid Until",
            (date.today() + timedelta(days=90)).isoformat(),
            width=260,
        )
        base = self.text("Base Value", 220, "0")
        discount = self.text("Discount", 220, "0")
        vat = self.text("VAT Rate %", 180, "15")
        status = self.dropdown("Status", QUOTATION_STATUSES, "Draft", 240)
        revision = self.text("Revision", 140, "0")
        quotation_owner_options = self.owner_options(quotation_only=False)
        owner = ft.Dropdown(
            label="Eligible for Quotation Assignment",
            width=360,
            options=quotation_owner_options,
            value=(
                quotation_owner_options[0].key
                if quotation_owner_options else None
            ),
        )
        notes = self.text("Notes", 650, multiline=True)

        def save(_):
            try:
                if not opportunity.value:
                    raise ValueError("A quotation must be linked to an opportunity.")
                valid_date(qdate.value)
                valid_date(valid.value)

                selected = opportunity_lookup.get(str(opportunity.value))
                if not selected:
                    raise ValueError("The selected opportunity no longer exists.")

                self.db.insert_quotation({
                    "quotation_number": number.value.strip(),
                    "opportunity_id": int(opportunity.value),
                    "account_id": selected["account_id"],
                    "quotation_date": qdate.value,
                    "valid_until": valid.value,
                    "base_value": safe_float(base.value),
                    "discount": safe_float(discount.value),
                    "vat_rate": safe_float(vat.value),
                    "status": status.value,
                    "revision": safe_int(revision.value),
                    "owner": owner.value,
                    "notes": notes.value,
                })
                self.page.close(dialog)
                self.notify("Quotation created and linked to the selected opportunity.")
                self.navigate("quotations")
            except Exception as error:
                self.notify(f"Could not create quotation: {error}", True)

        dialog = self.dialog(
            "New Quotation",
            [
                number,
                opportunity,
                linked_account,
                ft.Row([qdate_control, valid_control]),
                ft.Row([base, discount, vat]),
                ft.Row([status, revision, owner]),
                notes,
            ],
            save,
        )
        self.page.open(dialog)

    def poc_dialog(self):
        opps=self.db.opportunities(); accounts=self.db.accounts()
        if not accounts: self.notify("Create an account first.",True); return
        ref=self.text("Auto PoC ID",230,f"POC-{datetime.now().strftime('%y%m%d%H%M%S')}",read_only=True)
        opportunity=ft.Dropdown(label="Opportunity",width=520,options=self.opportunity_options(),value=str(opps[0]["id"]) if opps else None)
        account=ft.Dropdown(label="Account",width=420,options=self.account_options(),value=str(accounts[0]["id"]))
        title=self.text("PoC Title",520); solution=self.text("Solution / Technology",420)
        start=self.text("Start Date YYYY-MM-DD",260,date.today().isoformat())
        end=self.text("Planned End Date YYYY-MM-DD",260,(date.today()+timedelta(days=60)).isoformat())
        status=self.dropdown("Status",POC_STATUSES,"Planned",260); criteria=self.text("Success Criteria",650,multiline=True)
        cost=self.text("Estimated Cost",220,"0"); commercial=self.text("Commercial Value",220,"0")
        owner=self.text("Owner",240); outcome=self.text("Outcome",600,multiline=True)
        next_step=self.text("Next Step",600); notes=self.text("Notes",650,multiline=True)
        def save(_):
            try:
                valid_date(start.value); valid_date(end.value)
                self.db.insert_poc({"poc_reference":ref.value.strip(),"opportunity_id":int(opportunity.value) if opportunity.value else None,
                    "account_id":int(account.value),"poc_title":title.value.strip(),"solution":solution.value,
                    "start_date":start.value,"planned_end_date":end.value,"status":status.value,
                    "success_criteria":criteria.value,"estimated_cost":safe_float(cost.value),
                    "commercial_value":safe_float(commercial.value),"owner":owner.value,"outcome":outcome.value,
                    "next_step":next_step.value,"notes":notes.value})
                self.page.close(d); self.notify("PoC created."); self.navigate("pocs")
            except Exception as e: self.notify(f"Could not create PoC: {e}",True)
        d=self.dialog("New PoC",[ft.Row([ref,account]),opportunity,title,solution,ft.Row([start,end,status]),
                                criteria,ft.Row([cost,commercial,owner]),outcome,next_step,notes],save)
        self.page.open(d)

    def meeting_dialog(self):
        accounts = self.db.accounts()
        opportunities = self.db.opportunities()
        if not accounts:
            self.notify("Create an account first.", True)
            return

        meeting_id_value = f"MTG-{datetime.now().strftime('%y%m%d%H%M%S')}"
        reference = self.labelled_readonly_field(
            "Auto Meeting ID",
            meeting_id_value,
            width=360,
        )
        account = ft.Dropdown(
            label="Related Account",
            width=650,
            options=self.account_options(),
            value=str(accounts[0]["id"]),
        )
        opportunity = ft.Dropdown(
            label="Related Opportunity (optional)",
            width=650,
            options=[ft.dropdown.Option(key="", text="No linked opportunity")]
                    + self.opportunity_options(),
            value="",
        )
        meeting_date = self.text(
            "Meeting Date YYYY-MM-DD",
            280,
            date.today().isoformat(),
        )
        meeting_type = self.dropdown(
            "Meeting Type",
            MEETING_TYPES,
            width=300,
        )
        subject = self.text("Meeting Subject", 650)
        location = self.text("Location / Online Platform", 420)
        attendees = self.text("Attendees", 650, multiline=True)
        owner = self.text("Meeting Owner", 300)
        outcome = self.text("Meeting Outcome", 650, multiline=True)
        next_action = self.text("Agreed Next Action", 650)
        next_date = self.text(
            "Next Action Due Date YYYY-MM-DD",
            320,
            (date.today() + timedelta(days=7)).isoformat(),
        )
        notes = self.text("Additional Notes", 650, multiline=True)

        opportunity_lookup = {
            str(row["id"]): row for row in opportunities
        }

        def opportunity_changed(_):
            selected = opportunity_lookup.get(str(opportunity.value))
            if selected and selected["account_id"]:
                account.value = str(selected["account_id"])
                account.update()

        opportunity.on_change = opportunity_changed

        def save(_):
            try:
                valid_date(meeting_date.value)
                valid_date(next_date.value)
                self.db.insert_meeting({
                    "meeting_reference": meeting_id_value,
                    "account_id": int(account.value),
                    "opportunity_id": int(opportunity.value) if opportunity.value else None,
                    "meeting_date": meeting_date.value,
                    "meeting_type": meeting_type.value,
                    "subject": subject.value,
                    "location": location.value,
                    "attendees": attendees.value,
                    "owner": owner.value,
                    "outcome": outcome.value,
                    "next_action": next_action.value,
                    "next_action_date": next_date.value,
                    "notes": notes.value,
                })
                self.page.close(dialog)
                self.notify("Meeting recorded.")
                self.navigate("meetings")
            except Exception as error:
                self.notify(f"Could not create meeting: {error}", True)

        dialog = self.dialog(
            "New Meeting",
            [
                reference,
                account,
                opportunity,
                ft.Row([meeting_date, meeting_type]),
                subject,
                location,
                attendees,
                owner,
                outcome,
                next_action,
                next_date,
                notes,
            ],
            save,
        )
        self.page.open(dialog)

    def activity_dialog(self):
        accounts=self.db.accounts(); opps=self.db.opportunities()
        if not accounts: self.notify("Create an account first.",True); return
        ref=self.text("Auto Activity ID",230,f"ACT-{datetime.now().strftime('%y%m%d%H%M%S')}",read_only=True)
        account=ft.Dropdown(label="Account",width=420,options=self.account_options(),value=str(accounts[0]["id"]))
        opportunity=ft.Dropdown(label="Opportunity",width=520,options=self.opportunity_options(),value=str(opps[0]["id"]) if opps else None)
        atype=self.dropdown("Activity Type",ACTIVITY_TYPES,width=270); subject=self.text("Subject",520)
        due=self.text("Due Date YYYY-MM-DD",260,(date.today()+timedelta(days=7)).isoformat())
        priority=self.dropdown("Priority",PRIORITIES,"Medium",220); status=self.dropdown("Status",ACTIVITY_STATUSES,"Open",220)
        owner=self.text("Owner",250); completed=self.text("Completed Date YYYY-MM-DD",260)
        notes=self.text("Notes",650,multiline=True)
        def save(_):
            try:
                valid_date(due.value); valid_date(completed.value)
                self.db.insert_activity({"activity_reference":ref.value.strip(),"account_id":int(account.value),
                    "opportunity_id":int(opportunity.value) if opportunity.value else None,"activity_type":atype.value,
                    "subject":subject.value,"due_date":due.value,"priority":priority.value,"status":status.value,
                    "owner":owner.value,"completed_date":completed.value,"notes":notes.value})
                self.page.close(d); self.notify("Activity created."); self.navigate("activities")
            except Exception as e: self.notify(f"Could not create activity: {e}",True)
        d=self.dialog("New Activity",[ft.Row([ref,account]),opportunity,ft.Row([atype,due]),subject,
                                     ft.Row([priority,status,owner]),completed,notes],save)
        self.page.open(d)

    def dialog(self, title, controls, save):
        d=ft.AlertDialog(modal=True,title=ft.Text(title,weight=ft.FontWeight.BOLD),
                         content=ft.Container(width=780,content=ft.Column(height=620,scroll=ft.ScrollMode.AUTO,controls=controls)),
                         actions=[ft.TextButton("Cancel",on_click=lambda _:self.page.close(d)),
                                  ft.FilledButton("Save",icon=ft.Icons.SAVE,on_click=save,style=ft.ButtonStyle(bgcolor="#0E5A47"))])
        return d

    def import_migration(self):
        counts=Counter()
        errors=[]
        try:
            def read_file(name):
                path=MIGRATION_DIR/name
                if not path.exists(): return []
                with path.open("r",encoding="utf-8-sig",newline="") as f:
                    return list(csv.DictReader(f))

            for r in read_file("accounts.csv"):
                if not r.get("account_name") or r["account_name"].startswith("Example"): continue
                try:
                    self.db.insert_account({
                        "account_name":r.get("account_name","").strip(),"account_type":r.get("account_type","Other"),
                        "industry":r.get("industry","Other"),"city":r.get("city",""),"country":r.get("country",""),
                        "website":r.get("website",""),"main_phone":r.get("main_phone",""),"owner":r.get("owner",""),
                        "status":r.get("status","Active"),"notes":r.get("notes","")
                    },ignore=True); counts["accounts"]+=1
                except Exception as e: errors.append(f"accounts.csv: {e}")

            for r in read_file("contacts.csv"):
                if not r.get("first_name") or r["first_name"]=="Ahmed": continue
                try:
                    self.db.insert_contact({
                        "account_id":self.db.account_id(r.get("account_name","")),"first_name":r.get("first_name",""),
                        "last_name":r.get("last_name",""),"job_title":r.get("job_title",""),
                        "department":r.get("department",""),"email":r.get("email",""),"mobile":r.get("mobile",""),
                        "phone":r.get("phone",""),"influence_level":r.get("influence_level","Medium"),
                        "relationship_status":r.get("relationship_status","New"),"owner":r.get("owner",""),
                        "notes":r.get("notes","")
                    }); counts["contacts"]+=1
                except Exception as e: errors.append(f"contacts.csv: {e}")

            for r in read_file("leads.csv"):
                if not r.get("lead_reference") or r["lead_reference"]=="LD-0001": continue
                try:
                    self.db.insert_lead({
                        "lead_reference":r.get("lead_reference",""),"company_name":r.get("company_name",""),
                        "contact_name":r.get("contact_name",""),"job_title":r.get("job_title",""),
                        "email":r.get("email",""),"mobile":r.get("mobile",""),"source":r.get("source","Other"),
                        "interest_area":r.get("interest_area","Other"),"lead_status":r.get("lead_status","New"),
                        "lead_score":safe_float(r.get("lead_score",0)),"owner":r.get("owner",""),
                        "next_action_date":valid_date(r.get("next_action_date","")),
                        "estimated_value":safe_float(r.get("estimated_value",0)),"notes":r.get("notes","")
                    },ignore=True); counts["leads"]+=1
                except Exception as e: errors.append(f"leads.csv: {e}")

            for r in read_file("opportunities.csv"):
                if not r.get("opportunity_reference") or r["opportunity_reference"]=="OPP-0001": continue
                try:
                    stage=r.get("stage","Qualification")
                    self.db.insert_opportunity({
                        "opportunity_reference":r.get("opportunity_reference",""),
                        "account_id":self.db.account_id(r.get("account_name","")),
                        "opportunity_name":r.get("opportunity_name",""),"project_type":r.get("project_type","Other"),
                        "stage":stage,"probability":safe_float(r.get("probability",STAGE_PROBABILITY.get(stage,0))),
                        "estimated_value":safe_float(r.get("estimated_value",0)),
                        "expected_close_date":valid_date(r.get("expected_close_date","")),
                        "sales_owner":r.get("sales_owner",""),"technical_owner":r.get("technical_owner",""),
                        "customer_budget":safe_float(r.get("customer_budget",0)),"competitors":r.get("competitors",""),
                        "next_step":r.get("next_step",""),"next_action_date":valid_date(r.get("next_action_date","")),
                        "status":r.get("status","Open"),"lost_reason":"","notes":r.get("notes","")
                    },ignore=True); counts["opportunities"]+=1
                except Exception as e: errors.append(f"opportunities.csv: {e}")

            for r in read_file("quotations.csv"):
                if not r.get("quotation_number") or r["quotation_number"]=="Q-0001": continue
                try:
                    self.db.insert_quotation({
                        "quotation_number":r.get("quotation_number",""),
                        "opportunity_id":self.db.opportunity_id(r.get("opportunity_reference","")),
                        "account_id":self.db.account_id(r.get("account_name","")),
                        "quotation_date":valid_date(r.get("quotation_date","")),
                        "valid_until":valid_date(r.get("valid_until","")),"base_value":safe_float(r.get("base_value",0)),
                        "discount":safe_float(r.get("discount",0)),"vat_rate":safe_float(r.get("vat_rate",15)),
                        "status":r.get("status","Draft"),"revision":safe_int(r.get("revision",0)),
                        "owner":r.get("owner",""),"notes":r.get("notes","")
                    },ignore=True); counts["quotations"]+=1
                except Exception as e: errors.append(f"quotations.csv: {e}")

            for r in read_file("pocs.csv"):
                if not r.get("poc_reference") or r["poc_reference"]=="POC-0001": continue
                try:
                    self.db.insert_poc({
                        "poc_reference":r.get("poc_reference",""),
                        "opportunity_id":self.db.opportunity_id(r.get("opportunity_reference","")),
                        "account_id":self.db.account_id(r.get("account_name","")),"poc_title":r.get("poc_title",""),
                        "solution":r.get("solution",""),"start_date":valid_date(r.get("start_date","")),
                        "planned_end_date":valid_date(r.get("planned_end_date","")),"status":r.get("status","Planned"),
                        "success_criteria":r.get("success_criteria",""),"estimated_cost":safe_float(r.get("estimated_cost",0)),
                        "commercial_value":safe_float(r.get("commercial_value",0)),"owner":r.get("owner",""),
                        "outcome":r.get("outcome",""),"next_step":r.get("next_step",""),"notes":r.get("notes","")
                    },ignore=True); counts["pocs"]+=1
                except Exception as e: errors.append(f"pocs.csv: {e}")

            for r in read_file("meetings.csv"):
                if not r.get("meeting_reference") or r["meeting_reference"]=="MTG-0001": continue
                try:
                    self.db.insert_meeting({
                        "meeting_reference":r.get("meeting_reference",""),
                        "account_id":self.db.account_id(r.get("account_name","")),
                        "opportunity_id":self.db.opportunity_id(r.get("opportunity_reference","")),
                        "meeting_date":valid_date(r.get("meeting_date","")),"meeting_type":r.get("meeting_type","Other"),
                        "subject":r.get("subject",""),"location":r.get("location",""),"attendees":r.get("attendees",""),
                        "owner":r.get("owner",""),"outcome":r.get("outcome",""),"next_action":r.get("next_action",""),
                        "next_action_date":valid_date(r.get("next_action_date","")),"notes":r.get("notes","")
                    },ignore=True); counts["meetings"]+=1
                except Exception as e: errors.append(f"meetings.csv: {e}")

            for r in read_file("activities.csv"):
                if not r.get("activity_reference") or r["activity_reference"]=="ACT-0001": continue
                try:
                    self.db.insert_activity({
                        "activity_reference":r.get("activity_reference",""),
                        "account_id":self.db.account_id(r.get("account_name","")),
                        "opportunity_id":self.db.opportunity_id(r.get("opportunity_reference","")),
                        "activity_type":r.get("activity_type","Other"),"subject":r.get("subject",""),
                        "due_date":valid_date(r.get("due_date","")),"priority":r.get("priority","Medium"),
                        "status":r.get("status","Open"),"owner":r.get("owner",""),
                        "completed_date":valid_date(r.get("completed_date","")),"notes":r.get("notes","")
                    },ignore=True); counts["activities"]+=1
                except Exception as e: errors.append(f"activities.csv: {e}")

            summary=", ".join(f"{k}: {v}" for k,v in counts.items()) or "No new records"
            if errors:
                log=EXPORT_DIR/"migration_errors.txt"
                log.write_text("\n".join(errors),encoding="utf-8")
                self.notify(f"Migration completed with warnings. {summary}. See migration_errors.txt",True)
            else:
                self.notify(f"Migration completed successfully. {summary}")
            self.navigate("migration")
        except Exception as e:
            self.notify(f"Migration failed: {e}",True)

    def export_excel(self):
        try:
            stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
            path=EXPORT_DIR/f"Saudi_Sensing_CRM_{stamp}.xlsx"
            wb=xlsxwriter.Workbook(str(path))
            title=wb.add_format({"bold":True,"font_size":18,"font_name":"Calibri","font_color":"#103F37"})
            hdr=wb.add_format({"bold":True,"bg_color":"#DDEBE6","border":1,"font_name":"Calibri"})
            txt=wb.add_format({"border":1,"font_name":"Calibri"})
            mon=wb.add_format({"border":1,"num_format":'#,##0.00 "SAR"',"font_name":"Calibri"})
            pct=wb.add_format({"border":1,"num_format":"0.0%","font_name":"Calibri"})

            d=self.db.dashboard()
            ws=wb.add_worksheet("Executive Dashboard")
            ws.set_column("A:A",30); ws.set_column("B:B",22)
            ws.write("A1","Saudi Sensing CRM",title)
            metrics=[
                ("Accounts",d["accounts"]),("Leads",d["leads"]),("Open Opportunities",d["open_opportunities"]),
                ("Pipeline",d["pipeline"]),("Weighted Pipeline",d["weighted_pipeline"]),
                ("Open Quotation Value",d["quotation_value"]),("Win Rate",d["win_rate"]/100),
                ("Lead Conversion",d["lead_conversion"]/100),("Active PoCs",d["active_pocs"]),
                ("Overdue Activities",d["overdue_activities"]),
            ]
            ws.write_row("A3",["Metric","Value"],hdr)
            for i,(label,val) in enumerate(metrics,start=3):
                ws.write(i,0,label,txt)
                if label in ("Win Rate","Lead Conversion"): ws.write_number(i,1,val,pct)
                elif label in ("Pipeline","Weighted Pipeline","Open Quotation Value"): ws.write_number(i,1,val,mon)
                else: ws.write(i,1,val,txt)

            def sheet(name,headers,rows,money_cols=()):
                sh=wb.add_worksheet(name); sh.freeze_panes(1,0)
                for c,h in enumerate(headers): sh.write(0,c,h,hdr)
                for ri,row in enumerate(rows,start=1):
                    for ci,v in enumerate(row):
                        if ci in money_cols: sh.write_number(ri,ci,float(v or 0),mon)
                        else: sh.write(ri,ci,"" if v is None else v,txt)
                sh.autofilter(0,0,max(1,len(rows)),len(headers)-1)
                sh.set_column(0,len(headers)-1,20)

            sheet("Accounts",["Account","Type","Industry","City","Country","Website","Phone","Owner","Status","Notes"],
                  [[r["account_name"],r["account_type"],r["industry"],r["city"],r["country"],r["website"],r["main_phone"],r["owner"],r["status"],r["notes"]] for r in self.db.accounts()])
            sheet("Contacts",["Account","First Name","Last Name","Job Title","Department","Email","Mobile","Phone","Influence","Relationship","Owner","Notes"],
                  [[r["account_name"],r["first_name"],r["last_name"],r["job_title"],r["department"],r["email"],r["mobile"],r["phone"],r["influence_level"],r["relationship_status"],r["owner"],r["notes"]] for r in self.db.contacts()])
            sheet("Leads",["Reference","Company","Contact","Title","Email","Mobile","Source","Interest","Status","Score","Owner","Next Action","Estimated Value","Notes"],
                  [[r["lead_reference"],r["company_name"],r["contact_name"],r["job_title"],r["email"],r["mobile"],r["source"],r["interest_area"],r["lead_status"],r["lead_score"],r["owner"],r["next_action_date"],r["estimated_value"],r["notes"]] for r in self.db.leads()],(12,))
            sheet("Opportunities",["Reference","Account","Opportunity","Type","Stage","Probability","Estimated Value","Weighted Value","Expected Close","Sales Owner","Technical Owner","Customer Budget","Competitors","Next Step","Next Action","Status","Lost Reason","Notes"],
                  [[r["opportunity_reference"],r["account_name"],r["opportunity_name"],r["project_type"],r["stage"],r["probability"],r["estimated_value"],float(r["estimated_value"])*float(r["probability"])/100,r["expected_close_date"],r["sales_owner"],r["technical_owner"],r["customer_budget"],r["competitors"],r["next_step"],r["next_action_date"],r["status"],r["lost_reason"],r["notes"]] for r in self.db.opportunities()],(6,7,11))
            sheet("Quotations",["Quotation","Opportunity","Account","Date","Valid Until","Base Value","Discount","VAT %","Status","Revision","Owner","Notes"],
                  [[r["quotation_number"],r["opportunity_reference"],r["account_name"],r["quotation_date"],r["valid_until"],r["base_value"],r["discount"],r["vat_rate"],r["status"],r["revision"],r["owner"],r["notes"]] for r in self.db.quotations()],(5,6))
            sheet("PoCs",["Reference","Opportunity","Account","Title","Solution","Start","End","Status","Success Criteria","Estimated Cost","Commercial Value","Owner","Outcome","Next Step","Notes"],
                  [[r["poc_reference"],r["opportunity_reference"],r["account_name"],r["poc_title"],r["solution"],r["start_date"],r["planned_end_date"],r["status"],r["success_criteria"],r["estimated_cost"],r["commercial_value"],r["owner"],r["outcome"],r["next_step"],r["notes"]] for r in self.db.pocs()],(9,10))
            sheet("Meetings",["Reference","Account","Opportunity","Date","Type","Subject","Location","Attendees","Owner","Outcome","Next Action","Next Action Date","Notes"],
                  [[r["meeting_reference"],r["account_name"],r["opportunity_reference"],r["meeting_date"],r["meeting_type"],r["subject"],r["location"],r["attendees"],r["owner"],r["outcome"],r["next_action"],r["next_action_date"],r["notes"]] for r in self.db.meetings()])
            sheet("Activities",["Reference","Account","Opportunity","Type","Subject","Due Date","Priority","Status","Owner","Completed Date","Notes"],
                  [[r["activity_reference"],r["account_name"],r["opportunity_reference"],r["activity_type"],r["subject"],r["due_date"],r["priority"],r["status"],r["owner"],r["completed_date"],r["notes"]] for r in self.db.activities()])

            legacy = self.db.legacy_opportunities()
            sheet("Migrated Pipeline",
                  ["Reference","Account","Customer Source","End User","Deal Name","Business Unit",
                   "Source Project Type","Industry","Competitive","Probability Band","Probability %",
                   "Gross","Source Currency","GM %","GM Value","GM Basis","Expected PO Year",
                   "Expected PO Month","Quarter","Delivery Date","Created By","Sales Owner",
                   "Assigned To","Include in Forecast","Source Stage","CRM Stage","Status",
                   "Opportunity Update","Must Win","Suspended","Quality Flags","Source Row"],
                  [[r["opportunity_reference"],r["account_name"],r["customer_name_source"],r["end_user"],
                    r["opportunity_name"],r["business_unit"],r["source_project_type"],r["industry"],
                    "Yes" if r["competitive"] else "No",r["probability_band"],r["probability"],
                    r["estimated_value"],r["source_currency"],r["forecast_gm_percent"],
                    r["forecast_gm_value"],r["gm_value_basis"],r["expected_po_year"],
                    r["expected_po_month"],r["quarter"],r["delivery_date"],r["created_by"],
                    r["sales_owner"],r["assigned_to"],"Yes" if r["include_in_forecast"] else "No",
                    r["source_stage"],r["stage"],r["status"],r["opportunity_update"],
                    "Yes" if r["must_win"] else "No","Yes" if r["suspended"] else "No",
                    r["quality_flags"],r["source_row"]] for r in legacy],(11,14))
            wb.close(); self.notify(f"Excel exported: {path.name}"); os.startfile(path)
        except Exception as e: self.notify(f"Excel export failed: {e}",True)

    def export_pdf(self):
        try:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = EXPORT_DIR / f"Saudi_Sensing_CRM_Executive_{stamp}.pdf"
            dashboard = self.calculate_filtered_dashboard()
            opportunities = self.filtered_opportunities()
            activities = self.db.activities()

            open_opportunities = [
                item for item in opportunities if item["status"] == "Open"
            ]
            won_opportunities = [
                item for item in opportunities
                if item["status"] == "Won" or item["stage"] == "Awarded"
            ]
            lost_opportunities = [
                item for item in opportunities
                if item["status"] == "Lost" or item["stage"] == "Lost"
            ]

            stage_values = defaultdict(float)
            stage_counts = defaultdict(int)
            owner_values = defaultdict(float)
            owner_weighted = defaultdict(float)
            close_forecast = defaultdict(float)

            for item in open_opportunities:
                value = float(item["estimated_value"] or 0)
                probability = float(item["probability"] or 0)
                weighted = value * probability / 100
                stage = item["stage"] or "Unspecified"
                owner = item["sales_owner"] or "Unassigned"

                stage_values[stage] += value
                stage_counts[stage] += 1
                owner_values[owner] += value
                owner_weighted[owner] += weighted

                close_date = item["expected_close_date"] or ""
                if len(close_date) >= 7:
                    close_forecast[close_date[:7]] += weighted

            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                "CRMTitle",
                parent=styles["Title"],
                fontName="Helvetica-Bold",
                fontSize=23,
                leading=27,
                alignment=1,
                textColor=colors.HexColor("#103F37"),
                spaceAfter=3 * mm,
            )
            subtitle_style = ParagraphStyle(
                "CRMSubtitle",
                parent=styles["Heading2"],
                fontName="Helvetica-Bold",
                fontSize=13,
                leading=16,
                alignment=1,
                textColor=colors.HexColor("#397064"),
                spaceAfter=4 * mm,
            )
            section_style = ParagraphStyle(
                "CRMSection",
                parent=styles["Heading2"],
                fontName="Helvetica-Bold",
                fontSize=14,
                leading=17,
                textColor=colors.HexColor("#103F37"),
                spaceBefore=2 * mm,
                spaceAfter=2 * mm,
            )
            body_style = ParagraphStyle(
                "CRMBody",
                parent=styles["BodyText"],
                fontName="Helvetica",
                fontSize=9.5,
                leading=13,
                textColor=colors.HexColor("#304B45"),
            )
            small_style = ParagraphStyle(
                "CRMSmall",
                parent=styles["BodyText"],
                fontName="Helvetica",
                fontSize=7.2,
                leading=9,
                textColor=colors.HexColor("#263D38"),
                wordWrap="CJK",
            )
            small_center_style = ParagraphStyle(
                "CRMSmallCenter",
                parent=small_style,
                alignment=1,
            )
            small_bold_style = ParagraphStyle(
                "CRMSmallBold",
                parent=small_style,
                fontName="Helvetica-Bold",
                textColor=colors.HexColor("#103F37"),
            )
            kpi_label_style = ParagraphStyle(
                "KpiLabel",
                parent=small_center_style,
                fontName="Helvetica-Bold",
                fontSize=7.8,
                leading=9.5,
                textColor=colors.HexColor("#234F46"),
            )
            kpi_value_style = ParagraphStyle(
                "KpiValue",
                parent=small_center_style,
                fontName="Helvetica-Bold",
                fontSize=10.5,
                leading=12,
                textColor=colors.HexColor("#103F37"),
            )

            def p(value, style=small_style):
                safe_value = "" if value is None else str(value)
                safe_value = (
                    safe_value.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace("\n", "<br/>")
                )
                return Paragraph(safe_value, style)

            def page_footer(canvas, document):
                canvas.saveState()
                canvas.setStrokeColor(colors.HexColor("#C8D8D3"))
                canvas.line(
                    document.leftMargin,
                    9 * mm,
                    landscape(A4)[0] - document.rightMargin,
                    9 * mm,
                )
                canvas.setFont("Helvetica", 7.5)
                canvas.setFillColor(colors.HexColor("#607C75"))
                canvas.drawString(
                    document.leftMargin,
                    5.5 * mm,
                    f"Saudi Sensing CRM | Generated {datetime.now().strftime('%d %b %Y %H:%M')}",
                )
                canvas.drawRightString(
                    landscape(A4)[0] - document.rightMargin,
                    5.5 * mm,
                    f"Page {document.page}",
                )
                canvas.restoreState()

            document = SimpleDocTemplate(
                str(path),
                pagesize=landscape(A4),
                leftMargin=11 * mm,
                rightMargin=11 * mm,
                topMargin=10 * mm,
                bottomMargin=13 * mm,
                title="Saudi Sensing CRM Executive Sales and Pipeline Report",
                author=self.user["full_name"] if self.user else "Saudi Sensing",
            )

            with tempfile.TemporaryDirectory() as temporary_directory:
                temp_dir = Path(temporary_directory)
                stage_chart = temp_dir / "stage_pipeline.png"
                forecast_chart = temp_dir / "weighted_forecast.png"
                owner_chart = temp_dir / "owner_pipeline.png"

                stage_order = [
                    stage for stage in OPPORTUNITY_STAGES if stage_values.get(stage)
                ]
                stage_chart_values = [stage_values[stage] for stage in stage_order]
                stage_colors = [
                    "#6C9E8E", "#3E7F6F", "#2E8B57", "#2D6A8E",
                    "#7C5C9E", "#C17D24", "#D35400", "#1F6F43",
                    "#B54735", "#7F8C8D",
                ]

                if stage_order:
                    plt.figure(figsize=(10.5, 4.15))
                    bars = plt.bar(
                        stage_order,
                        stage_chart_values,
                        color=stage_colors[:len(stage_order)],
                        width=0.62,
                    )
                    plt.title("Open Pipeline by CRM Stage", fontweight="bold", pad=12)
                    plt.ylabel("Pipeline Value")
                    plt.xticks(rotation=22, ha="right")
                    plt.grid(axis="y", alpha=0.16)
                    plt.ticklabel_format(style="plain", axis="y")
                    plt.gca().yaxis.set_major_formatter(
                        plt.FuncFormatter(lambda value, _: f"{value/1_000_000:.0f}M")
                    )
                    for bar, value in zip(bars, stage_chart_values):
                        plt.text(
                            bar.get_x() + bar.get_width() / 2,
                            bar.get_height(),
                            f"{value/1_000_000:.1f}M",
                            ha="center",
                            va="bottom",
                            fontsize=8,
                            fontweight="bold",
                        )
                    plt.box(False)
                    plt.tight_layout()
                    plt.savefig(stage_chart, dpi=190, bbox_inches="tight", facecolor="white")
                    plt.close()

                forecast_months = sorted(close_forecast)
                if forecast_months:
                    forecast_values = [close_forecast[month] for month in forecast_months]
                    plt.figure(figsize=(10.5, 3.75))
                    plt.plot(
                        forecast_months,
                        forecast_values,
                        marker="o",
                        linewidth=2.5,
                        color="#1667A8",
                    )
                    plt.fill_between(
                        forecast_months,
                        forecast_values,
                        alpha=0.12,
                        color="#1667A8",
                    )
                    plt.title("Weighted Opportunity Forecast by Expected Closing Month", fontweight="bold", pad=12)
                    plt.ylabel("Weighted Value")
                    plt.xticks(rotation=35, ha="right", fontsize=8)
                    plt.grid(axis="y", alpha=0.18)
                    plt.gca().yaxis.set_major_formatter(
                        plt.FuncFormatter(lambda value, _: f"{value/1_000_000:.0f}M")
                    )
                    plt.box(False)
                    plt.tight_layout()
                    plt.savefig(forecast_chart, dpi=190, bbox_inches="tight", facecolor="white")
                    plt.close()

                top_owners = sorted(
                    owner_values.items(),
                    key=lambda pair: pair[1],
                    reverse=True,
                )[:8]
                if top_owners:
                    owner_names = [item[0] for item in top_owners][::-1]
                    owner_pipeline_values = [item[1] for item in top_owners][::-1]
                    plt.figure(figsize=(8.5, 4.3))
                    bars = plt.barh(
                        owner_names,
                        owner_pipeline_values,
                        color="#3E7F6F",
                    )
                    plt.title("Top Sales Owners by Open Pipeline", fontweight="bold", pad=12)
                    plt.xlabel("Pipeline Value")
                    plt.grid(axis="x", alpha=0.16)
                    plt.gca().xaxis.set_major_formatter(
                        plt.FuncFormatter(lambda value, _: f"{value/1_000_000:.0f}M")
                    )
                    for bar, value in zip(bars, owner_pipeline_values):
                        plt.text(
                            bar.get_width(),
                            bar.get_y() + bar.get_height() / 2,
                            f" {value/1_000_000:.1f}M",
                            va="center",
                            fontsize=8,
                            fontweight="bold",
                        )
                    plt.box(False)
                    plt.tight_layout()
                    plt.savefig(owner_chart, dpi=190, bbox_inches="tight", facecolor="white")
                    plt.close()

                story = [
                    Table([[Image(str(REPORT_LOGO_PATH), width=62.00 * mm, height=28.27 * mm)]], colWidths=[landscape(A4)[0] - 22 * mm], style=TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),("VALIGN", (0, 0), (-1, -1), "MIDDLE"),("LEFTPADDING", (0, 0), (-1, -1), 0),("RIGHTPADDING", (0, 0), (-1, -1), 0),("TOPPADDING", (0, 0), (-1, -1), 0),("BOTTOMPADDING", (0, 0), (-1, -1), 0),])),
                    Spacer(1, 2 * mm),
                    Paragraph("Saudi Sensing CRM", title_style),
                    Paragraph("Executive Sales & Opportunity Pipeline Report", subtitle_style),
                    Paragraph(
                        f"Applied filters: {self.dashboard_filter_summary()}",
                        body_style,
                    ),
                    Spacer(1, 2 * mm),
                ]

                kpi_items = [
                    ("Accounts", dashboard["accounts"]),
                    ("Open Opportunities", dashboard["open_opportunities"]),
                    ("Total Pipeline", money(dashboard["pipeline"])),
                    ("Weighted Pipeline", money(dashboard["weighted_pipeline"])),
                    ("Win Rate", f'{dashboard["win_rate"]:.1f}%'),
                    ("Active PoCs", dashboard["active_pocs"]),
                    ("Open Quote Value", money(dashboard["quotation_value"])),
                    ("Overdue Actions", dashboard["overdue_activities"]),
                ]
                kpi_table_data = []
                for start_index in (0, 4):
                    labels = [
                        p(label, kpi_label_style)
                        for label, _ in kpi_items[start_index:start_index + 4]
                    ]
                    values = [
                        p(value, kpi_value_style)
                        for _, value in kpi_items[start_index:start_index + 4]
                    ]
                    kpi_table_data.extend([labels, values])

                kpi_table = Table(
                    kpi_table_data,
                    colWidths=[65 * mm] * 4,
                    rowHeights=[9 * mm, 12 * mm, 9 * mm, 12 * mm],
                )
                kpi_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBE6")),
                            ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#DDEBE6")),
                            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F7FAF9")),
                            ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#F7FAF9")),
                            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#9CB8B0")),
                            ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#BFD0CB")),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 5),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )
                story.extend([kpi_table, Spacer(1, 4 * mm)])

                pipeline_value = float(dashboard["pipeline"] or 0)
                weighted_value = float(dashboard["weighted_pipeline"] or 0)
                weighting_ratio = (
                    weighted_value / pipeline_value * 100 if pipeline_value else 0
                )
                executive_summary = (
                    f"The CRM currently contains {dashboard['open_opportunities']} open opportunities "
                    f"with a total pipeline of {money(pipeline_value)}. The weighted pipeline is "
                    f"{money(weighted_value)}, representing {weighting_ratio:.1f}% of the gross open pipeline. "
                    f"The recorded win rate is {dashboard['win_rate']:.1f}%. "
                    f"There are {dashboard['overdue_activities']} overdue commercial actions and "
                    f"{dashboard['due_7_days']} actions due within the next seven days."
                )
                story.extend(
                    [
                        Paragraph("Executive Summary", section_style),
                        Paragraph(executive_summary, body_style),
                        Spacer(1, 2 * mm),
                    ]
                )

                if stage_chart.exists():
                    story.append(Image(str(stage_chart), width=263 * mm, height=102 * mm))

                story.append(PageBreak())

                story.extend(
                    [
                        Paragraph("Forecast and Ownership Analysis", section_style),
                    ]
                )
                chart_row = []
                chart_widths = []
                if forecast_chart.exists():
                    chart_row.append(Image(str(forecast_chart), width=154 * mm, height=83 * mm))
                    chart_widths.append(158 * mm)
                if owner_chart.exists():
                    chart_row.append(Image(str(owner_chart), width=105 * mm, height=83 * mm))
                    chart_widths.append(109 * mm)
                if chart_row:
                    charts_table = Table([chart_row], colWidths=chart_widths)
                    charts_table.setStyle(
                        TableStyle(
                            [
                                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                            ]
                        )
                    )
                    story.append(charts_table)
                    story.append(Spacer(1, 3 * mm))

                stage_table_data = [
                    [
                        p("CRM Stage", small_bold_style),
                        p("Opportunity Count", small_bold_style),
                        p("Gross Pipeline", small_bold_style),
                        p("Share of Open Pipeline", small_bold_style),
                        p("Weighted Value", small_bold_style),
                    ]
                ]
                for stage in stage_order:
                    gross = stage_values[stage]
                    default_probability = STAGE_PROBABILITY.get(stage, 0)
                    stage_table_data.append(
                        [
                            p(stage),
                            p(stage_counts[stage], small_center_style),
                            p(money(gross), small_center_style),
                            p(f"{gross/pipeline_value*100:.1f}%" if pipeline_value else "0.0%", small_center_style),
                            p(money(gross * default_probability / 100), small_center_style),
                        ]
                    )
                stage_table = Table(
                    stage_table_data,
                    repeatRows=1,
                    colWidths=[55 * mm, 40 * mm, 50 * mm, 52 * mm, 55 * mm],
                )
                stage_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBE6")),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#AFC3BD")),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                                colors.white,
                                colors.HexColor("#F7FAF9"),
                            ]),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )
                story.extend(
                    [
                        Paragraph("Pipeline Stage Breakdown", section_style),
                        stage_table,
                    ]
                )

                story.append(PageBreak())

                top_opportunities = sorted(
                    open_opportunities,
                    key=lambda item: float(item["estimated_value"] or 0),
                    reverse=True,
                )[:15]

                top_data = [
                    [
                        p("Reference", small_bold_style),
                        p("Account", small_bold_style),
                        p("Opportunity", small_bold_style),
                        p("Stage", small_bold_style),
                        p("Probability", small_bold_style),
                        p("Gross Value", small_bold_style),
                        p("Weighted Value", small_bold_style),
                        p("Expected Close", small_bold_style),
                        p("Sales Owner", small_bold_style),
                    ]
                ]

                for item in top_opportunities:
                    gross_value = float(item["estimated_value"] or 0)
                    probability = float(item["probability"] or 0)
                    top_data.append(
                        [
                            p(item["opportunity_reference"], small_center_style),
                            p(item["account_name"] or "Unassigned"),
                            p(item["opportunity_name"]),
                            p(item["stage"], small_center_style),
                            p(f"{probability:.0f}%", small_center_style),
                            p(money(gross_value), small_center_style),
                            p(money(gross_value * probability / 100), small_center_style),
                            p(item["expected_close_date"] or "", small_center_style),
                            p(item["sales_owner"] or "Unassigned"),
                        ]
                    )

                top_table = Table(
                    top_data,
                    repeatRows=1,
                    colWidths=[
                        23 * mm,
                        34 * mm,
                        60 * mm,
                        30 * mm,
                        23 * mm,
                        31 * mm,
                        31 * mm,
                        27 * mm,
                        31 * mm,
                    ],
                )
                top_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBE6")),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#AFC3BD")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                                colors.white,
                                colors.HexColor("#F7FAF9"),
                            ]),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                            ("LEFTPADDING", (0, 0), (-1, -1), 4),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ]
                    )
                )
                story.extend(
                    [
                        Paragraph("Top 15 Open Opportunities by Gross Value", section_style),
                        Paragraph(
                            "Long opportunity and account names are wrapped inside their cells to preserve readability.",
                            body_style,
                        ),
                        Spacer(1, 2 * mm),
                        top_table,
                    ]
                )

                story.append(PageBreak())

                overdue_activities = []
                today = date.today()
                for activity in activities:
                    due_date = activity["due_date"] or ""
                    if not due_date or activity["status"] in ("Completed", "Cancelled"):
                        continue
                    try:
                        parsed_due_date = datetime.strptime(due_date, "%Y-%m-%d").date()
                    except ValueError:
                        continue
                    if parsed_due_date < today:
                        overdue_activities.append(activity)

                action_data = [
                    [
                        p("Activity", small_bold_style),
                        p("Account", small_bold_style),
                        p("Opportunity", small_bold_style),
                        p("Subject", small_bold_style),
                        p("Due Date", small_bold_style),
                        p("Priority", small_bold_style),
                        p("Owner", small_bold_style),
                    ]
                ]
                for activity in overdue_activities[:20]:
                    action_data.append(
                        [
                            p(activity["activity_reference"], small_center_style),
                            p(activity["account_name"] or ""),
                            p(activity["opportunity_reference"] or ""),
                            p(activity["subject"]),
                            p(activity["due_date"], small_center_style),
                            p(activity["priority"], small_center_style),
                            p(activity["owner"] or ""),
                        ]
                    )

                if len(action_data) == 1:
                    action_data.append(
                        [
                            p("No overdue activities", small_center_style),
                            p(""),
                            p(""),
                            p(""),
                            p(""),
                            p(""),
                            p(""),
                        ]
                    )

                action_table = Table(
                    action_data,
                    repeatRows=1,
                    colWidths=[
                        29 * mm,
                        43 * mm,
                        32 * mm,
                        78 * mm,
                        27 * mm,
                        24 * mm,
                        42 * mm,
                    ],
                )
                action_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2D7D5")),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C9AAA6")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                                colors.white,
                                colors.HexColor("#FFF9F8"),
                            ]),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ]
                    )
                )

                owner_table_data = [
                    [
                        p("Sales Owner", small_bold_style),
                        p("Open Opportunities", small_bold_style),
                        p("Gross Pipeline", small_bold_style),
                        p("Weighted Pipeline", small_bold_style),
                    ]
                ]
                for owner, gross in sorted(
                    owner_values.items(),
                    key=lambda pair: pair[1],
                    reverse=True,
                )[:15]:
                    owner_count = sum(
                        1 for item in open_opportunities
                        if (item["sales_owner"] or "Unassigned") == owner
                    )
                    owner_table_data.append(
                        [
                            p(owner),
                            p(owner_count, small_center_style),
                            p(money(gross), small_center_style),
                            p(money(owner_weighted[owner]), small_center_style),
                        ]
                    )

                owner_table = Table(
                    owner_table_data,
                    repeatRows=1,
                    colWidths=[70 * mm, 48 * mm, 65 * mm, 65 * mm],
                )
                owner_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBE6")),
                            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#AFC3BD")),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                                colors.white,
                                colors.HexColor("#F7FAF9"),
                            ]),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )

                story.extend(
                    [
                        Paragraph("Sales Owner Performance", section_style),
                        owner_table,
                        Spacer(1, 5 * mm),
                        Paragraph("Overdue Commercial Actions", section_style),
                        action_table,
                    ]
                )

                document.build(
                    story,
                    onFirstPage=page_footer,
                    onLaterPages=page_footer,
                )

            self.notify(f"PDF exported: {path.name}")
            os.startfile(path)
        except Exception as error:
            error_path = EXPORT_DIR / "crm_pdf_export_error.log"
            try:
                error_path.write_text(
                    f"{datetime.now().isoformat(timespec='seconds')}\n"
                    f"{type(error).__name__}: {error}",
                    encoding="utf-8",
                )
            except Exception:
                pass
            self.notify(
                f"PDF export failed: {error}. Details saved in {error_path.name}",
                True,
            )

    def backup(self):
        try:
            folder=EXPORT_DIR/"backups"; folder.mkdir(exist_ok=True)
            path=folder/f"crm_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy2(DB_PATH,path); self.notify(f"Backup created: {path.name}"); os.startfile(folder)
        except Exception as e: self.notify(f"Backup failed: {e}",True)


def main(page: ft.Page):
    CRMApp(page)


if __name__ == "__main__":
    print(f"CRM database backend: {DATABASE_BACKEND}")
    if WEB_MODE:
        ft.app(
            target=main,
            view=ft.AppView.WEB_BROWSER,
            host=APP_HOST,
            port=APP_PORT,
            assets_dir=str(APP_DIR / "assets"),
        )
    else:
        ft.app(
            target=main,
            assets_dir=str(APP_DIR / "assets"),
        )
