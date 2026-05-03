"""
MJR West Industrial Property Intelligence Agent
Single-file Flask application — main.py

pip install flask supabase twilio apscheduler openpyxl python-dotenv \
            openai pypdf python-docx Pillow
"""

# ── Imports ───────────────────────────────────────────────────────────────────
import os, json, logging, imaplib, email, re, io, smtplib, threading, base64, mimetypes
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from flask import (Flask, request, jsonify, render_template_string,
                   redirect, url_for, flash, session)
from werkzeug.utils import secure_filename
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from twilio.rest import Client as TwilioClient
from twilio.twiml.messaging_response import MessagingResponse
from supabase import create_client
import openpyxl
from dotenv import load_dotenv

load_dotenv()

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
SECRET_KEY          = os.getenv("SECRET_KEY", "change-me-in-production")
ADMIN_USER          = os.getenv("ADMIN_USER", "admin")
ADMIN_PASS_HASH     = generate_password_hash(os.getenv("ADMIN_PASSWORD", "admin123"))

SUPABASE_URL        = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY        = os.getenv("SUPABASE_KEY", "")

TWILIO_SID          = os.getenv("TWILIO_ACCOUNT_SID", "")
TWILIO_TOKEN        = os.getenv("TWILIO_AUTH_TOKEN", "")
TWILIO_WA_FROM      = os.getenv("TWILIO_WHATSAPP_FROM", "whatsapp:+14155238886")
WA_BROADCAST_LIST   = [n.strip() for n in os.getenv("WA_BROADCAST_LIST", "").split(",") if n.strip()]

GMAIL_USER          = os.getenv("GMAIL_USER", "")
GMAIL_APP_PASS      = os.getenv("GMAIL_APP_PASSWORD", "")
GMAIL_IMAP_HOST     = "imap.gmail.com"
GMAIL_IMAP_PORT     = 993

UPLOAD_FOLDER       = os.getenv("UPLOAD_FOLDER", "uploads")
ALLOWED_EXT         = {"xlsx","xls","pdf","docx","doc","txt","csv","eml",
                       "jpg","jpeg","png","gif","webp","bmp","tiff","msg"}
MAX_UPLOAD_MB       = int(os.getenv("MAX_UPLOAD_MB", 32))

# Public-facing base URL (no trailing slash). Set this when deploying behind a
# reverse proxy or to a custom domain so the Outlook add-in manifest is correct.
# Falls back to the inbound request's Host header at runtime when not set.
APP_URL             = os.getenv("APP_URL", "").rstrip("/")

BRIEFING_HOUR       = int(os.getenv("BRIEFING_HOUR", 8))
BRIEFING_MINUTE     = int(os.getenv("BRIEFING_MINUTE", 0))
TIMEZONE            = os.getenv("TIMEZONE", "Australia/Melbourne")

# ── OpenAI ────────────────────────────────────────────────────────────────────
OPENAI_API_KEY      = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL        = os.getenv("OPENAI_MODEL", "gpt-4o")
EMBEDDING_MODEL     = os.getenv("EMBEDDING_MODEL", "text-embedding-3-small")
EMBEDDING_DIMS      = 1536

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ── Flask app ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
# Trust X-Forwarded-For / X-Forwarded-Proto from one proxy layer (Render, Fly, etc.)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

def get_base_url() -> str:
    """Return the public-facing base URL with no trailing slash.
    Prefers APP_URL env var; falls back to the inbound request's scheme+host."""
    if APP_URL:
        return APP_URL
    return request.url_root.rstrip("/")

# ── Supabase ──────────────────────────────────────────────────────────────────
_sb = None
def get_sb():
    global _sb
    if _sb is None and SUPABASE_URL and SUPABASE_KEY:
        _sb = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _sb

def sb_select(table, filters=None, order=None, limit=None):
    sb = get_sb()
    if sb is None:
        return []
    try:
        q = sb.table(table).select("*")
        if filters:
            for col, val in filters.items():
                q = q.eq(col, val)
        if order:
            q = q.order(order, desc=True)
        if limit:
            q = q.limit(limit)
        return q.execute().data or []
    except Exception as e:
        log.error("sb_select %s: %s", table, e)
        return []

def sb_insert(table, data):
    sb = get_sb()
    if sb is None:
        return None
    try:
        return sb.table(table).insert(data).execute().data
    except Exception as e:
        log.error("sb_insert %s: %s", table, e)
        return None

def sb_update(table, match, data):
    sb = get_sb()
    if sb is None:
        return None
    try:
        q = sb.table(table).update(data)
        for col, val in match.items():
            q = q.eq(col, val)
        return q.execute().data
    except Exception as e:
        log.error("sb_update %s: %s", table, e)
        return None

def sb_delete(table, row_id):
    sb = get_sb()
    if sb is None:
        return None
    try:
        return sb.table(table).delete().eq("id", row_id).execute().data
    except Exception as e:
        log.error("sb_delete %s: %s", table, e)
        return None

def sb_count(table, filters=None):
    rows = sb_select(table, filters)
    return len(rows)

def init_schema():
    """Create tables if they don't exist via Supabase RPC (run once)."""
    sb = get_sb()
    if sb is None:
        log.warning("Supabase not configured — schema init skipped.")
        return
    ddl = """
    create table if not exists properties (
        id bigserial primary key,
        address text not null,
        suburb text default 'West Melbourne',
        property_type text default 'Warehouse',
        size_sqm numeric,
        land_sqm numeric,
        asking_price numeric,
        asking_rent_pa numeric,
        status text default 'Available',
        agent_name text,
        agent_phone text,
        agent_email text,
        year_built int,
        zoning text,
        notes text,
        source text,
        created_at timestamptz default now(),
        updated_at timestamptz default now()
    );
    create table if not exists contacts (
        id bigserial primary key,
        name text not null,
        company text,
        phone text,
        email text,
        contact_type text default 'Agent',
        notes text,
        whatsapp_opt_in bool default false,
        created_at timestamptz default now()
    );
    create table if not exists inquiries (
        id bigserial primary key,
        contact_name text,
        contact_phone text,
        contact_email text,
        property_id bigint references properties(id),
        source text default 'Web',
        message text,
        status text default 'New',
        created_at timestamptz default now()
    );
    create table if not exists deals (
        id bigserial primary key,
        property_id bigint references properties(id),
        deal_type text default 'Sale',
        price numeric,
        settlement_date date,
        buyer_name text,
        seller_name text,
        agent_name text,
        status text default 'Negotiation',
        notes text,
        created_at timestamptz default now()
    );
    create table if not exists briefings (
        id bigserial primary key,
        briefing_type text default 'Daily',
        content text,
        sent_to text,
        channel text default 'WhatsApp',
        created_at timestamptz default now()
    );
    create table if not exists email_logs (
        id              bigserial primary key,
        sender          text,
        sender_email    text,
        sender_name     text,
        subject         text,
        body            text,
        message_id      text,
        property_id     bigint,
        contact_id      bigint,
        source          text default 'gmail_imap',
        processed       bool default false,
        reply_status    text default 'pending',
        replied_at      timestamptz,
        action_items    jsonb default '[]',
        ai_summary      text,
        ai_priority     text default 'medium',
        requires_action bool default false,
        flagged_unanswered bool default false,
        save_triggered  bool default false,
        received_at     timestamptz default now()
    );
    -- Partial unique index prevents re-processing the same email
    create unique index if not exists email_logs_message_id_idx
        on email_logs(message_id) where message_id is not null;
    -- Backfill new columns on existing rows
    alter table email_logs add column if not exists sender_email    text;
    alter table email_logs add column if not exists sender_name     text;
    alter table email_logs add column if not exists message_id      text;
    alter table email_logs add column if not exists contact_id      bigint;
    alter table email_logs add column if not exists source          text default 'gmail_imap';
    alter table email_logs add column if not exists reply_status    text default 'pending';
    alter table email_logs add column if not exists replied_at      timestamptz;
    alter table email_logs add column if not exists action_items    jsonb default '[]';
    alter table email_logs add column if not exists ai_summary      text;
    alter table email_logs add column if not exists ai_priority     text default 'medium';
    alter table email_logs add column if not exists requires_action bool default false;
    alter table email_logs add column if not exists flagged_unanswered bool default false;
    alter table email_logs add column if not exists save_triggered  bool default false;
    create table if not exists whatsapp_logs (
        id bigserial primary key,
        direction text default 'Inbound',
        from_number text,
        to_number text,
        body text,
        created_at timestamptz default now()
    );

    -- West Melbourne domain tables
    create table if not exists vacancies (
        id              bigserial primary key,
        address         text,
        suburb          text default 'West Melbourne',
        size_sqm        numeric,
        vacating_tenant text,
        available_date  date,
        owner           text,
        agent           text,
        notes           text,
        document_source text,
        property_id     bigint references properties(id),
        document_id     bigint,
        created_at      timestamptz default now()
    );

    create table if not exists requirements (
        id                  bigserial primary key,
        company             text,
        size_min_sqm        numeric,
        size_max_sqm        numeric,
        preferred_location  text default 'West Melbourne',
        region              text,
        rating              text,
        agent               text,
        notes               text,
        document_source     text,
        contact_id          bigint references contacts(id),
        document_id         bigint,
        created_at          timestamptz default now()
    );
    alter table requirements add column if not exists region text;

    create table if not exists market_data (
        id                  bigserial primary key,
        address             text,
        suburb              text default 'West Melbourne',
        size_sqm            numeric,
        deal_type           text default 'Lease',
        tenant              text,
        landlord            text,
        lease_term_years    numeric,
        commencement_date   date,
        expiry_date         date,
        rent_pa             numeric,
        rent_psm            numeric,
        incentive_months    numeric,
        outgoings_psm       numeric,
        document_source     text,
        property_id         bigint references properties(id),
        document_id         bigint,
        created_at          timestamptz default now()
    );

    -- Extra columns for properties (asset register fields)
    alter table properties add column if not exists occupier    text;
    alter table properties add column if not exists landlord    text;
    alter table properties add column if not exists grade       text;
    alter table properties add column if not exists lease_expiry date;

    -- Extra columns for contacts
    alter table contacts add column if not exists last_contacted_at  timestamptz;
    alter table contacts add column if not exists lead_status        text default 'Warm';

    -- Extra columns for calendar_events (post-meeting AI processing)
    alter table calendar_events add column if not exists contact_id          bigint;
    alter table calendar_events add column if not exists ai_action_items     jsonb default '[]';
    alter table calendar_events add column if not exists follow_up_draft     text;
    alter table calendar_events add column if not exists notes_processed_at  timestamptz;

    -- Vacancy-requirement matches
    create table if not exists vacancy_matches (
        id              bigserial primary key,
        vacancy_id      bigint references vacancies(id),
        requirement_id  bigint references requirements(id),
        score           numeric default 0,
        alerted         bool default false,
        created_at      timestamptz default now()
    );

    -- Extra columns for deals (lease deal fields)
    alter table deals add column if not exists address          text;
    alter table deals add column if not exists suburb           text;
    alter table deals add column if not exists size_sqm         numeric;
    alter table deals add column if not exists rent_pa          numeric;
    alter table deals add column if not exists tenant_name      text;
    alter table deals add column if not exists landlord_name    text;
    alter table deals add column if not exists term_years       numeric;
    alter table deals add column if not exists commencement_date date;

    -- Call logs (Android CSV/JSON export)
    create table if not exists call_logs (
        id              bigserial primary key,
        call_date       timestamptz,
        duration_sec    int,
        number          text,
        direction       text default 'Unknown',
        contact_name    text,
        contact_id      bigint references contacts(id),
        inquiry_id      bigint references inquiries(id),
        notes           text,
        source_file     text,
        created_at      timestamptz default now()
    );

    -- Call recordings (Whisper-transcribed audio)
    create table if not exists call_recordings (
        id              bigserial primary key,
        filename        text,
        duration_sec    int,
        contact_name    text,
        contact_id      bigint references contacts(id),
        transcript      text,
        ai_summary      text,
        ai_action_items jsonb default '[]',
        follow_up_draft text,
        key_facts       jsonb default '{}',
        style_learned   bool default false,
        document_id     bigint,
        created_at      timestamptz default now()
    );

    -- Calendar events (ICS import)
    create table if not exists calendar_events (
        id                  bigserial primary key,
        uid                 text unique,
        title               text,
        start_dt            timestamptz,
        end_dt              timestamptz,
        location            text,
        description         text,
        is_property_related bool default false,
        brief_sent          bool default false,
        post_meeting_notes  text,
        created_at          timestamptz default now()
    );

    -- Fee schedules (institutional client rate cards)
    create table if not exists fee_schedules (
        id          bigserial primary key,
        client_name text not null,
        deal_type   text default 'Lease',
        fee_pct     numeric,
        flat_fee    numeric,
        min_fee     numeric,
        notes       text,
        created_at  timestamptz default now()
    );

    -- Fee records (calculated on closed deals)
    create table if not exists fees (
        id              bigserial primary key,
        deal_id         bigint references deals(id),
        property_id     bigint references properties(id),
        client_name     text,
        landlord        text,
        deal_type       text default 'Lease',
        gross_value     numeric,
        fee_pct         numeric,
        fee_amount      numeric,
        invoice_status  text default 'Pending',
        paid_date       date,
        notes           text,
        created_at      timestamptz default now()
    );

    -- Style profile (Michael's communication patterns from transcriptions)
    create table if not exists style_profile (
        id                  bigserial primary key,
        sample_type         text default 'call',
        raw_text            text,
        key_phrases         jsonb default '[]',
        communication_notes text,
        created_at          timestamptz default now()
    );
    """
    ddl_ai = """
    create extension if not exists vector;

    create table if not exists documents (
        id            bigserial primary key,
        filename      text,
        file_type     text,
        raw_text      text,
        ai_classification  text default 'unknown',
        ai_summary    text,
        ai_confidence numeric default 0,
        ai_urgency    text default 'low',
        action_items  jsonb default '[]',
        key_facts     jsonb default '{}',
        extracted_properties    int default 0,
        extracted_contacts      int default 0,
        extracted_inquiries     int default 0,
        extracted_deals         int default 0,
        extracted_vacancies     int default 0,
        extracted_requirements  int default 0,
        extracted_market_data   int default 0,
        mentioned_companies     jsonb default '[]',
        linked_property_ids     jsonb default '[]',
        linked_contact_ids      jsonb default '[]',
        linked_document_ids     jsonb default '[]',
        embedding     vector(1536),
        processing_status text default 'pending',
        error_message text,
        created_at    timestamptz default now()
    );
    -- backfill new columns on existing documents rows
    alter table documents add column if not exists extracted_vacancies    int default 0;
    alter table documents add column if not exists extracted_requirements int default 0;
    alter table documents add column if not exists extracted_market_data  int default 0;
    alter table documents add column if not exists mentioned_companies    jsonb default '[]';
    alter table documents add column if not exists linked_property_ids    jsonb default '[]';
    alter table documents add column if not exists linked_contact_ids     jsonb default '[]';
    alter table documents add column if not exists linked_document_ids    jsonb default '[]';

    create index if not exists documents_embedding_idx
        on documents using ivfflat (embedding vector_cosine_ops) with (lists = 100);

    create or replace function match_documents(
        query_embedding vector(1536),
        match_threshold float default 0.3,
        match_count     int   default 10
    )
    returns table (
        id                bigint,
        filename          text,
        ai_classification text,
        ai_summary        text,
        similarity        float
    )
    language sql stable as $$
        select id, filename, ai_classification, ai_summary,
               1 - (embedding <=> query_embedding) as similarity
        from   documents
        where  embedding is not null
          and  1 - (embedding <=> query_embedding) > match_threshold
        order  by embedding <=> query_embedding
        limit  match_count;
    $$;
    """
    for label, sql in [("core tables", ddl), ("AI/vector tables", ddl_ai)]:
        try:
            sb.rpc("exec_sql", {"sql": sql}).execute()
            log.info("Schema OK: %s", label)
        except Exception as e:
            log.warning("Schema init (%s) via RPC unavailable (%s) — run DDL manually.", label, e)

# ── Twilio / WhatsApp ─────────────────────────────────────────────────────────
def get_twilio():
    if TWILIO_SID and TWILIO_TOKEN:
        return TwilioClient(TWILIO_SID, TWILIO_TOKEN)
    return None

def send_whatsapp(to: str, body: str) -> bool:
    client = get_twilio()
    if not client:
        log.warning("Twilio not configured.")
        return False
    try:
        to_wa = to if to.startswith("whatsapp:") else f"whatsapp:{to}"
        client.messages.create(body=body, from_=TWILIO_WA_FROM, to=to_wa)
        sb_insert("whatsapp_logs", {"direction": "Outbound", "from_number": TWILIO_WA_FROM,
                                     "to_number": to_wa, "body": body})
        return True
    except Exception as e:
        log.error("send_whatsapp: %s", e)
        return False

def broadcast_whatsapp(body: str):
    sent, failed = 0, 0
    for num in WA_BROADCAST_LIST:
        if send_whatsapp(num, body):
            sent += 1
        else:
            failed += 1
    return sent, failed

def handle_inbound_whatsapp(from_num: str, body: str) -> str:
    sb_insert("whatsapp_logs", {"direction": "Inbound", "from_number": from_num,
                                 "to_number": TWILIO_WA_FROM, "body": body})
    cmd = body.strip().upper()
    if cmd in ("HI", "HELLO", "HELP", "START"):
        return ("*West Melbourne Industrial Property Agent* 🏭\n\n"
                "Commands:\n"
                "• *LIST* — available properties\n"
                "• *BRIEF* — today's market brief\n"
                "• *DEALS* — recent transactions\n"
                "• *STATS* — market statistics\n"
                "• *CONTACT* — agent contact info\n"
                "Reply with any of the above keywords.")
    if cmd == "LIST":
        props = sb_select("properties", filters={"status": "Available"}, limit=5)
        if not props:
            return "No available properties at this time. Check back soon."
        lines = ["*Available Properties — West Melbourne*\n"]
        for p in props:
            price = f"${p.get('asking_rent_pa',0):,.0f}/pa" if p.get("asking_rent_pa") else \
                    (f"${p.get('asking_price',0):,.0f}" if p.get("asking_price") else "POA")
            lines.append(f"📍 {p.get('address','N/A')}\n"
                         f"   {p.get('property_type','')} | {p.get('size_sqm','')} sqm | {price}\n")
        return "\n".join(lines)
    if cmd == "BRIEF":
        briefs = sb_select("briefings", order="created_at", limit=1)
        if briefs:
            return briefs[0].get("content", "No briefing available.")
        return build_market_brief()
    if cmd == "DEALS":
        deals = sb_select("deals", order="created_at", limit=5)
        if not deals:
            return "No recent deals recorded."
        lines = ["*Recent Transactions*\n"]
        for d in deals:
            lines.append(f"• {d.get('deal_type','')} — {d.get('status','')}\n"
                         f"  ${d.get('price',0):,.0f} | {d.get('agent_name','')}\n")
        return "\n".join(lines)
    if cmd == "STATS":
        total      = sb_count("properties")
        available  = sb_count("properties", {"status": "Available"})
        leased     = sb_count("properties", {"status": "Leased"})
        sold       = sb_count("properties", {"status": "Sold"})
        return (f"*Market Statistics — West Melbourne Industrial*\n\n"
                f"Total Listings: {total}\n"
                f"✅ Available: {available}\n"
                f"🔵 Leased: {leased}\n"
                f"🟣 Sold: {sold}\n"
                f"Vacancy Rate: {round(available/total*100 if total else 0, 1)}%")
    if cmd == "CONTACT":
        return ("*Agent Contact*\n\n"
                "West Melbourne Industrial Property\n"
                "📞 Contact your agent directly.\n"
                "🌐 Reply HELP for all commands.")
    return ("Sorry, I didn't understand that.\nReply *HELP* to see available commands.")

# ── Gmail IMAP ────────────────────────────────────────────────────────────────
def _extract_email_address(header: str) -> tuple:
    """Return (display_name, email_address) from a From/Reply-To header string."""
    m = re.match(r"^(.*?)\s*<([^>]+)>", header.strip())
    if m:
        return m.group(1).strip(' "\''), m.group(2).strip()
    m2 = re.search(r"[\w.+\-]+@[\w.\-]+", header)
    if m2:
        return header.split("@")[0].strip(), m2.group(0)
    return header.strip(), ""


def analyse_email_with_ai(subject: str, body: str, sender_name: str = "") -> dict:
    """GPT-4o triage: priority, action items, inquiry detection, reply requirement."""
    client = get_openai()
    if not client:
        # Heuristic fallback when OpenAI not configured
        urgent_kw = ["urgent","asap","immediately","today","deadline","critical","offer"]
        inq_kw    = ["lease","rent","warehouse","factory","available","inspect","sqm","m2",
                     "interested","require","looking","enquiry","inquiry"]
        text = (subject + " " + body).lower()
        return {
            "summary":       subject[:120],
            "priority":      "high" if any(k in text for k in urgent_kw) else "medium",
            "action_items":  [],
            "requires_reply": any(k in text for k in ["?","please","could you","can you"]),
            "is_inquiry":    any(k in text for k in inq_kw),
            "not_required":  any(k in text for k in ["unsubscribe","no-reply","noreply",
                                                      "do not reply","auto-reply"]),
            "contact_update": None,
        }
    try:
        resp = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role": "system", "content":
                "You are an email triage assistant for Michael, a West Melbourne industrial "
                "property agent (State Manager). Analyse this email and return JSON only."},
               {"role": "user", "content":
                f"From: {sender_name}\nSubject: {subject}\n\n{body[:4000]}\n\n"
                "Return JSON:\n"
                "{\n"
                "  \"summary\": \"1-2 sentence summary\",\n"
                "  \"priority\": \"critical|high|medium|low\",\n"
                "  \"action_items\": [\"string\"],\n"
                "  \"requires_reply\": true,\n"
                "  \"is_inquiry\": true,\n"
                "  \"inquiry_type\": \"lease|sale|inspection|market|general|none\",\n"
                "  \"properties_mentioned\": [\"address or description\"],\n"
                "  \"not_required\": false,\n"
                "  \"contact_update\": \"Hot|Warm|Cold|null — updated lead status if known contact\"\n"
                "}\n"
                "Priority rules: critical=deadline/offer/legal, high=active deal/inquiry, "
                "medium=general business, low=newsletter/FYI/auto. "
                "not_required=true for newsletters, auto-replies, no-reply senders."}],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=600,
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        log.error("analyse_email_with_ai: %s", e)
        return {"summary": subject, "priority": "medium", "action_items": [],
                "requires_reply": True, "is_inquiry": False, "not_required": False,
                "contact_update": None}


def process_email_full(sender: str, subject: str, body: str,
                       message_id: str = None, source: str = "gmail_imap",
                       received_at: str = None) -> dict | None:
    """
    Unified email processing pipeline.
    1. Dedup by message_id
    2. Extract sender info, match contact
    3. AI triage (priority, action items, inquiry flag)
    4. SAVE trigger → AI document intake
    5. Create/update inquiry record
    6. Store to email_logs
    Returns the inserted email_log row dict, or None if duplicate.
    """
    sb = get_sb()

    # ── 1. Dedup ──────────────────────────────────────────────────────────────
    if message_id and sb:
        try:
            existing = sb.table("email_logs").select("id").eq("message_id", message_id).execute().data
            if existing:
                log.debug("Email dedup: %s already processed", message_id)
                return None
        except Exception:
            pass

    # ── 2. Sender info ────────────────────────────────────────────────────────
    sender_name, sender_email = _extract_email_address(sender)
    is_no_reply = any(x in (sender_email + subject).lower()
                      for x in ["no-reply", "noreply", "do-not-reply", "donotreply",
                                 "mailer-daemon", "postmaster", "unsubscribe"])
    save_triggered = "SAVE" in subject.upper()

    # ── 3. Find or create contact ─────────────────────────────────────────────
    contact_id = None
    contact    = None
    if sender_email and not is_no_reply and sb:
        try:
            matches = sb.table("contacts").select("*").ilike("email", sender_email).execute().data
            if matches:
                contact   = matches[0]
                contact_id = contact["id"]
            elif sender_name:
                # Try matching by name
                name_matches = sb.table("contacts").select("*").ilike("name", f"%{sender_name}%").execute().data
                if name_matches:
                    contact    = name_matches[0]
                    contact_id = contact["id"]
        except Exception as e:
            log.warning("Contact lookup: %s", e)

    # ── 4. AI triage ──────────────────────────────────────────────────────────
    analysis = analyse_email_with_ai(subject, body, sender_name)
    priority        = analysis.get("priority", "medium")
    action_items    = analysis.get("action_items") or []
    requires_reply  = bool(analysis.get("requires_reply", True))
    is_inquiry      = bool(analysis.get("is_inquiry", False))
    not_required    = bool(analysis.get("not_required", False)) or is_no_reply
    ai_summary      = analysis.get("summary", subject)
    reply_status    = "not_required" if not_required else "pending"

    # ── 5. SAVE trigger → AI document intake ──────────────────────────────────
    if save_triggered:
        priority = "critical"
        requires_reply = True
        # write body to temp file and run through process_file_universal
        try:
            tmp_name = f"save_email_{re.sub(r'[^a-z0-9]', '_', subject[:30].lower())}.txt"
            tmp_path = os.path.join(UPLOAD_FOLDER, tmp_name)
            with open(tmp_path, "w", encoding="utf-8") as fh:
                fh.write(f"From: {sender}\nSubject: {subject}\n\n{body}")
            threading.Thread(
                target=lambda: process_file_universal(tmp_path, tmp_name),
                daemon=True
            ).start()
            log.info("SAVE trigger: queued AI intake for '%s'", subject)
        except Exception as e:
            log.error("SAVE trigger AI intake: %s", e)

    # ── 6. Create inquiry if needed ───────────────────────────────────────────
    property_id = _guess_property_from_text(subject + " " + body)
    if is_inquiry and sender_email and not is_no_reply:
        sb_insert("inquiries", {
            "contact_name":  sender_name or sender_email,
            "contact_email": sender_email,
            "source":        "Email",
            "message":       f"{subject}\n\n{body[:800]}",
            "status":        "New",
            "property_id":   property_id,
        })

    # ── 7. Update contact lead status if AI suggests it ───────────────────────
    if contact_id and analysis.get("contact_update") in ("Hot", "Warm", "Cold"):
        open_inqs = sb_select("inquiries", {"contact_name": sender_name})
        status_map = {"Hot": "New", "Warm": "New", "Cold": "Contacted"}
        for inq in open_inqs[:1]:
            sb_update("inquiries", inq["id"], {"status": status_map[analysis["contact_update"]]})

    # ── 8. Store email log ────────────────────────────────────────────────────
    row = {
        "sender":           sender,
        "sender_email":     sender_email,
        "sender_name":      sender_name,
        "subject":          subject,
        "body":             body[:8000],
        "message_id":       message_id,
        "property_id":      property_id,
        "contact_id":       contact_id,
        "source":           source,
        "processed":        True,
        "reply_status":     reply_status,
        "action_items":     json.dumps(action_items),
        "ai_summary":       ai_summary,
        "ai_priority":      priority,
        "requires_action":  requires_reply and not not_required,
        "flagged_unanswered": False,
        "save_triggered":   save_triggered,
    }
    if received_at:
        row["received_at"] = received_at
    inserted = sb_insert("email_logs", row)
    result = inserted[0] if inserted else row
    log.info("Email processed: [%s] %s → priority=%s save=%s",
             source, subject[:60], priority, save_triggered)
    return result


_last_unanswered_alert: datetime | None = None

def flag_unanswered_emails():
    """
    Hourly job: mark emails >24h old with no reply as flagged_unanswered.
    Sends one WhatsApp alert per day if there are pending unanswered emails.
    """
    global _last_unanswered_alert
    sb = get_sb()
    if sb is None:
        return
    cutoff = (datetime.now(timezone.utc) - timedelta(hours=24)).isoformat()
    try:
        rows = (sb.table("email_logs")
                  .select("id,sender_name,subject,ai_priority")
                  .eq("reply_status", "pending")
                  .eq("requires_action", True)
                  .eq("flagged_unanswered", False)
                  .lt("received_at", cutoff)
                  .execute().data or [])
    except Exception as e:
        log.error("flag_unanswered_emails query: %s", e)
        return

    if not rows:
        return

    ids = [r["id"] for r in rows]
    try:
        sb.table("email_logs").update({"flagged_unanswered": True}).in_("id", ids).execute()
    except Exception as e:
        log.error("flag_unanswered_emails update: %s", e)

    # WhatsApp alert — at most once per calendar day
    now = datetime.now(timezone.utc)
    if _last_unanswered_alert is None or _last_unanswered_alert.date() < now.date():
        critical = [r for r in rows if r.get("ai_priority") == "critical"]
        high     = [r for r in rows if r.get("ai_priority") == "high"]
        lines    = [f"📬 *{len(rows)} unanswered email{'s' if len(rows)!=1 else ''} (>24h)*"]
        if critical:
            lines.append(f"🔴 Critical: {critical[0].get('subject','')[:50]}")
        if high:
            lines.append(f"🟠 High: {high[0].get('subject','')[:50]}")
        if len(rows) > 2:
            lines.append(f"…and {len(rows)-2} more. Check /email-log for full list.")
        broadcast_whatsapp("\n".join(lines))
        _last_unanswered_alert = now
        log.info("Unanswered alert sent: %d emails", len(rows))


def get_email_priorities(query: str = "") -> dict:
    """
    Natural-language query over email_logs.
    Returns {'emails': [...], 'answer': str, 'filter_used': str}
    """
    emails = sb_select("email_logs", order="received_at", limit=200)
    if not emails:
        return {"emails": [], "answer": "No emails in the database yet.", "filter_used": "none"}

    q_lower = query.lower()

    # keyword dispatch
    if any(w in q_lower for w in ["unanswered","haven't responded","no reply","not replied","pending"]):
        filtered = [e for e in emails if e.get("reply_status") == "pending"
                    and e.get("requires_action")]
        filtered.sort(key=lambda e: (
            {"critical":0,"high":1,"medium":2,"low":3}.get(e.get("ai_priority","medium"),2),
            e.get("received_at","")
        ))
        answer = (f"{len(filtered)} email{'s' if len(filtered)!=1 else ''} awaiting a reply."
                  if filtered else "You're all caught up — no unanswered emails.")
        return {"emails": filtered[:20], "answer": answer, "filter_used": "unanswered"}

    if any(w in q_lower for w in ["priority","priorities","important","urgent","today","focus"]):
        filtered = [e for e in emails if e.get("ai_priority") in ("critical","high")]
        filtered.sort(key=lambda e: {"critical":0,"high":1}.get(e.get("ai_priority","high"),1))
        answer = (f"{len(filtered)} high-priority email{'s' if len(filtered)!=1 else ''} need attention."
                  if filtered else "No critical or high-priority emails right now.")
        return {"emails": filtered[:20], "answer": answer, "filter_used": "high_priority"}

    if any(w in q_lower for w in ["flagged","overdue","24","24h","follow up"]):
        filtered = [e for e in emails if e.get("flagged_unanswered")]
        answer = (f"{len(filtered)} email{'s' if len(filtered)!=1 else ''} flagged as overdue (>24h, no reply)."
                  if filtered else "No overdue emails.")
        return {"emails": filtered[:20], "answer": answer, "filter_used": "overdue"}

    if any(w in q_lower for w in ["save","saved","captured","document"]):
        filtered = [e for e in emails if e.get("save_triggered")]
        answer = f"{len(filtered)} email{'s' if len(filtered)!=1 else ''} triggered SAVE processing."
        return {"emails": filtered[:20], "answer": answer, "filter_used": "save_triggered"}

    # "from [name]" pattern
    from_match = re.search(r"from\s+([a-zA-Z\s]{2,30})", q_lower)
    if from_match:
        name = from_match.group(1).strip()
        filtered = [e for e in emails
                    if name in (e.get("sender_name") or e.get("sender") or "").lower()]
        answer = (f"{len(filtered)} email{'s' if len(filtered)!=1 else ''} from '{name}'."
                  if filtered else f"No emails found from '{name}'.")
        return {"emails": filtered[:20], "answer": answer, "filter_used": f"from:{name}"}

    # default: all emails sorted by priority
    filtered = sorted(emails, key=lambda e: (
        {"critical":0,"high":1,"medium":2,"low":3}.get(e.get("ai_priority","medium"),2),
        "" if not e.get("flagged_unanswered") else "z"
    ))
    answer = f"{len(filtered)} total emails. Use the query bar to filter."
    return {"emails": filtered[:50], "answer": answer, "filter_used": "all"}


def decode_mime_words(s: str) -> str:
    if not s:
        return ""
    parts = decode_header(s)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(str(part))
    return " ".join(result)

def fetch_gmail_emails(max_emails=20):
    if not GMAIL_USER or not GMAIL_APP_PASS:
        log.warning("Gmail credentials not configured.")
        return 0
    new_count = 0
    try:
        conn = imaplib.IMAP4_SSL(GMAIL_IMAP_HOST, GMAIL_IMAP_PORT)
        conn.login(GMAIL_USER, GMAIL_APP_PASS)
        conn.select("INBOX")
        _, msg_ids = conn.search(None, "UNSEEN")
        ids = (msg_ids[0].split() or [])[-max_emails:]
        for mid in ids:
            try:
                _, data = conn.fetch(mid, "(RFC822)")
                raw = data[0][1]
                msg = email.message_from_bytes(raw)
                sender  = decode_mime_words(msg.get("From", ""))
                subject = decode_mime_words(msg.get("Subject", ""))
                message_id = (msg.get("Message-ID") or "").strip()

                # Extract plain-text body
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        ct = part.get_content_type()
                        cd = str(part.get("Content-Disposition", ""))
                        if ct == "text/plain" and "attachment" not in cd:
                            body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                            break
                else:
                    body = msg.get_payload(decode=True).decode("utf-8", errors="replace")

                # Parse received date
                date_str = msg.get("Date", "")
                received_at = None
                if date_str:
                    try:
                        from email.utils import parsedate_to_datetime
                        received_at = parsedate_to_datetime(date_str).isoformat()
                    except Exception:
                        pass

                result = process_email_full(
                    sender=sender,
                    subject=subject,
                    body=body,
                    message_id=message_id or None,
                    source="gmail_imap",
                    received_at=received_at,
                )
                if result is not None:
                    new_count += 1
            except Exception as e:
                log.error("fetch_gmail_emails processing msg %s: %s", mid, e)

        conn.logout()
        log.info("Gmail: processed %d new emails.", new_count)
    except Exception as e:
        log.error("fetch_gmail_emails: %s", e)
    return new_count

def _guess_property_from_text(text: str):
    """Try to match email text to a property by address keywords."""
    props = sb_select("properties")
    for p in props:
        addr = (p.get("address") or "").lower()
        if addr and len(addr) > 5 and addr[:10] in text.lower():
            return p.get("id")
    return None

# ── Excel processing ──────────────────────────────────────────────────────────
PROPERTY_COLUMN_MAP = {
    "address": ["address", "property address", "street", "location"],
    "suburb": ["suburb", "city", "area"],
    "property_type": ["type", "property type", "asset type"],
    "size_sqm": ["size", "sqm", "area sqm", "building area", "floor area", "gfa"],
    "land_sqm": ["land", "land area", "land sqm", "site area"],
    "asking_price": ["price", "asking price", "sale price", "list price"],
    "asking_rent_pa": ["rent", "rent pa", "annual rent", "asking rent", "lease pa"],
    "status": ["status", "availability"],
    "agent_name": ["agent", "agent name", "listing agent"],
    "agent_phone": ["phone", "agent phone", "mobile", "contact number"],
    "agent_email": ["email", "agent email", "contact email"],
    "year_built": ["year", "year built", "built"],
    "zoning": ["zone", "zoning", "planning zone"],
    "notes": ["notes", "comments", "remarks", "description"],
}

def _normalise_header(h: str) -> str:
    return re.sub(r"[^a-z0-9 ]", "", str(h).lower()).strip()

def _match_col(header: str, field: str) -> bool:
    norm = _normalise_header(header)
    return any(alias in norm for alias in PROPERTY_COLUMN_MAP.get(field, []))

def process_excel_file(filepath: str) -> dict:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {"imported": 0, "skipped": 0, "errors": ["Empty workbook"]}

    # Find header row (first row with >3 non-None cells)
    header_idx = 0
    for i, row in enumerate(rows):
        if sum(1 for c in row if c is not None) >= 3:
            header_idx = i
            break

    headers = [str(c) if c is not None else "" for c in rows[header_idx]]
    col_map = {}
    for col_idx, h in enumerate(headers):
        for field in PROPERTY_COLUMN_MAP:
            if _match_col(h, field) and field not in col_map.values():
                col_map[col_idx] = field
                break

    imported, skipped, errors = 0, 0, []
    for row_num, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if all(c is None for c in row):
            continue
        record = {}
        for col_idx, field in col_map.items():
            if col_idx < len(row):
                val = row[col_idx]
                if val is not None:
                    record[field] = val
        if not record.get("address"):
            skipped += 1
            continue
        record.setdefault("suburb", "West Melbourne")
        record.setdefault("status", "Available")
        record.setdefault("property_type", "Warehouse")
        record["source"] = "Excel Import"
        # Coerce numerics
        for num_field in ("size_sqm", "land_sqm", "asking_price", "asking_rent_pa", "year_built"):
            if num_field in record:
                try:
                    raw = str(record[num_field]).replace("$", "").replace(",", "").strip()
                    record[num_field] = float(raw) if "." in raw else int(raw)
                except (ValueError, TypeError):
                    record.pop(num_field, None)
        result = sb_insert("properties", record)
        if result:
            imported += 1
        else:
            errors.append(f"Row {row_num}: DB insert failed.")
            skipped += 1

    return {"imported": imported, "skipped": skipped, "errors": errors}

# ── AI Intake Engine ──────────────────────────────────────────────────────────
_oai_client = None

def get_openai():
    global _oai_client
    if _oai_client is None and OPENAI_API_KEY and not OPENAI_API_KEY.startswith("PASTE"):
        from openai import OpenAI
        _oai_client = OpenAI(api_key=OPENAI_API_KEY)
    return _oai_client

INTAKE_SYSTEM_PROMPT = """You are an expert document analyst for MJR West, a commercial real estate agency \
specialising in West Melbourne industrial property.

Classify the document into ONE primary type and extract every piece of data present.
Return ONLY valid JSON — no markdown fences, no commentary, just the JSON object.

DOCUMENT TYPE RULES:
- asset_register    → spreadsheet/list of properties with occupier, landlord, lease details
- vacancy_schedule  → list of available/upcoming-vacant premises
- requirements_listing → tenants/buyers searching for space (size, location, budget)
- deal_tracker      → pipeline of current lease/sale negotiations
- lease_contract    → executed lease or agreement — the full legal document
- sales_contract    → executed sale/purchase contract
- market_report     → market commentary, research, trend data
- inquiry_email     → inbound enquiry from tenant/buyer
- inspection_report → property condition report
- invoice           → financial invoice
- general_correspondence → letters, memos, emails
- unknown           → cannot determine

JSON SCHEMA:
{
  "document_type": "<type from list above>",
  "confidence": 0.85,
  "summary": "2–3 sentence plain-English summary of what this document is and key numbers.",
  "urgency": "low|medium|high|critical",
  "mentioned_companies": ["Exact Pty Ltd name as written", "..."],
  "table_inserts": {

    "properties": [
      {
        "address": "", "suburb": "West Melbourne",
        "property_type": "Warehouse|Factory|Office/Warehouse|Cold Storage|Land|Showroom",
        "size_sqm": null, "land_sqm": null,
        "asking_price": null, "asking_rent_pa": null,
        "status": "Available|Under Offer|Leased|Sold",
        "occupier": "", "landlord": "", "grade": "A|B|C|",
        "lease_expiry": null,
        "agent_name": "", "agent_phone": "", "agent_email": "",
        "year_built": null, "zoning": "", "notes": ""
      }
    ],

    "contacts": [
      {
        "name": "", "company": "", "phone": "", "email": "",
        "contact_type": "Agent|Buyer|Tenant|Owner|Developer|Solicitor|Other",
        "notes": ""
      }
    ],

    "inquiries": [
      {
        "contact_name": "", "contact_phone": "", "contact_email": "",
        "source": "Email|WhatsApp|Phone|Web|Referral",
        "message": "", "status": "New"
      }
    ],

    "deals": [
      {
        "deal_type": "Sale|Lease",
        "address": "", "suburb": "West Melbourne",
        "size_sqm": null, "rent_pa": null, "price": null,
        "tenant_name": "", "landlord_name": "",
        "buyer_name": "", "seller_name": "",
        "agent_name": "", "term_years": null,
        "commencement_date": null, "settlement_date": null,
        "status": "Negotiation|Under Offer|Exchanged|Completed",
        "notes": ""
      }
    ],

    "vacancies": [
      {
        "address": "", "suburb": "West Melbourne",
        "size_sqm": null, "vacating_tenant": "",
        "available_date": null, "owner": "", "agent": "", "notes": ""
      }
    ],

    "requirements": [
      {
        "company": "", "size_min_sqm": null, "size_max_sqm": null,
        "preferred_location": "West Melbourne",
        "region": "W",
        "rating": "Hot|Warm|Cold|",
        "agent": "", "notes": ""
      }
    ],

    "market_data": [
      {
        "address": "", "suburb": "West Melbourne",
        "size_sqm": null, "deal_type": "Lease|Sale",
        "tenant": "", "landlord": "",
        "lease_term_years": null,
        "commencement_date": null, "expiry_date": null,
        "rent_pa": null, "rent_psm": null,
        "incentive_months": null, "outgoings_psm": null,
        "notes": ""
      }
    ]

  },
  "action_items": ["string"],
  "key_facts": {"key": "value"}
}

STRICT EXTRACTION RULES:
1. Populate ONLY arrays where real data was found — empty arrays [] are fine
2. Numeric fields must be numbers or null — NEVER strings
3. Dates must be "YYYY-MM-DD" strings or null — NEVER free text like "June 2025"
4. Do not invent or guess — only extract data explicitly present in the document
5. mentioned_companies: list every company/business name exactly as written
6. For asset_register: create one properties row per property listed; use occupier/landlord/lease_expiry
7. For vacancy_schedule: create one vacancies row per premises; also create a properties row marked Available
8. For requirements_listing: create one requirements row per requirement; also create a contacts row
9. For deal_tracker: create one deals row per deal in the pipeline
10. For lease_contract or sales_contract: create one market_data row with full lease/sale terms; \
    also create a deals row with status=Completed
11. For all document types containing contacts: also populate the contacts array
12. rent_psm = rent_pa / size_sqm when both are present but psm is missing
13. grade must be A, B, C, or empty string — never null
14. REGION FILTER FOR REQUIREMENTS (critical): When processing a requirements_listing or any \
    spreadsheet with a location/region/area column, identify that column. \
    Read the raw region value for every row and store it in the `region` field ALWAYS. \
    Then apply this filter: ONLY include a row in table_inserts.requirements if its region \
    value is "W", "West", "West Melbourne", or any clear West Melbourne indicator. \
    Rows with region "N", "North", "NW", "SE", "South East", "E", "East", or any \
    non-West value must be EXCLUDED from table_inserts.requirements entirely. \
    If the document has no region/location column, include all requirements and set region="W". \
    The `region` field must always contain the raw value from the source document (e.g. "W", \
    "N", "SE") — never leave it null when a region column exists in the source."""


def extract_text_from_file(filepath: str, filename: str) -> tuple:
    """Return (content, is_image). Images return base64 data-URL; others return plain text."""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    # ── Images → base64 for GPT-4o vision ────────────────────────────────────
    if ext in ("jpg", "jpeg", "png", "gif", "webp", "bmp", "tiff"):
        try:
            with open(filepath, "rb") as fh:
                b64 = base64.b64encode(fh.read()).decode()
            mime = {"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png",
                    "gif":"image/gif","webp":"image/webp","bmp":"image/bmp",
                    "tiff":"image/tiff"}.get(ext, "image/jpeg")
            return f"data:{mime};base64,{b64}", True
        except Exception as exc:
            log.error("Image read: %s", exc)
            return "", True

    # ── PDF ───────────────────────────────────────────────────────────────────
    if ext == "pdf":
        try:
            import pypdf
            reader = pypdf.PdfReader(filepath)
            pages = [page.extract_text() or "" for page in reader.pages[:40]]
            text = "\n\n".join(p for p in pages if p.strip())
            return text or "[PDF contained no extractable text]", False
        except Exception as exc:
            log.error("PDF read: %s", exc)
            return f"[PDF read failed: {exc}]", False

    # ── Word ──────────────────────────────────────────────────────────────────
    if ext in ("docx", "doc"):
        try:
            import docx as _docx
            doc = _docx.Document(filepath)
            paras = [p.text for p in doc.paragraphs if p.text.strip()]
            tbl_rows = []
            for tbl in doc.tables:
                for row in tbl.rows:
                    tbl_rows.append(" | ".join(c.text.strip() for c in row.cells))
            return "\n".join(paras + tbl_rows), False
        except Exception as exc:
            log.error("DOCX read: %s", exc)
            return f"[Word read failed: {exc}]", False

    # ── Excel ─────────────────────────────────────────────────────────────────
    if ext in ("xlsx", "xls"):
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            lines = []
            for sheet in wb.worksheets[:5]:
                lines.append(f"[Sheet: {sheet.title}]")
                for row in list(sheet.iter_rows(values_only=True))[:300]:
                    row_str = " | ".join("" if c is None else str(c) for c in row)
                    if row_str.replace("|","").strip():
                        lines.append(row_str)
            return "\n".join(lines), False
        except Exception as exc:
            log.error("Excel read: %s", exc)
            return f"[Excel read failed: {exc}]", False

    # ── Email (.eml) ──────────────────────────────────────────────────────────
    if ext == "eml":
        try:
            with open(filepath, "rb") as fh:
                msg = email.message_from_bytes(fh.read())
            sender  = decode_mime_words(msg.get("From", ""))
            subject = decode_mime_words(msg.get("Subject", ""))
            date    = msg.get("Date", "")
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                        break
            else:
                raw_pay = msg.get_payload(decode=True)
                body = raw_pay.decode("utf-8", errors="replace") if raw_pay else ""
            return f"FROM: {sender}\nDATE: {date}\nSUBJECT: {subject}\n\n{body}", False
        except Exception as exc:
            log.error("EML read: %s", exc)
            return f"[Email read failed: {exc}]", False

    # ── Plain text / CSV / TSV ────────────────────────────────────────────────
    if ext in ("txt", "csv", "tsv", "md"):
        try:
            with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
                return fh.read()[:60000], False
        except Exception as exc:
            return f"[Text read failed: {exc}]", False

    # ── Generic fallback ──────────────────────────────────────────────────────
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read()[:20000], False
    except Exception:
        return f"[Unreadable file type: {ext}]", False


def analyze_with_gpt4o(content: str, filename: str, is_image: bool = False) -> dict:
    """Send content to GPT-4o; return parsed JSON analysis dict."""
    client = get_openai()
    if not client:
        return {"document_type":"unknown","confidence":0,"summary":"OpenAI not configured — add OPENAI_API_KEY to .env",
                "urgency":"low","table_inserts":{},"action_items":[],"key_facts":{}}
    try:
        if is_image:
            messages = [
                {"role":"system","content":INTAKE_SYSTEM_PROMPT},
                {"role":"user","content":[
                    {"type":"text","text":f"Analyse this document: {filename}"},
                    {"type":"image_url","image_url":{"url":content,"detail":"high"}}
                ]}
            ]
        else:
            truncated = content[:120000]   # ~30k tokens
            messages = [
                {"role":"system","content":INTAKE_SYSTEM_PROMPT},
                {"role":"user","content":f"Filename: {filename}\n\nDocument content:\n\n{truncated}"}
            ]
        resp = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=messages,
            response_format={"type":"json_object"},
            temperature=0.1,
            max_tokens=4096
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as exc:
        log.error("GPT-4o analysis: %s", exc)
        return {"document_type":"unknown","confidence":0,"summary":f"Analysis failed: {exc}",
                "urgency":"low","table_inserts":{},"action_items":[],"key_facts":{},
                "error":str(exc)}


def embed_text(text: str) -> list:
    """Generate 1536-dim embedding vector. Returns list[float] or empty list."""
    client = get_openai()
    if not client or not text.strip():
        return []
    try:
        clean = text[:8000].replace("\n", " ")
        resp  = client.embeddings.create(
            model=EMBEDDING_MODEL,
            input=clean,
            dimensions=EMBEDDING_DIMS
        )
        return resp.data[0].embedding
    except Exception as exc:
        log.error("Embedding: %s", exc)
        return []


def _cosine_sim(a: list, b: list) -> float:
    if not a or not b or len(a) != len(b):
        return 0.0
    dot = sum(x * y for x, y in zip(a, b))
    na  = sum(x * x for x in a) ** 0.5
    nb  = sum(x * x for x in b) ** 0.5
    return dot / (na * nb) if na and nb else 0.0


def semantic_search(query: str, threshold: float = 0.25, n: int = 12) -> list:
    """Search the documents table by semantic meaning. Returns ranked list."""
    q_vec = embed_text(query)
    if not q_vec:
        return []
    sb = get_sb()
    if sb is None:
        return []
    # Try Supabase pgvector RPC (fast, server-side)
    try:
        result = sb.rpc("match_documents", {
            "query_embedding": q_vec,
            "match_threshold": threshold,
            "match_count": n
        }).execute()
        if result.data:
            return result.data
    except Exception:
        pass
    # Fallback: pull all embeddings, rank client-side
    try:
        docs = sb.table("documents").select(
            "id,filename,ai_classification,ai_summary,ai_confidence,embedding,created_at"
        ).execute().data or []
        scored = []
        for doc in docs:
            raw_emb = doc.get("embedding")
            if isinstance(raw_emb, str):
                try:
                    raw_emb = json.loads(raw_emb)
                except Exception:
                    continue
            sim = _cosine_sim(q_vec, raw_emb or [])
            if sim >= threshold:
                scored.append({**doc, "similarity": round(sim, 4)})
        return sorted(scored, key=lambda x: x["similarity"], reverse=True)[:n]
    except Exception as exc:
        log.error("semantic_search fallback: %s", exc)
        return []


def _coerce_numeric(d: dict, fields: list):
    """In-place: convert listed fields to float or None."""
    for f in fields:
        v = d.get(f)
        if v not in (None, ""):
            try:
                d[f] = float(v)
            except (ValueError, TypeError):
                d[f] = None
        else:
            d[f] = None

def _coerce_date(d: dict, fields: list):
    """In-place: accept YYYY-MM-DD strings; set invalid/empty to None."""
    for f in fields:
        v = d.get(f)
        if not v:
            d[f] = None
            continue
        v = str(v).strip()
        import re as _re
        if _re.match(r"^\d{4}-\d{2}-\d{2}$", v):
            d[f] = v
        else:
            d[f] = None

def _normalise_company(name: str) -> str:
    """Lowercase, strip Pty Ltd / P/L / Pty. Ltd. punctuation for fuzzy matching."""
    import re as _re
    if not name:
        return ""
    n = name.lower()
    n = _re.sub(r"\b(pty\.?\s*ltd\.?|p/l|pty|ltd|limited|lp|llc|inc\.?|corp\.?)\b", "", n)
    n = _re.sub(r"[^\w\s]", " ", n)
    return " ".join(n.split())


def store_ai_results(analysis: dict, filename: str) -> dict:
    """Write all table_inserts from GPT-4o analysis to Supabase. Returns insert counts."""
    counts = {k: 0 for k in ("properties","contacts","inquiries","deals",
                              "vacancies","requirements","market_data")}
    inserts = analysis.get("table_inserts") or {}

    # ── Properties ─────────────────────────────────────────────────────────────
    for prop in (inserts.get("properties") or []):
        if not (prop.get("address") or "").strip():
            continue
        prop.setdefault("suburb", "West Melbourne")
        prop.setdefault("source", f"AI Intake: {filename}")
        prop.setdefault("status", "Available")
        _coerce_numeric(prop, ["size_sqm","land_sqm","asking_price","asking_rent_pa","year_built"])
        _coerce_date(prop, ["lease_expiry"])
        # grade must be A/B/C or empty
        if prop.get("grade") not in ("A","B","C",""):
            prop["grade"] = ""
        if sb_insert("properties", prop):
            counts["properties"] += 1

    # ── Contacts ───────────────────────────────────────────────────────────────
    for contact in (inserts.get("contacts") or []):
        if not (contact.get("name") or "").strip():
            continue
        if sb_insert("contacts", contact):
            counts["contacts"] += 1

    # ── Inquiries ──────────────────────────────────────────────────────────────
    for inq in (inserts.get("inquiries") or []):
        if not ((inq.get("contact_name") or "") + (inq.get("contact_email") or "")).strip():
            continue
        inq.setdefault("source", "AI Intake")
        inq.setdefault("status", "New")
        if sb_insert("inquiries", inq):
            counts["inquiries"] += 1

    # ── Deals ──────────────────────────────────────────────────────────────────
    for deal in (inserts.get("deals") or []):
        _coerce_numeric(deal, ["price","size_sqm","rent_pa","term_years"])
        _coerce_date(deal, ["commencement_date","settlement_date"])
        deal.setdefault("document_source", filename)
        if sb_insert("deals", deal):
            counts["deals"] += 1

    # ── Vacancies ──────────────────────────────────────────────────────────────
    for vac in (inserts.get("vacancies") or []):
        if not (vac.get("address") or "").strip():
            continue
        vac.setdefault("suburb", "West Melbourne")
        vac["document_source"] = filename
        _coerce_numeric(vac, ["size_sqm"])
        _coerce_date(vac, ["available_date"])
        if sb_insert("vacancies", vac):
            counts["vacancies"] += 1

    # ── Requirements ── W-region only; region field preserved for future routing ─
    _WEST_REGIONS = {"w", "west", "west melbourne", "western", "wm"}
    _NON_WEST     = {"n", "north", "ne", "nw", "se", "south east", "southeast",
                     "e", "east", "s", "south", "sw", "c", "central", "cbd"}
    for req in (inserts.get("requirements") or []):
        if not (req.get("company") or req.get("agent") or "").strip():
            continue
        region_raw = (req.get("region") or "").strip()
        region_key = region_raw.lower()
        # If region column exists and is a known non-West value, skip this record.
        # If region is blank or absent, assume West (no column in source doc).
        if region_key and region_key in _NON_WEST:
            log.debug("Skipping non-West requirement (region=%s): %s", region_raw,
                      req.get("company","?"))
            continue
        # Normalise: store the raw value; default to "W" when absent
        if not region_raw:
            req["region"] = "W"
        req.setdefault("preferred_location", "West Melbourne")
        req["document_source"] = filename
        _coerce_numeric(req, ["size_min_sqm","size_max_sqm"])
        if sb_insert("requirements", req):
            counts["requirements"] += 1

    # ── Market Data (executed leases / sales evidence) ─────────────────────────
    for md in (inserts.get("market_data") or []):
        if not (md.get("address") or md.get("tenant") or "").strip():
            continue
        md.setdefault("suburb", "West Melbourne")
        md.setdefault("deal_type", "Lease")
        md["document_source"] = filename
        _coerce_numeric(md, ["size_sqm","rent_pa","rent_psm","lease_term_years",
                              "incentive_months","outgoings_psm"])
        _coerce_date(md, ["commencement_date","expiry_date"])
        # derive rent_psm if missing
        if md.get("rent_pa") and md.get("size_sqm") and not md.get("rent_psm"):
            try:
                md["rent_psm"] = round(md["rent_pa"] / md["size_sqm"], 2)
            except Exception:
                pass
        if sb_insert("market_data", md):
            counts["market_data"] += 1

    return counts


def link_companies(doc_id: int, mentioned_companies: list) -> dict:
    """
    Cross-reference companies in this document against the full database.
    Returns dict with lists of linked property_ids, contact_ids, document_ids.
    Also back-links the other documents that share companies.
    """
    if not mentioned_companies or doc_id is None:
        return {"property_ids": [], "contact_ids": [], "document_ids": []}

    sb = get_sb()
    if sb is None:
        return {"property_ids": [], "contact_ids": [], "document_ids": []}

    norm_names = {_normalise_company(c): c for c in mentioned_companies if c}

    linked_prop_ids  = []
    linked_cont_ids  = []
    linked_doc_ids   = []

    # ── Match against contacts.company ────────────────────────────────────────
    try:
        all_contacts = sb.table("contacts").select("id,company").execute().data or []
        for ct in all_contacts:
            cname = ct.get("company") or ""
            if _normalise_company(cname) in norm_names:
                linked_cont_ids.append(ct["id"])
    except Exception as e:
        log.warning("link_companies contacts: %s", e)

    # ── Match against properties.occupier + properties.landlord ───────────────
    try:
        all_props = sb.table("properties").select("id,occupier,landlord").execute().data or []
        for pr in all_props:
            for field in ("occupier", "landlord"):
                if _normalise_company(pr.get(field) or "") in norm_names:
                    if pr["id"] not in linked_prop_ids:
                        linked_prop_ids.append(pr["id"])
    except Exception as e:
        log.warning("link_companies properties: %s", e)

    # ── Match against other documents' mentioned_companies ────────────────────
    try:
        other_docs = (sb.table("documents")
                        .select("id,mentioned_companies")
                        .neq("id", doc_id)
                        .execute().data or [])
        for od in other_docs:
            raw = od.get("mentioned_companies") or []
            if isinstance(raw, str):
                try: raw = json.loads(raw)
                except Exception: raw = []
            for oc in raw:
                if _normalise_company(oc) in norm_names:
                    linked_doc_ids.append(od["id"])
                    break
    except Exception as e:
        log.warning("link_companies documents: %s", e)

    # ── Update this document with its linked IDs ───────────────────────────────
    try:
        sb.table("documents").update({
            "linked_property_ids": json.dumps(linked_prop_ids),
            "linked_contact_ids":  json.dumps(linked_cont_ids),
            "linked_document_ids": json.dumps(linked_doc_ids),
        }).eq("id", doc_id).execute()
    except Exception as e:
        log.warning("link_companies update doc: %s", e)

    # ── Back-link the other documents that share at least one company ──────────
    for other_id in linked_doc_ids:
        try:
            od_row = (sb.table("documents").select("linked_document_ids")
                       .eq("id", other_id).execute().data or [{}])[0]
            existing = od_row.get("linked_document_ids") or []
            if isinstance(existing, str):
                try: existing = json.loads(existing)
                except Exception: existing = []
            if doc_id not in existing:
                existing.append(doc_id)
                sb.table("documents").update({
                    "linked_document_ids": json.dumps(existing)
                }).eq("id", other_id).execute()
        except Exception as e:
            log.warning("link_companies back-link %s: %s", other_id, e)

    return {
        "property_ids": linked_prop_ids,
        "contact_ids":  linked_cont_ids,
        "document_ids": linked_doc_ids,
    }


def process_file_universal(filepath: str, filename: str) -> dict:
    """
    Universal AI intake pipeline.
    1. Extract full text/image from any file type
    2. Analyse with GPT-4o  →  structured JSON (7-table schema)
    3. Store records to all relevant tables
    4. Embed full raw text (richer semantic signal than summary alone)
    5. Persist document record + embedding
    6. Cross-link companies across the database
    Returns full result dict.
    """
    result = {
        "filename":          filename,
        "status":            "processing",
        "ai_classification": "unknown",
        "ai_summary":        "",
        "ai_confidence":     0.0,
        "ai_urgency":        "low",
        "counts":            {},
        "action_items":      [],
        "key_facts":         {},
        "mentioned_companies": [],
        "links":             {},
        "error":             None,
        "doc_id":            None,
    }

    # ── 1. Extract ────────────────────────────────────────────────────────────
    raw_content, is_image = extract_text_from_file(filepath, filename)
    if not raw_content:
        result.update(status="error", error="Could not extract content from file.")
        return result

    # ── 2. GPT-4o analysis ────────────────────────────────────────────────────
    analysis = analyze_with_gpt4o(raw_content, filename, is_image)

    result["ai_classification"]  = analysis.get("document_type", "unknown")
    result["ai_summary"]         = analysis.get("summary", "")
    result["ai_confidence"]      = float(analysis.get("confidence") or 0)
    result["ai_urgency"]         = analysis.get("urgency", "low")
    result["action_items"]       = analysis.get("action_items") or []
    result["key_facts"]          = analysis.get("key_facts") or {}
    result["mentioned_companies"] = analysis.get("mentioned_companies") or []

    if analysis.get("error") and result["ai_confidence"] == 0:
        result["error"] = analysis["error"]

    # ── 3. Store extracted records ────────────────────────────────────────────
    counts = store_ai_results(analysis, filename)
    result["counts"] = counts

    # ── 4. Embed full text for maximum semantic coverage ──────────────────────
    if is_image:
        embed_src = f"{filename} {result['ai_summary']} {' '.join(str(v) for v in (result['key_facts'] or {}).values())}"
    else:
        # Use filename + summary + full raw text (trimmed to ~6k tokens)
        embed_src = f"{filename}\n{result['ai_summary']}\n{raw_content[:15000]}"
    embedding = embed_text(embed_src)

    # ── 5. Persist document record ────────────────────────────────────────────
    doc_record = {
        "filename":               filename,
        "file_type":              filename.rsplit(".", 1)[-1].lower() if "." in filename else "unknown",
        "raw_text":               "" if is_image else raw_content[:50000],
        "ai_classification":      result["ai_classification"],
        "ai_summary":             result["ai_summary"],
        "ai_confidence":          result["ai_confidence"],
        "ai_urgency":             result["ai_urgency"],
        "action_items":           json.dumps(result["action_items"]),
        "key_facts":              json.dumps(result["key_facts"]),
        "mentioned_companies":    json.dumps(result["mentioned_companies"]),
        "extracted_properties":   counts.get("properties", 0),
        "extracted_contacts":     counts.get("contacts", 0),
        "extracted_inquiries":    counts.get("inquiries", 0),
        "extracted_deals":        counts.get("deals", 0),
        "extracted_vacancies":    counts.get("vacancies", 0),
        "extracted_requirements": counts.get("requirements", 0),
        "extracted_market_data":  counts.get("market_data", 0),
        "processing_status":      "error" if result["error"] else "complete",
        "error_message":          result["error"],
    }
    if embedding:
        doc_record["embedding"] = embedding

    inserted = sb_insert("documents", doc_record)
    doc_id = inserted[0].get("id") if inserted else None
    result["doc_id"] = doc_id

    # ── 6. Cross-link companies ───────────────────────────────────────────────
    if doc_id and result["mentioned_companies"]:
        links = link_companies(doc_id, result["mentioned_companies"])
        result["links"] = links
        log.info("Company links for doc %s: %d props, %d contacts, %d related docs",
                 doc_id, len(links["property_ids"]),
                 len(links["contact_ids"]), len(links["document_ids"]))

    result["status"] = "error" if result["error"] else "complete"
    log.info("AI intake: %s → %s (conf=%.2f) inserts=%s",
             filename, result["ai_classification"], result["ai_confidence"], counts)
    return result

# ── Outlook capture helpers ───────────────────────────────────────────────────
OUTLOOK_MANIFEST_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="MailApp">
  <Id>mjrwest-agent-addin-001</Id>
  <Version>1.0.0</Version>
  <ProviderName>MJR West</ProviderName>
  <DefaultLocale>en-AU</DefaultLocale>
  <DisplayName DefaultValue="MJR West Agent"/>
  <Description DefaultValue="Capture property emails to the West Melbourne intelligence agent"/>
  <IconUrl DefaultValue="{base_url}/static/icon.png"/>
  <HighResolutionIconUrl DefaultValue="{base_url}/static/icon.png"/>
  <SupportUrl DefaultValue="{base_url}"/>
  <Hosts><Host Name="Mailbox"/></Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1"><Set Name="Mailbox"/></Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="{base_url}/outlook-taskpane"/>
        <RequestedHeight>280</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
</OfficeApp>"""

OUTLOOK_TASKPANE_HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width"/>
<title>MJR West Agent</title>
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
<style>
  body{font-family:-apple-system,sans-serif;margin:0;padding:12px;
       background:#0f1623;color:#e2e8f0;font-size:13px}
  .brand{color:#f59e0b;font-weight:800;font-size:14px;margin-bottom:8px}
  button{width:100%;padding:10px;border:none;border-radius:8px;
         background:#f59e0b;color:#000;font-weight:700;cursor:pointer;font-size:13px;
         margin-bottom:8px}
  button:disabled{background:#334155;color:#64748b;cursor:not-allowed}
  .status{font-size:11px;color:#64748b;min-height:16px;margin-top:4px}
  .ok{color:#10b981} .err{color:#ef4444}
  .sub{font-size:11px;color:#475569;margin-top:4px}
</style>
</head>
<body>
<div class="brand">&#127979; MJR West Agent</div>
<p style="color:#64748b;font-size:11px;margin-bottom:10px">
  Capture the selected email to the property intelligence database.</p>
<button id="captureBtn" onclick="captureEmail()" disabled>Capture Email</button>
<div class="status" id="status">Initialising Office.js…</div>
<div class="sub" id="preview"></div>
<script>
var BASE = window.location.origin;
Office.onReady(function(info){
  if(info.host===Office.HostType.Outlook){
    document.getElementById('captureBtn').disabled=false;
    document.getElementById('status').textContent='Ready';
    var item=Office.context.mailbox.item;
    if(item) document.getElementById('preview').textContent=
      (item.subject||'').substring(0,60);
  }
});
function captureEmail(){
  var btn=document.getElementById('captureBtn');
  var st=document.getElementById('status');
  btn.disabled=true; st.textContent='Reading email…';
  var item=Office.context.mailbox.item;
  var payload={
    subject:item.subject,
    from_email:item.from?item.from.emailAddress:'',
    from_name:item.from?item.from.displayName:'',
    date:item.dateTimeCreated?item.dateTimeCreated.toISOString():'',
    message_id:item.internetMessageId||''
  };
  item.body.getAsync(Office.CoercionType.Text,function(r){
    payload.body=r.status===Office.AsyncResultStatus.Succeeded?r.value:'';
    fetch(BASE+'/outlook-capture',{
      method:'POST',
      headers:{'Content-Type':'application/json',
               'X-Outlook-Token':Office.context.mailbox.userProfile.emailAddress||''},
      body:JSON.stringify(payload)
    }).then(function(res){return res.json()})
      .then(function(d){
        st.className='status ok';
        st.textContent=d.message||'Captured!';
        btn.disabled=false;
      }).catch(function(e){
        st.className='status err';
        st.textContent='Error: '+e.message;
        btn.disabled=false;
      });
  });
}
</script>
</body>
</html>"""


# ── Call log helpers ──────────────────────────────────────────────────────────
def _find_contact_by_number(number: str):
    """Return first matching contact record by phone number, or None."""
    if not number:
        return None
    clean = re.sub(r"[\s\-\(\)\+]", "", number)
    all_c = sb_select("contacts")
    for c in all_c:
        cphone = re.sub(r"[\s\-\(\)\+]", "", c.get("phone") or "")
        if cphone and (cphone == clean or cphone.endswith(clean[-8:]) or clean.endswith(cphone[-8:])):
            return c
    return None

def parse_call_log_file(filepath: str, filename: str) -> dict:
    """Parse Android call log CSV or JSON. Returns {'rows': [...], 'errors': [...]}"""
    ext = filename.rsplit(".", 1)[-1].lower()
    rows, errors = [], []
    try:
        if ext == "json":
            with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
                data = json.load(fh)
            if isinstance(data, list):
                raw_rows = data
            elif isinstance(data, dict):
                raw_rows = data.get("calls") or data.get("logs") or list(data.values())[0] if data else []
            else:
                raw_rows = []
            for r in raw_rows:
                rows.append({
                    "number":       str(r.get("number") or r.get("phone") or r.get("phoneNumber") or ""),
                    "direction":    str(r.get("type") or r.get("direction") or r.get("callType") or "Unknown"),
                    "call_date":    str(r.get("date") or r.get("time") or r.get("dateTime") or ""),
                    "duration_sec": int(r.get("duration") or r.get("durationSeconds") or r.get("duration_sec") or 0),
                    "contact_name": str(r.get("name") or r.get("contactName") or r.get("contact") or ""),
                })
        else:
            import csv as _csv
            with open(filepath, "r", encoding="utf-8-sig", errors="replace") as fh:
                reader = _csv.DictReader(fh)
                for r in reader:
                    # flexible column name matching
                    def _col(*keys):
                        for k in keys:
                            for col in r.keys():
                                if col.lower().replace(" ","_") == k.lower().replace(" ","_"):
                                    return r[col]
                        return ""
                    rows.append({
                        "number":       _col("number","phone","phone_number","tel"),
                        "direction":    _col("type","direction","call_type","calltype"),
                        "call_date":    _col("date","time","datetime","call_date","timestamp"),
                        "duration_sec": int(_col("duration","duration_sec","duration_seconds") or 0),
                        "contact_name": _col("name","contact","contact_name","display_name"),
                    })
    except Exception as e:
        errors.append(str(e))
    return {"rows": rows, "errors": errors}


def store_call_logs(rows: list, source_file: str) -> dict:
    """Cross-reference call log rows against contacts, store to DB."""
    saved = skipped = updated_leads = 0
    for r in rows:
        number = r.get("number", "")
        if not number:
            skipped += 1
            continue
        contact = _find_contact_by_number(number)
        row = {
            "number":       number,
            "direction":    r.get("direction", "Unknown"),
            "duration_sec": r.get("duration_sec", 0),
            "contact_name": r.get("contact_name") or (contact["name"] if contact else ""),
            "contact_id":   contact["id"] if contact else None,
            "source_file":  source_file,
        }
        # parse date
        raw_date = r.get("call_date", "")
        if raw_date:
            try:
                from email.utils import parsedate_to_datetime
                row["call_date"] = parsedate_to_datetime(raw_date).isoformat()
            except Exception:
                try:
                    from datetime import datetime as _dt
                    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M",
                                "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y %H:%M:%S"):
                        try:
                            row["call_date"] = _dt.strptime(raw_date.strip(), fmt).isoformat()
                            break
                        except ValueError:
                            continue
                except Exception:
                    pass
        if sb_insert("call_logs", row):
            saved += 1
            # If contact has open inquiries, note the call
            if contact:
                open_inq = sb_select("inquiries", {"contact_name": contact["name"], "status": "New"})
                for inq in open_inq[:1]:
                    sb_update("inquiries", inq["id"], {"status": "Contacted"})
                    updated_leads += 1
    return {"saved": saved, "skipped": skipped, "updated_leads": updated_leads}


# ── Call recording / Whisper ──────────────────────────────────────────────────
AUDIO_EXTENSIONS = {"mp3","mp4","m4a","wav","webm","ogg","flac","aac","mpga","mpeg"}

def transcribe_audio_file(filepath: str, filename: str) -> str:
    """Transcribe audio with OpenAI Whisper. Returns transcript text."""
    client = get_openai()
    if not client:
        return "[OpenAI not configured — add OPENAI_API_KEY]"
    try:
        with open(filepath, "rb") as fh:
            resp = client.audio.transcriptions.create(
                model="whisper-1",
                file=fh,
                response_format="text",
                language="en"
            )
        return resp if isinstance(resp, str) else getattr(resp, "text", str(resp))
    except Exception as e:
        log.error("Whisper transcription: %s", e)
        return f"[Transcription failed: {e}]"


def extract_call_intel(transcript: str, contact_name: str = "") -> dict:
    """GPT-4o analyses a call transcript for property intel + follow-up draft."""
    client = get_openai()
    if not client:
        return {}
    prompt = f"""You are analysing a call transcript for MJR West, a West Melbourne industrial property agent.
Contact: {contact_name or 'Unknown'}

Transcript:
{transcript[:8000]}

Return JSON only:
{{
  "summary": "2-3 sentence summary",
  "contact_name": "name if identified",
  "properties_discussed": ["address or description"],
  "key_facts": {{"key": "value"}},
  "action_items": ["string"],
  "lead_status": "Hot|Warm|Cold|Not a lead",
  "follow_up_email_draft": "Full draft email from Michael to the contact. \
Professional but conversational tone. \
West Melbourne industrial focus. Subject line on first line prefixed SUBJECT:"
}}"""
    try:
        resp = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role":"user","content":prompt}],
            response_format={"type":"json_object"},
            temperature=0.3,
            max_tokens=2000
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        log.error("extract_call_intel: %s", e)
        return {"summary": f"Analysis failed: {e}", "action_items": [], "follow_up_email_draft": ""}


def update_style_profile(transcript: str, summary: str):
    """Store transcript sample for style learning."""
    if not transcript.strip():
        return
    # Extract key phrases via GPT-4o mini (cheap)
    client = get_openai()
    phrases = []
    if client:
        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role":"user","content":
                    f"Extract 5 characteristic phrases or expressions that reveal the speaker's "
                    f"communication style. Return JSON: {{\"phrases\": [\"...\"]}}.\n\n{transcript[:3000]}"}],
                response_format={"type":"json_object"},
                temperature=0.2,
                max_tokens=300
            )
            phrases = json.loads(resp.choices[0].message.content).get("phrases", [])
        except Exception:
            pass
    sb_insert("style_profile", {
        "sample_type":         "call",
        "raw_text":            transcript[:5000],
        "key_phrases":         json.dumps(phrases),
        "communication_notes": summary,
    })


def get_style_context(n: int = 5) -> str:
    """Return recent style samples as a context string for email drafting."""
    samples = sb_select("style_profile", order="created_at", limit=n)
    if not samples:
        return ""
    parts = ["Michael's communication style (from recent calls):"]
    for s in samples:
        phrases = s.get("key_phrases") or []
        if isinstance(phrases, str):
            try: phrases = json.loads(phrases)
            except Exception: phrases = []
        if phrases:
            parts.append("  Phrases: " + "; ".join(phrases[:3]))
    return "\n".join(parts)


# ── Calendar / ICS ────────────────────────────────────────────────────────────
def parse_ics_file(filepath: str) -> list:
    """Parse ICS calendar file. Returns list of event dicts."""
    events = []
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
            content = fh.read()
    except Exception as e:
        log.error("ICS read: %s", e)
        return events

    def _ics_val(block: str, key: str) -> str:
        import re as _re
        m = _re.search(rf"^{key}[;:](.+?)(?=\r?\n[A-Z]|\Z)", block,
                       _re.MULTILINE | _re.DOTALL)
        if not m:
            return ""
        val = m.group(1).strip().replace("\r\n ", "").replace("\n ", "")
        # unescape ICS
        return val.replace("\\n", "\n").replace("\\,", ",").replace("\\;", ";")

    def _parse_dt(s: str):
        s = s.split(";")[-1].replace("VALUE=DATE-TIME:", "").replace("VALUE=DATE:", "").strip()
        s = s.replace("Z", "").strip()
        for fmt in ("%Y%m%dT%H%M%S", "%Y%m%dT%H%M", "%Y%m%d"):
            try:
                from datetime import datetime as _dt
                return _dt.strptime(s, fmt).isoformat()
            except ValueError:
                continue
        return None

    # split into VEVENT blocks
    import re as _re
    for block in _re.split(r"BEGIN:VEVENT", content):
        if "END:VEVENT" not in block:
            continue
        block = block.split("END:VEVENT")[0]
        uid       = _ics_val(block, "UID")
        title     = _ics_val(block, "SUMMARY")
        location  = _ics_val(block, "LOCATION")
        desc      = _ics_val(block, "DESCRIPTION")
        start_raw = _ics_val(block, "DTSTART")
        end_raw   = _ics_val(block, "DTEND")
        if not title:
            continue
        events.append({
            "uid":      uid,
            "title":    title,
            "start_dt": _parse_dt(start_raw) if start_raw else None,
            "end_dt":   _parse_dt(end_raw)   if end_raw   else None,
            "location": location,
            "description": desc,
            "is_property_related": any(kw in (title + location + desc).lower()
                                       for kw in ["property","lease","warehouse","factory",
                                                  "inspection","listing","tenant","landlord",
                                                  "west melbourne","meeting"]),
        })
    return events


def check_calendar_briefs():
    """
    APScheduler job (every 5 min): send rich WhatsApp pre-meeting brief 30 min before
    any property-related calendar event.
    Includes: property record, last 3 emails from that contact/company,
    semantic search context, GPT-4o talking points.
    """
    sb = get_sb()
    if sb is None:
        return
    now          = datetime.now(timezone.utc)
    window_start = now + timedelta(minutes=25)
    window_end   = now + timedelta(minutes=35)
    try:
        events = (sb.table("calendar_events")
                    .select("*")
                    .eq("brief_sent", False)
                    .eq("is_property_related", True)
                    .execute().data or [])
    except Exception:
        return

    for ev in events:
        if not ev.get("start_dt"):
            continue
        from datetime import datetime as _dt
        try:
            start = _dt.fromisoformat(ev["start_dt"].replace("Z", ""))
            if start.tzinfo is None:
                try:
                    import pytz
                    start = pytz.timezone(TIMEZONE).localize(start)
                except ImportError:
                    start = start.replace(tzinfo=timezone.utc)
            start_utc = start.astimezone(timezone.utc)
        except Exception:
            continue
        if not (window_start <= start_utc <= window_end):
            continue

        title    = ev.get("title", "Meeting")
        location = ev.get("location") or "No location"
        time_str = start.strftime("%H:%M")

        lines = [
            f"🗓️ *Pre-Meeting Brief — {title}*",
            f"⏰ {time_str}  📍 {location}",
            "",
        ]

        # ── Property record ───────────────────────────────────────────────────
        prop = None
        if location and location != "No location":
            loc_key = location.split(",")[0].strip().lower()
            all_props = sb_select("properties")
            for p in all_props:
                if loc_key and loc_key[:12] in (p.get("address") or "").lower():
                    prop = p
                    break
        if not prop:
            # try extracting address-like text from event title
            for p in sb_select("properties"):
                addr_frag = (p.get("address") or "")[:15].lower()
                if addr_frag and addr_frag in title.lower():
                    prop = p
                    break
        if prop:
            price = (f"${prop.get('asking_rent_pa',0):,.0f}/pa" if prop.get("asking_rent_pa")
                     else (f"${prop.get('asking_price',0):,.0f}" if prop.get("asking_price") else "POA"))
            lines.append("*Property*")
            lines.append(f"• {prop.get('address','?')} — {prop.get('size_sqm','?')} sqm")
            lines.append(f"• {prop.get('property_type','?')} | {prop.get('status','?')} | {price}")
            if prop.get("occupier"):
                lines.append(f"• Occupier: {prop['occupier']}")
            if prop.get("landlord"):
                lines.append(f"• Landlord: {prop['landlord']}")
            if prop.get("lease_expiry"):
                lines.append(f"• Lease Expiry: {str(prop['lease_expiry'])[:10]}")
            lines.append("")

        # ── Last 3 emails from this company/person ────────────────────────────
        company_hint = ""
        if prop:
            company_hint = prop.get("occupier") or prop.get("landlord") or ""
        if not company_hint:
            # pull company name from title words
            company_hint = title.split(" ")[0] if title else ""

        recent_emails = []
        if company_hint:
            try:
                all_emails = sb_select("email_logs", order="received_at", limit=200)
                hint_lower = company_hint.lower()[:12]
                recent_emails = [
                    e for e in all_emails
                    if hint_lower in (e.get("sender_name") or e.get("sender") or "").lower()
                    or hint_lower in (e.get("subject") or "").lower()
                ][:3]
            except Exception:
                pass
        if recent_emails:
            lines.append("*Recent Emails*")
            for e in recent_emails:
                nm  = (e.get("sender_name") or e.get("sender") or "?")[:20]
                sub = (e.get("subject") or "")[:40]
                dt  = (e.get("received_at") or "")[:10]
                lines.append(f"• {dt} {nm} — {sub}")
            lines.append("")

        # ── Semantic search context ───────────────────────────────────────────
        sem_context = []
        search_query = f"{title} {company_hint} West Melbourne industrial"
        try:
            sem_results = semantic_search(search_query, threshold=0.3, n=3)
            sem_context = [f"• {r.get('ai_summary','')[:80]}" for r in sem_results if r.get("ai_summary")]
        except Exception:
            pass
        if sem_context:
            lines.append("*Knowledge Base*")
            lines.extend(sem_context[:3])
            lines.append("")

        # ── GPT-4o talking points ─────────────────────────────────────────────
        client = get_openai()
        if client:
            context_text = "\n".join(lines)
            try:
                resp = client.chat.completions.create(
                    model=OPENAI_MODEL,
                    messages=[{"role": "user", "content":
                        f"You are helping Michael, a West Melbourne industrial property agent, "
                        f"prepare for a meeting.\n\nContext:\n{context_text}\n\n"
                        "Give 3 concise bullet-point talking points for this meeting. "
                        "Be specific, actionable, and property-focused. "
                        "Return plain text, one bullet per line, no JSON."}],
                    temperature=0.4,
                    max_tokens=300,
                )
                talking_pts = resp.choices[0].message.content.strip()
                lines.append("*Talking Points*")
                for pt in talking_pts.split("\n")[:3]:
                    if pt.strip():
                        lines.append(pt.strip())
                lines.append("")
            except Exception as e:
                log.warning("Pre-meeting talking points GPT: %s", e)

        lines.append("Good luck — check /calendar for full details.")
        broadcast_whatsapp("\n".join(lines))

        try:
            sb.table("calendar_events").update({"brief_sent": True}).eq("id", ev["id"]).execute()
        except Exception:
            pass
        log.info("Enhanced pre-meeting brief sent for: %s", title)


# ── Fee engine ────────────────────────────────────────────────────────────────
INSTITUTIONAL_CLIENTS = [
    "Dexus", "GPT Group", "Charter Hall", "Goodman", "ESR",
    "Logos", "Centuria", "ISPT", "Mirvac", "Growthpoint"
]

def parse_fee_schedule_excel(filepath: str) -> dict:
    """Parse institutional fee schedule spreadsheet. Returns {imported, errors}."""
    imported, errors = 0, []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        headers = [str(c.value or "").strip().lower() for c in list(ws.iter_rows(min_row=1, max_row=1, values_only=False))[0]]
        def _col(row, *keys):
            for k in keys:
                for i, h in enumerate(headers):
                    if k in h:
                        v = list(row)[i].value
                        return v
            return None
        for row in ws.iter_rows(min_row=2, values_only=False):
            client = _col(row, "client", "landlord", "company")
            if not client:
                continue
            fee_pct = _col(row, "fee %", "fee_pct", "rate", "%")
            flat    = _col(row, "flat", "fixed")
            min_fee = _col(row, "min", "minimum")
            deal_t  = _col(row, "type", "deal_type")
            notes   = _col(row, "notes", "comment")
            rec = {
                "client_name": str(client),
                "deal_type":   str(deal_t or "Lease"),
                "notes":       str(notes or ""),
            }
            for field, val in [("fee_pct", fee_pct), ("flat_fee", flat), ("min_fee", min_fee)]:
                try: rec[field] = float(val) if val not in (None, "") else None
                except Exception: rec[field] = None
            if sb_insert("fee_schedules", rec):
                imported += 1
    except Exception as e:
        errors.append(str(e))
    return {"imported": imported, "errors": errors}


def calculate_fee_for_deal(deal: dict) -> float | None:
    """Look up fee schedule and calculate fee for a closed deal."""
    landlord = deal.get("landlord_name") or deal.get("seller_name") or ""
    deal_type = deal.get("deal_type", "Lease")
    # find matching schedule
    schedules = sb_select("fee_schedules")
    schedule = None
    for s in schedules:
        if s.get("client_name", "").lower() in landlord.lower() or \
           landlord.lower() in s.get("client_name", "").lower():
            if s.get("deal_type", "Lease").lower() == deal_type.lower():
                schedule = s
                break
    if not schedule:
        # default: 10% of first year rent (Lease) or 2% of sale price (Sale)
        default_pct = 10.0 if deal_type == "Lease" else 2.0
        schedule = {"fee_pct": default_pct, "flat_fee": None, "min_fee": None}

    gross = float(deal.get("rent_pa") or deal.get("price") or 0)
    if gross == 0:
        return None

    fee_pct  = float(schedule.get("fee_pct") or 0)
    flat_fee = schedule.get("flat_fee")
    min_fee  = schedule.get("min_fee")

    if flat_fee:
        fee = float(flat_fee)
    else:
        fee = gross * fee_pct / 100.0

    if min_fee:
        fee = max(fee, float(min_fee))
    return round(fee, 2)


# ── Landlord portfolio builder ─────────────────────────────────────────────────
def build_landlord_portfolios() -> list:
    """Aggregate all properties by landlord with tenants, expiries, vacancies, fees."""
    props    = sb_select("properties")
    deals    = sb_select("deals")
    fees     = sb_select("fees")
    vacancies = sb_select("vacancies")

    landlords = {}
    for p in props:
        ll = (p.get("landlord") or "").strip()
        if not ll:
            ll = "Unknown Landlord"
        if ll not in landlords:
            landlords[ll] = {
                "name":       ll,
                "properties": [],
                "total_sqm":  0,
                "vacant_sqm": 0,
                "expiring_soon": [],  # within 12 months
                "total_fees": 0.0,
                "deal_count": 0,
            }
        landlords[ll]["properties"].append(p)
        sqm = float(p.get("size_sqm") or 0)
        landlords[ll]["total_sqm"] += sqm
        if p.get("status") in ("Available",):
            landlords[ll]["vacant_sqm"] += sqm
        # check lease expiry
        expiry = p.get("lease_expiry")
        if expiry:
            try:
                from datetime import date as _date
                exp_date = _date.fromisoformat(str(expiry)[:10])
                months_away = (exp_date - _date.today()).days / 30
                if 0 <= months_away <= 12:
                    landlords[ll]["expiring_soon"].append({
                        "address": p.get("address"),
                        "occupier": p.get("occupier"),
                        "expiry": str(expiry)[:10],
                        "months_away": round(months_away, 1),
                        "size_sqm": sqm,
                    })
            except Exception:
                pass

    # attach fees
    for f in fees:
        ll = (f.get("landlord") or f.get("client_name") or "").strip()
        if ll and ll in landlords:
            landlords[ll]["total_fees"] += float(f.get("fee_amount") or 0)
            landlords[ll]["deal_count"] += 1

    # attach active vacancies
    for v in vacancies:
        addr  = (v.get("address") or "").lower()
        owner = (v.get("owner") or "").strip()
        for ll, data in landlords.items():
            if owner and owner.lower() in ll.lower():
                data["vacant_sqm"] += float(v.get("size_sqm") or 0)

    return sorted(landlords.values(), key=lambda x: x["total_sqm"], reverse=True)


# ── Market Briefings ──────────────────────────────────────────────────────────
_MOTIVATIONAL_QUOTES = [
    "The best investment on earth is earth. — Louis Glickman",
    "Buy land, they're not making it anymore. — Mark Twain",
    "In real estate, you make 10% of your money because you're a genius and 90% because you catch a great wave. — Jeff Greene",
    "Ninety percent of all millionaires become so through owning real estate. — Andrew Carnegie",
    "Real estate is an imperishable asset, ever increasing in value. It is the most solid security that human ingenuity has devised. — Russell Sage",
    "Every person who invests in well-selected real estate in a growing section of a prosperous community adopts the surest and safest method of becoming independent. — Theodore Roosevelt",
    "The trouble with real estate is that it's local. You have to understand the local market. — Robert Kiyosaki",
    "Don't wait to buy real estate. Buy real estate and wait. — Will Rogers",
]

def build_market_brief() -> str:
    """Build the daily WhatsApp briefing with live Supabase data."""
    from datetime import date as _date
    import random

    total     = sb_count("properties")
    available = sb_count("properties", {"status": "Available"})
    leased    = sb_count("properties", {"status": "Leased"})
    sold      = sb_count("properties", {"status": "Sold"})
    under     = sb_count("properties", {"status": "Under Offer"})
    new_inq   = sb_count("inquiries",  {"status": "New"})
    vacancy   = round(available / total * 100 if total else 0, 1)
    date_str  = datetime.now().strftime("%A, %d %B %Y")

    lines = [
        "🏭 *West Melbourne Industrial — Daily Brief*",
        f"📅 {date_str}",
        "",
        "*Market Snapshot*",
        f"• Total Listings: {total}",
        f"• Available: {available} ({vacancy}% vacancy)",
        f"• Under Offer: {under}  |  Leased: {leased}  |  Sold: {sold}",
        f"• New Inquiries: {new_inq}",
        "",
    ]

    # ── Today's calendar meetings ──────────────────────────────────────────────
    try:
        today_start = datetime.now().strftime("%Y-%m-%d") + "T00:00:00"
        today_end   = datetime.now().strftime("%Y-%m-%d") + "T23:59:59"
        sb = get_sb()
        if sb:
            cal_rows = (sb.table("calendar_events")
                          .select("title,start_dt,location,is_property_related")
                          .gte("start_dt", today_start)
                          .lte("start_dt", today_end)
                          .execute().data or [])
        else:
            cal_rows = []
    except Exception:
        cal_rows = []
    if cal_rows:
        lines.append("*Today's Meetings*")
        for ev in cal_rows[:5]:
            t   = (ev.get("start_dt") or "")[:16].replace("T", " ")[-5:]
            tag = " 🏭" if ev.get("is_property_related") else ""
            lines.append(f"🗓️ {t} — {ev.get('title','?')[:45]}{tag}")
            if ev.get("location"):
                lines.append(f"   📍 {ev['location'][:40]}")
        lines.append("")

    # ── Unanswered emails ──────────────────────────────────────────────────────
    try:
        unanswered = sb_select("email_logs", order="received_at", limit=100) if get_sb() else []
        pending = [e for e in unanswered
                   if e.get("reply_status") == "pending" and e.get("requires_action")]
    except Exception:
        pending = []
    if pending:
        lines.append(f"*📬 Unanswered Emails: {len(pending)}*")
        for e in pending[:4]:
            nm = (e.get("sender_name") or e.get("sender") or "?")[:25]
            sub = (e.get("subject") or "")[:40]
            pri = (e.get("ai_priority") or "").upper()[:4]
            lines.append(f"• [{pri}] {nm} — {sub}")
        if len(pending) > 4:
            lines.append(f"  …and {len(pending)-4} more")
        lines.append("")

    # ── Top 5 hot leads by score ───────────────────────────────────────────────
    try:
        all_inq = sb_select("inquiries", order="created_at", limit=50)
        scored  = sorted([dict(i, _s=score_lead(i)) for i in all_inq],
                         key=lambda x: x["_s"], reverse=True)
        hot_leads = [i for i in scored if i["_s"] >= 3][:5]
    except Exception:
        hot_leads = []
    if hot_leads:
        lines.append("*🔥 Hot Leads*")
        for i in hot_leads:
            nm  = (i.get("contact_name") or "?")[:22]
            ph  = i.get("contact_phone") or i.get("contact_email") or "—"
            src = i.get("source", "")[:10]
            lines.append(f"• {nm}  {ph}  [{src}]")
        lines.append("")

    # ── Lease expiries within 90 days ─────────────────────────────────────────
    try:
        today   = _date.today()
        in_90   = _date(today.year + (1 if today.month > 9 else 0),
                        (today.month + 3 - 1) % 12 + 1, today.day)
        all_props = sb_select("properties")
        expiring = []
        for p in all_props:
            exp = p.get("lease_expiry")
            if not exp:
                continue
            try:
                exp_d = _date.fromisoformat(str(exp)[:10])
                days  = (exp_d - today).days
                if 0 <= days <= 90:
                    expiring.append((days, p))
            except Exception:
                pass
        expiring.sort(key=lambda x: x[0])
    except Exception:
        expiring = []
    if expiring:
        lines.append("*⏰ Leases Expiring < 90 Days*")
        for days, p in expiring[:4]:
            lines.append(
                f"• {p.get('address','?')[:35]} — "
                f"{p.get('occupier') or p.get('landlord') or '?'} "
                f"({days}d)"
            )
        lines.append("")

    # ── New requirements ───────────────────────────────────────────────────────
    try:
        recent_reqs = sb_select("requirements", order="created_at", limit=3)
    except Exception:
        recent_reqs = []
    if recent_reqs:
        lines.append("*📋 Recent Requirements (West)*")
        for r in recent_reqs:
            sz = f"{r.get('size_min_sqm','?')}–{r.get('size_max_sqm','?')} sqm"
            lines.append(f"• {r.get('company','?')[:25]} — {sz} — {r.get('preferred_location','West Melbourne')}")
        lines.append("")

    # ── Featured available listings ───────────────────────────────────────────
    try:
        featured = sb_select("properties", filters={"status": "Available"}, order="created_at", limit=3)
    except Exception:
        featured = []
    if featured:
        lines.append("*📍 Featured Listings*")
        for p in featured:
            price = (f"${p.get('asking_rent_pa', 0):,.0f}/pa" if p.get("asking_rent_pa")
                     else (f"${p.get('asking_price', 0):,.0f}" if p.get("asking_price") else "POA"))
            lines.append(f"• {p.get('address','N/A')[:35]}")
            lines.append(f"  {p.get('property_type','')} | {p.get('size_sqm','')} sqm | {price}")
        lines.append("")

    # ── Quote of the day ───────────────────────────────────────────────────────
    lines.append(f"💬 _{random.choice(_MOTIVATIONAL_QUOTES)}_")
    lines.append("")
    lines.append("Reply HELP for commands.")
    return "\n".join(lines)

def send_daily_briefing():
    log.info("Running daily briefing job.")
    content = build_market_brief()
    sent, failed = broadcast_whatsapp(content)
    sb_insert("briefings", {
        "briefing_type": "Daily",
        "content": content,
        "sent_to": json.dumps(WA_BROADCAST_LIST),
        "channel": "WhatsApp"
    })
    log.info("Daily briefing sent to %d, failed %d.", sent, failed)

def send_weekly_briefing():
    log.info("Running weekly briefing job.")
    deals = sb_select("deals", order="created_at", limit=5)
    base  = build_market_brief()
    extra_lines = ["\n*📊 Weekly Transaction Summary*"]
    if deals:
        for d in deals:
            price = f"${d.get('price', 0):,.0f}" if d.get("price") else "Undisclosed"
            extra_lines.append(f"• {d.get('deal_type', 'Deal')} | {price} | {d.get('status', '')}")
    else:
        extra_lines.append("• No transactions recorded this week.")
    content = base + "\n".join(extra_lines)
    broadcast_whatsapp(content)
    sb_insert("briefings", {
        "briefing_type": "Weekly",
        "content": content,
        "sent_to": json.dumps(WA_BROADCAST_LIST),
        "channel": "WhatsApp"
    })

def check_gmail_job():
    log.info("Running Gmail check job.")
    fetch_gmail_emails()


# ── Relationship decay alerts ─────────────────────────────────────────────────
def check_relationship_decay():
    """
    Weekly job: WhatsApp alert for contacts not touched in 30+ days
    who have a property with a lease expiring within 24 months.
    """
    sb = get_sb()
    if sb is None:
        return
    cutoff = (datetime.now(timezone.utc) - timedelta(days=30)).isoformat()
    from datetime import date as _date
    today     = _date.today()
    exp_limit = _date(today.year + 2, today.month, today.day).isoformat()
    try:
        contacts = (sb.table("contacts")
                      .select("id,name,company,phone,email,last_contacted_at")
                      .execute().data or [])
        props = sb_select("properties")
    except Exception as e:
        log.error("check_relationship_decay: %s", e)
        return

    # index properties with upcoming expiry by landlord/occupier name
    expiring_props = []
    for p in props:
        exp = p.get("lease_expiry")
        if not exp:
            continue
        try:
            exp_d = _date.fromisoformat(str(exp)[:10])
            if today <= exp_d <= _date.fromisoformat(exp_limit[:10]):
                expiring_props.append(p)
        except Exception:
            pass

    stale = []
    for c in contacts:
        lca = c.get("last_contacted_at")
        if lca:
            try:
                lca_dt = datetime.fromisoformat(str(lca).replace("Z", "+00:00").replace(" ", "T"))
                if lca_dt.replace(tzinfo=timezone.utc) > datetime.fromisoformat(cutoff):
                    continue  # contacted recently
            except Exception:
                pass
        else:
            pass  # never contacted — always include if they have expiring leases

        # check if this contact is linked to an expiring property (by company name match)
        cname = _normalise_company(c.get("company") or c.get("name") or "")
        for p in expiring_props:
            ll  = _normalise_company(p.get("landlord") or "")
            occ = _normalise_company(p.get("occupier") or "")
            if cname and (cname in ll or cname in occ or ll in cname or occ in cname):
                months_away = round(
                    (_date.fromisoformat(str(p["lease_expiry"])[:10]) - today).days / 30, 1
                )
                stale.append({
                    "contact": c,
                    "property": p,
                    "months_away": months_away,
                })
                break

    if not stale:
        log.info("Relationship decay check: no stale contacts with expiring leases.")
        return

    lines = [f"🤝 *Relationship Decay Alert — {len(stale)} contact{'s' if len(stale)!=1 else ''}*",
             "These landlords/tenants haven't been contacted in 30+ days and have leases expiring:"]
    for item in sorted(stale, key=lambda x: x["months_away"])[:8]:
        c  = item["contact"]
        p  = item["property"]
        mo = item["months_away"]
        ph = c.get("phone") or c.get("email") or "no details"
        lines.append(
            f"• {c.get('name','?')} ({c.get('company','')}) — "
            f"{p.get('address','?')[:35]} expires {mo}mo · {ph}"
        )
    broadcast_whatsapp("\n".join(lines))
    log.info("Relationship decay alert sent: %d contacts", len(stale))


# ── Vacancy ↔ Requirements matching ──────────────────────────────────────────
def match_vacancies_to_requirements():
    """
    Daily job: match new vacancies against active requirements by size range + region.
    Sends WhatsApp alert for new matches. Stores results in vacancy_matches.
    """
    sb = get_sb()
    if sb is None:
        return
    try:
        vacancies    = sb_select("vacancies",    order="created_at", limit=200)
        requirements = sb_select("requirements", order="created_at", limit=200)
    except Exception as e:
        log.error("match_vacancies_to_requirements: %s", e)
        return

    _WEST = {"w", "west", "west melbourne", "western", "wm", ""}

    new_matches = []
    for vac in vacancies:
        vac_sqm = float(vac.get("size_sqm") or 0)
        vac_sub = (vac.get("suburb") or "West Melbourne").lower()
        vac_addr = vac.get("address") or "?"
        for req in requirements:
            # region gate — only West Melbourne requirements
            region_key = (req.get("region") or "w").strip().lower()
            if region_key not in _WEST:
                continue
            sz_min = float(req.get("size_min_sqm") or 0)
            sz_max = float(req.get("size_max_sqm") or 0)
            if sz_max == 0:
                sz_max = sz_min * 2 if sz_min else 99999
            # size match with 15% tolerance
            if vac_sqm == 0 or not (sz_min * 0.85 <= vac_sqm <= sz_max * 1.15):
                continue
            # suburb match — prefer same suburb but accept West Melbourne broadly
            req_loc = (req.get("preferred_location") or "").lower()
            if req_loc and req_loc not in vac_sub and vac_sub not in req_loc:
                if not any(w in req_loc for w in ["west", "any", "all", ""]):
                    continue

            # check if already matched
            try:
                existing = (sb.table("vacancy_matches")
                              .select("id")
                              .eq("vacancy_id",     vac["id"])
                              .eq("requirement_id", req["id"])
                              .execute().data or [])
                if existing:
                    continue
            except Exception:
                pass

            score = 100
            if vac_sqm > 0 and sz_min <= vac_sqm <= sz_max:
                score = 100  # perfect fit
            elif sz_min * 0.85 <= vac_sqm <= sz_max * 1.15:
                score = 80   # tolerance fit

            sb_insert("vacancy_matches", {
                "vacancy_id":     vac["id"],
                "requirement_id": req["id"],
                "score":          score,
                "alerted":        False,
            })
            new_matches.append({
                "vac":  vac,
                "req":  req,
                "score": score,
            })

    if not new_matches:
        log.info("Vacancy matching: no new matches found.")
        return

    lines = [f"🏭 *{len(new_matches)} New Vacancy Match{'es' if len(new_matches)!=1 else ''}*"]
    for m in new_matches[:6]:
        v, r = m["vac"], m["req"]
        lines.append(
            f"• {v.get('address','?')[:35]} ({v.get('size_sqm','?')} sqm) "
            f"↔ {r.get('company','?')[:25]} ({r.get('size_min_sqm','?')}–"
            f"{r.get('size_max_sqm','?')} sqm)"
        )
    if len(new_matches) > 6:
        lines.append(f"…and {len(new_matches)-6} more — check the dashboard.")
    broadcast_whatsapp("\n".join(lines))

    # mark all as alerted
    try:
        ids = []
        for m in new_matches:
            rows = (sb.table("vacancy_matches")
                      .select("id")
                      .eq("vacancy_id",     m["vac"]["id"])
                      .eq("requirement_id", m["req"]["id"])
                      .execute().data or [])
            ids.extend([r["id"] for r in rows])
        if ids:
            sb.table("vacancy_matches").update({"alerted": True}).in_("id", ids).execute()
    except Exception as e:
        log.error("vacancy_matches alerted update: %s", e)

    log.info("Vacancy matching: %d new matches alerted.", len(new_matches))


# ── Post-meeting note AI processing ──────────────────────────────────────────
def process_meeting_notes(event: dict, notes_text: str):
    """
    GPT-4o processes post-meeting notes:
    - Extracts action items
    - Drafts a follow-up email
    - Extracts property intel (size, rent, sentiment)
    - Updates contact last_contacted_at
    """
    client = get_openai()
    if not client or not notes_text.strip():
        return
    sb = get_sb()
    event_id = event.get("id")

    prompt = f"""You are analysing post-meeting notes for Michael, a West Melbourne industrial property agent.

Meeting: {event.get('title','')}
Location: {event.get('location','')}
Date: {(event.get('start_dt') or '')[:16]}

Notes:
{notes_text[:4000]}

Return JSON only:
{{
  "action_items": ["string"],
  "follow_up_email": "Full draft email from Michael. Include SUBJECT: on first line.",
  "contact_name": "name of person met",
  "contact_update": "Hot|Warm|Cold|null",
  "property_intel": {{
    "address": "if mentioned",
    "size_sqm": null,
    "rent_pa": null,
    "sentiment": "positive|neutral|negative"
  }},
  "summary": "2-3 sentence summary of the meeting"
}}"""

    try:
        resp = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.3,
            max_tokens=1500,
        )
        data = json.loads(resp.choices[0].message.content)
    except Exception as e:
        log.error("process_meeting_notes GPT: %s", e)
        return

    # Update calendar event with AI results
    if sb and event_id:
        try:
            sb.table("calendar_events").update({
                "ai_action_items":    json.dumps(data.get("action_items", [])),
                "follow_up_draft":    data.get("follow_up_email", ""),
                "notes_processed_at": datetime.now(timezone.utc).isoformat(),
            }).eq("id", event_id).execute()
        except Exception as e:
            log.error("process_meeting_notes calendar update: %s", e)

    # Update contact last_contacted_at
    contact_name = data.get("contact_name") or ""
    if contact_name and sb:
        try:
            matches = (sb.table("contacts")
                         .select("id")
                         .ilike("name", f"%{contact_name.split()[0]}%")
                         .execute().data or [])
            for c in matches[:1]:
                sb.table("contacts").update({
                    "last_contacted_at": datetime.now(timezone.utc).isoformat(),
                }).eq("id", c["id"]).execute()
                if data.get("contact_update") in ("Hot", "Warm", "Cold"):
                    sb.table("contacts").update({
                        "lead_status": data["contact_update"]
                    }).eq("id", c["id"]).execute()
        except Exception as e:
            log.error("process_meeting_notes contact update: %s", e)

    log.info("Meeting notes processed for event %s", event_id)


# ── APScheduler ───────────────────────────────────────────────────────────────
scheduler = BackgroundScheduler(timezone=TIMEZONE)
scheduler.add_job(send_daily_briefing,  CronTrigger(hour=BRIEFING_HOUR, minute=BRIEFING_MINUTE),
                  id="daily_brief",   replace_existing=True)
scheduler.add_job(send_weekly_briefing, CronTrigger(day_of_week="mon", hour=BRIEFING_HOUR, minute=BRIEFING_MINUTE),
                  id="weekly_brief",  replace_existing=True)
scheduler.add_job(check_gmail_job,      "interval", minutes=15,
                  id="gmail_check",   replace_existing=True)
scheduler.add_job(check_calendar_briefs, "interval", minutes=5,
                  id="calendar_briefs", replace_existing=True)
scheduler.add_job(flag_unanswered_emails, "interval", hours=1,
                  id="unanswered_check", replace_existing=True)
scheduler.add_job(check_relationship_decay, CronTrigger(day_of_week="fri", hour=7, minute=30),
                  id="decay_check", replace_existing=True)
scheduler.add_job(match_vacancies_to_requirements, CronTrigger(hour=6, minute=0),
                  id="vacancy_match", replace_existing=True)

# ── Auth decorator ────────────────────────────────────────────────────────────
def login_required(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

# ── Lead scoring ─────────────────────────────────────────────────────────────
def score_lead(inq: dict) -> int:
    """Return urgency score 1–4 for an inquiry."""
    score = 1
    msg = ((inq.get("message") or "") + " " + (inq.get("contact_name") or "")).lower()
    hot_words = ["urgent", "asap", "today", "tomorrow", "immediately", "must", "deadline", "eoi"]
    warm_words = ["interested", "inspect", "viewing", "keen", "ready", "buy", "lease"]
    if any(w in msg for w in hot_words): score += 2
    elif any(w in msg for w in warm_words): score += 1
    if inq.get("source") in ("WhatsApp", "Phone"): score += 1
    try:
        age = (datetime.now(timezone.utc) -
               datetime.fromisoformat(str(inq.get("created_at", "")).replace("Z", "+00:00")
                                      .replace(" ", "T"))).total_seconds() / 3600
        if age < 4: score += 1
    except Exception: pass
    return min(score, 4)

def urgency_badge(score: int) -> str:
    labels = {1: ("LOW",  "#374151", "#9ca3af"),
              2: ("WARM", "#78350f", "#fbbf24"),
              3: ("HOT",  "#7c2d12", "#f97316"),
              4: ("FIRE", "#7f1d1d", "#ef4444")}
    bg, fg = labels.get(score, labels[1])[0:3:2]  # noqa
    lbl = labels.get(score, labels[1])[0]
    return (f'<span style="background:{bg};color:{fg};font-size:.65rem;font-weight:700;'
            f'padding:.25em .55em;border-radius:4px;letter-spacing:.06em">{lbl}</span>')

# ── HTML Templates ────────────────────────────────────────────────────────────
_LAYOUT = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>{{ title }} — MJR West Agent</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"/>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css"/>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{--sb:240px;--amber:#f59e0b;--amber-lt:#fbbf24;--amber-dim:#78350f;
      --bg:#0a0d14;--bg2:#0f1521;--bg3:#131c2e;--border:#1e2d45;
      --text:#e2e8f0;--text2:#64748b;--text3:#94a3b8}
*{box-sizing:border-box}
body{background:var(--bg);color:var(--text);min-height:100vh;
     font-family:'Segoe UI',system-ui,sans-serif;font-size:14px}
/* ─── Scrollbar ─── */
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--bg2)}
::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:3px}
/* ─── Sidebar ─── */
.sidebar{position:fixed;top:0;left:0;width:var(--sb);height:100vh;
  background:var(--bg2);border-right:1px solid var(--border);
  overflow-y:auto;z-index:200;display:flex;flex-direction:column}
.brand{padding:1.4rem 1.1rem 1rem;border-bottom:1px solid var(--border)}
.brand-logo{display:flex;align-items:center;gap:.6rem;margin-bottom:.25rem}
.brand-icon{width:32px;height:32px;background:var(--amber);border-radius:8px;
  display:flex;align-items:center;justify-content:center;color:#000;font-size:.95rem;font-weight:900}
.brand-name{font-size:.88rem;font-weight:800;color:#fff;letter-spacing:-.01em}
.brand-sub{color:var(--text2);font-size:.68rem;padding-left:38px}
.nav-section{color:#2d4a6b;font-size:.6rem;font-weight:800;letter-spacing:.12em;
  text-transform:uppercase;padding:.9rem 1.2rem .2rem}
.nav-link{color:var(--text2);padding:.42rem .9rem;border-radius:7px;
  margin:.08rem .45rem;font-size:.81rem;display:flex;align-items:center;gap:.55rem;
  text-decoration:none;transition:all .15s}
.nav-link i{font-size:.9rem;width:17px;text-align:center;opacity:.7}
.nav-link:hover{color:var(--amber-lt);background:rgba(245,158,11,.07)}
.nav-link.active{color:var(--amber);background:rgba(245,158,11,.12);
  border-left:2px solid var(--amber);padding-left:calc(.9rem - 2px)}
.nav-link.active i{opacity:1}
.nav-badge{background:var(--amber);color:#000;font-size:.6rem;font-weight:800;
  padding:.15em .45em;border-radius:20px;margin-left:auto}
/* ─── Main ─── */
.main{margin-left:var(--sb);padding:1.6rem 1.8rem;min-height:100vh}
/* ─── Page header ─── */
.ph{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:1.4rem}
.ph h4{font-weight:800;color:#fff;margin:0;font-size:1.15rem;letter-spacing:-.02em}
.ph p{color:var(--text2);font-size:.8rem;margin:.2rem 0 0}
/* ─── Cards ─── */
.card{background:var(--bg3);border:1px solid var(--border);border-radius:12px;overflow:hidden}
.card-header{background:transparent;border-bottom:1px solid var(--border);
  font-weight:700;font-size:.8rem;color:var(--text3);letter-spacing:.04em;
  text-transform:uppercase;padding:.75rem 1.1rem}
.card-body{padding:1.1rem}
/* ─── KPI tiles ─── */
.kpi{background:var(--bg3);border:1px solid var(--border);border-radius:12px;
  padding:1.1rem 1.2rem;height:100%}
.kpi-val{font-size:1.8rem;font-weight:900;color:#fff;line-height:1;letter-spacing:-.03em}
.kpi-val.amber{color:var(--amber)}
.kpi-val.green{color:#10b981}
.kpi-val.red{color:#ef4444}
.kpi-val.blue{color:#38bdf8}
.kpi-label{font-size:.68rem;color:var(--text2);text-transform:uppercase;
  letter-spacing:.07em;margin-top:.3rem}
.kpi-delta{font-size:.7rem;margin-top:.4rem}
.kpi-icon{width:36px;height:36px;border-radius:8px;display:flex;align-items:center;
  justify-content:center;font-size:1rem}
/* ─── Morning brief banner ─── */
.brief-banner{background:linear-gradient(135deg,#111827 0%,#1c2e1a 100%);
  border:1px solid #1f3d1a;border-radius:12px;padding:1.1rem 1.4rem;margin-bottom:1.4rem}
.brief-banner .date-str{color:var(--amber);font-size:.72rem;font-weight:700;
  letter-spacing:.08em;text-transform:uppercase}
.brief-banner h5{color:#fff;font-weight:800;margin:.2rem 0 .1rem;font-size:1rem}
.brief-banner p{color:#86efac;font-size:.8rem;margin:0}
/* ─── Call list ─── */
.call-item{display:flex;align-items:center;gap:.75rem;padding:.65rem .8rem;
  border-radius:8px;border:1px solid var(--border);margin-bottom:.4rem;
  background:var(--bg2);transition:border-color .15s}
.call-item:hover{border-color:var(--amber-dim)}
.call-avatar{width:34px;height:34px;border-radius:8px;background:var(--amber-dim);
  color:var(--amber);display:flex;align-items:center;justify-content:center;
  font-weight:800;font-size:.8rem;flex-shrink:0}
.call-name{font-weight:700;color:#fff;font-size:.82rem;line-height:1.2}
.call-detail{color:var(--text2);font-size:.72rem}
.call-actions{margin-left:auto;display:flex;gap:.3rem}
.btn-call{background:rgba(245,158,11,.15);color:var(--amber);border:1px solid var(--amber-dim);
  border-radius:6px;padding:.25rem .55rem;font-size:.72rem;font-weight:700;
  cursor:pointer;transition:all .15s;text-decoration:none}
.btn-call:hover{background:var(--amber);color:#000}
.btn-done{background:rgba(16,185,129,.1);color:#10b981;border:1px solid #065f46;
  border-radius:6px;padding:.25rem .55rem;font-size:.72rem;font-weight:700;cursor:pointer}
/* ─── Tables ─── */
.table{color:var(--text);--bs-table-bg:transparent;--bs-table-hover-bg:rgba(245,158,11,.04)}
.table th{font-size:.67rem;text-transform:uppercase;letter-spacing:.07em;
  color:#2d4a6b;font-weight:800;border-color:var(--border);border-top:none;padding:.6rem .8rem}
.table td{font-size:.81rem;vertical-align:middle;border-color:var(--border);padding:.55rem .8rem}
.table-hover tbody tr:hover{--bs-table-accent-bg:rgba(245,158,11,.04)}
/* ─── Badges ─── */
.badge{font-size:.65rem;font-weight:700;padding:.3em .6em;border-radius:5px}
.badge-available{background:#052e16;color:#4ade80;border:1px solid #14532d}
.badge-leased{background:#0c1a35;color:#60a5fa;border:1px solid #1e3a5f}
.badge-sold{background:#1a0533;color:#c084fc;border:1px solid #3b0764}
.badge-offer{background:#2d1a00;color:#fbbf24;border:1px solid #78350f}
.badge-new{background:#2d000a;color:#f87171;border:1px solid #7f1d1d}
/* ─── Buttons ─── */
.btn-primary{background:var(--amber);border-color:var(--amber);color:#000;font-weight:700}
.btn-primary:hover{background:var(--amber-lt);border-color:var(--amber-lt);color:#000}
.btn-outline-secondary{border-color:var(--border);color:var(--text3)}
.btn-outline-secondary:hover{background:rgba(255,255,255,.06);color:var(--text);border-color:#334155}
.btn-sm{font-size:.75rem;padding:.3rem .65rem}
.btn-xs{font-size:.68rem;padding:.2rem .5rem}
/* ─── Forms ─── */
.form-control,.form-select{background:var(--bg2);border:1px solid var(--border);
  color:var(--text);font-size:.83rem;border-radius:8px}
.form-control:focus,.form-select:focus{background:var(--bg2);color:var(--text);
  border-color:var(--amber);box-shadow:0 0 0 2px rgba(245,158,11,.2)}
.form-control::placeholder{color:var(--text2)}
.form-label{font-size:.78rem;font-weight:700;color:var(--text3);margin-bottom:.3rem}
.form-select option{background:#1e293b}
.modal-content{background:var(--bg3);border:1px solid var(--border);color:var(--text)}
.modal-header{border-bottom:1px solid var(--border)}
.modal-footer{border-top:1px solid var(--border)}
.modal-title{font-size:.9rem;font-weight:800;color:#fff}
.btn-close{filter:invert(1) opacity(.5)}
/* ─── Alert/toast ─── */
.toast-container{position:fixed;top:1rem;right:1rem;z-index:9999}
.toast{background:var(--bg3);border:1px solid var(--border);color:var(--text)}
/* ─── Misc ─── */
pre{color:var(--text3);font-size:.78rem}
code{color:var(--amber)}
.alert-info{background:rgba(56,189,248,.07);border-color:#0e4f6b;color:#7dd3fc}
.text-muted{color:var(--text2)!important}
hr{border-color:var(--border)}
</style>
</head>
<body>
<div class="sidebar">
  <div class="brand">
    <div class="brand-logo">
      <div class="brand-icon">W</div>
      <span class="brand-name">MJR West</span>
    </div>
    <div class="brand-sub">Industrial Property Agent</div>
  </div>
  <nav class="mt-1 flex-grow-1">
    <div class="nav-section">Command Centre</div>
    <a href="/dashboard" class="nav-link {{ 'active' if active=='dashboard' }}"><i class="bi bi-grid-1x2-fill"></i>Dashboard</a>
    <a href="/analytics" class="nav-link {{ 'active' if active=='analytics' }}"><i class="bi bi-bar-chart-fill"></i>Analytics</a>
    <div class="nav-section">Pipeline</div>
    <a href="/inquiries"  class="nav-link {{ 'active' if active=='inquiries'  }}"><i class="bi bi-lightning-fill"></i>Lead Pipeline</a>
    <a href="/properties" class="nav-link {{ 'active' if active=='properties' }}"><i class="bi bi-buildings-fill"></i>Properties</a>
    <a href="/contacts"   class="nav-link {{ 'active' if active=='contacts'   }}"><i class="bi bi-people-fill"></i>Contacts</a>
    <a href="/deals"      class="nav-link {{ 'active' if active=='deals'      }}"><i class="bi bi-handshake-fill"></i>Deals</a>
    <div class="nav-section">Intelligence</div>
    <a href="/briefings"  class="nav-link {{ 'active' if active=='briefings'  }}"><i class="bi bi-newspaper"></i>Briefings</a>
    <a href="/whatsapp"   class="nav-link {{ 'active' if active=='whatsapp'   }}"><i class="bi bi-whatsapp"></i>WhatsApp</a>
    <a href="/email-log"  class="nav-link {{ 'active' if active=='email'      }}"><i class="bi bi-envelope-fill"></i>Email Inbox</a>
    <a href="/outlook-setup" class="nav-link {{ 'active' if active=='outlook' }}"><i class="bi bi-envelope-check-fill"></i>Outlook Setup</a>
    <a href="/calendar"   class="nav-link {{ 'active' if active=='calendar'   }}"><i class="bi bi-calendar3"></i>Calendar</a>
    <div class="nav-section">Field Tools</div>
    <a href="/call-logs"  class="nav-link {{ 'active' if active=='calllogs'   }}"><i class="bi bi-telephone-fill"></i>Call Logs</a>
    <a href="/recordings" class="nav-link {{ 'active' if active=='recordings' }}"><i class="bi bi-mic-fill"></i>Recordings</a>
    <div class="nav-section">AI Knowledge</div>
    <a href="/upload"     class="nav-link {{ 'active' if active=='upload'     }}"><i class="bi bi-cloud-arrow-up-fill"></i>AI Intake</a>
    <a href="/documents"  class="nav-link {{ 'active' if active=='documents'  }}"><i class="bi bi-file-earmark-text-fill"></i>Documents</a>
    <a href="/search"     class="nav-link {{ 'active' if active=='search'     }}"><i class="bi bi-search-heart-fill"></i>Semantic Search</a>
    <a href="/email-draft" class="nav-link {{ 'active' if active=='emaildraft' }}"><i class="bi bi-pencil-square"></i>Email Drafter</a>
    <div class="nav-section">Business</div>
    <a href="/fees"       class="nav-link {{ 'active' if active=='fees'       }}"><i class="bi bi-currency-dollar"></i>Fee Tracking</a>
    <a href="/landlords"  class="nav-link {{ 'active' if active=='landlords'  }}"><i class="bi bi-person-workspace"></i>Landlord Portfolios</a>
    <div class="nav-section">Tools</div>
    <a href="/settings"   class="nav-link {{ 'active' if active=='settings'   }}"><i class="bi bi-gear-fill"></i>Settings</a>
    <a href="/logout"     class="nav-link"><i class="bi bi-box-arrow-left"></i>Logout</a>
  </nav>
  <div style="padding:.8rem 1.1rem;border-top:1px solid var(--border);font-size:.67rem;color:#1e3a5f">
    MJR West Agent &bull; v3.0
  </div>
</div>
<div class="main">
{% with msgs = get_flashed_messages(with_categories=true) %}{% if msgs %}
<div class="toast-container">
{% for cat, msg in msgs %}
<div class="toast show border-0 mb-2" role="alert"
     style="background:{{ '#052e16' if cat!='error' else '#2d000a' }};border:1px solid {{ '#14532d' if cat!='error' else '#7f1d1d' }}!important">
  <div class="d-flex align-items-center">
    <div class="toast-body" style="color:{{ '#4ade80' if cat!='error' else '#f87171' }};font-size:.82rem;font-weight:600">
      <i class="bi bi-{{ 'check-circle-fill' if cat!='error' else 'exclamation-circle-fill' }} me-2"></i>{{ msg }}
    </div>
    <button type="button" class="btn-close me-2" data-bs-dismiss="toast"></button>
  </div>
</div>
{% endfor %}
</div>{% endif %}{% endwith %}
{{ content | safe }}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>document.querySelectorAll('[data-bs-dismiss="toast"]').forEach(b=>b.addEventListener('click',e=>e.target.closest('.toast').remove()));</script>
</body>
</html>"""

_LOGIN_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>MJR West Agent — Sign In</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a0d14;min-height:100vh;display:flex;align-items:center;
     justify-content:center;font-family:'Segoe UI',system-ui,sans-serif}
.wrap{width:100%;max-width:380px;padding:1rem}
.card{background:#0f1521;border:1px solid #1e2d45;border-radius:16px;padding:2.2rem 2rem}
.logo{display:flex;align-items:center;gap:.7rem;margin-bottom:1.6rem}
.logo-icon{width:40px;height:40px;background:#f59e0b;border-radius:10px;
  display:flex;align-items:center;justify-content:center;font-size:1.1rem;font-weight:900;color:#000}
.logo-text{font-size:1rem;font-weight:800;color:#fff}
.logo-sub{font-size:.72rem;color:#64748b;margin-top:.1rem}
label{display:block;font-size:.75rem;font-weight:700;color:#94a3b8;
  letter-spacing:.05em;text-transform:uppercase;margin-bottom:.4rem}
input{width:100%;background:#131c2e;border:1px solid #1e2d45;border-radius:8px;
  color:#e2e8f0;padding:.65rem .85rem;font-size:.88rem;outline:none;transition:border .15s}
input:focus{border-color:#f59e0b;box-shadow:0 0 0 2px rgba(245,158,11,.18)}
.mb{margin-bottom:1rem}
.btn{width:100%;background:#f59e0b;border:none;border-radius:8px;color:#000;
  font-weight:800;font-size:.88rem;padding:.72rem;cursor:pointer;
  margin-top:.4rem;transition:background .15s;letter-spacing:.01em}
.btn:hover{background:#fbbf24}
.err{background:#2d000a;border:1px solid #7f1d1d;color:#f87171;
  border-radius:8px;padding:.6rem .9rem;font-size:.8rem;margin-bottom:1rem}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="logo">
      <div class="logo-icon">W</div>
      <div>
        <div class="logo-text">MJR West Agent</div>
        <div class="logo-sub">Industrial Property Intelligence</div>
      </div>
    </div>
    {% with msgs = get_flashed_messages(with_categories=true) %}
    {% for cat, msg in msgs %}<div class="err"><i class="bi bi-exclamation-circle me-1"></i>{{ msg }}</div>{% endfor %}
    {% endwith %}
    <form method="POST" action="/login">
      <div class="mb">
        <label>Username</label>
        <input name="username" type="text" placeholder="admin" autocomplete="username" required/>
      </div>
      <div class="mb">
        <label>Password</label>
        <input name="password" type="password" autocomplete="current-password" required/>
      </div>
      <button type="submit" class="btn"><i class="bi bi-box-arrow-in-right me-1"></i>Sign In</button>
    </form>
  </div>
</div>
</body>
</html>"""

def render_layout(content, title="Dashboard", active=""):
    return render_template_string(_LAYOUT, content=content, title=title, active=active)

# ── Route helpers ─────────────────────────────────────────────────────────────
def status_badge(status):
    m = {"Available": "badge-available", "Leased": "badge-leased",
         "Sold": "badge-sold", "Under Offer": "badge-offer",
         "New": "badge-new", "Contacted": "badge-offer", "Resolved": "badge-leased"}
    cls = m.get(status, "badge-leased")
    return f'<span class="badge {cls}">{status}</span>'

def fmt_currency(val):
    if val is None:
        return "—"
    try:
        return f"${float(val):,.0f}"
    except (ValueError, TypeError):
        return str(val)

def fmt_date(s):
    if not s:
        return "—"
    try:
        return datetime.fromisoformat(str(s).replace("Z", "+00:00")).strftime("%d %b %Y")
    except Exception:
        return str(s)[:10]

# ── Views ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return redirect(url_for("dashboard"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "")
        p = request.form.get("password", "")
        if u == ADMIN_USER and check_password_hash(ADMIN_PASS_HASH, p):
            session["logged_in"] = True
            session["user"] = u
            return redirect(url_for("dashboard"))
        flash("Invalid credentials.", "error")
    return render_template_string(_LOGIN_HTML)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/dashboard")
@login_required
def dashboard():
    total     = sb_count("properties")
    available = sb_count("properties", {"status": "Available"})
    leased    = sb_count("properties", {"status": "Leased"})
    sold      = sb_count("properties", {"status": "Sold"})
    under     = sb_count("properties", {"status": "Under Offer"})
    new_inq   = sb_count("inquiries",  {"status": "New"})
    all_inq   = sb_select("inquiries", order="created_at", limit=30)
    recent_props = sb_select("properties", order="created_at", limit=8)
    vacancy   = round(available / total * 100 if total else 0, 1)
    now_str   = datetime.now().strftime("%A, %d %B %Y")
    hour      = datetime.now().hour
    greeting  = "Good morning" if hour < 12 else ("Good afternoon" if hour < 17 else "Good evening")

    # Score and sort all inquiries for lead pipeline + call list
    scored_inq = sorted(
        [dict(i, _score=score_lead(i)) for i in all_inq],
        key=lambda x: x["_score"], reverse=True
    )
    hot_leads  = [i for i in scored_inq if i["_score"] >= 3][:8]
    call_list  = [i for i in scored_inq if i.get("status") == "New"][:6]

    # Call list HTML
    def initials(name):
        parts = (name or "?").split()
        return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()

    call_items = ""
    for i in call_list:
        nm  = i.get("contact_name") or "Unknown"
        ph  = i.get("contact_phone") or i.get("contact_email") or "—"
        sc  = i["_score"]
        src = i.get("source", "")
        src_icon = {"WhatsApp": "bi-whatsapp", "Email": "bi-envelope-fill",
                    "Phone": "bi-telephone-fill"}.get(src, "bi-chat-left-dots-fill")
        call_items += f"""
<div class="call-item" id="call-{i['id']}">
  <div class="call-avatar">{initials(nm)}</div>
  <div style="flex:1;min-width:0">
    <div class="call-name">{nm}</div>
    <div class="call-detail"><i class="bi {src_icon} me-1"></i>{ph[:28]}</div>
  </div>
  {urgency_badge(sc)}
  <div class="call-actions ms-2">
    <a href="tel:{ph}" class="btn-call"><i class="bi bi-telephone-fill"></i></a>
    <button class="btn-done" onclick="markCalled({i['id']})"><i class="bi bi-check"></i></button>
  </div>
</div>"""
    if not call_items:
        call_items = '<p style="color:#2d4a6b;font-size:.82rem;text-align:center;padding:1.5rem 0">No active leads — <a href="/inquiries" style="color:#f59e0b">view pipeline</a></p>'

    # Hot leads pipeline rows
    pipeline_rows = ""
    for i in hot_leads:
        sc = i["_score"]
        pipeline_rows += f"""<tr>
          <td style="font-weight:700;color:#fff">{i.get('contact_name','—')}</td>
          <td style="color:#64748b;font-size:.76rem">{i.get('source','—')}</td>
          <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
                     color:#94a3b8;font-size:.76rem">{(i.get('message') or '')[:50]}</td>
          <td>{urgency_badge(sc)}</td>
          <td>{fmt_date(i.get('created_at'))}</td>
          <td><a href="/inquiries" class="btn btn-xs btn-outline-secondary">Open</a></td>
        </tr>"""
    if not pipeline_rows:
        pipeline_rows = '<tr><td colspan="6" style="text-align:center;color:#2d4a6b;padding:2rem">No hot leads yet</td></tr>'

    # Properties table
    prop_rows = "".join(f"""<tr>
        <td style="font-weight:700;color:#fff">{p.get('address','—')}</td>
        <td style="color:#64748b">{p.get('property_type','—')}</td>
        <td style="color:#94a3b8">{p.get('size_sqm','—')}</td>
        <td style="color:#f59e0b">{fmt_currency(p.get('asking_rent_pa') or p.get('asking_price'))}</td>
        <td>{status_badge(p.get('status','—'))}</td>
        <td><a href="/properties/{p['id']}/edit" class="btn btn-xs btn-outline-secondary">Edit</a></td>
    </tr>""" for p in recent_props)

    content = f"""
<!-- Morning brief banner -->
<div class="brief-banner">
  <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:.5rem">
    <div>
      <div class="date-str"><i class="bi bi-sun-fill me-1"></i>{now_str}</div>
      <h5>{greeting}, MJR West.</h5>
      <p>
        {new_inq} new lead{'s' if new_inq!=1 else ''} &bull;
        {available} propert{'ies' if available!=1 else 'y'} available &bull;
        {vacancy}% vacancy &bull;
        {len(hot_leads)} hot lead{'s' if len(hot_leads)!=1 else ''} need attention
      </p>
    </div>
    <div class="d-flex gap-2 mt-1">
      <a href="/briefings/send" class="btn btn-sm btn-primary">
        <i class="bi bi-send-fill me-1"></i>Send Morning Brief</a>
      <a href="/email-log/check" class="btn btn-sm btn-outline-secondary">
        <i class="bi bi-arrow-clockwise me-1"></i>Sync Email</a>
    </div>
  </div>
</div>

<!-- KPIs -->
<div class="row g-3 mb-4">
  <div class="col-6 col-lg-3">
    <div class="kpi">
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <div>
          <div class="kpi-val amber">{total}</div>
          <div class="kpi-label">Total Listings</div>
        </div>
        <div class="kpi-icon" style="background:rgba(245,158,11,.12)">
          <i class="bi bi-buildings-fill" style="color:#f59e0b"></i>
        </div>
      </div>
    </div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="kpi">
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <div>
          <div class="kpi-val green">{available}</div>
          <div class="kpi-label">Available</div>
          <div class="kpi-delta" style="color:#10b981">{vacancy}% vacancy rate</div>
        </div>
        <div class="kpi-icon" style="background:rgba(16,185,129,.1)">
          <i class="bi bi-check-circle-fill" style="color:#10b981"></i>
        </div>
      </div>
    </div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="kpi">
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <div>
          <div class="kpi-val blue">{leased}</div>
          <div class="kpi-label">Leased / Sold</div>
          <div class="kpi-delta" style="color:#64748b">{sold} sold &bull; {under} under offer</div>
        </div>
        <div class="kpi-icon" style="background:rgba(56,189,248,.1)">
          <i class="bi bi-key-fill" style="color:#38bdf8"></i>
        </div>
      </div>
    </div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="kpi">
      <div style="display:flex;justify-content:space-between;align-items:flex-start">
        <div>
          <div class="kpi-val red">{new_inq}</div>
          <div class="kpi-label">New Inquiries</div>
          <div class="kpi-delta" style="color:#f97316">{len(hot_leads)} hot leads</div>
        </div>
        <div class="kpi-icon" style="background:rgba(239,68,68,.1)">
          <i class="bi bi-lightning-fill" style="color:#ef4444"></i>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- Main content: call list + market doughnut -->
<div class="row g-3 mb-3">
  <div class="col-lg-4">
    <div class="card h-100">
      <div class="card-header d-flex justify-content-between align-items-center">
        <span><i class="bi bi-telephone-fill me-1" style="color:#f59e0b"></i>Today's Call List</span>
        <a href="/inquiries" class="btn btn-xs btn-outline-secondary">Full Pipeline</a>
      </div>
      <div class="card-body" style="padding:.8rem">
        {call_items}
      </div>
    </div>
  </div>
  <div class="col-lg-4">
    <div class="card h-100">
      <div class="card-header">Market Composition</div>
      <div class="card-body d-flex flex-column align-items-center justify-content-center py-3">
        <canvas id="vacancyChart" style="max-width:180px;max-height:180px"></canvas>
        <div class="mt-3" style="display:grid;grid-template-columns:1fr 1fr;gap:.4rem .9rem;width:100%;padding:0 .5rem">
          <div style="display:flex;align-items:center;gap:.4rem;font-size:.74rem;color:#94a3b8">
            <span style="width:8px;height:8px;border-radius:2px;background:#10b981;flex-shrink:0"></span>Available ({available})
          </div>
          <div style="display:flex;align-items:center;gap:.4rem;font-size:.74rem;color:#94a3b8">
            <span style="width:8px;height:8px;border-radius:2px;background:#38bdf8;flex-shrink:0"></span>Leased ({leased})
          </div>
          <div style="display:flex;align-items:center;gap:.4rem;font-size:.74rem;color:#94a3b8">
            <span style="width:8px;height:8px;border-radius:2px;background:#c084fc;flex-shrink:0"></span>Sold ({sold})
          </div>
          <div style="display:flex;align-items:center;gap:.4rem;font-size:.74rem;color:#94a3b8">
            <span style="width:8px;height:8px;border-radius:2px;background:#f59e0b;flex-shrink:0"></span>Under Offer ({under})
          </div>
        </div>
      </div>
    </div>
  </div>
  <div class="col-lg-4">
    <div class="card h-100">
      <div class="card-header">Scheduled Today</div>
      <div class="card-body" style="padding:.85rem">
        <div style="display:flex;flex-direction:column;gap:.5rem">
          <div style="display:flex;align-items:center;gap:.7rem;padding:.5rem .7rem;
               background:rgba(245,158,11,.06);border:1px solid rgba(245,158,11,.15);border-radius:8px">
            <div style="width:36px;height:36px;background:rgba(245,158,11,.15);border-radius:7px;
                 display:flex;align-items:center;justify-content:center;color:#f59e0b;font-size:.85rem">
              <i class="bi bi-sun-fill"></i></div>
            <div>
              <div style="font-size:.78rem;font-weight:700;color:#fff">Morning Brief</div>
              <div style="font-size:.7rem;color:#64748b">{BRIEFING_HOUR:02d}:{BRIEFING_MINUTE:02d} AEST — WhatsApp broadcast</div>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:.7rem;padding:.5rem .7rem;
               background:rgba(56,189,248,.05);border:1px solid rgba(56,189,248,.1);border-radius:8px">
            <div style="width:36px;height:36px;background:rgba(56,189,248,.1);border-radius:7px;
                 display:flex;align-items:center;justify-content:center;color:#38bdf8;font-size:.85rem">
              <i class="bi bi-envelope-fill"></i></div>
            <div>
              <div style="font-size:.78rem;font-weight:700;color:#fff">Email Sync</div>
              <div style="font-size:.7rem;color:#64748b">Every 15 min — Gmail IMAP</div>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:.7rem;padding:.5rem .7rem;
               background:rgba(16,185,129,.05);border:1px solid rgba(16,185,129,.1);border-radius:8px">
            <div style="width:36px;height:36px;background:rgba(16,185,129,.1);border-radius:7px;
                 display:flex;align-items:center;justify-content:center;color:#10b981;font-size:.85rem">
              <i class="bi bi-calendar-check-fill"></i></div>
            <div>
              <div style="font-size:.78rem;font-weight:700;color:#fff">Weekly Summary</div>
              <div style="font-size:.7rem;color:#64748b">Monday — market transactions</div>
            </div>
          </div>
        </div>
        <a href="/briefings/preview" class="btn btn-sm btn-outline-secondary w-100 mt-3" style="font-size:.74rem">
          <i class="bi bi-eye me-1"></i>Preview Today's Brief</a>
      </div>
    </div>
  </div>
</div>

<!-- Hot Leads Pipeline -->
<div class="card mb-3">
  <div class="card-header d-flex justify-content-between align-items-center">
    <span><i class="bi bi-fire me-1" style="color:#f97316"></i>Hot Lead Pipeline</span>
    <a href="/inquiries" class="btn btn-xs btn-outline-secondary">All Leads</a>
  </div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Contact</th><th>Source</th><th>Note</th><th>Urgency</th><th>Date</th><th></th></tr></thead>
      <tbody>{pipeline_rows}</tbody>
    </table></div>
  </div>
</div>

<!-- Recent Properties -->
<div class="card">
  <div class="card-header d-flex justify-content-between align-items-center">
    <span><i class="bi bi-buildings me-1" style="color:#f59e0b"></i>Recent Listings</span>
    <div class="d-flex gap-2">
      <a href="/upload" class="btn btn-xs btn-outline-secondary">Import Excel</a>
      <a href="/properties" class="btn btn-xs btn-outline-secondary">View All</a>
    </div>
  </div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Address</th><th>Type</th><th>Size</th><th>Price / Rent</th><th>Status</th><th></th></tr></thead>
      <tbody>{prop_rows or '<tr><td colspan="6" style="text-align:center;color:#2d4a6b;padding:2rem">No properties yet — <a href="/upload" style="color:#f59e0b">import Excel</a></td></tr>'}</tbody>
    </table></div>
  </div>
</div>

<script>
Chart.defaults.color = '#64748b';
new Chart(document.getElementById('vacancyChart'), {{
  type:'doughnut',
  data:{{
    labels:['Available','Leased','Sold','Under Offer'],
    datasets:[{{
      data:[{available},{leased},{sold},{under}],
      backgroundColor:['#10b981','#38bdf8','#c084fc','#f59e0b'],
      borderWidth:2,borderColor:'#131c2e',hoverOffset:6
    }}]
  }},
  options:{{cutout:'70%',plugins:{{legend:{{display:false}},
    tooltip:{{callbacks:{{label:function(c){{return ' '+c.label+': '+c.raw}}}}}}
  }}}}
}});
function markCalled(id) {{
  fetch('/api/inquiries/'+id+'/status',{{method:'PATCH',
    headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{status:'Contacted'}})}})
    .then(r=>r.json()).then(()=>{{
      const el=document.getElementById('call-'+id);
      if(el){{el.style.opacity='.3';el.style.transition='opacity .4s';
        setTimeout(()=>el.remove(),450);}}
    }});
}}
</script>
"""
    return render_layout(content, "Dashboard", "dashboard")

@app.route("/properties")
@login_required
def properties_page():
    status_filter = request.args.get("status", "")
    search = request.args.get("q", "")
    props = sb_select("properties", filters={"status": status_filter} if status_filter else None, order="created_at")
    if search:
        props = [p for p in props if search.lower() in (p.get("address") or "").lower()
                 or search.lower() in (p.get("suburb") or "").lower()]
    rows = "".join(f"""<tr>
        <td class="fw-semibold">{p.get('address','—')}</td>
        <td>{p.get('suburb','—')}</td>
        <td>{p.get('property_type','—')}</td>
        <td>{p.get('size_sqm','—')}</td>
        <td>{fmt_currency(p.get('asking_price'))}</td>
        <td>{fmt_currency(p.get('asking_rent_pa'))}</td>
        <td>{status_badge(p.get('status','—'))}</td>
        <td>{p.get('agent_name','—')}</td>
        <td>{fmt_date(p.get('created_at'))}</td>
        <td>
          <a href="/properties/{p['id']}/edit" class="btn btn-sm btn-outline-secondary me-1">Edit</a>
          <button class="btn btn-sm btn-outline-danger" onclick="deleteRow({p['id']},'property')">Del</button>
        </td>
    </tr>""" for p in props)

    content = f"""
<div class="ph">
  <div><h4>Properties</h4><p>{len(props)} listings</p></div>
  <div class="d-flex gap-2">
    <a href="/upload" class="btn btn-sm btn-outline-secondary"><i class="bi bi-upload me-1"></i>Import Excel</a>
    <button class="btn btn-sm btn-primary" data-bs-toggle="modal" data-bs-target="#addPropertyModal">
      <i class="bi bi-plus me-1"></i>Add Property</button>
  </div>
</div>
<div class="card mb-3">
  <div class="card-body py-2">
    <form class="row g-2 align-items-center" method="GET" action="/properties">
      <div class="col-auto"><input name="q" class="form-control form-control-sm" placeholder="Search address…" value="{search}"/></div>
      <div class="col-auto">
        <select name="status" class="form-select form-select-sm" onchange="this.form.submit()">
          <option value="">All Status</option>
          {"".join(f'<option value="{s}" {"selected" if status_filter==s else ""}>{s}</option>' for s in ["Available","Leased","Sold","Under Offer"])}
        </select>
      </div>
      <div class="col-auto"><button class="btn btn-sm btn-primary">Filter</button></div>
    </form>
  </div>
</div>
<div class="card">
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Address</th><th>Suburb</th><th>Type</th><th>Size (sqm)</th><th>Sale Price</th><th>Rent/pa</th><th>Status</th><th>Agent</th><th>Added</th><th></th></tr></thead>
      <tbody>{rows or '<tr><td colspan="10" class="text-center text-muted py-4">No properties found.</td></tr>'}</tbody>
    </table></div>
  </div>
</div>

<!-- Add Property Modal -->
<div class="modal fade" id="addPropertyModal" tabindex="-1">
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
      <div class="modal-header"><h5 class="modal-title">Add Property</h5>
        <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <form method="POST" action="/api/properties">
      <div class="modal-body">
        <div class="row g-3">
          <div class="col-md-8"><label class="form-label fw-semibold">Address *</label>
            <input name="address" class="form-control" required/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Suburb</label>
            <input name="suburb" class="form-control" value="West Melbourne"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Type</label>
            <select name="property_type" class="form-select">
              <option>Warehouse</option><option>Factory</option><option>Office/Warehouse</option>
              <option>Cold Storage</option><option>Land</option><option>Showroom</option>
            </select></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Size (sqm)</label>
            <input name="size_sqm" type="number" class="form-control"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Land (sqm)</label>
            <input name="land_sqm" type="number" class="form-control"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Asking Price</label>
            <input name="asking_price" type="number" class="form-control"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Rent/pa</label>
            <input name="asking_rent_pa" type="number" class="form-control"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Status</label>
            <select name="status" class="form-select">
              <option>Available</option><option>Under Offer</option>
              <option>Leased</option><option>Sold</option>
            </select></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Year Built</label>
            <input name="year_built" type="number" class="form-control"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Zoning</label>
            <input name="zoning" class="form-control" placeholder="e.g. Industrial 1"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Agent Name</label>
            <input name="agent_name" class="form-control"/></div>
          <div class="col-md-4"><label class="form-label fw-semibold">Agent Phone</label>
            <input name="agent_phone" class="form-control"/></div>
          <div class="col-12"><label class="form-label fw-semibold">Notes</label>
            <textarea name="notes" class="form-control" rows="2"></textarea></div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
        <button type="submit" class="btn btn-primary">Save Property</button>
      </div>
      </form>
    </div>
  </div>
</div>
<script>
function deleteRow(id, type) {{
  if (!confirm('Delete this record?')) return;
  fetch('/api/'+type+'s/'+id, {{method:'DELETE'}})
    .then(r=>r.json()).then(d=>{{ if(d.ok) location.reload(); else alert(d.error); }});
}}
</script>
"""
    return render_layout(content, "Properties", "properties")

@app.route("/properties/<int:pid>/edit", methods=["GET", "POST"])
@login_required
def edit_property(pid):
    if request.method == "POST":
        data = {k: v for k, v in request.form.items() if v != ""}
        for f in ("size_sqm","land_sqm","asking_price","asking_rent_pa","year_built"):
            if f in data:
                try: data[f] = float(data[f])
                except ValueError: data.pop(f)
        data["updated_at"] = datetime.now(timezone.utc).isoformat()
        sb_update("properties", {"id": pid}, data)
        flash("Property updated.", "success")
        return redirect(url_for("properties_page"))
    props = sb_select("properties", filters={"id": pid})
    p = props[0] if props else {}
    def val(k): return p.get(k, "") or ""
    def sel(field, opt):
        return "selected" if str(val(field)) == str(opt) else ""
    content = f"""
<div class="ph"><div><h4>Edit Property</h4><p>{val('address')}</p></div>
<div class="card"><div class="card-body">
<form method="POST" action="/properties/{pid}/edit">
  <div class="row g-3">
    <div class="col-md-8"><label class="form-label fw-semibold">Address *</label>
      <input name="address" class="form-control" value="{val('address')}" required/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Suburb</label>
      <input name="suburb" class="form-control" value="{val('suburb')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Type</label>
      <select name="property_type" class="form-select">
        {"".join(f'<option {sel("property_type",t)}>{t}</option>' for t in ["Warehouse","Factory","Office/Warehouse","Cold Storage","Land","Showroom"])}
      </select></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Size (sqm)</label>
      <input name="size_sqm" type="number" class="form-control" value="{val('size_sqm')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Land (sqm)</label>
      <input name="land_sqm" type="number" class="form-control" value="{val('land_sqm')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Asking Price</label>
      <input name="asking_price" type="number" class="form-control" value="{val('asking_price')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Rent/pa</label>
      <input name="asking_rent_pa" type="number" class="form-control" value="{val('asking_rent_pa')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Status</label>
      <select name="status" class="form-select">
        {"".join(f'<option {sel("status",s)}>{s}</option>' for s in ["Available","Under Offer","Leased","Sold"])}
      </select></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Year Built</label>
      <input name="year_built" type="number" class="form-control" value="{val('year_built')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Zoning</label>
      <input name="zoning" class="form-control" value="{val('zoning')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Agent Name</label>
      <input name="agent_name" class="form-control" value="{val('agent_name')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Agent Phone</label>
      <input name="agent_phone" class="form-control" value="{val('agent_phone')}"/></div>
    <div class="col-md-4"><label class="form-label fw-semibold">Agent Email</label>
      <input name="agent_email" type="email" class="form-control" value="{val('agent_email')}"/></div>
    <div class="col-12"><label class="form-label fw-semibold">Notes</label>
      <textarea name="notes" class="form-control" rows="3">{val('notes')}</textarea></div>
    <div class="col-12 d-flex gap-2">
      <button type="submit" class="btn btn-primary">Save Changes</button>
      <a href="/properties" class="btn btn-outline-secondary">Cancel</a>
    </div>
  </div>
</form>
</div></div>"""
    return render_layout(content, "Edit Property", "properties")

@app.route("/contacts")
@login_required
def contacts_page():
    contacts = sb_select("contacts", order="created_at")
    rows = "".join(f"""<tr>
        <td class="fw-semibold">{c.get('name','—')}</td>
        <td>{c.get('company','—')}</td>
        <td>{c.get('contact_type','—')}</td>
        <td>{c.get('phone','—')}</td>
        <td>{c.get('email','—')}</td>
        <td>{'✅' if c.get('whatsapp_opt_in') else '—'}</td>
        <td>{fmt_date(c.get('created_at'))}</td>
        <td><button class="btn btn-sm btn-outline-danger" onclick="deleteRow({c['id']},'contact')">Del</button></td>
    </tr>""" for c in contacts)
    content = f"""
<div class="ph">
  <div><h4>Contacts</h4><p>{len(contacts)} records</p></div>
  <button class="btn btn-sm btn-primary" data-bs-toggle="modal" data-bs-target="#addContactModal">
    <i class="bi bi-plus me-1"></i>Add Contact</button>
</div>
<div class="card">
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Name</th><th>Company</th><th>Type</th><th>Phone</th><th>Email</th><th>WA Opt-In</th><th>Added</th><th></th></tr></thead>
      <tbody>{rows or '<tr><td colspan="8" class="text-center text-muted py-4">No contacts yet.</td></tr>'}</tbody>
    </table></div>
  </div>
</div>
<!-- Add Contact Modal -->
<div class="modal fade" id="addContactModal" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header"><h5 class="modal-title">Add Contact</h5>
        <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
      <form method="POST" action="/api/contacts">
      <div class="modal-body">
        <div class="row g-3">
          <div class="col-12"><label class="form-label fw-semibold">Name *</label>
            <input name="name" class="form-control" required/></div>
          <div class="col-md-6"><label class="form-label fw-semibold">Company</label>
            <input name="company" class="form-control"/></div>
          <div class="col-md-6"><label class="form-label fw-semibold">Type</label>
            <select name="contact_type" class="form-select">
              <option>Agent</option><option>Buyer</option><option>Tenant</option>
              <option>Owner</option><option>Developer</option>
            </select></div>
          <div class="col-md-6"><label class="form-label fw-semibold">Phone</label>
            <input name="phone" class="form-control"/></div>
          <div class="col-md-6"><label class="form-label fw-semibold">Email</label>
            <input name="email" type="email" class="form-control"/></div>
          <div class="col-12">
            <div class="form-check">
              <input class="form-check-input" type="checkbox" name="whatsapp_opt_in" value="true" id="waOptIn"/>
              <label class="form-check-label" for="waOptIn">WhatsApp Opt-In</label>
            </div>
          </div>
          <div class="col-12"><label class="form-label fw-semibold">Notes</label>
            <textarea name="notes" class="form-control" rows="2"></textarea></div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
        <button type="submit" class="btn btn-primary">Save</button>
      </div>
      </form>
    </div>
  </div>
</div>
<script>
function deleteRow(id, type) {{
  if (!confirm('Delete?')) return;
  fetch('/api/'+type+'s/'+id, {{method:'DELETE'}}).then(r=>r.json()).then(d=>{{if(d.ok) location.reload();}});
}}
</script>"""
    return render_layout(content, "Contacts", "contacts")

@app.route("/inquiries")
@login_required
def inquiries_page():
    inquiries = sb_select("inquiries", order="created_at")
    scored = sorted([dict(i, _score=score_lead(i)) for i in inquiries],
                    key=lambda x: x["_score"], reverse=True)
    rows = "".join(f"""<tr>
        <td style="font-weight:700;color:#fff">{i.get('contact_name','—')}</td>
        <td style="color:#64748b;font-size:.77rem">{i.get('contact_email','—')}</td>
        <td style="color:#94a3b8">{i.get('contact_phone','—')}</td>
        <td>{i.get('source','—')}</td>
        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#94a3b8;font-size:.77rem">{(i.get('message') or '')[:60]}</td>
        <td>{urgency_badge(i['_score'])}</td>
        <td>{status_badge(i.get('status','—'))}</td>
        <td>{fmt_date(i.get('created_at'))}</td>
        <td>
          <button class="btn btn-xs btn-outline-secondary me-1" onclick="markDone({i['id']})">Done</button>
          <button class="btn btn-xs" style="background:#2d000a;color:#f87171;border:1px solid #7f1d1d;border-radius:5px" onclick="deleteRow({i['id']},'inquir')">Del</button>
        </td>
    </tr>""" for i in scored)
    fire_count = sum(1 for i in scored if i["_score"] >= 3)
    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-lightning-fill me-2" style="color:#f59e0b"></i>Lead Pipeline</h4>
    <p>{len(inquiries)} total &bull; <span style="color:#f97316;font-weight:700">{fire_count} hot leads</span></p>
  </div>
  <button class="btn btn-sm btn-primary" data-bs-toggle="modal" data-bs-target="#addInqModal">
    <i class="bi bi-plus me-1"></i>Add Lead</button>
</div>
<div class="card">
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Name</th><th>Email</th><th>Phone</th><th>Source</th><th>Message</th><th>Urgency</th><th>Status</th><th>Date</th><th></th></tr></thead>
      <tbody>{rows or '<tr><td colspan="9" style="text-align:center;color:#2d4a6b;padding:2rem">No leads yet.</td></tr>'}</tbody>
    </table></div>
  </div>
</div>
<!-- Add Inquiry Modal -->
<div class="modal fade" id="addInqModal" tabindex="-1">
  <div class="modal-dialog"><div class="modal-content">
    <div class="modal-header"><h5 class="modal-title">Add Inquiry</h5>
      <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
    <form method="POST" action="/api/inquiries">
    <div class="modal-body">
      <div class="row g-3">
        <div class="col-md-6"><label class="form-label fw-semibold">Contact Name</label>
          <input name="contact_name" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Contact Email</label>
          <input name="contact_email" type="email" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Phone</label>
          <input name="contact_phone" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Source</label>
          <select name="source" class="form-select">
            <option>Web</option><option>WhatsApp</option><option>Email</option><option>Phone</option><option>Referral</option>
          </select></div>
        <div class="col-12"><label class="form-label fw-semibold">Message</label>
          <textarea name="message" class="form-control" rows="3"></textarea></div>
      </div>
    </div>
    <div class="modal-footer">
      <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
      <button type="submit" class="btn btn-primary">Save</button>
    </div>
    </form>
  </div></div>
</div>
<script>
function deleteRow(id, type) {{
  if (!confirm('Delete?')) return;
  fetch('/api/'+type+'ies/'+id, {{method:'DELETE'}}).then(r=>r.json()).then(d=>{{if(d.ok) location.reload();}});
}}
function markDone(id) {{
  fetch('/api/inquiries/'+id+'/status', {{method:'PATCH', headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{status:'Resolved'}})}}).then(r=>r.json()).then(d=>{{if(d.ok) location.reload();}});
}}
</script>"""
    return render_layout(content, "Inquiries", "inquiries")

@app.route("/deals")
@login_required
def deals_page():
    deals = sb_select("deals", order="created_at")
    rows = "".join(f"""<tr>
        <td>{d.get('deal_type','—')}</td>
        <td>{fmt_currency(d.get('price'))}</td>
        <td>{d.get('buyer_name','—')}</td>
        <td>{d.get('seller_name','—')}</td>
        <td>{d.get('agent_name','—')}</td>
        <td>{status_badge(d.get('status','—'))}</td>
        <td>{fmt_date(d.get('settlement_date'))}</td>
        <td>{fmt_date(d.get('created_at'))}</td>
        <td><button class="btn btn-sm btn-outline-danger" onclick="deleteDeal({d['id']})">Del</button></td>
    </tr>""" for d in deals)
    content = f"""
<div class="ph">
  <div><h4>Deals</h4><p>{len(deals)} transactions</p></div>
  <button class="btn btn-sm btn-primary" data-bs-toggle="modal" data-bs-target="#addDealModal">
    <i class="bi bi-plus me-1"></i>Add Deal</button>
</div>
<div class="card">
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Type</th><th>Price</th><th>Buyer</th><th>Seller</th><th>Agent</th><th>Status</th><th>Settlement</th><th>Added</th><th></th></tr></thead>
      <tbody>{rows or '<tr><td colspan="9" class="text-center text-muted py-4">No deals yet.</td></tr>'}</tbody>
    </table></div>
  </div>
</div>
<div class="modal fade" id="addDealModal" tabindex="-1">
  <div class="modal-dialog"><div class="modal-content">
    <div class="modal-header"><h5 class="modal-title">Add Deal</h5>
      <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
    <form method="POST" action="/api/deals">
    <div class="modal-body">
      <div class="row g-3">
        <div class="col-md-6"><label class="form-label fw-semibold">Type</label>
          <select name="deal_type" class="form-select"><option>Sale</option><option>Lease</option></select></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Price/Rent</label>
          <input name="price" type="number" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Buyer / Lessee</label>
          <input name="buyer_name" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Seller / Lessor</label>
          <input name="seller_name" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Agent</label>
          <input name="agent_name" class="form-control"/></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Status</label>
          <select name="status" class="form-select">
            <option>Negotiation</option><option>Unconditional</option><option>Settled</option>
          </select></div>
        <div class="col-md-6"><label class="form-label fw-semibold">Settlement Date</label>
          <input name="settlement_date" type="date" class="form-control"/></div>
        <div class="col-12"><label class="form-label fw-semibold">Notes</label>
          <textarea name="notes" class="form-control" rows="2"></textarea></div>
      </div>
    </div>
    <div class="modal-footer">
      <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
      <button type="submit" class="btn btn-primary">Save</button>
    </div>
    </form>
  </div></div>
</div>
<script>
function deleteDeal(id) {{
  if(!confirm('Delete?')) return;
  fetch('/api/deals/'+id,{{method:'DELETE'}}).then(r=>r.json()).then(d=>{{if(d.ok) location.reload();}});
}}
</script>"""
    return render_layout(content, "Deals", "deals")

@app.route("/whatsapp")
@login_required
def whatsapp_page():
    logs = sb_select("whatsapp_logs", order="created_at", limit=50)
    rows = "".join(f"""<tr>
        <td><span class="badge {'bg-primary' if l.get('direction')=='Inbound' else 'bg-success'}">{l.get('direction','—')}</span></td>
        <td>{l.get('from_number','—')}</td>
        <td>{l.get('to_number','—')}</td>
        <td style="max-width:300px" class="text-truncate">{l.get('body','—')[:100]}</td>
        <td>{fmt_date(l.get('created_at'))}</td>
    </tr>""" for l in logs)
    broadcast_list = ", ".join(WA_BROADCAST_LIST) if WA_BROADCAST_LIST else "None configured"
    content = f"""
<div class="ph">
  <div><h4>WhatsApp</h4><p>Webhook endpoint: <code>/webhook/whatsapp</code></p></div>
  <button class="btn btn-sm btn-primary" data-bs-toggle="modal" data-bs-target="#sendMsgModal">
    <i class="bi bi-send me-1"></i>Send Message</button>
</div>
<div class="row g-3 mb-4">
  <div class="col-md-4">
    <div class="kpi">
      <div class="kpi-val">{len(logs)}</div>
      <div class="kpi-label">Messages (last 50)</div>
    </div>
  </div>
  <div class="col-md-4">
    <div class="kpi">
      <div class="kpi-val">{sum(1 for l in logs if l.get('direction')=='Inbound')}</div>
      <div class="kpi-label">Inbound</div>
    </div>
  </div>
  <div class="col-md-4">
    <div class="kpi">
      <div class="kpi-val">{len(WA_BROADCAST_LIST)}</div>
      <div class="kpi-label">Broadcast Recipients</div>
    </div>
  </div>
</div>
<div class="alert alert-info d-flex align-items-center mb-3" style="font-size:.85rem">
  <i class="bi bi-info-circle-fill me-2"></i>
  Broadcast list: <strong class="ms-1">{broadcast_list}</strong>
</div>
<div class="card">
  <div class="card-header">Message Log</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Direction</th><th>From</th><th>To</th><th>Message</th><th>Date</th></tr></thead>
      <tbody>{rows or '<tr><td colspan="5" class="text-center text-muted py-4">No messages yet.</td></tr>'}</tbody>
    </table></div>
  </div>
</div>
<!-- Send Modal -->
<div class="modal fade" id="sendMsgModal" tabindex="-1">
  <div class="modal-dialog"><div class="modal-content">
    <div class="modal-header"><h5 class="modal-title">Send WhatsApp Message</h5>
      <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
    <form method="POST" action="/api/whatsapp/send">
    <div class="modal-body">
      <div class="mb-3"><label class="form-label fw-semibold">To (number with country code)</label>
        <input name="to" class="form-control" placeholder="+61412345678"/></div>
      <div class="mb-3"><label class="form-label fw-semibold">Message</label>
        <textarea name="body" class="form-control" rows="4" required></textarea></div>
      <div class="form-check">
        <input class="form-check-input" type="checkbox" name="broadcast" id="broadcastChk"/>
        <label class="form-check-label" for="broadcastChk">Send to all broadcast recipients instead</label>
      </div>
    </div>
    <div class="modal-footer">
      <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
      <button type="submit" class="btn btn-primary">Send</button>
    </div>
    </form>
  </div></div>
</div>"""
    return render_layout(content, "WhatsApp", "whatsapp")

@app.route("/email-log")
@login_required
def email_page():
    all_emails = sb_select("email_logs", order="received_at", limit=200) or []

    total      = len(all_emails)
    unanswered = sum(1 for e in all_emails if e.get("reply_status") == "pending" and e.get("requires_action"))
    overdue    = sum(1 for e in all_emails if e.get("flagged_unanswered"))
    hi_pri     = sum(1 for e in all_emails if e.get("ai_priority") in ("critical", "high"))

    def _pri_order(e):
        return {"critical": 0, "high": 1, "medium": 2, "low": 3}.get(e.get("ai_priority", "medium"), 2)

    sorted_emails = sorted(all_emails, key=_pri_order)

    def _badge_source(src):
        src = (src or "").lower()
        if "gmail" in src:
            return '<span style="font-size:.6rem;padding:.1rem .35rem;border-radius:4px;background:#ea4335;color:#fff">Gmail</span>'
        if "android" in src:
            return '<span style="font-size:.6rem;padding:.1rem .35rem;border-radius:4px;background:#3ddc84;color:#000">Android</span>'
        if "outlook" in src:
            return '<span style="font-size:.6rem;padding:.1rem .35rem;border-radius:4px;background:#0078d4;color:#fff">Outlook</span>'
        return f'<span style="font-size:.6rem;padding:.1rem .35rem;border-radius:4px;background:rgba(255,255,255,.1);color:#aaa">{src or "?"}</span>'

    def _priority_pill(p):
        cols = {"critical": ("#ef4444", "#fff"), "high": ("#f97316", "#fff"),
                "medium": ("#eab308", "#000"), "low": ("#6b7280", "#fff")}
        bg, fg = cols.get(p or "medium", ("#6b7280", "#fff"))
        return (f'<span style="font-size:.6rem;padding:.1rem .4rem;border-radius:4px;'
                f'background:{bg};color:{fg};font-weight:600;text-transform:uppercase">'
                f'{(p or "med")[:4]}</span>')

    def _email_row(e):
        eid        = e.get("id", "")
        sender     = (e.get("sender_name") or e.get("sender") or "Unknown")[:30]
        subject    = (e.get("subject") or "(no subject)")[:60]
        summary    = (e.get("ai_summary") or (e.get("body") or "")[:100])[:100]
        received   = fmt_date(e.get("received_at"))
        replied    = e.get("reply_status") == "replied"
        flagged    = bool(e.get("flagged_unanswered"))
        saved      = bool(e.get("save_triggered"))
        pri        = e.get("ai_priority", "medium")
        items_raw  = e.get("action_items") or "[]"
        try:
            items = json.loads(items_raw) if isinstance(items_raw, str) else (items_raw or [])
        except Exception:
            items = []
        items_html = ""
        if items:
            bullets = "".join(f'<li style="font-size:.68rem;color:#94a3b8">{ai}</li>' for ai in items[:3])
            items_html = f'<ul style="margin:.3rem 0 0 1rem;padding:0">{bullets}</ul>'

        overdue_badge = ('<span style="font-size:.6rem;padding:.1rem .35rem;border-radius:4px;'
                         'background:#ef4444;color:#fff;margin-left:.3rem">OVERDUE</span>' if flagged else "")
        save_badge    = ('<span style="font-size:.6rem;padding:.1rem .35rem;border-radius:4px;'
                         'background:#7c3aed;color:#fff;margin-left:.3rem">SAVED</span>' if saved else "")
        reply_btn = ""
        if not replied and e.get("requires_action"):
            reply_btn = (f'<button onclick="markReplied({eid}, this)" '
                         f'class="btn btn-sm" style="font-size:.65rem;padding:.15rem .5rem;'
                         f'background:rgba(16,185,129,.15);color:#10b981;border:1px solid #10b981">'
                         f'Mark Replied</button>')
        replied_label = ('<span style="font-size:.65rem;color:#10b981">Replied</span>' if replied else "")
        row_border = "border-left:3px solid #ef4444" if pri == "critical" else (
                     "border-left:3px solid #f97316" if pri == "high" else
                     "border-left:3px solid rgba(255,255,255,.05)")
        return f"""
<div style="background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.07);
            {row_border};border-radius:8px;padding:.75rem 1rem;margin-bottom:.5rem">
  <div style="display:flex;align-items:center;gap:.5rem;flex-wrap:wrap;margin-bottom:.25rem">
    {_priority_pill(pri)}
    {_badge_source(e.get("source"))}
    {overdue_badge}{save_badge}
    <span style="font-weight:600;font-size:.82rem;flex:1;min-width:0;overflow:hidden;
                 text-overflow:ellipsis;white-space:nowrap">{subject}</span>
    <span style="font-size:.7rem;color:#64748b;white-space:nowrap">{received}</span>
  </div>
  <div style="font-size:.75rem;color:#94a3b8;margin-bottom:.25rem">
    <i class="bi bi-person-fill me-1"></i>{sender}
  </div>
  <div style="font-size:.73rem;color:#64748b">{summary}</div>
  {items_html}
  <div style="margin-top:.4rem;display:flex;align-items:center;gap:.5rem">
    {reply_btn}{replied_label}
  </div>
</div>"""

    sections = [
        ("critical", "Critical", "#ef4444"),
        ("high",     "High Priority", "#f97316"),
        ("medium",   "Medium", "#eab308"),
        ("low",      "Low / FYI", "#6b7280"),
    ]
    sections_html = ""
    for pri_key, pri_label, pri_col in sections:
        group = [e for e in sorted_emails if (e.get("ai_priority") or "medium") == pri_key]
        if not group:
            continue
        rows_html = "".join(_email_row(e) for e in group)
        sections_html += f"""
<div style="margin-bottom:1.5rem">
  <div style="display:flex;align-items:center;gap:.5rem;margin-bottom:.6rem">
    <span style="width:10px;height:10px;border-radius:50%;background:{pri_col};display:inline-block"></span>
    <span style="font-weight:600;font-size:.85rem;color:{pri_col}">{pri_label}</span>
    <span style="font-size:.72rem;color:#64748b">{len(group)} email{'s' if len(group)!=1 else ''}</span>
  </div>
  {rows_html}
</div>"""

    content = f"""
<div class="ph">
  <div>
    <h4>Email Inbox</h4>
    <p>Priority-sorted inbox — Gmail IMAP + Android automation + Outlook</p>
  </div>
  <a href="/email-log/check" class="btn btn-sm btn-outline-secondary">
    <i class="bi bi-arrow-clockwise me-1"></i>Check Gmail Now</a>
</div>

<!-- KPI strip -->
<div class="row g-3 mb-4">
  <div class="col-6 col-md-3">
    <div class="kpi"><div class="kpi-val">{total}</div><div class="kpi-label">Total Emails</div></div>
  </div>
  <div class="col-6 col-md-3">
    <div class="kpi"><div class="kpi-val" style="color:#f97316">{unanswered}</div><div class="kpi-label">Awaiting Reply</div></div>
  </div>
  <div class="col-6 col-md-3">
    <div class="kpi"><div class="kpi-val" style="color:#ef4444">{overdue}</div><div class="kpi-label">Overdue (&gt;24h)</div></div>
  </div>
  <div class="col-6 col-md-3">
    <div class="kpi"><div class="kpi-val" style="color:#a78bfa">{hi_pri}</div><div class="kpi-label">High Priority</div></div>
  </div>
</div>

<!-- Natural-language query bar -->
<div class="card mb-4">
  <div class="card-body" style="padding:.75rem 1rem">
    <div style="display:flex;gap:.5rem;align-items:center">
      <i class="bi bi-search" style="color:#64748b"></i>
      <input id="emailQuery" type="text" class="form-control form-control-sm"
             placeholder="Ask: 'who haven't I responded to' · 'priorities today' · 'from John Smith' · 'overdue'"
             style="background:rgba(255,255,255,.05);border-color:rgba(255,255,255,.12);color:#e2e8f0"
             onkeydown="if(event.key==='Enter')queryEmails()">
      <button class="btn btn-sm btn-primary" onclick="queryEmails()">Ask</button>
      <button class="btn btn-sm btn-outline-secondary" onclick="clearQuery()">Clear</button>
    </div>
    <div id="queryAnswer" style="margin-top:.5rem;font-size:.78rem;color:#10b981;display:none"></div>
  </div>
</div>

<!-- Priority sections -->
<div id="emailSections">
  {sections_html or '<div class="text-center text-muted py-5">No emails logged yet. Click Check Gmail Now to fetch.</div>'}
</div>

<!-- Query results container (hidden until query runs) -->
<div id="queryResults" style="display:none">
  <div style="display:flex;align-items:center;gap:.5rem;margin-bottom:.75rem">
    <span style="font-weight:600;font-size:.85rem">Query Results</span>
  </div>
  <div id="queryResultRows"></div>
</div>

<script>
function markReplied(id, btn) {{
  btn.disabled = true;
  btn.textContent = 'Saving…';
  fetch('/api/email/' + id + '/replied', {{method:'PATCH', headers:{{'X-CSRFToken':''}}}})
    .then(r => r.json())
    .then(() => {{
      btn.textContent = 'Replied';
      btn.style.background = 'rgba(16,185,129,.05)';
      btn.style.color = '#64748b';
      btn.style.borderColor = '#64748b';
    }})
    .catch(() => {{ btn.disabled = false; btn.textContent = 'Mark Replied'; }});
}}

function queryEmails() {{
  const q = document.getElementById('emailQuery').value.trim();
  if (!q) return;
  const ans = document.getElementById('queryAnswer');
  ans.style.display = 'block';
  ans.textContent = 'Searching…';
  fetch('/api/email/query', {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{query: q}})
  }})
  .then(r => r.json())
  .then(data => {{
    ans.textContent = data.answer || '';
    document.getElementById('emailSections').style.display = 'none';
    const container = document.getElementById('queryResults');
    const rows = document.getElementById('queryResultRows');
    container.style.display = 'block';
    if (!data.emails || data.emails.length === 0) {{
      rows.innerHTML = '<div class="text-muted" style="font-size:.8rem;padding:1rem">No matching emails.</div>';
      return;
    }}
    rows.innerHTML = data.emails.map(e => `
      <div style="background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.07);
                  border-radius:8px;padding:.65rem 1rem;margin-bottom:.4rem">
        <div style="display:flex;align-items:center;gap:.5rem;flex-wrap:wrap">
          <span style="font-weight:600;font-size:.8rem;flex:1">${{e.subject || '(no subject)'}}</span>
          <span style="font-size:.7rem;color:#64748b">${{e.received_at ? e.received_at.slice(0,16).replace('T',' ') : ''}}</span>
        </div>
        <div style="font-size:.72rem;color:#94a3b8;margin-top:.15rem">${{e.sender_name || e.sender || ''}}</div>
        <div style="font-size:.7rem;color:#64748b;margin-top:.1rem">${{e.ai_summary || ''}}</div>
      </div>`).join('');
  }})
  .catch(() => {{ ans.textContent = 'Query failed. Please try again.'; }});
}}

function clearQuery() {{
  document.getElementById('emailQuery').value = '';
  document.getElementById('queryAnswer').style.display = 'none';
  document.getElementById('queryResults').style.display = 'none';
  document.getElementById('emailSections').style.display = 'block';
}}
</script>"""
    return render_layout(content, "Email Inbox", "email")


@app.route("/email-log/check")
@login_required
def email_check_now():
    t = threading.Thread(target=fetch_gmail_emails)
    t.daemon = True
    t.start()
    flash("Gmail check triggered in background.", "success")
    return redirect(url_for("email_page"))

@app.route("/api/android-email", methods=["POST"])
def android_email_webhook():
    """
    Webhook for Android automation (Tasker / MacroDroid / HTTP Shortcuts).
    Accepts JSON or form data with fields: subject, from, body, timestamp, app.
    Optional header: X-Agent-Token matching SECRET_KEY for lightweight auth.
    """
    token = request.headers.get("X-Agent-Token", "")
    if token and token != SECRET_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    if request.is_json:
        data = request.get_json(force=True) or {}
    else:
        data = request.form.to_dict()

    sender  = data.get("from") or data.get("sender") or "Android <unknown>"
    subject = data.get("subject") or "(no subject)"
    body    = data.get("body") or data.get("text") or data.get("content") or ""
    ts      = data.get("timestamp") or data.get("date") or None
    source  = "android_" + (data.get("app") or "automation").lower().replace(" ", "_")

    # Normalise timestamp — accept epoch millis, epoch seconds, or ISO string
    received_at = None
    if ts:
        try:
            ts_str = str(ts)
            if re.match(r"^\d{13}$", ts_str):
                received_at = datetime.fromtimestamp(int(ts_str)/1000, tz=timezone.utc).isoformat()
            elif re.match(r"^\d{10}$", ts_str):
                received_at = datetime.fromtimestamp(int(ts_str), tz=timezone.utc).isoformat()
            else:
                received_at = ts_str
        except Exception:
            pass

    result = process_email_full(
        sender=sender,
        subject=subject,
        body=body,
        message_id=None,
        source=source,
        received_at=received_at,
    )
    if result is None:
        return jsonify({"status": "duplicate", "message": "Already processed"}), 200
    return jsonify({"status": "ok", "id": result.get("id"), "priority": result.get("ai_priority")}), 201


@app.route("/api/email/<int:email_id>/replied", methods=["PATCH"])
@login_required
def mark_email_replied(email_id):
    """Mark an email as replied and clear the unanswered flag."""
    if get_sb() is None:
        return jsonify({"error": "Database not configured"}), 503
    result = sb_update("email_logs", {"id": email_id}, {
        "reply_status":       "replied",
        "replied_at":         datetime.now(timezone.utc).isoformat(),
        "flagged_unanswered": False,
    })
    if result is None:
        return jsonify({"error": "Update failed — check server logs"}), 500
    return jsonify({"status": "ok", "id": email_id})


@app.route("/api/email/query", methods=["POST"])
@login_required
def email_query_api():
    """Natural-language query over email_logs. POST {query: str}"""
    data  = request.get_json(force=True) or {}
    query = data.get("query", "")
    result = get_email_priorities(query)
    return jsonify(result)


@app.route("/api/calendar/<int:event_id>/ai-results")
@login_required
def calendar_ai_results(event_id):
    """Return AI-processed meeting notes for a calendar event."""
    evs = sb_select("calendar_events", {"id": event_id})
    if not evs:
        return jsonify({"error": "Not found"}), 404
    ev = evs[0]
    items = ev.get("ai_action_items") or []
    if isinstance(items, str):
        try:   items = json.loads(items)
        except Exception: items = []
    return jsonify({
        "action_items":    items,
        "follow_up_draft": ev.get("follow_up_draft") or "",
        "summary":         ev.get("post_meeting_notes", "")[:200],
        "processed_at":    ev.get("notes_processed_at") or "",
    })


@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload_page():
    ai_result  = None
    xls_result = None

    if request.method == "POST":
        mode = request.form.get("mode", "ai")
        f = request.files.get("file")
        if not f or f.filename == "":
            flash("No file selected.", "error")
        elif not allowed_file(f.filename):
            flash("File type not supported.", "error")
        else:
            fname = secure_filename(f.filename)
            fpath = os.path.join(UPLOAD_FOLDER, fname)
            f.save(fpath)
            try:
                ext = fname.rsplit(".", 1)[-1].lower()
                if mode == "excel" or ext in {"xlsx", "xls"}:
                    xls_result = process_excel_file(fpath)
                    flash(f"Excel import: {xls_result['imported']} properties added, {xls_result['skipped']} skipped.", "success")
                else:
                    ai_result = process_file_universal(fpath, fname)
                    flash(f"AI intake complete — {ai_result.get('ai_classification','?')} detected.", "success")
            except Exception as e:
                log.exception("upload error")
                flash(f"Processing error: {e}", "error")
            finally:
                try: os.remove(fpath)
                except OSError: pass

    # ── Build result panel ────────────────────────────────────────────────────
    result_html = ""
    if ai_result:
        conf_pct = int((ai_result.get("ai_confidence") or 0) * 100)
        conf_col = "#10b981" if conf_pct >= 70 else ("#f59e0b" if conf_pct >= 40 else "#ef4444")
        counts   = ai_result.get("counts") or {}
        links    = ai_result.get("links") or {}

        # Counts strip — only show non-zero
        count_labels = {
            "properties": ("bi-buildings-fill","#f59e0b","Properties"),
            "contacts":   ("bi-people-fill","#38bdf8","Contacts"),
            "inquiries":  ("bi-lightning-fill","#ef4444","Inquiries"),
            "deals":      ("bi-handshake-fill","#c084fc","Deals"),
            "vacancies":  ("bi-door-open-fill","#10b981","Vacancies"),
            "requirements":("bi-search","#fb923c","Requirements"),
            "market_data":("bi-graph-up","#a78bfa","Market Data"),
        }
        count_chips = ""
        for k, (icon, col, lbl) in count_labels.items():
            n = counts.get(k, 0)
            if n:
                count_chips += (f'<span style="display:flex;align-items:center;gap:.3rem;'
                                f'background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);'
                                f'border-radius:6px;padding:.2rem .55rem;font-size:.72rem;color:{col}">'
                                f'<i class="bi {icon}"></i>{n} {lbl}</span>')

        companies = ai_result.get("mentioned_companies") or []
        company_chips = " ".join(
            f'<span class="badge" style="background:rgba(56,189,248,.1);color:#38bdf8;font-size:.62rem">{c}</span>'
            for c in companies[:8]
        )
        link_line = ""
        if links.get("document_ids"):
            link_line = (f'<div style="font-size:.7rem;color:#64748b;margin-top:.4rem">'
                         f'<i class="bi bi-link-45deg me-1"></i>Linked to '
                         f'{len(links["document_ids"])} related document(s) via shared companies</div>')

        result_html = f"""
<div class="card mt-3">
  <div class="card-header" style="background:rgba(16,185,129,.08);border-color:rgba(16,185,129,.2)">
    <i class="bi bi-cpu-fill me-1" style="color:#10b981"></i>AI Analysis Complete
  </div>
  <div class="card-body" style="font-size:.82rem">
    <div class="row g-2 mb-2">
      <div class="col-6">
        <div style="color:#64748b;font-size:.7rem">Type Detected</div>
        <div style="color:#f59e0b;font-weight:700">{(ai_result.get('ai_classification') or '—').replace('_',' ').title()}</div>
      </div>
      <div class="col-6">
        <div style="color:#64748b;font-size:.7rem">Confidence</div>
        <div style="color:{conf_col};font-weight:700">{conf_pct}%</div>
      </div>
    </div>
    <div style="color:#94a3b8;font-size:.78rem;margin-bottom:.7rem">{ai_result.get('ai_summary','')}</div>
    {f'<div style="display:flex;gap:.3rem;flex-wrap:wrap;margin-bottom:.6rem">{count_chips}</div>' if count_chips else ''}
    {f'<div style="margin-bottom:.4rem"><span style="color:#64748b;font-size:.68rem">Companies: </span>{company_chips}</div>' if company_chips else ''}
    {link_line}
    {'<div style="color:#64748b;font-size:.7rem;margin-top:.4rem"><i class="bi bi-vector-pen me-1"></i>Full text embedded — document is semantically searchable</div>' if ai_result.get('doc_id') else ''}
    <div class="d-flex gap-2 mt-2">
      <a href="/documents" class="btn btn-sm btn-outline-secondary">
        <i class="bi bi-file-earmark-text me-1"></i>Document Library</a>
      <a href="/search" class="btn btn-sm btn-outline-secondary">
        <i class="bi bi-search me-1"></i>Search</a>
    </div>
  </div>
</div>"""
    elif xls_result:
        err_li = "".join(f"<li>{e}</li>" for e in xls_result.get("errors", []))
        result_html = f"""
<div class="alert alert-{'success' if not xls_result['errors'] else 'warning'} mt-3">
  <strong>Excel Import:</strong> {xls_result['imported']} imported, {xls_result['skipped']} skipped.
  {'<ul class="mb-0 mt-2">'+err_li+'</ul>' if err_li else ''}
</div>"""

    ai_enabled = bool(OPENAI_API_KEY)
    ai_badge   = ('<span class="badge" style="background:#10b981;color:#000;font-size:.65rem">AI Ready</span>'
                  if ai_enabled else
                  '<span class="badge bg-secondary" style="font-size:.65rem">No API Key</span>')

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-cloud-arrow-up-fill me-2" style="color:#f59e0b"></i>AI Document Intake</h4>
    <p>Upload any document — AI extracts, classifies, summarises and embeds it for search.</p>
  </div>
  <div>{ai_badge}</div>
</div>

<div class="row g-3">
  <!-- Drop zone -->
  <div class="col-lg-6">
    <div class="card h-100">
      <div class="card-header">
        <i class="bi bi-cloud-upload me-1" style="color:#f59e0b"></i>Upload File
      </div>
      <div class="card-body">
        <form method="POST" action="/upload" enctype="multipart/form-data" id="uploadForm">
          <input type="hidden" name="mode" id="uploadMode" value="ai"/>
          <div id="dropZone" onclick="document.getElementById('fileInput').click()"
               style="border:2px dashed rgba(245,158,11,.35);border-radius:12px;padding:2.5rem 1.5rem;
                      text-align:center;cursor:pointer;background:rgba(245,158,11,.03);
                      transition:all .2s" onmouseover="this.style.borderColor='#f59e0b'"
               onmouseout="this.style.borderColor='rgba(245,158,11,.35)'">
            <i class="bi bi-file-earmark-arrow-up" style="font-size:2rem;color:#f59e0b;opacity:.7"></i>
            <div style="color:#94a3b8;margin-top:.6rem;font-size:.85rem">Click or drag a file here</div>
            <div style="color:#475569;font-size:.72rem;margin-top:.3rem">
              Excel, PDF, Word, TXT, CSV, Image, EML</div>
          </div>
          <input type="file" name="file" id="fileInput" class="d-none"
                 accept=".xlsx,.xls,.pdf,.docx,.doc,.txt,.csv,.eml,.jpg,.jpeg,.png,.gif,.webp,.bmp"/>
          <div id="fileName" style="font-size:.78rem;color:#64748b;margin-top:.5rem;min-height:1.2em"></div>

          <div class="d-grid gap-2 mt-3">
            <button type="submit" class="btn btn-primary" id="submitBtn" disabled>
              <i class="bi bi-cpu-fill me-1"></i>Run AI Intake
            </button>
          </div>
          <div class="mt-2 text-center">
            <button type="button" class="btn btn-link btn-sm" style="font-size:.73rem;color:#475569"
                    onclick="switchToExcel()">
              Upload .xlsx as properties spreadsheet instead →
            </button>
          </div>
        </form>
        {result_html}
      </div>
    </div>
  </div>

  <!-- Info panel -->
  <div class="col-lg-6">
    <div class="card mb-3">
      <div class="card-header"><i class="bi bi-magic me-1" style="color:#c084fc"></i>Document Types Recognised</div>
      <div class="card-body" style="font-size:.8rem;padding:.8rem">
        <div style="display:flex;flex-direction:column;gap:.5rem">
          <div style="display:flex;gap:.6rem;align-items:flex-start">
            <span style="min-width:28px;font-size:.7rem;font-weight:700;color:#f59e0b;padding-top:.1rem">REG</span>
            <div><strong style="color:#fff">Asset Register</strong>
              <div style="color:#64748b;font-size:.72rem">→ properties (occupier, landlord, grade, lease_expiry), contacts</div>
            </div>
          </div>
          <div style="display:flex;gap:.6rem;align-items:flex-start">
            <span style="min-width:28px;font-size:.7rem;font-weight:700;color:#10b981;padding-top:.1rem">VAC</span>
            <div><strong style="color:#fff">Vacancy Schedule</strong>
              <div style="color:#64748b;font-size:.72rem">→ vacancies (available_date, vacating_tenant, owner, agent), properties</div>
            </div>
          </div>
          <div style="display:flex;gap:.6rem;align-items:flex-start">
            <span style="min-width:28px;font-size:.7rem;font-weight:700;color:#fb923c;padding-top:.1rem">REQ</span>
            <div><strong style="color:#fff">Requirements / Listing</strong>
              <div style="color:#64748b;font-size:.72rem">→ requirements (company, size range, location, rating), contacts</div>
            </div>
          </div>
          <div style="display:flex;gap:.6rem;align-items:flex-start">
            <span style="min-width:28px;font-size:.7rem;font-weight:700;color:#c084fc;padding-top:.1rem">DEL</span>
            <div><strong style="color:#fff">Deal Tracker</strong>
              <div style="color:#64748b;font-size:.72rem">→ deals (tenant, address, size, rent, term, landlord), contacts</div>
            </div>
          </div>
          <div style="display:flex;gap:.6rem;align-items:flex-start">
            <span style="min-width:28px;font-size:.7rem;font-weight:700;color:#a78bfa;padding-top:.1rem">LSE</span>
            <div><strong style="color:#fff">Executed Lease / Contract</strong>
              <div style="color:#64748b;font-size:.72rem">→ market_data (all lease terms as evidence), deals (Completed)</div>
            </div>
          </div>
          <div style="display:flex;gap:.6rem;align-items:flex-start;border-top:1px solid rgba(255,255,255,.05);padding-top:.5rem">
            <span style="min-width:28px;font-size:.7rem;font-weight:700;color:#38bdf8;padding-top:.1rem">ALL</span>
            <div><strong style="color:#fff">Every document</strong>
              <div style="color:#64748b;font-size:.72rem">Full text embedded as 1536-dim vector. Companies auto-linked across all records.</div>
            </div>
          </div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-header"><i class="bi bi-file-earmark-spreadsheet me-1" style="color:#10b981"></i>Excel Column Map</div>
      <div class="card-body" style="max-height:220px;overflow-y:auto">
        <table class="table table-sm mb-0" style="font-size:.74rem">
          <thead><tr><th>Field</th><th>Accepted Headers</th></tr></thead>
          <tbody>
            {"".join(f'<tr><td class="fw-semibold" style="color:#f59e0b">{fi}</td><td style="color:#94a3b8">{", ".join(al)}</td></tr>' for fi,al in PROPERTY_COLUMN_MAP.items())}
          </tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<script>
const fileInput = document.getElementById('fileInput');
const submitBtn = document.getElementById('submitBtn');
const fileNameEl = document.getElementById('fileName');

fileInput.addEventListener('change', function() {{
  if (this.files.length > 0) {{
    const file = this.files[0];
    fileNameEl.textContent = '⏳ Processing ' + file.name + '...';
    submitBtn.disabled = true;
    const fd = new FormData();
    fd.append('file', file);
    fd.append('mode', document.getElementById('uploadMode').value || 'ai');
    fetch('/upload', {{method:'POST', body:fd}})
      .then(r => r.text())
      .then(() => {{
        fileNameEl.innerHTML = '✅ <strong>' + file.name + '</strong> done! <a href="/documents" style="color:#f59e0b">View documents</a>';
        fileInput.value = '';
      }})
      .catch(() => {{
        fileNameEl.textContent = '❌ Failed - try again';
        submitBtn.disabled = false;
      }});
  }}
}});

const dz = document.getElementById('dropZone');
dz.addEventListener('dragover', e => {{ e.preventDefault(); dz.style.borderColor='#f59e0b'; dz.style.background='rgba(245,158,11,.06)'; }});
dz.addEventListener('dragleave', () => {{ dz.style.borderColor='rgba(245,158,11,.35)'; dz.style.background='rgba(245,158,11,.03)'; }});
dz.addEventListener('drop', e => {{
  e.preventDefault();
  dz.style.borderColor='rgba(245,158,11,.35)'; dz.style.background='rgba(245,158,11,.03)';
  const dt = new DataTransfer();
  if (e.dataTransfer.files.length > 0) {{
    dt.items.add(e.dataTransfer.files[0]);
    fileInput.files = dt.files;
    fileInput.dispatchEvent(new Event('change'));
  }}
}});

function switchToExcel() {{
  document.getElementById('uploadMode').value = 'excel';
  submitBtn.innerHTML = '<i class="bi bi-file-earmark-spreadsheet-fill me-1"></i>Import as Properties';
  submitBtn.className = 'btn btn-success';
}}
</script>"""
    return render_layout(content, "AI Intake", "upload")

# ── Documents library ─────────────────────────────────────────────────────────
@app.route("/documents")
@login_required
def documents_page():
    docs = sb_select("documents", order="created_at", limit=200)

    type_meta = {
        "asset_register":       ("#f59e0b", "bi-list-columns-reverse",  "Asset Register"),
        "vacancy_schedule":     ("#10b981", "bi-door-open-fill",         "Vacancy Schedule"),
        "requirements_listing": ("#fb923c", "bi-search-heart",           "Requirements"),
        "deal_tracker":         ("#c084fc", "bi-kanban-fill",             "Deal Tracker"),
        "lease_contract":       ("#a78bfa", "bi-file-earmark-check-fill","Lease Contract"),
        "sales_contract":       ("#38bdf8", "bi-file-earmark-check-fill","Sales Contract"),
        "market_report":        ("#10b981", "bi-bar-chart-fill",         "Market Report"),
        "inquiry_email":        ("#f97316", "bi-envelope-fill",          "Inquiry Email"),
        "inspection_report":    ("#64748b", "bi-clipboard2-check-fill",  "Inspection"),
        "invoice":              ("#64748b", "bi-receipt",                "Invoice"),
        "general_correspondence":("#475569","bi-chat-left-text-fill",    "Correspondence"),
        "unknown":              ("#334155", "bi-file-earmark-fill",      "Unknown"),
    }

    extract_types = [
        ("extracted_properties",  "#f59e0b", "P"),
        ("extracted_contacts",    "#38bdf8", "C"),
        ("extracted_deals",       "#c084fc", "D"),
        ("extracted_vacancies",   "#10b981", "V"),
        ("extracted_requirements","#fb923c", "R"),
        ("extracted_market_data", "#a78bfa", "M"),
    ]

    cards = ""
    for d in docs:
        cls  = (d.get("ai_classification") or "unknown").lower()
        col, icon, label = type_meta.get(cls, ("#334155","bi-file-earmark-fill",cls.replace("_"," ").title()))
        conf = int((d.get("ai_confidence") or 0) * 100)
        conf_col = "#10b981" if conf >= 70 else ("#f59e0b" if conf >= 40 else "#475569")

        # Extract-count badges (non-zero only)
        badge_html = ""
        for field, bcol, letter in extract_types:
            n = d.get(field) or 0
            if n:
                badge_html += (f'<span title="{field.replace("extracted_","").replace("_"," ").title()}" '
                               f'style="width:18px;height:18px;border-radius:4px;background:{bcol}22;'
                               f'color:{bcol};font-size:.6rem;font-weight:700;display:flex;align-items:center;'
                               f'justify-content:center">{letter}{n}</span>')

        # Companies
        companies = d.get("mentioned_companies") or []
        if isinstance(companies, str):
            try: companies = json.loads(companies)
            except Exception: companies = []
        comp_html = " ".join(
            f'<span style="background:rgba(56,189,248,.08);color:#38bdf8;border-radius:4px;'
            f'padding:.05rem .3rem;font-size:.6rem">{c[:22]}</span>'
            for c in companies[:3]
        )

        # Linked docs count
        linked = d.get("linked_document_ids") or []
        if isinstance(linked, str):
            try: linked = json.loads(linked)
            except Exception: linked = []
        link_badge = (f'<span title="{len(linked)} linked doc(s)" style="font-size:.62rem;color:#38bdf8">'
                      f'<i class="bi bi-link-45deg"></i>{len(linked)}</span>') if linked else ""

        cards += f"""
<div class="col-md-6 col-lg-4">
  <div class="card h-100" style="border-color:rgba(255,255,255,.05)">
    <div class="card-body" style="padding:.85rem">
      <div style="display:flex;align-items:flex-start;gap:.65rem;margin-bottom:.55rem">
        <div style="min-width:34px;height:34px;border-radius:7px;background:{col}18;
             display:flex;align-items:center;justify-content:center;color:{col};font-size:.85rem;flex-shrink:0">
          <i class="bi {icon}"></i>
        </div>
        <div style="flex:1;min-width:0">
          <div style="font-weight:700;font-size:.78rem;color:#fff;white-space:nowrap;
               overflow:hidden;text-overflow:ellipsis" title="{d.get('filename','')}">
            {(d.get('filename') or 'Untitled')[:38]}
          </div>
          <div style="font-size:.66rem;color:#334155">{fmt_date(d.get('created_at'))}</div>
        </div>
        <div style="display:flex;align-items:center;gap:.3rem">
          {link_badge}
          <span style="background:{col}18;color:{col};border-radius:4px;padding:.08rem .35rem;
                font-size:.6rem;font-weight:700;white-space:nowrap">{label}</span>
        </div>
      </div>
      <div style="font-size:.74rem;color:#94a3b8;margin-bottom:.5rem;
           display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">
        {d.get('ai_summary') or '<span style="color:#1e3a5f;font-style:italic">No summary</span>'}
      </div>
      <div style="display:flex;gap:.25rem;flex-wrap:wrap;align-items:center;
           justify-content:space-between;margin-top:.35rem">
        <div style="display:flex;gap:.25rem;flex-wrap:wrap;align-items:center">
          {badge_html}
          {comp_html}
        </div>
        <span style="font-size:.62rem;color:{conf_col}">{conf}%</span>
      </div>
    </div>
  </div>
</div>"""

    if not cards:
        cards = """
<div class="col-12">
  <div style="text-align:center;padding:4rem 1rem;color:#1e3a5f">
    <i class="bi bi-file-earmark-plus" style="font-size:2.5rem;opacity:.4;display:block;margin-bottom:.8rem"></i>
    No documents yet — <a href="/upload" style="color:#f59e0b">upload your first file</a>
  </div>
</div>"""

    total_docs = len(docs)
    embedded   = sum(1 for d in docs if d.get("embedding"))
    linked_any = sum(1 for d in docs if (d.get("linked_document_ids") or []))
    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-file-earmark-text-fill me-2" style="color:#f59e0b"></i>Document Library</h4>
    <p>{total_docs} document{'s' if total_docs!=1 else ''} &bull;
       {embedded} embedded &bull;
       {linked_any} with cross-links</p>
  </div>
  <a href="/upload" class="btn btn-sm btn-primary">
    <i class="bi bi-cloud-arrow-up-fill me-1"></i>Add Document</a>
</div>

<div class="row g-3">
  {cards}
</div>"""
    return render_layout(content, "Documents", "documents")


# ── Semantic search ───────────────────────────────────────────────────────────
@app.route("/search")
@login_required
def search_page():
    q       = request.args.get("q", "").strip()
    results = []
    error   = ""

    if q:
        if not OPENAI_API_KEY:
            error = "OpenAI API key not configured — semantic search unavailable."
        else:
            try:
                results = semantic_search(q, top_k=12)
            except Exception as e:
                log.error("semantic_search error: %s", e)
                error = f"Search failed: {e}"

    result_rows = ""
    for r in results:
        sim = r.get("similarity") or 0
        pct = int(sim * 100)
        bar_col = "#10b981" if pct >= 70 else ("#f59e0b" if pct >= 40 else "#64748b")
        cls = r.get("ai_classification") or "Other"
        tags = r.get("ai_tags") or []
        if isinstance(tags, str):
            try: tags = json.loads(tags)
            except Exception: tags = []
        tag_html = " ".join(f'<span class="badge bg-secondary" style="font-size:.6rem">{t}</span>' for t in tags[:3])
        result_rows += f"""
<div class="card mb-2">
  <div class="card-body" style="padding:.75rem 1rem">
    <div style="display:flex;align-items:flex-start;gap:.75rem">
      <div style="min-width:44px;text-align:center">
        <div style="font-size:1rem;font-weight:800;color:{bar_col}">{pct}%</div>
        <div style="font-size:.6rem;color:#475569">match</div>
      </div>
      <div style="flex:1;min-width:0">
        <div style="display:flex;align-items:center;gap:.5rem;margin-bottom:.25rem">
          <span style="font-weight:700;color:#fff;font-size:.82rem">{r.get('filename') or 'Untitled'}</span>
          <span class="badge" style="background:rgba(245,158,11,.15);color:#f59e0b;font-size:.62rem">{cls}</span>
        </div>
        <div style="font-size:.76rem;color:#94a3b8;margin-bottom:.35rem">{r.get('ai_summary') or ''}</div>
        {f'<div style="display:flex;gap:.25rem;flex-wrap:wrap">{tag_html}</div>' if tag_html else ''}
      </div>
      <div style="font-size:.68rem;color:#475569;white-space:nowrap">{fmt_date(r.get('created_at'))}</div>
    </div>
  </div>
</div>"""

    empty_html = ""
    if q and not results and not error:
        empty_html = """
<div style="text-align:center;padding:3rem 1rem;color:#1e3a5f">
  <i class="bi bi-search" style="font-size:2rem;opacity:.3;display:block;margin-bottom:.7rem"></i>
  No documents matched. Try different keywords or upload more files.
</div>"""

    ai_status = ('<span style="color:#10b981;font-size:.73rem"><i class="bi bi-check-circle-fill me-1"></i>AI Ready</span>'
                 if OPENAI_API_KEY else
                 '<span style="color:#ef4444;font-size:.73rem"><i class="bi bi-exclamation-circle-fill me-1"></i>No API Key</span>')

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-search-heart-fill me-2" style="color:#f59e0b"></i>Semantic Search</h4>
    <p>Search documents by meaning — not just keywords. {ai_status}</p>
  </div>
</div>

<div class="card mb-4">
  <div class="card-body" style="padding:1.1rem">
    <form method="GET" action="/search">
      <div class="d-flex gap-2">
        <input type="text" name="q" class="form-control" placeholder='e.g. "warehouse lease Sunshine" or "buyer looking for 2000sqm"…'
               value="{q}" autofocus style="font-size:.88rem"/>
        <button type="submit" class="btn btn-primary px-4">
          <i class="bi bi-search me-1"></i>Search
        </button>
      </div>
    </form>
    <div class="mt-2 d-flex gap-2 flex-wrap" style="font-size:.72rem;color:#475569">
      <span>Try:</span>
      <a href="/search?q=lease+enquiry+West+Melbourne" style="color:#f59e0b">lease enquiry West Melbourne</a>
      <a href="/search?q=industrial+warehouse+for+sale" style="color:#f59e0b">industrial warehouse for sale</a>
      <a href="/search?q=tenant+renewing+contract" style="color:#f59e0b">tenant renewing contract</a>
    </div>
  </div>
</div>

{'<div class="alert" style="background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.2);color:#ef4444;border-radius:8px;padding:.6rem 1rem;font-size:.8rem">'+error+'</div>' if error else ''}

{f'<div style="font-size:.78rem;color:#64748b;margin-bottom:.75rem">{len(results)} result{"s" if len(results)!=1 else ""} for <em style=\'color:#f59e0b\'>{q}</em></div>' if q and results else ''}

{result_rows}
{empty_html}

{'' if q else '''
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:.8rem;margin-top:1rem">
  <div class="card" style="border-color:rgba(56,189,248,.15)">
    <div class="card-body" style="padding:.85rem">
      <div style="color:#38bdf8;font-size:.8rem;font-weight:700;margin-bottom:.3rem">
        <i class="bi bi-vector-pen me-1"></i>How it works</div>
      <div style="color:#64748b;font-size:.73rem">Your query is converted to a 1536-dim vector. Results are ranked by cosine similarity to that vector using pgvector in Supabase.</div>
    </div>
  </div>
  <div class="card" style="border-color:rgba(192,132,252,.15)">
    <div class="card-body" style="padding:.85rem">
      <div style="color:#c084fc;font-size:.8rem;font-weight:700;margin-bottom:.3rem">
        <i class="bi bi-lightbulb-fill me-1"></i>Best results</div>
      <div style="color:#64748b;font-size:.73rem">Use natural language — describe what you are looking for. More documents uploaded = better coverage.</div>
    </div>
  </div>
  <div class="card" style="border-color:rgba(16,185,129,.15)">
    <div class="card-body" style="padding:.85rem">
      <div style="color:#10b981;font-size:.8rem;font-weight:700;margin-bottom:.3rem">
        <i class="bi bi-file-earmark-plus me-1"></i>Add more docs</div>
      <div style="color:#64748b;font-size:.73rem">Upload leases, inspection reports, market data, emails — anything relevant to West Melbourne industrial.</div>
    </div>
  </div>
</div>'''}
"""
    return render_layout(content, "Search", "search")


@app.route("/briefings")
@login_required
def briefings_page():
    briefs = sb_select("briefings", order="created_at", limit=20)
    rows = "".join(f"""<tr>
        <td><span class="badge bg-secondary">{b.get('briefing_type','—')}</span></td>
        <td>{b.get('channel','—')}</td>
        <td class="text-truncate" style="max-width:350px">{(b.get('content') or '')[:100]}</td>
        <td>{fmt_date(b.get('created_at'))}</td>
        <td>
          <button class="btn btn-sm btn-outline-secondary" data-bs-toggle="modal"
            data-bs-target="#viewBriefModal" onclick="showBrief(`{(b.get('content') or '').replace('`','').replace(chr(10),'\\n')}`)">View</button>
        </td>
    </tr>""" for b in briefs)
    content = f"""
<div class="ph">
  <div><h4>Briefings</h4><p>Automated market reports sent via WhatsApp</p></div>
  <div class="d-flex gap-2">
    <a href="/briefings/send" class="btn btn-sm btn-primary">
      <i class="bi bi-send me-1"></i>Send Daily Brief Now</a>
    <a href="/briefings/preview" class="btn btn-sm btn-outline-secondary">
      <i class="bi bi-eye me-1"></i>Preview</a>
  </div>
</div>
<div class="alert alert-info mb-3" style="font-size:.84rem">
  <i class="bi bi-clock me-1"></i>
  Daily briefing scheduled at <strong>{BRIEFING_HOUR:02d}:{BRIEFING_MINUTE:02d} {TIMEZONE}</strong> —
  Weekly briefing every Monday.
  Gmail checked every <strong>15 minutes</strong>.
</div>
<div class="card">
  <div class="card-header">Briefing History</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Type</th><th>Channel</th><th>Preview</th><th>Sent At</th><th></th></tr></thead>
      <tbody>{rows or '<tr><td colspan="5" class="text-center text-muted py-4">No briefings sent yet.</td></tr>'}</tbody>
    </table></div>
  </div>
</div>
<div class="modal fade" id="viewBriefModal" tabindex="-1">
  <div class="modal-dialog"><div class="modal-content">
    <div class="modal-header"><h5 class="modal-title">Briefing Content</h5>
      <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
    <div class="modal-body">
      <pre id="briefContent" style="white-space:pre-wrap;font-size:.85rem;background:#f8f9fa;padding:1rem;border-radius:8px;max-height:400px;overflow-y:auto"></pre>
    </div>
  </div></div>
</div>
<script>
function showBrief(txt) {{
  document.getElementById('briefContent').textContent = txt.replace(/\\\\n/g,'\\n');
}}
</script>"""
    return render_layout(content, "Briefings", "briefings")

@app.route("/briefings/send")
@login_required
def briefing_send_now():
    t = threading.Thread(target=send_daily_briefing)
    t.daemon = True
    t.start()
    flash("Daily briefing sent in background.", "success")
    return redirect(url_for("briefings_page"))

@app.route("/briefings/preview")
@login_required
def briefing_preview():
    content_text = build_market_brief()
    content = f"""
<div class="ph"><div><h4>Briefing Preview</h4><p>This is what will be sent</p></div>
<div class="card"><div class="card-body">
  <pre style="white-space:pre-wrap;font-size:.88rem;background:#f8f9fa;padding:1.2rem;border-radius:8px">{content_text}</pre>
  <a href="/briefings/send" class="btn btn-primary mt-3"><i class="bi bi-send me-1"></i>Send Now</a>
  <a href="/briefings" class="btn btn-outline-secondary mt-3 ms-2">Back</a>
</div></div>"""
    return render_layout(content, "Brief Preview", "briefings")

@app.route("/analytics")
@login_required
def analytics_page():
    from collections import defaultdict

    props   = sb_select("properties")
    deals   = sb_select("deals",    order="created_at", limit=500)
    fees    = sb_select("fees",     order="created_at", limit=500)
    all_inq = sb_select("inquiries", order="created_at", limit=500)

    # ── Property breakdown ────────────────────────────────────────────────────
    by_type, by_status, size_data = {}, {}, []
    for p in props:
        t = p.get("property_type") or "Other"
        s = p.get("status")        or "Unknown"
        by_type[t]   = by_type.get(t, 0)   + 1
        by_status[s] = by_status.get(s, 0) + 1
        try: size_data.append(float(p["size_sqm"]))
        except (TypeError, ValueError, KeyError): pass
    avg_size = round(sum(size_data)/len(size_data), 0) if size_data else 0

    # ── Fee pipeline by month ──────────────────────────────────────────────────
    fee_by_month   = defaultdict(float)
    fee_by_landlord = defaultdict(float)
    total_fees = paid_fees = 0.0
    for f in fees:
        amt = float(f.get("fee_amount") or 0)
        total_fees += amt
        if f.get("invoice_status") == "Paid":
            paid_fees += amt
        raw_dt = f.get("created_at") or ""
        try:
            month_key = str(raw_dt)[:7]   # "YYYY-MM"
            fee_by_month[month_key]       += amt
        except Exception:
            pass
        ll = f.get("landlord") or f.get("client_name") or "Other"
        fee_by_landlord[ll] += amt

    # sort months chronologically, last 12
    sorted_months = sorted(fee_by_month.keys())[-12:]
    fee_months_labels = sorted_months
    fee_months_values = [round(fee_by_month[m], 0) for m in sorted_months]

    # top 7 landlords by revenue
    top_landlords = sorted(fee_by_landlord.items(), key=lambda x: -x[1])[:7]
    ll_labels = [x[0][:22] for x in top_landlords]
    ll_values = [round(x[1], 0) for x in top_landlords]

    # ── Deals by suburb ────────────────────────────────────────────────────────
    deals_by_suburb = defaultdict(int)
    deals_by_type   = defaultdict(int)
    deals_by_month  = defaultdict(int)
    for d in deals:
        sub = d.get("suburb") or "Unknown"
        deals_by_suburb[sub] += 1
        deals_by_type[d.get("deal_type") or "Unknown"] += 1
        raw_dt = d.get("created_at") or ""
        try:  deals_by_month[str(raw_dt)[:7]] += 1
        except Exception: pass
    top_suburbs = sorted(deals_by_suburb.items(), key=lambda x: -x[1])[:8]
    sub_labels  = [x[0][:20] for x in top_suburbs]
    sub_values  = [x[1]      for x in top_suburbs]

    # ── Inquiry source breakdown ───────────────────────────────────────────────
    inq_by_source = defaultdict(int)
    for i in all_inq:
        inq_by_source[i.get("source") or "Web"] += 1

    # deal volume by month (last 12)
    sorted_deal_months = sorted(deals_by_month.keys())[-12:]
    deal_month_vals    = [deals_by_month[m] for m in sorted_deal_months]

    kpi_total   = len(props)
    kpi_avail   = by_status.get("Available", 0)
    kpi_vac_pct = round(kpi_avail / kpi_total * 100 if kpi_total else 0, 1)
    kpi_deals   = len(deals)

    content = f"""
<div class="ph"><div><h4>Analytics</h4><p>Live market intelligence — {kpi_total} properties, {kpi_deals} deals, ${total_fees:,.0f} total fees</p></div></div>

<!-- KPI row -->
<div class="row g-3 mb-4">
  <div class="col-6 col-md-3"><div class="kpi">
    <div class="kpi-val amber">{kpi_total}</div><div class="kpi-label">Total Properties</div>
  </div></div>
  <div class="col-6 col-md-3"><div class="kpi">
    <div class="kpi-val green">{kpi_avail}</div><div class="kpi-label">Available ({kpi_vac_pct}%)</div>
  </div></div>
  <div class="col-6 col-md-3"><div class="kpi">
    <div class="kpi-val blue">{kpi_deals}</div><div class="kpi-label">Total Deals</div>
  </div></div>
  <div class="col-6 col-md-3"><div class="kpi">
    <div class="kpi-val" style="color:#10b981">${total_fees:,.0f}</div><div class="kpi-label">Fee Revenue</div>
  </div></div>
</div>

<!-- Row 1: Fee pipeline + Deals volume -->
<div class="row g-3 mb-3">
  <div class="col-lg-7">
    <div class="card">
      <div class="card-header">Fee Pipeline by Month</div>
      <div class="card-body"><canvas id="feeMonthChart" style="max-height:240px"></canvas></div>
    </div>
  </div>
  <div class="col-lg-5">
    <div class="card">
      <div class="card-header">Deal Volume by Month</div>
      <div class="card-body"><canvas id="dealMonthChart" style="max-height:240px"></canvas></div>
    </div>
  </div>
</div>

<!-- Row 2: Deals by suburb + Revenue by landlord -->
<div class="row g-3 mb-3">
  <div class="col-lg-6">
    <div class="card">
      <div class="card-header">Deals by Suburb</div>
      <div class="card-body"><canvas id="suburbChart" style="max-height:260px"></canvas></div>
    </div>
  </div>
  <div class="col-lg-6">
    <div class="card">
      <div class="card-header">Revenue by Landlord / Client</div>
      <div class="card-body"><canvas id="landlordChart" style="max-height:260px"></canvas></div>
    </div>
  </div>
</div>

<!-- Row 3: Property type + Inquiry source + Status -->
<div class="row g-3">
  <div class="col-md-4">
    <div class="card">
      <div class="card-header">By Property Type</div>
      <div class="card-body"><canvas id="typeChart" style="max-height:220px"></canvas></div>
    </div>
  </div>
  <div class="col-md-4">
    <div class="card">
      <div class="card-header">Inquiry Sources</div>
      <div class="card-body"><canvas id="inqChart" style="max-height:220px"></canvas></div>
    </div>
  </div>
  <div class="col-md-4">
    <div class="card">
      <div class="card-header">By Status</div>
      <div class="card-body"><canvas id="statusChart" style="max-height:220px"></canvas></div>
    </div>
  </div>
</div>

<script>
const PAL = ['#f59e0b','#10b981','#3b82f6','#ef4444','#8b5cf6','#06b6d4','#f97316','#a78bfa'];
const GRID = {{ color:'rgba(255,255,255,.05)' }};
const TICK = {{ color:'#475569', font:{{size:11}} }};
const DARK_OPTS = {{
  plugins: {{ legend: {{ labels: {{ color:'#94a3b8', font:{{size:11}} }} }} }},
  scales:  {{ x: {{ grid:GRID, ticks:TICK }}, y: {{ grid:GRID, ticks:TICK, beginAtZero:true }} }}
}};

// Fee pipeline by month
new Chart(document.getElementById('feeMonthChart'), {{
  type:'bar',
  data:{{
    labels: {json.dumps(fee_months_labels)},
    datasets:[{{
      label:'Fees ($)', data:{json.dumps(fee_months_values)},
      backgroundColor:'rgba(245,158,11,.7)', borderColor:'#f59e0b',
      borderWidth:1, borderRadius:4
    }}]
  }},
  options:{{...DARK_OPTS, plugins:{{legend:{{display:false}}}}}}
}});

// Deal volume by month
new Chart(document.getElementById('dealMonthChart'), {{
  type:'line',
  data:{{
    labels: {json.dumps(sorted_deal_months)},
    datasets:[{{
      label:'Deals', data:{json.dumps(deal_month_vals)},
      borderColor:'#10b981', backgroundColor:'rgba(16,185,129,.1)',
      pointBackgroundColor:'#10b981', tension:.35, fill:true, borderWidth:2
    }}]
  }},
  options:{{...DARK_OPTS, plugins:{{legend:{{display:false}}}}}}
}});

// Deals by suburb
new Chart(document.getElementById('suburbChart'), {{
  type:'bar',
  data:{{
    labels:{json.dumps(sub_labels)},
    datasets:[{{data:{json.dumps(sub_values)},backgroundColor:PAL,borderRadius:4,borderWidth:0}}]
  }},
  options:{{...DARK_OPTS, indexAxis:'y', plugins:{{legend:{{display:false}}}}}}
}});

// Revenue by landlord
new Chart(document.getElementById('landlordChart'), {{
  type:'bar',
  data:{{
    labels:{json.dumps(ll_labels)},
    datasets:[{{data:{json.dumps(ll_values)},backgroundColor:'rgba(167,139,250,.7)',
                borderColor:'#a78bfa',borderRadius:4,borderWidth:1}}]
  }},
  options:{{...DARK_OPTS, indexAxis:'y', plugins:{{legend:{{display:false}}}}}}
}});

// Property type
new Chart(document.getElementById('typeChart'), {{
  type:'doughnut',
  data:{{
    labels:{json.dumps(list(by_type.keys()))},
    datasets:[{{data:{json.dumps(list(by_type.values()))},backgroundColor:PAL,borderWidth:0,hoverOffset:6}}]
  }},
  options:{{cutout:'58%',plugins:{{legend:{{position:'bottom',labels:{{color:'#94a3b8',font:{{size:10}}}}}}}}}}
}});

// Inquiry source
new Chart(document.getElementById('inqChart'), {{
  type:'doughnut',
  data:{{
    labels:{json.dumps(list(inq_by_source.keys()))},
    datasets:[{{data:{json.dumps(list(inq_by_source.values()))},backgroundColor:PAL,borderWidth:0,hoverOffset:6}}]
  }},
  options:{{cutout:'58%',plugins:{{legend:{{position:'bottom',labels:{{color:'#94a3b8',font:{{size:10}}}}}}}}}}
}});

// Status
new Chart(document.getElementById('statusChart'), {{
  type:'doughnut',
  data:{{
    labels:{json.dumps(list(by_status.keys()))},
    datasets:[{{data:{json.dumps(list(by_status.values()))},backgroundColor:PAL,borderWidth:0,hoverOffset:6}}]
  }},
  options:{{cutout:'58%',plugins:{{legend:{{position:'bottom',labels:{{color:'#94a3b8',font:{{size:10}}}}}}}}}}
}});
</script>"""
    return render_layout(content, "Analytics", "analytics")

@app.route("/settings")
@login_required
def settings_page():
    jobs = scheduler.get_jobs()
    job_rows = "".join(f"""<tr>
        <td class="fw-semibold">{j.id}</td>
        <td><code>{j.trigger}</code></td>
        <td>{str(j.next_run_time)[:19] if j.next_run_time else '—'}</td>
        <td><span class="badge bg-success">Active</span></td>
    </tr>""" for j in jobs)
    content = f"""
<div class="ph"><div><h4>Settings</h4><p>Environment configuration and scheduler status</p></div>
<div class="row g-3">
  <div class="col-md-6">
    <div class="card"><div class="card-header">Integration Status</div>
      <div class="card-body">
        <table class="table table-sm mb-0">
          <tbody>
            <tr><td>Supabase</td><td>{'<span class="badge bg-success">Connected</span>' if SUPABASE_URL else '<span class="badge bg-danger">Not Configured</span>'}</td></tr>
            <tr><td>Twilio / WhatsApp</td><td>{'<span class="badge bg-success">Connected</span>' if TWILIO_SID else '<span class="badge bg-danger">Not Configured</span>'}</td></tr>
            <tr><td>Gmail IMAP</td><td>{'<span class="badge bg-success">Connected</span>' if GMAIL_USER else '<span class="badge bg-danger">Not Configured</span>'}</td></tr>
            <tr><td>WA Broadcast List</td><td><code style="font-size:.78rem">{len(WA_BROADCAST_LIST)} recipient(s)</code></td></tr>
            <tr><td>Daily Briefing</td><td><code>{BRIEFING_HOUR:02d}:{BRIEFING_MINUTE:02d} {TIMEZONE}</code></td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
  <div class="col-md-6">
    <div class="card"><div class="card-header">Scheduled Jobs</div>
      <div class="card-body p-0">
        <table class="table table-sm mb-0">
          <thead><tr><th>Job ID</th><th>Trigger</th><th>Next Run</th><th>State</th></tr></thead>
          <tbody>{job_rows or '<tr><td colspan="4" class="text-center text-muted py-3">Scheduler not started.</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  </div>
  <div class="col-12">
    <div class="card"><div class="card-header">Required Environment Variables</div>
      <div class="card-body">
        <pre style="font-size:.8rem;background:#f8f9fa;padding:1rem;border-radius:8px"># .env file
SECRET_KEY=your-secret-key
ADMIN_USER=admin
ADMIN_PASSWORD=your-secure-password

SUPABASE_URL=https://xxxx.supabase.co
SUPABASE_KEY=your-anon-or-service-key

TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxxxx
TWILIO_AUTH_TOKEN=your-auth-token
TWILIO_WHATSAPP_FROM=whatsapp:+14155238886
WA_BROADCAST_LIST=+61412345678,+61498765432

GMAIL_USER=you@gmail.com
GMAIL_APP_PASSWORD=xxxx-xxxx-xxxx-xxxx

BRIEFING_HOUR=8
BRIEFING_MINUTE=0
TIMEZONE=Australia/Melbourne</pre>
      </div>
    </div>
  </div>
</div>"""
    return render_layout(content, "Settings", "settings")

# ── API routes ─────────────────────────────────────────────────────────────────
@app.route("/api/properties", methods=["POST"])
@login_required
def api_add_property():
    data = {k: v for k, v in request.form.items() if v != ""}
    for f in ("size_sqm","land_sqm","asking_price","asking_rent_pa","year_built"):
        if f in data:
            try: data[f] = float(data[f])
            except ValueError: data.pop(f)
    data["source"] = "Manual"
    sb_insert("properties", data)
    flash("Property added.", "success")
    return redirect(url_for("properties_page"))

@app.route("/api/properties/<int:pid>", methods=["DELETE"])
@login_required
def api_delete_property(pid):
    sb_delete("properties", pid)
    return jsonify({"ok": True})

@app.route("/api/contacts", methods=["POST"])
@login_required
def api_add_contact():
    data = {k: v for k, v in request.form.items() if v != ""}
    data["whatsapp_opt_in"] = "whatsapp_opt_in" in request.form
    sb_insert("contacts", data)
    flash("Contact added.", "success")
    return redirect(url_for("contacts_page"))

@app.route("/api/contacts/<int:cid>", methods=["DELETE"])
@login_required
def api_delete_contact(cid):
    sb_delete("contacts", cid)
    return jsonify({"ok": True})

@app.route("/api/inquiries", methods=["POST"])
@login_required
def api_add_inquiry():
    data = {k: v for k, v in request.form.items() if v != ""}
    sb_insert("inquiries", data)
    flash("Inquiry added.", "success")
    return redirect(url_for("inquiries_page"))

@app.route("/api/inquiries/<int:iid>", methods=["DELETE"])
@login_required
def api_delete_inquiry(iid):
    sb_delete("inquiries", iid)
    return jsonify({"ok": True})

@app.route("/api/inquiries/<int:iid>/status", methods=["PATCH"])
@login_required
def api_update_inquiry_status(iid):
    data = request.get_json(silent=True) or {}
    sb_update("inquiries", {"id": iid}, {"status": data.get("status", "Resolved")})
    return jsonify({"ok": True})

@app.route("/api/deals", methods=["POST"])
@login_required
def api_add_deal():
    data = {k: v for k, v in request.form.items() if v != ""}
    if "price" in data:
        try: data["price"] = float(data["price"])
        except ValueError: data.pop("price")
    sb_insert("deals", data)
    flash("Deal added.", "success")
    return redirect(url_for("deals_page"))

@app.route("/api/deals/<int:did>", methods=["DELETE"])
@login_required
def api_delete_deal(did):
    sb_delete("deals", did)
    return jsonify({"ok": True})

@app.route("/api/whatsapp/send", methods=["POST"])
@login_required
def api_send_whatsapp():
    body = request.form.get("body", "").strip()
    broadcast = request.form.get("broadcast") == "on"
    if not body:
        flash("Message body is required.", "error")
        return redirect(url_for("whatsapp_page"))
    if broadcast:
        sent, failed = broadcast_whatsapp(body)
        flash(f"Broadcast sent to {sent}, failed {failed}.", "success")
    else:
        to = request.form.get("to", "").strip()
        if not to:
            flash("Recipient number is required.", "error")
            return redirect(url_for("whatsapp_page"))
        ok = send_whatsapp(to, body)
        flash("Message sent." if ok else "Failed to send — check Twilio config.", "success" if ok else "error")
    return redirect(url_for("whatsapp_page"))

# ── Webhooks ──────────────────────────────────────────────────────────────────
@app.route("/webhook/whatsapp", methods=["POST"])
def webhook_whatsapp():
    from_num = request.form.get("From", "")
    body     = request.form.get("Body", "").strip()
    log.info("WhatsApp inbound: %s — %s", from_num, body[:80])
    reply = handle_inbound_whatsapp(from_num, body)
    resp = MessagingResponse()
    resp.message(reply)
    return str(resp), 200, {"Content-Type": "text/xml"}

@app.route("/webhook/whatsapp/status", methods=["POST"])
def webhook_whatsapp_status():
    sid    = request.form.get("MessageSid", "")
    status = request.form.get("MessageStatus", "")
    log.info("WA status callback: %s → %s", sid, status)
    return "", 204

# ── Outlook add-in ────────────────────────────────────────────────────────────
@app.route("/outlook-manifest.xml")
def outlook_manifest():
    base_url = get_base_url()
    xml = OUTLOOK_MANIFEST_TEMPLATE.replace("{base_url}", base_url)
    return xml, 200, {"Content-Type": "application/xml"}

@app.route("/outlook-taskpane")
def outlook_taskpane():
    return OUTLOOK_TASKPANE_HTML, 200, {"Content-Type": "text/html"}

@app.route("/outlook-capture", methods=["POST"])
def outlook_capture():
    data = request.get_json(silent=True) or {}
    if not data:
        return jsonify({"ok": False, "message": "No data received"}), 400
    subject  = data.get("subject", "(no subject)")
    from_email = data.get("from_email", "")
    from_name  = data.get("from_name", "")
    body       = data.get("body", "")
    date_str   = data.get("date", "")
    # Store to email_logs
    sb_insert("email_logs", {
        "sender":  f"{from_name} <{from_email}>",
        "subject": subject,
        "body":    body[:5000],
        "processed": False,
    })
    # Check for inquiry keywords
    inq_keywords = ["available","lease","rent","warehouse","factory","looking",
                    "require","need","space","sqm","m2","inquiry","enquiry"]
    is_inquiry = any(kw in (subject + body).lower() for kw in inq_keywords)
    inq_id = None
    if is_inquiry and from_email:
        inserted = sb_insert("inquiries", {
            "contact_name":  from_name or from_email,
            "contact_email": from_email,
            "source":        "Outlook Capture",
            "message":       f"{subject}\n\n{body[:1000]}",
            "status":        "New",
        })
        if inserted:
            inq_id = inserted[0].get("id")
    msg = f"Email captured {'+ inquiry created' if inq_id else ''}. Subject: {subject[:60]}"
    log.info("Outlook capture: %s", msg)
    return jsonify({"ok": True, "message": msg, "inquiry_id": inq_id})

@app.route("/outlook-setup")
@login_required
def outlook_setup():
    base_url = get_base_url()
    manifest_url = f"{base_url}/outlook-manifest.xml"
    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-envelope-check-fill me-2" style="color:#f59e0b"></i>Outlook Add-in Setup</h4>
    <p>Capture property emails from Outlook with one click. No automatic scanning of your work inbox.</p>
  </div>
</div>

<div class="row g-3">
  <div class="col-lg-6">
    <div class="card">
      <div class="card-header"><i class="bi bi-plug-fill me-1" style="color:#38bdf8"></i>Option 1 — Install Add-in</div>
      <div class="card-body" style="font-size:.82rem">
        <p style="color:#94a3b8">Adds a "MJR West Agent" button to every email you read in Outlook.</p>
        <div style="background:rgba(245,158,11,.06);border:1px solid rgba(245,158,11,.2);
             border-radius:8px;padding:.75rem;margin-bottom:.8rem">
          <div style="font-size:.7rem;color:#64748b;margin-bottom:.3rem">Manifest URL</div>
          <code style="color:#f59e0b;font-size:.75rem;word-break:break-all">{manifest_url}</code>
        </div>
        <ol style="color:#94a3b8;font-size:.78rem;padding-left:1.2rem;line-height:1.9">
          <li>Open Outlook → <strong style="color:#fff">File → Manage Add-ins</strong> (or Admin Center)</li>
          <li>Click <strong style="color:#fff">+ Add a custom add-in → From URL</strong></li>
          <li>Paste the manifest URL above</li>
          <li>Open any email → the <strong style="color:#fff">MJR West Agent</strong> button appears in the ribbon</li>
          <li>Click it to send the email to this system</li>
        </ol>
        <div style="background:rgba(56,189,248,.05);border:1px solid rgba(56,189,248,.1);
             border-radius:6px;padding:.6rem;font-size:.72rem;color:#64748b;margin-top:.5rem">
          <i class="bi bi-info-circle me-1" style="color:#38bdf8"></i>
          Outlook requires HTTPS for add-ins. Set the <code>APP_URL</code> environment variable
          to your public domain (e.g. <code>https://agent.mjrwest.com.au</code>) once deployed.
        </div>
      </div>
    </div>
  </div>
  <div class="col-lg-6">
    <div class="card mb-3">
      <div class="card-header"><i class="bi bi-forward-fill me-1" style="color:#10b981"></i>Option 2 — BCC Forwarding</div>
      <div class="card-body" style="font-size:.82rem">
        <p style="color:#94a3b8">Simpler: just BCC any property email to the agent address. The Gmail IMAP monitor processes it automatically.</p>
        <div style="background:rgba(16,185,129,.06);border:1px solid rgba(16,185,129,.2);
             border-radius:8px;padding:.75rem;margin-bottom:.8rem;text-align:center">
          <div style="font-size:.7rem;color:#64748b;margin-bottom:.2rem">BCC address</div>
          <code style="color:#10b981;font-size:.9rem">{GMAIL_USER}</code>
        </div>
        <ul style="color:#94a3b8;font-size:.78rem;padding-left:1.2rem;line-height:1.9">
          <li>Add <code style="color:#10b981">{GMAIL_USER}</code> to BCC on any property email</li>
          <li>The Gmail monitor picks it up within 15 minutes</li>
          <li>Inquiry records are created automatically</li>
          <li>No Outlook configuration required</li>
        </ul>
      </div>
    </div>
    <div class="card">
      <div class="card-header"><i class="bi bi-clock-history me-1" style="color:#f59e0b"></i>Recent Captures</div>
      <div class="card-body" style="padding:.7rem">
        {"".join(f'<div style="padding:.4rem 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:.76rem">'
                 f'<span style="color:#fff">{e.get("subject","")[:50]}</span>'
                 f'<span style="color:#475569;font-size:.68rem;margin-left:.5rem">{fmt_date(e.get("received_at"))}</span>'
                 f'</div>'
                 for e in sb_select("email_logs", order="received_at", limit=5)) or
         '<p style="color:#1e3a5f;font-size:.78rem;text-align:center;padding:1rem 0">No captures yet</p>'}
      </div>
    </div>
  </div>
</div>"""
    return render_layout(content, "Outlook Setup", "outlook")


# ── Call Logs ─────────────────────────────────────────────────────────────────
@app.route("/call-logs", methods=["GET", "POST"])
@login_required
def call_logs_page():
    upload_result = None
    if request.method == "POST":
        f = request.files.get("file")
        if f and f.filename:
            fname = secure_filename(f.filename)
            fpath = os.path.join(UPLOAD_FOLDER, fname)
            f.save(fpath)
            try:
                parsed = parse_call_log_file(fpath, fname)
                if parsed["errors"]:
                    flash(f"Parse errors: {'; '.join(parsed['errors'])}", "error")
                else:
                    upload_result = store_call_logs(parsed["rows"], fname)
                    flash(f"Call log: {upload_result['saved']} calls saved, "
                          f"{upload_result['updated_leads']} leads updated.", "success")
            except Exception as e:
                flash(f"Error: {e}", "error")
            finally:
                try: os.remove(fpath)
                except OSError: pass

    logs = sb_select("call_logs", order="call_date", limit=100)
    rows_html = ""
    for cl in logs:
        dur = cl.get("duration_sec") or 0
        dur_str = f"{dur//60}:{dur%60:02d}" if dur else "—"
        direction = cl.get("direction","")
        dir_col = {"Incoming":"#10b981","Outgoing":"#38bdf8","Missed":"#ef4444"}.get(direction,"#64748b")
        rows_html += f"""<tr>
          <td style="color:#fff">{fmt_date(cl.get('call_date'))}</td>
          <td style="color:{dir_col}">{direction}</td>
          <td style="color:#fff">{cl.get('contact_name') or '—'}</td>
          <td style="color:#94a3b8">{cl.get('number','—')}</td>
          <td style="color:#64748b">{dur_str}</td>
          <td style="color:#475569;font-size:.72rem">{cl.get('source_file','')[:20]}</td>
        </tr>"""
    if not rows_html:
        rows_html = '<tr><td colspan="6" style="text-align:center;color:#1e3a5f;padding:2rem">No call logs yet</td></tr>'

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-telephone-fill me-2" style="color:#f59e0b"></i>Call Logs</h4>
    <p>Import Android call log exports (CSV or JSON). Cross-referenced against contacts.</p>
  </div>
</div>
<div class="row g-3 mb-3">
  <div class="col-lg-5">
    <div class="card">
      <div class="card-header">Upload Call Log Export</div>
      <div class="card-body">
        <form method="POST" action="/call-logs" enctype="multipart/form-data">
          <input type="file" name="file" accept=".csv,.json" class="form-control mb-3"
                 style="font-size:.82rem"/>
          <button class="btn btn-primary w-100">
            <i class="bi bi-upload me-1"></i>Import Call Log</button>
        </form>
        <div style="font-size:.72rem;color:#475569;margin-top:.8rem">
          <strong style="color:#94a3b8">Supported formats:</strong>
          <ul style="margin:.3rem 0 0;padding-left:1.2rem;line-height:1.8">
            <li>Android call log CSV (Name, Number, Type, Date, Duration)</li>
            <li>JSON array with number/date/duration/type fields</li>
            <li>DRPU, Call Log Backup, or any standard export</li>
          </ul>
        </div>
      </div>
    </div>
  </div>
  <div class="col-lg-7">
    <div class="row g-3">
      <div class="col-4"><div class="kpi">
        <div class="kpi-val amber">{sb_count("call_logs")}</div>
        <div class="kpi-label">Total Calls</div>
      </div></div>
      <div class="col-4"><div class="kpi">
        <div class="kpi-val green">{len([c for c in logs if c.get("contact_id")])}</div>
        <div class="kpi-label">Matched</div>
      </div></div>
      <div class="col-4"><div class="kpi">
        <div class="kpi-val blue">{len([c for c in logs if c.get("direction")=="Outgoing"])}</div>
        <div class="kpi-label">Outgoing</div>
      </div></div>
    </div>
  </div>
</div>
<div class="card">
  <div class="card-header">Call History</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Date</th><th>Direction</th><th>Contact</th><th>Number</th>
             <th>Duration</th><th>Source</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table></div>
  </div>
</div>"""
    return render_layout(content, "Call Logs", "calllogs")


# ── Call Recordings ───────────────────────────────────────────────────────────
@app.route("/recordings", methods=["GET", "POST"])
@login_required
def recordings_page():
    result = None
    if request.method == "POST":
        f = request.files.get("file")
        if f and f.filename:
            fname = secure_filename(f.filename)
            ext   = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
            if ext not in AUDIO_EXTENSIONS:
                flash(f"Unsupported audio format. Accepted: {', '.join(sorted(AUDIO_EXTENSIONS))}", "error")
            else:
                fpath = os.path.join(UPLOAD_FOLDER, fname)
                f.save(fpath)
                try:
                    transcript = transcribe_audio_file(fpath, fname)
                    intel      = extract_call_intel(transcript)
                    style_ctx  = get_style_context()
                    # If we have style context, refine the follow-up draft
                    follow_up  = intel.get("follow_up_email_draft", "")
                    contact_name = intel.get("contact_name", "")
                    # Find matching contact
                    contact = None
                    if contact_name:
                        all_c = sb_select("contacts")
                        for c in all_c:
                            if contact_name.lower() in (c.get("name") or "").lower():
                                contact = c
                                break
                    rec = {
                        "filename":        fname,
                        "contact_name":    contact_name,
                        "contact_id":      contact["id"] if contact else None,
                        "transcript":      transcript[:20000],
                        "ai_summary":      intel.get("summary",""),
                        "ai_action_items": json.dumps(intel.get("action_items",[])),
                        "follow_up_draft": follow_up,
                        "key_facts":       json.dumps(intel.get("key_facts",{})),
                        "style_learned":   False,
                    }
                    inserted = sb_insert("call_recordings", rec)
                    rec_id = inserted[0]["id"] if inserted else None
                    # Update style profile
                    update_style_profile(transcript, intel.get("summary",""))
                    if inserted:
                        sb_update("call_recordings", rec_id, {"style_learned": True})
                    # Update lead if contact found
                    if contact:
                        lead_status = intel.get("lead_status","")
                        new_status = {"Hot":"New","Warm":"New","Cold":"Contacted"}.get(lead_status)
                        open_inqs  = sb_select("inquiries", {"contact_name": contact["name"]})
                        for inq in open_inqs[:1]:
                            if new_status:
                                sb_update("inquiries", inq["id"], {"status": new_status})
                    result = {"rec_id": rec_id, "intel": intel, "transcript": transcript}
                    flash(f"Recording transcribed and analysed. Contact: {contact_name or 'unknown'}", "success")
                except Exception as e:
                    log.exception("recording error")
                    flash(f"Error: {e}", "error")
                finally:
                    try: os.remove(fpath)
                    except OSError: pass

    recordings = sb_select("call_recordings", order="created_at", limit=50)
    result_html = ""
    if result:
        intel = result["intel"]
        actions = "".join(f"<li style='color:#94a3b8;font-size:.78rem'>{a}</li>"
                          for a in (intel.get("action_items") or []))
        follow_up = (intel.get("follow_up_email_draft") or "").strip()
        result_html = f"""
<div class="card mb-3">
  <div class="card-header" style="background:rgba(16,185,129,.08);border-color:rgba(16,185,129,.2)">
    <i class="bi bi-mic-fill me-1" style="color:#10b981"></i>Transcription & Analysis Complete
  </div>
  <div class="card-body">
    <div class="row g-3">
      <div class="col-md-6">
        <div style="color:#64748b;font-size:.7rem;margin-bottom:.3rem">SUMMARY</div>
        <div style="color:#94a3b8;font-size:.8rem">{intel.get('summary','')}</div>
        {f'<ul class="mt-2 mb-0">{actions}</ul>' if actions else ''}
      </div>
      <div class="col-md-6">
        <div style="color:#64748b;font-size:.7rem;margin-bottom:.3rem">TRANSCRIPT (first 500 chars)</div>
        <div style="color:#475569;font-size:.72rem;font-family:monospace;
             background:rgba(255,255,255,.02);padding:.5rem;border-radius:6px;
             max-height:100px;overflow:hidden">{result['transcript'][:500]}</div>
      </div>
    </div>
    {f'''<div class="mt-3">
      <div style="color:#64748b;font-size:.7rem;margin-bottom:.4rem">FOLLOW-UP EMAIL DRAFT</div>
      <textarea class="form-control" rows="8" style="font-family:monospace;font-size:.76rem">{follow_up}</textarea>
    </div>''' if follow_up else ''}
  </div>
</div>"""

    rec_rows = ""
    for r in recordings:
        actions = r.get("ai_action_items") or []
        if isinstance(actions, str):
            try: actions = json.loads(actions)
            except Exception: actions = []
        rec_rows += f"""<tr>
          <td style="color:#fff;font-size:.78rem">{r.get('filename','')[:35]}</td>
          <td style="color:#f59e0b">{r.get('contact_name') or '—'}</td>
          <td style="color:#94a3b8;font-size:.74rem;max-width:300px;white-space:nowrap;
              overflow:hidden;text-overflow:ellipsis">{r.get('ai_summary','')[:80]}</td>
          <td style="color:#64748b;font-size:.7rem">{len(actions)} items</td>
          <td style="color:#475569;font-size:.7rem">{'Yes' if r.get('style_learned') else '—'}</td>
          <td>{fmt_date(r.get('created_at'))}</td>
        </tr>"""
    if not rec_rows:
        rec_rows = '<tr><td colspan="6" style="text-align:center;color:#1e3a5f;padding:2rem">No recordings yet</td></tr>'

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-mic-fill me-2" style="color:#f59e0b"></i>Call Recordings</h4>
    <p>Upload audio from your call recorder. Whisper transcribes, GPT-4o extracts intel and drafts follow-ups.</p>
  </div>
</div>
{result_html}
<div class="row g-3 mb-3">
  <div class="col-lg-5">
    <div class="card">
      <div class="card-header">Upload Recording</div>
      <div class="card-body">
        <form method="POST" action="/recordings" enctype="multipart/form-data">
          <input type="file" name="file"
                 accept=".mp3,.mp4,.m4a,.wav,.webm,.ogg,.flac,.aac"
                 class="form-control mb-3" style="font-size:.82rem"/>
          <button class="btn btn-primary w-100">
            <i class="bi bi-cpu-fill me-1"></i>Transcribe & Analyse</button>
        </form>
        <div style="font-size:.72rem;color:#475569;margin-top:.8rem">
          Supported: MP3, M4A, MP4, WAV, WebM, OGG, FLAC, AAC<br/>
          Max {MAX_UPLOAD_MB}MB. Transcribed with OpenAI Whisper.
        </div>
      </div>
    </div>
  </div>
  <div class="col-lg-7">
    <div class="row g-3">
      <div class="col-4"><div class="kpi">
        <div class="kpi-val amber">{sb_count("call_recordings")}</div>
        <div class="kpi-label">Recordings</div>
      </div></div>
      <div class="col-4"><div class="kpi">
        <div class="kpi-val green">{sb_count("style_profile")}</div>
        <div class="kpi-label">Style Samples</div>
      </div></div>
      <div class="col-4"><div class="kpi">
        <div class="kpi-val blue">{len([r for r in recordings if r.get("follow_up_draft")])}</div>
        <div class="kpi-label">Drafts Ready</div>
      </div></div>
    </div>
  </div>
</div>
<div class="card">
  <div class="card-header">Recording History</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>File</th><th>Contact</th><th>Summary</th>
             <th>Actions</th><th>Style</th><th>Date</th></tr></thead>
      <tbody>{rec_rows}</tbody>
    </table></div>
  </div>
</div>"""
    return render_layout(content, "Recordings", "recordings")


# ── Calendar ──────────────────────────────────────────────────────────────────
@app.route("/calendar", methods=["GET", "POST"])
@login_required
def calendar_page():
    if request.method == "POST":
        action = request.form.get("action","")
        if action == "notes":
            eid  = request.form.get("event_id")
            note = request.form.get("notes","")
            if eid:
                sb_update("calendar_events", {"id": int(eid)}, {"post_meeting_notes": note})
                flash("Post-meeting notes saved. AI processing in background…", "success")
                # Fetch the event and run AI processing in background thread
                evs = sb_select("calendar_events", {"id": int(eid)})
                if evs and note.strip():
                    ev = evs[0]
                    threading.Thread(
                        target=process_meeting_notes,
                        args=(ev, note),
                        daemon=True
                    ).start()
        else:
            f = request.files.get("file")
            if f and f.filename:
                fname = secure_filename(f.filename)
                fpath = os.path.join(UPLOAD_FOLDER, fname)
                f.save(fpath)
                try:
                    events = parse_ics_file(fpath)
                    imported = skipped = 0
                    for ev in events:
                        if ev.get("uid"):
                            # upsert by uid
                            existing = sb_select("calendar_events", {"uid": ev["uid"]})
                            if existing:
                                skipped += 1
                                continue
                        if sb_insert("calendar_events", ev):
                            imported += 1
                    flash(f"Calendar: {imported} events imported, {skipped} already existed.", "success")
                except Exception as e:
                    flash(f"ICS parse error: {e}", "error")
                finally:
                    try: os.remove(fpath)
                    except OSError: pass

    events = sb_select("calendar_events", order="start_dt", limit=100)
    now_iso = datetime.now(timezone.utc).isoformat()

    upcoming = [e for e in events if (e.get("start_dt") or "") >= now_iso]
    past     = [e for e in events if (e.get("start_dt") or "") < now_iso]

    def event_row(e, show_notes_btn=False):
        prop_badge = ('<span class="badge" style="background:rgba(245,158,11,.15);'
                      'color:#f59e0b;font-size:.6rem">Property</span> '
                      if e.get("is_property_related") else "")
        brief_badge = ('<span class="badge bg-secondary" style="font-size:.6rem">Brief Sent</span> '
                       if e.get("brief_sent") else "")
        notes_snippet = (e.get("post_meeting_notes") or "")[:60]
        return (f'<tr>'
                f'<td style="color:#fff;font-size:.78rem">{e.get("title","")[:50]}</td>'
                f'<td style="color:#64748b;font-size:.74rem">{(e.get("start_dt") or "")[:16].replace("T"," ")}</td>'
                f'<td style="color:#94a3b8;font-size:.74rem">{e.get("location","")[:30] or "—"}</td>'
                f'<td>{prop_badge}{brief_badge}</td>'
                f'<td style="color:#475569;font-size:.7rem">{notes_snippet}</td>'
                f'<td style="white-space:nowrap">'
                f'{"<button class=\"btn btn-xs btn-outline-secondary me-1\" onclick=\"openNotes(" + str(e["id"]) + ",`" + (e.get("title","")[:40].replace("`","")) + "`)\">Notes</button>" if show_notes_btn else ""}'
                f'{"<button class=\"btn btn-xs btn-outline-warning\" onclick=\"viewAiResults(" + str(e["id"]) + ")\">AI</button>" if e.get("notes_processed_at") else ""}'
                f'</td>'
                f'</tr>')

    up_rows = "".join(event_row(e, True) for e in upcoming[:20]) or \
              '<tr><td colspan="6" style="text-align:center;color:#1e3a5f;padding:2rem">No upcoming events</td></tr>'
    past_rows = "".join(event_row(e, True) for e in past[:10]) or \
                '<tr><td colspan="6" style="text-align:center;color:#1e3a5f;padding:1rem">None</td></tr>'

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-calendar3 me-2" style="color:#f59e0b"></i>Calendar</h4>
    <p>Import ICS from Outlook. Pre-meeting briefs sent 30 min before property meetings.</p>
  </div>
</div>
<div class="row g-3 mb-3">
  <div class="col-lg-4">
    <div class="card">
      <div class="card-header">Import ICS File</div>
      <div class="card-body">
        <form method="POST" action="/calendar" enctype="multipart/form-data">
          <input type="file" name="file" accept=".ics" class="form-control mb-3"
                 style="font-size:.82rem"/>
          <button class="btn btn-primary w-100">
            <i class="bi bi-upload me-1"></i>Import Calendar</button>
        </form>
        <div style="font-size:.72rem;color:#475569;margin-top:.8rem">
          In Outlook: File → Open &amp; Export → Import/Export → Export to .ics
        </div>
      </div>
    </div>
  </div>
  <div class="col-lg-8">
    <div class="row g-3">
      <div class="col-4"><div class="kpi">
        <div class="kpi-val amber">{len(upcoming)}</div>
        <div class="kpi-label">Upcoming</div>
      </div></div>
      <div class="col-4"><div class="kpi">
        <div class="kpi-val green">{len([e for e in events if e.get("is_property_related")])}</div>
        <div class="kpi-label">Property Meetings</div>
      </div></div>
      <div class="col-4"><div class="kpi">
        <div class="kpi-val blue">{len([e for e in events if e.get("brief_sent")])}</div>
        <div class="kpi-label">Briefs Sent</div>
      </div></div>
    </div>
  </div>
</div>

<div class="card mb-3">
  <div class="card-header"><i class="bi bi-calendar-event me-1" style="color:#10b981"></i>Upcoming Events</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Title</th><th>When</th><th>Location</th><th>Tags</th><th>Notes</th><th></th></tr></thead>
      <tbody>{up_rows}</tbody>
    </table></div>
  </div>
</div>
<div class="card">
  <div class="card-header">Past Events</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0">
      <thead><tr><th>Title</th><th>When</th><th>Location</th><th>Tags</th><th>Notes</th><th></th></tr></thead>
      <tbody>{past_rows}</tbody>
    </table></div>
  </div>
</div>

<!-- Notes modal -->
<div id="notesModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);
     z-index:9999;align-items:center;justify-content:center">
  <div style="background:#1e293b;border-radius:12px;padding:1.5rem;width:92%;max-width:540px">
    <h6 style="color:#fff;margin-bottom:1rem" id="notesTitle">Post-meeting Notes</h6>
    <form method="POST" action="/calendar">
      <input type="hidden" name="action" value="notes"/>
      <input type="hidden" name="event_id" id="notesEventId"/>
      <textarea name="notes" id="notesTextarea" class="form-control mb-2" rows="6"
                style="background:rgba(255,255,255,.05);border-color:rgba(255,255,255,.12);color:#e2e8f0"
                placeholder="Key outcomes, follow-ups, observations… GPT-4o will extract action items and draft a follow-up email automatically."></textarea>
      <div style="font-size:.7rem;color:#475569;margin-bottom:.75rem">
        <i class="bi bi-robot me-1"></i>GPT-4o will auto-extract action items, update the contact record, and draft a follow-up email.
      </div>
      <div class="d-flex gap-2">
        <button type="submit" class="btn btn-primary">Save &amp; Process</button>
        <button type="button" class="btn btn-outline-secondary"
                onclick="document.getElementById('notesModal').style.display='none'">Cancel</button>
      </div>
    </form>
  </div>
</div>

<!-- AI results panel (shown after notes processed) -->
<div id="aiResultsModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);
     z-index:9999;align-items:center;justify-content:center">
  <div style="background:#1e293b;border-radius:12px;padding:1.5rem;width:92%;max-width:600px;
              max-height:80vh;overflow-y:auto">
    <h6 style="color:#f59e0b;margin-bottom:1rem">
      <i class="bi bi-robot me-1"></i>AI Meeting Analysis
    </h6>
    <div id="aiResultsContent" style="font-size:.8rem;color:#e2e8f0"></div>
    <button type="button" class="btn btn-outline-secondary mt-3"
            onclick="document.getElementById('aiResultsModal').style.display='none'">Close</button>
  </div>
</div>

<script>
function openNotes(id, title) {{
  document.getElementById('notesEventId').value = id;
  document.getElementById('notesTitle').textContent = 'Notes: ' + title;
  document.getElementById('notesTextarea').value = '';
  document.getElementById('notesModal').style.display = 'flex';
}}
function viewAiResults(id) {{
  fetch('/api/calendar/' + id + '/ai-results')
    .then(r => r.json())
    .then(d => {{
      let html = '';
      if (d.summary) html += '<p style="color:#94a3b8">' + d.summary + '</p>';
      if (d.action_items && d.action_items.length) {{
        html += '<div style="margin-bottom:.75rem"><div style="color:#f59e0b;font-weight:600;margin-bottom:.3rem">Action Items</div>';
        d.action_items.forEach(a => {{ html += '<div style="color:#e2e8f0;padding:.2rem 0">• ' + a + '</div>'; }});
        html += '</div>';
      }}
      if (d.follow_up_draft) {{
        html += '<div style="margin-bottom:.5rem"><div style="color:#f59e0b;font-weight:600;margin-bottom:.3rem">Follow-up Draft</div>';
        html += '<pre style="background:rgba(255,255,255,.05);padding:.75rem;border-radius:6px;white-space:pre-wrap;font-size:.72rem;color:#cbd5e1">' + d.follow_up_draft + '</pre></div>';
      }}
      if (!html) html = '<p style="color:#475569">AI processing pending or no results yet.</p>';
      document.getElementById('aiResultsContent').innerHTML = html;
      document.getElementById('aiResultsModal').style.display = 'flex';
    }})
    .catch(() => alert('Could not load AI results.'));
}}
</script>"""
    return render_layout(content, "Calendar", "calendar")


# ── Fee Tracking ──────────────────────────────────────────────────────────────
@app.route("/fees", methods=["GET", "POST"])
@login_required
def fees_page():
    if request.method == "POST":
        action = request.form.get("action","")
        f = request.files.get("file")
        if action == "schedule" and f and f.filename:
            fname = secure_filename(f.filename)
            fpath = os.path.join(UPLOAD_FOLDER, fname)
            f.save(fpath)
            try:
                r = parse_fee_schedule_excel(fpath)
                flash(f"Fee schedule: {r['imported']} client rates imported.", "success")
            except Exception as e:
                flash(f"Error: {e}", "error")
            finally:
                try: os.remove(fpath)
                except OSError: pass
        elif action == "forecast" and f and f.filename:
            fname = secure_filename(f.filename)
            fpath = os.path.join(UPLOAD_FOLDER, fname)
            f.save(fpath)
            try:
                result = process_file_universal(fpath, fname)
                flash(f"Fee forecast processed via AI intake — {result.get('ai_classification','')} detected.", "success")
            except Exception as e:
                flash(f"Error: {e}", "error")
            finally:
                try: os.remove(fpath)
                except OSError: pass
        elif action == "calc_deal":
            deal_id = request.form.get("deal_id")
            if deal_id:
                deals_list = sb_select("deals", {"id": int(deal_id)})
                if deals_list:
                    deal = deals_list[0]
                    fee_amt = calculate_fee_for_deal(deal)
                    if fee_amt:
                        sb_insert("fees", {
                            "deal_id":    deal["id"],
                            "client_name": deal.get("landlord_name") or deal.get("seller_name",""),
                            "landlord":   deal.get("landlord_name",""),
                            "deal_type":  deal.get("deal_type","Lease"),
                            "gross_value": float(deal.get("rent_pa") or deal.get("price") or 0),
                            "fee_amount": fee_amt,
                            "invoice_status": "Pending",
                        })
                        flash(f"Fee calculated: ${fee_amt:,.0f}", "success")

    fees_list     = sb_select("fees", order="created_at", limit=100)
    schedules     = sb_select("fee_schedules", limit=50)
    deals_closed  = sb_select("deals", {"status": "Completed"}, limit=50)
    total_fees    = sum(float(f.get("fee_amount") or 0) for f in fees_list)
    paid_fees     = sum(float(f.get("fee_amount") or 0) for f in fees_list if f.get("invoice_status")=="Paid")
    pending_fees  = total_fees - paid_fees

    # Revenue by client
    by_client = {}
    for f in fees_list:
        cl = f.get("client_name") or "Unknown"
        by_client.setdefault(cl, 0)
        by_client[cl] += float(f.get("fee_amount") or 0)
    client_rows = "".join(
        f'<tr><td style="color:#fff">{cl}</td><td style="color:#f59e0b">${amt:,.0f}</td>'
        f'<td style="color:#64748b">{round(amt/total_fees*100 if total_fees else 0,1)}%</td></tr>'
        for cl, amt in sorted(by_client.items(), key=lambda x: -x[1])
    ) or '<tr><td colspan="3" style="text-align:center;color:#1e3a5f;padding:1.5rem">No fees recorded</td></tr>'

    fee_rows = ""
    for f in fees_list:
        status = f.get("invoice_status","Pending")
        st_col = {"Paid":"#10b981","Pending":"#f59e0b","Overdue":"#ef4444"}.get(status,"#64748b")
        fee_rows += (f'<tr>'
                     f'<td style="color:#fff">{f.get("client_name","—")}</td>'
                     f'<td style="color:#94a3b8">{f.get("deal_type","—")}</td>'
                     f'<td style="color:#f59e0b">${float(f.get("fee_amount") or 0):,.0f}</td>'
                     f'<td><span style="color:{st_col};font-size:.74rem">{status}</span></td>'
                     f'<td style="color:#475569;font-size:.7rem">{fmt_date(f.get("created_at"))}</td>'
                     f'</tr>')
    if not fee_rows:
        fee_rows = '<tr><td colspan="5" style="text-align:center;color:#1e3a5f;padding:2rem">No fees yet — calculate from a closed deal below</td></tr>'

    deal_options = "".join(
        f'<option value="{d["id"]}">{d.get("address") or d.get("tenant_name","Deal "+str(d["id"]))} — {d.get("deal_type","")}</option>'
        for d in deals_closed
    )

    schedule_rows = "".join(
        f'<tr><td style="color:#fff">{s.get("client_name","")}</td>'
        f'<td style="color:#64748b">{s.get("deal_type","")}</td>'
        f'<td style="color:#f59e0b">{s.get("fee_pct") or "—"}%</td>'
        f'<td style="color:#94a3b8">${float(s.get("flat_fee") or 0):,.0f}</td>'
        f'<td style="color:#94a3b8">${float(s.get("min_fee") or 0):,.0f}</td></tr>'
        for s in schedules
    ) or '<tr><td colspan="5" style="text-align:center;color:#1e3a5f;padding:1rem">No schedules — upload one below</td></tr>'

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-currency-dollar me-2" style="color:#f59e0b"></i>Fee Tracking</h4>
    <p>Calculate and track brokerage fees on closed deals. Upload institutional client rate cards.</p>
  </div>
</div>
<div class="row g-3 mb-4">
  <div class="col-6 col-lg-3"><div class="kpi">
    <div class="kpi-val amber">${total_fees:,.0f}</div>
    <div class="kpi-label">Total Fees</div>
  </div></div>
  <div class="col-6 col-lg-3"><div class="kpi">
    <div class="kpi-val green">${paid_fees:,.0f}</div>
    <div class="kpi-label">Paid</div>
  </div></div>
  <div class="col-6 col-lg-3"><div class="kpi">
    <div class="kpi-val red">${pending_fees:,.0f}</div>
    <div class="kpi-label">Pending</div>
  </div></div>
  <div class="col-6 col-lg-3"><div class="kpi">
    <div class="kpi-val blue">{len(fees_list)}</div>
    <div class="kpi-label">Invoices</div>
  </div></div>
</div>

<div class="row g-3 mb-3">
  <!-- Calculate fee -->
  <div class="col-lg-4">
    <div class="card mb-3">
      <div class="card-header">Calculate Fee from Deal</div>
      <div class="card-body">
        <form method="POST" action="/fees">
          <input type="hidden" name="action" value="calc_deal"/>
          <div class="mb-3">
            <label class="form-label" style="font-size:.8rem">Closed Deal</label>
            <select name="deal_id" class="form-select form-select-sm">
              <option value="">— Select deal —</option>{deal_options}
            </select>
          </div>
          <button class="btn btn-primary w-100 btn-sm">
            <i class="bi bi-calculator me-1"></i>Calculate Fee</button>
        </form>
      </div>
    </div>
    <div class="card">
      <div class="card-header">Upload Rate Card</div>
      <div class="card-body">
        <form method="POST" action="/fees" enctype="multipart/form-data">
          <input type="hidden" name="action" value="schedule"/>
          <input type="file" name="file" accept=".xlsx,.xls" class="form-control form-control-sm mb-2"/>
          <button class="btn btn-outline-secondary w-100 btn-sm">
            <i class="bi bi-upload me-1"></i>Import Fee Schedule</button>
        </form>
        <form method="POST" action="/fees" enctype="multipart/form-data" class="mt-2">
          <input type="hidden" name="action" value="forecast"/>
          <input type="file" name="file" accept=".xlsx,.xls,.pdf" class="form-control form-control-sm mb-2"/>
          <button class="btn btn-outline-secondary w-100 btn-sm">
            <i class="bi bi-cpu me-1"></i>AI Import Fee Forecast</button>
        </form>
        <div style="font-size:.7rem;color:#475569;margin-top:.6rem">
          Institutional clients: {', '.join(INSTITUTIONAL_CLIENTS[:5])} etc.
        </div>
      </div>
    </div>
  </div>

  <!-- Revenue by client -->
  <div class="col-lg-4">
    <div class="card h-100">
      <div class="card-header">Revenue by Client</div>
      <div class="card-body p-0">
        <table class="table table-sm mb-0" style="font-size:.78rem">
          <thead><tr><th>Client</th><th>Fees</th><th>Share</th></tr></thead>
          <tbody>{client_rows}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Fee schedule -->
  <div class="col-lg-4">
    <div class="card h-100">
      <div class="card-header">Client Rate Cards</div>
      <div class="card-body p-0">
        <table class="table table-sm mb-0" style="font-size:.76rem">
          <thead><tr><th>Client</th><th>Type</th><th>%</th><th>Flat</th><th>Min</th></tr></thead>
          <tbody>{schedule_rows}</tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<div class="card">
  <div class="card-header">All Fee Records</div>
  <div class="card-body p-0">
    <div class="table-responsive">
    <table class="table table-hover mb-0" style="font-size:.8rem">
      <thead><tr><th>Client / Landlord</th><th>Type</th><th>Fee</th><th>Status</th><th>Date</th></tr></thead>
      <tbody>{fee_rows}</tbody>
    </table></div>
  </div>
</div>"""
    return render_layout(content, "Fee Tracking", "fees")


# ── Landlord Portfolio ─────────────────────────────────────────────────────────
@app.route("/landlords")
@login_required
def landlords_page():
    portfolios = build_landlord_portfolios()
    cards = ""
    for p in portfolios:
        if p["name"] == "Unknown Landlord" and not p["properties"]:
            continue
        vacancy_pct = round(p["vacant_sqm"] / p["total_sqm"] * 100 if p["total_sqm"] else 0, 1)
        vac_col = "#ef4444" if vacancy_pct > 30 else ("#f59e0b" if vacancy_pct > 10 else "#10b981")

        def _prop_row(pr):
            exp = pr.get("lease_expiry")
            exp_html = (f'<span style="font-size:.65rem;color:#ef4444">Exp {str(exp)[:10]}</span>'
                        if exp else "")
            return (f'<div style="display:flex;gap:.5rem;align-items:flex-start;padding:.35rem 0;'
                    f'border-bottom:1px solid rgba(255,255,255,.04)">'
                    f'<div style="flex:1;min-width:0">'
                    f'<div style="font-size:.76rem;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'
                    f'{pr.get("address","—")}</div>'
                    f'<div style="font-size:.68rem;color:#475569">'
                    f'{pr.get("occupier") or "Vacant"} &bull; {pr.get("size_sqm","—")} sqm</div>'
                    f'</div>'
                    f'<div style="text-align:right;flex-shrink:0">{exp_html}</div></div>')
        prop_items = "".join(_prop_row(pr) for pr in p["properties"][:6])
        more = len(p["properties"]) - 6
        if more > 0:
            prop_items += f'<div style="font-size:.68rem;color:#334155;text-align:center;padding:.3rem">+{more} more</div>'

        expiring_html = ""
        if p["expiring_soon"]:
            expiring_html = '<div style="margin-top:.5rem"><div style="font-size:.68rem;color:#f59e0b;margin-bottom:.25rem">EXPIRING &lt;12 MONTHS</div>'
            for ex in sorted(p["expiring_soon"], key=lambda x: x["months_away"]):
                urgency = "#ef4444" if ex["months_away"] < 3 else "#f59e0b"
                expiring_html += (f'<div style="font-size:.7rem;color:{urgency};padding:.15rem 0">'
                                  f'{ex["address"][:35]} — {ex["occupier"] or "?"} '
                                  f'({ex["months_away"]}mo)</div>')
            expiring_html += '</div>'

        cards += f"""
<div class="col-lg-6">
  <div class="card">
    <div class="card-header" style="display:flex;justify-content:space-between;align-items:center">
      <span style="font-size:.85rem;font-weight:800;color:#fff">{p['name']}</span>
      <div style="display:flex;gap:.5rem">
        <span style="background:rgba(245,158,11,.15);color:#f59e0b;border-radius:5px;
              padding:.15rem .45rem;font-size:.65rem">{len(p['properties'])} bldgs</span>
        <span style="background:{vac_col}22;color:{vac_col};border-radius:5px;
              padding:.15rem .45rem;font-size:.65rem">{vacancy_pct}% vacant</span>
      </div>
    </div>
    <div class="card-body" style="padding:.85rem">
      <div class="row g-2 mb-3">
        <div class="col-4" style="text-align:center">
          <div style="font-size:1.1rem;font-weight:800;color:#f59e0b">{p['total_sqm']:,.0f}</div>
          <div style="font-size:.62rem;color:#475569">Total sqm</div>
        </div>
        <div class="col-4" style="text-align:center">
          <div style="font-size:1.1rem;font-weight:800;color:{vac_col}">{p['vacant_sqm']:,.0f}</div>
          <div style="font-size:.62rem;color:#475569">Vacant sqm</div>
        </div>
        <div class="col-4" style="text-align:center">
          <div style="font-size:1.1rem;font-weight:800;color:#10b981">${p['total_fees']:,.0f}</div>
          <div style="font-size:.62rem;color:#475569">Fees earned</div>
        </div>
      </div>
      {prop_items}
      {expiring_html}
    </div>
  </div>
</div>"""

    if not cards:
        cards = """<div class="col-12"><div style="text-align:center;padding:4rem 1rem;color:#1e3a5f">
          <i class="bi bi-buildings" style="font-size:2rem;opacity:.3;display:block;margin-bottom:.8rem"></i>
          No landlord data yet — upload an asset register via
          <a href="/upload" style="color:#f59e0b">AI Intake</a>
        </div></div>"""

    total_landlords = len([p for p in portfolios if p["name"] != "Unknown Landlord"])
    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-person-workspace me-2" style="color:#f59e0b"></i>Landlord Portfolios</h4>
    <p>{total_landlords} landlords &bull; built automatically from asset register data</p>
  </div>
  <a href="/upload" class="btn btn-sm btn-outline-secondary">
    <i class="bi bi-cloud-arrow-up-fill me-1"></i>Update from Asset Register</a>
</div>
<div class="row g-3">
  {cards}
</div>"""
    return render_layout(content, "Landlord Portfolios", "landlords")


@app.route("/email-draft", methods=["GET", "POST"])
@login_required
def email_draft_page():
    draft = ""
    error = ""
    form_to      = request.form.get("to", "")
    form_subject = request.form.get("subject", "")
    form_prompt  = request.form.get("prompt", "")

    if request.method == "POST" and form_prompt.strip():
        client = get_openai()
        if not client:
            error = "OpenAI API key not configured — add OPENAI_API_KEY to .env"
        else:
            style_ctx = get_style_context(n=5)
            system_msg = (
                "You are a ghostwriter for Michael, a West Melbourne industrial property agent "
                "(State Manager). Write professional but warm emails in his voice. "
                "Never use stiff corporate language. Be direct and specific.\n\n"
                + (f"{style_ctx}\n\n" if style_ctx else "")
                + "Return ONLY the email text, starting with SUBJECT: on the first line, "
                "then a blank line, then the email body. No commentary."
            )
            try:
                resp = client.chat.completions.create(
                    model=OPENAI_MODEL,
                    messages=[
                        {"role": "system", "content": system_msg},
                        {"role": "user",   "content":
                            f"Write an email{' to ' + form_to if form_to else ''}.\n"
                            f"{'Subject hint: ' + form_subject + chr(10) if form_subject else ''}"
                            f"Instruction: {form_prompt}"},
                    ],
                    temperature=0.6,
                    max_tokens=1200,
                )
                draft = resp.choices[0].message.content.strip()
            except Exception as e:
                error = f"GPT error: {e}"

    style_samples = sb_select("style_profile", order="created_at", limit=3)
    style_count   = len(sb_select("style_profile", limit=200))

    content = f"""
<div class="ph">
  <div>
    <h4><i class="bi bi-pencil-square me-2" style="color:#f59e0b"></i>Email Drafter</h4>
    <p>Write emails in your voice — style learned from {style_count} call recording{'s' if style_count!=1 else ''}.</p>
  </div>
</div>

<div class="row g-3">
  <div class="col-lg-5">
    <div class="card">
      <div class="card-header">Compose</div>
      <div class="card-body">
        <form method="POST" action="/email-draft">
          <div class="mb-3">
            <label class="form-label" style="font-size:.78rem;color:#94a3b8">To (optional)</label>
            <input type="text" name="to" class="form-control"
                   style="background:rgba(255,255,255,.05);border-color:rgba(255,255,255,.12);color:#e2e8f0"
                   value="{form_to}" placeholder="Recipient name or company"/>
          </div>
          <div class="mb-3">
            <label class="form-label" style="font-size:.78rem;color:#94a3b8">Subject hint (optional)</label>
            <input type="text" name="subject" class="form-control"
                   style="background:rgba(255,255,255,.05);border-color:rgba(255,255,255,.12);color:#e2e8f0"
                   value="{form_subject}" placeholder="e.g. Follow-up re: 45 Appleton St inspection"/>
          </div>
          <div class="mb-3">
            <label class="form-label" style="font-size:.78rem;color:#94a3b8">What to write <span style="color:#ef4444">*</span></label>
            <textarea name="prompt" class="form-control" rows="5"
                      style="background:rgba(255,255,255,.05);border-color:rgba(255,255,255,.12);color:#e2e8f0"
                      placeholder="e.g. Follow up on yesterday's inspection at 45 Appleton St West Melbourne. The tenant seemed interested in 2,500 sqm. Rent was $180/sqm. Ask if they want to submit an offer."
                      >{form_prompt}</textarea>
          </div>
          {'<div class="alert alert-danger py-2" style="font-size:.78rem">' + error + '</div>' if error else ''}
          <button class="btn btn-primary w-100">
            <i class="bi bi-robot me-1"></i>Generate Draft
          </button>
        </form>

        {'<div style="margin-top:1.2rem"><div style="font-size:.72rem;color:#64748b;margin-bottom:.5rem">Style profile — recent phrases:</div>' + "".join(
            '<div style="font-size:.7rem;color:#475569;padding:.2rem 0">'
            + "".join('<span style="background:rgba(245,158,11,.1);color:#f59e0b;border-radius:4px;padding:.1rem .35rem;margin:.1rem;display:inline-block;font-size:.67rem">' + ph + '</span>'
                      for ph in (json.loads(s.get("key_phrases","[]")) if isinstance(s.get("key_phrases"), str) else (s.get("key_phrases") or []))[:4])
            + '</div>'
            for s in style_samples
        ) + '</div>' if style_samples else '<div style="margin-top:1rem;font-size:.72rem;color:#334155">Upload call recordings to build your style profile.</div>'}
      </div>
    </div>
  </div>

  <div class="col-lg-7">
    <div class="card" style="height:100%">
      <div class="card-header" style="display:flex;justify-content:space-between;align-items:center">
        <span>Draft</span>
        {'<button onclick="copyDraft()" class="btn btn-xs btn-outline-secondary"><i class="bi bi-clipboard me-1"></i>Copy</button>' if draft else ''}
      </div>
      <div class="card-body">
        {'<pre id="draftText" style="white-space:pre-wrap;font-size:.8rem;color:#e2e8f0;background:rgba(255,255,255,.03);padding:1rem;border-radius:8px;min-height:300px;border:1px solid rgba(255,255,255,.06)">' + draft + '</pre>' if draft else
         '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:300px;color:#334155">'
         '<i class="bi bi-pencil-square" style="font-size:2.5rem;margin-bottom:.75rem;opacity:.25"></i>'
         '<div style="font-size:.82rem">Describe what to write in the form →</div>'
         '</div>'}
      </div>
    </div>
  </div>
</div>

<script>
function copyDraft() {{
  var txt = document.getElementById('draftText');
  if (!txt) return;
  navigator.clipboard.writeText(txt.innerText)
    .then(() => {{ var b = event.target.closest('button'); b.textContent = 'Copied!'; setTimeout(() => b.innerHTML = '<i class=\"bi bi-clipboard me-1\"></i>Copy', 2000); }});
}}
</script>"""
    return render_layout(content, "Email Drafter", "emaildraft")


@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "supabase": bool(SUPABASE_URL),
        "twilio": bool(TWILIO_SID),
        "gmail": bool(GMAIL_USER),
        "scheduler_running": scheduler.running,
        "jobs": [j.id for j in scheduler.get_jobs()],
        "timestamp": datetime.now(timezone.utc).isoformat()
    })

# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    init_schema()
    if not scheduler.running:
        scheduler.start()
        log.info("APScheduler started. Jobs: %s", [j.id for j in scheduler.get_jobs()])
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "false").lower() == "true"
    log.info("Starting West Melbourne Industrial Property Agent on port %d", port)
    app.run(host="0.0.0.0", port=port, debug=debug, use_reloader=False)
