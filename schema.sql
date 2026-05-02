-- ============================================================
-- MJR West Industrial Property Intelligence Agent
-- Full database schema — paste into Supabase SQL Editor and run
-- Run PART 1 first, then PART 2 (requires pgvector extension)
-- ============================================================


-- ============================================================
-- PART 1 — Core tables, indexes, and backfill columns
-- ============================================================

create table if not exists properties (
    id             bigserial primary key,
    address        text not null,
    suburb         text default 'West Melbourne',
    property_type  text default 'Warehouse',
    size_sqm       numeric,
    land_sqm       numeric,
    asking_price   numeric,
    asking_rent_pa numeric,
    status         text default 'Available',
    agent_name     text,
    agent_phone    text,
    agent_email    text,
    year_built     int,
    zoning         text,
    notes          text,
    source         text,
    created_at     timestamptz default now(),
    updated_at     timestamptz default now()
);

create table if not exists contacts (
    id               bigserial primary key,
    name             text not null,
    company          text,
    phone            text,
    email            text,
    contact_type     text default 'Agent',
    notes            text,
    whatsapp_opt_in  bool default false,
    created_at       timestamptz default now()
);

create table if not exists inquiries (
    id             bigserial primary key,
    contact_name   text,
    contact_phone  text,
    contact_email  text,
    property_id    bigint references properties(id),
    source         text default 'Web',
    message        text,
    status         text default 'New',
    created_at     timestamptz default now()
);

create table if not exists deals (
    id              bigserial primary key,
    property_id     bigint references properties(id),
    deal_type       text default 'Sale',
    price           numeric,
    settlement_date date,
    buyer_name      text,
    seller_name     text,
    agent_name      text,
    status          text default 'Negotiation',
    notes           text,
    created_at      timestamptz default now()
);

create table if not exists briefings (
    id             bigserial primary key,
    briefing_type  text default 'Daily',
    content        text,
    sent_to        text,
    channel        text default 'WhatsApp',
    created_at     timestamptz default now()
);

create table if not exists email_logs (
    id                 bigserial primary key,
    sender             text,
    sender_email       text,
    sender_name        text,
    subject            text,
    body               text,
    message_id         text,
    property_id        bigint,
    contact_id         bigint,
    source             text default 'gmail_imap',
    processed          bool default false,
    reply_status       text default 'pending',
    replied_at         timestamptz,
    action_items       jsonb default '[]',
    ai_summary         text,
    ai_priority        text default 'medium',
    requires_action    bool default false,
    flagged_unanswered bool default false,
    save_triggered     bool default false,
    received_at        timestamptz default now()
);

-- Prevents re-processing the same email across IMAP polls
create unique index if not exists email_logs_message_id_idx
    on email_logs(message_id) where message_id is not null;

-- Backfill for email_logs (safe on existing databases)
alter table email_logs add column if not exists sender_email       text;
alter table email_logs add column if not exists sender_name        text;
alter table email_logs add column if not exists message_id         text;
alter table email_logs add column if not exists contact_id         bigint;
alter table email_logs add column if not exists source             text default 'gmail_imap';
alter table email_logs add column if not exists reply_status       text default 'pending';
alter table email_logs add column if not exists replied_at         timestamptz;
alter table email_logs add column if not exists action_items       jsonb default '[]';
alter table email_logs add column if not exists ai_summary         text;
alter table email_logs add column if not exists ai_priority        text default 'medium';
alter table email_logs add column if not exists requires_action    bool default false;
alter table email_logs add column if not exists flagged_unanswered bool default false;
alter table email_logs add column if not exists save_triggered     bool default false;

create table if not exists whatsapp_logs (
    id           bigserial primary key,
    direction    text default 'Inbound',
    from_number  text,
    to_number    text,
    body         text,
    created_at   timestamptz default now()
);

-- ── West Melbourne domain tables ──────────────────────────────────────────────

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
    id                 bigserial primary key,
    company            text,
    size_min_sqm       numeric,
    size_max_sqm       numeric,
    preferred_location text default 'West Melbourne',
    region             text,
    rating             text,
    agent              text,
    notes              text,
    document_source    text,
    contact_id         bigint references contacts(id),
    document_id        bigint,
    created_at         timestamptz default now()
);

alter table requirements add column if not exists region text;

create table if not exists market_data (
    id                bigserial primary key,
    address           text,
    suburb            text default 'West Melbourne',
    size_sqm          numeric,
    deal_type         text default 'Lease',
    tenant            text,
    landlord          text,
    lease_term_years  numeric,
    commencement_date date,
    expiry_date       date,
    rent_pa           numeric,
    rent_psm          numeric,
    incentive_months  numeric,
    outgoings_psm     numeric,
    document_source   text,
    property_id       bigint references properties(id),
    document_id       bigint,
    created_at        timestamptz default now()
);

-- ── Extra columns backfill ────────────────────────────────────────────────────

-- properties — asset register fields
alter table properties add column if not exists occupier     text;
alter table properties add column if not exists landlord     text;
alter table properties add column if not exists grade        text;
alter table properties add column if not exists lease_expiry date;

-- contacts — relationship tracking
alter table contacts add column if not exists last_contacted_at  timestamptz;
alter table contacts add column if not exists lead_status        text default 'Warm';

-- deals — lease deal fields
alter table deals add column if not exists address           text;
alter table deals add column if not exists suburb            text;
alter table deals add column if not exists size_sqm          numeric;
alter table deals add column if not exists rent_pa           numeric;
alter table deals add column if not exists tenant_name       text;
alter table deals add column if not exists landlord_name     text;
alter table deals add column if not exists term_years        numeric;
alter table deals add column if not exists commencement_date date;

-- ── Vacancy–requirement matching ──────────────────────────────────────────────

create table if not exists vacancy_matches (
    id             bigserial primary key,
    vacancy_id     bigint references vacancies(id),
    requirement_id bigint references requirements(id),
    score          numeric default 0,
    alerted        bool default false,
    created_at     timestamptz default now()
);

-- ── Field tools ───────────────────────────────────────────────────────────────

create table if not exists call_logs (
    id           bigserial primary key,
    call_date    timestamptz,
    duration_sec int,
    number       text,
    direction    text default 'Unknown',
    contact_name text,
    contact_id   bigint references contacts(id),
    inquiry_id   bigint references inquiries(id),
    notes        text,
    source_file  text,
    created_at   timestamptz default now()
);

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

-- ── Calendar ──────────────────────────────────────────────────────────────────

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

-- calendar_events — post-meeting AI processing columns
alter table calendar_events add column if not exists contact_id         bigint;
alter table calendar_events add column if not exists ai_action_items    jsonb default '[]';
alter table calendar_events add column if not exists follow_up_draft    text;
alter table calendar_events add column if not exists notes_processed_at timestamptz;

-- ── Fee tracking ──────────────────────────────────────────────────────────────

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

create table if not exists fees (
    id             bigserial primary key,
    deal_id        bigint references deals(id),
    property_id    bigint references properties(id),
    client_name    text,
    landlord       text,
    deal_type      text default 'Lease',
    gross_value    numeric,
    fee_pct        numeric,
    fee_amount     numeric,
    invoice_status text default 'Pending',
    paid_date      date,
    notes          text,
    created_at     timestamptz default now()
);

-- ── Style learning ────────────────────────────────────────────────────────────

create table if not exists style_profile (
    id                  bigserial primary key,
    sample_type         text default 'call',
    raw_text            text,
    key_phrases         jsonb default '[]',
    communication_notes text,
    created_at          timestamptz default now()
);


-- ============================================================
-- PART 2 — pgvector extension, documents table, and
--           match_documents RPC function
-- Enable the "vector" extension in Supabase first:
--   Dashboard → Database → Extensions → search "vector" → Enable
-- Then run this block.
-- ============================================================

create extension if not exists vector;

create table if not exists documents (
    id                     bigserial primary key,
    filename               text,
    file_type              text,
    raw_text               text,
    ai_classification      text default 'unknown',
    ai_summary             text,
    ai_confidence          numeric default 0,
    ai_urgency             text default 'low',
    action_items           jsonb default '[]',
    key_facts              jsonb default '{}',
    extracted_properties   int default 0,
    extracted_contacts     int default 0,
    extracted_inquiries    int default 0,
    extracted_deals        int default 0,
    extracted_vacancies    int default 0,
    extracted_requirements int default 0,
    extracted_market_data  int default 0,
    mentioned_companies    jsonb default '[]',
    linked_property_ids    jsonb default '[]',
    linked_contact_ids     jsonb default '[]',
    linked_document_ids    jsonb default '[]',
    embedding              vector(1536),
    processing_status      text default 'pending',
    error_message          text,
    created_at             timestamptz default now()
);

-- Backfill for documents (safe on existing databases)
alter table documents add column if not exists extracted_vacancies    int default 0;
alter table documents add column if not exists extracted_requirements int default 0;
alter table documents add column if not exists extracted_market_data  int default 0;
alter table documents add column if not exists mentioned_companies    jsonb default '[]';
alter table documents add column if not exists linked_property_ids    jsonb default '[]';
alter table documents add column if not exists linked_contact_ids     jsonb default '[]';
alter table documents add column if not exists linked_document_ids    jsonb default '[]';

-- IVFFlat index for fast cosine similarity search
-- Note: requires at least 100 rows to be effective; safe to create on empty table
create index if not exists documents_embedding_idx
    on documents using ivfflat (embedding vector_cosine_ops) with (lists = 100);

-- RPC function called by the app for server-side semantic search
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
