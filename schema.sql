-- LangMind existing Business Memory schema + additive Quotation AI schema.
-- Safe to run more than once. No existing table is dropped or replaced.

CREATE TABLE IF NOT EXISTS business_memories (
 id UUID PRIMARY KEY,
 workspace_id VARCHAR(100) NOT NULL DEFAULT 'default',
 category VARCHAR(30) NOT NULL,
 title VARCHAR(300) NOT NULL,
 content TEXT NOT NULL,
 data JSONB NOT NULL DEFAULT '{}'::jsonb,
 entity_key VARCHAR(200),
 source_filename VARCHAR(255),
 source_type VARCHAR(50),
 source_date TIMESTAMPTZ,
 tags TEXT[] NOT NULL DEFAULT '{}',
 confidence DOUBLE PRECISION,
 status VARCHAR(20) NOT NULL DEFAULT 'active',
 created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
 updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
 deleted_at TIMESTAMPTZ
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_business_memory_entity
 ON business_memories(workspace_id, category, entity_key)
 WHERE entity_key IS NOT NULL AND status = 'active';
CREATE INDEX IF NOT EXISTS ix_business_memory_category ON business_memories(workspace_id, category, status);
CREATE INDEX IF NOT EXISTS ix_business_memory_updated ON business_memories(workspace_id, updated_at DESC);

CREATE TABLE IF NOT EXISTS business_memory_revisions (
 revision_id BIGSERIAL PRIMARY KEY,
 memory_id UUID NOT NULL,
 workspace_id VARCHAR(100) NOT NULL,
 action VARCHAR(20) NOT NULL,
 old_record JSONB,
 new_record JSONB,
 change_reason VARCHAR(500),
 changed_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Quotation AI: temporary review queue.
CREATE TABLE IF NOT EXISTS quotation_drafts (
 id UUID PRIMARY KEY,
 workspace_id VARCHAR(100) NOT NULL DEFAULT 'default',
 source_filename VARCHAR(255) NOT NULL,
 source_mime_type VARCHAR(100) NOT NULL,
 extracted_data JSONB NOT NULL,
 match_report JSONB NOT NULL DEFAULT '{}'::jsonb,
 reviewer_note TEXT,
 status VARCHAR(20) NOT NULL DEFAULT 'pending_review',
 created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
 updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
 confirmed_quotation_id UUID
);
CREATE INDEX IF NOT EXISTS ix_quotation_drafts_workspace ON quotation_drafts(workspace_id, status, created_at DESC);

-- Confirmed quotation header and line items.
CREATE TABLE IF NOT EXISTS quotations (
 id UUID PRIMARY KEY,
 draft_id UUID NOT NULL UNIQUE REFERENCES quotation_drafts(id),
 workspace_id VARCHAR(100) NOT NULL DEFAULT 'default',
 quotation_number VARCHAR(120),
 quotation_date DATE,
 valid_until DATE,
 customer_name VARCHAR(500),
 customer_code VARCHAR(200),
 currency VARCHAR(10) NOT NULL DEFAULT 'THB',
 subtotal NUMERIC(18,2) NOT NULL DEFAULT 0,
 discount_total NUMERIC(18,2) NOT NULL DEFAULT 0,
 tax_rate NUMERIC(8,4) NOT NULL DEFAULT 0,
 tax_total NUMERIC(18,2) NOT NULL DEFAULT 0,
 grand_total NUMERIC(18,2) NOT NULL DEFAULT 0,
 notes TEXT,
 source_filename VARCHAR(255),
 created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS ix_quotations_workspace_date ON quotations(workspace_id, quotation_date DESC);

CREATE TABLE IF NOT EXISTS quotation_items (
 id UUID PRIMARY KEY,
 quotation_id UUID NOT NULL REFERENCES quotations(id) ON DELETE CASCADE,
 line_no INTEGER NOT NULL,
 product_code VARCHAR(200),
 product_name VARCHAR(1000) NOT NULL,
 description TEXT,
 quantity NUMERIC(18,4) NOT NULL DEFAULT 0,
 unit VARCHAR(80),
 unit_price NUMERIC(18,2) NOT NULL DEFAULT 0,
 discount NUMERIC(18,2) NOT NULL DEFAULT 0,
 line_total NUMERIC(18,2) NOT NULL DEFAULT 0,
 match_status VARCHAR(30) NOT NULL DEFAULT 'unmatched'
);
CREATE INDEX IF NOT EXISTS ix_quotation_items_quotation ON quotation_items(quotation_id, line_no);
