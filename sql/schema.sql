-- =============================================================
-- schema.sql - Create tables and indexes (run once)
-- =============================================================

-- Fail fast if another session holds a lock (avoids silent hang)
SET lock_timeout = '30s';

-- 1. Main target table
CREATE TABLE IF NOT EXISTS estabelecimentos_crm (
    cnpj VARCHAR(14) PRIMARY KEY,
    nome_fantasia TEXT,
    razao_social TEXT,
    data_inicio_atividade VARCHAR(8),
    cnae_fiscal_principal VARCHAR(7),
    cnae_fiscal_secundaria TEXT,
    tipo_logradouro TEXT,
    logradouro TEXT,
    numero TEXT,
    complemento TEXT,
    bairro TEXT,
    cep VARCHAR(8),
    uf CHAR(2),
    municipio VARCHAR(4),
    municipio_nome TEXT,
    telefone_1 TEXT,
    telefone_2 TEXT,
    email TEXT,
    regime_tributario TEXT,
    primeiro_contato DATE,
    situacao TEXT,
    segundo_contato DATE
);

-- 2. Company table (legal name source)
CREATE TABLE IF NOT EXISTS empresas_dados (
    cnpj_basico VARCHAR(8) PRIMARY KEY,
    razao_social TEXT,
    natureza_juridica VARCHAR(4),
    capital_social TEXT,
    porte_empresa VARCHAR(10)
);

-- 3. CNAE category rules derived only from Receita/government fields.
-- match_type:
--   exact  = cnae must equal cnae_inicio
--   prefix = cnae must start with cnae_inicio
--   range  = first two CNAE digits must be between cnae_inicio and cnae_fim
CREATE TABLE IF NOT EXISTS cnae_categoria_map (
    categoria TEXT NOT NULL,
    match_type TEXT NOT NULL CHECK (match_type IN ('exact', 'prefix', 'range')),
    cnae_inicio VARCHAR(7) NOT NULL,
    cnae_fim VARCHAR(7) NOT NULL DEFAULT '',
    prioridade INTEGER NOT NULL DEFAULT 100,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    observacao TEXT,
    PRIMARY KEY (categoria, match_type, cnae_inicio, cnae_fim)
);

-- 4. Per-client category matches. One CNPJ can match multiple categories.
CREATE TABLE IF NOT EXISTS estabelecimentos_categorias (
    cnpj VARCHAR(14) NOT NULL REFERENCES estabelecimentos_crm(cnpj) ON DELETE CASCADE,
    categoria TEXT NOT NULL,
    categoria_principal BOOLEAN NOT NULL DEFAULT FALSE,
    cnae_match VARCHAR(7),
    cnae_origem TEXT NOT NULL CHECK (cnae_origem IN ('principal', 'secundaria', 'governo')),
    match_type TEXT NOT NULL,
    prioridade INTEGER NOT NULL,
    classificacao_fonte TEXT NOT NULL DEFAULT 'Receita CNAE',
    PRIMARY KEY (cnpj, categoria, cnae_match, cnae_origem)
);

-- 5. Public-data lookup for provided active-client examples, independent from lead filters.
-- Existing local databases that used the pre-cleanup public identifier are migrated in place.
DO $$
BEGIN
    IF to_regclass('public.active_clients_public_enrichment') IS NULL
       AND to_regclass('public.clientes_ativos_governo') IS NOT NULL
       AND EXISTS (
           SELECT 1
           FROM pg_class c
           JOIN pg_namespace n ON n.oid = c.relnamespace
           WHERE n.nspname = 'public'
             AND c.relname = 'clientes_ativos_governo'
             AND c.relkind = 'r'
       ) THEN
        ALTER TABLE clientes_ativos_governo RENAME TO active_clients_public_enrichment;
    END IF;
END $$;

-- This persists CNAE data for the provided CNPJ list so the classifier can enrich it later.
CREATE TABLE IF NOT EXISTS active_clients_public_enrichment (
    cnpj VARCHAR(14) PRIMARY KEY,
    cnpj_basico VARCHAR(8),
    razao_social TEXT,
    natureza_juridica VARCHAR(4),
    capital_social TEXT,
    porte_empresa VARCHAR(10),
    nome_fantasia TEXT,
    situacao_cadastral VARCHAR(2),
    data_inicio_atividade VARCHAR(8),
    cnae_fiscal_principal VARCHAR(7),
    cnae_fiscal_secundaria TEXT,
    uf CHAR(2),
    municipio VARCHAR(4),
    atualizado_em TIMESTAMP DEFAULT NOW()
);

-- Backward-compatible alias for scripts or ad-hoc queries that still use the old public identifier.
CREATE OR REPLACE VIEW clientes_ativos_governo AS
SELECT * FROM active_clients_public_enrichment;

-- 6. Import log
CREATE TABLE IF NOT EXISTS import_log (
    id SERIAL PRIMARY KEY,
    fase TEXT,
    arquivo TEXT,
    linhas_importadas INT,
    data_import TIMESTAMP DEFAULT NOW()
);

-- 7. Main indexes
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_data ON estabelecimentos_crm(data_inicio_atividade);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_uf ON estabelecimentos_crm(uf);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_municipio ON estabelecimentos_crm(municipio);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_cnae ON estabelecimentos_crm(cnae_fiscal_principal);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_cnpj_basico ON estabelecimentos_crm((LEFT(cnpj, 8)));
CREATE INDEX IF NOT EXISTS idx_empresas_cnpj_basico ON empresas_dados(cnpj_basico);
CREATE INDEX IF NOT EXISTS idx_import_log_data ON import_log(data_import);
CREATE INDEX IF NOT EXISTS idx_cnae_categoria_map_active ON cnae_categoria_map(ativo, match_type, cnae_inicio, cnae_fim);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_categorias_cnpj ON estabelecimentos_categorias(cnpj);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_categorias_categoria ON estabelecimentos_categorias(categoria);
CREATE INDEX IF NOT EXISTS idx_active_clients_public_enrichment_cnpj_basico ON active_clients_public_enrichment(cnpj_basico);
CREATE INDEX IF NOT EXISTS idx_active_clients_public_enrichment_cnae ON active_clients_public_enrichment(cnae_fiscal_principal);

-- 8. Comments
COMMENT ON TABLE estabelecimentos_crm IS 'Qualified CNPJ leads for CRM - PR/SC region';
COMMENT ON TABLE empresas_dados IS 'Company legal name data used for joins';
COMMENT ON TABLE cnae_categoria_map IS 'CNAE-to-business-category rules based only on Receita/government CNAE data';
COMMENT ON TABLE estabelecimentos_categorias IS 'Business category matches generated from establishment CNAE fields';
COMMENT ON TABLE active_clients_public_enrichment IS 'Public CNAE/company lookup for provided active-client examples, captured during import before cleanup';
COMMENT ON VIEW clientes_ativos_governo IS 'Compatibility view for active_clients_public_enrichment';
COMMENT ON TABLE import_log IS 'Execution log for pipeline runs';
