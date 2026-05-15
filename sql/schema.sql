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

-- 3. Import log
CREATE TABLE IF NOT EXISTS import_log (
    id SERIAL PRIMARY KEY,
    fase TEXT,
    arquivo TEXT,
    linhas_importadas INT,
    data_import TIMESTAMP DEFAULT NOW()
);

-- 4. Main indexes
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_data ON estabelecimentos_crm(data_inicio_atividade);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_uf ON estabelecimentos_crm(uf);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_municipio ON estabelecimentos_crm(municipio);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_cnae ON estabelecimentos_crm(cnae_fiscal_principal);
CREATE INDEX IF NOT EXISTS idx_estabelecimentos_cnpj_basico ON estabelecimentos_crm((LEFT(cnpj, 8)));
CREATE INDEX IF NOT EXISTS idx_empresas_cnpj_basico ON empresas_dados(cnpj_basico);
CREATE INDEX IF NOT EXISTS idx_import_log_data ON import_log(data_import);

-- 5. Comments
COMMENT ON TABLE estabelecimentos_crm IS 'Filtered CNPJ establishments for downstream export';
COMMENT ON TABLE empresas_dados IS 'Company legal name data used for joins';
COMMENT ON TABLE import_log IS 'Execution log for pipeline runs';
