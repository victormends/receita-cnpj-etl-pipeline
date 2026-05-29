-- PostgreSQL staging tables for enriched existing-client imports.

CREATE TABLE IF NOT EXISTS clientes_staging (
    id_normalizado text,
    cnpj varchar(14),
    nome text,
    razao_social text,
    cidade text,
    estado text,
    regime_tributario text,
    postgres_versao text,
    arquitetura_os text,
    provedor text,
    cpf_cnpj_internet text,
    nome_internet text,
    atualizando text,
    flag_postgres_17_ou_nuvem boolean,
    flag_postgres_desatualizado_32_bits boolean,
    flag_postgres_atualizado_sem_replicacao boolean,
    flag_multiplos_computadores boolean,
    flag_sistema_web boolean,
    flag_sistema_deluito_contador boolean,
    cnae_fiscal_principal text,
    cnae_fiscal_secundaria text,
    categoria_principal text,
    categorias_detectadas text,
    categoria_transporte boolean,
    categoria_gas_combustivel boolean,
    categoria_farmacia boolean,
    categoria_construtora boolean,
    categoria_industria boolean,
    categoria_agro_produtor_rural boolean,
    categoria_telecom boolean,
    categoria_ti_software boolean,
    categoria_saude_clinica boolean,
    categoria_financeiro boolean,
    categoria_cooperativa boolean,
    categoria_imobiliario boolean,
    categoria_grafica boolean,
    categoria_frigorifico boolean,
    categoria_importador_exportador_status text,
    categoria_beneficio_fiscal_status text
);

CREATE INDEX IF NOT EXISTS idx_clientes_staging_id_normalizado ON clientes_staging(id_normalizado);
CREATE INDEX IF NOT EXISTS idx_clientes_staging_cnpj ON clientes_staging(cnpj);
CREATE INDEX IF NOT EXISTS idx_clientes_staging_categoria_principal ON clientes_staging(categoria_principal);

CREATE TABLE IF NOT EXISTS clientes_categorias (
    id_normalizado text,
    cnpj varchar(14),
    categoria text NOT NULL,
    categoria_principal boolean NOT NULL DEFAULT false,
    cnae_match text,
    cnae_origem text,
    match_type text,
    prioridade integer,
    classificacao_fonte text
);

CREATE INDEX IF NOT EXISTS idx_clientes_categorias_id_normalizado ON clientes_categorias(id_normalizado);
CREATE INDEX IF NOT EXISTS idx_clientes_categorias_cnpj ON clientes_categorias(cnpj);
CREATE INDEX IF NOT EXISTS idx_clientes_categorias_categoria ON clientes_categorias(categoria);
