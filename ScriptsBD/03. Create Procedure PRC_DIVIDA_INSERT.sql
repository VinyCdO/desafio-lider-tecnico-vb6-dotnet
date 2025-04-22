CREATE OR REPLACE PROCEDURE prcdividainsert(
    cpf VARCHAR(11),
    valor NUMERIC(10,2),
    vencimento DATE
)
LANGUAGE plpgsql
AS $$
BEGIN
    -- Validar CPF (formato básico)
    IF LENGTH(cpf) != 11 OR cpf ~ '[^0-9]' THEN
        RAISE EXCEPTION 'CPF inválido. Deve conter exatamente 11 dígitos numéricos.';
    END IF;
    
    -- Validar valor positivo
    IF valor <= 0 THEN
        RAISE EXCEPTION 'O valor original deve ser positivo.';
    END IF;
       
	-- Inserir os dados
    INSERT INTO dividas (cpf, valor_original, data_vencimento)
    VALUES (cpf, valor, vencimento);
    
    
    RAISE NOTICE 'Dívida inserida com sucesso para o CPF: %', cpf;
EXCEPTION
    WHEN OTHERS THEN
        ROLLBACK;
        RAISE EXCEPTION 'Erro ao inserir dívida: %', SQLERRM;
END;
$$;
