CREATE TABLE dividas (
    id SERIAL PRIMARY KEY,
    cpf CHAR(11) NOT NULL,
    valor_original NUMERIC(12,2) NOT NULL,
    data_vencimento DATE NOT NULL
);

CREATE TABLE negociacoes (
	id SERIAL PRIMARY KEY, 
	id_divida SERIAL NOT NULL,
	qtd_parcelas INTEGER NOT NULL, 
	taxa_juros NUMERIC(5,2) NOT NULL, 
	valor_total NUMERIC(12,2) NOT NULL,
	data_negociacao DATE NOT NULL,
	CONSTRAINT fk_divida FOREIGN KEY (id_divida) REFERENCES dividas(id)
);
