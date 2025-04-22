CREATE OR REPLACE FUNCTION fnCalculaJurosComposto (valorInicial numeric(12,2), qtdParcelas integer, taxaJuros numeric(5,2))
RETURNS numeric(12,2)
AS $$
DECLARE
    ValorCalculado NUMERIC(12,2) := valorInicial;
	CountLoop INT := 1;
BEGIN
	WHILE CountLoop <= qtdParcelas LOOP
		ValorCalculado := ValorCalculado + (ValorCalculado * (taxaJuros / 100));
		CountLoop := CountLoop + 1;
	END LOOP;
	
    RETURN ValorCalculado;
END;
$$ LANGUAGE plpgsql;