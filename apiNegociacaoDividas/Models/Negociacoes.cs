namespace apiNegociacaoDividas.Models
{
    public class Negociacoes
    {
        public int id { get; set; }
        public int id_divida { get; set; }
        public int qtd_parcelas { get; set; }
        public decimal taxa_juros { get; set; }
        public decimal valor_total { get; set; }
        public DateOnly data_negociacao { get; set; }        
    }
}
