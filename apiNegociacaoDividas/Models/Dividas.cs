namespace apiNegociacaoDividas.Models
{
    public class Dividas
    {
        public int id { get; set; }

        public string? cpf { get; set; }

        public decimal valor_original { get; set; }

        public DateOnly data_vencimento { get; set; }               
    }
}
