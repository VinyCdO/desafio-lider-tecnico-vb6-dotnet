public class NegociacaoCreateDto
{
    public int IdDivida { get; set; }
    public int QtdParcelas { get; set; }
    public decimal TaxaJuros { get; set; }
    public decimal ValorTotal { get; set; }
    public DateOnly DataNegociacao { get; set; }
}
