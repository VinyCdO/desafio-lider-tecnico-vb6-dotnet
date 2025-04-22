using apiNegociacaoDividas.Models;

public interface IDividasService
{
    Task<IEnumerable<Dividas>> ObterDividasPorCpfAsync(string cpf);
}
