using apiNegociacaoDividas.Models;

public interface INegociacoesService
{
    Task<IEnumerable<Negociacoes>> ObterNegociacoesPorIdDividaAsync(int idDivida);
}
