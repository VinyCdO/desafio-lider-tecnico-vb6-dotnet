using apiNegociacaoDividas.Models;

public interface INegociacoesRepository
{
    Task<IEnumerable<Negociacoes>> GetNegociacoesByIdDividaAsync(int id_divida);
    Task<int> AddNegociacaoAsync(Negociacoes negociacao);
}
