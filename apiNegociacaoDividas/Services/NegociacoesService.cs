using apiNegociacaoDividas.Models;

public class NegociacoesService : INegociacoesService
{
    private readonly INegociacoesRepository _negociacoesRepository;
    public NegociacoesService(INegociacoesRepository negociacoesRepository)
    {
        _negociacoesRepository = negociacoesRepository;
    }
    public async Task<IEnumerable<Negociacoes>> ObterNegociacoesPorIdDividaAsync(int idDivida)
    {
        if (idDivida <= 0)
        {
            throw new ArgumentException("O ID da dívida é obrigatório.");
        }
        
        return await _negociacoesRepository.GetNegociacoesByIdDividaAsync(idDivida);
    }
}