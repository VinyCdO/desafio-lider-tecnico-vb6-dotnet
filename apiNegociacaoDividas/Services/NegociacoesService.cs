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

    public async Task<int> CriarNegociacaoAsync(NegociacaoCreateDto negociacaoDto)
    {
        if (negociacaoDto.IdDivida <= 0)
            throw new ArgumentException("O ID da dívida é obrigatório.");

        var negociacao = new Negociacoes
        {
            id_divida = negociacaoDto.IdDivida,
            qtd_parcelas = negociacaoDto.QtdParcelas,
            taxa_juros = negociacaoDto.TaxaJuros,
            valor_total = negociacaoDto.ValorTotal,
            data_negociacao = negociacaoDto.DataNegociacao
        };

        return await _negociacoesRepository.AddNegociacaoAsync(negociacao);
    }
}