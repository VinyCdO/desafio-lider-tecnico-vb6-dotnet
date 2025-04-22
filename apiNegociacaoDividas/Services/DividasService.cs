using apiNegociacaoDividas.Models;

public class DividasService : IDividasService
{
    private readonly IDividasRepository _dividasRepository;

    public DividasService(IDividasRepository dividasRepository)
    {
        _dividasRepository = dividasRepository;
    }

    public async Task<IEnumerable<Dividas>> ObterDividasPorCpfAsync(string cpf)
    {
        if (string.IsNullOrWhiteSpace(cpf))
        {
            throw new ArgumentException("O CPF é obrigatório.");
        }

        return await _dividasRepository.GetDividasByCpfAsync(cpf);
    }
}
