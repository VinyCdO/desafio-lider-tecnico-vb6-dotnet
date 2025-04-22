using apiNegociacaoDividas.Models;

public interface IDividasRepository
{
    Task<IEnumerable<Dividas>> GetDividasByCpfAsync(string cpf);
}
