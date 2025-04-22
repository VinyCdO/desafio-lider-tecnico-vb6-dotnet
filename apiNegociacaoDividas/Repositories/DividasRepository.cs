using apiNegociacaoDividas.Models;
using Npgsql;

public class DividasRepository : IDividasRepository
{
    private readonly string _connectionString;

    public DividasRepository(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection");
    }

    public async Task<IEnumerable<Dividas>> GetDividasByCpfAsync(string cpf)
    {
        var dividas = new List<Dividas>();

        await using var connection = new NpgsqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = "SELECT id, cpf, valor_original, data_vencimento FROM dividas WHERE cpf = @cpf";

        await using var command = new NpgsqlCommand(query, connection);
        command.Parameters.AddWithValue("cpf", cpf);

        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            dividas.Add(new Dividas
            {
                id = reader.GetInt32(0),
                cpf = reader.GetString(1),
                valor_original = reader.GetDecimal(2),
                data_vencimento = reader.GetFieldValue<DateOnly>(3)
            });
        }

        return dividas;
    }
}