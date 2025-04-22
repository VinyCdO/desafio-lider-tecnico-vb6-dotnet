using apiNegociacaoDividas.Models;
using Npgsql;

public class NegociacoesRepository : INegociacoesRepository
{
    private readonly string _connectionString;
    public NegociacoesRepository(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection");
    }

    public async Task<IEnumerable<Negociacoes>> GetNegociacoesByIdDividaAsync(int id_divida)
    {
        var negociacoes = new List<Negociacoes>();

        await using var connection = new NpgsqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = "SELECT id, id_divida, valor_total, data_negociacao FROM negociacoes WHERE id_divida = @id_divida";
        
        await using var command = new NpgsqlCommand(query, connection);
        command.Parameters.AddWithValue("id_divida", id_divida);
        
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            negociacoes.Add(new Negociacoes
            {
                id = reader.GetInt32(0),
                id_divida = reader.GetInt32(1),
                valor_total = reader.GetDecimal(2),
                data_negociacao = reader.GetFieldValue<DateOnly>(3)
            });
        }

        return negociacoes;
    }
}

