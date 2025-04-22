using apiNegociacaoDividas.Models;
using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("[controller]")]
public class DividasController : ControllerBase
{
    private readonly IDividasService _dividasService;

    public DividasController(IDividasService dividasService)
    {
        _dividasService = dividasService;
    }

    [HttpGet(Name = "GetDividas")]
    public async Task<ActionResult<IEnumerable<Dividas>>> Get([FromQuery] string cpf)
    {
        try
        {
            var dividas = await _dividasService.ObterDividasPorCpfAsync(cpf);

            if (!dividas.Any())
            {
                return NoContent();
            }

            return Ok(dividas);
        }
        catch (ArgumentException ex)
        {
            return BadRequest(ex.Message);
        }
        catch (Exception ex)
        {
            return StatusCode(500, "Erro interno do servidor.");
        }
    }
}