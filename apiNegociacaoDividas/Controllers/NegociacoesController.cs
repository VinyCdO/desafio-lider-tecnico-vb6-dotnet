using apiNegociacaoDividas.Models;
using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("[controller]")]
public class NegociacoesController : ControllerBase
{
    private readonly INegociacoesService _negociacoesService;

    public NegociacoesController(INegociacoesService negociacoesService)
    {
        _negociacoesService = negociacoesService;
    }

    [HttpGet(Name = "GetNegociacoes")]
    public async Task<ActionResult<IEnumerable<Negociacoes>>> Get([FromQuery] int id_divida)
    {
        try
        {
            var negociacoes = await _negociacoesService.ObterNegociacoesPorIdDividaAsync(id_divida);

            if (!negociacoes.Any())
            {
                return NoContent();
            }

            return Ok(negociacoes);
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

    [HttpPost]
    public async Task<ActionResult<int>> Post([FromBody] NegociacaoCreateDto negociacaoDto)
    {
        try
        {
            var id = await _negociacoesService.CriarNegociacaoAsync(negociacaoDto);
            return CreatedAtAction(nameof(Get), new { id_divida = negociacaoDto.IdDivida }, id);
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
