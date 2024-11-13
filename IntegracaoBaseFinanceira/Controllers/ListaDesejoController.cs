using Dapper;
using IntegracaoBaseFinanceira.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Data;
using System.Data.SqlClient;

namespace IntegracaoBaseFinanceira.Controllers
{
    [ApiController]
    [Route("/api/financeiro/listaDesejo")]
    public class ListaDesejoController : ControllerBase
    {
        private readonly IDbConnection _con;
        private readonly ILogger<ListaDesejoController> _log;

        public ListaDesejoController(IDbConnection con, ILogger<ListaDesejoController> log)
        {
            _con = con;
            _log = log;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<ListaDesejo>>> GetAllProducts()
        {
            try
            {
                var sql = "SELECT * FROM LISTA_DESEJO";

                var products = await _con.QueryAsync(sql);

                _log.LogInformation("CONSULTA REALIZADA COM SUCESSO");
                return Ok(products);
            }
            catch (Exception ex)
            {
                _log.LogError("ERRO NA CONSULTA: " + ex.Message);
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost]
        public async Task<ActionResult> CreateProduct(ListaDesejo listaDesejo)
        {
            try
            {
                var sql = "INSERT INTO LISTA_DESEJO (ID, PRODUTO, PRECO, LINKACESSO) VALUES (@ID, @PRODUTO, @PRECO, @LINKACESSO)";

                Guid UUID = Guid.NewGuid();

                await _con.ExecuteAsync(sql, new
                    {
                        ID = UUID.ToString(),
                        PRODUTO = listaDesejo.Produto,
                        PRECO = listaDesejo.Preco,
                        LINKACESSO = listaDesejo.LinkAcesso
                    }
                );

                _log.LogInformation("CRIADO PRODUTO NA LISTA DE DESEJO: " + listaDesejo.Produto);
                return Ok(listaDesejo);
            }
            catch(Exception ex)
            {
                _log.LogError("ERRO NA CRIAÇÃO: " + ex.Message);
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPut("{id}")]
        public async Task<ActionResult> UpdateProduct(string id, ListaDesejo listaDesejo)
        {
            try
            {
                var sql = "UPDATE LISTA_DESEJO SET PRODUTO = @PRODUTO, PRECO = @PRECO, LINKACESSO = @LINKACESSO WHERE ID = @ID";

                await _con.ExecuteAsync(sql, new
                    {
                        ID = id,
                        PRODUTO = listaDesejo.Produto,
                        PRECO = listaDesejo.Preco,
                        LINKACESSO = listaDesejo.LinkAcesso
                    }
                );

                _log.LogInformation("LISTA DE DESEJO ATUALIZADA COM SUCESSO");
                return Ok(listaDesejo);
            }
            catch (Exception ex)
            {
                _log.LogError("ERRO NA ATUALIZAÇÃO: " + ex.Message);
                return StatusCode(500, ex.Message);
            }
        }

        [HttpDelete("{id}")]
        public async Task<ActionResult> DeleteProduct(string id)
        {
            try
            {
                var sql = "DELETE FROM LISTA_DESEJO WHERE ID = @ID";

                await _con.ExecuteAsync(sql, new
                    {
                        ID = id,
                    }
                );

                _log.LogInformation("REMOVIDO O ITEM DA LISTA DE DESEJO: " + id);
                return Ok("REMOVIDO O ITEM DA LISTA DE DESEJO: " + id);
            }
            catch (Exception ex)
            {
                _log.LogError("ERRO AO REMOVER: " + ex.Message);
                return StatusCode(500, ex.Message);
            }
        }
    }
}
