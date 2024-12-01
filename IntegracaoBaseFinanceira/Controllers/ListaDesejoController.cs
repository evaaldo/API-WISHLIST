using Dapper;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using IntegracaoBaseFinanceira.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using System;
using System.Data;
using System.Data.SqlClient;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace IntegracaoBaseFinanceira.Controllers
{
    [ApiController]
    [Route("/api/financeiro/listaDesejo")]
    public class ListaDesejoController : ControllerBase
    {
        private readonly IDbConnection _con;
        private readonly ILogger<ListaDesejoController> _log;
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string AplicationName = "ListaDesejo";
        static readonly string SpreadsheetId = "15Af0Utm_QHiC0OygF3fYxlNsDc0NUJ4KsOtO3tdxVXQ";
        static readonly string sheet = "LISTA DE DESEJO";
        static SheetsService service;

        public ListaDesejoController(IDbConnection con, ILogger<ListaDesejoController> log)
        {
            _con = con;
            _log = log;
        }
        public static void ConnectToGoogle()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = AplicationName,
            });
        }

        [HttpGet("/api/financeiro/database/listaDesejo")]
        public async Task<IEnumerable<ListaDesejoSheet>> GetDatabaseProducts()
        {
            try
            {
                var sql = "SELECT * FROM LISTA_DESEJO";

                var products = await _con.QueryAsync<ListaDesejoSheet>(sql);

                return products;
            }
            catch (Exception ex)
            {
                _log.LogError("ERRO NA CONSULTA: " + ex.Message);
                throw;
            }
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

        [HttpGet("/api/financeiro/listaDesejo/readSheetsProducts")]
        public async Task<ActionResult> ReadWishListSheet()
        {
            try
            {
                if (service == null)
                {
                    ConnectToGoogle();
                }

                var range = $"{sheet}!A1:C9";
                var request = service!.Spreadsheets.Values.Get(SpreadsheetId, range);

                var response = request!.Execute();
                var values = response!.Values;

                if (values != null && values.Count > 0)
                {
                    return Ok(values);
                }
                else
                {
                    _log.LogWarning("Nenhum dado encontrado na planilha.");
                    return StatusCode(204);
                }
            }
            catch (Exception ex)
            {
                _log.LogError("Erro ao ler entradas da planilha: " + ex.Message);
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost("/api/financeiro/listaDesejo/addProductToSheet")]
        public async Task<ActionResult> AddProductToSheet(ListaDesejoSheet listaDesejoSheet)
        {
            try
            {
                if(service == null)
                {
                    ConnectToGoogle();
                }

                string readRange = $"{sheet}!A1:C9";
                string appendRange = $"{sheet}!A1:C9";  

                var requestRead = service!.Spreadsheets.Values.Get(SpreadsheetId, readRange);

                var responseRead = requestRead!.Execute();
                var values = responseRead.Values ?? new List<IList<object>>();
                int nextRowIndex = values.Count + 1;

                var newRow = new List<object> { listaDesejoSheet.Produto, listaDesejoSheet.Preco, listaDesejoSheet.LinkAcesso };

                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { newRow }
                };

                var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, appendRange);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();

                return Ok($"Linha adicionada com sucesso na posição {nextRowIndex}");
            }
            catch (Exception ex)
            {
                _log.LogError("Erro ao atualizar planilha: " + ex.Message);
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost("/api/financeiro/listaDesejo/updateSheetWithDatabase")]
        public async Task<ActionResult> UpdateSheetsWithDatabase()
        {
            try
            {
                if(service == null)
                {
                    ConnectToGoogle();
                }

                string updateRange = $"{sheet}!A1:C9";
                var products = await this.GetDatabaseProducts();
                var productsList = new List<IList<object>>
                {
                    new object[] { "PRODUTO", "PREÇO", "LINK DE ACESSO" }
                };

                foreach (var product in products)
                {
                    productsList.Add(new object[]
                    {
                        product.Produto,
                        product.Preco,
                        product.LinkAcesso
                    });
                }

                var valueRange = new ValueRange
                {
                    Values = productsList
                };

                var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, updateRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var updateResponse = updateRequest.Execute();

                return Ok("PLANILHA ATUALIZADA DE ACORDO COM BANCO DE DADOS");
            }
            catch (Exception ex)
            {
                _log.LogError("ERRO NA ATUALIZAÇÃO: " + ex.Message);
                throw;
            }
        }

    }
}
