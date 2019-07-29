using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Threading.Tasks;
using ART3WebAPI.Models.Domains;

namespace ART3WebAPI.Controllers
{
    public class PostDataController : ApiController
    {
        #region [ Relatório Produto Flag ]

        [HttpPost]
        public async Task<HttpResponseMessage> RelatorioProdutoFlagPost(string paginaId, string usuario, string codFabricante, string codProduto, short flag)
        {
            try
            {
                if ((usuario ?? "").Length == 0 || (codFabricante ?? "").Length == 0 || (codProduto ?? "").Length == 0)
                    throw new HttpResponseException(HttpStatusCode.PartialContent);

                await RelatorioProdutoFlagBD.RelatorioProdutoFlagPostAsync(paginaId, usuario, codFabricante, codProduto, flag);
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
            }

            return Request.CreateResponse(HttpStatusCode.OK);
        }

        #endregion
    }
}
