using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Threading.Tasks;
using System.Web;
using System.IO;
using System.Net.Http.Headers;
using System;
using ART3WebAPI.Models.Repository;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;
using System.Text;

namespace ART3WebAPI.Controllers
{
    public class FarolController : ApiController
    {
        [HttpPost]
        public async Task<HttpResponseMessage> GetXLSReport(string usuario, string dt_inicio, string dt_termino, string fabricante, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string perc_est_cresc, string loja, string visao)
        {
			const string NOME_DESTA_ROTINA = "FarolController.GetXLSReport()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", dt_inicio=" + (dt_inicio ?? "") + ", dt_termino=" + (dt_termino ?? "") + ", fabricante=" + (fabricante ?? "") + ", grupo=" + (grupo ?? "") + ", subgrupo=" + (subgrupo ?? "") + ", btu=" + (btu ?? "") + ", ciclo=" + (ciclo ?? "") + ", pos_mercado=" + (pos_mercado ?? "") + ", perc_est_cresc=" + (perc_est_cresc ?? "") + ", loja=" + (loja ?? "") + ", visao=" + (visao ?? "") + ")";
			Global.gravaLogAtividade(httpRequestId, msg);

			if (string.IsNullOrEmpty(dt_inicio.ToString())) throw new Exception("Não foi informada a data inicial do período de vendas.");
            if (string.IsNullOrEmpty(dt_termino.ToString())) throw new Exception("Não foi informada a data final do período de vendas.");

            HttpResponseMessage result = null;

            DateTime data = DateTime.Now;
            string fileName = "Farol_" + data.ToString("yyyyMMdd_HHmmss");
            fileName = fileName + ".xlsx";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/FarolResumido/" + fileName);
            StringBuilder xmlResponse = new StringBuilder();
            Log s_log = new Log();
            s_log.complemento = "Farol Resumido Nome do arquivo: " + fileName;
            s_log.operacao = "Farol Resumido";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "";

            try
            {
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_dt_periodo_inicio", dt_inicio);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_dt_periodo_termino", dt_termino);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_perc_est_cresc", string.IsNullOrEmpty(perc_est_cresc) ? "" : perc_est_cresc);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_fabricante", string.IsNullOrEmpty(fabricante) ? "" : fabricante.Replace("_", ", "));
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_grupo", string.IsNullOrEmpty(grupo) ? "" : grupo.Replace("|", ", "));
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_subgrupo", string.IsNullOrEmpty(subgrupo) ? "" : subgrupo.Replace("|", ", "));
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_potencia_BTU", string.IsNullOrEmpty(btu) ? "" : btu);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_ciclo", string.IsNullOrEmpty(ciclo) ? "" : ciclo);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_posicao_mercado", string.IsNullOrEmpty(pos_mercado) ? "" : pos_mercado);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_loja", string.IsNullOrEmpty(loja) ? "" : loja);
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|rb_visao", visao);

                DataFarol datasource = new DataFarol();
                List<Farol> relFarolList = datasource.Get(dt_inicio, dt_termino, fabricante, grupo, subgrupo, btu, ciclo, pos_mercado, loja).ToList();

                if (relFarolList.Count != 0)
                {
                    await ART3WebAPI.Models.Domains.FarolGeradorRelatorio.GenerateXLS(relFarolList, filePath, dt_inicio, dt_termino, fabricante, grupo, subgrupo, btu, ciclo, pos_mercado, perc_est_cresc, loja, visao);
                    statusResponse = "OK";

                    LogDAO.insere(httpRequestId, usuario, s_log, out strMsgErro);
                }
                else {
                    statusResponse = "Vazio";
                    MsgErroException = "Nenhum registro foi encontrado!";
                }

            }
            catch (Exception e)
            {
                statusResponse = "Falha";
                MsgErroException = e.Message;
            }

			xmlResponse.Append("{ \"fileName\" : \"" + fileName + "\", " + "\"Status\" : \"" + statusResponse + "\", " + "\"Exception\" : " + System.Web.Helpers.Json.Encode(MsgErroException) + "}");
            result = Request.CreateResponse(HttpStatusCode.OK);
			result.Content = new StringContent(xmlResponse.ToString(), Encoding.UTF8, "application/json");

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
        }

        [HttpPost]
        public HttpResponseMessage downloadXLS(string fileName)
        {
			const string NOME_DESTA_ROTINA = "FarolController.downloadXLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

			string filePath = HttpContext.Current.Server.MapPath("~/Report/FarolResumido/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }
    }
}
