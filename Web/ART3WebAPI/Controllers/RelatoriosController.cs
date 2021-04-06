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
    public class RelatoriosController : ApiController
    {
        #region [CadIndicadoresListagemCSV]
        #region [GetCadIndicadoresListagemCSV]
        [HttpPost]
        public async Task<HttpResponseMessage> GetCadIndicadoresListagemCSV(string loja, string usuario)
        {
            DataCadIndicadores datasource = new DataCadIndicadores();
            List<Indicador> indicadorList = datasource.GetIndicador(loja).ToList();

            DateTime data = DateTime.Now;
            string fileName = "CadIndicadoresListagemCSV_" + data.ToString("yyyyMMdd_HHmmss");
            fileName = fileName + ".csv";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);

            Log s_log = new Log();
            s_log.complemento = "Nome do arquivo: " + fileName;
            s_log.operacao = "Relatório CadIndicadoresListagemCSV";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "";

            HttpResponseMessage result = null;

            StringBuilder xmlResponse = new StringBuilder();

            try
            {
                await ART3WebAPI.Models.Domains.CadIndicadoresGeradorRelatorio.GerarListagemCSV(indicadorList, filePath);

                statusResponse = "OK";

                LogDAO.insere(usuario, s_log, out strMsgErro);

            }
            catch (Exception e)
            {
                statusResponse = "Falha";
                MsgErroException = e.Message;
            }

            xmlResponse.Append("{ 'fileName' : '" + fileName + "', " + "'Status' : '" + statusResponse + "', " + "'Exception' : '" + MsgErroException + "' }");
            xmlResponse = xmlResponse.Replace("'", ((char)34).ToString());

            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StringContent(xmlResponse.ToString(), Encoding.UTF8, "text/html");

            return result;
        }
        #endregion

        #region [downloadCadIndicadoresListagemCSV]
        [HttpPost]
        public HttpResponseMessage downloadCadIndicadoresListagemCSV(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }
        #endregion
        #endregion

        #region [ EstatisticasOcorrencia ]
        [HttpPost]
        public async Task<HttpResponseMessage> GetXLSReport(string usuario, string dt_inicio, string dt_termino, string motivo_ocorrencia, string tp_ocorrencia, string transportadora, string vendedor, string indicador, string UF, string loja)
        {

            if (string.IsNullOrEmpty(dt_inicio.ToString())) throw new Exception("Não foi informada a data inicial do período de vendas.");
            if (string.IsNullOrEmpty(dt_termino.ToString())) throw new Exception("Não foi informada a data final do período de vendas.");



            DateTime data = DateTime.Now;
            string fileName = "EstatisticasOcorrencias_" + data.ToString("yyyyMMdd_HHmmss");
            fileName = fileName + ".xlsx";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);
            StringBuilder xmlResponse = new StringBuilder();
            Log s_log = new Log();
            s_log.complemento = " Nome do arquivo: " + fileName;
            s_log.operacao = "Relatório EstatisticasOcorrencias";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "";
            HttpResponseMessage result = null;

            try
            {

                DataEstatisticasOcorrencias datasource = new DataEstatisticasOcorrencias();
                List<Ocorrencias> relEstOcorrenciasList = datasource.Get(dt_inicio, dt_termino, motivo_ocorrencia, tp_ocorrencia, transportadora, vendedor, indicador, UF, loja).ToList();
                if (relEstOcorrenciasList.Count != 0)
                {
                    await ART3WebAPI.Models.Domains.EstOcorrenciasGeradorRelatorio.GenerateXLS(relEstOcorrenciasList, filePath, dt_inicio, dt_termino, motivo_ocorrencia, tp_ocorrencia, transportadora, vendedor, indicador, UF, loja);
                    statusResponse = "OK";
                    LogDAO.insere(usuario, s_log, out strMsgErro);
                }
                else
                {
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

            return result;
        }

        [HttpPost]
        public HttpResponseMessage downloadXLS(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }
        #endregion


        #region [ OcorrenciaStatus ]
        [HttpPost]
        public async Task<HttpResponseMessage> GetXLSReport2(string usuario, string oc_status, string transportadora, string loja)
        {


            DateTime data = DateTime.Now;
            string fileName = "Ocorrencias_" + data.ToString("yyyyMMdd_HHmmss");
            fileName = fileName + ".xlsx";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);
            StringBuilder xmlResponse = new StringBuilder();
            Log s_log = new Log();
            s_log.complemento = " Nome do arquivo: " + fileName;
            s_log.operacao = "Relatório Ocorrencias";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "";
            HttpResponseMessage result = null;

            try
            {

                DataOcorrencias datasource = new DataOcorrencias();
                List<OcorrenciasStatus> relOcorrenciasList = datasource.Get(oc_status, transportadora, loja).ToList();
                if (relOcorrenciasList.Count != 0)
                {
                    await ART3WebAPI.Models.Domains.OcorrenciasGeradorRelatorio.GenerateXLS(relOcorrenciasList, filePath, oc_status, loja, transportadora);
                    statusResponse = "OK";
                    LogDAO.insere(usuario, s_log, out strMsgErro);
                }
                else
                {
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

            return result;
        }

        [HttpPost]
        public HttpResponseMessage downloadXLS2(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }
        #endregion

        #region [ Compras 2 ]

        #region [ Get CSV COMPRAS ]
        [HttpPost]
        public async Task<HttpResponseMessage> GetCompras2CSV(string usuario, string tipo_periodo, string dt_inicio, string dt_termino, string fabricante, string produto, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string nf, string dt_nf_inicio, string dt_nf_termino, string visao, string detalhamento)
        {


            DateTime data = DateTime.Now;
            string fileName = "Compras2_" + data.ToString("yyyyMMdd_HHmmss");
            fileName = fileName + ".xlsx";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);
            StringBuilder xmlResponse = new StringBuilder();
            Log s_log = new Log();
            s_log.complemento = " Nome do arquivo: " + fileName;
            s_log.operacao = "Relatório Compras2";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "";
            HttpResponseMessage result = null;

			try
			{
				#region [ Consistências ]
				if ((tipo_periodo ?? "").Length == 0)
				{
					statusResponse = "Falha";
					MsgErroException = "Não foi informado o tipo de período da consulta!";
				}
				else if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
				{
					if ((dt_inicio ?? "").Length == 0)
					{
						statusResponse = "Falha";
						MsgErroException = "Data de início do período da consulta não foi informado!";
					}

					if ((dt_termino ?? "").Length == 0)
					{
						statusResponse = "Falha";
						MsgErroException = "Data de término do período da consulta não foi informado!";
					}
				}
				else if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_EMISSAO_NF_ENTRADA))
				{
					if ((dt_nf_inicio ?? "").Length == 0)
					{
						statusResponse = "Falha";
						MsgErroException = "Data de início do período da consulta não foi informado!";
					}

					if ((dt_nf_termino ?? "").Length == 0)
					{
						statusResponse = "Falha";
						MsgErroException = "Data de término do período da consulta não foi informado!";
					}
				}
				else
				{
					statusResponse = "Falha";
					MsgErroException = "Tipo de período da consulta informado é inválido (" + tipo_periodo + ")!";
				}
				#endregion

				if (MsgErroException.Length == 0)
				{
					Global.setDefaultBD(usuario, "RelCompras2Filtro|rb_periodo", string.IsNullOrEmpty(tipo_periodo) ? "" : tipo_periodo);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_dt_periodo_inicio", string.IsNullOrEmpty(dt_inicio) ? "" : dt_inicio);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_dt_periodo_termino", string.IsNullOrEmpty(dt_termino) ? "" : dt_termino);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_fabricante", string.IsNullOrEmpty(fabricante) ? "" : fabricante.Replace("_", ", "));
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_grupo", string.IsNullOrEmpty(grupo) ? "" : grupo.Replace("_", ", "));
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_subgrupo", string.IsNullOrEmpty(subgrupo) ? "" : subgrupo.Replace("_", ", "));
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_potencia_BTU", string.IsNullOrEmpty(btu) ? "" : btu);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_ciclo", string.IsNullOrEmpty(ciclo) ? "" : ciclo);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_posicao_mercado", string.IsNullOrEmpty(pos_mercado) ? "" : pos_mercado);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_dt_nf_inicio", string.IsNullOrEmpty(dt_nf_inicio) ? "" : dt_nf_inicio);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|c_dt_nf_termino", string.IsNullOrEmpty(dt_nf_termino) ? "" : dt_nf_termino);
					Global.setDefaultBD(usuario, "RelCompras2Filtro|rb_detalhe", detalhamento);

					DataCompras2 datasource = new DataCompras2();
					List<Compras> relCompras2List = datasource.Get(tipo_periodo, dt_inicio, dt_termino, fabricante, produto, grupo, subgrupo, btu, ciclo, pos_mercado, nf, dt_nf_inicio, dt_nf_termino, visao, detalhamento).ToList();
					if (relCompras2List.Count != 0)
					{
						await ART3WebAPI.Models.Domains.Compras2GeradorRelatorio.GenerateXLS(relCompras2List, filePath, tipo_periodo, dt_inicio, dt_termino, fabricante, produto, grupo, subgrupo, btu, ciclo, pos_mercado, nf, dt_nf_inicio, dt_nf_termino, visao, detalhamento);
						statusResponse = "OK";
						LogDAO.insere(usuario, s_log, out strMsgErro);
					}
					else
					{
						statusResponse = "Vazio";
						MsgErroException = "Nenhum registro foi encontrado!";
					}
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

            return result;
        }
        
        #endregion

        #region [ Download ]
        [HttpPost]
        public HttpResponseMessage downloadCompras2CSV(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }
        #endregion

        #endregion

        #region [ Devolução de Produtos 2 ]
        [HttpPost]
        public async Task<HttpResponseMessage> GeraDevolucaoProdutos2XLS(string usuario, string dt_devolucao_inicio, string dt_devolucao_termino, string fabricante, string produto, string pedido, string vendedor, string indicador, string captador, string lojas)
        {
            #region [ Declarações ]
            DateTime dataAtual;
            string fileName;
            string filePath;
            StringBuilder sbXmlResponse = new StringBuilder();
            Log s_log = new Log();
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "";
            HttpResponseMessage result = null;
            List<DevolucaoProduto2Entity> DevolProd2Lista;
            DataDevolucaoProdutos2 datasource;
            #endregion

            #region [ Gera nome do arquivo ]
            dataAtual = DateTime.Now;
            fileName = "DevolucaoProdutos2_" + dataAtual.ToString("yyyyMMdd_HHmmss");
            fileName = fileName + ".xlsx";
            filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);
            #endregion

            s_log.complemento = " Nome do arquivo: " + fileName;
            s_log.operacao = "Rel Devol Prod II";
            s_log.usuario = usuario;

            try
            {
                datasource = new DataDevolucaoProdutos2();
                DevolProd2Lista = datasource.Get(usuario, dt_devolucao_inicio, dt_devolucao_termino, fabricante, produto, pedido, vendedor, indicador, captador, lojas);
                if (DevolProd2Lista.Count > 0)
                {
                    await DevolucaoProdutos2GeradorRelatorio.GeraXLS(DevolProd2Lista, filePath, dt_devolucao_inicio, dt_devolucao_termino, fabricante, produto, pedido, vendedor, indicador, captador, lojas);
                    statusResponse = "OK";
                    LogDAO.insere(usuario, s_log, out strMsgErro);
                }
                else
                {
                    statusResponse = "Vazio";
                    MsgErroException = "Nenhum registro foi encontrado!";
                }
            }
            catch (Exception ex)
            {
                statusResponse = "Falha";
                MsgErroException = ex.Message;
            }

            sbXmlResponse.Append("{ \"fileName\" : \"" + fileName + "\", " + "\"Status\" : \"" + statusResponse + "\", " + "\"Exception\" : " + System.Web.Helpers.Json.Encode(MsgErroException) + "}");
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StringContent(sbXmlResponse.ToString(), Encoding.UTF8, "application/json");

            return result;
        }

        #region [ Download ]
        [HttpPost]
        public HttpResponseMessage downloadDevolucaoProdutos2XLS(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }
        #endregion
        #endregion

    }
}
