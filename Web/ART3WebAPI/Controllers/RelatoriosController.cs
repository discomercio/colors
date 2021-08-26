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
			const string NOME_DESTA_ROTINA = "RelatoriosController.GetCadIndicadoresListagemCSV()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", loja=" + (loja ?? "") + ")";
			Global.gravaLogAtividade(httpRequestId, msg);

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

                LogDAO.insere(httpRequestId, usuario, s_log, out strMsgErro);

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

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
        }
        #endregion

        #region [downloadCadIndicadoresListagemCSV]
        [HttpPost]
        public HttpResponseMessage downloadCadIndicadoresListagemCSV(string fileName)
        {
			const string NOME_DESTA_ROTINA = "RelatoriosController.downloadCadIndicadoresListagemCSV()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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
			const string NOME_DESTA_ROTINA = "RelatoriosController.GetXLSReport()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", dt_inicio=" + (dt_inicio ?? "") + ", dt_termino=" + (dt_termino ?? "") + ", motivo_ocorrencia=" + (motivo_ocorrencia ?? "") + ", tp_ocorrencia=" + (tp_ocorrencia ?? "") + ", transportadora=" + (transportadora ?? "") + ", vendedor=" + (vendedor ?? "") + ", indicador=" + (indicador ?? "") + ", UF=" + (UF ?? "") + ", loja=" + (loja ?? "") + ")";
			Global.gravaLogAtividade(httpRequestId, msg);

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
                    LogDAO.insere(httpRequestId, usuario, s_log, out strMsgErro);
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

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
        }

        [HttpPost]
        public HttpResponseMessage downloadXLS(string fileName)
        {
			const string NOME_DESTA_ROTINA = "RelatoriosController.downloadXLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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
			const string NOME_DESTA_ROTINA = "RelatoriosController.GetXLSReport2()";
			Guid httpRequestId = Request.GetCorrelationId();
			DateTime data = DateTime.Now;
			string msg;
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
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", oc_status=" + (oc_status ?? "") + ", loja=" + (loja ?? "") + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				DataOcorrencias datasource = new DataOcorrencias();
                List<OcorrenciasStatus> relOcorrenciasList = datasource.Get(oc_status, transportadora, loja).ToList();
                if (relOcorrenciasList.Count != 0)
                {
                    await ART3WebAPI.Models.Domains.OcorrenciasGeradorRelatorio.GenerateXLS(relOcorrenciasList, filePath, oc_status, loja, transportadora);
                    statusResponse = "OK";
                    LogDAO.insere(httpRequestId, usuario, s_log, out strMsgErro);
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

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
        }

        [HttpPost]
        public HttpResponseMessage downloadXLS2(string fileName)
        {
			const string NOME_DESTA_ROTINA = "RelatoriosController.downloadXLS2()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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

        #region [ Requisição ]
        [HttpPost]
        public async Task<HttpResponseMessage> GetCompras2CSV(string usuario, string tipo_periodo, string dt_inicio, string dt_termino, string fabricante, string produto, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string nf, string dt_nf_inicio, string dt_nf_termino, string visao, string detalhamento)
        {
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "RelatoriosController.GetCompras2CSV()";
			Guid httpRequestId = Request.GetCorrelationId();
			DateTime data = DateTime.Now;
			string msg;
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
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", tipo_periodo=" + (tipo_periodo ?? "") + ", dt_inicio=" + (dt_inicio ?? "") + ", dt_termino=" + (dt_termino ?? "") + ", fabricante=" + (fabricante ?? "") + ", produto=" + (produto ?? "") + ", grupo=" + (grupo ?? "") + ", subgrupo=" + (subgrupo ?? "") + ", btu=" + (btu ?? "") + ", ciclo=" + (ciclo ?? "") + ", pos_mercado=" + (pos_mercado ?? "") + ", nf=" + (nf ?? "") + ", dt_nf_inicio=" + (dt_nf_inicio ?? "") + ", dt_nf_termino=" + (dt_nf_termino ?? "") + ", visao=" + (visao ?? "") + ", detalhamento=" + (detalhamento ?? "") + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

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
					#region [ Salva parâmetros no BD como valores default do usuário ]
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
					#endregion

					DataCompras2 datasource = new DataCompras2();
					List<Compras> relCompras2List = datasource.Get(tipo_periodo, dt_inicio, dt_termino, fabricante, produto, grupo, subgrupo, btu, ciclo, pos_mercado, nf, dt_nf_inicio, dt_nf_termino, visao, detalhamento).ToList();
					if (relCompras2List.Count != 0)
					{
						await ART3WebAPI.Models.Domains.Compras2GeradorRelatorio.GenerateXLS(relCompras2List, filePath, tipo_periodo, dt_inicio, dt_termino, fabricante, produto, grupo, subgrupo, btu, ciclo, pos_mercado, nf, dt_nf_inicio, dt_nf_termino, visao, detalhamento);
						statusResponse = "OK";
						LogDAO.insere(httpRequestId, usuario, s_log, out strMsgErro);
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

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
        }
        
        #endregion

        #region [ Download ]
        [HttpPost]
        public HttpResponseMessage downloadCompras2CSV(string fileName)
        {
			const string NOME_DESTA_ROTINA = "RelatoriosController.downloadCompras2CSV()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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

		#region [ Requisição ]
		[HttpPost]
        public async Task<HttpResponseMessage> GeraDevolucaoProdutos2XLS(string usuario, string dt_devolucao_inicio, string dt_devolucao_termino, string fabricante, string produto, string pedido, string vendedor, string indicador, string captador, string lojas)
        {
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "RelatoriosController.GeraDevolucaoProdutos2XLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			DateTime dataAtual;
			string msg;
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

			msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", dt_devolucao_inicio=" + (dt_devolucao_inicio ?? "") + ", dt_devolucao_termino=" + (dt_devolucao_termino ?? "") + ", fabricante=" + (fabricante ?? "") + ", produto=" + (produto ?? "") + ", pedido=" + (pedido ?? "") + ", vendedor=" + (vendedor ?? "") + ", indicador=" + (indicador ?? "") + ", captador=" + (captador ?? "") + ", lojas=" + (lojas ?? "") + ")";
			Global.gravaLogAtividade(httpRequestId, msg);

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
                    LogDAO.insere(httpRequestId, usuario, s_log, out strMsgErro);
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

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
        }
		#endregion

		#region [ Download ]
		[HttpPost]
        public HttpResponseMessage downloadDevolucaoProdutos2XLS(string fileName)
        {
			const string NOME_DESTA_ROTINA = "RelatoriosController.downloadDevolucaoProdutos2XLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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

		#region [ Relatório de Pré-Devolução ]

		#region [ Requisição ]
		[HttpGet]
		public async Task<HttpResponseMessage> RelPedidoPreDevolucaoXLS(string usuario, string loja, string sessionToken, string filtro_status, string filtro_data_inicio, string filtro_data_termino, string filtro_lojas)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "RelatoriosController.RelPedidoPreDevolucaoXLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			DateTime dataAtual;
			string msg;
			string msg_erro;
			string fileName;
			string filePath;
			StringBuilder sbResponse = new StringBuilder();
			string statusResponse = "";
			string strMsgErro = "";
			string MsgErroException = "";
			HttpResponseMessage result = null;
			Usuario usuarioBD;
			DataRelPreDevolucao datasource;
			List<RelPreDevolucaoEntity> dataRel;
			Log log = new Log();
			#endregion

			msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", loja=" + (loja ?? "") + ", sessionToken=" + (sessionToken ?? "") + ", filtro_status=" + (filtro_status ?? "") + ", filtro_data_inicio=" + (filtro_data_inicio ?? "") + ", filtro_data_termino=" + (filtro_data_termino ?? "") + ", filtro_lojas=" + (filtro_lojas ?? "") + ")";
			Global.gravaLogAtividade(httpRequestId, msg);

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				msg = "Não foi informada a identificação do usuário!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				msg = "Não foi informado o token da sessão do usuário!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				msg = "Falha ao tentar validar usuário!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				msg = "Token de sessão inválido!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}
			#endregion

			#region [ Gera nome do arquivo ]
			dataAtual = DateTime.Now;
			fileName = "RelPreDevolucao_" + dataAtual.ToString("yyyyMMdd_HHmmss");
			fileName = fileName + ".xlsx";
			filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);
			#endregion

			log.operacao = "RelPedPreDevolXLS";
			log.complemento = " Nome do arquivo: " + fileName + " (filtro_status=" + filtro_status + ", filtro_data_inicio=" + filtro_data_inicio + ", filtro_data_termino=" + filtro_data_termino + ", filtro_lojas=" + filtro_lojas + ")";
			log.usuario = usuario;
			log.loja = loja;

			try
			{
				datasource = new DataRelPreDevolucao();
				dataRel = datasource.Get(httpRequestId, usuario, loja, filtro_status, filtro_data_inicio, filtro_data_termino, filtro_lojas);
				if (dataRel.Count > 0)
				{
					await RelPreDevolucaoGeradorRelatorio.GeraXLS(dataRel, filePath, usuario, loja, filtro_status, filtro_data_inicio, filtro_data_termino, filtro_lojas);
					statusResponse = "OK";
					LogDAO.insere(httpRequestId, usuario, log, out strMsgErro);
				}
				else
				{
					statusResponse = "Vazio";
					MsgErroException = "Nenhum registro foi encontrado!";
					msg = NOME_DESTA_ROTINA + ": " + MsgErroException;
					Global.gravaLogAtividade(httpRequestId, msg);
				}
			}
			catch (Exception ex)
			{
				statusResponse = "Falha";
				MsgErroException = ex.Message;
				msg = NOME_DESTA_ROTINA + ": Exception - " + ex.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);
			}

			sbResponse.Append("{ \"fileName\" : \"" + fileName + "\", " + "\"Status\" : \"" + statusResponse + "\", " + "\"Exception\" : " + System.Web.Helpers.Json.Encode(MsgErroException) + "}");
			result = Request.CreateResponse(HttpStatusCode.OK);
			result.Content = new StringContent(sbResponse.ToString(), Encoding.UTF8, "application/json");

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ Download ]
		[HttpPost]
		public HttpResponseMessage DownloadRelPedidoPreDevolucaoXLS(string fileName)
		{
			const string NOME_DESTA_ROTINA = "RelatoriosController.DownloadRelPedidoPreDevolucaoXLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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

		#region [ Relatório Estoque de Venda ]

		#region [ Requisição ]
		[HttpGet]
		public async Task<HttpResponseMessage> RelEstoqueVendaXLS(string usuario, string loja, string sessionToken, string filtro_estoque, string filtro_detalhe, string filtro_consolidacao_codigos, string filtro_empresa, string filtro_fabricante, string filtro_produto, string filtro_fabricante_multiplo, string filtro_grupo, string filtro_subgrupo, string filtro_potencia_BTU, string filtro_ciclo, string filtro_posicao_mercado)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "RelatoriosController.RelEstoqueVendaXLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			DateTime dataAtual;
			string ID_RELATORIO;
			string msg;
			string msg_erro;
			string fileName;
			string filePath;
			StringBuilder sbResponse = new StringBuilder();
			string statusResponse = "";
			string strMsgErro = "";
			string MsgErroException = "";
			HttpResponseMessage result = null;
			Usuario usuarioBD;
			DataRelEstoqueVenda datasource;
			List<RelEstoqueVendaEntity> dataRel;
			Log log = new Log();
			#endregion

			msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", loja=" + (loja ?? "") + ", sessionToken=" + (sessionToken ?? "") + ", filtro_estoque=" + (filtro_estoque ?? "") + ", filtro_detalhe=" + (filtro_detalhe ?? "") + ", filtro_consolidacao_codigos=" + (filtro_consolidacao_codigos ?? "") + ", filtro_empresa=" + (filtro_empresa ?? "") + ", filtro_fabricante=" + (filtro_fabricante ?? "") + ", filtro_produto=" + (filtro_produto ?? "") + ", filtro_fabricante_multiplo=" + (filtro_fabricante_multiplo ?? "") + ", filtro_grupo=" + (filtro_grupo ?? "") + ", filtro_subgrupo=" + (filtro_subgrupo ?? "") + ", filtro_potencia_BTU=" + (filtro_potencia_BTU ?? "") + ", filtro_ciclo=" + (filtro_ciclo ?? "") + ", filtro_posicao_mercado=" + (filtro_posicao_mercado ?? "") + ")";
			Global.gravaLogAtividade(httpRequestId, msg);

			#region [ Validação de segurança: session token confere? ]
			if ((usuario ?? "").Trim().Length == 0)
			{
				msg = "Não foi informada a identificação do usuário!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}

			if ((sessionToken ?? "").Trim().Length == 0)
			{
				msg = "Não foi informado o token da sessão do usuário!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}

			usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
			if (usuarioBD == null)
			{
				msg = "Falha ao tentar validar usuário!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}

			if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
			{
				msg = "Token de sessão inválido!";
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}
			#endregion

			#region [ Validação dos filtros obrigatórios ]

			#region [ filtro_estoque ]
			if (!(filtro_estoque ?? "").Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_ESTOQUE.ESTOQUE_VENDA))
			{
				msg = "Filtro com valor inválido: filtro_estoque=" + (filtro_estoque ?? "");
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}
			#endregion

			#region [ filtro_detalhe ]
			if (
				(!(filtro_detalhe ?? "").Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.SINTETICO))
				&&
				(!(filtro_detalhe ?? "").Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
				)
			{
				msg = "Filtro com valor inválido: filtro_detalhe=" + (filtro_detalhe ?? "");
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}
			#endregion

			#region [ filtro_consolidacao_codigos ]
			if (
				(!(filtro_consolidacao_codigos ?? "").Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_CONSOLIDACAO_CODIGOS.NORMAIS))
				&&
				(!(filtro_consolidacao_codigos ?? "").Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_CONSOLIDACAO_CODIGOS.UNIFICADOS))
				)
			{
				msg = "Filtro com valor inválido: filtro_consolidacao_codigos=" + (filtro_consolidacao_codigos ?? "");
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + msg);
				throw new Exception(msg);
			}
			#endregion
			#endregion

			#region [ Salva parâmetros no BD como valores default do usuário ]
			if ((loja ?? "").Length > 0)
			{
				ID_RELATORIO = "LOJA/RelEstoqueVendaCmvPv";
			}
			else
			{
				ID_RELATORIO = "CENTRAL/RelEstoqueVendaCmvPv";
			}

			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "rb_detalhe", (filtro_detalhe ?? ""));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "rb_exportacao", (filtro_consolidacao_codigos ?? ""));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "c_fabricante_multiplo", (filtro_fabricante_multiplo ?? "").Replace("_", ", "));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "c_grupo", (filtro_grupo ?? "").Replace("_", ", "));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "c_subgrupo", (filtro_subgrupo ?? "").Replace("_", ", "));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "c_potencia_BTU", (filtro_potencia_BTU ?? ""));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "c_ciclo", (filtro_ciclo ?? ""));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "c_posicao_mercado", (filtro_posicao_mercado ?? ""));
			Global.setDefaultBD(usuario, ID_RELATORIO + "|" + "rb_saida", "XLS");
			#endregion

			#region [ Gera nome do arquivo ]
			dataAtual = DateTime.Now;
			fileName = "RelEstoqueVenda_" + dataAtual.ToString("yyyyMMdd_HHmmss");
			fileName = fileName + ".xlsx";
			filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileName);
			#endregion

			log.operacao = "RelEstoqueVendaXLS";
			log.complemento = " Nome do arquivo: " + fileName + " (filtro_estoque=" + (filtro_estoque ?? "") + ", filtro_detalhe=" + (filtro_detalhe ?? "") + ", filtro_consolidacao_codigos=" + (filtro_consolidacao_codigos ?? "") + ", filtro_empresa=" + (filtro_empresa ?? "") + ", filtro_fabricante=" + (filtro_fabricante ?? "") + ", filtro_produto=" + (filtro_produto ?? "") + ", filtro_fabricante_multiplo=" + (filtro_fabricante_multiplo ?? "") + ", filtro_grupo=" + (filtro_grupo ?? "") + ", filtro_subgrupo=" + (filtro_subgrupo ?? "") + ", filtro_potencia_BTU=" + (filtro_potencia_BTU ?? "") + ", filtro_ciclo=" + (filtro_ciclo ?? "") + ", filtro_posicao_mercado=" + (filtro_posicao_mercado ?? "") + ")";
			log.usuario = usuario;
			log.loja = loja;

			try
			{
				datasource = new DataRelEstoqueVenda();
				dataRel = datasource.Get(httpRequestId, usuario, loja, filtro_estoque, filtro_detalhe, filtro_consolidacao_codigos, filtro_empresa, filtro_fabricante, filtro_produto, filtro_fabricante_multiplo, filtro_grupo, filtro_subgrupo, filtro_potencia_BTU, filtro_ciclo, filtro_posicao_mercado);
				if (dataRel.Count > 0)
				{
					await RelEstoqueVendaGeradorRelatorio.GeraXLS(dataRel, filePath, usuario, loja, filtro_estoque, filtro_detalhe, filtro_consolidacao_codigos, filtro_empresa, filtro_fabricante, filtro_produto, filtro_fabricante_multiplo, filtro_grupo, filtro_subgrupo, filtro_potencia_BTU, filtro_ciclo, filtro_posicao_mercado);
					statusResponse = "OK";
					LogDAO.insere(httpRequestId, usuario, log, out strMsgErro);
				}
				else
				{
					statusResponse = "Vazio";
					MsgErroException = "Nenhum registro foi encontrado!";
					msg = NOME_DESTA_ROTINA + ": " + MsgErroException;
					Global.gravaLogAtividade(httpRequestId, msg);
				}
			}
			catch (Exception ex)
			{
				statusResponse = "Falha";
				MsgErroException = ex.Message;
				msg = NOME_DESTA_ROTINA + ": Exception - " + ex.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);
			}

			sbResponse.Append("{ \"fileName\" : \"" + fileName + "\", " + "\"Status\" : \"" + statusResponse + "\", " + "\"Exception\" : " + System.Web.Helpers.Json.Encode(MsgErroException) + "}");
			result = Request.CreateResponse(HttpStatusCode.OK);
			result.Content = new StringContent(sbResponse.ToString(), Encoding.UTF8, "application/json");

			msg = NOME_DESTA_ROTINA + ": Status=" + statusResponse + ", fileName=" + fileName;
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ Download ]
		[HttpPost]
		public HttpResponseMessage DownloadRelEstoqueVendaXLS(string fileName)
		{
			const string NOME_DESTA_ROTINA = "RelatoriosController.DownloadRelEstoqueVendaXLS()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": fileName=" + (fileName ?? "");
			Global.gravaLogAtividade(httpRequestId, msg);

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
