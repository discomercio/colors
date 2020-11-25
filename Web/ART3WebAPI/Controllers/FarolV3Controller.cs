﻿using System.Collections.Generic;
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
using System.Data;

namespace ART3WebAPI.Controllers
{
    public class FarolV3Controller : ApiController
    {
        public const string COD_CONSULTA_POR_PERIODO_CADASTRO = "CADASTRO";
        public const string COD_CONSULTA_POR_PERIODO_ENTREGA = "ENTREGA";

        [HttpPost]
        public async Task<HttpResponseMessage> GetXLSReport(string usuario, string opcao_periodo, string dt_inicio, string dt_termino, string fabricante, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string perc_est_cresc, string loja, string visao)
        {
            if (string.IsNullOrEmpty(opcao_periodo.ToString())) throw new Exception("Não foi informado o tipo de período de consulta!");
            if (opcao_periodo.Equals(COD_CONSULTA_POR_PERIODO_CADASTRO))
            {
                if (string.IsNullOrEmpty(dt_inicio.ToString())) throw new Exception("Não foi informada a data inicial do período de vendas!");
                if (string.IsNullOrEmpty(dt_termino.ToString())) throw new Exception("Não foi informada a data final do período de vendas!");
            }
            else if (opcao_periodo.Equals(COD_CONSULTA_POR_PERIODO_ENTREGA))
            {
                if (string.IsNullOrEmpty(dt_inicio.ToString())) throw new Exception("Não foi informada a data inicial do período de entrega!");
                if (string.IsNullOrEmpty(dt_termino.ToString())) throw new Exception("Não foi informada a data final do período de entrega!");
            }
            else
            {
                throw new Exception("Opção de período de consulta inválido!");
            }

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
                Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|rb_periodo", opcao_periodo);
                if (opcao_periodo.Equals(COD_CONSULTA_POR_PERIODO_CADASTRO))
                {
                    Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_dt_periodo_inicio", dt_inicio);
                    Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_dt_periodo_termino", dt_termino);
                }
                else if (opcao_periodo.Equals(COD_CONSULTA_POR_PERIODO_ENTREGA)){
                    Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_dt_entregue_inicio", dt_inicio);
                    Global.setDefaultBD(usuario, "RelFarolResumidoFiltro|c_dt_entregue_termino", dt_termino);
                }
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
                List<Farol> relFarolList = datasource.GetV3(opcao_periodo, dt_inicio, dt_termino, fabricante, grupo, subgrupo, btu, ciclo, pos_mercado, loja).ToList();

                if (relFarolList.Count != 0)
                {
                    await ART3WebAPI.Models.Domains.FarolGeradorRelatorio.GenerateXLSv3(relFarolList, filePath, opcao_periodo, dt_inicio, dt_termino, fabricante, grupo, subgrupo, btu, ciclo, pos_mercado, perc_est_cresc, loja, visao);
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
