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
    public class CiagroupController : ApiController
    {
        [HttpPost]
        public async Task<HttpResponseMessage> GetCSVReport(int id, string usuario)
        {
            DataCiagroup datasource = new DataCiagroup();
            List<Indicador> indicadorList = datasource.Get(id).ToList();
            List<string> vendedoresList = new List<string>();

            indicadorList.ForEach(delegate(Indicador ind)
            {
                vendedoresList.Add(ind.Vendedor);
            });

            string strVendedores = "";
            string vendedor_a = "xxx";
            vendedoresList.Sort();

            string[] vendedoresArray = vendedoresList.Distinct().ToArray();

            for (int i = 0; i < vendedoresArray.ToArray().GetLength(0); i++)
            {
                if (vendedoresArray[i] != vendedor_a)
                {
                    if (i < 2)
                    {
                        if (strVendedores != "") strVendedores = strVendedores + "__";
                        strVendedores = strVendedores + vendedoresArray[i];
                        vendedor_a = vendedoresArray[i];
                    }
                    if (i == 2 && vendedoresArray.ToArray().GetLength(0) > 2)
                    {
                        strVendedores = strVendedores + "__ETC";
                    }
                }
            }

            strVendedores = strVendedores.Replace(" ", "_");
            strVendedores = Global.filtraAcentuacao(strVendedores);

            DateTime data = DateTime.Now;
            string fileName = "Ciagroup_" + data.ToString("yyyyMMdd_HHmmss") + "_" + strVendedores;
            fileName = fileName + ".csv";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Ciagroup/" + fileName);

            Log s_log = new Log();
            s_log.complemento = "t_COMISSAO_INDICADOR_N1.id = " + id + ", Nome do arquivo: " + fileName;
            s_log.operacao = "Arq Ciagroup";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "n";

            HttpResponseMessage result = null;

            StringBuilder xmlResponse = new StringBuilder();

            try
            {
                await ART3WebAPI.Models.Domains.CiagroupGeradorRelatorio.GenerateCSV(indicadorList, filePath);

                statusResponse = "OK";

                LogDAO.insere(usuario, s_log, out strMsgErro);

            }
            catch (Exception e)
            {
                statusResponse = "Falha";
                MsgErroException = e.Message;
            }

            xmlResponse.Append("{ 'fileName' : '" + fileName + "', " + "'Status' : '" + statusResponse + "', " + "'Exception' : '" + MsgErroException + "' }");
            xmlResponse = xmlResponse.Replace("'", ((char) 34).ToString());

            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StringContent(xmlResponse.ToString(), Encoding.UTF8, "text/html");

            return result;
        }

        [HttpPost]
        public HttpResponseMessage downloadCSV(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Ciagroup/" + fileName);

            HttpResponseMessage result = null;
            result = Request.CreateResponse(HttpStatusCode.OK);
            result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = fileName;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return result;
        }

        [HttpPost]
        public async Task<HttpResponseMessage> GetXLSReport(int id, string usuario)
        {
            DataCiagroup datasource = new DataCiagroup();
            List<Indicador> indicadorList = datasource.Get(id).ToList();
            List<string> vendedoresList = new List<string>();

            indicadorList.ForEach(delegate(Indicador ind)
            {
                vendedoresList.Add(ind.Vendedor);
            });

            string strVendedores = "";
            string vendedor_a = "xxx";
            int lengthArray = vendedoresList.ToArray().GetLength(0);
            vendedoresList.Sort();

            string[] vendedoresArray = vendedoresList.Distinct().ToArray();

            vendedor_a = "xxx";
            for (int i = 0; i < vendedoresArray.ToArray().GetLength(0); i++)
            {
                if (vendedoresArray[i] != vendedor_a)
                {
                    if (i < 2)
                    {
                        if (strVendedores != "") strVendedores = strVendedores + "__";
                        strVendedores = strVendedores + vendedoresArray[i];
                        vendedor_a = vendedoresArray[i];
                    }
                    if (i == 2 && vendedoresArray.ToArray().GetLength(0) > 2)
                    {
                        strVendedores = strVendedores + "__ETC";
                    }
                }
            }

            strVendedores = strVendedores.Replace(" ", "_");
            strVendedores = Global.filtraAcentuacao(strVendedores);

            DateTime data = DateTime.Now;
            string fileName = "Ciagroup_" + data.ToString("yyyyMMdd_HHmmss") + "_" + strVendedores;
            fileName = fileName + ".xlsx";
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Ciagroup/" + fileName);

            Log s_log = new Log();
            s_log.complemento = "t_COMISSAO_INDICADOR_N1.id = " + id + ", Nome do arquivo: " + fileName;
            s_log.operacao = "Arq Ciagroup";
            s_log.usuario = usuario;
            string strMsgErro = "";
            string statusResponse = "";
            string MsgErroException = "n";

            HttpResponseMessage result = null;

            StringBuilder xmlResponse = new StringBuilder();

            try
            {
                await ART3WebAPI.Models.Domains.CiagroupGeradorRelatorio.GenerateXLS(indicadorList, filePath);

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

        [HttpPost]
        public HttpResponseMessage downloadXLS(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/Report/Ciagroup/" + fileName);

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
