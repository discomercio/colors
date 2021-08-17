using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Text;
using System.IO;
using System.Xml;
using Newtonsoft.Json;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using ART3WebAPI.Models.Domains;

namespace ART3WebAPI.Controllers
{
	public class UploadedFileController : ApiController
	{
		#region [ Teste ]
		[HttpGet]
		public HttpResponseMessage Teste()
		{
			const string NOME_DESTA_ROTINA = "UploadedFileController.Teste()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": Requisição recebida";
			Global.gravaLogAtividade(httpRequestId, msg);

			HttpResponseMessage result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Versao.M_ID);

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion

		#region [ ConvertXmlToJson ]
		[HttpGet]
		public HttpResponseMessage ConvertXmlToJson(Guid? id)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "UploadedFileController.ConvertXmlToJson()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string guid = "";
			string s;
			string msg_erro_aux;
			string sXml;
			string sJson;
			HttpResponseMessage result;
			UploadStoredFileInfo fileInfo = null;
			XmlDocument doc = new XmlDocument();
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (id=" + (id != null ? id.ToString() : "") + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				if (id != null) guid = id.ToString();

				if (guid.Length == 0)
				{
					throw new Exception("Identificador do arquivo é inválido!");
				}

				fileInfo = UploadFileDAO.getUploadStoredFileInfoByGuid(guid, out msg_erro_aux);

				if (fileInfo == null)
				{
					throw new Exception("Não foi localizado o registro com as informações do arquivo referente ao identificador " + guid);
				}

				if (!Path.GetExtension(fileInfo.stored_full_file_name).ToUpper().Equals(".XML"))
				{
					throw new Exception("O arquivo '" + fileInfo.original_file_name + "' não possui a extensão XML");
				}

				if ((fileInfo.file_content == null) && (fileInfo.file_content_text == null) && (!File.Exists(fileInfo.stored_full_file_name)))
				{
					throw new Exception("Não foi encontrado o arquivo referente ao identificador " + guid);
				}

				#region [ Lê o conteúdo do arquivo XML ]
				if (fileInfo.file_content != null)
				{
					MemoryStream ms = new MemoryStream();
					ms.Write(fileInfo.file_content, 0, fileInfo.file_content.Length);
					ms.Flush();
					ms.Position = 0;
					StreamReader sr = new StreamReader(ms);
					sXml = sr.ReadToEnd();
				}
				else if ((fileInfo.file_content_text ?? "").Trim().Length > 0)
				{
					sXml = fileInfo.file_content_text;
				}
				else
				{
					sXml = File.ReadAllText(fileInfo.stored_full_file_name);
				}

				if ((sXml ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi possível recuperar o conteúdo XML do arquivo especificado (" + fileInfo.original_file_name + ")!");
				}

				doc.LoadXml(sXml);
				sJson = JsonConvert.SerializeXmlNode(doc);
				#endregion

				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(sJson, Encoding.UTF8, "text/html");

				msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return result;
			}
			catch (Exception e)
			{
				s = "";
				if (fileInfo != null) s = ", arquivo=" + fileInfo.original_file_name;
				s = NOME_DESTA_ROTINA + " - [t_UPLOAD_FILE.guid=" + guid + s + "] - Exception: " + e.ToString();
				Global.gravaLogAtividade(httpRequestId, s);
				return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
			}
		}
		#endregion
	}
}
