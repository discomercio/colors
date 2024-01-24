using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Script.Serialization;
using System.ServiceModel.Channels;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using ART3WebAPI.Models.Domains;
using System.IO.Compression;
using System.Net.Http.Headers;

namespace ART3WebAPI.Controllers
{
	public class DownloadCompressedFileController : ApiController
	{
		#region [ GetIp ]
		public string GetIp()
		{
			return GetClientIp();
		}
		#endregion

		#region [ GetClientIp ]
		private string GetClientIp(HttpRequestMessage request = null)
		{
			request = request ?? Request;

			if (request.Properties.ContainsKey("MS_HttpContext"))
			{
				return ((HttpContextWrapper)request.Properties["MS_HttpContext"]).Request.UserHostAddress;
			}
			else if (request.Properties.ContainsKey(RemoteEndpointMessageProperty.Name))
			{
				RemoteEndpointMessageProperty prop = (RemoteEndpointMessageProperty)request.Properties[RemoteEndpointMessageProperty.Name];
				return prop.Address;
			}
			else if (HttpContext.Current != null)
			{
				return HttpContext.Current.Request.UserHostAddress;
			}
			else
			{
				return null;
			}
		}
		#endregion

		#region [ Teste ]
		[HttpGet]
		public HttpResponseMessage Teste()
		{
			const string NOME_DESTA_ROTINA = "DownloadCompressedFileController.Teste()";
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

		#region [ MultipleNFeXml ]
		// POST api/<controller>
		[HttpPost]
		public HttpResponseMessage MultipleNFeXml([FromBody] MultipleNFeXmlPostRequest dadosReq)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DownloadCompressedFileController.MultipleNFeXml()";
			string compressedFileName = "";
			string compressedFileNameWithoutExt = "";
			string compressedFullFileName = "";
			string compressedTempFolderName;
			string compressedFullTempFolderName;
			string xmlFileName;
			string xmlFileNameAux;
			string xmlFilePath;
			string xmlFileNameFull;
			string msg;
			string msg_erro_aux;
			bool blnXmlFileExists;
			int idx = 0;
			int qtdeArquivosNFeXmlCompactados = 0;
			StringBuilder sbMsgErro;
			StringBuilder sbMsg;
			Usuario usuarioBD;
			Guid httpRequestId = Request.GetCorrelationId();
			NFeEmitente emitente;
			List<NFeEmitente> listaEmitentes;
			MultipleNFeXmlDownloadProc nfeXmlProc;
			List<MultipleNFeXmlDownloadProc> listaNFeXmlProc = new List<MultipleNFeXmlDownloadProc>();
			DownloadCompressedFileResponse downloadFileResponse = new DownloadCompressedFileResponse();
			HttpResponseMessage result;
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (" + new JavaScriptSerializer().Serialize(dadosReq) + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				#region [ Parâmetro: identificação do usuário ]
				if ((dadosReq.formData.download_parameter__user_id == 0) && ((dadosReq.formData.download_parameter__username ?? "").Trim().Length == 0))
				{
					throw new WebApiException("Não foi informada a identificação do usuário!");
				}
				#endregion

				#region [ Validação de segurança: session token confere? ]
				if (dadosReq.formData.download_parameter__user_id > 0)
				{
					usuarioBD = GeralDAO.getUsuario(dadosReq.formData.download_parameter__user_id, out msg_erro_aux);
				}
				else
				{
					usuarioBD = GeralDAO.getUsuario(dadosReq.formData.download_parameter__username, out msg_erro_aux);
				}

				if (usuarioBD == null)
				{
					throw new WebApiException("Falha ao tentar validar usuário!");
				}

				if ((!usuarioBD.SessionTokenModuloCentral.Equals(dadosReq.formData.download_parameter__sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(dadosReq.formData.download_parameter__sessionToken)))
				{
					throw new WebApiException("Token de sessão inválido!");
				}
				#endregion

				#region [ Há dados de NFe XML? ]
				if (dadosReq.NFeXmlDownloadList == null)
				{
					throw new WebApiException("Não foi solicitado nenhum arquivo de NFe XML para download!");
				}

				if (dadosReq.NFeXmlDownloadList.Count == 0)
				{
					throw new WebApiException("Não foi informado nenhum arquivo de NFe XML para download!");
				}

				sbMsgErro = new StringBuilder("");
				foreach (MultipleNFeXmlPostRequestNFeXmlId nfeXml in dadosReq.NFeXmlDownloadList)
				{
					idx++;

					if (nfeXml.id_nfe_emitente == 0)
					{
						sbMsgErro.AppendLine("NFe XML (linha " + idx.ToString() + "): não foi informado o Id do emitente");
					}

					if (nfeXml.NFe_serie_NF.Trim().Length == 0)
					{
						sbMsgErro.AppendLine("NFe XML (linha " + idx.ToString() + "): não foi informado o nº série da NFe");
					}

					if (nfeXml.NFe_numero_NF.Trim().Length == 0)
					{
						sbMsgErro.AppendLine("NFe XML (linha " + idx.ToString() + "): não foi informado o nº da NFe");
					}
				}

				if (sbMsgErro.Length > 0)
				{
					throw new WebApiException("Há dados inconsistentes na solicitação de NFe XML para download!");
				}
				#endregion

				#region [ Definição do nome do arquivo compactado ]
				if ((dadosReq.formData.download_parameter__CompressedFile_Filename_Prefix ?? "").Trim().Length > 0)
				{
					compressedFileName = dadosReq.formData.download_parameter__CompressedFile_Filename_Prefix.Trim();
				}

				if (dadosReq.formData.download_parameter__CompressedFile_FileName_IncludeGuid == 1)
				{
					if ((compressedFileName.Length > 0) && (!compressedFileName.EndsWith("_"))) compressedFileName += "_";
					compressedFileName += BD.gera_uid(httpRequestId);
				}

				if (dadosReq.formData.download_parameter__CompressedFile_Filename_IncludeDateTime == 1)
				{
					if ((compressedFileName.Length > 0) && (!compressedFileName.EndsWith("_"))) compressedFileName += "_";
					compressedFileName += DateTime.Now.ToString(Global.Cte.DataHora.FmtYYYYMMDD) + "_" + DateTime.Now.ToString(Global.Cte.DataHora.FmtHHMMSS);
				}

				if (dadosReq.formData.download_parameter__CompressedFile_Filename_IncludeUserId == 1)
				{
					if ((compressedFileName.Length > 0) && (!compressedFileName.EndsWith("_"))) compressedFileName += "_";
					compressedFileName += "Usuario" + usuarioBD.Id.ToString();
				}

				// Se nenhum nome foi gerado pelas definições dos parâmetros, define um Guid
				if (compressedFileName.Trim().Length == 0) compressedFileName = BD.gera_uid(httpRequestId);

				compressedFileNameWithoutExt = compressedFileName;
				compressedFileName = compressedFileNameWithoutExt + ".zip";
				compressedTempFolderName = compressedFileNameWithoutExt;
				compressedFullFileName = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + compressedFileName);
				compressedFullTempFolderName = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + compressedTempFolderName);

				#region [ Se o nome já estiver em uso, adiciona um Guid ao final ]
				if (File.Exists(compressedFullFileName) || Directory.Exists(compressedFullTempFolderName))
				{
					if (!compressedFileNameWithoutExt.EndsWith("_")) compressedFileNameWithoutExt += "_";
					compressedFileNameWithoutExt += BD.gera_uid(httpRequestId);
					compressedFileName = compressedFileNameWithoutExt + ".zip";
					compressedTempFolderName = compressedFileNameWithoutExt;
					compressedFullFileName = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + compressedFileName);
					compressedFullTempFolderName = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + compressedTempFolderName);
				}
				#endregion

				#endregion

				listaEmitentes = NFeEmitenteDAO.getAllNFeEmitente();

				#region [ Para cada NFe solicitada, tenta localizar o arquivo XML ]
				idx = 0;
				foreach (MultipleNFeXmlPostRequestNFeXmlId nfeXmlId in dadosReq.NFeXmlDownloadList)
				{
					idx++;
					blnXmlFileExists = false;

					nfeXmlProc = new MultipleNFeXmlDownloadProc();
					nfeXmlProc.id_nfe_emitente = nfeXmlId.id_nfe_emitente;
					nfeXmlProc.NFe_serie_NF = nfeXmlId.NFe_serie_NF;
					nfeXmlProc.NFe_numero_NF = nfeXmlId.NFe_numero_NF;

					try // try-finally (adiciona na lista)
					{
						try
						{
							emitente = listaEmitentes.Single(e => e.id == nfeXmlId.id_nfe_emitente);
						}
						catch (Exception)
						{
							emitente = null;
						}

						if (emitente == null)
						{
							nfeXmlProc.msg_erro = "Falha ao localizar os dados do emitente " + nfeXmlId.id_nfe_emitente.ToString() + " (linha " + idx.ToString() + ")";
							continue;
						}

						foreach (NFeEmitenteCfgDanfe cfgDanfe in emitente.listaCfgDanfe)
						{
							xmlFileName = "";
							xmlFileNameAux = (cfgDanfe.convencao_nome_arq_xml_nfe ?? "").Trim();
							if (xmlFileNameAux.Length > 0)
							{
								xmlFileName = xmlFileNameAux.Replace("[NUMERO_NFE]", nfeXmlId.NFe_numero_NF).Replace("[SERIE_NFE]", nfeXmlId.NFe_serie_NF);
								xmlFilePath = cfgDanfe.diretorio_xml_nfe;
								xmlFileNameFull = xmlFilePath;
								if (!xmlFileNameFull.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) xmlFileNameFull += Path.DirectorySeparatorChar;
								xmlFileNameFull += xmlFileName;
								blnXmlFileExists = File.Exists(xmlFileNameFull);
								if (blnXmlFileExists)
								{
									nfeXmlProc.existe_arquivo = true;
									nfeXmlProc.diretorio = cfgDanfe.diretorio_xml_nfe;
									nfeXmlProc.nome_arquivo = xmlFileName;
									nfeXmlProc.nome_arquivo_completo = xmlFileNameFull;
									break;
								}
							}
						}

						if (!blnXmlFileExists)
						{
							nfeXmlProc.msg_erro = "Arquivo NFe XML não encontrado para o emitente " + nfeXmlId.id_nfe_emitente.ToString() + ", nº série " + nfeXmlId.NFe_serie_NF + ", nº NF " + nfeXmlId.NFe_numero_NF + " (linha " + idx.ToString() + ")";
						}
					}
					finally
					{
						listaNFeXmlProc.Add(nfeXmlProc);
					}
				}
				#endregion

				#region [ Copia arquivos para pasta temporária ]
				Directory.CreateDirectory(compressedFullTempFolderName);
				try // try-finally (remove o diretório temporário)
				{
					foreach (MultipleNFeXmlDownloadProc regNFeXmlProc in listaNFeXmlProc)
					{
						if (!regNFeXmlProc.existe_arquivo) continue;

						xmlFileNameAux = compressedFullTempFolderName + Path.DirectorySeparatorChar.ToString() + regNFeXmlProc.nome_arquivo;
						File.Copy(regNFeXmlProc.nome_arquivo_completo, xmlFileNameAux);
						qtdeArquivosNFeXmlCompactados++;
					}
					#endregion

					#region [ Compacta a pasta temporária ]
					ZipFile.CreateFromDirectory(compressedFullTempFolderName, compressedFullFileName);
					#endregion
				}
				finally
				{
					#region [ Remove a pasta temporária ]
					Directory.Delete(compressedFullTempFolderName, recursive: true);
					#endregion
				}

				downloadFileResponse.Status = "OK";
				downloadFileResponse.CompressedFileName = compressedFileName;
				downloadFileResponse.CompressedFilesQuantity = qtdeArquivosNFeXmlCompactados;

				sbMsg = new StringBuilder();
				foreach (MultipleNFeXmlDownloadProc regNFeXmlProc in listaNFeXmlProc)
				{
					if ((regNFeXmlProc.msg_erro ?? "").Trim().Length > 0)
					{
						sbMsg.AppendLine(regNFeXmlProc.msg_erro);
					}
				}

				if (sbMsg.Length > 0)
				{
					if ((downloadFileResponse.Message ?? "").Trim().Length > 0) downloadFileResponse.Message += "\n";
					downloadFileResponse.Message += sbMsg.ToString();
				}

				var jsonOk = new JavaScriptSerializer().Serialize(downloadFileResponse);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(jsonOk, Encoding.UTF8, "text/html");

				msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return result;
			}
			catch (System.Exception e)
			{
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + " - Exception: " + e.ToString());

				if (e is WebApiException)
				{
					downloadFileResponse.Status = "ERROR";
					downloadFileResponse.Message = e.Message;
					var jsonError = new JavaScriptSerializer().Serialize(downloadFileResponse);
					result = Request.CreateResponse(HttpStatusCode.OK);
					result.Content = new StringContent(jsonError, Encoding.UTF8, "text/html");
					return result;
				}
				else
				{
					return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
				}
			}
		}
		#endregion

		#region [ DownloadCompressedMultipleNFeXml ]
		[HttpPost]
		public HttpResponseMessage DownloadCompressedMultipleNFeXml(string fileName)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DownloadCompressedFileController.DownloadCompressedMultipleNFeXml()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro_aux;
			string fileNameToDownload;
			string fileNameFromQueryString;
			string fileNameFromFormData;
			string userName;
			string userId;
			string sessionToken;
			Usuario usuarioBD;
			DownloadCompressedFileResponse downloadFileResponse = new DownloadCompressedFileResponse();
			HttpResponseMessage result;
			#endregion

			try
			{
				userName = (HttpContext.Current.Request["download_compressed_file_parameter__username"] ?? "");
				userId = (HttpContext.Current.Request["download_compressed_file_parameter__user_id"] ?? "");
				sessionToken = (HttpContext.Current.Request["download_compressed_file_parameter__sessionToken"] ?? "");
				fileNameFromQueryString = fileName;
				fileNameFromFormData = (HttpContext.Current.Request["download_compressed_file_parameter__compressedFileName"] ?? "");

				#region [ Validação de segurança: session token confere? ]
				if ((Global.converteInteiro(userId) == 0)&&(userName.Trim().Length == 0))
				{
					throw new WebApiException("Não foi informada a identificação do usuário!");
				}

				if (Global.converteInteiro(userId) > 0)
				{
					usuarioBD = GeralDAO.getUsuario((int)Global.converteInteiro(userId), out msg_erro_aux);
				}
				else
				{
					usuarioBD = GeralDAO.getUsuario(userName, out msg_erro_aux);
				}

				if (usuarioBD == null)
				{
					throw new WebApiException("Falha ao tentar validar usuário!");
				}

				if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
				{
					throw new WebApiException("Token de sessão inválido!");
				}
				#endregion

				#region [ Define o nome de arquivo a ser usado, sendo que a opção via form data tem preferência ]
				if ((fileNameFromFormData ?? "").Trim().Length > 0)
				{
					fileNameToDownload = fileNameFromFormData;
				}
				else
				{
					fileNameToDownload = fileNameFromQueryString;
				}
				#endregion

				msg = NOME_DESTA_ROTINA + ": fileName=" + (fileNameToDownload ?? "") + " (fileNameFromFormData=" + fileNameFromFormData + ", fileNameFromQueryString=" + fileNameFromQueryString + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				string filePath = HttpContext.Current.Server.MapPath("~/Report/Relatorios/" + fileNameToDownload);

				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StreamContent(new FileStream(filePath, FileMode.Open));
				result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
				result.Content.Headers.ContentDisposition.FileName = fileNameToDownload;
				result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/zip");

				return result;
			}
			catch (Exception e)
			{
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + " - Exception: " + e.ToString());

				if (e is WebApiException)
				{
					downloadFileResponse.Status = "ERROR";
					downloadFileResponse.Message = e.Message;
					var jsonError = new JavaScriptSerializer().Serialize(downloadFileResponse);
					result = Request.CreateResponse(HttpStatusCode.OK);
					result.Content = new StringContent(jsonError, Encoding.UTF8, "text/html");
					return result;
				}
				else
				{
					return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
				}
			}
		}
		#endregion
	}

	public class MultipleNFeXmlPostRequestNFeXmlId
	{
		public int id_nfe_emitente { get; set; }
		public string NFe_serie_NF { get; set; }
		public string NFe_numero_NF { get; set; }
	}

	public class MultipleNFeXmlPostRequestFormData
	{
		public string SessionCtrlInfo { get; set; }
		public string download_parameter__username { get; set; }
		public int download_parameter__user_id { get; set; }
		public string download_parameter__sessionToken { get; set; }
		public string download_parameter__CompressedFile_Filename_Prefix { get; set; }
		public int download_parameter__CompressedFile_Filename_IncludeDateTime { get; set; }
		public int download_parameter__CompressedFile_FileName_IncludeGuid { get; set; }
		public int download_parameter__CompressedFile_Filename_IncludeUserId { get; set; }
	}

	public class MultipleNFeXmlPostRequest
	{
		public MultipleNFeXmlPostRequestFormData formData { get; set; }
		public List<MultipleNFeXmlPostRequestNFeXmlId> NFeXmlDownloadList { get; set; }
	}

	public class MultipleNFeXmlDownloadProc
	{
		public int id_nfe_emitente { get; set; }
		public string NFe_serie_NF { get; set; }
		public string NFe_numero_NF { get; set; }
		public string nome_arquivo { get; set; }
		public string diretorio { get; set; }
		public string nome_arquivo_completo { get; set; }
		public bool existe_arquivo { get; set; }
		public string msg_erro { get; set; }
	}
}