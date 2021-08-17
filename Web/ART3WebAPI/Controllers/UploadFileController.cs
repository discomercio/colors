using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

namespace ART3WebAPI.Controllers
{
	public class UploadFileController : ApiController
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
			const string NOME_DESTA_ROTINA = "UploadFileController.Teste()";
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

		#region [ PostFile ]
		/// <summary>
		/// Parâmetros obtidos através da leitura de campos do formulário:
		///		'upload_parameter__user_id': identificação do usuário (obrigatório).
		///		'upload_parameter__sessionToken': token da sessão do usuário usado para assegurar que a requisição está sendo realizada por um usuário autenticado.
		///		'upload_parameter__folder_name': nome da pasta em que deve ser armazenado o arquivo (obrigatório).
		///		'upload_parameter__is_temp_file': indica se o arquivo é temporário, ou seja, se será excluído automaticamente pela rotina automática do servidor.
		///		'upload_parameter__is_confirmation_required': indica se será necessária a confirmação para considerar o arquivo como válido. A confirmação é feita
		///				através do campo t_UPLOAD_FILE.st_confirmation_ok
		///				Os arquivos com t_UPLOAD_FILE.st_confirmation_required = 1 e t_UPLOAD_FILE.st_confirmation_ok = 0 serão excluídos automaticamente.
		///		'upload_parameter__save_file_content_in_db': indica se o conteúdo do arquivo deve ser salvo no banco de dados (binário).
		///		'upload_parameter__save_file_content_in_db_as_text': indica se o conteúdo do arquivo deve ser salvo no banco de dados como texto.
		///
		/// Método para ser usado no upload de arquivos para o servidor.
		/// Utiliza a seguinte estrutura de diretórios:
		///		UploadedFiles
		///			StoredFiles
		///			BackupRecentFiles
		///			TemporaryFiles
		///			Temp
		///		
		/// Diretório 'StoredFiles': armazena os arquivos que devem ser preservados por tempo indefinido. Para facilitar o backup desses arquivos, essa pasta
		///		possui sub-pastas cujos nomes são no formato YYYY-MM e, dentro destas, há sub-pastas que são definidas pelo parâmetro 'upload_parameter__folder_name'.
		///		Caso este parâmetro não tenha sido informado, será armazenado na pasta 'Default'.
		///		
		/// Diretório 'BackupRecentFiles': utiliza a mesma estrutura interna da pasta 'StoredFiles', porém, mantém os arquivos somente pelo tempo necessário para
		///		a realização do backup mensal dos arquivos. Esta pasta é utilizada para que a rotina de backup diário salve os arquivos mais recentes. A gravação da
		///		cópia de backup depende do parâmetro 'WebAPI_UploadFile_FlagHabilitacao_BackupRecentFiles' gravado em t_PARAMETRO
		///		
		/// Diretório 'TemporaryFiles': pasta usada para fazer upload de arquivos usados em atividades curtas e que não necessitam que os arquivos sejam preservados.
		///		O conteúdo desta pasta é limpo diariamente.
		///		
		/// Diretório 'Temp': é usado para gravar automaticamente os arquivos recebidos na operação de upload (ex: BodyPart_2ac172ce-b0b1-4c17-88b8-bf78d13142a6)
		/// </summary>
		/// <returns>
		///		Retorna uma resposta JSON com o objeto 'UploadFileResponse' e as seguintes informações:
		///			'Status'
		///			Array de itens com informações de cada arquivo:
		///				'field_name': nome do elemento html usado no upload do arquivo.
		///				'original_file_name': nome original do arquivo no computador de origem.
		///				'stored_file_name': nome usado para armazenar o arquivo no servidor (formato: yyyyMMdd_HHmmss_fff__GUID.ext).
		///				'folder_name': pasta (relativa) em que o arquivo foi armazenado.
		///				'stored_file_guid': identificador do arquivo no formato GUID para ser usado em consultas posteriores.
		/// </returns>
		[HttpPost]
		public async Task<HttpResponseMessage> PostFile()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "UploadFile.PostFile()";
			Guid httpRequestId = Request.GetCorrelationId();
			int qtdeTentativas;
			int intParametro;
			int id_module_folder_name;
			int maxSizeInBytesToSaveFileContentInDb = 0;
			int maxSizeInCharsToSaveFileContentInDbAsText = 0;
			long localFileSizeBytes;
			long localFileSizeChars = 0;
			bool blnTempFile = false;
			bool blnConfirmationRequired = false;
			bool blnSaveFileContentInDb = false;
			bool blnSaveFileContentInDbAsText = false;
			bool blnBackupRecentFiles;
			string s;
			string msg;
			string msg_erro_aux;
			string fileName = "";
			string fullFileName = "";
			string fullBackupFileName;
			string sModuleFolderName;
			string sParamFolderName;
			string sParamValue;
			string sUserId;
			string sessionToken;
			string sGuid;
			string root;
			string dirDestRoot;
			string dirBackupRoot;
			string dirTemp;
			string dirDestTemporaryRelativo;
			string dirDestTemporary;
			string dirDestRelativo;
			string dirDestFull;
			string dirDestFullAux;
			string dirBackupFull;
			Usuario usuarioBD;
			UploadFileItemResponse uploadFileItem;
			UploadStoredFileInfo storedFileInfo;
			UploadFileModuleFolderName uploadFileModuleFolderName;
			UploadFileModuleFolderName insertUploadFileModuleFolderName;
			List<UploadFileItemResponse> vUploadFileItem = new List<UploadFileItemResponse>();
			UploadFileResponse uploadFileResponse = new UploadFileResponse();
			HttpResponseMessage result;
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida";
				Global.gravaLogAtividade(httpRequestId, msg);

				// Check if the request contains multipart/form-data.
				if (!Request.Content.IsMimeMultipartContent())
				{
					throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
				}

				#region [ Parâmetro referente à gravação da cópia de backup ]
				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.FLAG_HABILITACAO_UPLOAD_FILE_BACKUP_RECENT_FILES);
				blnBackupRecentFiles = (intParametro != 0) ? true : false;
				#endregion

				#region [ Prepara diretórios (nível base) ]
				root = HttpContext.Current.Server.MapPath("~/UploadedFiles");
				if (!Directory.Exists(root)) Directory.CreateDirectory(root);

				dirDestRoot = Path.Combine(root, "StoredFiles");
				if (!Directory.Exists(dirDestRoot)) Directory.CreateDirectory(dirDestRoot);

				dirBackupRoot = Path.Combine(root, "BackupRecentFiles");
				if (blnBackupRecentFiles)
				{
					if (!Directory.Exists(dirBackupRoot)) Directory.CreateDirectory(dirBackupRoot);
				}

				dirDestTemporaryRelativo = "TemporaryFiles";
				dirDestTemporary = Path.Combine(root, dirDestTemporaryRelativo);
				if (!Directory.Exists(dirDestTemporary)) Directory.CreateDirectory(dirDestTemporary);

				dirTemp = Path.Combine(root, "Temp");
				if (!Directory.Exists(dirTemp)) Directory.CreateDirectory(dirTemp);
				#endregion

				#region [ Prepara diretórios (nível específico) ]
				dirDestRelativo = DateTime.Now.ToString("yyyy-MM");

				dirDestFull = Path.Combine(dirDestRoot, dirDestRelativo);
				if (!Directory.Exists(dirDestFull)) Directory.CreateDirectory(dirDestFull);

				dirBackupFull = Path.Combine(dirBackupRoot, dirDestRelativo);
				if (blnBackupRecentFiles)
				{
					if (!Directory.Exists(dirBackupFull)) Directory.CreateDirectory(dirBackupFull);
				}
				#endregion

				var provider = new MultipartFormDataStreamProvider(dirTemp);

				// Read the form data and return an async task.
				// Obs: os arquivos serão gravados automaticamente no diretório especificado na criação do objeto 'provider' (ex: BodyPart_2ac172ce-b0b1-4c17-88b8-bf78d13142a6)
				await Request.Content.ReadAsMultipartAsync(provider);

				try // Finally: assegura que todos os arquivos temporários gerados pelo 'provider' foram apagados
				{
					#region [ Obtém os parâmetros ]

					#region [ Parâmetro: upload_parameter__user_id ]
					sUserId = "";
					sParamValue = provider.FormData.Get("upload_parameter__user_id");
					if ((sParamValue ?? "").Length > 0) sUserId = sParamValue;
					if (sUserId.Length == 0) throw new WebApiException("Não foi informado a identificação do usuário responsável pela operação!");
					#endregion

					#region [ Parâmetro: upload_parameter__sessionToken ]
					sessionToken = "";
					sParamValue = provider.FormData.Get("upload_parameter__sessionToken");
					if ((sParamValue ?? "").Length > 0) sessionToken = sParamValue;
					if (sessionToken.Length == 0) throw new WebApiException("Não foi informado o token da sessão do usuário!");
					#endregion

					#region [ Validação de segurança: session token confere? ]
					usuarioBD = GeralDAO.getUsuario(sUserId, out msg_erro_aux);
					if (usuarioBD == null)
					{
						throw new WebApiException("Falha ao tentar validar usuário!");
					}

					if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
					{
						throw new WebApiException("Token de sessão inválido!");
					}
					#endregion

					#region [ Parâmetro: upload_parameter__folder_name ]
					sParamFolderName = "";
					sParamValue = provider.FormData.Get("upload_parameter__folder_name");
					if ((sParamValue ?? "").Length > 0)
					{
						try
						{
							dirDestFullAux = Path.Combine(dirDestFull, sParamValue);
							s = Path.GetFullPath(dirDestFullAux);
							if ((s ?? "").Length > 0)
							{
								sParamFolderName = sParamValue;
							}
						}
						catch (Exception ex)
						{
							Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + " - Exception: " + ex.Message);
							// Path.Combine() e Path.GetFullPath() throw exceptions if the path is invalid
							throw new WebApiException("O nome da pasta informado é inválido!\n" + ex.Message);
						}
					}

					if (sParamFolderName.Length == 0) sParamFolderName = "Default";
					sModuleFolderName = sParamFolderName;
					dirDestRelativo = Path.Combine(dirDestRelativo, sParamFolderName);
					dirDestFull = Path.Combine(dirDestRoot, dirDestRelativo);
					if (!Directory.Exists(dirDestFull)) Directory.CreateDirectory(dirDestFull);
					dirBackupFull = Path.Combine(dirBackupRoot, dirDestRelativo);
					if (blnBackupRecentFiles)
					{
						if (!Directory.Exists(dirBackupFull)) Directory.CreateDirectory(dirBackupFull);
					}
					#endregion

					#region [ Parâmetro: upload_parameter__is_temp_file ]
					sParamValue = provider.FormData.Get("upload_parameter__is_temp_file");
					if (sParamValue != null)
					{
						sParamValue = sParamValue.ToUpper();
						if (sParamValue.Equals("1") || sParamValue.Equals("S") || sParamValue.Equals("SIM") || sParamValue.Equals("Y") || sParamValue.Equals("YES")) blnTempFile = true;
					}
					#endregion

					#region [ Parâmetro: upload_parameter__is_confirmation_required ]
					sParamValue = provider.FormData.Get("upload_parameter__is_confirmation_required");
					if (sParamValue != null)
					{
						sParamValue = sParamValue.ToUpper();
						if (sParamValue.Equals("1") || sParamValue.Equals("S") || sParamValue.Equals("SIM") || sParamValue.Equals("Y") || sParamValue.Equals("YES")) blnConfirmationRequired = true;
					}
					#endregion

					#region [ Parâmetro: upload_parameter__save_file_content_in_db ]
					sParamValue = provider.FormData.Get("upload_parameter__save_file_content_in_db");
					if (sParamValue != null)
					{
						sParamValue = sParamValue.ToUpper();
						if (sParamValue.Equals("1") || sParamValue.Equals("S") || sParamValue.Equals("SIM") || sParamValue.Equals("Y") || sParamValue.Equals("YES")) blnSaveFileContentInDb = true;
					}

					if (blnSaveFileContentInDb)
					{
						maxSizeInBytesToSaveFileContentInDb = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.UPLOAD_FILE_SAVE_FILE_CONTENT_IN_DB_MAX_SIZE_IN_BYTES);
					}
					#endregion

					#region [ Parâmetro: upload_parameter__save_file_content_in_db_as_text ]
					sParamValue = provider.FormData.Get("upload_parameter__save_file_content_in_db_as_text");
					if (sParamValue != null)
					{
						sParamValue = sParamValue.ToUpper();
						if (sParamValue.Equals("1") || sParamValue.Equals("S") || sParamValue.Equals("SIM") || sParamValue.Equals("Y") || sParamValue.Equals("YES")) blnSaveFileContentInDbAsText = true;
					}

					if (blnSaveFileContentInDbAsText)
					{
						maxSizeInCharsToSaveFileContentInDbAsText = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.UPLOAD_FILE_SAVE_FILE_CONTENT_IN_DB_AS_TEXT_MAX_SIZE_IN_CHARS);
					}
					#endregion

					#endregion

					#region [ Tratamento para associar o 'folder_name' ao módulo ao qual pertence(m) o(s) arquivo(s) ]
					uploadFileModuleFolderName = UploadFileDAO.getUploadFileModuleFolderName(sModuleFolderName, out msg_erro_aux);
					if (uploadFileModuleFolderName != null)
					{
						id_module_folder_name = uploadFileModuleFolderName.id;
					}
					else
					{
						insertUploadFileModuleFolderName = new UploadFileModuleFolderName();
						insertUploadFileModuleFolderName.module_folder_name = sModuleFolderName;
						UploadFileDAO.insertUploadFileModuleFolderName(httpRequestId, insertUploadFileModuleFolderName, out msg_erro_aux);
						id_module_folder_name = insertUploadFileModuleFolderName.id;
					}
					#endregion

					#region [ Laço para renomear o nome de cada arquivo salvo ]
					foreach (MultipartFileData fileData in provider.FileData)
					{
						// Se no formulário houver vários campos input (type=file) e algum deles não estiver preenchido, o conteúdo de 'FileName' será duas aspas duplas
						if (fileData.Headers.ContentDisposition.FileName.Replace("\"", "").Length > 0)
						{
							// Remove as aspas duplas que envolvem o nome do arquivo (no início e no final)
							FileInfo fileInfo = new FileInfo(fileData.Headers.ContentDisposition.FileName.Replace("\"", ""));

							#region [ Cria o nome do arquivo a ser usado na gravação no servidor ]
							qtdeTentativas = 0;
							while (true)
							{
								qtdeTentativas++;
								sGuid = BD.gera_uid(httpRequestId);
								if ((sGuid ?? "").Trim().Length > 0)
								{
									fileName = System.DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "__" + sGuid + Path.GetExtension(fileInfo.Name);
									fullFileName = (blnTempFile ? Path.Combine(dirDestTemporary, fileName) : Path.Combine(dirDestFull, fileName));
									// Verifica se o nome já está em uso por outro arquivo
									if (!File.Exists(fullFileName)) break;
								}
								if (qtdeTentativas > 20) throw new WebApiException("Falha ao tentar definir o nome do arquivo para armazenamento no servidor!");
							}
							#endregion

							#region [ Obtém o tamanho do arquivo (binário) ]
							// Quando o upload é feito pelo Google Chrome, fileInfo.Length lança um exception do tipo System.IO.FileNotFoundException (ex: Não foi possível localizar o arquivo 'nome_arquivo.txt'.)
							localFileSizeBytes = new System.IO.FileInfo(fileData.LocalFileName).Length;
							#endregion

							#region [ Obtém o tamanho do arquivo (texto) ]
							if (blnSaveFileContentInDbAsText)
							{
								try
								{
									localFileSizeChars = (File.ReadAllText(fileData.LocalFileName)).Length;
								}
								catch (Exception ex)
								{
									s = NOME_DESTA_ROTINA + " - Exception ao tentar ler o conteúdo do arquivo como texto (arquivo: " + fileInfo.Name + "): " + ex.Message;
									Global.gravaLogAtividade(httpRequestId, s);
									throw new WebApiException("Falha ao tentar ler o conteúdo do arquivo como texto (" + fileInfo.Name + ")");
								}
							}
							#endregion

							#region [ Caso a opção de salvar o arquivo no BD esteja ativa, verifica tamanho máximo ]
							if (blnSaveFileContentInDb && (maxSizeInBytesToSaveFileContentInDb > 0))
							{
								if (localFileSizeBytes > maxSizeInBytesToSaveFileContentInDb)
								{
									s = NOME_DESTA_ROTINA + ": arquivo " + fileInfo.Name + " excede o limite máximo permitido para salvar no banco de dados (" + maxSizeInBytesToSaveFileContentInDb.ToString() + " bytes)";
									Global.gravaLogAtividade(httpRequestId, s);
									throw new WebApiException("O tamanho do arquivo excede o limite máximo permitido para salvar no banco de dados (" + maxSizeInBytesToSaveFileContentInDb.ToString() + " bytes)!");
								}
							}

							if (blnSaveFileContentInDbAsText && (maxSizeInCharsToSaveFileContentInDbAsText > 0))
							{
								if (localFileSizeChars > maxSizeInCharsToSaveFileContentInDbAsText)
								{
									s = NOME_DESTA_ROTINA + ": arquivo " + fileInfo.Name + " excede o limite máximo permitido para salvar no banco de dados como texto (" + maxSizeInCharsToSaveFileContentInDbAsText.ToString() + " caracteres)";
									Global.gravaLogAtividade(httpRequestId, s);
									throw new WebApiException("O tamanho do arquivo excede o limite máximo permitido para salvar no banco de dados como texto (" + maxSizeInCharsToSaveFileContentInDbAsText.ToString() + " caracteres)!");
								}
							}
							#endregion

							// Copia o arquivo para o destino final
							File.Copy(fileData.LocalFileName, fullFileName);

							msg = NOME_DESTA_ROTINA + ": Arquivo salvo: " + fileData.LocalFileName + " => " + fullFileName + " (nome original: " + fileInfo.FullName + ")";
							Global.gravaLogAtividade(httpRequestId, msg);

							if (!blnTempFile)
							{
								fullBackupFileName = Path.Combine(dirBackupFull, fileName);
								if (blnBackupRecentFiles)
								{
									File.Copy(fileData.LocalFileName, fullBackupFileName);

									msg = NOME_DESTA_ROTINA + ": Arquivo salvo: " + fileData.LocalFileName + " => " + fullBackupFileName + " (nome original: " + fileInfo.FullName + ")";
									Global.gravaLogAtividade(httpRequestId, msg);
								}
							}

							#region [ Dados para serem retornados na resposta ]
							uploadFileItem = new UploadFileItemResponse();
							uploadFileItem.field_name = fileData.Headers.ContentDisposition.Name.Replace(((char)34).ToString(), string.Empty);
							uploadFileItem.original_file_name = fileInfo.Name;
							uploadFileItem.stored_file_name = fileName;
							uploadFileItem.folder_name = (blnTempFile ? dirDestTemporaryRelativo : dirDestRelativo);
							#endregion

							#region [ Grava as informações do arquivo transferido na tabela t_UPLOAD_FILE ]
							storedFileInfo = new UploadStoredFileInfo();
							storedFileInfo.guid = sGuid;
							storedFileInfo.usuario_cadastro = sUserId;
							storedFileInfo.st_temporary_file = (byte)(blnTempFile ? 1 : 0);
							storedFileInfo.st_confirmation_required = (byte)(blnConfirmationRequired ? 1 : 0);
							storedFileInfo.original_file_name = fileInfo.Name;
							storedFileInfo.original_full_file_name = fileInfo.FullName;
							storedFileInfo.stored_file_name = fileName;
							storedFileInfo.stored_full_file_name = fullFileName;
							storedFileInfo.stored_relative_path = (blnTempFile ? dirDestTemporaryRelativo : dirDestRelativo);
							storedFileInfo.id_module_folder_name = id_module_folder_name;
							// Quando o upload é feito pelo Google Chrome, fileInfo.Length lança um exception do tipo System.IO.FileNotFoundException (ex: Não foi possível localizar o arquivo 'nome_arquivo.txt'.)
							storedFileInfo.file_size = localFileSizeBytes;
							storedFileInfo.remote_IP = GetIp();

							#region [ Conteúdo do arquivo (binário) ]
							if (blnSaveFileContentInDb)
							{
								using (var stream = new FileStream(fullFileName, FileMode.Open, FileAccess.Read))
								{
									using (var reader = new BinaryReader(stream))
									{
										storedFileInfo.file_content = reader.ReadBytes((int)stream.Length);
									}
								}
							}
							#endregion

							#region [ Conteúdo do arquivo (texto) ]
							if (blnSaveFileContentInDbAsText)
							{
								storedFileInfo.file_content_text = File.ReadAllText(fullFileName);
							}
							#endregion

							if (UploadFileDAO.insertUploadedFileInfo(httpRequestId, storedFileInfo, out msg_erro_aux))
							{
								uploadFileItem.stored_file_guid = storedFileInfo.guid;
							}
							else
							{
								throw new WebApiException("Falha ao tentar salvar informações sobre o arquivo no banco de dados!");
							}
							#endregion

							#region [ Armazena os dados para serem retornados na resposta ]
							vUploadFileItem.Add(uploadFileItem);
							#endregion
						}

						// Apaga o arquivo gravado automaticamente
						File.Delete(fileData.LocalFileName);
					}
					#endregion
				}
				finally
				{
					#region [ Garante que os arquivos gerados automaticamente serão todos excluídos ]
					foreach (MultipartFileData fileData in provider.FileData)
					{
						if (File.Exists(fileData.LocalFileName)) File.Delete(fileData.LocalFileName);
					}
					#endregion
				}

				#region [ Monta resposta ]
				if (vUploadFileItem.Count == 0)
				{
					#region [ Nenhum arquivo foi transferido ]
					uploadFileResponse.Status = "ERROR";
					uploadFileResponse.Message = "Nenhum arquivo foi transferido!";
					var jsonError = new JavaScriptSerializer().Serialize(uploadFileResponse);
					result = Request.CreateResponse(HttpStatusCode.OK);
					result.Content = new StringContent(jsonError, Encoding.UTF8, "text/html");
					#endregion
				}
				else
				{
					#region [ Monta resposta de sucesso ]
					uploadFileResponse.Status = "OK";
					uploadFileResponse.files = vUploadFileItem.ToArray();

					var jsonOk = new JavaScriptSerializer().Serialize(uploadFileResponse);
					result = Request.CreateResponse(HttpStatusCode.OK);
					result.Content = new StringContent(jsonOk, Encoding.UTF8, "text/html");
					#endregion
				}
				#endregion

				msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return result;
			}
			catch (System.Exception e)
			{
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + " - Exception: " + e.Message);

				if (e is WebApiException)
				{
					uploadFileResponse.Status = "ERROR";
					uploadFileResponse.Message = e.Message;
					var jsonError = new JavaScriptSerializer().Serialize(uploadFileResponse);
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
}
