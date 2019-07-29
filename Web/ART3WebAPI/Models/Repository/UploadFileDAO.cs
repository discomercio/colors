using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Threading;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;

namespace ART3WebAPI.Models.Repository
{
	public class UploadFileDAO
	{
		#region [ uploadStoredFileInfoLoadFromDataRow ]
		public static UploadStoredFileInfo uploadStoredFileInfoLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			UploadStoredFileInfo fileInfo = new UploadStoredFileInfo();
			#endregion

			fileInfo.id = BD.readToInt(rowDados["id"]);
			fileInfo.guid = BD.readToString(rowDados["guid"]);
			fileInfo.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			fileInfo.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			fileInfo.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			fileInfo.st_temporary_file = BD.readToByte(rowDados["st_temporary_file"]);
			fileInfo.st_confirmation_required = BD.readToByte(rowDados["st_confirmation_required"]);
			fileInfo.original_file_name = BD.readToString(rowDados["original_file_name"]);
			fileInfo.original_full_file_name = BD.readToString(rowDados["original_full_file_name"]);
			fileInfo.stored_file_name = BD.readToString(rowDados["stored_file_name"]);
			fileInfo.stored_full_file_name = BD.readToString(rowDados["stored_full_file_name"]);
			fileInfo.stored_relative_path = BD.readToString(rowDados["stored_relative_path"]);
			fileInfo.id_module_folder_name = BD.readToInt(rowDados["id_module_folder_name"]);
			fileInfo.file_size = BD.readToInt64(rowDados["file_size"]);
			fileInfo.remote_IP = BD.readToString(rowDados["remote_IP"]);
			fileInfo.file_content = BD.readToVarBinary(rowDados["file_content"]);
			fileInfo.file_content_text = BD.readToString(rowDados["file_content_text"]);

			return fileInfo;
		}
		#endregion

		#region [ getUploadStoredFileInfoById ]
		public static UploadStoredFileInfo getUploadStoredFileInfoById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			UploadStoredFileInfo fileInfo;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if (id <= 0)
				{
					msg_erro = "O ID informado é inválido!";
					return null;
				}
				#endregion

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_UPLOAD_FILE" +
							" WHERE" +
								" (id = " + id.ToString() + ")";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Nenhum registro localizado com o ID = " + id.ToString();
						return null;
					}

					fileInfo = uploadStoredFileInfoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return fileInfo;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getUploadStoredFileInfoByGuid ]
		public static UploadStoredFileInfo getUploadStoredFileInfoByGuid(string guid, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			UploadStoredFileInfo fileInfo;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((guid ?? "").Length == 0)
				{
					msg_erro = "O GUID informado é inválido!";
					return null;
				}
				#endregion

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_UPLOAD_FILE" +
							" WHERE" +
								" (guid = '" + guid + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Nenhum registro localizado com o GUID = " + guid;
						return null;
					}

					fileInfo = uploadStoredFileInfoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return fileInfo;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ insertUploadedFileInfo ]
		public static bool insertUploadedFileInfo(UploadStoredFileInfo storedFileInfo, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "UploadFileDAO.insertUploadedFileInfo()";
			const int TAMANHO_CAMPO_USUARIO_CADASTRO = 20;
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string strSql;
			string msg_erro_aux = "";
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			UploadStoredFileInfo fileInfoBD;
			Log log;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_UPLOAD_FILE (" +
								"guid, " +
								"usuario_cadastro, " +
								"st_temporary_file, " +
								"st_confirmation_required, " +
								"original_file_name, " +
								"original_full_file_name, " +
								"stored_file_name, " +
								"stored_full_file_name, " +
								"stored_relative_path, " +
								"id_module_folder_name, " +
								"file_size, " +
								"remote_IP, " +
								"file_content, " +
								"file_content_text" +
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@guid, " +
								"@usuario_cadastro, " +
								"@st_temporary_file, " +
								"@st_confirmation_required, " +
								"@original_file_name, " +
								"@original_full_file_name, " +
								"@stored_file_name, " +
								"@stored_full_file_name, " +
								"@stored_relative_path, " +
								"@id_module_folder_name, " +
								"@file_size, " +
								"@remote_IP, " +
								"@file_content, " +
								"@file_content_text" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@guid", SqlDbType.UniqueIdentifier);
					cmInsert.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, TAMANHO_CAMPO_USUARIO_CADASTRO);
					cmInsert.Parameters.Add("@st_temporary_file", SqlDbType.TinyInt);
					cmInsert.Parameters.Add("@st_confirmation_required", SqlDbType.TinyInt);
					cmInsert.Parameters.Add("@original_file_name", SqlDbType.VarChar, 256);
					cmInsert.Parameters.Add("@original_full_file_name", SqlDbType.VarChar, 1024);
					cmInsert.Parameters.Add("@stored_file_name", SqlDbType.VarChar, 256);
					cmInsert.Parameters.Add("@stored_full_file_name", SqlDbType.VarChar, 1024);
					cmInsert.Parameters.Add("@stored_relative_path", SqlDbType.VarChar, 1024);
					cmInsert.Parameters.Add("@id_module_folder_name", SqlDbType.Int);
					cmInsert.Parameters.Add("@file_size", SqlDbType.BigInt);
					cmInsert.Parameters.Add("@remote_IP", SqlDbType.VarChar, 50);
					cmInsert.Parameters.Add("@file_content", SqlDbType.VarBinary, -1); // varbinary(max)
					cmInsert.Parameters.Add("@file_content_text", SqlDbType.NVarChar, -1); // nvarchar(max)
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							if ((storedFileInfo.guid ?? "").Trim().Length == 0) storedFileInfo.guid = BD.gera_uid();

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@guid"].Value = new Guid(storedFileInfo.guid);
							cmInsert.Parameters["@usuario_cadastro"].Value = Global.leftStr((storedFileInfo.usuario_cadastro ?? ""), TAMANHO_CAMPO_USUARIO_CADASTRO);
							cmInsert.Parameters["@st_temporary_file"].Value = storedFileInfo.st_temporary_file;
							cmInsert.Parameters["@st_confirmation_required"].Value = storedFileInfo.st_confirmation_required;
							cmInsert.Parameters["@original_file_name"].Value = (storedFileInfo.original_file_name ?? "");
							cmInsert.Parameters["@original_full_file_name"].Value = (storedFileInfo.original_full_file_name ?? "");
							cmInsert.Parameters["@stored_file_name"].Value = (storedFileInfo.stored_file_name ?? "");
							cmInsert.Parameters["@stored_full_file_name"].Value = (storedFileInfo.stored_full_file_name ?? "");
							cmInsert.Parameters["@stored_relative_path"].Value = (storedFileInfo.stored_relative_path ?? "");
							cmInsert.Parameters["@id_module_folder_name"].Value = storedFileInfo.id_module_folder_name;
							cmInsert.Parameters["@file_size"].Value = storedFileInfo.file_size;
							cmInsert.Parameters["@remote_IP"].Value = (storedFileInfo.remote_IP ?? "");
							cmInsert.Parameters["@file_content"].Value = (storedFileInfo.file_content ?? Convert.DBNull);
							cmInsert.Parameters["@file_content_text"].Value = (storedFileInfo.file_content_text ?? Convert.DBNull);
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								storedFileInfo.id = generatedId;
							}
							catch (Exception ex)
							{
								generatedId = 0;
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception:\n" + ex.ToString());
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								fileInfoBD = getUploadStoredFileInfoById(generatedId, out msg_erro_aux);
								storedFileInfo.guid = fileInfoBD.guid;
								storedFileInfo.dt_cadastro = fileInfoBD.dt_cadastro;
								storedFileInfo.dt_hr_cadastro = fileInfoBD.dt_hr_cadastro;

								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion

						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Grava o log ]
						if (blnSucesso)
						{
							log = new Log();
							log.usuario = storedFileInfo.usuario_cadastro;
							log.operacao = "UploadFile";
							log.complemento = sbLog.ToString();
							LogDAO.insere(storedFileInfo.usuario_cadastro, log, msg_erro_aux);
						}
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados as informações sobre o arquivo transferido para o servidor após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ uploadFileModuleFolderNameLoadFromDataRow ]
		public static UploadFileModuleFolderName uploadFileModuleFolderNameLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			UploadFileModuleFolderName uploadFileModuleFolderNameLoadFromDataRow = new UploadFileModuleFolderName();
			#endregion

			uploadFileModuleFolderNameLoadFromDataRow.id = BD.readToInt(rowDados["id"]);
			uploadFileModuleFolderNameLoadFromDataRow.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			uploadFileModuleFolderNameLoadFromDataRow.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			uploadFileModuleFolderNameLoadFromDataRow.module_folder_name = BD.readToString(rowDados["module_folder_name"]);

			return uploadFileModuleFolderNameLoadFromDataRow;
		}
		#endregion

		#region [ getUploadFileModuleFolderName ]
		public static UploadFileModuleFolderName getUploadFileModuleFolderName(string module_folder_name, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			UploadFileModuleFolderName uploadFileModuleFolderName;
			#endregion

			msg_erro = "";
			try
			{
				if ((module_folder_name ?? "").Trim().Length == 0)
				{
					msg_erro = "Nome da pasta não foi fornecida!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_UPLOAD_FILE_MODULE_FOLDER_NAME" +
							" WHERE" +
								" (module_folder_name = @module_folder_name)";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					cmCommand.Parameters.Add("@module_folder_name", SqlDbType.VarChar, 512);
					cmCommand.Parameters["@module_folder_name"].Value = (module_folder_name ?? "");
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Registro não encontrado!";
						return null;
					}

					uploadFileModuleFolderName = uploadFileModuleFolderNameLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return uploadFileModuleFolderName;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ insertUploadFileModuleFolderName ]
		public static bool insertUploadFileModuleFolderName(UploadFileModuleFolderName uploadFileModuleFolderName, out string msg_erro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string strSql;
			string msg_erro_aux = "";
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			Log log;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_UPLOAD_FILE_MODULE_FOLDER_NAME (" +
								"module_folder_name" +
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@module_folder_name" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@module_folder_name", SqlDbType.VarChar, 512);
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@module_folder_name"].Value = (uploadFileModuleFolderName.module_folder_name ?? "");
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								uploadFileModuleFolderName.id = generatedId;
							}
							catch (Exception)
							{
								generatedId = 0;
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion

						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Grava o log ]
						if (blnSucesso)
						{
							log = new Log();
							log.usuario = Global.Cte.Usuario.ID_USUARIO_SISTEMA;
							log.operacao = "UpFileModFolderName";
							log.complemento = sbLog.ToString();
							LogDAO.insere(Global.Cte.Usuario.ID_USUARIO_SISTEMA, log, msg_erro_aux);
						}
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados o registro em t_UPLOAD_FILE_MODULE_FOLDER_NAME após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion
	}
}