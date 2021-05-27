using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Domains;
using System.Threading;

namespace ART3WebAPI.Models.Repository
{
	public class GeralDAO
	{
		#region [ criptografaTexto ]
		public static bool criptografaTexto(string textoDecriptografado, out string textoCriptografado, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			textoCriptografado = "";
			msg_erro = "";
			try
			{
				if (textoDecriptografado == null) return true;
				if ((textoDecriptografado ?? "").Trim().Length == 0) return true;

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT dbo.SqlClrUtilCodificaTextoHex(@textoDecriptografado) AS textoCriptografado";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@textoDecriptografado", SqlDbType.VarChar, -1); // varchar(max)
					cmSelect.Prepare();
					cmSelect.Parameters["@textoDecriptografado"].Value = (textoDecriptografado ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Falha ao tentar criptografar o texto!";
						return false;
					}

					textoCriptografado = BD.readToString(dtbResultado.Rows[0][0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ decriptografaTexto ]
		public static bool decriptografaTexto(string textoCriptografado, out string textoDecriptografado, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			textoDecriptografado = "";
			msg_erro = "";
			try
			{
				if (textoCriptografado == null) return true;
				if ((textoCriptografado ?? "").Trim().Length == 0) return true;

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT dbo.SqlClrUtilDecodificaTextoHex(@textoCriptografado) AS textoDecriptografado";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@textoCriptografado", SqlDbType.VarChar, -1); // varchar(max)
					cmSelect.Prepare();
					cmSelect.Parameters["@textoCriptografado"].Value = (textoCriptografado ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Falha ao tentar decriptografar o texto!";
						return false;
					}

					textoDecriptografado = BD.readToString(dtbResultado.Rows[0][0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ codigoDescricaoLoadFromDataRow ]
		public static CodigoDescricao codigoDescricaoLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			CodigoDescricao codigoDescricao = new CodigoDescricao();
			#endregion

			codigoDescricao.grupo = BD.readToString(rowDados["grupo"]);
			codigoDescricao.codigo = BD.readToString(rowDados["codigo"]);
			codigoDescricao.ordenacao = BD.readToInt(rowDados["ordenacao"]);
			codigoDescricao.st_inativo = BD.readToByte(rowDados["st_inativo"]);
			codigoDescricao.descricao = BD.readToString(rowDados["descricao"]);
			codigoDescricao.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			codigoDescricao.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			codigoDescricao.dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["dt_hr_ult_atualizacao"]);
			codigoDescricao.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);
			codigoDescricao.st_possui_sub_codigo = BD.readToByte(rowDados["st_possui_sub_codigo"]);
			codigoDescricao.st_eh_sub_codigo = BD.readToByte(rowDados["st_eh_sub_codigo"]);
			codigoDescricao.grupo_pai = BD.readToString(rowDados["grupo_pai"]);
			codigoDescricao.codigo_pai = BD.readToString(rowDados["codigo_pai"]);
			codigoDescricao.lojas_habilitadas = BD.readToString(rowDados["lojas_habilitadas"]);
			codigoDescricao.parametro_1_campo_flag = BD.readToByte(rowDados["parametro_1_campo_flag"]);
			codigoDescricao.parametro_2_campo_flag = BD.readToByte(rowDados["parametro_2_campo_flag"]);
			codigoDescricao.parametro_3_campo_flag = BD.readToByte(rowDados["parametro_3_campo_flag"]);
			codigoDescricao.parametro_4_campo_flag = BD.readToByte(rowDados["parametro_4_campo_flag"]);
			codigoDescricao.parametro_5_campo_flag = BD.readToByte(rowDados["parametro_5_campo_flag"]);
			codigoDescricao.parametro_campo_inteiro = BD.readToInt(rowDados["parametro_campo_inteiro"]);
			codigoDescricao.parametro_campo_monetario = BD.readToDecimal(rowDados["parametro_campo_monetario"]);
			codigoDescricao.parametro_campo_real = BD.readToSingle(rowDados["parametro_campo_real"]);
			codigoDescricao.parametro_campo_data = BD.readToDateTime(rowDados["parametro_campo_data"]);
			codigoDescricao.parametro_campo_texto = BD.readToString(rowDados["parametro_campo_texto"]);
			codigoDescricao.parametro_2_campo_texto = BD.readToString(rowDados["parametro_2_campo_texto"]);
			codigoDescricao.parametro_3_campo_texto = BD.readToString(rowDados["parametro_3_campo_texto"]);
			codigoDescricao.descricao_parametro = BD.readToString(rowDados["descricao_parametro"]);

			return codigoDescricao;
		}
		#endregion

		#region [ getCodigoDescricao ]
		public static CodigoDescricao getCodigoDescricao(string grupo, string codigo, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			CodigoDescricao codigoDescricao;
			#endregion

			msg_erro = "";
			try
			{
				if ((grupo ?? "").Trim().Length == 0)
				{
					msg_erro = "Identificação do grupo não foi informado!";
					return null;
				}

				if ((codigo ?? "").Trim().Length == 0)
				{
					msg_erro = "Identificação do código não foi informado!";
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
							" FROM t_CODIGO_DESCRICAO" +
							" WHERE" +
								" (grupo = '" + grupo + "')" +
								" AND (codigo = '" + codigo + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Código não encontrado!";
						return null;
					}

					codigoDescricao = codigoDescricaoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return codigoDescricao;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getCodigoDescricaoByGrupo ]
		public static List<CodigoDescricao> getCodigoDescricaoByGrupo(string grupo, Global.eFiltroFlagStInativo st_inativo, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			string strWhere;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			CodigoDescricao codigoDescricao;
			List<CodigoDescricao> listaCodigoDescricao = new List<CodigoDescricao>();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta cláusula Where ]
					strWhere = "";
					if ((grupo ?? "").Trim().Length > 0)
					{
						if (strWhere.Length > 0) strWhere += " AND";
						strWhere += " (grupo = '" + grupo + "')";
					}
					if (st_inativo != Global.eFiltroFlagStInativo.FLAG_IGNORADO)
					{
						if (strWhere.Length > 0) strWhere += " AND";
						strWhere += " (st_inativo = " + st_inativo.ToString() + ")";
					}

					if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;
					#endregion

					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_CODIGO_DESCRICAO" +
							strWhere +
							" ORDER BY" +
								" grupo," +
								" ordenacao";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						codigoDescricao = codigoDescricaoLoadFromDataRow(dtbResultado.Rows[i]);
						listaCodigoDescricao.Add(codigoDescricao);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaCodigoDescricao;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ usuarioLoadFromDataRow ]
		public static Usuario usuarioLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			Usuario usuario = new Usuario();
			#endregion

			usuario.usuario = BD.readToString(rowDados["usuario"]);
			usuario.nome = BD.readToString(rowDados["nome"]);
			usuario.datastamp = BD.readToString(rowDados["datastamp"]);

			if (rowDados["bloqueado"].ToString().Equals("0"))
				usuario.bloqueado = false;
			else
				usuario.bloqueado = true;

			if (rowDados["dt_ult_alteracao_senha"] == DBNull.Value)
				usuario.senhaExpirada = true;
			else
				usuario.senhaExpirada = false;

			usuario.SessionTokenModuloCentral = BD.readToString(rowDados["SessionTokenModuloCentral"]);
			usuario.SessionTokenModuloLoja = BD.readToString(rowDados["SessionTokenModuloLoja"]);

			return usuario;
		}
		#endregion

		#region [ getUsuario ]
		public static Usuario getUsuario(string id_usuario, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			Usuario usuario;
			#endregion

			msg_erro = "";
			try
			{
				if ((id_usuario ?? "").Trim().Length == 0)
				{
					msg_erro = "Identificação do usuário não foi fornecida!";
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
					strSql = "SELECT" +
								" usuario," +
								" nome," +
								" datastamp," +
								" bloqueado," +
								" dt_ult_alteracao_senha," +
								" Convert(varchar(36), SessionTokenModuloCentral) AS SessionTokenModuloCentral," +
								" Convert(varchar(36), SessionTokenModuloLoja) AS SessionTokenModuloLoja" +
							" FROM t_USUARIO" +
							" WHERE" +
								" (usuario = '" + id_usuario + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Usuário inválido!";
						return null;
					}

					usuario = usuarioLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return usuario;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getCampoDataTabelaParametro ]
		public static DateTime getCampoDataTabelaParametro(String nomeParametro)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			DateTime dtHrResultado = DateTime.MinValue;
			SqlConnection cn;
			SqlCommand cmCommand;
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			#endregion

			strSql = "SELECT " +
						Global.sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador("campo_data") +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = objResultado.ToString();
				if ((strResultado != null) && (strResultado.Length > 0)) dtHrResultado = Global.converteYyyyMmDdHhMmSsParaDateTime(strResultado);
			}
			return dtHrResultado;
		}
		#endregion

		#region [ getCampoInteiroTabelaParametro ]
		public static int getCampoInteiroTabelaParametro(String nomeParametro)
		{
			return getCampoInteiroTabelaParametro(nomeParametro, 0);
		}

		public static int getCampoInteiroTabelaParametro(String nomeParametro, int valorDefault)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			int intResultado;
			SqlConnection cn;
			SqlCommand cmCommand;
			#endregion

			intResultado = valorDefault;

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			#endregion

			strSql = "SELECT " +
						"campo_inteiro" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				intResultado = BD.readToInt(objResultado);
			}
			return intResultado;
		}
		#endregion

		#region [ getCampoTextoTabelaParametro ]
		public static String getCampoTextoTabelaParametro(String nomeParametro)
		{
			return getCampoTextoTabelaParametro(nomeParametro, "");
		}

		public static String getCampoTextoTabelaParametro(String nomeParametro, String valorDefault)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			SqlConnection cn;
			SqlCommand cmCommand;
			#endregion

			strResultado = valorDefault;

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			#endregion

			strSql = "SELECT " +
						"campo_texto" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = BD.readToString(objResultado);
			}
			return strResultado;
		}
		#endregion

		#region [ setCampoDataTabelaParametro ]
		public static bool setCampoDataTabelaParametro(String nomeParametro, DateTime dtHrValorParametro)
		{
			#region [ Declarações ]
			String strSql;
			String strValorParametro;
			SqlConnection cn;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				#endregion

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Prepara o valor do parâmetro p/ o SQL ]
				if (dtHrValorParametro == DateTime.MinValue)
				{
					strValorParametro = "NULL";
				}
				else
				{
					strValorParametro = Global.sqlMontaDateTimeParaSqlDateTime(dtHrValorParametro);
				}
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_data = " + strValorParametro +
								", dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_data, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								strValorParametro + ", " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				intQtdeUpdated = cmCommand.ExecuteNonQuery();
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_data no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ setCampoInteiroTabelaParametro ]
		public static bool setCampoInteiroTabelaParametro(String nomeParametro, int valorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				#endregion

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_inteiro = " + valorParametro.ToString() +
								", dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_inteiro, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								valorParametro.ToString() + ", " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				intQtdeUpdated = cmCommand.ExecuteNonQuery();
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_inteiro no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ setCampoTextoTabelaParametro ]
		public static bool setCampoTextoTabelaParametro(String nomeParametro, String valorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				#endregion

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_texto = @campo_texto," +
								" dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_texto, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								"@campo_texto, " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				cmCommand.Parameters.Add("@campo_texto", SqlDbType.VarChar, 1024);
				cmCommand.Parameters["@campo_texto"].Value = valorParametro;
				intQtdeUpdated = cmCommand.ExecuteNonQuery();
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_texto no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaTabelaControleNsu ]
		private static bool atualizaTabelaControleNsu(ref SqlConnection cn, ref SqlTransaction trx,
													String id_nsu,
													String nsu_novo,
													String nsu_atual,
													out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.atualizaTabelaControleNsu()";
			bool blnSucesso = false;
			bool blnAbriuConexao = false;
			int intRetorno;
			string strSql;
			SqlCommand cmUpdateTabelaControleNsu;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_nsu == null)
				{
					strMsgErro = "Não foi informado o identificador do NSU!!";
					return false;
				}

				if (id_nsu.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do NSU!!";
					return false;
				}

				if (nsu_novo == null)
				{
					strMsgErro = "Não foi informado o valor do novo NSU!!";
					return false;
				}

				if (nsu_novo.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o valor do novo NSU!!";
					return false;
				}

				if (nsu_atual == null)
				{
					strMsgErro = "Não foi informado o valor do NSU atual!!";
					return false;
				}
				#endregion

				#region [ Prepara acesso ao BD ]
				if (cn == null)
				{
					cn = new SqlConnection(BD.getConnectionString());
					cn.Open();
					blnAbriuConexao = true;
				}

				// Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes através
				// do controle que impede o campo 'dt_ult_atualizacao' de receber uma data menor do que aquela que ele já possui
				strSql = "UPDATE t_CONTROLE SET " +
							"nsu = @nsu_novo, " +
							"dt_ult_atualizacao = CASE WHEN dt_ult_atualizacao > " + Global.sqlMontaGetdateSomenteData() + " THEN dt_ult_atualizacao ELSE " + Global.sqlMontaGetdateSomenteData() + " END" +
						" WHERE" +
							" (id_nsu = @id_nsu)" +
							" AND (nsu = @nsu_atual)";
				cmUpdateTabelaControleNsu = new SqlCommand();
				cmUpdateTabelaControleNsu.Connection = cn;
				if (trx != null) cmUpdateTabelaControleNsu.Transaction = trx;
				cmUpdateTabelaControleNsu.CommandText = strSql;
				cmUpdateTabelaControleNsu.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
				cmUpdateTabelaControleNsu.Parameters.Add("@nsu_novo", SqlDbType.VarChar, 12);
				cmUpdateTabelaControleNsu.Parameters.Add("@nsu_atual", SqlDbType.VarChar, 12);
				cmUpdateTabelaControleNsu.Prepare();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Preenche o valor dos parâmetros ]
					cmUpdateTabelaControleNsu.Parameters["@id_nsu"].Value = id_nsu;
					cmUpdateTabelaControleNsu.Parameters["@nsu_novo"].Value = nsu_novo;
					cmUpdateTabelaControleNsu.Parameters["@nsu_atual"].Value = nsu_atual;
					#endregion

					#region [ Tenta alterar o registro ]
					try
					{
						intRetorno = cmUpdateTabelaControleNsu.ExecuteNonQuery();
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
						Global.gravaLogAtividade(strMsgErro);
					}

					if (intRetorno == 1)
					{
						blnSucesso = true;
					}
					else
					{
						blnSucesso = false;
					}
					#endregion
				}
				finally
				{
					if (blnAbriuConexao) BD.fechaConexao(ref cn);
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaTabelaControleNsuComLetraSeq ]
		private static bool atualizaTabelaControleNsuComLetraSeq(ref SqlConnection cn, ref SqlTransaction trx,
													String id_nsu,
													String nsu_novo,
													String nsu_atual,
													String ano_letra_seq_novo,
													out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.atualizaTabelaControleNsuComLetraSeq()";
			bool blnSucesso = false;
			bool blnAbriuConexao = false;
			int intRetorno;
			string strSql;
			SqlCommand cmUpdateTabelaControleNsuComLetraSeq;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_nsu == null)
				{
					strMsgErro = "Não foi informado o identificador do NSU!!";
					return false;
				}

				if (id_nsu.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do NSU!!";
					return false;
				}

				if (nsu_novo == null)
				{
					strMsgErro = "Não foi informado o valor do novo NSU!!";
					return false;
				}

				if (nsu_novo.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o valor do novo NSU!!";
					return false;
				}

				if (nsu_atual == null)
				{
					strMsgErro = "Não foi informado o valor do NSU atual!!";
					return false;
				}
				#endregion

				#region [ Prepara acesso ao BD ]
				if (cn == null)
				{
					cn = new SqlConnection(BD.getConnectionString());
					cn.Open();
					blnAbriuConexao = true;
				}

				// Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes através
				// do controle que impede o campo 'dt_ult_atualizacao' de receber uma data menor do que aquela que ele já possui
				strSql = "UPDATE t_CONTROLE SET " +
							"nsu = @nsu_novo, " +
							"ano_letra_seq = @ano_letra_seq_novo, " +
							"dt_ult_atualizacao = CASE WHEN dt_ult_atualizacao > " + Global.sqlMontaGetdateSomenteData() + " THEN dt_ult_atualizacao ELSE " + Global.sqlMontaGetdateSomenteData() + " END" +
						" WHERE" +
							" (id_nsu = @id_nsu)" +
							" AND (nsu = @nsu_atual)";
				cmUpdateTabelaControleNsuComLetraSeq = new SqlCommand();
				cmUpdateTabelaControleNsuComLetraSeq.Connection = cn;
				if (trx != null) cmUpdateTabelaControleNsuComLetraSeq.Transaction = trx;
				cmUpdateTabelaControleNsuComLetraSeq.CommandText = strSql;
				cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
				cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@nsu_novo", SqlDbType.VarChar, 12);
				cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@nsu_atual", SqlDbType.VarChar, 12);
				cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@ano_letra_seq_novo", SqlDbType.VarChar, 1);
				cmUpdateTabelaControleNsuComLetraSeq.Prepare();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Preenche o valor dos parâmetros ]
					cmUpdateTabelaControleNsuComLetraSeq.Parameters["@id_nsu"].Value = id_nsu;
					cmUpdateTabelaControleNsuComLetraSeq.Parameters["@nsu_novo"].Value = nsu_novo;
					cmUpdateTabelaControleNsuComLetraSeq.Parameters["@nsu_atual"].Value = nsu_atual;
					cmUpdateTabelaControleNsuComLetraSeq.Parameters["@ano_letra_seq_novo"].Value = ano_letra_seq_novo;
					#endregion

					#region [ Tenta alterar o registro ]
					try
					{
						intRetorno = cmUpdateTabelaControleNsuComLetraSeq.ExecuteNonQuery();
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
						Global.gravaLogAtividade(strMsgErro);
					}

					if (intRetorno == 1)
					{
						blnSucesso = true;
					}
					else
					{
						blnSucesso = false;
					}
					#endregion
				}
				finally
				{
					if (blnAbriuConexao) BD.fechaConexao(ref cn);
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ geraNsuUsandoTabelaControle ]
		public static bool geraNsuUsandoTabelaControle(ref SqlConnection cn, ref SqlTransaction trx, String id_nsu, out String nsu_novo, out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.geraNsuUsandoTabelaControle()";
			const int MAX_TENTATIVAS = 10;
			int intQtdeTentativas = 0;
			bool blnRetorno;
			#endregion

			nsu_novo = "";
			strMsgErro = "";

			try
			{
				while (true)
				{
					intQtdeTentativas++;

					blnRetorno = executaGeraNsuUsandoTabelaControle(ref cn, ref trx, id_nsu, out nsu_novo, out strMsgErro);
					if (blnRetorno) return true;

					if (intQtdeTentativas > MAX_TENTATIVAS)
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar gerar o NSU após " + MAX_TENTATIVAS.ToString() + "!!" + strMsgErro;
						return false;
					}

					Thread.Sleep(100);
				}
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaGeraNsuUsandoTabelaControle ]
		private static bool executaGeraNsuUsandoTabelaControle(ref SqlConnection cn, ref SqlTransaction trx, String id_nsu, out String nsu_novo, out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.executaGeraNsuUsandoTabelaControle()";
			int n_nsu;
			bool blnAbriuConexao = false;
			String strNsuNovo;
			String strNsuAtual = "";
			String strLetraSeqNovo;
			String strSql;
			SqlCommand cmCommand;
			SqlCommand cmUpdateTabelaControleNsuAcquireXLock;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			nsu_novo = "";
			strMsgErro = "";
			try
			{
				#region [ Cria objetos de BD ]
				if (cn == null)
				{
					cn = new SqlConnection(BD.getConnectionString());
					cn.Open();
					blnAbriuConexao = true;
				}

				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				if (trx != null) cmCommand.Transaction = trx;
				daAdapter = new SqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Consistências ]
					if (id_nsu == null)
					{
						strMsgErro = "Não foi informado o NSU a ser gerado!!";
						return false;
					}

					if (id_nsu.ToString().Trim().Length == 0)
					{
						strMsgErro = "Não foi especificado o NSU a ser gerado!!";
						return false;
					}
					#endregion

					strMsgErro = "";
					n_nsu = -1;

					#region [ Bloqueia registro p/ evitar acesso concorrente ]
					if (Global.Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO)
					{
						cmUpdateTabelaControleNsuAcquireXLock = new SqlCommand();
						cmUpdateTabelaControleNsuAcquireXLock.Connection = cn;
						if (trx != null) cmUpdateTabelaControleNsuAcquireXLock.Transaction = trx;
						strSql = "UPDATE t_CONTROLE SET" +
									" dummy = ~dummy" +
								" WHERE" +
									" (id_nsu = @id_nsu)";
						cmUpdateTabelaControleNsuAcquireXLock.CommandText = strSql;
						cmUpdateTabelaControleNsuAcquireXLock.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
						cmUpdateTabelaControleNsuAcquireXLock.Prepare();
						cmUpdateTabelaControleNsuAcquireXLock.Parameters["@id_nsu"].Value = id_nsu;
						cmUpdateTabelaControleNsuAcquireXLock.ExecuteNonQuery();
					}
					#endregion

					strSql = "SELECT * FROM t_CONTROLE WHERE (id_nsu = '" + id_nsu + "')";

					#region [ Executa a consulta no BD ]
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					#endregion

					#region [ Ainda não existe registro de controle para gerar este NSU? ]
					if (dtbConsulta.Rows.Count == 0)
					{
						strMsgErro = "Não existe registro na tabela de controle para poder gerar este NSU!!";
						return false;
					}
					#endregion

					rowConsulta = dtbConsulta.Rows[0];
					if (!Convert.IsDBNull(rowConsulta["nsu"]))
					{
						strNsuAtual = BD.readToString(rowConsulta["nsu"]);
						if (strNsuAtual.Trim().Length > 0)
						{
							n_nsu = (int)Global.converteInteiro(strNsuAtual);
							if (BD.readToInt(rowConsulta["seq_anual"]) != 0)
							{
								// Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes
								if (DateTime.Today.Year > BD.readToDateTime(rowConsulta["dt_ult_atualizacao"]).Year)
								{
									// Se mudou o ano, reinicia a contagem do NSU
									strNsuNovo = "".PadLeft(Global.Cte.Etc.TAM_MAX_NSU, '0');
									n_nsu = 0;
									if (BD.readToString(rowConsulta["ano_letra_seq"]).Trim().Length > 0)
									{
										strLetraSeqNovo = BD.readToString(rowConsulta["ano_letra_seq"]);
										strLetraSeqNovo = Texto.chr((short)(Texto.asc(strLetraSeqNovo[0]) + BD.readToInt(rowConsulta["ano_letra_step"]))).ToString();
										if (!atualizaTabelaControleNsuComLetraSeq(ref cn, ref trx, id_nsu, strNsuNovo, strNsuAtual, strLetraSeqNovo, out strMsgErro))
										{
											if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
											strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
											return false;
										}
									}
									else
									{
										if (!atualizaTabelaControleNsu(ref cn, ref trx, id_nsu, strNsuNovo, strNsuAtual, out strMsgErro))
										{
											if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
											strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
											return false;
										}
									}
								}
							}
						}
					}

					if (n_nsu < 0)
					{
						strMsgErro = "O NSU gerado é inválido!!";
						return false;
					}

					n_nsu++;
					strNsuNovo = n_nsu.ToString().PadLeft(Global.Cte.Etc.TAM_MAX_NSU, '0');
					if (!atualizaTabelaControleNsu(ref cn, ref trx, id_nsu, strNsuNovo, strNsuAtual, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar atualizar a tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
						return false;
					}
				}
				finally
				{
					if (blnAbriuConexao) BD.fechaConexao(ref cn);
				}

				nsu_novo = strNsuNovo;
				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion
	}
}