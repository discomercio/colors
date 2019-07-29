using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Domains;

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
	}
}