#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class ComumDAO
	{
		#region [ getContaCorrenteNumeroConta ]
		public static String getContaCorrenteNumeroConta(byte id_conta_corrente)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado = "";
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT" +
						" conta" +
					 " FROM t_FIN_CONTA_CORRENTE" +
					 " WHERE" +
						" (id = " + id_conta_corrente.ToString() + ")";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null) strResultado = objResultado.ToString();
			return strResultado;
		}
		#endregion

		#region [ getContaCorrenteDescricao ]
		public static String getContaCorrenteDescricao(byte id_conta_corrente)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado = "";
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT" +
						" descricao" +
					 " FROM t_FIN_CONTA_CORRENTE" +
					 " WHERE" +
						" (id = " + id_conta_corrente.ToString() + ")";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null) strResultado = objResultado.ToString();
			return strResultado;
		}
		#endregion

		#region [ getPlanoContasEmpresaDescricao ]
		public static String getPlanoContasEmpresaDescricao(byte id_plano_contas_empresa)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado = "";
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT" +
						" descricao" +
					 " FROM t_FIN_PLANO_CONTAS_EMPRESA" +
					 " WHERE" +
						" (id = " + id_plano_contas_empresa.ToString() + ")";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null) strResultado = objResultado.ToString();
			return strResultado;
		}
		#endregion

		#region [ getPlanoContasGrupoDescricao ]
		public static String getPlanoContasGrupoDescricao(int id_plano_contas_grupo)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado = "";
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT" +
						" descricao" +
					 " FROM t_FIN_PLANO_CONTAS_GRUPO" +
					 " WHERE" +
						" (id = " + id_plano_contas_grupo.ToString() + ")";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null) strResultado = objResultado.ToString();
			return strResultado;
		}
		#endregion

		#region [ getPlanoContasContaDescricao ]
		public static String getPlanoContasContaDescricao(int id_plano_contas_conta)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado = "";
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT" +
						" descricao" +
					 " FROM t_FIN_PLANO_CONTAS_CONTA" +
					 " WHERE" +
						" (id = " + id_plano_contas_conta.ToString() + ")";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null) strResultado = objResultado.ToString();
			return strResultado;
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
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT " +
						Global.sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador("campo_data") +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = BD.criaSqlCommand();
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
			SqlCommand cmCommand;
			#endregion

			intResultado = valorDefault;

			strSql = "SELECT " +
						"campo_inteiro" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				intResultado = BD.readToInt(objResultado);
			}
			return intResultado;
		}
		#endregion

		#region [ getCampoStringTabelaParametro ]
		public static String getCampoStringTabelaParametro(String nomeParametro)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado = "";
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT" +
						" Coalesce(campo_texto, '') AS campo_texto" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = objResultado.ToString();
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
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

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
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
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

		#region [ setCampoStringTabelaParametro ]
		public static bool setCampoStringTabelaParametro(String nomeParametro, String strValorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlCommand cmCommandGravacao;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

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
								" (id = @id)";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_texto, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"@id, " +
								"@campo_texto, " +
								"getdate()" +
							")";
				}
				cmCommandGravacao = BD.criaSqlCommand();
				cmCommandGravacao.CommandText = strSql;
				cmCommandGravacao.Parameters.Add("@id", SqlDbType.VarChar, 80);
				cmCommandGravacao.Parameters.Add("@campo_texto", SqlDbType.VarChar, 1024);
				cmCommandGravacao.Prepare();
				cmCommandGravacao.Parameters["@id"].Value = nomeParametro;
				cmCommandGravacao.Parameters["@campo_texto"].Value = strValorParametro;
				intQtdeUpdated = BD.executaNonQuery(ref cmCommandGravacao);
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

		#region [ setCampoInteiroTabelaParametro ]
		public static bool setCampoInteiroTabelaParametro(String nomeParametro, int valorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

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
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
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
	}
}
