using System.Configuration;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models;
using ART3WebAPI.Models.Domains;
using System;

namespace ART3WebAPI.Models.Repository
{
	static class BD
	{
		private static string servidorBanco = ConfigurationManager.ConnectionStrings["ServidorBanco"].ConnectionString;
		private static string nomeBanco = ConfigurationManager.ConnectionStrings["NomeBanco"].ConnectionString;
		private static string loginBanco = ConfigurationManager.ConnectionStrings["LoginBanco"].ConnectionString;
		private static string senhaBanco = Domains.Criptografia.Descriptografa(ConfigurationManager.ConnectionStrings["SenhaBanco"].ConnectionString);

		public const char CARACTER_CURINGA_TODOS = '%';

		#region [ getConnectionString ]
		public static string getConnectionString()
		{
			string connectionString = string.Format("{0};{1};{2};Password={3}", servidorBanco, nomeBanco, loginBanco, senhaBanco);

			return connectionString;
		}
		#endregion

		#region [ fechaConexao ]
		public static void fechaConexao(ref SqlConnection cn)
		{
			try
			{
				if (cn != null)
				{
					if (cn.State != ConnectionState.Closed) cn.Close();
				}
			}
			catch (Exception)
			{
				// NOP
			}
		}
		#endregion

		#region [ gera_uid ]
		public static string gera_uid()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "gera_uid()";
			string strUID = "";
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			try
			{

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // Finally: cn.Close()
				{
					cmCommand = new SqlCommand();
					cmCommand.Connection = cn;
					daDataAdapter = new SqlDataAdapter();
					#endregion

					strSql = "SELECT Convert(varchar(36), NEWID()) AS uid";
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					if (dtbResultado.Rows.Count > 0)
					{
						strUID = BD.readToString(dtbResultado.Rows[0]["uid"]);
					}

					return strUID;
				}
				finally
				{
					cn.Close();
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + ex.Message);
				return "";
			}
		}
		#endregion

		#region [ obtem_descricao_tabela_t_codigo_descricao ]
		public static string obtem_descricao_tabela_t_codigo_descricao(string grupo, string codigo)
		{
			SqlConnection cn = new SqlConnection(BD.getConnectionString());
			string sql, _novo;
			sql = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='" + grupo + "') AND (codigo='" + codigo + "')";
			_novo = "";
			cn.Open();
			try // Finally: cn.Close()
			{
				SqlCommand cmd = new SqlCommand();
				cmd.Connection = cn;
				cmd.CommandText = sql.ToString();
				IDataReader reader = cmd.ExecuteReader();

				try
				{
					int idxDescricao = reader.GetOrdinal("descricao");

					while (reader.Read())
					{
						_novo = reader.IsDBNull(idxDescricao) ? "" : reader.GetString(idxDescricao);
					}
				}
				finally
				{
					reader.Close();
				}
			}
			finally
			{
				cn.Close();
			}

			return _novo;
		}
		#endregion

		#region [ readToString ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo texto
		/// </param>
		/// <returns>
		/// Retorna o texto armazenado no campo. Caso o conteúdo seja DBNull, retorna uma String vazia.
		/// </returns>
		public static String readToString(object campo)
		{
			return !Convert.IsDBNull(campo) ? campo.ToString() : "";
		}
		#endregion

		#region [ readToDateTime ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo data
		/// </param>
		/// <returns>
		/// Retorna a data armazenada no campo. Caso o conteúdo seja DBNull, retorna DateTime.MinValue
		/// </returns>
		public static DateTime readToDateTime(object campo)
		{
			return !Convert.IsDBNull(campo) ? (DateTime)campo : DateTime.MinValue;
		}
		#endregion

		#region [ readToSingle ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo real
		/// </param>
		/// <returns>
		/// Retorna o número real armazenado no campo
		/// </returns>
		public static Single readToSingle(object campo)
		{
			return (Single)campo;
		}
		#endregion

		#region [ readToByte ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo byte
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static byte readToByte(object campo)
		{
			return (byte)campo;
		}
		#endregion

		#region [ readToShort ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo short
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static short readToShort(object campo)
		{
			return (short)campo;
		}
		#endregion

		#region [ readToInt ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo int
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static int readToInt(object campo)
		{
			if (campo.GetType().Name.Equals("Int16"))
			{
				return (int)(Int16)campo;
			}
			else
			{
				return (int)campo;
			}
		}
		#endregion

		#region [ readToInt16 ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo System.Int16
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static System.Int16 readToInt16(object campo)
		{
			return (System.Int16)campo;
		}
		#endregion

		#region [ readToInt64 ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo Int64
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static Int64 readToInt64(object campo)
		{
			if (campo.GetType().Name.Equals("Int16"))
			{
				return (Int64)(Int16)campo;
			}
			else
			{
				return (Int64)campo;
			}
		}
		#endregion

		#region [ readToChar ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo char
		/// </param>
		/// <returns>
		/// Retorna o caracter armazenado no campo. Caso o conteúdo seja DBNull, retorna um caracter nulo.
		/// </returns>
		public static char readToChar(object campo)
		{
			String s;
			char c = '\0';

			if (!Convert.IsDBNull(campo))
			{
				s = campo.ToString();
				if (s.Length > 0) c = s[0];
			}

			return c;
		}
		#endregion

		#region [ readToDecimal ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo decimal
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static decimal readToDecimal(object campo)
		{
			return (decimal)campo;
		}
		#endregion

		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo"></param>
		/// Retorna o conteúdo do campo como um array de bytes. Caso o conteúdo seja DBNull, retorna null.
		/// <returns></returns>
		public static byte[]readToVarBinary(object campo)
		{
			return !Convert.IsDBNull(campo) ? ((byte[])campo) : null;
		}
	}
}