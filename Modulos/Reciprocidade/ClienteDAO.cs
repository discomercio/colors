#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class ClienteDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmAtualizaInfoEnvioCNPJ;
		#endregion

		#region [ Construtor estático ]
		static ClienteDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmAtualizaInfoEnvioCNPJ ]
			strSql = "UPDATE t_SERASA_CLIENTE " +
					 "SET st_enviado_serasa = @st_enviado_serasa, " +
					 "dt_enviado_serasa = @dt_enviado_serasa, " +
					 "id_serasa_arq_remessa_normal = @id_serasa_arq_remessa_normal " +
					 "WHERE cnpj = @cnpj ";

			cmAtualizaInfoEnvioCNPJ = BD.criaSqlCommand();
			cmAtualizaInfoEnvioCNPJ.CommandText = strSql;
			cmAtualizaInfoEnvioCNPJ.Parameters.Add("@st_enviado_serasa", SqlDbType.TinyInt);
			cmAtualizaInfoEnvioCNPJ.Parameters.Add("@dt_enviado_serasa", SqlDbType.DateTime);
			cmAtualizaInfoEnvioCNPJ.Parameters.Add("@id_serasa_arq_remessa_normal", SqlDbType.Int);
			cmAtualizaInfoEnvioCNPJ.Parameters.Add("@cnpj", SqlDbType.VarChar, 14);
			cmAtualizaInfoEnvioCNPJ.Prepare();
			#endregion
		}
		#endregion

		#region [ Atualiza Informações de Envio do CNPJ a Serasa ]
		public static bool atualizaInfoEnvioCNPJ(int st_enviado_serasa,
												DateTime dt_enviado_serasa,
												int id_serasa_arq_remessa_normal,
												String cnpj)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_CLIENTE";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaInfoEnvioCNPJ.Parameters["@st_enviado_serasa"].Value = st_enviado_serasa;
			cmAtualizaInfoEnvioCNPJ.Parameters["@dt_enviado_serasa"].Value = dt_enviado_serasa;
			cmAtualizaInfoEnvioCNPJ.Parameters["@id_serasa_arq_remessa_normal"].Value = id_serasa_arq_remessa_normal;
			cmAtualizaInfoEnvioCNPJ.Parameters["@cnpj"].Value = cnpj;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaInfoEnvioCNPJ);
			}
			catch (Exception ex)
			{
				intRetorno = 0;
				Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
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

			return blnSucesso;
		}
		#endregion

		#region [ radicalCNPJSacadoJaEnviado ]
		public static bool radicalCNPJSacadoJaEnviado(string radicalCNPJSacado)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			object ret = null;
			int status;
			const int CNPJ_NAO_ENVIADO = 0;
			#endregion

			cmCommand = BD.criaSqlCommand();

			strSql = "SELECT COUNT(*) " +
					 "FROM t_SERASA_CLIENTE " +
					 "WHERE raiz_cnpj = @raiz_cnpj " +
						"AND st_enviado_serasa = 1 ";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.Add("@raiz_cnpj", SqlDbType.VarChar, 8);
			cmCommand.Parameters["@raiz_cnpj"].Value = radicalCNPJSacado.Trim();

			ret = cmCommand.ExecuteScalar();
			status = BD.readToInt(ret);

			if (status == CNPJ_NAO_ENVIADO)
			{
				return false;
			}

			return true;
		}
		#endregion
	}
}
