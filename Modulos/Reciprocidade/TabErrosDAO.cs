#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class TabErrosDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		#endregion

		#region [ Construtor estático ]
		static TabErrosDAO()
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

			#region [ cmInsere ]
			strSql = "INSERT INTO t_SERASA_RETORNO_NORMAL_TAB_ERROS " +
						"(id, id_serasa_arq_retorno_normal, numero_mensagem, descricao_msg_erro) " +
					 "VALUES " +
						"(@id, @id_serasa_arq_retorno_normal, @numero_mensagem, @descricao_msg_erro) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_arq_retorno_normal", SqlDbType.Int);
			cmInsere.Parameters.Add("@numero_mensagem", SqlDbType.VarChar, 3);
			cmInsere.Parameters.Add("@descricao_msg_erro", SqlDbType.VarChar, 70);
			cmInsere.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   int id_serasa_arq_retorno_normal,
								   String numero_mensagem,
								   String descricao_msg_erro)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_RETORNO_NORMAL_TAB_ERROS";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@id_serasa_arq_retorno_normal"].Value = id_serasa_arq_retorno_normal;
			cmInsere.Parameters["@numero_mensagem"].Value = numero_mensagem;
			cmInsere.Parameters["@descricao_msg_erro"].Value = descricao_msg_erro;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmInsere);
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

		#region [selecionaTabErrosPorArqRetorno]
		public static DataTable selecionaTabErrosPorArqRetorno(int id_serasa_arq_retorno_normal)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbErros = new DataTable();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"* " +
					 "FROM t_SERASA_RETORNO_NORMAL_TAB_ERROS " +
					 "WHERE id_serasa_arq_retorno_normal = @id_serasa_arq_retorno_normal ";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.AddWithValue("@id_serasa_arq_retorno_normal", id_serasa_arq_retorno_normal);
			daDataAdapter.Fill(dtbErros);

			return dtbErros;
		}
		#endregion
	}
}
