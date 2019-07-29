#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class DetTempoRelacDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		#endregion

		#region [ Construtor estático ]
		static DetTempoRelacDAO()
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
			strSql = "INSERT INTO t_SERASA_REMESSA_DET_TEMPO_RELAC " +
						"(id, id_serasa_arq_remessa_normal, id_serasa_cliente, id_registro_dados, cnpj_cliente, tipo_dados, dt_cliente_desde, tipo_cliente) " +
					 "VALUES " +
						"(@id, @id_serasa_arq_remessa_normal, @id_serasa_cliente, @id_registro_dados, @cnpj_cliente, @tipo_dados, @dt_cliente_desde, @tipo_cliente) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_arq_remessa_normal", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_cliente", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_registro_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@cnpj_cliente", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@tipo_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@dt_cliente_desde", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@tipo_cliente", SqlDbType.VarChar, 1);
			cmInsere.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   int id_serasa_arq_remessa_normal,
								   int id_serasa_cliente,
								   String id_registro_dados,
								   String cnpj_cliente,
								   String tipo_dados,
								   DateTime dt_cliente_desde,
								   String tipo_cliente)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_REMESSA_DET_TEMPO_RELAC";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@id_serasa_arq_remessa_normal"].Value = id_serasa_arq_remessa_normal;
			cmInsere.Parameters["@id_serasa_cliente"].Value = id_serasa_cliente;
			cmInsere.Parameters["@id_registro_dados"].Value = id_registro_dados;
			cmInsere.Parameters["@cnpj_cliente"].Value = cnpj_cliente;
			cmInsere.Parameters["@tipo_dados"].Value = tipo_dados;
			cmInsere.Parameters["@dt_cliente_desde"].Value = dt_cliente_desde;
			cmInsere.Parameters["@tipo_cliente"].Value = tipo_cliente;

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
	}
}
