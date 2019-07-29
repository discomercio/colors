#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class DetTituloDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		#endregion

		#region [ Construtor estático ]
		static DetTituloDAO()
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
			strSql = "INSERT INTO t_SERASA_REMESSA_DET_TITULO " +
						"(id, id_serasa_arq_remessa_normal, id_serasa_titulo_movimento, id_serasa_cliente, id_registro_dados, " +
						"cnpj_cliente, tipo_dados, num_titulo, dt_emissao, vl_titulo, dt_vencto, dt_pagto, indicador_num_titulo_estendido, num_titulo_estendido) " +
					 "VALUES " +
						"(@id, @id_serasa_arq_remessa_normal, @id_serasa_titulo_movimento, @id_serasa_cliente, @id_registro_dados, " +
						" @cnpj_cliente, @tipo_dados, @num_titulo, @dt_emissao, @vl_titulo, @dt_vencto, @dt_pagto, @indicador_num_titulo_estendido, @num_titulo_estendido) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_arq_remessa_normal", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_titulo_movimento", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_cliente", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_registro_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@cnpj_cliente", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@tipo_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_titulo", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@dt_emissao", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@vl_titulo", SqlDbType.Money);
			cmInsere.Parameters.Add("@dt_vencto", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@dt_pagto", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@indicador_num_titulo_estendido", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_titulo_estendido", SqlDbType.VarChar, 32);
			cmInsere.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   int id_serasa_arq_remessa_normal,
								   int id_serasa_titulo_movimento,
								   int id_serasa_cliente,
								   String id_registro_dados,
								   String cnpj_cliente,
								   String tipo_dados,
								   String num_titulo,
								   DateTime dt_emissao,
								   Decimal vl_titulo,
								   DateTime dt_vencto,
								   DateTime dt_pagto,
								   String indicador_num_titulo_estendido,
								   String num_titulo_estendido)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_REMESSA_DET_TITULO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@id_serasa_arq_remessa_normal"].Value = id_serasa_arq_remessa_normal;
			cmInsere.Parameters["@id_serasa_titulo_movimento"].Value = id_serasa_titulo_movimento;
			cmInsere.Parameters["@id_serasa_cliente"].Value = id_serasa_cliente;
			cmInsere.Parameters["@id_registro_dados"].Value = id_registro_dados;
			cmInsere.Parameters["@cnpj_cliente"].Value = cnpj_cliente;
			cmInsere.Parameters["@tipo_dados"].Value = tipo_dados;
			cmInsere.Parameters["@num_titulo"].Value = num_titulo;
			cmInsere.Parameters["@dt_emissao"].Value = dt_emissao;
			cmInsere.Parameters["@vl_titulo"].Value = vl_titulo;
			cmInsere.Parameters["@dt_vencto"].Value = dt_vencto;

			if (dt_pagto == DateTime.MinValue)
			{
				cmInsere.Parameters["@dt_pagto"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@dt_pagto"].Value = dt_pagto;
			}

			cmInsere.Parameters["@indicador_num_titulo_estendido"].Value = indicador_num_titulo_estendido;
			cmInsere.Parameters["@num_titulo_estendido"].Value = num_titulo_estendido;

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
