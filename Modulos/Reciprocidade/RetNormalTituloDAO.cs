#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class RetNormalTituloDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		#endregion

		#region [ Construtor estático ]
		static RetNormalTituloDAO()
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
			strSql = "INSERT INTO t_SERASA_RETORNO_NORMAL_TITULO " +
						"(id, id_serasa_arq_retorno_normal, id_registro_dados, cnpj_cliente, tipo_dados, num_titulo, s_data_emissao, s_valor_titulo, " +
						" s_data_vencto, s_data_pagto, indicador_num_titulo_estendido, num_titulo_estendido, codigos_erro) " +
					 "VALUES " +
						"(@id, @id_serasa_arq_retorno_normal, @id_registro_dados, @cnpj_cliente, @tipo_dados, @num_titulo, @s_data_emissao, @s_valor_titulo, " +
						"@s_data_vencto, @s_data_pagto, @indicador_num_titulo_estendido, @num_titulo_estendido, @codigos_erro) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_arq_retorno_normal", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_registro_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@cnpj_cliente", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@tipo_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_titulo", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@s_data_emissao", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@s_valor_titulo", SqlDbType.VarChar, 13);
			cmInsere.Parameters.Add("@s_data_vencto", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@s_data_pagto", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@indicador_num_titulo_estendido", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_titulo_estendido", SqlDbType.VarChar, 32);
			cmInsere.Parameters.Add("@codigos_erro", SqlDbType.VarChar, 90);
			cmInsere.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   int id_serasa_arq_retorno_normal,
								   String id_registro_dados,
								   String cnpj_cliente,
								   String tipo_dados,
								   String num_titulo,
								   String s_data_emissao,
								   String s_valor_titulo,
								   String s_data_vencto,
								   String s_data_pagto,
								   String indicador_num_titulo_estendido,
								   String num_titulo_estendido,
								   String codigos_erro)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_RETORNO_NORMAL_TITULO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@id_serasa_arq_retorno_normal"].Value = id_serasa_arq_retorno_normal;
			cmInsere.Parameters["@id_registro_dados"].Value = id_registro_dados;
			cmInsere.Parameters["@cnpj_cliente"].Value = cnpj_cliente;
			cmInsere.Parameters["@tipo_dados"].Value = tipo_dados;
			cmInsere.Parameters["@num_titulo"].Value = num_titulo;
			cmInsere.Parameters["@s_data_emissao"].Value = s_data_emissao;
			cmInsere.Parameters["@s_valor_titulo"].Value = s_valor_titulo;
			cmInsere.Parameters["@s_data_vencto"].Value = s_data_vencto;
			cmInsere.Parameters["@s_data_pagto"].Value = s_data_pagto;
			cmInsere.Parameters["@indicador_num_titulo_estendido"].Value = indicador_num_titulo_estendido;
			cmInsere.Parameters["@num_titulo_estendido"].Value = num_titulo_estendido;
			cmInsere.Parameters["@codigos_erro"].Value = codigos_erro;

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
