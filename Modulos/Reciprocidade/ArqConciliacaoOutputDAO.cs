#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class ArqConciliacaoOutputDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		private static SqlCommand cmAtualizaStatusGeracao;
		#endregion

		#region [ Construtor estático ]
		static ArqConciliacaoOutputDAO()
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
			strSql = "INSERT INTO t_SERASA_ARQ_CONCILIACAO_OUTPUT " +
						"(id, dt_geracao, dt_hr_geracao, usuario_geracao, cnpj_empresa_conveniada, s_periodo_termino, dt_periodo_termino, " +
						" periodicidade_remessa, reservado_serasa_1, num_id_grupo_relato_segmento, id_versao_layout, num_versao_layout, qtde_reg_titulos, duracao_proc_em_seg, " +
						" nome_arq_remessa, caminho_arq_remessa, st_geracao, msg_erro_geracao) " +
					 "VALUES " +
						"(@id, @dt_geracao, @dt_hr_geracao, @usuario_geracao, @cnpj_empresa_conveniada, @s_periodo_termino, @dt_periodo_termino, " +
						" @periodicidade_remessa, @reservado_serasa_1, @num_id_grupo_relato_segmento, @id_versao_layout, @num_versao_layout, @qtde_reg_titulos, @duracao_proc_em_seg, " +
						" @nome_arq_remessa, @caminho_arq_remessa, @st_geracao, @msg_erro_geracao) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@dt_geracao", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@dt_hr_geracao", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@usuario_geracao", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@cnpj_empresa_conveniada", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@s_periodo_termino", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_periodo_termino", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@periodicidade_remessa", SqlDbType.VarChar, 1);
			cmInsere.Parameters.Add("@reservado_serasa_1", SqlDbType.VarChar, 15);
			cmInsere.Parameters.Add("@num_id_grupo_relato_segmento", SqlDbType.VarChar, 3);
			cmInsere.Parameters.Add("@id_versao_layout", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_versao_layout", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@qtde_reg_titulos", SqlDbType.Int);
			cmInsere.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmInsere.Parameters.Add("@nome_arq_remessa", SqlDbType.VarChar, 40);
			cmInsere.Parameters.Add("@caminho_arq_remessa", SqlDbType.VarChar, 1024);
			cmInsere.Parameters.Add("@st_geracao", SqlDbType.SmallInt);
			cmInsere.Parameters.Add("@msg_erro_geracao", SqlDbType.VarChar, 1024);
			cmInsere.Prepare();
			#endregion

			#region [ cmAtualizaStatusGeracao ]
			strSql = "UPDATE t_SERASA_ARQ_CONCILIACAO_OUTPUT " +
					 "SET st_geracao = @st_geracao, " +
						 "msg_erro_geracao = @msg_erro_geracao " +
						 "WHERE id = @id_serasa_arq_conciliacao_output ";

			cmAtualizaStatusGeracao = BD.criaSqlCommand();
			cmAtualizaStatusGeracao.CommandText = strSql;
			cmAtualizaStatusGeracao.Parameters.Add("@st_geracao", SqlDbType.SmallInt);
			cmAtualizaStatusGeracao.Parameters.Add("@msg_erro_geracao", SqlDbType.VarChar, 1024);
			cmAtualizaStatusGeracao.Parameters.Add("@id_serasa_arq_conciliacao_output", SqlDbType.Int);
			cmAtualizaStatusGeracao.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   DateTime dt_geracao,
								   DateTime dt_hr_geracao,
								   String usuario_geracao,
								   String cnpj_empresa_conveniada,
								   String s_periodo_termino,
								   DateTime dt_periodo_termino,
								   String periodicidade_remessa,
								   String reservado_serasa_1,
								   String num_id_grupo_relato_segmento,
								   String id_versao_layout,
								   String num_versao_layout,
								   int qtde_reg_titulos,
								   int duracao_proc_em_seg,
								   String nome_arq_remessa,
								   String caminho_arq_remessa,
								   int st_geracao,
								   String msg_erro_geracao)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_ARQ_CONCILIACAO_OUTPUT";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@dt_geracao"].Value = dt_geracao.Date;
			cmInsere.Parameters["@dt_hr_geracao"].Value = dt_hr_geracao;
			cmInsere.Parameters["@usuario_geracao"].Value = usuario_geracao;
			cmInsere.Parameters["@cnpj_empresa_conveniada"].Value = cnpj_empresa_conveniada;
			cmInsere.Parameters["@s_periodo_termino"].Value = s_periodo_termino;
			cmInsere.Parameters["@dt_periodo_termino"].Value = dt_periodo_termino;
			cmInsere.Parameters["@periodicidade_remessa"].Value = periodicidade_remessa;
			cmInsere.Parameters["@reservado_serasa_1"].Value = reservado_serasa_1;
			cmInsere.Parameters["@num_id_grupo_relato_segmento"].Value = num_id_grupo_relato_segmento;
			cmInsere.Parameters["@id_versao_layout"].Value = id_versao_layout;
			cmInsere.Parameters["@num_versao_layout"].Value = num_versao_layout;
			cmInsere.Parameters["@qtde_reg_titulos"].Value = qtde_reg_titulos;
			cmInsere.Parameters["@duracao_proc_em_seg"].Value = duracao_proc_em_seg;
			cmInsere.Parameters["@nome_arq_remessa"].Value = nome_arq_remessa;
			cmInsere.Parameters["@caminho_arq_remessa"].Value = caminho_arq_remessa;
			cmInsere.Parameters["@st_geracao"].Value = st_geracao;

			if (msg_erro_geracao == null)
			{
				cmInsere.Parameters["@msg_erro_geracao"].Value = DBNull.Value;
			}
			else if (msg_erro_geracao.Length > 1024)
			{
				cmInsere.Parameters["@msg_erro_geracao"].Value = msg_erro_geracao.Substring(0, 1024);
			}
			else
			{
				cmInsere.Parameters["@msg_erro_geracao"].Value = msg_erro_geracao;
			}

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

		#region [ atualizaStatusGeracao ]
		public static bool atualizaStatusGeracao(int st_geracao,
												 String msg_erro_geracao,
												 int id_serasa_arq_conciliacao_output)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_ARQ_CONCILIACAO_OUTPUT";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaStatusGeracao.Parameters["@st_geracao"].Value = st_geracao;

			if (msg_erro_geracao == null)
			{
				cmAtualizaStatusGeracao.Parameters["@msg_erro_geracao"].Value = DBNull.Value;
			}
			else if (msg_erro_geracao.Length > 1024)
			{
				cmAtualizaStatusGeracao.Parameters["@msg_erro_geracao"].Value = msg_erro_geracao.Substring(0, 1024);
			}
			else
			{
				cmAtualizaStatusGeracao.Parameters["@msg_erro_geracao"].Value = msg_erro_geracao;
			}

			cmAtualizaStatusGeracao.Parameters["@id_serasa_arq_conciliacao_output"].Value = id_serasa_arq_conciliacao_output;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaStatusGeracao);
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
