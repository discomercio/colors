#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class ArqRemessaDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		private static SqlCommand cmAtualizaStatusGeracao;
		#endregion

		#region [ Construtor estático ]
		static ArqRemessaDAO()
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
			strSql = "INSERT INTO t_SERASA_ARQ_REMESSA_NORMAL " +
						"(id, dt_geracao, dt_hr_geracao, usuario_geracao, cnpj_empresa_conveniada, s_periodo_inicio, dt_periodo_inicio, s_periodo_termino, " +
						" dt_periodo_termino, periodicidade_remessa, reservado_serasa_1, num_id_grupo_relato_segmento, id_versao_layout, num_versao_layout, " +
						" qtde_reg_tempo_relac, qtde_reg_titulos, duracao_proc_em_seg, nome_arq_remessa, caminho_arq_remessa, st_geracao, msg_erro_geracao) " +
					 "VALUES " +
						"(@id, @dt_geracao, @dt_hr_geracao, @usuario_geracao, @cnpj_empresa_conveniada, @s_periodo_inicio, @dt_periodo_inicio, @s_periodo_termino, " +
						"@dt_periodo_termino, @periodicidade_remessa, @reservado_serasa_1, @num_id_grupo_relato_segmento, @id_versao_layout, @num_versao_layout, " +
						"@qtde_reg_tempo_relac, @qtde_reg_titulos, @duracao_proc_em_seg, @nome_arq_remessa, @caminho_arq_remessa, @st_geracao, @msg_erro_geracao) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@dt_geracao", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@dt_hr_geracao", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@usuario_geracao", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@cnpj_empresa_conveniada", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@s_periodo_inicio", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_periodo_inicio", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@s_periodo_termino", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_periodo_termino", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@periodicidade_remessa", SqlDbType.VarChar, 1);
			cmInsere.Parameters.Add("@reservado_serasa_1", SqlDbType.VarChar, 15);
			cmInsere.Parameters.Add("@num_id_grupo_relato_segmento", SqlDbType.VarChar, 3);
			cmInsere.Parameters.Add("@id_versao_layout", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_versao_layout", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@qtde_reg_tempo_relac", SqlDbType.Int);
			cmInsere.Parameters.Add("@qtde_reg_titulos", SqlDbType.Int);
			cmInsere.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmInsere.Parameters.Add("@nome_arq_remessa", SqlDbType.VarChar, 40);
			cmInsere.Parameters.Add("@caminho_arq_remessa", SqlDbType.VarChar, 1024);
			cmInsere.Parameters.Add("@st_geracao", SqlDbType.SmallInt);
			cmInsere.Parameters.Add("@msg_erro_geracao", SqlDbType.VarChar, 1024);
			cmInsere.Prepare();
			#endregion

			#region [ cmAtualizaStatusGeracao ]
			strSql = "UPDATE t_SERASA_ARQ_REMESSA_NORMAL " +
					 "SET st_geracao = @st_geracao, " +
						 "msg_erro_geracao = @msg_erro_geracao " +
						 "WHERE id = @id_serasa_arq_remessa_normal ";

			cmAtualizaStatusGeracao = BD.criaSqlCommand();
			cmAtualizaStatusGeracao.CommandText = strSql;
			cmAtualizaStatusGeracao.Parameters.Add("@st_geracao", SqlDbType.SmallInt);
			cmAtualizaStatusGeracao.Parameters.Add("@msg_erro_geracao", SqlDbType.VarChar, 1024);
			cmAtualizaStatusGeracao.Parameters.Add("@id_serasa_arq_remessa_normal", SqlDbType.Int);
			cmAtualizaStatusGeracao.Prepare();
			#endregion
		}
		#endregion

		#region [ obtemPeriodoUltRemessa ]
		public static bool obtemPeriodoUltRemessa(out int id_arq_remessa_normal, out DateTime dt_periodo_inicio, out DateTime dt_periodo_termino, out DateTime dt_geracao, out String strMsgErro)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Inicialização ]
			id_arq_remessa_normal = 0;
			dt_periodo_inicio = DateTime.MinValue;
			dt_periodo_termino = DateTime.MinValue;
			dt_geracao = DateTime.MinValue;
			strMsgErro = "";
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();
				daDataAdapter = BD.criaSqlDataAdapter();
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

				strSql = "SELECT TOP 1" +
							" id," +
							" dt_periodo_inicio," +
							" dt_periodo_termino," +
							" dt_geracao" +
						" FROM t_SERASA_ARQ_REMESSA_NORMAL" +
						" WHERE" +
							" (st_geracao = 1)" +
						" ORDER BY" +
							" dt_periodo_termino DESC," +
							" id DESC";
				cmCommand.CommandText = strSql;
				daDataAdapter.Fill(dtbResultado);

				if (dtbResultado.Rows.Count == 0)
				{
					// Ao gerar a 1ª remessa, não haverá nenhuma remessa anterior, portanto, não ter encontrado a remessa anterior é uma situação válida
					return true;
				}

				rowResultado = dtbResultado.Rows[0];

				id_arq_remessa_normal = BD.readToInt(rowResultado["id"]);
				dt_periodo_inicio = BD.readToDateTime(rowResultado["dt_periodo_inicio"]);
				dt_periodo_termino = BD.readToDateTime(rowResultado["dt_periodo_termino"]);
				dt_geracao = BD.readToDateTime(rowResultado["dt_geracao"]);

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   DateTime dt_geracao,
								   DateTime dt_hr_geracao,
								   String usuario_geracao,
								   String cnpj_empresa_conveniada,
								   String s_periodo_inicio,
								   DateTime dt_periodo_inicio,
								   String s_periodo_termino,
								   DateTime dt_periodo_termino,
								   String periodicidade_remessa,
								   String reservado_serasa_1,
								   String num_id_grupo_relato_segmento,
								   String id_versao_layout,
								   String num_versao_layout,
								   int qtde_reg_tempo_relac,
								   int qtde_reg_titulos,
								   int duracao_proc_em_seg,
								   String nome_arq_remessa,
								   String caminho_arq_remessa,
								   int st_geracao,
								   String msg_erro_geracao)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_ARQ_REMESSA_NORMAL";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@dt_geracao"].Value = dt_geracao.Date;
			cmInsere.Parameters["@dt_hr_geracao"].Value = dt_hr_geracao;
			cmInsere.Parameters["@usuario_geracao"].Value = usuario_geracao;
			cmInsere.Parameters["@cnpj_empresa_conveniada"].Value = cnpj_empresa_conveniada;
			cmInsere.Parameters["@s_periodo_inicio"].Value = s_periodo_inicio;
			cmInsere.Parameters["@dt_periodo_inicio"].Value = dt_periodo_inicio;
			cmInsere.Parameters["@s_periodo_termino"].Value = s_periodo_termino;
			cmInsere.Parameters["@dt_periodo_termino"].Value = dt_periodo_termino;
			cmInsere.Parameters["@periodicidade_remessa"].Value = periodicidade_remessa;

			if (reservado_serasa_1 == null)
			{
				cmInsere.Parameters["@reservado_serasa_1"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@reservado_serasa_1"].Value = reservado_serasa_1;
			}

			if (num_id_grupo_relato_segmento == null)
			{
				cmInsere.Parameters["@num_id_grupo_relato_segmento"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@num_id_grupo_relato_segmento"].Value = num_id_grupo_relato_segmento;
			}

			cmInsere.Parameters["@id_versao_layout"].Value = id_versao_layout;
			cmInsere.Parameters["@num_versao_layout"].Value = num_versao_layout;
			cmInsere.Parameters["@qtde_reg_tempo_relac"].Value = qtde_reg_tempo_relac;
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
												 int id_serasa_arq_remessa_normal)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_ARQ_REMESSA_NORMAL";
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

			cmAtualizaStatusGeracao.Parameters["@id_serasa_arq_remessa_normal"].Value = id_serasa_arq_remessa_normal;

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
