#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class ArqRetornoDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		private static SqlCommand cmAtualizaDuracaoProcessamento;
		private static SqlCommand cmAtualizaStatusProcessamento;
		#endregion

		#region [ Construtor estático ]
		static ArqRetornoDAO()
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
			strSql = "INSERT INTO t_SERASA_ARQ_RETORNO_NORMAL " +
						"(id, dt_processamento, dt_hr_processamento, usuario_processamento, " +
						" cnpj_empresa_conveniada, s_periodo_inicio, dt_periodo_inicio, " +
						" s_periodo_termino, dt_periodo_termino, periodicidade_remessa, " +
						" reservado_serasa_1, num_id_grupo_relato_segmento, id_versao_layout, num_versao_layout, " +
						" qtde_total_registros, qtde_registros_ok, qtde_registros_erro, duracao_proc_em_seg, " +
						" nome_arq_retorno, caminho_arq_retorno, st_processamento, msg_erro_processamento) " +
					 "VALUES " +
						"(@id, @dt_processamento, @dt_hr_processamento, @usuario_processamento, " +
						" @cnpj_empresa_conveniada, @s_periodo_inicio, @dt_periodo_inicio, " +
						" @s_periodo_termino, @dt_periodo_termino, @periodicidade_remessa, " +
						" @reservado_serasa_1, @num_id_grupo_relato_segmento, @id_versao_layout, @num_versao_layout, " +
						" @qtde_total_registros, @qtde_registros_ok, @qtde_registros_erro, @duracao_proc_em_seg, " +
						"@nome_arq_retorno, @caminho_arq_retorno, @st_processamento, @msg_erro_processamento) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@dt_processamento", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@dt_hr_processamento", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@usuario_processamento", SqlDbType.VarChar, 10);
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
			cmInsere.Parameters.Add("@qtde_total_registros", SqlDbType.Int);
			cmInsere.Parameters.Add("@qtde_registros_ok", SqlDbType.Int);
			cmInsere.Parameters.Add("@qtde_registros_erro", SqlDbType.Int);
			cmInsere.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmInsere.Parameters.Add("@nome_arq_retorno", SqlDbType.VarChar, 40);
			cmInsere.Parameters.Add("@caminho_arq_retorno", SqlDbType.VarChar, 1024);
			cmInsere.Parameters.Add("@st_processamento", SqlDbType.SmallInt);
			cmInsere.Parameters.Add("@msg_erro_processamento", SqlDbType.VarChar, 1024);
			cmInsere.Prepare();
			#endregion

			#region [ cmAtualizaDuracaoProcessamento ]
			strSql = "UPDATE t_SERASA_ARQ_RETORNO_NORMAL " +
					 "SET duracao_proc_em_seg = @duracao_proc_em_seg " +
						 "WHERE id = @id ";
			cmAtualizaDuracaoProcessamento = BD.criaSqlCommand();
			cmAtualizaDuracaoProcessamento.CommandText = strSql;
			cmAtualizaDuracaoProcessamento.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmAtualizaDuracaoProcessamento.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaDuracaoProcessamento.Prepare();
			#endregion

			#region [ cmAtualizaStatusProcessamento ]
			strSql = "UPDATE t_SERASA_ARQ_RETORNO_NORMAL " +
					 "SET st_processamento = @st_processamento, " +
					 "msg_erro_processamento = @msg_erro_processamento " +
						 "WHERE id = @id ";
			cmAtualizaStatusProcessamento = BD.criaSqlCommand();
			cmAtualizaStatusProcessamento.CommandText = strSql;
			cmAtualizaStatusProcessamento.Parameters.Add("@st_processamento", SqlDbType.SmallInt);
			cmAtualizaStatusProcessamento.Parameters.Add("@msg_erro_processamento", SqlDbType.VarChar, 1024);
			cmAtualizaStatusProcessamento.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaStatusProcessamento.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   DateTime dt_processamento,
								   DateTime dt_hr_processamento,
								   String usuario_processamento,
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
								   int qtde_total_registros,
								   int qtde_registros_ok,
								   int qtde_registros_erro,
								   int duracao_proc_em_seg,
								   String nome_arq_retorno,
								   String caminho_arq_retorno,
								   int st_processamento,
								   String msg_erro_processamento)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_ARQ_RETORNO_NORMAL";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@dt_processamento"].Value = dt_processamento.Date;
			cmInsere.Parameters["@dt_hr_processamento"].Value = dt_hr_processamento;
			cmInsere.Parameters["@usuario_processamento"].Value = usuario_processamento;
			cmInsere.Parameters["@cnpj_empresa_conveniada"].Value = cnpj_empresa_conveniada;
			cmInsere.Parameters["@s_periodo_inicio"].Value = s_periodo_inicio;
			cmInsere.Parameters["@dt_periodo_inicio"].Value = dt_periodo_inicio;
			cmInsere.Parameters["@s_periodo_termino"].Value = s_periodo_termino;
			cmInsere.Parameters["@dt_periodo_termino"].Value = dt_periodo_termino;
			cmInsere.Parameters["@periodicidade_remessa"].Value = periodicidade_remessa;
			cmInsere.Parameters["@reservado_serasa_1"].Value = reservado_serasa_1;
			cmInsere.Parameters["@num_id_grupo_relato_segmento"].Value = num_id_grupo_relato_segmento;
			cmInsere.Parameters["@id_versao_layout"].Value = id_versao_layout;
			cmInsere.Parameters["@num_versao_layout"].Value = num_versao_layout;
			cmInsere.Parameters["@qtde_total_registros"].Value = qtde_total_registros;
			cmInsere.Parameters["@qtde_registros_ok"].Value = qtde_registros_ok;
			cmInsere.Parameters["@qtde_registros_erro"].Value = qtde_registros_erro;
			cmInsere.Parameters["@duracao_proc_em_seg"].Value = duracao_proc_em_seg;
			cmInsere.Parameters["@nome_arq_retorno"].Value = nome_arq_retorno;
			cmInsere.Parameters["@caminho_arq_retorno"].Value = caminho_arq_retorno;
			cmInsere.Parameters["@st_processamento"].Value = st_processamento;

			if (msg_erro_processamento == null)
			{
				cmInsere.Parameters["@msg_erro_processamento"].Value = DBNull.Value;
			}
			else if (msg_erro_processamento.Length > 1024)
			{
				cmInsere.Parameters["@msg_erro_processamento"].Value = msg_erro_processamento.Substring(0, 1024);
			}
			else
			{
				cmInsere.Parameters["@msg_erro_processamento"].Value = msg_erro_processamento;
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

		#region [ atualizaDuracaoProcessamento ]
		public static bool atualizaDuracaoProcessamento(int duracao_proc_em_seg,
															int id)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_ARQ_RETORNO_NORMAL";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaDuracaoProcessamento.Parameters["@duracao_proc_em_seg"].Value = duracao_proc_em_seg;
			cmAtualizaDuracaoProcessamento.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaDuracaoProcessamento);
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

		#region [ atualizaStatusProcessamento ]
		public static bool atualizaStatusProcessamento(int st_processamento,
														String msg_erro_processamento,
														int id)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_ARQ_RETORNO_NORMAL";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaStatusProcessamento.Parameters["@st_processamento"].Value = st_processamento;

			if (msg_erro_processamento == null)
			{
				cmAtualizaStatusProcessamento.Parameters["@msg_erro_processamento"].Value = DBNull.Value;
			}
			else if (msg_erro_processamento.Length > 1024)
			{
				cmAtualizaStatusProcessamento.Parameters["@msg_erro_processamento"].Value = msg_erro_processamento.Substring(0, 1024);
			}
			else
			{
				cmAtualizaStatusProcessamento.Parameters["@msg_erro_processamento"].Value = msg_erro_processamento;
			}

			cmAtualizaStatusProcessamento.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaStatusProcessamento);
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

		#region [selecionaDatasParaCombobox]
		public static DataTable selecionaDatasParaCombobox()
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT DISTINCT " +
						"t2.id, " +
						"t2.dt_hr_processamento " +
					 "FROM t_SERASA_TITULO_MOVIMENTO t1 " +
					 "INNER JOIN t_SERASA_ARQ_RETORNO_NORMAL t2 " +
					 "ON t1.id_serasa_arq_retorno_normal = t2.id " +
					 "WHERE t2.st_processamento = 2 " +
						"AND t1.st_enviado_serasa = 0 " +
						"AND t1.st_processado_serasa_sucesso = 0 " +
						"AND t1.st_editado_manual = 1 " +
					 "ORDER BY t2.id, t2.dt_hr_processamento ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);

			return dtbResultado;
		}
		#endregion

		#region [verificaSeArquivoJaFoiCarregadoAntes]
		public static bool verificaSeArquivoJaFoiCarregadoAntes(String nome_arq_retorno,
																String s_periodo_inicio,
																String s_periodo_termino,
																String cnpj_empresa)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			bool arquivoJaCarregado = true;
			#endregion

			cmCommand = BD.criaSqlCommand();

			strSql = "SELECT COUNT(*) " +
					 "FROM t_SERASA_ARQ_RETORNO_NORMAL " +
					 "WHERE " +
						"(s_periodo_inicio = @s_periodo_inicio) " +
						"AND (s_periodo_termino = @s_periodo_termino) " +
						"AND ( (cnpj_empresa_conveniada = @cnpj_empresa_conveniada) OR (nome_arq_retorno = @nome_arq_retorno) ) " +
						"AND (st_processamento = 2)";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.Add("@nome_arq_retorno", SqlDbType.VarChar, 40);
			cmCommand.Parameters.Add("@s_periodo_inicio", SqlDbType.VarChar, 8);
			cmCommand.Parameters.Add("@s_periodo_termino", SqlDbType.VarChar, 8);
			cmCommand.Parameters.Add("@cnpj_empresa_conveniada", SqlDbType.VarChar, 14);
			cmCommand.Parameters["@nome_arq_retorno"].Value = nome_arq_retorno;
			cmCommand.Parameters["@s_periodo_inicio"].Value = s_periodo_inicio;
			cmCommand.Parameters["@s_periodo_termino"].Value = s_periodo_termino;
			cmCommand.Parameters["@cnpj_empresa_conveniada"].Value = cnpj_empresa;

			int ret = BD.readToInt(cmCommand.ExecuteScalar());

			if (ret == 0)
			{
				arquivoJaCarregado = false;
			}

			return arquivoJaCarregado;
		}
		#endregion
	}
}
