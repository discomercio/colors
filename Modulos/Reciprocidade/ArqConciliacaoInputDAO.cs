#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class ArqConciliacaoInputDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		private static SqlCommand cmAtualizaDuracaoProcessamento;
		private static SqlCommand cmAtualizaStatusProcessamento;
		#endregion

		#region [ Construtor estático ]
		static ArqConciliacaoInputDAO()
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
			strSql = "INSERT INTO t_SERASA_ARQ_CONCILIACAO_INPUT " +
						"(id, dt_processamento, dt_hr_processamento, usuario_processamento, s_data_final_periodo, qtde_registros, duracao_proc_em_seg, " +
						"cnpj_empresa, nome_arq_retorno, caminho_arq_retorno, st_processamento, linha_header, linha_trailler, msg_erro_processamento) " +
					 "VALUES " +
						"(@id, @dt_processamento, @dt_hr_processamento, @usuario_processamento, @s_data_final_periodo, @qtde_registros, @duracao_proc_em_seg, " +
						"@cnpj_empresa, @nome_arq_retorno, @caminho_arq_retorno, @st_processamento, @linha_header, @linha_trailler, @msg_erro_processamento) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@dt_processamento", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@dt_hr_processamento", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@usuario_processamento", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@s_data_final_periodo", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@qtde_registros", SqlDbType.Int);
			cmInsere.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmInsere.Parameters.Add("@cnpj_empresa", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@nome_arq_retorno", SqlDbType.VarChar, 40);
			cmInsere.Parameters.Add("@caminho_arq_retorno", SqlDbType.VarChar, 1024);
			cmInsere.Parameters.Add("@st_processamento", SqlDbType.SmallInt);
			cmInsere.Parameters.Add("@linha_header", SqlDbType.VarChar, 240);
			cmInsere.Parameters.Add("@linha_trailler", SqlDbType.VarChar, 240);
			cmInsere.Parameters.Add("@msg_erro_processamento", SqlDbType.VarChar, 1024);
			cmInsere.Prepare();
			#endregion

			#region [ cmAtualizaDuracaoProcessamento ]
			strSql = "UPDATE t_SERASA_ARQ_CONCILIACAO_INPUT " +
					 "SET duracao_proc_em_seg = @duracao_proc_em_seg " +
						 "WHERE id = @id ";

			cmAtualizaDuracaoProcessamento = BD.criaSqlCommand();
			cmAtualizaDuracaoProcessamento.CommandText = strSql;
			cmAtualizaDuracaoProcessamento.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmAtualizaDuracaoProcessamento.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaDuracaoProcessamento.Prepare();
			#endregion

			#region [ cmAtualizaStatusProcessamento ]
			strSql = "UPDATE t_SERASA_ARQ_CONCILIACAO_INPUT " +
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
								   String s_data_final_periodo,
								   int qtde_registros,
								   int duracao_proc_em_seg,
								   String cnpj_empresa,
								   String nome_arq_retorno,
								   String caminho_arq_retorno,
								   int st_processamento,
								   String linha_header,
								   String linha_trailler,
								   String msg_erro_processamento)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_ARQ_CONCILIACAO_INPUT";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@dt_processamento"].Value = dt_processamento.Date;
			cmInsere.Parameters["@dt_hr_processamento"].Value = dt_hr_processamento;
			cmInsere.Parameters["@usuario_processamento"].Value = usuario_processamento;
			cmInsere.Parameters["@s_data_final_periodo"].Value = s_data_final_periodo;
			cmInsere.Parameters["@qtde_registros"].Value = qtde_registros;
			cmInsere.Parameters["@duracao_proc_em_seg"].Value = duracao_proc_em_seg;
			cmInsere.Parameters["@cnpj_empresa"].Value = cnpj_empresa;
			cmInsere.Parameters["@nome_arq_retorno"].Value = nome_arq_retorno;
			cmInsere.Parameters["@caminho_arq_retorno"].Value = caminho_arq_retorno;
			cmInsere.Parameters["@st_processamento"].Value = st_processamento;
			cmInsere.Parameters["@linha_header"].Value = linha_header;
			cmInsere.Parameters["@linha_trailler"].Value = linha_trailler;

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
			String strOperacao = "UPDATE t_SERASA_ARQ_CONCILIACAO_INPUT";
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
			String strOperacao = "UPDATE t_SERASA_ARQ_CONCILIACAO_INPUT";
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

			strSql = "SELECT DISTINCT TOP 1 " +
						"t2.id, " +
						"t2.dt_hr_processamento " +
					"FROM t_SERASA_CONCILIACAO_TITULO t1 " +
						"INNER JOIN t_SERASA_ARQ_CONCILIACAO_INPUT t2 " +
							"ON t1.id_serasa_arq_conciliacao_input = t2.id " +
					"WHERE " +
						"(t2.st_processamento = 2) " +
						"AND (t1.st_enviado_serasa = 0) " +
					"ORDER BY " +
						"t2.id DESC, " +
						"t2.dt_hr_processamento DESC";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);

			return dtbResultado;
		}
		#endregion

		#region [verificaSeArquivoJaFoiCarregadoAntes]
		public static bool verificaSeArquivoJaFoiCarregadoAntes(String nome_arq_retorno,
																String s_data_final_periodo,
																String cnpj_empresa)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			bool arquivoJaCarregado = true;
			#endregion

			cmCommand = BD.criaSqlCommand();

			strSql = "SELECT COUNT(*) " +
					 "FROM t_SERASA_ARQ_CONCILIACAO_INPUT " +
					 "WHERE " +
						"(s_data_final_periodo = @s_data_final_periodo) " +
						"AND ( (cnpj_empresa = @cnpj_empresa) OR (nome_arq_retorno = @nome_arq_retorno) ) " +
						"AND (st_processamento = 2)";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.Add("@nome_arq_retorno", SqlDbType.VarChar, 40);
			cmCommand.Parameters.Add("@s_data_final_periodo", SqlDbType.VarChar, 8);
			cmCommand.Parameters.Add("@cnpj_empresa", SqlDbType.VarChar, 14);
			cmCommand.Parameters["@nome_arq_retorno"].Value = nome_arq_retorno;
			cmCommand.Parameters["@s_data_final_periodo"].Value = s_data_final_periodo;
			cmCommand.Parameters["@cnpj_empresa"].Value = cnpj_empresa;

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
