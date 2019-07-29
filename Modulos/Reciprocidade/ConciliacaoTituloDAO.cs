#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class ConciliacaoTituloDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsere;
		private static SqlCommand cmTrataTitulo;
		private static SqlCommand cmAtualizaArqConciliacaoOutputEStEnviadoSerasa;
		#endregion

		#region [ Construtor estático ]
		static ConciliacaoTituloDAO()
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
			strSql = "INSERT INTO t_SERASA_CONCILIACAO_TITULO " +
						"(id, id_serasa_arq_conciliacao_input, st_titulo_tratado_manual, dt_tratado_manual, dt_hr_tratado_manual, usuario_tratado_manual, st_enviado_serasa, " +
						" id_serasa_arq_conciliacao_output, id_registro_dados, cnpj_cliente, tipo_dados, num_titulo, s_data_emissao, dt_data_emissao, s_valor_titulo_original, " +
						" vl_valor_titulo_original, s_valor_titulo_editado, vl_valor_titulo_editado, s_data_vencto_original, dt_data_vencto_original, s_data_vencto_editado, " +
						" dt_data_vencto_editado, s_data_pagto_original, dt_data_pagto_original, s_data_pagto_editado, dt_data_pagto_editado, indicador_num_titulo_estendido, " +
						" num_titulo_estendido) " +
					 "VALUES " +
						"(@id, @id_serasa_arq_conciliacao_input, @st_titulo_tratado_manual, @dt_tratado_manual, @dt_hr_tratado_manual, @usuario_tratado_manual, @st_enviado_serasa, " +
						" @id_serasa_arq_conciliacao_output, @id_registro_dados, @cnpj_cliente, @tipo_dados, @num_titulo, @s_data_emissao, @dt_data_emissao, @s_valor_titulo_original, " +
						" @vl_valor_titulo_original, @s_valor_titulo_editado, @vl_valor_titulo_editado, @s_data_vencto_original, @dt_data_vencto_original, @s_data_vencto_editado, " +
						" @dt_data_vencto_editado, @s_data_pagto_original, @dt_data_pagto_original, @s_data_pagto_editado, @dt_data_pagto_editado, @indicador_num_titulo_estendido, " +
						" @num_titulo_estendido) ";

			cmInsere = BD.criaSqlCommand();
			cmInsere.CommandText = strSql;
			cmInsere.Parameters.Add("@id", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_serasa_arq_conciliacao_input", SqlDbType.Int);
			cmInsere.Parameters.Add("@st_titulo_tratado_manual", SqlDbType.TinyInt);
			cmInsere.Parameters.Add("@dt_tratado_manual", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@dt_hr_tratado_manual", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@usuario_tratado_manual", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@st_enviado_serasa", SqlDbType.TinyInt);
			cmInsere.Parameters.Add("@id_serasa_arq_conciliacao_output", SqlDbType.Int);
			cmInsere.Parameters.Add("@id_registro_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@cnpj_cliente", SqlDbType.VarChar, 14);
			cmInsere.Parameters.Add("@tipo_dados", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_titulo", SqlDbType.VarChar, 10);
			cmInsere.Parameters.Add("@s_data_emissao", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_data_emissao", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@s_valor_titulo_original", SqlDbType.VarChar, 13);
			cmInsere.Parameters.Add("@vl_valor_titulo_original", SqlDbType.Money);
			cmInsere.Parameters.Add("@s_valor_titulo_editado", SqlDbType.VarChar, 13);
			cmInsere.Parameters.Add("@vl_valor_titulo_editado", SqlDbType.Money);
			cmInsere.Parameters.Add("@s_data_vencto_original", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_data_vencto_original", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@s_data_vencto_editado", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_data_vencto_editado", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@s_data_pagto_original", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_data_pagto_original", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@s_data_pagto_editado", SqlDbType.VarChar, 8);
			cmInsere.Parameters.Add("@dt_data_pagto_editado", SqlDbType.DateTime);
			cmInsere.Parameters.Add("@indicador_num_titulo_estendido", SqlDbType.VarChar, 2);
			cmInsere.Parameters.Add("@num_titulo_estendido", SqlDbType.VarChar, 32);
			cmInsere.Prepare();
			#endregion

			#region [ cmTrataTitulo ]
			strSql = "UPDATE t_SERASA_CONCILIACAO_TITULO " +
					 "SET st_titulo_tratado_manual = 1, " +
						 "dt_tratado_manual = @dt_tratado_manual, " +
						 "dt_hr_tratado_manual = @dt_hr_tratado_manual, " +
						 "usuario_tratado_manual = @usuario_tratado_manual, " +
						 "s_valor_titulo_editado = @s_valor_titulo_editado, " +
						 "vl_valor_titulo_editado = @vl_valor_titulo_editado, " +
						 "s_data_vencto_editado = @s_data_vencto_editado, " +
						 "dt_data_vencto_editado = @dt_data_vencto_editado, " +
						 "s_data_pagto_editado = @s_data_pagto_editado, " +
						 "dt_data_pagto_editado = @dt_data_pagto_editado " +
					 "WHERE id = @id ";

			cmTrataTitulo = BD.criaSqlCommand();
			cmTrataTitulo.CommandText = strSql;
			cmTrataTitulo.Parameters.Add("@dt_tratado_manual", SqlDbType.DateTime);
			cmTrataTitulo.Parameters.Add("@dt_hr_tratado_manual", SqlDbType.DateTime);
			cmTrataTitulo.Parameters.Add("@usuario_tratado_manual", SqlDbType.VarChar, 10);
			cmTrataTitulo.Parameters.Add("@s_valor_titulo_editado", SqlDbType.VarChar, 13);
			cmTrataTitulo.Parameters.Add("@vl_valor_titulo_editado", SqlDbType.Money);
			cmTrataTitulo.Parameters.Add("@s_data_vencto_editado", SqlDbType.VarChar, 8);
			cmTrataTitulo.Parameters.Add("@dt_data_vencto_editado", SqlDbType.DateTime);
			cmTrataTitulo.Parameters.Add("@s_data_pagto_editado", SqlDbType.VarChar, 8);
			cmTrataTitulo.Parameters.Add("@dt_data_pagto_editado", SqlDbType.DateTime);
			cmTrataTitulo.Parameters.Add("@id", SqlDbType.Int);
			cmTrataTitulo.Prepare();
			#endregion

			#region [ cmAtualizaArqConciliacaoOutputEStEnviadoSerasa ]
			strSql = "UPDATE t_SERASA_CONCILIACAO_TITULO " +
					 "SET id_serasa_arq_conciliacao_output = @id_serasa_arq_conciliacao_output, " +
						 "st_enviado_serasa = @st_enviado_serasa " +
						 "WHERE id = @id ";

			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa = BD.criaSqlCommand();
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.CommandText = strSql;
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Parameters.Add("@id_serasa_arq_conciliacao_output", SqlDbType.Int);
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Parameters.Add("@st_enviado_serasa", SqlDbType.TinyInt);
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(int id,
								   int id_serasa_arq_conciliacao_input,
								   int st_titulo_tratado_manual,
								   DateTime dt_tratado_manual,
								   DateTime dt_hr_tratado_manual,
								   String usuario_tratado_manual,
								   int st_enviado_serasa,
								   int id_serasa_arq_conciliacao_output,
								   String id_registro_dados,
								   String cnpj_cliente,
								   String tipo_dados,
								   String num_titulo,
								   String s_data_emissao,
								   DateTime dt_data_emissao,
								   String s_valor_titulo_original,
								   Decimal vl_valor_titulo_original,
								   String s_valor_titulo_editado,
								   Decimal vl_valor_titulo_editado,
								   String s_data_vencto_original,
								   DateTime dt_data_vencto_original,
								   String s_data_vencto_editado,
								   DateTime dt_data_vencto_editado,
								   String s_data_pagto_original,
								   DateTime dt_data_pagto_original,
								   String s_data_pagto_editado,
								   DateTime dt_data_pagto_editado,
								   String indicador_num_titulo_estendido,
								   String num_titulo_estendido)
		{
			#region [Declarações]
			String strOperacao = "INSERT t_SERASA_CONCILIACAO_TITULO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsere.Parameters["@id"].Value = id;
			cmInsere.Parameters["@id_serasa_arq_conciliacao_input"].Value = id_serasa_arq_conciliacao_input;
			cmInsere.Parameters["@st_titulo_tratado_manual"].Value = st_titulo_tratado_manual;

			if (dt_tratado_manual == DateTime.MinValue)
			{
				cmInsere.Parameters["@dt_tratado_manual"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@dt_tratado_manual"].Value = dt_tratado_manual.Date;
			}

			if (dt_hr_tratado_manual == DateTime.MinValue)
			{
				cmInsere.Parameters["@dt_hr_tratado_manual"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@dt_hr_tratado_manual"].Value = dt_hr_tratado_manual;
			}

			if (usuario_tratado_manual == null)
			{
				cmInsere.Parameters["@usuario_tratado_manual"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@usuario_tratado_manual"].Value = usuario_tratado_manual;
			}

			cmInsere.Parameters["@st_enviado_serasa"].Value = st_enviado_serasa;
			cmInsere.Parameters["@id_serasa_arq_conciliacao_output"].Value = id_serasa_arq_conciliacao_output;
			cmInsere.Parameters["@id_registro_dados"].Value = id_registro_dados;
			cmInsere.Parameters["@cnpj_cliente"].Value = cnpj_cliente;
			cmInsere.Parameters["@tipo_dados"].Value = tipo_dados;
			cmInsere.Parameters["@num_titulo"].Value = num_titulo;
			cmInsere.Parameters["@s_data_emissao"].Value = s_data_emissao;
			cmInsere.Parameters["@dt_data_emissao"].Value = dt_data_emissao;
			cmInsere.Parameters["@s_valor_titulo_original"].Value = s_valor_titulo_original;
			cmInsere.Parameters["@vl_valor_titulo_original"].Value = vl_valor_titulo_original;

			if (s_valor_titulo_editado == null)
			{
				cmInsere.Parameters["@s_valor_titulo_editado"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@s_valor_titulo_editado"].Value = s_valor_titulo_editado;
			}

			cmInsere.Parameters["@vl_valor_titulo_editado"].Value = vl_valor_titulo_editado;
			cmInsere.Parameters["@s_data_vencto_original"].Value = s_data_vencto_original;
			cmInsere.Parameters["@dt_data_vencto_original"].Value = dt_data_vencto_original;

			if (s_data_vencto_editado == null)
			{
				cmInsere.Parameters["@s_data_vencto_editado"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@s_data_vencto_editado"].Value = s_data_vencto_editado;
			}

			if (dt_data_vencto_editado == DateTime.MinValue)
			{
				cmInsere.Parameters["@dt_data_vencto_editado"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@dt_data_vencto_editado"].Value = dt_data_vencto_editado;
			}

			cmInsere.Parameters["@s_data_pagto_original"].Value = s_data_pagto_original;

			if (dt_data_pagto_original == DateTime.MinValue)
			{
				cmInsere.Parameters["@dt_data_pagto_original"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@dt_data_pagto_original"].Value = dt_data_pagto_original;
			}

			if (s_data_pagto_editado == null)
			{
				cmInsere.Parameters["@s_data_pagto_editado"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@s_data_pagto_editado"].Value = s_data_pagto_editado;
			}

			if (dt_data_pagto_editado == DateTime.MinValue)
			{
				cmInsere.Parameters["@dt_data_pagto_editado"].Value = DBNull.Value;
			}
			else
			{
				cmInsere.Parameters["@dt_data_pagto_editado"].Value = dt_data_pagto_editado;
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

		#region [ obtemArquivoRetornoParaTratamento ]
		public static bool obtemArquivoRetornoParaTratamento(out int id_serasa_arq_conciliacao_input, out DateTime data_final_periodo, out string msg_erro)
		{
			#region [Declarações]
			int id_aux;
			Object objResultado;
			String strSql;
			String s_data_final_periodo;
			SqlCommand cmCommand;
			#endregion

			#region [ Inicialização ]
			id_serasa_arq_conciliacao_input = 0;
			data_final_periodo = DateTime.MinValue;
			msg_erro = "";
			#endregion

			cmCommand = BD.criaSqlCommand();

			#region [ Localiza último arquivo de conciliação que foi carregado ]
			strSql = "SELECT" +
						" Coalesce(Max(id), 0) AS id" +
					" FROM t_SERASA_ARQ_CONCILIACAO_INPUT";
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado == null)
			{
				msg_erro = "Não há nenhum arquivo de conciliação disponível!!";
				return false;
			}
			else
			{
				id_aux = BD.readToInt(objResultado);
				if (id_aux == 0)
				{
					msg_erro = "Não há nenhum arquivo de conciliação disponível!!";
					return false;
				}
			}
			#endregion

			#region [ Obtém data final do período e verifica se o arquivo já foi enviado de volta ]
			strSql = "SELECT" +
						" s_data_final_periodo" +
					" FROM t_SERASA_ARQ_CONCILIACAO_INPUT" +
					" WHERE" +
						" (id = " + id_aux.ToString() + ")" +
						" AND " +
						"(" +
							"id NOT IN " +
							"(" +
								"SELECT TOP 1" +
									" id_serasa_arq_conciliacao_input" +
								" FROM t_SERASA_CONCILIACAO_TITULO" +
								" WHERE" +
									" (id_serasa_arq_conciliacao_input = " + id_aux.ToString() + ")" +
									" AND (st_enviado_serasa = 1)" +
							")" +
						")";
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado == null)
			{
				msg_erro = "O último arquivo de conciliação carregado já foi tratado!!";
				return false;
			}
			else
			{
				id_serasa_arq_conciliacao_input = id_aux;
				s_data_final_periodo = BD.readToString(objResultado);
				data_final_periodo = Global.converteYyyyMmDdSemSeparadorParaDateTime(s_data_final_periodo);
				return true;
			}
			#endregion
		}
		#endregion

		#region [selecionaBoletosParaTratamento]
		public static DataTable selecionaBoletosParaTratamento(int id_serasa_arq_conciliacao_input)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbTitulo = new DataTable();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_CONCILIACAO_TITULO" +
					" WHERE" +
						" (id_serasa_arq_conciliacao_input = " + id_serasa_arq_conciliacao_input.ToString() + ")" +
					" ORDER BY" +
						" dt_data_vencto_original," +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbTitulo);

			return dtbTitulo;
		}
		#endregion

		#region [ trataTitulo ]
		public static bool trataTitulo(DateTime dt_data_vencto_editado,
									   Decimal vl_valor_titulo_editado,
									   DateTime dt_data_pagto_editado,
									   int id,
									   bool blnTituloExcluido)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_CONCILIACAO_TITULO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmTrataTitulo.Parameters["@dt_tratado_manual"].Value = DateTime.Now.Date;
			cmTrataTitulo.Parameters["@dt_hr_tratado_manual"].Value = DateTime.Now;
			cmTrataTitulo.Parameters["@usuario_tratado_manual"].Value = Global.Usuario.usuario;

			if (blnTituloExcluido)
			{
				cmTrataTitulo.Parameters["@s_valor_titulo_editado"].Value = "9999999999999";
			}
			else
			{
				cmTrataTitulo.Parameters["@s_valor_titulo_editado"].Value = Global.formataMoedaSemSeparador(vl_valor_titulo_editado, 13);
			}

			cmTrataTitulo.Parameters["@vl_valor_titulo_editado"].Value = vl_valor_titulo_editado;
			cmTrataTitulo.Parameters["@s_data_vencto_editado"].Value = Global.formataDataYyyyMmDdSemSeparador(dt_data_vencto_editado);
			cmTrataTitulo.Parameters["@dt_data_vencto_editado"].Value = dt_data_vencto_editado;

			if (dt_data_pagto_editado == DateTime.MinValue)
			{
				cmTrataTitulo.Parameters["@s_data_pagto_editado"].Value = "";
				cmTrataTitulo.Parameters["@dt_data_pagto_editado"].Value = DBNull.Value;
			}
			else
			{
				cmTrataTitulo.Parameters["@s_data_pagto_editado"].Value = Global.formataDataYyyyMmDdSemSeparador(dt_data_pagto_editado);
				cmTrataTitulo.Parameters["@dt_data_pagto_editado"].Value = dt_data_pagto_editado;
			}

			cmTrataTitulo.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmTrataTitulo);
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

		#region [selecionaBoletosParaRemessa]
		public static DataSet selecionaBoletosParaRemessa(int id_serasa_arq_conciliacao_input)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbBoleto = new DataTable("DtbBoleto");
			DataTable dtbArqConciliacaoInput = new DataTable("dtbArqConciliacaoInput");
			DataRelation drlArqConciliacaoInputBoleto;
			DataSet dataset = new DataSet();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"t1.* " +
					"FROM t_SERASA_CONCILIACAO_TITULO t1 " +
						"INNER JOIN t_SERASA_ARQ_CONCILIACAO_INPUT t2 " +
							"ON t1.id_serasa_arq_conciliacao_input = t2.id " +
					"WHERE " +
						"(t2.st_processamento = 2) " +
						"AND (t1.st_enviado_serasa = 0) " +
						"AND (t2.id = @id) " +
					"ORDER BY " +
						"t1.id";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.Add("@id", SqlDbType.Int);
			cmCommand.Parameters["@id"].Value = id_serasa_arq_conciliacao_input;
			daDataAdapter.Fill(dtbBoleto);
			dataset.Tables.Add(dtbBoleto);

			strSql = "SELECT " +
						"* " +
					 "FROM t_SERASA_ARQ_CONCILIACAO_INPUT " +
					 "WHERE (id = @id)";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbArqConciliacaoInput);
			dataset.Tables.Add(dtbArqConciliacaoInput);

			drlArqConciliacaoInputBoleto = new DataRelation("dtbArqConciliacaoInput_dtbBoleto", dataset.Tables["dtbArqConciliacaoInput"].Columns["id"], dataset.Tables["DtbBoleto"].Columns["id_serasa_arq_conciliacao_input"]);
			dataset.Relations.Add(drlArqConciliacaoInputBoleto);

			return dataset;
		}
		#endregion

		#region [ atualizaArqConciliacaoOutputEStEnviadoSerasa ]
		public static bool atualizaArqConciliacaoOutputEStEnviadoSerasa(int id_serasa_arq_conciliacao_output,
																		int st_enviado_serasa,
																		int id)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_CONCILIACAO_TITULO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Parameters["@id_serasa_arq_conciliacao_output"].Value = id_serasa_arq_conciliacao_output;
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Parameters["@st_enviado_serasa"].Value = st_enviado_serasa;
			cmAtualizaArqConciliacaoOutputEStEnviadoSerasa.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaArqConciliacaoOutputEStEnviadoSerasa);
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

		#region [ selecionaBoletosComInfoPagtoDivergentes ]
		public static DataSet selecionaBoletosComInfoPagtoDivergentes()
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbBoleto = new DataTable("DtbBoleto");
			DataSet dataset = new DataSet();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT" +
						" tSCT.id," +
						" tSCT.cnpj_cliente," +
						" tSCT.num_titulo_estendido," +
						" tSCT.dt_data_emissao," +
						" tSCT.dt_data_vencto_original," +
						" tSCT.vl_valor_titulo_original," +
						" tFFC.dt_competencia," +
						" tFFC.descricao" +
					" FROM t_SERASA_CONCILIACAO_TITULO tSCT" +
						" INNER JOIN t_FIN_BOLETO_ITEM tFBI" +
							" ON (tFBI.nosso_numero_com_dv = tSCT.num_titulo_estendido)" +
						" INNER JOIN t_FIN_BOLETO tFB" +
							" ON (tFB.id = tFBI.id_boleto) AND (tFB.num_inscricao_sacado = tSCT.cnpj_cliente)" +
						" INNER JOIN t_FIN_FLUXO_CAIXA tFFC" +
							" ON (tFFC.ctrl_pagto_id_parcela = tFBI.id) AND (tFFC.ctrl_pagto_modulo = " + Global.Cte.FIN.CtrlPagtoModulo.BOLETO + ")" +
					" WHERE" +
						" (id_serasa_arq_conciliacao_input = (SELECT MAX(id) FROM t_SERASA_ARQ_CONCILIACAO_INPUT))" +
						" AND (tSCT.dt_data_vencto_original <= (SELECT MAX(Convert(datetime, s_data_final_periodo, 112)) FROM t_SERASA_ARQ_CONCILIACAO_INPUT))" +
						" AND (tSCT.dt_data_pagto_original IS NULL)" +
						" AND (tFFC.ctrl_pagto_modulo = " + Global.Cte.FIN.CtrlPagtoModulo.BOLETO + ")" +
						" AND (tFFC.st_sem_efeito = 0)" +
						" AND (tFFC.st_confirmacao_pendente = 0)" +
						" AND NOT EXISTS (SELECT 1 FROM t_SERASA_TITULO_MOVIMENTO WHERE (nosso_numero_com_dv = tSCT.num_titulo_estendido) AND (identificacao_ocorrencia_boleto IN ('06','15','17','09','10')) AND (st_envio_serasa_cancelado = 0))" +
					" ORDER BY" +
						" tSCT.dt_data_vencto_original," +
						" tSCT.vl_valor_titulo_original";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbBoleto);
			dataset.Tables.Add(dtbBoleto);

			return dataset;
		}
		#endregion
	}
}
