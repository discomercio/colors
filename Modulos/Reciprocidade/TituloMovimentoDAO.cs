#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Reciprocidade
{
	class TituloMovimentoDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmAtualizaStatusEnvioEIdArqRemessa;
		private static SqlCommand cmAtualizaTituloAposRetornoArquivo;
		private static SqlCommand cmTrataOcorrencia;
		private static SqlCommand cmCancelaEnvio;
		#endregion

		#region [ Construtor estático ]
		static TituloMovimentoDAO()
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

			#region [ cmAtualizaStatusEnvioEIdArqRemessa ]
			strSql = "UPDATE t_SERASA_TITULO_MOVIMENTO " +
					 "SET st_enviado_serasa = @st_enviado_serasa, " +
						 "id_serasa_arq_remessa_normal = @id_serasa_arq_remessa_normal " +
						 "WHERE id = @id_serasa_titulo_movimento ";

			cmAtualizaStatusEnvioEIdArqRemessa = BD.criaSqlCommand();
			cmAtualizaStatusEnvioEIdArqRemessa.CommandText = strSql;
			cmAtualizaStatusEnvioEIdArqRemessa.Parameters.Add("@st_enviado_serasa", SqlDbType.TinyInt);
			cmAtualizaStatusEnvioEIdArqRemessa.Parameters.Add("@id_serasa_arq_remessa_normal", SqlDbType.Int);
			cmAtualizaStatusEnvioEIdArqRemessa.Parameters.Add("@id_serasa_titulo_movimento", SqlDbType.Int);
			cmAtualizaStatusEnvioEIdArqRemessa.Prepare();
			#endregion

			#region [ cmAtualizaTituloAposRetornoArquivo ]
			strSql = "UPDATE t_SERASA_TITULO_MOVIMENTO " +
					 "SET st_retorno_serasa = @st_retorno_serasa, " +
						 "id_serasa_arq_retorno_normal = @id_serasa_arq_retorno_normal, " +
						 "st_processado_serasa_sucesso = @st_processado_serasa_sucesso, " +
						 "retorno_codigos_erro = @retorno_codigos_erro " +
						 "WHERE id = (select MAX(id) from t_SERASA_TITULO_MOVIMENTO where nosso_numero = @numeroTitulo) ";

			cmAtualizaTituloAposRetornoArquivo = BD.criaSqlCommand();
			cmAtualizaTituloAposRetornoArquivo.CommandText = strSql;
			cmAtualizaTituloAposRetornoArquivo.Parameters.Add("@st_retorno_serasa", SqlDbType.TinyInt);
			cmAtualizaTituloAposRetornoArquivo.Parameters.Add("@id_serasa_arq_retorno_normal", SqlDbType.Int);
			cmAtualizaTituloAposRetornoArquivo.Parameters.Add("@st_processado_serasa_sucesso", SqlDbType.TinyInt);
			cmAtualizaTituloAposRetornoArquivo.Parameters.Add("@retorno_codigos_erro", SqlDbType.VarChar, 90);
			cmAtualizaTituloAposRetornoArquivo.Parameters.Add("@numeroTitulo", SqlDbType.VarChar, 11);
			cmAtualizaTituloAposRetornoArquivo.Prepare();
			#endregion

			#region [ cmTrataOcorrencia ]
			strSql = "UPDATE t_SERASA_TITULO_MOVIMENTO " +
					 "SET nosso_numero = @nosso_numero, " +
						 "digito_nosso_numero = @digito_nosso_numero, " +
						 "dt_emissao = @dt_emissao, " +
						 "dt_vencto = @dt_vencto, " +
						 "vl_titulo = @vl_titulo, " +
						 "dt_pagto = @dt_pagto, " +
						 "vl_pago = @vl_pago, " +
						 "st_enviado_serasa = 0, " +
						 "st_retorno_serasa = 0, " +
						 "st_processado_serasa_sucesso = 0, " +
						 "st_editado_manual = 1, " +
						 "dt_editado_manual = @dt_editado_manual, " +
						 "dt_hr_editado_manual = @dt_hr_editado_manual, " +
						 "usuario_editado_manual = @usuario_editado_manual, " +
						 "qtde_vezes_editado_manual = @qtde_vezes_editado_manual " +
						 "WHERE id = @id ";

			cmTrataOcorrencia = BD.criaSqlCommand();
			cmTrataOcorrencia.CommandText = strSql;
			cmTrataOcorrencia.Parameters.Add("@nosso_numero", SqlDbType.VarChar, 11);
			cmTrataOcorrencia.Parameters.Add("@digito_nosso_numero", SqlDbType.VarChar, 1);
			cmTrataOcorrencia.Parameters.Add("@dt_emissao", SqlDbType.DateTime);
			cmTrataOcorrencia.Parameters.Add("@dt_vencto", SqlDbType.DateTime);
			cmTrataOcorrencia.Parameters.Add("@vl_titulo", SqlDbType.Money);
			cmTrataOcorrencia.Parameters.Add("@dt_pagto", SqlDbType.DateTime);
			cmTrataOcorrencia.Parameters.Add("@vl_pago", SqlDbType.Money);
			cmTrataOcorrencia.Parameters.Add("@dt_editado_manual", SqlDbType.DateTime);
			cmTrataOcorrencia.Parameters.Add("@dt_hr_editado_manual", SqlDbType.DateTime);
			cmTrataOcorrencia.Parameters.Add("@usuario_editado_manual", SqlDbType.VarChar, 10);
			cmTrataOcorrencia.Parameters.Add("@qtde_vezes_editado_manual", SqlDbType.Int);
			cmTrataOcorrencia.Parameters.Add("@id", SqlDbType.Int);
			cmTrataOcorrencia.Prepare();
			#endregion

			#region [ cmCancelaEnvio ]
			strSql = "UPDATE t_SERASA_TITULO_MOVIMENTO " +
					 "SET st_envio_serasa_cancelado = 1, " +
						 "dt_envio_serasa_cancelado = @dt_envio_serasa_cancelado, " +
						 "dt_hr_envio_serasa_cancelado = @dt_hr_envio_serasa_cancelado, " +
						 "usuario_envio_serasa_cancelado = @usuario_envio_serasa_cancelado " +
						 "WHERE id = @id ";

			cmCancelaEnvio = BD.criaSqlCommand();
			cmCancelaEnvio.CommandText = strSql;
			cmCancelaEnvio.Parameters.Add("@dt_envio_serasa_cancelado", SqlDbType.DateTime);
			cmCancelaEnvio.Parameters.Add("@dt_hr_envio_serasa_cancelado", SqlDbType.DateTime);
			cmCancelaEnvio.Parameters.Add("@usuario_envio_serasa_cancelado", SqlDbType.VarChar, 10);
			cmCancelaEnvio.Parameters.Add("@id", SqlDbType.Int);
			cmCancelaEnvio.Prepare();
			#endregion
		}
		#endregion

		#region [selecionaBoletosParaArqRemessa]
		public static DataSet selecionaBoletosParaArqRemessa(DateTime dtFinalPeriodo)
		{
			#region [Declarações]
			String strSqlTitulo;
			String strWhereTitulo;
			String strSqlCliente;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataRelation drlClienteBoleto;
			DataTable dtbBoleto = new DataTable("DtbBoleto");
			DataTable dtbCliente = new DataTable("DtbCliente");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strWhereTitulo = " (st_enviado_serasa = 0)" +
							" AND (st_editado_manual = 0)" +
							" AND (st_envio_serasa_cancelado = 0)" +
							" AND (dt_emissao <= " + Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtFinalPeriodo) + ")" +
							" AND (" +
								"(dt_pagto IS NULL)" +
								" OR " +
								"(dt_pagto <= " + Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtFinalPeriodo) + ")" +
								")";

			strSqlTitulo = "SELECT " +
								"*" +
							" FROM t_SERASA_TITULO_MOVIMENTO" +
							" WHERE" +
								strWhereTitulo +
							" ORDER BY" +
								" id";
			cmCommand.CommandText = strSqlTitulo;

			daDataAdapter.Fill(dtbBoleto);
			dsResultado.Tables.Add(dtbBoleto);

			strSqlCliente = "SELECT " +
								"*" +
							" FROM t_SERASA_CLIENTE" +
							" WHERE" +
								" id IN " +
									"(SELECT DISTINCT" +
										" id_serasa_cliente " +
									"FROM t_SERASA_TITULO_MOVIMENTO " +
									"WHERE " +
										strWhereTitulo +
									")";
			cmCommand.CommandText = strSqlCliente;
			daDataAdapter.Fill(dtbCliente);
			dsResultado.Tables.Add(dtbCliente);

			drlClienteBoleto = new DataRelation("dtbCliente_dtbBoleto", dsResultado.Tables["DtbCliente"].Columns["id"], dsResultado.Tables["DtbBoleto"].Columns["id_serasa_cliente"]);
			dsResultado.Relations.Add(drlClienteBoleto);

			return dsResultado;
		}
		#endregion

		#region [selecionaBoletosParaArqRemessaRetificacao]
		public static DataSet selecionaBoletosParaArqRemessaRetificacao(int id)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataRelation drlClienteBoleto;
			DataRelation drlArqRemessaBoleto;
			DataTable dtbBoleto = new DataTable("DtbBoleto");
			DataTable dtbCliente = new DataTable("DtbCliente");
			DataTable dtbArqRemessaNormal = new DataTable("DtbArqRemessaNormal");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"t1.* " +
					 "FROM t_SERASA_TITULO_MOVIMENTO t1 " +
					 "INNER JOIN t_SERASA_ARQ_RETORNO_NORMAL t2 " +
					 "ON t1.id_serasa_arq_retorno_normal = t2.id " +
					 "WHERE t2.st_processamento = 2 " +
						"AND t1.st_enviado_serasa = 0 " +
						"AND t1.st_envio_serasa_cancelado = 0 " +
						"AND t1.st_processado_serasa_sucesso = 0 " +
						"AND t1.st_editado_manual = 1 " +
						"AND t2.id = @id " +
					 "ORDER BY t1.id ";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.Add("@id", SqlDbType.Int);
			cmCommand.Parameters["@id"].Value = id;
			daDataAdapter.Fill(dtbBoleto);
			dsResultado.Tables.Add(dtbBoleto);

			strSql = "SELECT " +
					 "* " +
					 "FROM t_SERASA_CLIENTE ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbCliente);
			dsResultado.Tables.Add(dtbCliente);

			strSql = "SELECT " +
					 "* " +
					 "FROM t_SERASA_ARQ_REMESSA_NORMAL " +
					 "WHERE id = " + BD.readToInt(dtbBoleto.Rows[0]["id_serasa_arq_remessa_normal"]);

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbArqRemessaNormal);
			dsResultado.Tables.Add(dtbArqRemessaNormal);

			drlClienteBoleto = new DataRelation("dtbCliente_dtbBoleto", dsResultado.Tables["DtbCliente"].Columns["id"], dsResultado.Tables["DtbBoleto"].Columns["id_serasa_cliente"]);
			drlArqRemessaBoleto = new DataRelation("dtbArqRemessaNormal_dtbBoleto", dsResultado.Tables["DtbArqRemessaNormal"].Columns["id"], dsResultado.Tables["DtbBoleto"].Columns["id_serasa_arq_remessa_normal"]);
			dsResultado.Relations.Add(drlClienteBoleto);
			dsResultado.Relations.Add(drlArqRemessaBoleto);

			return dsResultado;
		}
		#endregion

		#region [ atualizaStatusEnvioEIdArqRemessa ]
		public static bool atualizaStatusEnvioEIdArqRemessa(int st_enviado_serasa,
															int id_serasa_arq_remessa_normal,
															int id_serasa_titulo_movimento)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_TITULO_MOVIMENTO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaStatusEnvioEIdArqRemessa.Parameters["@st_enviado_serasa"].Value = st_enviado_serasa;
			cmAtualizaStatusEnvioEIdArqRemessa.Parameters["@id_serasa_arq_remessa_normal"].Value = id_serasa_arq_remessa_normal;
			cmAtualizaStatusEnvioEIdArqRemessa.Parameters["@id_serasa_titulo_movimento"].Value = id_serasa_titulo_movimento;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaStatusEnvioEIdArqRemessa);
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

		#region [ atualizaTituloAposRetornoArquivo ]
		public static bool atualizaTituloAposRetornoArquivo(int st_retorno_serasa,
															int id_serasa_arq_retorno_normal,
															int st_processado_serasa_sucesso,
															String retorno_codigos_erro,
															String numeroTitulo)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_TITULO_MOVIMENTO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaTituloAposRetornoArquivo.Parameters["@st_retorno_serasa"].Value = st_retorno_serasa;
			cmAtualizaTituloAposRetornoArquivo.Parameters["@id_serasa_arq_retorno_normal"].Value = id_serasa_arq_retorno_normal;
			cmAtualizaTituloAposRetornoArquivo.Parameters["@st_processado_serasa_sucesso"].Value = st_processado_serasa_sucesso;
			cmAtualizaTituloAposRetornoArquivo.Parameters["@retorno_codigos_erro"].Value = retorno_codigos_erro;
			cmAtualizaTituloAposRetornoArquivo.Parameters["@numeroTitulo"].Value = numeroTitulo.Trim().Substring(0, numeroTitulo.Trim().Length - 1); //remove o DV

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaTituloAposRetornoArquivo);
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

		#region [selecionaBoletosParaTratamento]
		public static DataTable selecionaBoletosParaTratamento()
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbBoleto = new DataTable();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"* " +
					 "FROM t_SERASA_TITULO_MOVIMENTO " +
					 "WHERE st_enviado_serasa = 1 " +
						"AND st_retorno_serasa = 1 " +
						"AND st_processado_serasa_sucesso = 0 ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbBoleto);

			return dtbBoleto;
		}
		#endregion

		#region [ trataOcorrencia ]
		public static bool trataOcorrencia(String nosso_numero,
										   DateTime dt_emissao,
										   DateTime dt_vencto,
										   Decimal vl_titulo,
										   DateTime dt_pagto,
										   Decimal vl_pago,
										   int id)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_TITULO_MOVIMENTO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmTrataOcorrencia.Parameters["@nosso_numero"].Value = nosso_numero.Substring(0, nosso_numero.Length - 1);
			cmTrataOcorrencia.Parameters["@digito_nosso_numero"].Value = nosso_numero.Substring(nosso_numero.Length - 1, 1);
			cmTrataOcorrencia.Parameters["@dt_emissao"].Value = dt_emissao;
			cmTrataOcorrencia.Parameters["@dt_vencto"].Value = dt_vencto;
			cmTrataOcorrencia.Parameters["@vl_titulo"].Value = vl_titulo;

			if (dt_pagto == DateTime.MinValue)
			{
				cmTrataOcorrencia.Parameters["@dt_pagto"].Value = DBNull.Value;
			}
			else
			{
				cmTrataOcorrencia.Parameters["@dt_pagto"].Value = dt_pagto;
			}

			cmTrataOcorrencia.Parameters["@vl_pago"].Value = vl_pago;
			cmTrataOcorrencia.Parameters["@dt_editado_manual"].Value = DateTime.Now.Date;
			cmTrataOcorrencia.Parameters["@dt_hr_editado_manual"].Value = DateTime.Now;
			cmTrataOcorrencia.Parameters["@usuario_editado_manual"].Value = Global.Usuario.usuario;
			cmTrataOcorrencia.Parameters["@qtde_vezes_editado_manual"].Value = selecionaQtdeVezesEditado(id) + 1;
			cmTrataOcorrencia.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmTrataOcorrencia);
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

		#region [selecionaQtdeVezesEditado]
		public static int selecionaQtdeVezesEditado(int id_serasa_titulo_movimento)
		{
			#region [Declarações]
			String strSql;
			SqlCommand cmCommand;
			#endregion

			cmCommand = BD.criaSqlCommand();

			strSql = "SELECT qtde_vezes_editado_manual " +
					 "FROM t_SERASA_TITULO_MOVIMENTO " +
					 "WHERE id = @id ";

			cmCommand.CommandText = strSql;
			cmCommand.Parameters.AddWithValue("@id", id_serasa_titulo_movimento);

			object ret = cmCommand.ExecuteScalar();
			if (ret == DBNull.Value)
			{
				return 0;
			}

			return Convert.ToInt32(ret);
		}
		#endregion

		#region [ cancelaEnvio ]
		public static bool cancelaEnvio(int id)
		{
			#region [Declarações]
			String strOperacao = "UPDATE t_SERASA_TITULO_MOVIMENTO";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmCancelaEnvio.Parameters["@dt_envio_serasa_cancelado"].Value = DateTime.Now.Date;
			cmCancelaEnvio.Parameters["@dt_hr_envio_serasa_cancelado"].Value = DateTime.Now;
			cmCancelaEnvio.Parameters["@usuario_envio_serasa_cancelado"].Value = Global.Usuario.usuario;
			cmCancelaEnvio.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmCancelaEnvio);
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
