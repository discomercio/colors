#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.IO;
#endregion

namespace FinanceiroService
{
	static class GeralDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmUpdateTabelaControleNsu;
		private static SqlCommand cmUpdateTabelaControleNsuComLetraSeq;
		private static SqlCommand cmUpdateTabelaControleNsuAcquireXLock;
		private static SqlCommand cmUpdateUploadFileAsDeleted;
		private static SqlCommand cmInsereLog;
		private static SqlCommand cmInsereFinSvcLog;
		#endregion

		#region [ inicializaConstrutorEstatico ]
		public static void inicializaConstrutorEstatico()
		{
			// NOP
			// 1) The static constructor for a class executes before any instance of the class is created.
			// 2) The static constructor for a class executes before any of the static members for the class are referenced.
			// 3) The static constructor for a class executes after the static field initializers (if any) for the class.
			// 4) The static constructor for a class executes at most one time during a single program instantiation
			// 5) A static constructor does not take access modifiers or have parameters.
			// 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
			// 7) A static constructor cannot be called directly.
			// 8) The user has no control on when the static constructor is executed in the program.
			// 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
		}
		#endregion

		#region [ Construtor estático ]
		static GeralDAO()
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

			#region [ cmUpdateTabelaControleNsu ]
			// Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes através
			// do controle que impede o campo 'dt_ult_atualizacao' de receber uma data menor do que aquela que ele já possui
			strSql = "UPDATE t_CONTROLE SET " +
						"nsu = @nsu_novo, " +
						"dt_ult_atualizacao = CASE WHEN dt_ult_atualizacao > " + Global.sqlMontaGetdateSomenteData() + " THEN dt_ult_atualizacao ELSE " + Global.sqlMontaGetdateSomenteData() + " END" +
					" WHERE" +
						" (id_nsu = @id_nsu)" +
						" AND (nsu = @nsu_atual)";
			cmUpdateTabelaControleNsu = BD.criaSqlCommand();
			cmUpdateTabelaControleNsu.CommandText = strSql;
			cmUpdateTabelaControleNsu.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
			cmUpdateTabelaControleNsu.Parameters.Add("@nsu_novo", SqlDbType.VarChar, 12);
			cmUpdateTabelaControleNsu.Parameters.Add("@nsu_atual", SqlDbType.VarChar, 12);
			cmUpdateTabelaControleNsu.Prepare();
			#endregion

			#region [ cmUpdateTabelaControleNsuComLetraSeq ]
			// Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes através
			// do controle que impede o campo 'dt_ult_atualizacao' de receber uma data menor do que aquela que ele já possui
			strSql = "UPDATE t_CONTROLE SET " +
						"nsu = @nsu_novo, " +
						"ano_letra_seq = @ano_letra_seq_novo, " +
						"dt_ult_atualizacao = CASE WHEN dt_ult_atualizacao > " + Global.sqlMontaGetdateSomenteData() + " THEN dt_ult_atualizacao ELSE " + Global.sqlMontaGetdateSomenteData() + " END" +
					" WHERE" +
						" (id_nsu = @id_nsu)" +
						" AND (nsu = @nsu_atual)";
			cmUpdateTabelaControleNsuComLetraSeq = BD.criaSqlCommand();
			cmUpdateTabelaControleNsuComLetraSeq.CommandText = strSql;
			cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
			cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@nsu_novo", SqlDbType.VarChar, 12);
			cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@nsu_atual", SqlDbType.VarChar, 12);
			cmUpdateTabelaControleNsuComLetraSeq.Parameters.Add("@ano_letra_seq_novo", SqlDbType.VarChar, 1);
			cmUpdateTabelaControleNsuComLetraSeq.Prepare();
			#endregion

			#region [ cmUpdateTabelaControleNsuAcquireXLock ]
			// Usado para bloquear o registro de forma a evitar o acesso concorrente (realiza o flip em um campo bit apenas p/ adquirir o lock exclusivo)
			strSql = "UPDATE t_CONTROLE SET" +
						" dummy = ~dummy" +
					" WHERE"+
						" (id_nsu = @id_nsu)";
			cmUpdateTabelaControleNsuAcquireXLock = BD.criaSqlCommand();
			cmUpdateTabelaControleNsuAcquireXLock.CommandText = strSql;
			cmUpdateTabelaControleNsuAcquireXLock.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
			cmUpdateTabelaControleNsuAcquireXLock.Prepare();
			#endregion

			#region [ cmUpdateUploadFileAsDeleted ]
			strSql = "UPDATE t_UPLOAD_FILE SET" +
						" st_file_deleted = 1," +
						" usuario_file_deleted = @usuario_file_deleted," +
						" dt_file_deleted = " + Global.sqlMontaGetdateSomenteData() + "," +
						" dt_hr_file_deleted = getdate()," +
						" file_content = NULL," +
						" file_content_text = NULL" +
					" WHERE" +
						" (id = @id)";
			cmUpdateUploadFileAsDeleted = BD.criaSqlCommand();
			cmUpdateUploadFileAsDeleted.CommandText = strSql;
			cmUpdateUploadFileAsDeleted.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateUploadFileAsDeleted.Parameters.Add("@usuario_file_deleted", SqlDbType.VarChar, 20);
			cmUpdateUploadFileAsDeleted.Prepare();
			#endregion

			#region [ cmInsereLog ]
			strSql = "INSERT INTO t_LOG (" +
						"data, " +
						"usuario, " +
						"loja, " +
						"pedido, " +
						"id_cliente, " +
						"operacao, " +
						"complemento" +
					") VALUES (" +
						"getdate(), " +
						"@usuario, " +
						"@loja, " +
						"@pedido, " +
						"@id_cliente, " +
						"@operacao, " +
						"@complemento" +
					")";
			cmInsereLog = BD.criaSqlCommand();
			cmInsereLog.CommandText = strSql;
			cmInsereLog.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsereLog.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmInsereLog.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsereLog.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmInsereLog.Parameters.Add("@operacao", SqlDbType.VarChar, 20);
			cmInsereLog.Parameters.Add("@complemento", SqlDbType.Text, 2147483647);
			cmInsereLog.Prepare();
			#endregion

			#region [ cmInsereFinSvcLog ]
			strSql = "INSERT INTO t_FINSVC_LOG (" +
						"id, " +
						"operacao, " +
						"tabela, " +
						"descricao, " +
						"complemento_1, " +
						"complemento_2, " +
						"complemento_3, " +
						"complemento_4," +
						"complemento_5," +
						"complemento_6" +
					") VALUES (" +
						"@id, " +
						"@operacao, " +
						"@tabela, " +
						"@descricao, " +
						"@complemento_1, " +
						"@complemento_2, " +
						"@complemento_3, " +
						"@complemento_4," +
						"@complemento_5," +
						"@complemento_6" +
					")";
			cmInsereFinSvcLog = BD.criaSqlCommand();
			cmInsereFinSvcLog.CommandText = strSql;
			cmInsereFinSvcLog.Parameters.Add("@id", SqlDbType.Int);
			cmInsereFinSvcLog.Parameters.Add("@operacao", SqlDbType.VarChar, 160);
			cmInsereFinSvcLog.Parameters.Add("@tabela", SqlDbType.VarChar, 160);
			cmInsereFinSvcLog.Parameters.Add("@descricao", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Parameters.Add("@complemento_1", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Parameters.Add("@complemento_2", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Parameters.Add("@complemento_3", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Parameters.Add("@complemento_4", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Parameters.Add("@complemento_5", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Parameters.Add("@complemento_6", SqlDbType.VarChar, -1); // varchar(max)
			cmInsereFinSvcLog.Prepare();
			#endregion
		}
		#endregion

		#region [ getRegistroTabelaParametro ]
		public static RegistroTabelaParametro getRegistroTabelaParametro(String nomeParametro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.getRegistroTabelaParametro()";
			string strSql = "";
			string msg_erro_aux;
			RegistroTabelaParametro parametro = new RegistroTabelaParametro();
			DataTable dtbResultado = new DataTable();
			DataRow row;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			#endregion

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT " +
							"*" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				parametro.id = BD.readToString(row["id"]);
				parametro.campo_inteiro = BD.readToInt(row["campo_inteiro"]);
				parametro.campo_monetario = BD.readToDecimal(row["campo_monetario"]);
				parametro.campo_real = BD.readToSingle(row["campo_real"]);
				parametro.campo_data = BD.readToDateTime(row["campo_data"]);
				parametro.campo_texto = BD.readToString(row["campo_texto"]);
				parametro.campo_2_texto = BD.readToString(row["campo_2_texto"]);
				parametro.dt_hr_ult_atualizacao = BD.readToDateTime(row["dt_hr_ult_atualizacao"]);
				parametro.usuario_ult_atualizacao = BD.readToString(row["usuario_ult_atualizacao"]);
				parametro.obs = BD.readToString(row["obs"]);

				return parametro;
			}
			catch (Exception ex)
			{
				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "t_PARAMETRO.id=" + nomeParametro;
				svcLog.complemento_2 = Global.serializaObjectToXml(parametro);
				svcLog.complemento_3 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getCampoDataTabelaParametro ]
		public static DateTime getCampoDataTabelaParametro(String nomeParametro)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			DateTime dtHrResultado = DateTime.MinValue;
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT " +
						Global.sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador("campo_data") +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = objResultado.ToString();
				if ((strResultado != null) && (strResultado.Length > 0)) dtHrResultado = Global.converteYyyyMmDdHhMmSsParaDateTime(strResultado);
			}
			return dtHrResultado;
		}
		#endregion

		#region [ getCampoInteiroTabelaParametro ]
		public static int getCampoInteiroTabelaParametro(String nomeParametro)
		{
			return getCampoInteiroTabelaParametro(nomeParametro, 0);
		}

		public static int getCampoInteiroTabelaParametro(String nomeParametro, int valorDefault)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			int intResultado;
			SqlCommand cmCommand;
			#endregion

			intResultado = valorDefault;

			strSql = "SELECT " +
						"campo_inteiro" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				intResultado = BD.readToInt(objResultado);
			}
			return intResultado;
		}
		#endregion

		#region [ getCampoTextoTabelaParametro ]
		public static String getCampoTextoTabelaParametro(String nomeParametro)
		{
			return getCampoTextoTabelaParametro(nomeParametro, "");
		}

		public static String getCampoTextoTabelaParametro(String nomeParametro, String valorDefault)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			SqlCommand cmCommand;
			#endregion

			strResultado = valorDefault;

			strSql = "SELECT " +
						"campo_texto" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = BD.readToString(objResultado);
			}
			return strResultado;
		}
		#endregion

		#region [ nfeEmitenteCarregaFromDataRow ]
		private static NfeEmitente nfeEmitenteCarregaFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			NfeEmitente emitente = new NfeEmitente();
			#endregion

			emitente.id = BD.readToInt(rowDados["id"]);
			emitente.id_boleto_cedente = BD.readToInt(rowDados["id_boleto_cedente"]);
			emitente.braspag_id_boleto_cedente = BD.readToInt(rowDados["braspag_id_boleto_cedente"]);
			emitente.st_ativo = BD.readToByte(rowDados["st_ativo"]);
			emitente.apelido = BD.readToString(rowDados["apelido"]);
			emitente.cnpj = BD.readToString(rowDados["cnpj"]);
			emitente.razao_social = BD.readToString(rowDados["razao_social"]);
			emitente.endereco = BD.readToString(rowDados["endereco"]);
			emitente.endereco_numero = BD.readToString(rowDados["endereco_numero"]);
			emitente.endereco_complemento = BD.readToString(rowDados["endereco_complemento"]);
			emitente.bairro = BD.readToString(rowDados["bairro"]);
			emitente.cidade = BD.readToString(rowDados["cidade"]);
			emitente.uf = BD.readToString(rowDados["uf"]);
			emitente.cep = BD.readToString(rowDados["cep"]);
			emitente.NFe_st_emitente_padrao = BD.readToByte(rowDados["NFe_st_emitente_padrao"]);
			emitente.NFe_T1_servidor_BD = BD.readToString(rowDados["NFe_T1_servidor_BD"]);
			emitente.NFe_T1_nome_BD = BD.readToString(rowDados["NFe_T1_nome_BD"]);
			emitente.NFe_T1_usuario_BD = BD.readToString(rowDados["NFe_T1_usuario_BD"]);
			emitente.NFe_T1_senha_BD = Criptografia.Descriptografa(BD.readToString(rowDados["NFe_T1_senha_BD"]));
			emitente.st_habilitado_ctrl_estoque = BD.readToByte(rowDados["st_habilitado_ctrl_estoque"]);
			emitente.ordem = BD.readToInt(rowDados["ordem"]);
			emitente.texto_fixo_especifico = BD.readToString(rowDados["texto_fixo_especifico"]);
			emitente.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			emitente.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			emitente.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			emitente.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			emitente.dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["dt_hr_ult_atualizacao"]);
			emitente.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);

			return emitente;
		}
		#endregion

		#region [ getListaNfeEmitente ]
		public static List<NfeEmitente> getListaNfeEmitente(Global.eOpcaoFiltroStAtivo filtroStAtivo)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.getListaNfeEmitente()";
			string strSql = "";
			string strWhere = "";
			string msg_erro_aux;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			NfeEmitente emitente;
			List<NfeEmitente> listaNfeEmitente = new List<NfeEmitente>();
			#endregion

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Monta SQL ]
				if (filtroStAtivo == Global.eOpcaoFiltroStAtivo.SELECIONAR_SOMENTE_ATIVOS)
				{
					strWhere = " (st_ativo = 1)";
				}
				else if (filtroStAtivo == Global.eOpcaoFiltroStAtivo.SELECIONAR_SOMENTE_INATIVOS)
				{
					strWhere = " (st_ativo = 0)";
				}

				if (strWhere.Length > 0) strWhere = " WHERE " + strWhere;

				strSql = "SELECT " +
							"*" +
						" FROM t_NFe_EMITENTE" +
						strWhere +
						" ORDER BY" +
							" ordem";
				#endregion

				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					emitente = nfeEmitenteCarregaFromDataRow(row);
					listaNfeEmitente.Add(emitente);
				}

				return listaNfeEmitente;
			}
			catch (Exception ex)
			{
				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "filtroStAtivo=" + filtroStAtivo.ToString();
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ setCampoDataTabelaParametro ]
		public static bool setCampoDataTabelaParametro(String nomeParametro, DateTime dtHrValorParametro)
		{
			#region [ Declarações ]
			String strSql;
			String strValorParametro;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Prepara o valor do parâmetro p/ o SQL ]
				if (dtHrValorParametro == DateTime.MinValue)
				{
					strValorParametro = "NULL";
				}
				else
				{
					strValorParametro = Global.sqlMontaDateTimeParaSqlDateTime(dtHrValorParametro);
				}
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_data = " + strValorParametro +
								", dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_data, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								strValorParametro + ", " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_data no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ setCampoInteiroTabelaParametro ]
		public static bool setCampoInteiroTabelaParametro(String nomeParametro, int valorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_inteiro = " + valorParametro.ToString() +
								", dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_inteiro, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								valorParametro.ToString() + ", " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_inteiro no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ setCampoTextoTabelaParametro ]
		public static bool setCampoTextoTabelaParametro(String nomeParametro, String valorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_texto = @campo_texto," +
								" dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_texto, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								"@campo_texto, " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				cmCommand.Parameters.Add("@campo_texto", SqlDbType.VarChar, 1024);
				cmCommand.Parameters["@campo_texto"].Value = valorParametro;
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_texto no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ setCampo2TextoTabelaParametro ]
		public static bool setCampo2TextoTabelaParametro(String nomeParametro, String valorParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_2_texto = @campo_2_texto," +
								" dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"campo_2_texto, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								"@campo_2_texto, " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				cmCommand.Parameters.Add("@campo_2_texto", SqlDbType.VarChar, 1024);
				cmCommand.Parameters["@campo_2_texto"].Value = valorParametro;
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao gravar em t_PARAMETRO.campo_2_texto no registro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ resetRegistroTabelaParametro ]
		public static bool resetRegistroTabelaParametro(String nomeParametro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			int intQtdeCount;
			int intQtdeUpdated;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();

				#region [ Registro existe? ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM t_PARAMETRO" +
						" WHERE" +
							" (id = '" + nomeParametro + "')";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				#endregion

				#region [ Grava o novo valor do parâmetro ]
				if (intQtdeCount == 1)
				{
					strSql = "UPDATE" +
								" t_PARAMETRO" +
							" SET" +
								" campo_inteiro = 0," +
								" campo_monetario = 0," +
								" campo_real = 0," +
								" campo_data = NULL," +
								" campo_texto = NULL," +
								" campo_2_texto = NULL," +
								" dt_hr_ult_atualizacao = getdate()," +
								" usuario_ult_atualizacao = NULL" +
							" WHERE" +
								" (id = '" + nomeParametro + "')";
				}
				else
				{
					strSql = "INSERT INTO t_PARAMETRO (" +
								"id, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + nomeParametro + "', " +
								"getdate()" +
							")";
				}
				cmCommand.CommandText = strSql;
				intQtdeUpdated = BD.executaNonQuery(ref cmCommand);
				#endregion

				#region [ Sucesso ou falha? ]
				if (intQtdeUpdated == 1)
					return true;
				else
					return false;
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao tentar realizar o reset dos campos do parâmetro '" + nomeParametro + "'\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ getDataHoraUltCargaArqRetornoBoleto ]
		public static DateTime getDataHoraUltCargaArqRetornoBoleto()
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			DateTime dtHrResultado = DateTime.MinValue;
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT " +
						Global.sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador("Max(dt_hr_processamento)") +
					" FROM t_FIN_BOLETO_ARQ_RETORNO" +
					" WHERE" +
						" (st_processamento = " + Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.SUCESSO + ")";
			cmCommand = BD.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = objResultado.ToString();
				if ((strResultado != null) && (strResultado.Length > 0)) dtHrResultado = Global.converteYyyyMmDdHhMmSsParaDateTime(strResultado);
			}
			return dtHrResultado;
		}
		#endregion

		#region [ atualizaTabelaControleNsu ]
		public static bool atualizaTabelaControleNsu(String id_nsu,
													String nsu_novo,
													String nsu_atual,
													out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.atualizaTabelaControleNsu()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_nsu == null)
				{
					strMsgErro = "Não foi informado o identificador do NSU!!";
					return false;
				}

				if (id_nsu.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do NSU!!";
					return false;
				}

				if (nsu_novo == null)
				{
					strMsgErro = "Não foi informado o valor do novo NSU!!";
					return false;
				}

				if (nsu_novo.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o valor do novo NSU!!";
					return false;
				}

				if (nsu_atual == null)
				{
					strMsgErro = "Não foi informado o valor do NSU atual!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateTabelaControleNsu.Parameters["@id_nsu"].Value = id_nsu;
				cmUpdateTabelaControleNsu.Parameters["@nsu_novo"].Value = nsu_novo;
				cmUpdateTabelaControleNsu.Parameters["@nsu_atual"].Value = nsu_atual;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateTabelaControleNsu);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaTabelaControleNsuComLetraSeq ]
		public static bool atualizaTabelaControleNsuComLetraSeq(String id_nsu,
													String nsu_novo,
													String nsu_atual,
													String ano_letra_seq_novo,
													out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.atualizaTabelaControleNsuComLetraSeq()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_nsu == null)
				{
					strMsgErro = "Não foi informado o identificador do NSU!!";
					return false;
				}

				if (id_nsu.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do NSU!!";
					return false;
				}

				if (nsu_novo == null)
				{
					strMsgErro = "Não foi informado o valor do novo NSU!!";
					return false;
				}

				if (nsu_novo.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o valor do novo NSU!!";
					return false;
				}

				if (nsu_atual == null)
				{
					strMsgErro = "Não foi informado o valor do NSU atual!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateTabelaControleNsuComLetraSeq.Parameters["@id_nsu"].Value = id_nsu;
				cmUpdateTabelaControleNsuComLetraSeq.Parameters["@nsu_novo"].Value = nsu_novo;
				cmUpdateTabelaControleNsuComLetraSeq.Parameters["@nsu_atual"].Value = nsu_atual;
				cmUpdateTabelaControleNsuComLetraSeq.Parameters["@ano_letra_seq_novo"].Value = ano_letra_seq_novo;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateTabelaControleNsuComLetraSeq);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ geraNsuUsandoTabelaControle ]
		public static bool geraNsuUsandoTabelaControle(String id_nsu, out String nsu_novo, out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.geraNsuUsandoTabelaControle()";
			const int MAX_TENTATIVAS = 20;
			int intQtdeTentativas = 0;
			bool blnRetorno;
			#endregion

			nsu_novo = "";
			strMsgErro = "";

			try
			{
				while (true)
				{
					intQtdeTentativas++;

					blnRetorno = executaGeraNsuUsandoTabelaControle(id_nsu, out nsu_novo, out strMsgErro);
					if (blnRetorno) return true;

					if (intQtdeTentativas > MAX_TENTATIVAS)
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar gerar o NSU após " + MAX_TENTATIVAS.ToString() + "!!" + strMsgErro;
						return false;
					}

					Thread.Sleep(100);
				}
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaGeraNsuUsandoTabelaControle ]
		private static bool executaGeraNsuUsandoTabelaControle(String id_nsu, out String nsu_novo, out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.executaGeraNsuUsandoTabelaControle()";
			int n_nsu;
			String strNsuNovo;
			String strNsuAtual = "";
			String strLetraSeqNovo;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			nsu_novo = "";
			strMsgErro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Consistências ]
				if (id_nsu == null)
				{
					strMsgErro = "Não foi informado o NSU a ser gerado!!";
					return false;
				}

				if (id_nsu.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi especificado o NSU a ser gerado!!";
					return false;
				}
				#endregion

				strMsgErro = "";
				n_nsu = -1;

				#region [ Bloqueia registro p/ evitar acesso concorrente ]
				if (Global.Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO)
				{
					cmUpdateTabelaControleNsuAcquireXLock.Parameters["@id_nsu"].Value = id_nsu;
					BD.executaNonQuery(ref cmUpdateTabelaControleNsuAcquireXLock);
				}
				#endregion

				strSql = "SELECT * FROM t_CONTROLE WHERE (id_nsu = '" + id_nsu + "')";

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Ainda não existe registro de controle para gerar este NSU? ]
				if (dtbConsulta.Rows.Count == 0)
				{
					strMsgErro = "Não existe registro na tabela de controle para poder gerar este NSU!!";
					return false;
				}
				#endregion

				rowConsulta = dtbConsulta.Rows[0];
				if (!Convert.IsDBNull(rowConsulta["nsu"]))
				{
					strNsuAtual = BD.readToString(rowConsulta["nsu"]);
					if (strNsuAtual.Trim().Length > 0)
					{
						n_nsu = (int)Global.converteInteiro(strNsuAtual);
						if (BD.readToInt(rowConsulta["seq_anual"]) != 0)
						{
							// Caso o relógio do servidor seja alterado p/ datas futuras e passadas, evita que o campo 'ano_letra_seq' seja incrementado várias vezes
							if (DateTime.Today.Year > BD.readToDateTime(rowConsulta["dt_ult_atualizacao"]).Year)
							{
								// Se mudou o ano, reinicia a contagem do NSU
								strNsuNovo = "".PadLeft(Global.Cte.Etc.TAM_MAX_NSU, '0');
								n_nsu = 0;
								if (BD.readToString(rowConsulta["ano_letra_seq"]).Trim().Length > 0)
								{
									strLetraSeqNovo = BD.readToString(rowConsulta["ano_letra_seq"]);
									strLetraSeqNovo = Texto.chr((short)(Texto.asc(strLetraSeqNovo[0]) + BD.readToInt(rowConsulta["ano_letra_step"]))).ToString();
									if (!atualizaTabelaControleNsuComLetraSeq(id_nsu, strNsuNovo, strNsuAtual, strLetraSeqNovo, out strMsgErro))
									{
										if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
										strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
										return false;
									}
								}
								else
								{
									if (!atualizaTabelaControleNsu(id_nsu, strNsuNovo, strNsuAtual, out strMsgErro))
									{
										if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
										strMsgErro = "Falha ao tentar atualizar o registro da tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
										return false;
									}
								}
							}
						}
					}
				}

				if (n_nsu < 0)
				{
					strMsgErro = "O NSU gerado é inválido!!";
					return false;
				}

				n_nsu++;
				strNsuNovo = n_nsu.ToString().PadLeft(Global.Cte.Etc.TAM_MAX_NSU, '0');
				if (!atualizaTabelaControleNsu(id_nsu, strNsuNovo, strNsuAtual, out strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar a tabela de controle (id_nsu=" + id_nsu + ")!!" + strMsgErro;
					return false;
				}

				nsu_novo = strNsuNovo;
				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaManutencaoBdLogAntigo ]
		public static bool executaManutencaoBdLogAntigo(out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.executaManutencaoBdLogAntigo()";
			String strMsgErroAux;
			String strMsg;
			String strSqlCondicao;
			DateTime dtHrInicio = DateTime.Now;
			#endregion

			strMsgErro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				executaLimpezaTabela("t_LOG", "data", Global.Cte.ManutencaoLogBd.Corte.T_LOG__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_FIN_LOG", "data", Global.Cte.ManutencaoLogBd.Corte.T_FIN_LOG__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_ESTOQUE_LOG", "data", Global.Cte.ManutencaoLogBd.Corte.T_ESTOQUE_LOG__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_ESTOQUE_SALDO_DIARIO", "data", Global.Cte.ManutencaoLogBd.Corte.T_ESTOQUE_SALDO_DIARIO__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_ESTOQUE_VENDA_SALDO_DIARIO", "data", Global.Cte.ManutencaoLogBd.Corte.T_ESTOQUE_VENDA_SALDO_DIARIO__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_SESSAO_ABANDONADA", "SessaoAbandonadaDtHrInicio", Global.Cte.ManutencaoLogBd.Corte.T_SESSAO_ABANDONADA__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_SESSAO_HISTORICO", "DtHrInicio", Global.Cte.ManutencaoLogBd.Corte.T_SESSAO_HISTORICO__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_SESSAO_RESTAURADA", "DataHora", Global.Cte.ManutencaoLogBd.Corte.T_SESSAO_RESTAURADA__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_FINSVC_LOG", "data", Global.Cte.ManutencaoLogBd.Corte.T_FINSVC_LOG__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_EMAILSNDSVC_LOG", "dt_cadastro", Global.Cte.ManutencaoLogBd.Corte.T_EMAILSNDSVC_LOG__CORTE_EM_DIAS, out strMsgErroAux);
				executaLimpezaTabela("t_EMAILSNDSVC_LOG_ERRO", "dt_cadastro", Global.Cte.ManutencaoLogBd.Corte.T_EMAILSNDSVC_LOG_ERRO__CORTE_EM_DIAS, out strMsgErroAux);
                executaLimpezaTabela("t_CTRL_RELATORIO_USUARIO_X_PEDIDO", "data", Global.Cte.ManutencaoLogBd.Corte.T_CTRL_RELATORIO_USUARIO_X_PEDIDO__CORTE_EM_DIAS, out strMsgErroAux);

                strSqlCondicao = "(st_usado_cadastramento_pedido_erp = 0)";
				executaLimpezaTabelaCondicional("t_MAGENTO_API_PEDIDO_XML", "dt_cadastro", Global.Cte.ManutencaoLogBd.Corte.T_MAGENTO_API_PEDIDO_XML__INFO_DESCARTADA__CORTE_EM_DIAS, strSqlCondicao, out strMsgErroAux);

				strSqlCondicao = "(st_usado_cadastramento_pedido_erp = 1)";
				executaLimpezaTabelaCondicional("t_MAGENTO_API_PEDIDO_XML", "dt_cadastro", Global.Cte.ManutencaoLogBd.Corte.T_MAGENTO_API_PEDIDO_XML__INFO_UTILIZADA__CORTE_EM_DIAS, strSqlCondicao, out strMsgErroAux);

				strSqlCondicao = "(status = 0) AND (proc_comissao_status = 0) AND (proc_fluxo_caixa_status = 0)";
				executaLimpezaTabelaCondicional("t_COMISSAO_INDICADOR_NFSe_N1", "dt_cadastro", Global.Cte.ManutencaoLogBd.Corte.T_COMISSAO_INDICADOR_NFSe_N1__INFO_DESCARTADA__CORTE_EM_DIAS, strSqlCondicao, out strMsgErroAux);

				strMsg = "Rotina " + NOME_DESTA_ROTINA + " concluída com sucesso (duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio) + ")";
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaLimpezaTabela ]
		public static bool executaLimpezaTabela(String strNomeTabela, String strNomeCampoDataParaCorte, int intQtdeDiasAPreservar, out String strMsgErro)
		{
			#region [ Declarações ]
			const int MIN_REGISTROS_APOS_CORTE = 5000;
			String strSql;
			String strDataCorte;
			String strDataCorteSqlFormat;
			String strMsgLog = "";
			String strMsgErroAux;
			DateTime dtCorte;
			int intQtdeCount;
			int intQtdeRegistrosApagada;
			SqlCommand cmCommand;
			DateTime dtHrInicio = DateTime.Now;
			#endregion

			strMsgErro = "";

			#region [ Consistências ]

			#region [ Nome da tabela ]
			if (strNomeTabela == null)
			{
				strMsgErro = "Nome da tabela não foi informada!!";
				return false;
			}
			if (strNomeTabela.Trim().Length == 0)
			{
				strMsgErro = "Nome da tabela não foi fornecido!!";
				return false;
			}
			#endregion

			#region [ Nome do campo data ]
			if (strNomeCampoDataParaCorte == null)
			{
				strMsgErro = "Não foi informado o nome do campo na tabela que deve ser usado no corte!!";
				return false;
			}
			if (strNomeCampoDataParaCorte.Trim().Length == 0)
			{
				strMsgErro = "Não foi fornecido o nome do campo na tabela que deve ser usado no corte!!";
				return false;
			}
			#endregion

			#region [ Quantidade de dias a preservar ]
			if (intQtdeDiasAPreservar <= 0)
			{
				strMsgErro = "O período de corte dos dados antigos não foi informado!!";
				return false;
			}
			#endregion

			#endregion

			try
			{
				dtCorte = DateTime.Today.AddDays(-Math.Abs(intQtdeDiasAPreservar));
				strDataCorteSqlFormat = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtCorte);
				strDataCorte = dtCorte.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador);

				cmCommand = BD.criaSqlCommand();

				#region [ Executa o corte somente se for restar mais registros que o limite mínimo ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM " + strNomeTabela +
						" WHERE" +
							" (" + strNomeCampoDataParaCorte + " >= " + strDataCorteSqlFormat + ")";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				if (intQtdeCount <= MIN_REGISTROS_APOS_CORTE)
				{
					strMsgLog = "Eliminação de dados antigos da tabela " + strNomeTabela + " não foi feita porque restariam apenas " +
								Global.formataInteiro(intQtdeCount) + " registros posteriores à data de corte " + strDataCorte +
								" (limite mínimo: " + Global.formataInteiro(MIN_REGISTROS_APOS_CORTE) + ")";
					return true;
				}
				#endregion

				#region [ Apaga os dados antigos ]
				strSql = "DELETE" +
						 " FROM " + strNomeTabela +
						 " WHERE" +
							" (" + strNomeCampoDataParaCorte + " < " + strDataCorteSqlFormat + ")";
				cmCommand.CommandText = strSql;
				intQtdeRegistrosApagada = BD.executaNonQuery(ref cmCommand);
				strMsgLog = "Limpeza da tabela " + strNomeTabela + ": " + intQtdeRegistrosApagada.ToString() + " registros apagados, data de corte: " + strDataCorte + ", duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio);
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				strMsgLog = "Limpeza da tabela " + strNomeTabela + ": falha na operação!! Mensagem de erro: " + ex.ToString();
				return false;
			}
			finally
			{
				#region [ Grava log ]
				Global.gravaLogAtividade(strMsgLog);
				gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_EXECUTA_LIMPEZA_TABELA, strMsgLog, out strMsgErroAux);
				#endregion
			}
		}
		#endregion

		#region [ executaLimpezaTabelaCondicional ]
		public static bool executaLimpezaTabelaCondicional(String strNomeTabela, String strNomeCampoDataParaCorte, int intQtdeDiasAPreservar, String strSqlCondicao, out String strMsgErro)
		{
			#region [ Declarações ]
			const int MIN_REGISTROS_APOS_CORTE = 500;
			String strSql;
			String strDataCorte;
			String strDataCorteSqlFormat;
			String strMsgLog = "";
			String strMsgErroAux;
			DateTime dtCorte;
			int intQtdeCount;
			int intQtdeRegistrosApagada;
			SqlCommand cmCommand;
			DateTime dtHrInicio = DateTime.Now;
			#endregion

			strMsgErro = "";

			#region [ Consistências ]

			#region [ Nome da tabela ]
			if (strNomeTabela == null)
			{
				strMsgErro = "Nome da tabela não foi informada!!";
				return false;
			}
			if (strNomeTabela.Trim().Length == 0)
			{
				strMsgErro = "Nome da tabela não foi fornecido!!";
				return false;
			}
			#endregion

			#region [ Nome do campo data ]
			if (strNomeCampoDataParaCorte == null)
			{
				strMsgErro = "Não foi informado o nome do campo na tabela que deve ser usado no corte!!";
				return false;
			}
			if (strNomeCampoDataParaCorte.Trim().Length == 0)
			{
				strMsgErro = "Não foi fornecido o nome do campo na tabela que deve ser usado no corte!!";
				return false;
			}
			#endregion

			#region [ Quantidade de dias a preservar ]
			if (intQtdeDiasAPreservar <= 0)
			{
				strMsgErro = "O período de corte dos dados antigos não foi informado!!";
				return false;
			}
			#endregion

			#region [ SQL da condição ]
			if ((strSqlCondicao ?? "").Trim().Length == 0)
			{
				strMsgErro = "Não foi informado o SQL da condição a ser usada!";
				return false;
			}
			#endregion

			#endregion

			try
			{
				dtCorte = DateTime.Today.AddDays(-Math.Abs(intQtdeDiasAPreservar));
				strDataCorteSqlFormat = Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtCorte);
				strDataCorte = dtCorte.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador);

				cmCommand = BD.criaSqlCommand();

				#region [ Executa o corte somente se for restar mais registros que o limite mínimo ]
				strSql = "SELECT" +
							" Count(*)" +
						" FROM " + strNomeTabela +
						" WHERE" +
							" (" + strNomeCampoDataParaCorte + " >= " + strDataCorteSqlFormat + ")" +
							" AND (" + strSqlCondicao + ")";
				cmCommand.CommandText = strSql;
				intQtdeCount = (int)cmCommand.ExecuteScalar();
				if (intQtdeCount <= MIN_REGISTROS_APOS_CORTE)
				{
					strMsgLog = "Eliminação de dados antigos da tabela " + strNomeTabela + " não foi feita porque restariam apenas " +
								Global.formataInteiro(intQtdeCount) + " registros posteriores à data de corte " + strDataCorte +
								" (limite mínimo: " + Global.formataInteiro(MIN_REGISTROS_APOS_CORTE) + ")";
					return true;
				}
				#endregion

				#region [ Apaga os dados antigos ]
				strSql = "DELETE" +
						 " FROM " + strNomeTabela +
						 " WHERE" +
							" (" + strNomeCampoDataParaCorte + " < " + strDataCorteSqlFormat + ")" +
							" AND (" + strSqlCondicao + ")";
				cmCommand.CommandText = strSql;
				intQtdeRegistrosApagada = BD.executaNonQuery(ref cmCommand);
				strMsgLog = "Limpeza da tabela " + strNomeTabela + ": " + intQtdeRegistrosApagada.ToString() + " registros apagados, data de corte: " + strDataCorte + ", condição: '" + strSqlCondicao + "', duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio);
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				strMsgLog = "Limpeza da tabela " + strNomeTabela + ": falha na operação!! Mensagem de erro: " + ex.ToString();
				return false;
			}
			finally
			{
				#region [ Grava log ]
				Global.gravaLogAtividade(strMsgLog);
				gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_EXECUTA_LIMPEZA_TABELA, strMsgLog, out strMsgErroAux);
				#endregion
			}
		}
		#endregion

		#region [ executaLimpezaSessionToken ]
		public static bool executaLimpezaSessionToken(out string strMsgInformativa, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.executaLimpezaSessionToken()";
			string strMsg;
			string strSql;
			string strMsgLog = "";
			string strMsgErroAux;
			int intQtdeRegistros;
			DateTime dtRef;
			DateTime dtHrInicio = DateTime.Now;
			SqlCommand cmCommand;
			#endregion

			strMsgInformativa = "";
			strMsgErro = "";
			try
			{
				#region [ Prepara objetos de acesso ao banco de dados ]
				cmCommand = BD.criaSqlCommand();
				#endregion

				dtRef = DateTime.Now.AddHours(-6);

				#region [ Limpa campos SessionTokenModuloCentral ]
				strSql = "UPDATE" +
							" t_USUARIO" +
						" SET" +
							" SessionTokenModuloCentral = NULL" +
						" WHERE" +
							" (SessionTokenModuloCentral IS NOT NULL)" +
							" AND (DtHrSessionTokenModuloCentral IS NOT NULL)" +
							" AND (DtHrSessionTokenModuloCentral < " + Global.sqlMontaDateTimeParaSqlDateTime(dtRef) + ")";

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				cmCommand.CommandText = strSql;

				intQtdeRegistros = BD.executaNonQuery(ref cmCommand);
				if (strMsgLog.Length > 0) strMsgLog += ", ";
				strMsgLog += "SessionTokenModuloCentral: " + intQtdeRegistros.ToString() + " registros";
				#endregion

				#region [ Limpa campos SessionTokenModuloLoja ]
				strSql = "UPDATE" +
							" t_USUARIO" +
						" SET" +
							" SessionTokenModuloLoja = NULL" +
						" WHERE" +
							" (SessionTokenModuloLoja IS NOT NULL)" +
							" AND (DtHrSessionTokenModuloLoja IS NOT NULL)" +
							" AND (DtHrSessionTokenModuloLoja < " + Global.sqlMontaDateTimeParaSqlDateTime(dtRef) + ")";

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				cmCommand.CommandText = strSql;

				intQtdeRegistros = BD.executaNonQuery(ref cmCommand);
				if (strMsgLog.Length > 0) strMsgLog += ", ";
				strMsgLog += "SessionTokenModuloLoja: " + intQtdeRegistros.ToString() + " registros";
				#endregion

				if ((strMsgInformativa.Length == 0) && (strMsgLog.Length > 0)) strMsgInformativa = strMsgLog;

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
			finally
			{
				#region [ Grava log ]
				if (strMsgLog.Length > 0) strMsgLog = "Limpeza de session token: " + strMsgLog + ", duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio);
				Global.gravaLogAtividade(strMsgLog);
				gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_LIMPEZA_SESSION_TOKEN, strMsgLog, out strMsgErroAux);
				#endregion
			}
		}
		#endregion

		#region [ excluiArquivoUploadFile ]
		private static bool excluiArquivoUploadFile(int id_upload_file, string fullFilenameToDelete, string tipoArquivo, ref StringBuilder sbLogFalha, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.excluiArquivoUploadFile()";
			string strMsg;
			string strMsgErroAux;
			bool blnFileDeleted = false;
			#endregion

			strMsgErro = "";
			if (sbLogFalha == null) sbLogFalha = new StringBuilder("");
			try
			{
				if (fullFilenameToDelete.Length == 0)
				{
					strMsg = NOME_DESTA_ROTINA + " - Falha na exclusão do arquivo (nome completo não informado): ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete;
					Global.gravaLogAtividade(strMsg);

					strMsg = "Falha na exclusão de arquivo (tipo '" + tipoArquivo + "'): não há a informação do nome completo do arquivo (ID=" + id_upload_file.ToString() + ")";
					sbLogFalha.AppendLine(strMsg);
				}
				else
				{
					if (!File.Exists(fullFilenameToDelete))
					{
						strMsg = NOME_DESTA_ROTINA + " - Falha na exclusão do arquivo (arquivo não localizado): ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete;
						Global.gravaLogAtividade(strMsg);

						strMsg = "O arquivo (tipo '" + tipoArquivo + "') a ser excluído não existe e será marcado como já excluído (ID=" + id_upload_file.ToString() + "): " + fullFilenameToDelete;
						sbLogFalha.AppendLine(strMsg);

						#region [ Marca o arquivo como já excluído para não continuar aparecendo na consulta ]
						if (!updateUploadFileAsDeleted(id_upload_file, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out strMsgErroAux))
						{
							strMsg = NOME_DESTA_ROTINA + " - Falha ao marcar o arquivo como já excluído: ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete;
							if (strMsgErroAux.Length > 0) strMsg += "\n" + strMsgErroAux;
							Global.gravaLogAtividade(strMsg);

							strMsg = "Falha ao atualizar o banco de dados para marcar o arquivo como já excluído (ID=" + id_upload_file.ToString() + ")";
							if (strMsgErroAux.Length > 0) strMsg += ": " + strMsgErroAux;
							sbLogFalha.AppendLine(strMsg);
						}
						#endregion
					}
					else
					{
						try
						{
							File.Delete(fullFilenameToDelete);

							// Excluiu o arquivo?
							if (File.Exists(fullFilenameToDelete))
							{
								strMsg = NOME_DESTA_ROTINA + " - Falha na exclusão do arquivo (arquivo permanece gravado após comando de exclusão): ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete;
								Global.gravaLogAtividade(strMsg);

								strMsg = "Falha na exclusão do arquivo (ID=" + id_upload_file.ToString() + "): arquivo permanece gravado após comando de exclusão [" + fullFilenameToDelete + "]";
								sbLogFalha.AppendLine(strMsg);
							}
							else
							{
								blnFileDeleted = true;

								strMsg = NOME_DESTA_ROTINA + " - Arquivo excluído: ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete;
								Global.gravaLogAtividade(strMsg);

								#region [ Marca o arquivo como já excluído ]
								if (!updateUploadFileAsDeleted(id_upload_file, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out strMsgErroAux))
								{
									strMsg = NOME_DESTA_ROTINA + " - Falha ao atualizar o banco de dados para marcar o arquivo como já excluído: ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete;
									if (strMsgErroAux.Length > 0) strMsg += "\n" + strMsgErroAux;
									Global.gravaLogAtividade(strMsg);

									strMsg = "Falha ao atualizar o banco de dados para marcar o arquivo como já excluído (ID=" + id_upload_file.ToString() + ")";
									if (strMsgErroAux.Length > 0) strMsg += ": " + strMsgErroAux;
									sbLogFalha.AppendLine(strMsg);
								}
								#endregion
							}
						}
						catch (Exception ex)
						{
							strMsg = NOME_DESTA_ROTINA + " - Exception na exclusão do arquivo: ID=" + id_upload_file.ToString() + ", tipo=" + tipoArquivo + ", arquivo=" + fullFilenameToDelete + "\n" + ex.ToString();
							Global.gravaLogAtividade(strMsg);

							strMsg = "Falha ao tentar excluir o arquivo (ID=" + id_upload_file.ToString() + "): " + fullFilenameToDelete + "\n" + ex.Message;
							sbLogFalha.AppendLine(strMsg);
						}
					}
				}

				return blnFileDeleted;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaManutencaoArquivosUploadFile ]
		public static bool executaManutencaoArquivosUploadFile(out string strMsgInformativa, out string strLogFalha, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.executaManutencaoArquivosUploadFile()";
			string strSql;
			string strMsg;
			string strMsgLog = "";
			string strMsgErroAux;
			string sFullFileNameToDelete;
			string sInfoLogFileName;
			StringBuilder sbTempFileDeleted = new StringBuilder("");
			StringBuilder sbTempFileDeletionFailed = new StringBuilder("");
			StringBuilder sbPendingConfirmationFileDeleted = new StringBuilder("");
			StringBuilder sbPendingConfirmationFileDeletionFailed = new StringBuilder("");
			StringBuilder sbStoredFileDeleted = new StringBuilder("");
			StringBuilder sbStoredFileDeletionFailed = new StringBuilder("");
			StringBuilder sbLogFalha = new StringBuilder("");
			int qtdeTempFileDeleted = 0;
			int qtdeTempFileDeletionFailed = 0;
			int qtdePendingConfirmationFileDeleted = 0;
			int qtdePendingConfirmationFileDeletionFailed = 0;
			int qtdeStoredFileDeleted = 0;
			int qtdeStoredFileDeletionFailed = 0;
			int id_upload_file;
			DateTime dtHrInicio = DateTime.Now;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			strMsgInformativa = "";
			strLogFalha = "";
			strMsgErro = "";
			try
			{
				#region [ Prepara objetos de acesso ao banco de dados ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Limpeza de arquivos temporários ]

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"*" +
						" FROM t_UPLOAD_FILE" +
						" WHERE" +
							" (st_temporary_file = 1)" +
							" AND (st_file_deleted = 0)" +
						" ORDER BY" +
							" id";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Apaga cada um dos arquivos retornados na consulta ]
				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					id_upload_file = BD.readToInt(rowConsulta["id"]);
					sFullFileNameToDelete = BD.readToString(rowConsulta["stored_full_file_name"]);
					sInfoLogFileName = "ID=" + id_upload_file.ToString() + ", stored_file_name=" + BD.readToString(rowConsulta["stored_file_name"]) + ", original_file_name=" + BD.readToString(rowConsulta["original_file_name"]);
					if (excluiArquivoUploadFile(id_upload_file, sFullFileNameToDelete, "Temporary File", ref sbLogFalha, out strMsgErroAux))
					{
						qtdeTempFileDeleted++;
						sbTempFileDeleted.AppendLine(sInfoLogFileName);
					}
					else
					{
						qtdeTempFileDeletionFailed++;
						sbTempFileDeletionFailed.AppendLine(sInfoLogFileName);
					}
				}
				#endregion

				#endregion

				#region [ Limpeza dos arquivos pendentes de confirmação ]

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"*" +
						" FROM t_UPLOAD_FILE" +
						" WHERE" +
							" (st_confirmation_required = 1)" +
							" AND (st_confirmation_ok = 0)" +
							" AND (st_file_deleted = 0)" +
						" ORDER BY" +
							" id";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				dtbConsulta.Reset();
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Apaga cada um dos arquivos retornados na consulta ]
				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					id_upload_file = BD.readToInt(rowConsulta["id"]);
					sFullFileNameToDelete = BD.readToString(rowConsulta["stored_full_file_name"]);
					sInfoLogFileName = "ID=" + id_upload_file.ToString() + ", stored_file_name=" + BD.readToString(rowConsulta["stored_file_name"]) + ", original_file_name=" + BD.readToString(rowConsulta["original_file_name"]);
					if (excluiArquivoUploadFile(id_upload_file, sFullFileNameToDelete, "Pending Confirmation", ref sbLogFalha, out strMsgErroAux))
					{
						qtdePendingConfirmationFileDeleted++;
						sbPendingConfirmationFileDeleted.AppendLine(sInfoLogFileName);
					}
					else
					{
						qtdePendingConfirmationFileDeletionFailed++;
						sbPendingConfirmationFileDeletionFailed.AppendLine(sInfoLogFileName);
					}
				}
				#endregion

				#endregion

				#region [ Limpeza dos arquivos armazenados com exclusão agendada/solicitada ]

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"*" +
						" FROM t_UPLOAD_FILE" +
						" WHERE" +
							" (st_delete_file = 1)" +
							" AND ( (dt_delete_file_scheduled_date IS NULL) OR (dt_delete_file_scheduled_date < getdate()) )" +
							" AND (st_file_deleted = 0)" +
						" ORDER BY" +
							" id";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				dtbConsulta.Reset();
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Apaga cada um dos arquivos retornados na consulta ]
				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					id_upload_file = BD.readToInt(rowConsulta["id"]);
					sFullFileNameToDelete = BD.readToString(rowConsulta["stored_full_file_name"]);
					sInfoLogFileName = "ID=" + id_upload_file.ToString() + ", stored_file_name=" + BD.readToString(rowConsulta["stored_file_name"]) + ", original_file_name=" + BD.readToString(rowConsulta["original_file_name"]);
					if (excluiArquivoUploadFile(id_upload_file, sFullFileNameToDelete, "Stored File", ref sbLogFalha, out strMsgErroAux))
					{
						qtdeStoredFileDeleted++;
						sbStoredFileDeleted.AppendLine(sInfoLogFileName);
					}
					else
					{
						qtdeStoredFileDeletionFailed++;
						sbStoredFileDeletionFailed.AppendLine(sInfoLogFileName);
					}
				}
				#endregion

				#endregion

				#region [ Mensagem informativa de retorno para a rotina chamadora ]
				strMsgInformativa = "'Temporary files' excluídos: " + qtdeTempFileDeleted.ToString() + ", " +
									"'Temporary files' com falha na exclusão: " + qtdeTempFileDeletionFailed.ToString() + ", " +
									"'Pending Confirmation files' excluídos: " + qtdePendingConfirmationFileDeleted.ToString() + ", " +
									"'Pending Confirmation files' com falha na exclusão: " + qtdePendingConfirmationFileDeletionFailed.ToString() + ", " +
									"'Stored files' excluídos: " + qtdeStoredFileDeleted.ToString() + ", " +
									"'Stored files' com falha na exclusão: " + qtdeStoredFileDeletionFailed.ToString();
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
			finally
			{
				#region [ Grava log ]
				if (strMsgLog.Length > 0) strMsgLog += "; ";
				strMsgLog += "duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio);
				strMsgLog = "Manutenção de arquivos salvos no servidor através da WebAPI (UploadFile): " + strMsgLog;
				if (strMsgInformativa.Length > 0) strMsgLog += "; " + strMsgInformativa;
				Global.gravaLogAtividade(strMsgLog);

				strMsgLog += "\n'Temporary files' excluídos:\n" + (sbTempFileDeleted.ToString().Length == 0 ? "(nenhum)" : sbTempFileDeleted.ToString()) +
						"\n'Temporary files' com falha na exclusão:\n" + (sbTempFileDeletionFailed.ToString().Length == 0 ? "(nenhum)" : sbTempFileDeletionFailed.ToString()) +
						"\n'Pending Confirmation files' excluídos:\n" + (sbPendingConfirmationFileDeleted.ToString().Length == 0 ? "(nenhum)" : sbPendingConfirmationFileDeleted.ToString()) +
						"\n'Pending Confirmation files' com falha na exclusão:\n" + (sbPendingConfirmationFileDeletionFailed.ToString().Length == 0 ? "(nenhum)" : sbPendingConfirmationFileDeletionFailed.ToString()) +
						"\n'Stored files' excluídos:\n" + (sbStoredFileDeleted.ToString().Length == 0 ? "(nenhum)" : sbStoredFileDeleted.ToString()) +
						"\n'Stored files' com falha na exclusão:\n" + (sbStoredFileDeletionFailed.ToString().Length == 0 ? "(nenhum)" : sbStoredFileDeletionFailed.ToString());
				gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_MANUTENCAO_ARQUIVOS_UPLOAD_FILE, strMsgLog, out strMsgErroAux);
				#endregion

				#region [ Se ocorreu alguma falha imprevista, retorna para a rotina chamadora ]
				strLogFalha = sbLogFalha.ToString();
				#endregion
			}
		}
		#endregion

		#region [ gravaLog ]
		/// <summary>
		/// Grava dados no log do BD (tabela t_LOG)
		/// </summary>
		/// <param name="operacao">Identificação da operação realizada</param>
		/// <param name="complemento">Descrição da operação e/ou dados complementares</param>
		/// <param name="strMsgErro">Mensagem de erro retornada em caso de falha na gravação</param>
		/// <returns>
		/// true: sucesso na gravação do log
		/// false: falha na gravação do log
		/// </returns>
		public static bool gravaLog(String operacao, String complemento, out String strMsgErro)
		{
			return gravaLog(Global.Cte.LogBd.Usuario.ID_USUARIO_LOG, "", "", "", operacao, complemento, out strMsgErro);
		}

		public static bool gravaLog(String operacao, String pedido, String complemento, out String strMsgErro)
		{
			return gravaLog(Global.Cte.LogBd.Usuario.ID_USUARIO_LOG, "", pedido, "", operacao, complemento, out strMsgErro);
		}

		/// <summary>
		/// Grava dados no log do BD (tabela t_LOG)
		/// </summary>
		/// <param name="usuario">Identificação do usuário que realizou a operação</param>
		/// <param name="loja">Nº da loja, se houver</param>
		/// <param name="pedido">Nº do pedido, se houver</param>
		/// <param name="id_cliente">Identificação do cliente, se houver</param>
		/// <param name="operacao">Identificação da operação realizada</param>
		/// <param name="complemento">Descrição da operação e/ou dados complementares</param>
		/// <param name="strMsgErro">Mensagem de erro retornada em caso de falha na gravação</param>
		/// <returns>
		/// true: sucesso na gravação do log
		/// false: falha na gravação do log
		/// </returns>
		public static bool gravaLog(String usuario,
									String loja,
									String pedido,
									String id_cliente,
									String operacao,
									String complemento,
									out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.gravaLog()";
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";

					#region [ Preenche o valor dos parâmetros ]
					cmInsereLog.Parameters["@usuario"].Value = ((usuario == null) ? "" : usuario);
					cmInsereLog.Parameters["@loja"].Value = ((loja == null) ? "" : loja);
					cmInsereLog.Parameters["@pedido"].Value = ((pedido == null) ? "" : pedido);
					cmInsereLog.Parameters["@id_cliente"].Value = ((id_cliente == null) ? "" : id_cliente);
					cmInsereLog.Parameters["@operacao"].Value = ((operacao == null) ? "" : operacao);
					cmInsereLog.Parameters["@complemento"].Value = complemento;
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsereLog.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsereLog);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
					if (intRetorno == 1)
					{
						blnSucesso = true;
					}
					else
					{
						Thread.Sleep(100);
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao gravar o log no BD após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha: Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ gravaFinSvcLog ]
		public static bool gravaFinSvcLog(FinSvcLog log, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.gravaFinSvcLog()";
			string strMsgLogAtividade;
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			bool blnGerouNsu;
			int idFinSvcLog = 0;
			#endregion

			try
			{
				if (log.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_FINSVC_LOG, out idFinSvcLog, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro da tabela " + Global.Cte.FIN.NSU.T_FINSVC_LOG + "\n" + msg_erro;
						return false;
					}
					log.id = idFinSvcLog;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idFinSvcLog = log.id;
				}

				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;

					msg_erro = "";

					#region [ Preenche o valor dos parâmetros ]
					cmInsereFinSvcLog.Parameters["@id"].Value = idFinSvcLog;
					cmInsereFinSvcLog.Parameters["@operacao"].Value = ((log.operacao == null) ? "" : log.operacao);
					cmInsereFinSvcLog.Parameters["@tabela"].Value = ((log.tabela == null) ? "" : log.tabela);
					cmInsereFinSvcLog.Parameters["@descricao"].Value = ((log.descricao == null) ? "" : log.descricao);
					cmInsereFinSvcLog.Parameters["@complemento_1"].Value = ((log.complemento_1 == null) ? "" : log.complemento_1);
					cmInsereFinSvcLog.Parameters["@complemento_2"].Value = ((log.complemento_2 == null) ? "" : log.complemento_2);
					cmInsereFinSvcLog.Parameters["@complemento_3"].Value = ((log.complemento_3 == null) ? "" : log.complemento_3);
					cmInsereFinSvcLog.Parameters["@complemento_4"].Value = ((log.complemento_4 == null) ? "" : log.complemento_4);
					cmInsereFinSvcLog.Parameters["@complemento_5"].Value = ((log.complemento_5 == null) ? "" : log.complemento_5);
					cmInsereFinSvcLog.Parameters["@complemento_6"].Value = ((log.complemento_6 == null) ? "" : log.complemento_6);
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsereFinSvcLog);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception!!\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
					if (intRetorno == 1)
					{
						blnSucesso = true;
					}
					else
					{
						Thread.Sleep(100);
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					msg_erro = "Falha ao gravar o log no BD após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				msg_erro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception\n" + ex.ToString());
				return false;
			}
			finally
			{
				strMsgLogAtividade = "Gravação de log em t_FINSVC_LOG (id=" + log.id.ToString() + "): " + log.operacao + " - " + log.descricao;
				Global.gravaLogAtividade(strMsgLogAtividade);
			}
		}
		#endregion

		#region [ isPedidoECommerce ]
		public static bool isPedidoECommerce(string numeroPedidoEC, out string numeroPedidoERP)
		{
			#region [ Declarações ]
			string strSql;
			string strPedido;
			string strStEntrega;
			string strNumeroPedidoErpStEntregaAny = "";
			string strNumeroPedidoErpStEntregaValido = "";
			SqlCommand cmCommand;
			SqlDataReader dr;
			#endregion

			numeroPedidoERP = "";

			if ((numeroPedidoEC ?? "").Length == 0) return false;

			#region [ Tenta localizar o pedido no banco de dados ]
			cmCommand = BD.criaSqlCommand();
			strSql = "SELECT" +
						" pedido," +
						" st_entrega" +
					" FROM t_PEDIDO" +
					" WHERE" +
						" (pedido_bs_x_ac = '" + numeroPedidoEC + "')" +
					" ORDER BY" +
						" data_hora DESC";
			cmCommand.CommandText = strSql;
			dr = cmCommand.ExecuteReader();
			try
			{
				if (!dr.HasRows) return false;

				while (dr.Read())
				{
					strPedido = "";
					strStEntrega = "";
					if (!Convert.IsDBNull(dr["pedido"])) strPedido = dr["pedido"].ToString();
					if (!Convert.IsDBNull(dr["st_entrega"])) strStEntrega = dr["st_entrega"].ToString();

					// Armazena o número do pedido mais recente apenas
					if (strNumeroPedidoErpStEntregaAny.Length == 0) strNumeroPedidoErpStEntregaAny = strPedido;

					if (strNumeroPedidoErpStEntregaValido.Length == 0)
					{
						if (strStEntrega.Length > 0)
						{
							if (!strStEntrega.Equals(Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
							{
								strNumeroPedidoErpStEntregaValido = strPedido;
							}
						}
					}
				}  // while

				// Encontrou pedido c/ status de entrega válido?
				if (strNumeroPedidoErpStEntregaValido.Length > 0)
				{
					numeroPedidoERP = strNumeroPedidoErpStEntregaValido;
					return true;
				}

				// Se não encontrou pedido c/ status válido, retorna pedido c/ qualquer status, se houver
				if (strNumeroPedidoErpStEntregaAny.Length > 0)
				{
					numeroPedidoERP = strNumeroPedidoErpStEntregaAny;
					return true;
				}

				return false;
			}
			finally
			{
				dr.Close();
			}
			#endregion
		}
		#endregion

		#region [ updateUploadFileAsDeleted ]
		public static bool updateUploadFileAsDeleted(int id, string usuario, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GeralDAO.updateUploadFileAsDeleted()";
			string msg_erro_aux;
			int intRetorno;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche parâmetros ]
				cmUpdateUploadFileAsDeleted.Parameters["@id"].Value = id;
				cmUpdateUploadFileAsDeleted.Parameters["@usuario_file_deleted"].Value = (usuario ?? "");
				#endregion

				#region [ Tenta realizar o update ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateUploadFileAsDeleted);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = "ID = " + id.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "ID = " + id.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion
	}
}
