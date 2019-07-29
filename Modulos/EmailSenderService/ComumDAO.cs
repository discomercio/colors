#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace EmailSenderService
{
	static class ComumDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmUpdateTabelaControleNsu;
		private static SqlCommand cmInsereLog;
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
		static ComumDAO()
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
			strSql = "UPDATE t_CONTROLE SET " +
						"nsu = @nsu_novo, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() +
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
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			int intResultado = 0;
			SqlCommand cmCommand;
			#endregion

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

		#region [ atualizaTabelaControleNsu ]
		public static bool atualizaTabelaControleNsu(String id_nsu,
													String nsu_novo,
													String nsu_atual,
													out String strMsgErro)
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "atualizaTabelaControleNsu()";
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
					strMsgErro = strNomeDestaRotina + " - Tentativa resultou em exception!!\n" + ex.ToString();
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
				strMsgErro = strNomeDestaRotina + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ geraNsuUsandoTabelaControle ]
		public static bool geraNsuUsandoTabelaControle(String id_nsu, out String nsu_novo, out String strMsgErro)
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "geraNsuUsandoTabelaControle()";
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
				strMsgErro = strNomeDestaRotina + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaGeraNsuUsandoTabelaControle ]
		private static bool executaGeraNsuUsandoTabelaControle(String id_nsu, out String nsu_novo, out String strMsgErro)
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "executaGeraNsuUsandoTabelaControle()";
			int n_nsu;
			String strNsuNovo;
			String strNsuAtual = "";
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
							if (DateTime.Today.Year != BD.readToDateTime(rowConsulta["dt_ult_atualizacao"]).Year)
							{
								// Se mudou o ano, reinicia a contagem do NSU
								strNsuNovo = "".PadLeft(Global.Cte.Etc.TAM_MAX_NSU, '0');
								n_nsu = 0;
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
				strMsgErro = strNomeDestaRotina + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
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
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			String strOperacao = "Gravação do log";
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
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
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
				Global.gravaLogAtividade(strOperacao + " - Falha: Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion
	}
}
