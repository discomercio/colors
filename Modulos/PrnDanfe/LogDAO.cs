#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
#endregion

namespace PrnDANFE
{
	class LogDAO
	{
		#region [ Atributos ]
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
		static LogDAO()
		{
			String strSql;

			#region [ cmInsereFinLog ]
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
			cmInsereLog.Parameters.Add("@complemento", SqlDbType.Text, -1);
			cmInsereLog.Prepare();
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ insere ]
		/// <summary>
		/// Grava novo registro no log
		/// </summary>
		/// <param name="usuario">
		/// Identificação do usuário que realizou a operação
		/// </param>
		/// <param name="log">
		/// Objeto que representa um registro do log contendo os dados para gravar
		/// </param>
		/// <param name="strMsgErro">
		/// Retorna a mensagem de erro no caso de ocorrer exception
		/// </param>
		/// <returns>
		/// true: gravação efetuada com sucesso
		/// false: falha na gravação
		/// </returns>
		public static bool insere(
								   String usuario,
								   Log log,
								   ref String strMsgErro
								)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			String strOperacao = "Gravação de log";
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
					cmInsereLog.Parameters["@usuario"].Value = log.usuario;
					cmInsereLog.Parameters["@loja"].Value = (log.loja == null ? "" : log.loja);
					cmInsereLog.Parameters["@pedido"].Value = (log.pedido == null ? "" : log.pedido);
					cmInsereLog.Parameters["@id_cliente"].Value = (log.id_cliente == null ? "" : log.id_cliente);
					cmInsereLog.Parameters["@operacao"].Value = (log.operacao == null ? "" : log.operacao);
					cmInsereLog.Parameters["@complemento"].Value = (log.complemento == null ? "" : log.complemento);
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsereLog.Parameters)
					{
						if (!item.ParameterName.Equals("@complemento"))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
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
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados da operação = " + log.complemento + "; Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
					if (intRetorno == 1)
					{
						Global.gravaLogAtividade(strOperacao + " - Sucesso: Dados da operação = " + log.complemento + "; Dados do registro do log = " + sbLog.ToString());
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
					strMsgErro = "Falha ao gravar no banco de dados o log após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: Dados da operação = " + log.complemento + "; Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#endregion
	}
}
