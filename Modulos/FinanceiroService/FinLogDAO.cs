using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Threading;

namespace FinanceiroService
{
	class FinLogDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsereFinLog;
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
		static FinLogDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmInsereFinLog ]
			strSql = "INSERT INTO t_FIN_LOG (" +
						"data, " +
						"data_hora, " +
						"usuario, " +
						"operacao, " +
						"natureza, " +
						"tipo_cadastro, " +
						"fin_modulo, " +
						"cod_tabela_origem, " +
						"id_registro_origem, " +
						"id_conta_corrente, " +
						"id_plano_contas_empresa, " +
						"id_plano_contas_grupo, " +
						"id_plano_contas_conta, " +
						"id_boleto_cedente, " +
						"id_cliente, " +
						"cnpj_cpf, " +
						"descricao" +
					") VALUES (" +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario, " +
						"@operacao, " +
						"@natureza, " +
						"@tipo_cadastro, " +
						"@fin_modulo, " +
						"@cod_tabela_origem, " +
						"@id_registro_origem, " +
						"@id_conta_corrente, " +
						"@id_plano_contas_empresa, " +
						"@id_plano_contas_grupo, " +
						"@id_plano_contas_conta, " +
						"@id_boleto_cedente, " +
						"@id_cliente, " +
						"@cnpj_cpf, " +
						"@descricao" +
					")";
			cmInsereFinLog = BD.criaSqlCommand();
			cmInsereFinLog.CommandText = strSql;
			cmInsereFinLog.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsereFinLog.Parameters.Add("@operacao", SqlDbType.VarChar, 12);
			cmInsereFinLog.Parameters.Add("@natureza", SqlDbType.Char, 1);
			cmInsereFinLog.Parameters.Add("@tipo_cadastro", SqlDbType.Char, 1);
			cmInsereFinLog.Parameters.Add("@fin_modulo", SqlDbType.Char, 3);
			cmInsereFinLog.Parameters.Add("@cod_tabela_origem", SqlDbType.TinyInt);
			cmInsereFinLog.Parameters.Add("@id_registro_origem", SqlDbType.Int);
			cmInsereFinLog.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
			cmInsereFinLog.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
			cmInsereFinLog.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
			cmInsereFinLog.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
			cmInsereFinLog.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmInsereFinLog.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmInsereFinLog.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
			cmInsereFinLog.Parameters.Add("@descricao", SqlDbType.VarChar, 8000);
			cmInsereFinLog.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		/// <summary>
		/// Grava novo registro no log (módulo financeiro)
		/// </summary>
		/// <param name="usuario">
		/// Identificação do usuário que realizou a operação
		/// </param>
		/// <param name="finLog">
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
								   FinLog finLog,
								   ref String strMsgErro
								)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			String strOperacao = "Gravação de log (financeiro)";
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
					cmInsereFinLog.Parameters["@usuario"].Value = finLog.usuario;
					cmInsereFinLog.Parameters["@operacao"].Value = finLog.operacao;
					cmInsereFinLog.Parameters["@natureza"].Value = finLog.natureza;
					cmInsereFinLog.Parameters["@tipo_cadastro"].Value = finLog.tipo_cadastro;
					cmInsereFinLog.Parameters["@fin_modulo"].Value = ((finLog.fin_modulo == null) ? "" : finLog.fin_modulo);
					cmInsereFinLog.Parameters["@cod_tabela_origem"].Value = finLog.cod_tabela_origem;
					cmInsereFinLog.Parameters["@id_registro_origem"].Value = finLog.id_registro_origem;
					cmInsereFinLog.Parameters["@id_conta_corrente"].Value = finLog.id_conta_corrente;
					cmInsereFinLog.Parameters["@id_plano_contas_empresa"].Value = finLog.id_plano_contas_empresa;
					cmInsereFinLog.Parameters["@id_plano_contas_grupo"].Value = finLog.id_plano_contas_grupo;
					cmInsereFinLog.Parameters["@id_plano_contas_conta"].Value = finLog.id_plano_contas_conta;
					cmInsereFinLog.Parameters["@id_boleto_cedente"].Value = finLog.id_boleto_cedente;
					cmInsereFinLog.Parameters["@id_cliente"].Value = ((finLog.id_cliente == null) ? "" : finLog.id_cliente);
					cmInsereFinLog.Parameters["@cnpj_cpf"].Value = Global.digitos(finLog.cnpj_cpf);
					// Certifica-se de que não vai exceder o tamanho do campo
					if (finLog.descricao == null)
						cmInsereFinLog.Parameters["@descricao"].Value = "";
					else if (finLog.descricao.Length > Global.Cte.FIN.TamanhoCampo.FIN_LOG_DESCRICAO)
						cmInsereFinLog.Parameters["@descricao"].Value = finLog.descricao.Substring(0, Global.Cte.FIN.TamanhoCampo.FIN_LOG_DESCRICAO);
					else
						cmInsereFinLog.Parameters["@descricao"].Value = finLog.descricao;
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsereFinLog.Parameters)
					{
						if (!item.ParameterName.Equals("@descricao"))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsereFinLog);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados da operação = " + finLog.descricao + "; Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
					if (intRetorno == 1)
					{
						Global.gravaLogAtividade(strOperacao + " - Sucesso: Dados da operação = " + finLog.descricao + "; Dados do registro do log = " + sbLog.ToString());
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
					strMsgErro = "Falha ao gravar no banco de dados o log (financeiro) após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: Dados da operação = " + finLog.descricao + "; Dados do registro do log = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#endregion
	}
}
