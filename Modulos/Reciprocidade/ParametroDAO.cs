#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading;
#endregion

namespace Reciprocidade
{
	class ParametroDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsert;
		private static SqlCommand cmUpdate;
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
		static ParametroDAO()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmInsert ]
			strSql = "INSERT INTO t_PARAMETRO (" +
						"id, " +
						"campo_inteiro, " +
						"campo_monetario, " +
						"campo_real, " +
						"campo_data, " +
						"campo_texto, " +
						"dt_hr_ult_atualizacao, " +
						"usuario_ult_atualizacao" +
					") VALUES (" +
						"@id, " +
						"@campo_inteiro, " +
						"@campo_monetario, " +
						"@campo_real, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@campo_data") + ", " +
						"@campo_texto, " +
						"getdate(), " +
						"@usuario_ult_atualizacao" +
					")";
			cmInsert = BD.criaSqlCommand();
			cmInsert.CommandText = strSql;
			cmInsert.Parameters.Add("@id", SqlDbType.VarChar, 80);
			cmInsert.Parameters.Add("@campo_inteiro", SqlDbType.Int);
			cmInsert.Parameters.Add("@campo_monetario", SqlDbType.Money);
			cmInsert.Parameters.Add("@campo_real", SqlDbType.Real);
			cmInsert.Parameters.Add("@campo_data", SqlDbType.VarChar, 19);
			cmInsert.Parameters.Add("@campo_texto", SqlDbType.VarChar, 1024);
			cmInsert.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmInsert.Prepare();
			#endregion

			#region [ cmUpdate ]
			strSql = "UPDATE t_PARAMETRO SET " +
						"campo_inteiro = @campo_inteiro, " +
						"campo_monetario = @campo_monetario, " +
						"campo_real = @campo_real, " +
						"campo_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@campo_data") + ", " +
						"campo_texto = @campo_texto, " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao" +
					" WHERE (id = @id)";
			cmUpdate = BD.criaSqlCommand();
			cmUpdate.CommandText = strSql;
			cmUpdate.Parameters.Add("@id", SqlDbType.VarChar, 80);
			cmUpdate.Parameters.Add("@campo_inteiro", SqlDbType.Int);
			cmUpdate.Parameters.Add("@campo_monetario", SqlDbType.Money);
			cmUpdate.Parameters.Add("@campo_real", SqlDbType.Real);
			cmUpdate.Parameters.Add("@campo_data", SqlDbType.VarChar, 19);
			cmUpdate.Parameters.Add("@campo_texto", SqlDbType.VarChar, 1024);
			cmUpdate.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmUpdate.Prepare();
			#endregion
		}
		#endregion

		#region [ insere ]
		public static bool insere(String usuario,
								Parametro parametro,
								out String strDescricaoLog,
								out String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			String strOperacao = "Inserção de registro na tabela de parâmetros";
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			strDescricaoLog = "";
			strMsgErro = "";

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;
					strMsgErro = "";

					#region [ Tenta gravar o registro ]

					#region [ Preenche os parâmetros ]
					cmInsert.Parameters["@id"].Value = parametro.id;
					cmInsert.Parameters["@campo_inteiro"].Value = parametro.campo_inteiro;
					cmInsert.Parameters["@campo_monetario"].Value = parametro.campo_monetario;
					cmInsert.Parameters["@campo_real"].Value = parametro.campo_real;
					cmInsert.Parameters["@campo_data"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(parametro.campo_data);
					cmInsert.Parameters["@campo_texto"].Value = parametro.campo_texto;
					cmInsert.Parameters["@usuario_ult_atualizacao"].Value = usuario;
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsert.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsert);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
					if (intRetorno == 1)
					{
						strDescricaoLog = sbLog.ToString();
						Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
						blnSucesso = true;
					}
					else
					{
						Thread.Sleep(100);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar gravar na tabela de parâmetros após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ altera ]
		public static bool altera(String usuario,
								  Parametro parametro,
								  out String strDescricaoLog,
								  out String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do parâmetro '" + parametro.id + "'";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strDescricaoLog = "";
			strMsgErro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdate.Parameters["@id"].Value = parametro.id;
				cmUpdate.Parameters["@campo_inteiro"].Value = parametro.campo_inteiro;
				cmUpdate.Parameters["@campo_monetario"].Value = parametro.campo_monetario;
				cmUpdate.Parameters["@campo_real"].Value = parametro.campo_real;
				cmUpdate.Parameters["@campo_data"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(parametro.campo_data);
				cmUpdate.Parameters["@campo_texto"].Value = parametro.campo_texto;
				cmUpdate.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdate);
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

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar o registro do parâmetro '" + parametro.id + "'!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ salva ]
		public static bool salva(String usuario,
								Parametro parametro,
								out String strDescricaoLog,
								out String strMsgErro)
		{
			#region [ Declarações ]
			Parametro parametroAux;
			String strMsgException;
			String strOperacao = "Salva os dados do registro do parâmetro '" + parametro.id + "'";
			#endregion

			strDescricaoLog = "";
			strMsgErro = "";

			try
			{
				#region [ Verifica se o parâmetro já está cadastrado ou não ]
				try
				{
					parametroAux = getParametro(parametro.id);
				}
				catch (Exception ex)
				{
					strMsgException = ex.Message;
					parametroAux = null;
				}
				#endregion

				if (parametroAux == null)
				{
					#region [ Se parâmetro ainda não está cadastrado, cadastra agora ]
					insere(usuario, parametro, out strDescricaoLog, out strMsgErro);
					#endregion
				}
				else
				{
					#region [ Parâmetro já está cadastrado, então atualiza ]
					altera(usuario, parametro, out strDescricaoLog, out strMsgErro);
					#endregion
				}

				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ getParametro ]
		public static Parametro getParametro(String id)
		{
			#region [ Declarações ]
			String strSql;
			Parametro parametro = new Parametro();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id.Trim().Length == 0) throw new Exception("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa consulta ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + id + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) throw new Exception("Registro id=" + id + " não localizado na tabela de parâmetros!!");
			#endregion

			#region [ Transfere dados para objeto a ser retornado ]
			rowResultado = dtbResultado.Rows[0];

			parametro.id = rowResultado["id"].ToString();
			parametro.campo_inteiro = (int)rowResultado["campo_inteiro"];
			parametro.campo_monetario = (decimal)rowResultado["campo_monetario"];
			parametro.campo_real = (Single)rowResultado["campo_real"];
			parametro.campo_data = !Convert.IsDBNull(rowResultado["campo_data"]) ? (DateTime)rowResultado["campo_data"] : DateTime.MinValue;
			parametro.campo_texto = !Convert.IsDBNull(rowResultado["campo_texto"]) ? rowResultado["campo_texto"].ToString() : "";
			parametro.dt_hr_ult_atualizacao = !Convert.IsDBNull(rowResultado["dt_hr_ult_atualizacao"]) ? (DateTime)rowResultado["dt_hr_ult_atualizacao"] : DateTime.MinValue;
			parametro.usuario_ult_atualizacao = !Convert.IsDBNull(rowResultado["usuario_ult_atualizacao"]) ? rowResultado["usuario_ult_atualizacao"].ToString() : "";
			#endregion

			return parametro;
		}
		#endregion
	}
}
