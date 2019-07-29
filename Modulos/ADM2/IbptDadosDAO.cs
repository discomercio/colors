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

namespace ADM2
{
	public class IbptDadosDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
		private SqlCommand cmInsertTabelaTemporaria;
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

		#region [ Construtor ]
		public IbptDadosDAO(ref BancoDados bd)
		{
			_bd = bd;
			inicializaObjetos();
		}
		#endregion

		#region [ inicializaObjetos ]
		public void inicializaObjetos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmInsert ]
			strSql = "INSERT INTO " + Global.Cte.ADM2.Tabelas.PREFIXO_TABELA_TEMPORARIA + Global.Cte.ADM2.Tabelas.T_IBPT + " (" +
						"codigo, " +
						"ex, " +
						"tabela, " +
						"descricao, " +
						"aliqNac, " +
						"aliqImp, " +
						"percAliqNac, " +
						"percAliqImp" +
					") VALUES (" +
						"@codigo, " +
						"@ex, " +
						"@tabela, " +
						"@descricao, " +
						"@aliqNac, " +
						"@aliqImp, " +
						"@percAliqNac, " +
						"@percAliqImp" +
					")";
			cmInsertTabelaTemporaria = _bd.criaSqlCommand();
			cmInsertTabelaTemporaria.CommandText = strSql;
			cmInsertTabelaTemporaria.Parameters.Add("@codigo", SqlDbType.VarChar, 9);
			cmInsertTabelaTemporaria.Parameters.Add("@ex", SqlDbType.VarChar, 3);
			cmInsertTabelaTemporaria.Parameters.Add("@tabela", SqlDbType.VarChar, 1);
			cmInsertTabelaTemporaria.Parameters.Add("@descricao", SqlDbType.VarChar, 500);
			cmInsertTabelaTemporaria.Parameters.Add("@aliqNac", SqlDbType.VarChar, 8);
			cmInsertTabelaTemporaria.Parameters.Add("@aliqImp", SqlDbType.VarChar, 8);
			cmInsertTabelaTemporaria.Parameters.Add("@percAliqNac", SqlDbType.Real);
			cmInsertTabelaTemporaria.Parameters.Add("@percAliqImp", SqlDbType.Real);
			cmInsertTabelaTemporaria.Prepare();
			#endregion
		}
		#endregion

		#region [ insereTabelaTemporaria ]
		public bool insereTabelaTemporaria(String usuario,
												LinhaDadosArquivoIbptCsv ibptDados,
												ref String strDescricaoLog,
												ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			String strOperacao = "Gravação dos dados do arquivo IBPT na tabela temporária do BD";
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;
					strMsgErro = "";

					#region [ Tenta gravar o registro ]

					#region [ Preenche os parâmetros ]
					cmInsertTabelaTemporaria.Parameters["@codigo"].Value = ibptDados.codigo;
					cmInsertTabelaTemporaria.Parameters["@ex"].Value = ibptDados.ex;
					cmInsertTabelaTemporaria.Parameters["@tabela"].Value = ibptDados.tabela;
					cmInsertTabelaTemporaria.Parameters["@descricao"].Value = ibptDados.descricao;
					cmInsertTabelaTemporaria.Parameters["@aliqNac"].Value = ibptDados.aliqNac;
					cmInsertTabelaTemporaria.Parameters["@aliqImp"].Value = ibptDados.aliqImp;
					cmInsertTabelaTemporaria.Parameters["@percAliqNac"].Value = Global.converteNumeroDouble(ibptDados.aliqNac);
					cmInsertTabelaTemporaria.Parameters["@percAliqImp"].Value = Global.converteNumeroDouble(ibptDados.aliqImp);
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsertTabelaTemporaria.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = _bd.executaNonQuery(ref cmInsertTabelaTemporaria);
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
					strMsgErro = "Falha ao tentar gravar na tabela temporária do banco de dados os dados do arquivo do IBPT após " + intQtdeTentativas.ToString() + " tentativas!!";
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

		#region [ isTabelaTemporariaCriada ]
		public bool isTabelaTemporariaCriada()
		{
			#region [ Declarações ]
			bool blnTabelaTemporariaCriada = false;
			String strSql;
			String strNomeTabelaTemporaria;
			SqlCommand cmdConsulta;
			SqlDataReader drConsulta;
			#endregion

			cmdConsulta = _bd.criaSqlCommand();

			strNomeTabelaTemporaria = Global.Cte.ADM2.Tabelas.PREFIXO_TABELA_TEMPORARIA +
									  Global.Cte.ADM2.Tabelas.T_IBPT;
			
			strSql = "SELECT" +
						" name" +
					" FROM sysobjects" +
					" WHERE" +
						" (type = 'U')" +
						" AND (name = '" + strNomeTabelaTemporaria + "')";

			cmdConsulta.CommandText = strSql;
			drConsulta = cmdConsulta.ExecuteReader();
			try
			{
				while (drConsulta.Read())
				{
					if (drConsulta["name"].ToString().ToUpper().Equals(strNomeTabelaTemporaria.ToUpper()))
					{
						blnTabelaTemporariaCriada = true;
						break;
					}
				}
			}
			finally
			{
				drConsulta.Close();
			}

			return blnTabelaTemporariaCriada;
		}
		#endregion

		#region [ criaTabelaTemporaria ]
		public bool criaTabelaTemporaria()
		{
			#region [ Declarações ]
			String strSql;
			String strNomeTabelaTemporaria;
			String strMsgErro;
			String strOperacao = "Criação da tabela temporária no BD para gravação dos dados do arquivo do IBPT";
			SqlCommand cmdConsulta;
			#endregion

			try
			{
				if (isTabelaTemporariaCriada()) dropTabelaTemporaria();

				cmdConsulta = _bd.criaSqlCommand();

				strNomeTabelaTemporaria = Global.Cte.ADM2.Tabelas.PREFIXO_TABELA_TEMPORARIA +
										  Global.Cte.ADM2.Tabelas.T_IBPT;

				strSql = "SELECT TOP 0 *" +
							" INTO " + strNomeTabelaTemporaria +
						" FROM " + Global.Cte.ADM2.Tabelas.T_IBPT;
				cmdConsulta.CommandText = strSql;
				cmdConsulta.ExecuteNonQuery();
				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ dropTabelaTemporaria ]
		public bool dropTabelaTemporaria()
		{
			#region [ Declarações ]
			String strSql;
			String strNomeTabelaTemporaria;
			String strMsgErro;
			String strOperacao = "Exclui a tabela temporária do BD usada para gravação dos dados do arquivo do IBPT";
			SqlCommand cmdConsulta;
			#endregion

			try
			{
				if (!isTabelaTemporariaCriada()) return true;

				strNomeTabelaTemporaria = Global.Cte.ADM2.Tabelas.PREFIXO_TABELA_TEMPORARIA +
										  Global.Cte.ADM2.Tabelas.T_IBPT;

				strSql = "DROP TABLE " + strNomeTabelaTemporaria;
				cmdConsulta = _bd.criaSqlCommand();
				cmdConsulta.CommandText = strSql;
				cmdConsulta.ExecuteNonQuery();
				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ transfereDadosTabelaTemporariaParaTabelaProducao ]
		public bool transfereDadosTabelaTemporariaParaTabelaProducao(ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnIniciouTransacao = false;
			bool blnSucesso = false;
			String strSql;
			String strNomeTabelaTemporaria;
			String strOperacao = "Transfere os dados do IBPT da tabela temporária para a tabela de produção";
			SqlCommand cmdConsulta;
			#endregion

			strMsgErro = "";

			try
			{
				if (!isTabelaTemporariaCriada()) return false;

				if (!_bd.isTransacaoEmAndamento)
				{
					blnIniciouTransacao = true;
					_bd.iniciaTransacao();
				}

				try
				{
					cmdConsulta = _bd.criaSqlCommand();

					strNomeTabelaTemporaria = Global.Cte.ADM2.Tabelas.PREFIXO_TABELA_TEMPORARIA +
											  Global.Cte.ADM2.Tabelas.T_IBPT;

					#region [ Limpa a tabela de produção ]
					strSql = "DELETE FROM " + Global.Cte.ADM2.Tabelas.T_IBPT;
					cmdConsulta.CommandText = strSql;
					cmdConsulta.ExecuteNonQuery();
					#endregion

					#region [ Transfere os dados da tabela temporária para a tabela de produção ]
					strSql = "INSERT INTO " + Global.Cte.ADM2.Tabelas.T_IBPT + " SELECT * FROM " + strNomeTabelaTemporaria;
					cmdConsulta.CommandText = strSql;
					cmdConsulta.ExecuteNonQuery();
					#endregion

					blnSucesso = true;
				}
				finally
				{
					if (blnIniciouTransacao)
					{
						if (blnSucesso)
						{
							_bd.commitTransacao();
						}
						else
						{
							_bd.rollbackTransacao();
						}
					}
				}

				return blnSucesso;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion
	}
}
