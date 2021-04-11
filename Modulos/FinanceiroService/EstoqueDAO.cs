#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace FinanceiroService
{
	static class EstoqueDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsertLogEstoque;
		private static SqlCommand cmInsertEstoqueMovimento;
		private static SqlCommand cmUpdateEstoqueMovimentoAnula;
		private static SqlCommand cmUpdateEstoqueItemDevidoEstorno;
		private static SqlCommand cmUpdateEstoqueItemDevidoSaida;
		private static SqlCommand cmUpdateEstoqueDataUltMovimento;
		private static SqlCommand cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca;
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
		static EstoqueDAO()
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

			#region [ cmInsertLogEstoque ]
			strSql = "INSERT INTO t_ESTOQUE_LOG (" +
						"data, " +
						"data_hora, " +
						"usuario, " +
						"fabricante, " +
						"produto, " +
						"qtde_solicitada, " +
						"qtde_atendida, " +
						"operacao, " +
						"cod_estoque_origem, " +
						"cod_estoque_destino, " +
						"loja_estoque_origem, " +
						"loja_estoque_destino, " +
						"pedido_estoque_origem, " +
						"pedido_estoque_destino, " +
						"documento, " +
						"complemento, " +
						"id_ordem_servico, " +
						"id_nfe_emitente" +
					") VALUES (" +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario, " +
						"@fabricante, " +
						"@produto, " +
						"@qtde_solicitada, " +
						"@qtde_atendida, " +
						"@operacao, " +
						"@cod_estoque_origem, " +
						"@cod_estoque_destino, " +
						"@loja_estoque_origem, " +
						"@loja_estoque_destino, " +
						"@pedido_estoque_origem, " +
						"@pedido_estoque_destino, " +
						"@documento, " +
						"@complemento, " +
						"@id_ordem_servico, " +
						"@id_nfe_emitente" +
					")";
			cmInsertLogEstoque = BD.criaSqlCommand();
			cmInsertLogEstoque.CommandText = strSql;
			cmInsertLogEstoque.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertLogEstoque.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmInsertLogEstoque.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmInsertLogEstoque.Parameters.Add("@qtde_solicitada", SqlDbType.SmallInt);
			cmInsertLogEstoque.Parameters.Add("@qtde_atendida", SqlDbType.SmallInt);
			cmInsertLogEstoque.Parameters.Add("@operacao", SqlDbType.VarChar, 3);
			cmInsertLogEstoque.Parameters.Add("@cod_estoque_origem", SqlDbType.VarChar, 3);
			cmInsertLogEstoque.Parameters.Add("@cod_estoque_destino", SqlDbType.VarChar, 3);
			cmInsertLogEstoque.Parameters.Add("@loja_estoque_origem", SqlDbType.VarChar, 3);
			cmInsertLogEstoque.Parameters.Add("@loja_estoque_destino", SqlDbType.VarChar, 3);
			cmInsertLogEstoque.Parameters.Add("@pedido_estoque_origem", SqlDbType.VarChar, 9);
			cmInsertLogEstoque.Parameters.Add("@pedido_estoque_destino", SqlDbType.VarChar, 9);
			cmInsertLogEstoque.Parameters.Add("@documento", SqlDbType.VarChar, 30);
			cmInsertLogEstoque.Parameters.Add("@complemento", SqlDbType.VarChar, 80);
			cmInsertLogEstoque.Parameters.Add("@id_ordem_servico", SqlDbType.VarChar, 12);
			cmInsertLogEstoque.Parameters.Add("@id_nfe_emitente", SqlDbType.Int);
			cmInsertLogEstoque.Prepare();
			#endregion

			#region [ cmInsertEstoqueMovimento ]
			strSql = "INSERT INTO t_ESTOQUE_MOVIMENTO (" +
						"id_movimento, " +
						"data, " +
						"hora, " +
						"usuario, " +
						"pedido, " +
						"fabricante, " +
						"produto, " +
						"id_estoque, " +
						"qtde, " +
						"operacao, " +
						"estoque, " +
						"loja, " +
						"kit, " +
						"kit_id_estoque" +
					") VALUES (" +
						"@id_movimento, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"Replace(" + Global.sqlMontaGetdateSomenteHora() + ", ':', ''), " +
						"@usuario, " +
						"@pedido, " +
						"@fabricante, " +
						"@produto, " +
						"@id_estoque, " +
						"@qtde, " +
						"@operacao, " +
						"@estoque, " +
						"@loja, " +
						"@kit, " +
						"@kit_id_estoque" +
					")";
			cmInsertEstoqueMovimento = BD.criaSqlCommand();
			cmInsertEstoqueMovimento.CommandText = strSql;
			cmInsertEstoqueMovimento.Parameters.Add("@id_movimento", SqlDbType.VarChar, 12);
			cmInsertEstoqueMovimento.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertEstoqueMovimento.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertEstoqueMovimento.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmInsertEstoqueMovimento.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmInsertEstoqueMovimento.Parameters.Add("@id_estoque", SqlDbType.VarChar, 12);
			cmInsertEstoqueMovimento.Parameters.Add("@qtde", SqlDbType.SmallInt);
			cmInsertEstoqueMovimento.Parameters.Add("@operacao", SqlDbType.VarChar, 3);
			cmInsertEstoqueMovimento.Parameters.Add("@estoque", SqlDbType.VarChar, 3);
			cmInsertEstoqueMovimento.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmInsertEstoqueMovimento.Parameters.Add("@kit", SqlDbType.SmallInt);
			cmInsertEstoqueMovimento.Parameters.Add("@kit_id_estoque", SqlDbType.VarChar, 12);
			cmInsertEstoqueMovimento.Prepare();
			#endregion

			#region [ cmUpdateEstoqueMovimentoAnula ]
			strSql = "UPDATE t_ESTOQUE_MOVIMENTO SET " +
						"anulado_status = 1, " +
						"anulado_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"anulado_hora = Replace(" + Global.sqlMontaGetdateSomenteHora() + ", ':', ''), " +
						"anulado_usuario = @anulado_usuario " +
					"WHERE (id_movimento = @id_movimento)";
			cmUpdateEstoqueMovimentoAnula = BD.criaSqlCommand();
			cmUpdateEstoqueMovimentoAnula.CommandText = strSql;
			cmUpdateEstoqueMovimentoAnula.Parameters.Add("@id_movimento", SqlDbType.VarChar, 12);
			cmUpdateEstoqueMovimentoAnula.Parameters.Add("@anulado_usuario", SqlDbType.VarChar, 10);
			cmUpdateEstoqueMovimentoAnula.Prepare();
			#endregion

			#region [ cmUpdateEstoqueItemDevidoEstorno ]
			strSql = "UPDATE t_ESTOQUE_ITEM SET" +
						" qtde_utilizada = qtde_utilizada - @qtde_estorno," +
						" data_ult_movimento = " + Global.sqlMontaGetdateSomenteData() +
					" WHERE" +
						" (id_estoque = @id_estoque)" +
						" AND (fabricante = @fabricante)" +
						" AND (produto = @produto)";
			cmUpdateEstoqueItemDevidoEstorno = BD.criaSqlCommand();
			cmUpdateEstoqueItemDevidoEstorno.CommandText = strSql;
			cmUpdateEstoqueItemDevidoEstorno.Parameters.Add("@id_estoque", SqlDbType.VarChar, 12);
			cmUpdateEstoqueItemDevidoEstorno.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmUpdateEstoqueItemDevidoEstorno.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmUpdateEstoqueItemDevidoEstorno.Parameters.Add("@qtde_estorno", SqlDbType.SmallInt);
			cmUpdateEstoqueItemDevidoEstorno.Prepare();
			#endregion

			#region [ cmUpdateEstoqueItemDevidoSaida ]
			strSql = "UPDATE t_ESTOQUE_ITEM SET" +
						" qtde_utilizada = qtde_utilizada + @qtde_saida," +
						" data_ult_movimento = " + Global.sqlMontaGetdateSomenteData() +
					" WHERE" +
						" (id_estoque = @id_estoque)" +
						" AND (fabricante = @fabricante)" +
						" AND (produto = @produto)";
			cmUpdateEstoqueItemDevidoSaida = BD.criaSqlCommand();
			cmUpdateEstoqueItemDevidoSaida.CommandText = strSql;
			cmUpdateEstoqueItemDevidoSaida.Parameters.Add("@id_estoque", SqlDbType.VarChar, 12);
			cmUpdateEstoqueItemDevidoSaida.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmUpdateEstoqueItemDevidoSaida.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmUpdateEstoqueItemDevidoSaida.Parameters.Add("@qtde_saida", SqlDbType.SmallInt);
			cmUpdateEstoqueItemDevidoSaida.Prepare();
			#endregion

			#region [ cmUpdateEstoqueDataUltMovimento ]
			strSql = "UPDATE t_ESTOQUE SET" +
						" data_ult_movimento = " + Global.sqlMontaGetdateSomenteData() +
					" WHERE" +
						" (id_estoque = @id_estoque)";
			cmUpdateEstoqueDataUltMovimento = BD.criaSqlCommand();
			cmUpdateEstoqueDataUltMovimento.CommandText = strSql;
			cmUpdateEstoqueDataUltMovimento.Parameters.Add("@id_estoque", SqlDbType.VarChar, 12);
			cmUpdateEstoqueDataUltMovimento.Prepare();
			#endregion

			#region [ cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca ]
			strSql = "UPDATE t_ESTOQUE_MOVIMENTO SET" +
						" anulado_status = 1," +
						" anulado_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" anulado_hora = Replace(" + Global.sqlMontaGetdateSomenteHora() + ", ':', '')," +
						" anulado_usuario = @anulado_usuario" +
					" WHERE" +
						" (anulado_status = 0)" +
						" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA + "')" +
						" AND (pedido = @pedido)" +
						" AND (fabricante = @fabricante)" +
						" AND (produto = @produto)";
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca = BD.criaSqlCommand();
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.CommandText = strSql;
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters.Add("@anulado_usuario", SqlDbType.VarChar, 10);
			cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Prepare();
			#endregion
		}
		#endregion

		#region [ insereEstoqueMovimento ]
		public static bool insereEstoqueMovimento(String id_movimento,
												  String usuario,
												  String pedido,
												  String fabricante,
												  String produto,
												  String id_estoque,
												  int qtde,
												  String operacao,
												  String estoque,
												  String loja,
												  int kit,
												  String kit_id_estoque,
												  out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.insereEstoqueMovimento()";
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
					cmInsertEstoqueMovimento.Parameters["@id_movimento"].Value = id_movimento;
					cmInsertEstoqueMovimento.Parameters["@usuario"].Value = usuario;
					cmInsertEstoqueMovimento.Parameters["@pedido"].Value = pedido;
					cmInsertEstoqueMovimento.Parameters["@fabricante"].Value = fabricante;
					cmInsertEstoqueMovimento.Parameters["@produto"].Value = produto;
					cmInsertEstoqueMovimento.Parameters["@id_estoque"].Value = id_estoque;
					cmInsertEstoqueMovimento.Parameters["@qtde"].Value = qtde;
					cmInsertEstoqueMovimento.Parameters["@operacao"].Value = operacao;
					cmInsertEstoqueMovimento.Parameters["@estoque"].Value = estoque;
					cmInsertEstoqueMovimento.Parameters["@loja"].Value = loja;
					cmInsertEstoqueMovimento.Parameters["@kit"].Value = kit;
					cmInsertEstoqueMovimento.Parameters["@kit_id_estoque"].Value = kit_id_estoque;
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsertEstoqueMovimento.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertEstoqueMovimento);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados do registro = " + sbLog.ToString() + "\n" + ex.ToString());
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
					strMsgErro = "Falha ao tentar gravar no BD após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha: Dados do registro = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaEstoqueMovimentoAnula ]
		public static bool atualizaEstoqueMovimentoAnula(String id_movimento,
														 String anulado_usuario,
														 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.atualizaEstoqueMovimentoAnula()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_movimento == null)
				{
					strMsgErro = "Não foi informado o identificador do registro de movimento do estoque!!";
					return false;
				}

				if (id_movimento.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do registro de movimento do estoque!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateEstoqueMovimentoAnula.Parameters["@id_movimento"].Value = id_movimento;
				cmUpdateEstoqueMovimentoAnula.Parameters["@anulado_usuario"].Value = anulado_usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateEstoqueMovimentoAnula);
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
					strMsgErro = "Falha ao tentar atualizar o registro da tabela de movimentação do estoque (id_movimento=" + id_movimento + ")!!" + strMsgErro;
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

		#region [ atualizaEstoqueItemDevidoEstorno ]
		public static bool atualizaEstoqueItemDevidoEstorno(String id_estoque,
															String fabricante,
															String produto,
															int qtde_estorno,
														 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.atualizaEstoqueItemDevidoEstorno()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_estoque == null)
				{
					strMsgErro = "Não foi informado o identificador do registro do estoque!!";
					return false;
				}

				if (id_estoque.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do registro do estoque!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateEstoqueItemDevidoEstorno.Parameters["@id_estoque"].Value = id_estoque;
				cmUpdateEstoqueItemDevidoEstorno.Parameters["@fabricante"].Value = fabricante;
				cmUpdateEstoqueItemDevidoEstorno.Parameters["@produto"].Value = produto;
				cmUpdateEstoqueItemDevidoEstorno.Parameters["@qtde_estorno"].Value = qtde_estorno;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateEstoqueItemDevidoEstorno);
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
					strMsgErro = "Falha ao tentar atualizar o registro da tabela do estoque (id_estoque=" + id_estoque + ")!!" + strMsgErro;
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

		#region [ atualizaEstoqueItemDevidoSaida ]
		public static bool atualizaEstoqueItemDevidoSaida(String id_estoque,
														  String fabricante,
														  String produto,
														  int qtde_saida,
														  out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.atualizaEstoqueItemDevidoSaida()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_estoque == null)
				{
					strMsgErro = "Não foi informado o identificador do registro do estoque!!";
					return false;
				}

				if (id_estoque.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do registro do estoque!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateEstoqueItemDevidoSaida.Parameters["@id_estoque"].Value = id_estoque;
				cmUpdateEstoqueItemDevidoSaida.Parameters["@fabricante"].Value = fabricante;
				cmUpdateEstoqueItemDevidoSaida.Parameters["@produto"].Value = produto;
				cmUpdateEstoqueItemDevidoSaida.Parameters["@qtde_saida"].Value = qtde_saida;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateEstoqueItemDevidoSaida);
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
					strMsgErro = "Falha ao tentar atualizar o registro da tabela do estoque (id_estoque=" + id_estoque + ")!!" + strMsgErro;
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

		#region [ atualizaEstoqueDataUltMovimento ]
		public static bool atualizaEstoqueDataUltMovimento(String id_estoque,
														   out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.atualizaEstoqueDataUltMovimento()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (id_estoque == null)
				{
					strMsgErro = "Não foi informado o identificador do registro principal do estoque!!";
					return false;
				}

				if (id_estoque.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o identificador do registro principal do estoque!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateEstoqueDataUltMovimento.Parameters["@id_estoque"].Value = id_estoque;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateEstoqueDataUltMovimento);
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
					strMsgErro = "Falha ao tentar atualizar o registro principal do estoque (id_estoque=" + id_estoque + ")!!" + strMsgErro;
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

		#region [ atualizaEstoqueMovimentoCancelaPendenciaListaSemPresenca ]
		public static bool atualizaEstoqueMovimentoCancelaPendenciaListaSemPresenca(String pedido,
																					String fabricante,
																					String produto,
																					String anulado_usuario,
																					out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.atualizaEstoqueMovimentoCancelaPendenciaListaSemPresenca()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!!";
					return false;
				}

				if (pedido.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!!";
					return false;
				}

				if (fabricante == null)
				{
					strMsgErro = "Não foi informado o código do fabricante!!";
					return false;
				}

				if (fabricante.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o código do fabricante!!";
					return false;
				}

				if (produto == null)
				{
					strMsgErro = "Não foi informado o código do produto!!";
					return false;
				}

				if (produto.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o código do produto!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters["@pedido"].Value = pedido;
				cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters["@fabricante"].Value = fabricante;
				cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters["@produto"].Value = produto;
				cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca.Parameters["@anulado_usuario"].Value = anulado_usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateEstoqueMovimentoCancelaPendenciaListaSemPresenca);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					blnSucesso = false;
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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
					strMsgErro = "Falha ao tentar cancelar a pendência na lista de produtos sem presença no estoque (pedido=" + pedido + ", fabricante=" + fabricante + ", produto=" + produto + ")!!" + strMsgErro;
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

		#region [ estoquePedidoCancela ]
		/// <summary>
		/// Esta função processa o cancelamento do pedido com relação aos produtos no estoque.
		/// Portanto, os produtos que estiverem no "estoque vendido" serão estornados ao "estoque de venda".
		/// Os produtos que estiverem na lista de produtos vendidos "sem presença no estoque" serão cancelados.
		/// O log da movimentação no estoque (T_ESTOQUE_LOG) é gravado dentro das rotinas chamadas por esta:
		///		1) estoqueProdutoEstorna()
		///		2) estoqueProdutoCancelaListaSemPresenca()
		///	IMPORTANTE: sempre chame esta rotina dentro de uma transação para garantir a consistência dos registros entre as várias tabelas.
		/// </summary>
		/// <param name="usuario">Identificação do usuário</param>
		/// <param name="pedido">Número do pedido</param>
		/// <param name="strInfoLog">Informações para log</param>
		/// <param name="strMsgErro">Mensagem de erro, caso ocorra um</param>
		/// <returns>
		/// false - ocorreu falha ao tentar movimentar o estoque
		/// true - conseguiu fazer a movimentação do estoque
		/// </returns>
		public static bool estoquePedidoCancela(string usuario,
												string pedido,
												out string strInfoLog,
												out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.estoquePedidoCancela()";
			int qtde_estornada;
			int qtde_cancelada;
			String strSql;
			String strLogEstorno;
			String strLogCancela;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			List<cl_FABRICANTE_PRODUTO> vProduto = new List<cl_FABRICANTE_PRODUTO>();
			#endregion

			strInfoLog = "";
			strMsgErro = "";

			try
			{
				strLogEstorno = "";
				strLogCancela = "";

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT" +
							" fabricante," +
							" produto" +
						" FROM t_PEDIDO_ITEM" +
						" WHERE" +
							" (pedido = '" + pedido + "')";
				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					vProduto.Add(new cl_FABRICANTE_PRODUTO(BD.readToString(rowConsulta["fabricante"]), BD.readToString(rowConsulta["produto"])));
				}

				for (int i = 0; i < vProduto.Count; i++)
				{
					if (!estoqueProdutoEstorna(usuario, pedido, vProduto[i].fabricante, vProduto[i].produto, Global.Cte.Etc.COD_NEGATIVO_UM, out qtde_estornada, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao estornar o produto (" + vProduto[i].fabricante + ")" + vProduto[i].produto + " durante o cancelamento automático do pedido " + pedido + "!!" +
									 strMsgErro;
						return false;
					}

					if (!estoqueProdutoCancelaListaSemPresenca(usuario, pedido, vProduto[i].fabricante, vProduto[i].produto, Global.Cte.Etc.COD_NEGATIVO_UM, out qtde_cancelada, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao cancelar a pendência na lista de produtos sem presença no estoque do produto (" + vProduto[i].fabricante + ")" + vProduto[i].produto + " durante o cancelamento automático do pedido " + pedido + "!!" +
									 strMsgErro;
						return false;
					}

					if (qtde_estornada > 0) strLogEstorno += Global.logEstoqueMontaIncremento(qtde_estornada, vProduto[i].fabricante, vProduto[i].produto);
					if (qtde_cancelada > 0) strLogCancela += Global.logEstoqueMontaIncremento(qtde_cancelada, vProduto[i].fabricante, vProduto[i].produto);
				} // for (int i = 0; i < vProduto.Count; i++)

				if (strLogEstorno.Length > 0) strLogEstorno = "Produtos estornados do estoque vendido para o estoque de venda:" + strLogEstorno;
				if (strLogCancela.Length > 0) strLogCancela = "Produtos cancelados da lista de produtos vendidos sem presença no estoque:" + strLogCancela;

				strInfoLog = strLogEstorno;
				if (strInfoLog.Length > 0) strInfoLog += "\n";
				strInfoLog += strLogCancela;

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

		#region [ estoqueProcessaProdutosVendidosSemPresenca ]
		/// <summary>
		/// Esta função verifica a lista de produtos que foram vendidos sem presença no estoque para alocar os produtos
		/// que já estejam disponíveis aos pedidos mais antigos primeiro.
		/// O log da movimentação do estoque (T_ESTOQUE_LOG) é gravado dentro das rotinas chamadas por esta rotina:
		/// 	1) estoque_produto_vendido_sem_presenca_saida()
		/// IMPORTANTE: sempre chame esta rotina dentro de uma transação para garantir a consistência dos registros.
		/// </summary>
		/// <param name="usuario">Identificação do usuário</param>
		/// <param name="strMsgErro">Mensagem de erro, caso ocorra um</param>
		/// <returns>
		/// false - ocorreu falha ao tentar alterar os dados do estoque
		/// true - conseguiu alterar os dados do estoque
		/// </returns>
		public static bool estoqueProcessaProdutosVendidosSemPresenca(int id_nfe_emitente,
																	  string usuario,
																	  out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.estoqueProcessaProdutosVendidosSemPresenca()";
			bool blnAchou;
			int qtde_estoque_vendido;
			int qtde_estoque_sem_presenca;
			int total_estoque_sem_presenca;
			int total_estoque_vendido;
			String strStEntrega;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			List<cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA> v = new List<cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA>();
			List<String> vPedido = new List<String>();
			StringBuilder sbLog = new StringBuilder("");
			String strLog;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT DISTINCT" +
							" t_PEDIDO.data_hora," +
							" t_ESTOQUE_MOVIMENTO.pedido," +
							" t_ESTOQUE_MOVIMENTO.fabricante," +
							" t_ESTOQUE_MOVIMENTO.produto," +
							" (CASE" +
								" WHEN (t_PEDIDO__BASE.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK.ToString() + ") AND (t_PEDIDO.st_etg_imediata = " + Global.Cte.T_PEDIDO__ENTREGA_IMEDIATA_STATUS.ETG_IMEDIATA_SIM.ToString() + ") THEN 1" +
								" WHEN (t_PEDIDO__BASE.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK.ToString() + ") AND (t_PEDIDO.st_etg_imediata = " + Global.Cte.T_PEDIDO__ENTREGA_IMEDIATA_STATUS.ETG_IMEDIATA_NAO.ToString() + ") THEN 2" +
								" ELSE 9" +
							" END) AS Prioridade" +
						" FROM t_ESTOQUE_MOVIMENTO" +
							" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" +
							" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" +
							" INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto))" +
						" WHERE" +
							" (anulado_status = 0)" +
							" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA + "')" +
							" AND ((t_ESTOQUE_ITEM.qtde - t_ESTOQUE_ITEM.qtde_utilizada) > 0)";

				if (id_nfe_emitente > 0)
				{
					strSql += " AND (t_PEDIDO.id_nfe_emitente = " + id_nfe_emitente.ToString() + ")";
				}

				strSql += " ORDER BY" +
							" Prioridade," +
							" t_PEDIDO.data_hora";

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Não há o que processar!! ]
				if (dtbConsulta.Rows.Count == 0) return true;
				#endregion

				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					v.Add(new cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA(BD.readToString(rowConsulta["pedido"]), BD.readToString(rowConsulta["fabricante"]), BD.readToString(rowConsulta["produto"])));
				}

				#region [ Os pedidos mais antigos devem atendidos primeiro ]
				for (int i = 0; i < v.Count; i++)
				{
					if (!estoqueProdutoVendidoSemPresencaSaida(usuario, v[i].pedido, v[i].fabricante, v[i].produto, out qtde_estoque_vendido, out qtde_estoque_sem_presenca, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar processar a saída da lista de produtos sem presença no estoque (pedido=" + v[i].pedido + ", fabricante=" + v[i].fabricante + ", produto=" + v[i].produto + ")!!" +
									 strMsgErro;
						return false;
					}

					#region [ Se houve produto alocado p/ este pedido, então inclui o pedido na lista que será analisada quanto ao "status de entrega" ]
					if (qtde_estoque_vendido > 0)
					{
						blnAchou = false;
						for (int j = (vPedido.Count - 1); j >= 0; j--)
						{
							if (vPedido[j].Equals(v[i].pedido))
							{
								blnAchou = true;
								break;
							}
						}

						if (!blnAchou)
						{
							vPedido.Add(v[i].pedido);
						}

						#region [ Informações para o log ]
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(v[i].pedido + Global.logProdutoMonta(qtde_estoque_vendido, v[i].fabricante, v[i].produto) + " SPE=" + qtde_estoque_sem_presenca.ToString());
						#endregion
					} // if (qtde_estoque_vendido > 0)
					#endregion
				} // for (int i = 0; i < v.Count; i++)
				#endregion

				#region [ Atualiza o "status de entrega" dos pedidos ]
				for (int i = 0; i < vPedido.Count; i++)
				{
					total_estoque_sem_presenca = 0;

					strSql = "SELECT" +
								" Coalesce(Sum(qtde), 0) AS total" +
							" FROM t_ESTOQUE_MOVIMENTO" +
							" WHERE" +
								" (anulado_status = 0)" +
								" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA + "')" +
								" AND (pedido = '" + vPedido[i] + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					if (dtbConsulta.Rows.Count > 0)
					{
						rowConsulta = dtbConsulta.Rows[0];
						total_estoque_sem_presenca = BD.readToInt(rowConsulta["total"]);
					}

					total_estoque_vendido = 0;

					strSql = "SELECT" +
								" Coalesce(Sum(qtde), 0) AS total" +
							" FROM t_ESTOQUE_MOVIMENTO" +
							" WHERE" +
								" (anulado_status = 0)" +
								" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_VENDIDO + "')" +
								" AND (pedido = '" + vPedido[i] + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					if (dtbConsulta.Rows.Count > 0)
					{
						rowConsulta = dtbConsulta.Rows[0];
						total_estoque_vendido = BD.readToInt(rowConsulta["total"]);
					}

					strSql = "SELECT " +
								"*" +
							" FROM t_PEDIDO" +
							" WHERE" +
								" (pedido = '" + vPedido[i] + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					if (dtbConsulta.Rows.Count == 0)
					{
						strMsgErro = "Pedido " + vPedido[i] + " não foi encontrado!!";
						return false;
					}

					rowConsulta = dtbConsulta.Rows[0];

					#region [ Status de entrega ]
					if (total_estoque_vendido == 0)
					{
						strStEntrega = Global.Cte.StEntregaPedido.ST_ENTREGA_ESPERAR;
					}
					else if (total_estoque_sem_presenca == 0)
					{
						strStEntrega = Global.Cte.StEntregaPedido.ST_ENTREGA_SEPARAR;
					}
					else
					{
						strStEntrega = Global.Cte.StEntregaPedido.ST_ENTREGA_SPLIT_POSSIVEL;
					}
					#endregion

					if (!strStEntrega.Equals(BD.readToString(rowConsulta["st_entrega"])))
					{
						if (!PedidoDAO.atualizaPedidoStEntrega(vPedido[i], strStEntrega, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar atualizar o status de entrega do pedido (pedido=" + vPedido[i] + ")!!" +
										 strMsgErro;
							return false;
						}
					}
				} // for (int i = 0; i < vPedido.Count; i++)
				#endregion

				if (sbLog.Length > 0)
				{
					strLog = "Processamento automático da lista de produtos vendidos sem presença no estoque: " + sbLog.ToString();
					if (!GeralDAO.gravaLog(usuario, "", "", "", Global.Cte.LogBd.Operacao.OP_LOG_ESTOQUE_PROCESSA_SP, strLog, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao gravar o log da operação!!" +
									 strMsgErro;
						return false;
					}
				}

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

		#region [ estoqueProdutoVendidoSemPresencaSaida ]
		public static bool estoqueProdutoVendidoSemPresencaSaida(String usuario,
																 String pedido,
																 String fabricante,
																 String produto,
																 out int qtde_estoque_vendido,
																 out int qtde_estoque_sem_presenca,
																 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.estoqueProdutoVendidoSemPresencaSaida()";
			int qtde_a_sair;
			int qtde_disponivel;
			int qtde_movimentada;
			int qtde_movto;
			int qtde_aux;
			int qtde_utilizada_aux;
			int intKit;
			int id_nfe_emitente;
			String strLoja;
			String strKitIdEstoque;
			String strSql;
			String strNsuIdMovimento;
			String strIdEstoque;
			String strLojaEstoqueOrigem;
			String strLojaEstoqueDestino;
			String strDocumento;
			String strComplemento;
			String strIdOrdemServico;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			List<String> v_estoque = new List<String>();
			#endregion

			qtde_estoque_vendido = 0;
			qtde_estoque_sem_presenca = 0;
			strMsgErro = "";

			try
			{
				#region [ Consistência ]
				if (pedido == null) return true;
				if (pedido.Trim().Length == 0) return true;

				if (produto == null) return true;
				if (produto.Trim().Length == 0) return true;
				#endregion

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Obtém a quantidade vendida sem presença no estoque ]
				strSql = "SELECT" +
							" Coalesce(Sum(qtde), 0) AS total" +
						" FROM t_ESTOQUE_MOVIMENTO" +
						" WHERE" +
							" (anulado_status = 0)" +
							" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA + "')" +
							" AND (pedido = '" + pedido + "')" +
							" AND (fabricante = '" + fabricante + "')" +
							" AND (produto = '" + produto + "')";

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				qtde_a_sair = 0;
				if (dtbConsulta.Rows.Count > 0)
				{
					rowConsulta = dtbConsulta.Rows[0];
					qtde_a_sair = BD.readToInt(rowConsulta["total"]);
				}
				#endregion

				#region [ Não há produtos pendentes na lista de "sem presença" ]
				if (qtde_a_sair <= 0) return true;
				#endregion

				#region [ Obtém a empresa (CD) do pedido ]
				strSql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" + pedido + "')";
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				if (dtbConsulta.Rows.Count == 0)
				{
					strMsgErro = "Falha ao tentar localizar o registro do pedido " + pedido;
					return false;
				}

				rowConsulta= dtbConsulta.Rows[0];
				id_nfe_emitente = BD.readToInt(rowConsulta["id_nfe_emitente"]);
				#endregion

				#region [ Obtém os "lotes" do produto disponíveis no estoque (política FIFO) ]
				strSql = "SELECT" +
							" t_ESTOQUE.id_estoque," +
							" (qtde - qtde_utilizada) AS saldo" +
						" FROM t_ESTOQUE" +
							" INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" +
						" WHERE" +
							" (t_ESTOQUE.id_nfe_emitente = " + id_nfe_emitente.ToString() + ")" +
							" AND (t_ESTOQUE.fabricante = '" + fabricante + "')" +
							" AND (produto = '" + produto + "')" +
							" AND ((qtde - qtde_utilizada) > 0)" +
						" ORDER BY" +
							" data_entrada," +
							" t_ESTOQUE.id_estoque";
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);

				#region [ Armazena as entradas no estoque candidatas à saída de produtos ]
				qtde_disponivel = 0;
				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					v_estoque.Add(BD.readToString(rowConsulta["id_estoque"]));
					qtde_disponivel += BD.readToInt(rowConsulta["saldo"]);
				}
				#endregion

				#endregion

				#region [ Há produtos disponíveis? ]
				if (qtde_disponivel <= 0) return true;
				#endregion

				#region [ Realiza a saída do estoque!! ]
				qtde_movimentada = 0;
				for (int iv = 0; iv < v_estoque.Count; iv++)
				{
					#region [ A quantidade necessária já foi retirada do estoque!! ]
					if (qtde_movimentada >= qtde_a_sair) break;
					#endregion

					#region [ t_ESTOQUE_ITEM: saída de produtos ]
					strSql = "SELECT" +
								" qtde," +
								" qtde_utilizada," +
								" data_ult_movimento" +
							" FROM t_ESTOQUE_ITEM" +
							" WHERE" +
								" (id_estoque = '" + v_estoque[iv] + "')" +
								" AND (fabricante = '" + fabricante + "')" +
								" AND (produto = '" + produto + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);

					qtde_movto = 0;
					qtde_aux = 0;
					qtde_utilizada_aux = 0;
					if (dtbConsulta.Rows.Count > 0)
					{
						rowConsulta = dtbConsulta.Rows[0];
						qtde_aux = BD.readToInt(rowConsulta["qtde"]);
						qtde_utilizada_aux = BD.readToInt(rowConsulta["qtde_utilizada"]);
						if ((qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux))
						{
							// Quantidade de produtos deste item de estoque é insuficiente p/ atender o pedido
							qtde_movto = qtde_aux - qtde_utilizada_aux;
						}
						else
						{
							// Quantidade de produtos deste item sozinho é suficiente p/ atender o pedido
							qtde_movto = qtde_a_sair - qtde_movimentada;
						}

						#region [ Atualiza o registro em t_ESTOQUE_ITEM p/ contabilizar a saída ]
						if (!atualizaEstoqueItemDevidoSaida(v_estoque[iv], fabricante, produto, qtde_movto, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar alterar o registro do estoque do produto (" + fabricante + ")" + produto + " (id_estoque=" + v_estoque[iv] + ")" +
										 strMsgErro;
							return false;
						}
						#endregion
					} // if (dtbConsulta.Rows.Count > 0)

					#endregion

					#region [ Contabiliza quantidade movimentada ]
					qtde_movimentada = qtde_movimentada + qtde_movto;
					#endregion

					#region [ Registra o movimento de saída no estoque ]
					if (qtde_movto > 0)
					{
						// Registra o movimento que contabiliza os produtos vendidos sem presença no estoque
						if (!GeralDAO.geraNsuUsandoTabelaControle(Global.Cte.Nsu.NSU_ID_ESTOQUE_MOVTO, out strNsuIdMovimento, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar gerar o identificador do registro de movimento do estoque!!" +
										 strMsgErro;
							return false;
						}

						strLoja = "";
						intKit = 0;
						strKitIdEstoque = "";
						if (!insereEstoqueMovimento(strNsuIdMovimento, usuario, pedido, fabricante, produto, v_estoque[iv], qtde_movto, Global.Cte.OperacaoMovimentoEstoque.OP_ESTOQUE_VENDA, Global.Cte.TipoEstoque.ID_ESTOQUE_VENDIDO, strLoja, intKit, strKitIdEstoque, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar gravar o movimento de saída do estoque!!" +
										 strMsgErro;
							return false;
						}

						#region [ t_ESTOQUE: atualiza data do último movimento ]
						if (!atualizaEstoqueDataUltMovimento(v_estoque[iv], out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar alterar o registro principal do estoque do produto (" + fabricante + ")" + produto + " (id_estoque=" + v_estoque[iv] + ")" +
										 strMsgErro;
							return false;
						}
						#endregion
					}
					#endregion

					#region [ Já conseguiu alocar tudo? ]
					if (qtde_movimentada >= qtde_a_sair) break;
					#endregion
				} // for (int iv = 0; iv < v_estoque.Count; iv++)
				#endregion

				#region [ Anula o registro do produto deste pedido na lista "sem presença no estoque" ]
				if (!atualizaEstoqueMovimentoCancelaPendenciaListaSemPresenca(pedido, fabricante, produto, usuario, out strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar cancelar a pendência na lista de produtos sem presença no estoque (pedido=" + pedido + ", fabricante=" + fabricante + ", produto=" + produto + ")!!" +
								 strMsgErro;
					return false;
				}
				#endregion

				#region [ Resíduo faltante: registra a venda sem presença no estoque p/ a diferença que ainda falta ]
				if (qtde_movimentada < qtde_a_sair)
				{
					#region [ Registra o movimento de saída no estoque ]
					if (!GeralDAO.geraNsuUsandoTabelaControle(Global.Cte.Nsu.NSU_ID_ESTOQUE_MOVTO, out strNsuIdMovimento, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar gerar o identificador do registro de movimento do estoque!!" +
									 strMsgErro;
						return false;
					}

					qtde_estoque_sem_presenca = qtde_a_sair - qtde_movimentada;

					strIdEstoque = "";
					strLoja = "";
					intKit = 0;
					strKitIdEstoque = "";
					if (!insereEstoqueMovimento(strNsuIdMovimento, usuario, pedido, fabricante, produto, strIdEstoque, qtde_estoque_sem_presenca, Global.Cte.OperacaoMovimentoEstoque.OP_ESTOQUE_VENDA, Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA, strLoja, intKit, strKitIdEstoque, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar gravar o movimento do estoque referente ao saldo residual na lista de produtos sem presença (pedido=" + pedido + ", fabricante=" + fabricante + ", produto=" + produto + ")!!" +
									 strMsgErro;
						return false;
					}
					#endregion
				}
				#endregion

				qtde_estoque_vendido = qtde_movimentada;

				#region [ Log de movimentação do estoque ]
				strLojaEstoqueOrigem = "";
				strLojaEstoqueDestino = "";
				strDocumento = "";
				strComplemento = "";
				strIdOrdemServico = "";
				if (!gravaLogEstoque(usuario, id_nfe_emitente, fabricante, produto, qtde_a_sair, qtde_estoque_vendido, Global.Cte.OperacaoLogEstoque.OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA, Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA, Global.Cte.TipoEstoque.ID_ESTOQUE_VENDIDO, strLojaEstoqueOrigem, strLojaEstoqueDestino, pedido, pedido, strDocumento, strComplemento, strIdOrdemServico, out strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gravar o log da movimentação no estoque!!" +
								strMsgErro;
					return false;
				}
				#endregion

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

		#region [ estoqueProdutoEstorna ]
		/// <summary>
		/// Esta função estorna a quantidade de produtos indicada no parâmetro 'qtde_a_estornar' do 'estoque vendido' para o 'estoque de venda'.
		/// Se o parâmetro 'qtde_a_estornar' for especificado com o valor 'COD_NEGATIVO_UM', então o estorno será integral.
		/// IMPORTANTE: sempre chame esta rotina dentro de uma transação para garantir a consistência dos registros entre as várias tabelas.
		/// </summary>
		/// <param name="usuario">Identificação do usuário</param>
		/// <param name="pedido">Número do pedido</param>
		/// <param name="fabricante">Código do fabricante do produto</param>
		/// <param name="produto">Código do produto</param>
		/// <param name="qtde_a_estornar">Quantidade a estornar</param>
		/// <param name="qtde_estornada">Quantidade estornada</param>
		/// <param name="strMsgErro">Mensagem de erro, caso ocorra um</param>
		/// <returns>
		/// false - ocorreu falha ao tentar movimentar o estoque.
		/// true - conseguiu fazer a movimentação do estoque.
		/// </returns>
		public static bool estoqueProdutoEstorna(string usuario,
												 string pedido,
												 string fabricante,
												 string produto,
												 int qtde_a_estornar,
												 out int qtde_estornada,
												 out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.estoqueProdutoEstorna()";
			int qtde_aux;
			int qtde_movto;
			int qtde_utilizada_aux;
			int intKit;
			int intQtdeEstornarAux;
			int id_nfe_emitente;
			bool blnGravarLog;
			String strLoja;
			String strKitIdEstoque;
			String strSql;
			String operacao_aux;
			String id_estoque_aux;
			String strNsuIdMovimento;
			String strLojaEstoqueOrigem;
			String strLojaEstoqueDestino;
			String strPedidoEstoqueDestino;
			String strDocumento;
			String strComplemento;
			String strIdOrdemServico;
			List<String> v_estoque = new List<String>();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			qtde_estornada = 0;
			strMsgErro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Obtém a empresa (CD) do pedido ]
				strSql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" + pedido + "')";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				if (dtbConsulta.Rows.Count == 0)
				{
					strMsgErro = "Falha ao tentar localizar o registro do pedido " + pedido;
					return false;
				}
				rowConsulta = dtbConsulta.Rows[0];
				id_nfe_emitente = BD.readToInt(rowConsulta["id_nfe_emitente"]);
				#endregion

				// 1) LEMBRE-SE DE QUE PODE HAVER MAIS DE UM REGISTRO EM T_ESTOQUE_MOVIMENTO 
				//    P/ CADA PRODUTO, POIS PODEM TER SIDO USADOS DIFERENTES LOTES DO ESTOQUE 
				//    P/ ATENDER A UM ÚNICO PEDIDO!!
				// 2) LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
				//    OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
				//    FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
				strSql = "SELECT" +
							" id_movimento" +
						" FROM t_ESTOQUE" +
							" INNER JOIN t_ESTOQUE_MOVIMENTO ON (t_ESTOQUE.id_estoque = t_ESTOQUE_MOVIMENTO.id_estoque)" +
						" WHERE" +
							" (anulado_status = 0)" +
							" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_VENDIDO + "')" +
							" AND (pedido = '" + pedido + "')" +
							" AND (t_ESTOQUE_MOVIMENTO.fabricante = '" + fabricante + "')" +
							" AND (produto = '" + produto + "')" +
						" ORDER BY" +
							" t_ESTOQUE.data_entrada DESC," +
							" t_ESTOQUE.id_estoque DESC";

				#region [ Executa a consulta no BD ]
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					v_estoque.Add(BD.readToString(rowConsulta["id_movimento"]));
				}

				for (int i = 0; i < v_estoque.Count; i++)
				{
					#region [ Já estornou tudo? ]
					if (qtde_a_estornar != Global.Cte.Etc.COD_NEGATIVO_UM)
					{
						if (qtde_estornada >= qtde_a_estornar) break;
					}
					#endregion

					#region [ Recupera o registro do movimento ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ESTOQUE_MOVIMENTO" +
							" WHERE" +
								" (anulado_status = 0)" +
								" AND (id_movimento = '" + v_estoque[i] + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					if (dtbConsulta.Rows.Count == 0)
					{
						strMsgErro = "Falha ao acessar o registro de movimento no estoque do produto (" + fabricante + ")" + produto + " (id_movimento=" + v_estoque[i] + ")";
						return false;
					}

					rowConsulta = dtbConsulta.Rows[0];
					#endregion

					#region [ Memoriza os dados do movimento ]
					id_estoque_aux = BD.readToString(rowConsulta["id_estoque"]);
					qtde_aux = BD.readToInt(rowConsulta["qtde"]);
					operacao_aux = BD.readToString(rowConsulta["operacao"]);
					qtde_movto = qtde_aux;
					#endregion

					#region [ É para estornar tudo ou uma quantidade especificada? ]
					if (qtde_a_estornar != Global.Cte.Etc.COD_NEGATIVO_UM)
					{
						// A quantidade que falta ser estornada é menor que a quantidade do movimento
						if ((qtde_a_estornar - qtde_estornada) < qtde_aux)
						{
							qtde_movto = qtde_a_estornar - qtde_estornada;
						}
					}
					#endregion

					#region [ Anula o movimento ]
					if (!atualizaEstoqueMovimentoAnula(v_estoque[i], Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar alterar o registro de movimento no estoque do produto (" + fabricante + ")" + produto + " (id_movimento=" + v_estoque[i] + ")" +
									 strMsgErro;
						return false;
					}
					#endregion

					#region [ Estorno parcial ]
					// Estorno parcial: o movimento original foi anulado e um novo movimento c/ a quantidade restante deve ser gravado!!
					if (qtde_movto < qtde_aux)
					{
						// Registra o movimento que contabiliza os produtos vendidos sem presença no estoque
						if (!GeralDAO.geraNsuUsandoTabelaControle(Global.Cte.Nsu.NSU_ID_ESTOQUE_MOVTO, out strNsuIdMovimento, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar gerar o identificador do registro de movimento do estoque!!" +
										 strMsgErro;
							return false;
						}

						strLoja = "";
						intKit = 0;
						strKitIdEstoque = "";
						if (!insereEstoqueMovimento(strNsuIdMovimento, usuario, pedido, fabricante, produto, id_estoque_aux, (qtde_aux - qtde_movto), operacao_aux, Global.Cte.TipoEstoque.ID_ESTOQUE_VENDIDO, strLoja, intKit, strKitIdEstoque, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar gravar o movimento do estoque com o saldo residual!!" +
										 strMsgErro;
							return false;
						}
					} // if (qtde_movto < qtde_aux)
					#endregion

					#region [ t_ESTOQUE_ITEM: estorna produtos ao saldo ]
					strSql = "SELECT" +
								" data_ult_movimento," +
								" qtde_utilizada" +
							" FROM t_ESTOQUE_ITEM" +
							" WHERE" +
								" (id_estoque = '" + id_estoque_aux + "')" +
								" AND (fabricante = '" + fabricante + "')" +
								" AND (produto = '" + produto + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					if (dtbConsulta.Rows.Count == 0)
					{
						strMsgErro = "Falha ao acessar o registro do estoque do produto (" + fabricante + ")" + produto + " (id_estoque=" + id_estoque_aux + ")";
						return false;
					}

					rowConsulta = dtbConsulta.Rows[0];

					qtde_utilizada_aux = BD.readToInt(rowConsulta["qtde_utilizada"]);

					// Precaução (p/ garantir que "qtde_utilizada" nunca ficará c/ valor negativo)!!
					intQtdeEstornarAux = qtde_movto;
					if (qtde_utilizada_aux < qtde_movto) intQtdeEstornarAux = qtde_utilizada_aux;

					if (!atualizaEstoqueItemDevidoEstorno(id_estoque_aux, fabricante, produto, intQtdeEstornarAux, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar alterar o registro do estoque do produto (" + fabricante + ")" + produto + " (id_estoque=" + id_estoque_aux + ")" +
									 strMsgErro;
						return false;
					}

					#region [ Contabiliza quantidade estornada ]
					qtde_estornada = qtde_estornada + qtde_movto;
					#endregion

					#region [ t_ESTOQUE: atualiza data do último movimento ]
					if (!atualizaEstoqueDataUltMovimento(id_estoque_aux, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar alterar o registro principal do estoque do produto (" + fabricante + ")" + produto + " (id_estoque=" + id_estoque_aux + ")" +
									 strMsgErro;
						return false;
					}
					#endregion

					#endregion
				} // for (int i = 0; i < v_estoque.Count; i++)

				blnGravarLog = true;
				if ((qtde_a_estornar == Global.Cte.Etc.COD_NEGATIVO_UM) && (qtde_estornada == 0)) blnGravarLog = false;

				if (blnGravarLog)
				{
					strLojaEstoqueOrigem = "";
					strLojaEstoqueDestino = "";
					strPedidoEstoqueDestino = "";
					strDocumento = "";
					strComplemento = "";
					strIdOrdemServico = "";
					if (!gravaLogEstoque(usuario, id_nfe_emitente, fabricante, produto, qtde_a_estornar, qtde_estornada, Global.Cte.OperacaoLogEstoque.OP_ESTOQUE_LOG_ESTORNO, Global.Cte.TipoEstoque.ID_ESTOQUE_VENDIDO, Global.Cte.TipoEstoque.ID_ESTOQUE_VENDA, strLojaEstoqueOrigem, strLojaEstoqueDestino, pedido, strPedidoEstoqueDestino, strDocumento, strComplemento, strIdOrdemServico, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao gravar o log da movimentação no estoque!!" +
									strMsgErro;
						return false;
					}
				}

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

		#region [ estoqueProdutoCancelaListaSemPresenca ]
		/// <summary>
		/// Esta função cancela a quantidade de produtos indicada no parâmetro 'qtde_a_cancelar' da lista de produtos vendidos sem presença no estoque.
		/// Se o parâmetro 'qtde_a_cancelar' for especificado com o valor 'COD_NEGATIVO_UM', então o cancelamento será integral.
		/// IMPORTANTE: sempre chame esta rotina dentro de uma transação para garantir a consistência dos registros entre as várias tabelas.
		/// </summary>
		/// <param name="usuario">Identificação do usuário</param>
		/// <param name="pedido">Número do pedido</param>
		/// <param name="fabricante">Código do fabricante do produto</param>
		/// <param name="produto">Código do produto</param>
		/// <param name="qtde_a_cancelar">Quantidade a cancelar</param>
		/// <param name="qtde_cancelada">Quantidade que a rotina conseguiu cancelar</param>
		/// <param name="strMsgErro">Mensagem de erro, caso ocorra um</param>
		/// <returns>
		/// false - ocorreu falha ao tentar movimentar o estoque
		/// true - conseguiu fazer a movimentação do estoque
		/// </returns>
		public static bool estoqueProdutoCancelaListaSemPresenca(string usuario,
																 string pedido,
																 string fabricante,
																 string produto,
																 int qtde_a_cancelar,
																 out int qtde_cancelada,
																 out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EstoqueDAO.estoqueProdutoCancelaListaSemPresenca()";
			bool blnGravarLog;
			int qtde_aux;
			int qtde_movto;
			int intKit;
			int id_nfe_emitente;
			String strIdEstoque;
			String strLoja;
			String strKitIdEstoque;
			String operacao_aux;
			String strSql;
			String strNsuIdMovimento;
			String strCodEstoqueDestino;
			String strLojaEstoqueOrigem;
			String strLojaEstoqueDestino;
			String strPedidoEstoqueDestino;
			String strDocumento;
			String strComplemento;
			String strIdOrdemServico;
			List<String> v_estoque = new List<String>();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			qtde_cancelada = 0;
			strMsgErro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Obtém a empresa (CD) do pedido ]
				strSql = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" + pedido + "')";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				if (dtbConsulta.Rows.Count == 0)
				{
					strMsgErro = "Falha ao tentar localizar o registro do pedido " + pedido;
					return false;
				}

				rowConsulta = dtbConsulta.Rows[0];
				id_nfe_emitente = BD.readToInt(rowConsulta["id_nfe_emitente"]);
				#endregion

				// LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
				// OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
				// FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
				strSql = "SELECT" +
							" id_movimento" +
						" FROM t_ESTOQUE_MOVIMENTO" +
						" WHERE" +
							" (anulado_status = 0)" +
							" AND (estoque = '" + Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA + "')" +
							" AND (pedido = '" + pedido + "')" +
							" AND (fabricante = '" + fabricante + "')" +
							" AND (produto = '" + produto + "')";

				#region [ Executa a consulta no BD ]
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					v_estoque.Add(BD.readToString(rowConsulta["id_movimento"]));
				}

				for (int i = 0; i < v_estoque.Count; i++)
				{
					#region [ Já cancelou tudo? ]
					if (qtde_a_cancelar != Global.Cte.Etc.COD_NEGATIVO_UM)
					{
						if (qtde_cancelada >= qtde_a_cancelar) break;
					}
					#endregion

					#region [ Obtém o registro do movimento ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ESTOQUE_MOVIMENTO" +
							" WHERE" +
								" (anulado_status = 0)" +
								" AND (id_movimento = '" + v_estoque[i] + "')";
					dtbConsulta.Reset();
					cmCommand.CommandText = strSql;
					daAdapter.Fill(dtbConsulta);
					if (dtbConsulta.Rows.Count == 0)
					{
						strMsgErro = "Falha ao acessar o registro de movimento no estoque do produto " + produto + " do fabricante " + fabricante + " (id_movimento=" + v_estoque[i] + ")";
						return false;
					}

					rowConsulta = dtbConsulta.Rows[0];
					#endregion

					#region [ Memoriza os dados do movimento ]
					qtde_aux = BD.readToInt(rowConsulta["qtde"]);
					operacao_aux = BD.readToString(rowConsulta["operacao"]);
					qtde_movto = qtde_aux;
					#endregion

					#region [ É para cancelar tudo ou uma quantidade especificada? ]
					if (qtde_a_cancelar != Global.Cte.Etc.COD_NEGATIVO_UM)
					{
						// A quantidade que falta ser cancelada é menor que a quantidade do movimento
						if ((qtde_a_cancelar - qtde_cancelada) < qtde_aux)
						{
							qtde_movto = qtde_a_cancelar - qtde_cancelada;
						}
					}
					#endregion

					#region [ Anula o movimento ]
					if (!atualizaEstoqueMovimentoAnula(v_estoque[i], Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao tentar alterar o registro de movimento no estoque do produto " + produto + " do fabricante " + fabricante + " (id_movimento=" + v_estoque[i] + ")" +
									 strMsgErro;
						return false;
					}
					#endregion

					#region [ Cancelamento parcial ]
					// Cancelamento parcial: o movimento original foi anulado e um novo movimento c/ a quantidade restante deve ser gravado!!
					if (qtde_movto < qtde_aux)
					{
						// Registra o movimento que contabiliza os produtos vendidos sem presença no estoque
						if (!GeralDAO.geraNsuUsandoTabelaControle(Global.Cte.Nsu.NSU_ID_ESTOQUE_MOVTO, out strNsuIdMovimento, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar gerar o identificador do registro de movimento do estoque!!" +
										 strMsgErro;
							return false;
						}

						strIdEstoque = "";
						strLoja = "";
						intKit = 0;
						strKitIdEstoque = "";
						if (!insereEstoqueMovimento(strNsuIdMovimento, usuario, pedido, fabricante, produto, strIdEstoque, (qtde_aux - qtde_movto), operacao_aux, Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA, strLoja, intKit, strKitIdEstoque, out strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao tentar gravar o movimento do estoque com o saldo residual!!" +
										 strMsgErro;
							return false;
						}
					} // if (qtde_movto < qtde_aux)
					#endregion

					#region [ Contabiliza a quantidade cancelada ]
					qtde_cancelada += qtde_movto;
					#endregion
				}  // for

				blnGravarLog = true;
				if ((qtde_a_cancelar == Global.Cte.Etc.COD_NEGATIVO_UM) && (qtde_cancelada == 0)) blnGravarLog = false;

				if (blnGravarLog)
				{
					#region [ Grava o log de movimentação do estoque ]
					strCodEstoqueDestino = "";
					strLojaEstoqueOrigem = "";
					strLojaEstoqueDestino = "";
					strPedidoEstoqueDestino = "";
					strDocumento = "";
					strComplemento = "";
					strIdOrdemServico = "";
					if (!gravaLogEstoque(usuario, id_nfe_emitente, fabricante, produto, qtde_a_cancelar, qtde_cancelada, Global.Cte.OperacaoLogEstoque.OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA, Global.Cte.TipoEstoque.ID_ESTOQUE_SEM_PRESENCA, strCodEstoqueDestino, strLojaEstoqueOrigem, strLojaEstoqueDestino, pedido, strPedidoEstoqueDestino, strDocumento, strComplemento, strIdOrdemServico, out strMsgErro))
					{
						strMsgErro = "Falha ao gravar o log da movimentação no estoque!!";
						return false;
					}
					#endregion
				}

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

		#region [ gravaLogEstoque ]
		public static bool gravaLogEstoque(String strUsuario,
											int id_nfe_emitente,
											String strFabricante,
											String strProduto,
											int intQtdeSolicitada,
											int intQtdeAtendida,
											String strOperacao,
											String strCodEstoqueOrigem,
											String strCodEstoqueDestino,
											String strLojaEstoqueOrigem,
											String strLojaEstoqueDestino,
											String strPedidoEstoqueOrigem,
											String strPedidoEstoqueDestino,
											String strDocumento,
											String strComplemento,
											String strIdOrdemServico,
											out String strMsgErro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "EstoqueDAO.gravaLogEstoque()";
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
					cmInsertLogEstoque.Parameters["@usuario"].Value = ((strUsuario == null) ? "" : strUsuario);
					cmInsertLogEstoque.Parameters["@id_nfe_emitente"].Value = id_nfe_emitente;
					cmInsertLogEstoque.Parameters["@fabricante"].Value = ((strFabricante == null) ? "" : strFabricante);
					cmInsertLogEstoque.Parameters["@produto"].Value = ((strProduto == null) ? "" : strProduto);
					cmInsertLogEstoque.Parameters["@qtde_solicitada"].Value = intQtdeSolicitada;
					cmInsertLogEstoque.Parameters["@qtde_atendida"].Value = intQtdeAtendida;
					cmInsertLogEstoque.Parameters["@operacao"].Value = ((strOperacao == null) ? "" : strOperacao);
					cmInsertLogEstoque.Parameters["@cod_estoque_origem"].Value = ((strCodEstoqueOrigem == null) ? "" : strCodEstoqueOrigem);
					cmInsertLogEstoque.Parameters["@cod_estoque_destino"].Value = ((strCodEstoqueDestino == null) ? "" : strCodEstoqueDestino);
					cmInsertLogEstoque.Parameters["@loja_estoque_origem"].Value = ((strLojaEstoqueOrigem == null) ? "" : strLojaEstoqueOrigem);
					cmInsertLogEstoque.Parameters["@loja_estoque_destino"].Value = ((strLojaEstoqueDestino == null) ? "" : strLojaEstoqueDestino);
					cmInsertLogEstoque.Parameters["@pedido_estoque_origem"].Value = ((strPedidoEstoqueOrigem == null) ? "" : strPedidoEstoqueOrigem);
					cmInsertLogEstoque.Parameters["@pedido_estoque_destino"].Value = ((strPedidoEstoqueDestino == null) ? "" : strPedidoEstoqueDestino);
					cmInsertLogEstoque.Parameters["@documento"].Value = ((strDocumento == null) ? "" : strDocumento);
					cmInsertLogEstoque.Parameters["@complemento"].Value = ((strComplemento == null) ? "" : Texto.leftStr(strComplemento, 80));
					cmInsertLogEstoque.Parameters["@id_ordem_servico"].Value = ((strIdOrdemServico == null) ? "" : strIdOrdemServico);
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsertLogEstoque.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertLogEstoque);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados do registro do log do estoque = " + sbLog.ToString() + "\n" + ex.ToString());
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
					strMsgErro = "Falha ao gravar o log do estoque no BD após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha: Dados do registro do log do estoque = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion
	}

	#region [ cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA ]
	public class cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA
	{
		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private String _fabricante;
		public String fabricante
		{
			get { return _fabricante; }
			set { _fabricante = value; }
		}

		private String _produto;
		public String produto
		{
			get { return _produto; }
			set { _produto = value; }
		}

		#region [ Construtor ]
		public cl_PROCESSA_PRODUTOS_VENDIDOS_SEM_PRESENCA(String pedido, String fabricante, String produto)
		{
			_pedido = pedido;
			_fabricante = fabricante;
			_produto = produto;
		}
		#endregion
	}
	#endregion

	#region [ cl_FABRICANTE_PRODUTO ]
	public class cl_FABRICANTE_PRODUTO
	{
		private String _fabricante;
		public String fabricante
		{
			get { return _fabricante; }
			set { _fabricante = value; }
		}

		private String _produto;
		public String produto
		{
			get { return _produto; }
			set { _produto = value; }
		}

		#region [ Construtor ]
		public cl_FABRICANTE_PRODUTO(String fabricante, String produto)
		{
			_fabricante = fabricante;
			_produto = produto;
		}
		#endregion
	}
	#endregion
}
