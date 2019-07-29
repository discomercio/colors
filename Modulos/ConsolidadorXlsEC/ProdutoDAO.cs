using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	public class ProdutoDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
		private SqlCommand cmSelectEcProdutoComposto;
		private SqlCommand cmSelectEcProdutoCompostoItem;
		private SqlCommand cmSelectProduto;
		private SqlCommand cmSelectEstoqueVenda;
		private SqlCommand cmSelectProdutoLoja;
		private SqlCommand cmSelectTabelaPrecoLoja;
		private SqlCommand cmSelectPercentualCustoFinanceiroFornecedor;
		private SqlCommand cmSelectTabelaPercentualCustoFinanceiroFornecedor;
		private SqlCommand cmUpdateProdutoLoja;
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
		public ProdutoDAO(ref BancoDados bd)
		{
			_bd = bd;
			inicializaObjetos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetos ]
		public void inicializaObjetos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmSelectEcProdutoComposto ]
			strSql = "SELECT " +
						"*" +
					" FROM t_EC_PRODUTO_COMPOSTO" +
					" WHERE" +
						" (produto_composto = @produto_composto)";
			cmSelectEcProdutoComposto = _bd.criaSqlCommand();
			cmSelectEcProdutoComposto.CommandText = strSql;
			cmSelectEcProdutoComposto.Parameters.Add("@produto_composto", SqlDbType.VarChar, 8);
			cmSelectEcProdutoComposto.Prepare();
			#endregion

			#region [ cmSelectEcProdutoCompostoItem ]
			strSql = "SELECT " +
						"*" +
					" FROM t_EC_PRODUTO_COMPOSTO_ITEM" +
					" WHERE" +
						" (fabricante_composto = @fabricante_composto)" +
						" AND (produto_composto = @produto_composto)" +
						" AND (excluido_status = 0)" +
					" ORDER BY" +
						" sequencia";
			cmSelectEcProdutoCompostoItem = _bd.criaSqlCommand();
			cmSelectEcProdutoCompostoItem.CommandText = strSql;
			cmSelectEcProdutoCompostoItem.Parameters.Add("@fabricante_composto", SqlDbType.VarChar, 4);
			cmSelectEcProdutoCompostoItem.Parameters.Add("@produto_composto", SqlDbType.VarChar, 8);
			cmSelectEcProdutoCompostoItem.Prepare();
			#endregion

			#region [ cmSelectProduto ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PRODUTO" +
					" WHERE" +
						" (produto = @produto)";
			cmSelectProduto = _bd.criaSqlCommand();
			cmSelectProduto.CommandText = strSql;
			cmSelectProduto.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmSelectProduto.Prepare();
			#endregion

			#region [ cmSelectEstoqueVenda ]
			strSql = "SELECT" +
						" tP.fabricante," +
						" tP.produto," +
						" Coalesce(Sum(qtde - qtde_utilizada), 0) AS saldo," +
						" Coalesce(Sum((qtde - qtde_utilizada) * tEI.vl_custo2), 0) AS vl_total_custo2" +
					" FROM t_produto tP" +
						" LEFT JOIN t_ESTOQUE_ITEM tEI ON ((tEI.fabricante = tP.fabricante) AND (tEI.produto = tP.produto))" +
					" WHERE" +
						" (tP.fabricante = @fabricante)" +
						" AND (tP.produto = @produto)" +
						" AND ((qtde - qtde_utilizada) > 0)" +
					" GROUP BY" +
						" tP.fabricante," +
						" tP.produto";
			cmSelectEstoqueVenda = _bd.criaSqlCommand();
			cmSelectEstoqueVenda.CommandText = strSql;
			cmSelectEstoqueVenda.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmSelectEstoqueVenda.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmSelectEstoqueVenda.Prepare();
			#endregion

			#region [ cmSelectProdutoLoja ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PRODUTO_LOJA" +
					" WHERE" +
						" (fabricante = @fabricante)" +
						" AND (produto = @produto)" +
						" AND (loja = @loja)";
			cmSelectProdutoLoja = _bd.criaSqlCommand();
			cmSelectProdutoLoja.CommandText = strSql;
			cmSelectProdutoLoja.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmSelectProdutoLoja.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmSelectProdutoLoja.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmSelectProdutoLoja.Prepare();
			#endregion

			#region [ cmSelectTabelaPrecoLoja ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PRODUTO_LOJA" +
					" WHERE" +
						" (loja = @loja)" +
					" ORDER BY" +
						" loja," +
						" fabricante," +
						" produto";
			cmSelectTabelaPrecoLoja = _bd.criaSqlCommand();
			cmSelectTabelaPrecoLoja.CommandText = strSql;
			cmSelectTabelaPrecoLoja.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmSelectTabelaPrecoLoja.Prepare();
			#endregion

			#region [ cmSelectPercentualCustoFinanceiroFornecedor ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" +
					" WHERE" +
						" (fabricante = @fabricante)" +
					" ORDER BY" +
						" fabricante," +
						" tipo_parcelamento," +
						" qtde_parcelas";
			cmSelectPercentualCustoFinanceiroFornecedor = _bd.criaSqlCommand();
			cmSelectPercentualCustoFinanceiroFornecedor.CommandText = strSql;
			cmSelectPercentualCustoFinanceiroFornecedor.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmSelectPercentualCustoFinanceiroFornecedor.Prepare();
			#endregion

			#region [ cmSelectTabelaPercentualCustoFinanceiroFornecedor ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" +
					" ORDER BY" +
						" fabricante," +
						" tipo_parcelamento," +
						" qtde_parcelas";
			cmSelectTabelaPercentualCustoFinanceiroFornecedor = _bd.criaSqlCommand();
			cmSelectTabelaPercentualCustoFinanceiroFornecedor.CommandText = strSql;
			cmSelectTabelaPercentualCustoFinanceiroFornecedor.Prepare();
			#endregion

			#region [ cmUpdateProdutoLoja ]
			strSql = "UPDATE t_PRODUTO_LOJA" +
					" SET" +
						" preco_lista = @preco_lista," +
						" dt_ult_atualizacao = getdate()" +
					" WHERE" +
						" (fabricante = @fabricante)" +
						" AND (produto = @produto)" +
						" AND (loja = @loja)";
			cmUpdateProdutoLoja = _bd.criaSqlCommand();
			cmUpdateProdutoLoja.CommandText = strSql;
			cmUpdateProdutoLoja.Parameters.Add("@fabricante", SqlDbType.VarChar, 4);
			cmUpdateProdutoLoja.Parameters.Add("@produto", SqlDbType.VarChar, 8);
			cmUpdateProdutoLoja.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmUpdateProdutoLoja.Parameters.Add("@preco_lista", SqlDbType.Money);
			cmUpdateProdutoLoja.Prepare();
			#endregion
		}
		#endregion

		#region [ GetProdutoUnificadoCustoIntermediario ]
		public ProdutoEstoqueVenda GetProdutoUnificadoCustoIntermediario(string codigoProduto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetProdutoUnificadoCustoIntermediario()";
			int qtdeEstoque;
			int qtdeEstoqueAux;
			int qtdeSaldo;
			decimal valorComposto;
			decimal valor;
			List<ProdutoCompostoItem> vProdutoCompostoItem = new List<ProdutoCompostoItem>();
			ProdutoCompostoItem produtoCompostoItem;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			ProdutoEstoqueVenda produto = new ProdutoEstoqueVenda();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(codigoProduto))
				{
					msg_erro = "Código do produto não informado";
					return null;
				}

				codigoProduto = codigoProduto.Trim();
				codigoProduto = Global.normalizaCodigoProduto(codigoProduto);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Verifica se é um produto composto ]

				#region [ Executa a consulta ]
				cmSelectEcProdutoComposto.Parameters["@produto_composto"].Value = codigoProduto;
				daDataAdapter.SelectCommand = cmSelectEcProdutoComposto;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count > 0)
				{
					produto.isCadastrado = true;
					produto.isComposto = true;
					rowResultado = dtbResultado.Rows[0];
					produto.fabricante = BD.readToString(rowResultado["fabricante_composto"]);
					produto.produto = BD.readToString(rowResultado["produto_composto"]);
					produto.descricao = BD.readToString(rowResultado["descricao"]);
				}
				#endregion

				#region [ Se não é produto composto, tenta localizar no cadastro básico de produtos ]
				if (!produto.isComposto)
				{
					dtbResultado.Reset();
					cmSelectProduto.Parameters["@produto"].Value = codigoProduto;
					daDataAdapter.SelectCommand = cmSelectProduto;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					if (dtbResultado.Rows.Count > 0)
					{
						produto.isCadastrado = true;
						rowResultado = dtbResultado.Rows[0];
						produto.fabricante = BD.readToString(rowResultado["fabricante"]);
						produto.produto = BD.readToString(rowResultado["produto"]);
						produto.descricao = BD.readToString(rowResultado["descricao"]);
					}
				}
				#endregion

				#region [ Produto não cadastrado (t_EC_PRODUTO_COMPOSTO ou t_PRODUTO) ]
				if (!produto.isCadastrado) return produto;
				#endregion

				#region [ Produto normal ]
				if (!produto.isComposto)
				{
					dtbResultado.Reset();
					cmSelectEstoqueVenda.Parameters["@fabricante"].Value = produto.fabricante;
					cmSelectEstoqueVenda.Parameters["@produto"].Value = produto.produto;
					daDataAdapter.SelectCommand = cmSelectEstoqueVenda;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					if (dtbResultado.Rows.Count > 0)
					{
						rowResultado = dtbResultado.Rows[0];
						produto.qtdeEstoqueVenda = BD.readToInt(rowResultado["saldo"]);
						if (produto.qtdeEstoqueVenda > 0)
						{
							// Acerta arredondamento dos centavos
							valor = BD.readToDecimal(rowResultado["vl_total_custo2"]) / produto.qtdeEstoqueVenda;
							produto.vlCustoIntermediario = Global.arredondaParaMonetario(valor);
						}
					}

					return produto;
				}
				#endregion

				#region [ Produto Composto ]
				if (produto.isComposto)
				{
					qtdeEstoque = 0;
					qtdeEstoqueAux = 0;
					valorComposto = 0m;

					dtbResultado.Reset();
					cmSelectEcProdutoCompostoItem.Parameters["@fabricante_composto"].Value = produto.fabricante;
					cmSelectEcProdutoCompostoItem.Parameters["@produto_composto"].Value = produto.produto;
					daDataAdapter.SelectCommand = cmSelectEcProdutoCompostoItem;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					if (dtbResultado.Rows.Count > 0)
					{
						for (int i = 0; i < dtbResultado.Rows.Count; i++)
						{
							rowResultado = dtbResultado.Rows[i];
							produtoCompostoItem = new ProdutoCompostoItem();
							produtoCompostoItem.fabricante_composto = BD.readToString(rowResultado["fabricante_composto"]);
							produtoCompostoItem.produto_composto = BD.readToString(rowResultado["produto_composto"]);
							produtoCompostoItem.fabricante_item = BD.readToString(rowResultado["fabricante_item"]);
							produtoCompostoItem.produto_item = BD.readToString(rowResultado["produto_item"]);
							produtoCompostoItem.qtde = BD.readToInt(rowResultado["qtde"]);
							vProdutoCompostoItem.Add(produtoCompostoItem);
						}

						foreach (var item in vProdutoCompostoItem)
						{
							dtbResultado.Reset();
							cmSelectEstoqueVenda.Parameters["@fabricante"].Value = item.fabricante_item;
							cmSelectEstoqueVenda.Parameters["@produto"].Value = item.produto_item;
							daDataAdapter.SelectCommand = cmSelectEstoqueVenda;
							daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
							daDataAdapter.Fill(dtbResultado);
							if (dtbResultado.Rows.Count == 0)
							{
								#region [ Se não encontrou um produto da composição, então o estoque disponível é zero ]
								qtdeEstoque = 0;
								valorComposto = 0m;
								break;
								#endregion
							}
							else
							{
								rowResultado = dtbResultado.Rows[0];
								qtdeSaldo = BD.readToInt(rowResultado["saldo"]);
								qtdeEstoqueAux = qtdeSaldo / item.qtde;

								#region [ Se o estoque de um produto da composição está zerado, então o estoque disponível do produto composto é zero ]
								if ((qtdeSaldo == 0) || (qtdeEstoqueAux == 0))
								{
									qtdeEstoque = 0;
									valorComposto = 0m;
									break;
								}
								#endregion

								if (qtdeSaldo > 0)
								{
									valorComposto += (item.qtde * BD.readToDecimal(rowResultado["vl_total_custo2"])) / qtdeSaldo;
								}

								if (qtdeEstoque == 0)
								{
									qtdeEstoque = qtdeEstoqueAux;
								}
								else
								{
									if (qtdeEstoqueAux < qtdeEstoque) qtdeEstoque = qtdeEstoqueAux;
								}
							}
						}
					}

					produto.vlCustoIntermediario = Global.arredondaParaMonetario(valorComposto);
					produto.qtdeEstoqueVenda = qtdeEstoque;

					return produto;
				}
				#endregion

				return null;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ GetProdutoLoja ]
		public ProdutoLoja GetProdutoLoja(string codigoFabricante, string codigoProduto, string loja, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetProdutoLoja()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			ProdutoLoja produtoLoja = new ProdutoLoja();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(codigoFabricante))
				{
					msg_erro = "Código do fabricante não informado";
					return null;
				}

				if (string.IsNullOrEmpty(codigoProduto))
				{
					msg_erro = "Código do produto não informado";
					return null;
				}

				if (string.IsNullOrEmpty(loja))
				{
					msg_erro = "Número da loja não informado";
					return null;
				}

				codigoFabricante = codigoFabricante.Trim();
				codigoFabricante = Global.normalizaCodigoFabricante(codigoFabricante);

				codigoProduto = codigoProduto.Trim();
				codigoProduto = Global.normalizaCodigoProduto(codigoProduto);

				loja = loja.Trim();
				loja = Global.normalizaNumeroLoja(loja);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Executa a consulta ]
				cmSelectProdutoLoja.Parameters["@fabricante"].Value = codigoFabricante;
				cmSelectProdutoLoja.Parameters["@produto"].Value = codigoProduto;
				cmSelectProdutoLoja.Parameters["@loja"].Value = loja;
				daDataAdapter.SelectCommand = cmSelectProdutoLoja;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

				rowResultado = dtbResultado.Rows[0];

				produtoLoja.fabricante = BD.readToString(rowResultado["fabricante"]);
				produtoLoja.produto = BD.readToString(rowResultado["produto"]);
				produtoLoja.loja = BD.readToString(rowResultado["loja"]);
				produtoLoja.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
				produtoLoja.margem = BD.readToSingle(rowResultado["margem"]);
				produtoLoja.desc_max = BD.readToSingle(rowResultado["desc_max"]);
				produtoLoja.comissao = BD.readToSingle(rowResultado["comissao"]);
				produtoLoja.vendavel = BD.readToString(rowResultado["vendavel"]);
				produtoLoja.qtde_max_venda = BD.readToInt(rowResultado["qtde_max_venda"]);
				produtoLoja.cor = BD.readToString(rowResultado["cor"]);
				produtoLoja.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
				produtoLoja.dt_ult_atualizacao = BD.readToDateTime(rowResultado["dt_ult_atualizacao"]);
				produtoLoja.excluido_status = BD.readToInt(rowResultado["excluido_status"]);

				return produtoLoja;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ GetTabelaPrecoLoja ]
		public List<ProdutoLoja> GetTabelaPrecoLoja(string loja, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetTabelaPrecoLoja()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			ProdutoLoja produtoLoja;
			List<ProdutoLoja> listaProdutoLoja = new List<ProdutoLoja>();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(loja))
				{
					msg_erro = "Número da loja não informado";
					return null;
				}

				loja = loja.Trim();
				loja = Global.normalizaNumeroLoja(loja);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Executa a consulta ]
				cmSelectTabelaPrecoLoja.Parameters["@loja"].Value = loja;
				daDataAdapter.SelectCommand = cmSelectTabelaPrecoLoja;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					produtoLoja = new ProdutoLoja();
					produtoLoja.fabricante = BD.readToString(rowResultado["fabricante"]);
					produtoLoja.produto = BD.readToString(rowResultado["produto"]);
					produtoLoja.loja = BD.readToString(rowResultado["loja"]);
					produtoLoja.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
					produtoLoja.margem = BD.readToSingle(rowResultado["margem"]);
					produtoLoja.desc_max = BD.readToSingle(rowResultado["desc_max"]);
					produtoLoja.comissao = BD.readToSingle(rowResultado["comissao"]);
					produtoLoja.vendavel = BD.readToString(rowResultado["vendavel"]);
					produtoLoja.qtde_max_venda = BD.readToInt(rowResultado["qtde_max_venda"]);
					produtoLoja.cor = BD.readToString(rowResultado["cor"]);
					produtoLoja.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
					produtoLoja.dt_ult_atualizacao = BD.readToDateTime(rowResultado["dt_ult_atualizacao"]);
					produtoLoja.excluido_status = BD.readToInt(rowResultado["excluido_status"]);
					listaProdutoLoja.Add(produtoLoja);
				}

				return listaProdutoLoja;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ GetProdutoCompostoItem ]
		public List<ProdutoCompostoItem> GetProdutoCompostoItem(string fabricante_composto, string produto_composto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetProdutoCompostoItem()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			ProdutoCompostoItem produtoCompostoItem;
			List<ProdutoCompostoItem> vProdutoCompostoItem = new List<ProdutoCompostoItem>();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(fabricante_composto))
				{
					msg_erro = "Código do fabricante não informado";
					return null;
				}

				if (string.IsNullOrEmpty(produto_composto))
				{
					msg_erro = "Código do produto não informado";
					return null;
				}

				fabricante_composto = fabricante_composto.Trim();
				fabricante_composto = Global.normalizaCodigoFabricante(fabricante_composto);

				produto_composto = produto_composto.Trim();
				produto_composto = Global.normalizaCodigoProduto(produto_composto);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Executa a consulta ]
				cmSelectEcProdutoCompostoItem.Parameters["@fabricante_composto"].Value = fabricante_composto;
				cmSelectEcProdutoCompostoItem.Parameters["@produto_composto"].Value = produto_composto;
				daDataAdapter.SelectCommand = cmSelectEcProdutoCompostoItem;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);

				if (dtbResultado.Rows.Count == 0) return null;

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];

					produtoCompostoItem = new ProdutoCompostoItem();
					produtoCompostoItem.fabricante_composto = BD.readToString(rowResultado["fabricante_composto"]);
					produtoCompostoItem.produto_composto = BD.readToString(rowResultado["produto_composto"]);
					produtoCompostoItem.fabricante_item = BD.readToString(rowResultado["fabricante_item"]);
					produtoCompostoItem.produto_item = BD.readToString(rowResultado["produto_item"]);
					produtoCompostoItem.qtde = BD.readToInt(rowResultado["qtde"]);
					vProdutoCompostoItem.Add(produtoCompostoItem);
				}
				#endregion

				return vProdutoCompostoItem;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ GetProdutoCadastroBasico ]
		public ProdutoCadastroBasico GetProdutoCadastroBasico(string codigoProduto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetProdutoCadastroBasico()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			ProdutoCadastroBasico produto = new ProdutoCadastroBasico();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(codigoProduto))
				{
					msg_erro = "Código do produto não informado";
					return null;
				}

				codigoProduto = codigoProduto.Trim();
				codigoProduto = Global.normalizaCodigoProduto(codigoProduto);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Verifica se é um produto composto ]

				#region [ Executa a consulta ]
				cmSelectEcProdutoComposto.Parameters["@produto_composto"].Value = codigoProduto;
				daDataAdapter.SelectCommand = cmSelectEcProdutoComposto;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count > 0)
				{
					produto.isCadastrado = true;
					produto.isComposto = true;
					rowResultado = dtbResultado.Rows[0];
					produto.fabricante = BD.readToString(rowResultado["fabricante_composto"]);
					produto.produto = BD.readToString(rowResultado["produto_composto"]);
					produto.descricao = BD.readToString(rowResultado["descricao"]);
				}
				#endregion

				#region [ Se não é produto composto, tenta localizar no cadastro básico de produtos ]
				if (!produto.isComposto)
				{
					dtbResultado.Reset();
					cmSelectProduto.Parameters["@produto"].Value = codigoProduto;
					daDataAdapter.SelectCommand = cmSelectProduto;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					if (dtbResultado.Rows.Count > 0)
					{
						produto.isCadastrado = true;
						rowResultado = dtbResultado.Rows[0];
						produto.fabricante = BD.readToString(rowResultado["fabricante"]);
						produto.produto = BD.readToString(rowResultado["produto"]);
						produto.descricao = BD.readToString(rowResultado["descricao"]);
					}
				}
				#endregion

				return produto;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ GetPercentualCustoFinanceiroFornecedor ]
		public List<PercentualCustoFinanceiroFornecedor> GetPercentualCustoFinanceiroFornecedor(string codigoFabricante, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetPercentualCustoFinanceiroFornecedor()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			PercentualCustoFinanceiroFornecedor percentualCustoFinanceiroFornecedor;
			List<PercentualCustoFinanceiroFornecedor> listaPercentualCustoFinanceiroFornecedor = new List<PercentualCustoFinanceiroFornecedor>();
			#endregion

			msg_erro = "";
			try
			{
				if (string.IsNullOrEmpty(codigoFabricante))
				{
					msg_erro = "Código do fabricante não informado";
					return null;
				}

				codigoFabricante = codigoFabricante.Trim();
				codigoFabricante = Global.normalizaCodigoFabricante(codigoFabricante);

				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Executa a consulta ]
				cmSelectPercentualCustoFinanceiroFornecedor.Parameters["@fabricante"].Value = codigoFabricante;
				daDataAdapter.SelectCommand = cmSelectPercentualCustoFinanceiroFornecedor;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					percentualCustoFinanceiroFornecedor = new PercentualCustoFinanceiroFornecedor();
					percentualCustoFinanceiroFornecedor.fabricante = BD.readToString(rowResultado["fabricante"]);
					percentualCustoFinanceiroFornecedor.tipo_parcelamento = BD.readToString(rowResultado["tipo_parcelamento"]);
					percentualCustoFinanceiroFornecedor.qtde_parcelas = BD.readToInt(rowResultado["qtde_parcelas"]);
					percentualCustoFinanceiroFornecedor.coeficiente = BD.readToSingle(rowResultado["coeficiente"]);
					listaPercentualCustoFinanceiroFornecedor.Add(percentualCustoFinanceiroFornecedor);
				}

				return listaPercentualCustoFinanceiroFornecedor;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ GetTabelaPercentualCustoFinanceiroFornecedor ]
		public List<PercentualCustoFinanceiroFornecedor> GetTabelaPercentualCustoFinanceiroFornecedor(out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.GetTabelaPercentualCustoFinanceiroFornecedor()";
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			PercentualCustoFinanceiroFornecedor percentualCustoFinanceiroFornecedor;
			List<PercentualCustoFinanceiroFornecedor> listaPercentualCustoFinanceiroFornecedor = new List<PercentualCustoFinanceiroFornecedor>();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Executa a consulta ]
				daDataAdapter.SelectCommand = cmSelectTabelaPercentualCustoFinanceiroFornecedor;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					percentualCustoFinanceiroFornecedor = new PercentualCustoFinanceiroFornecedor();
					percentualCustoFinanceiroFornecedor.fabricante = BD.readToString(rowResultado["fabricante"]);
					percentualCustoFinanceiroFornecedor.tipo_parcelamento = BD.readToString(rowResultado["tipo_parcelamento"]);
					percentualCustoFinanceiroFornecedor.qtde_parcelas = BD.readToInt(rowResultado["qtde_parcelas"]);
					percentualCustoFinanceiroFornecedor.coeficiente = BD.readToSingle(rowResultado["coeficiente"]);
					listaPercentualCustoFinanceiroFornecedor.Add(percentualCustoFinanceiroFornecedor);
				}

				return listaPercentualCustoFinanceiroFornecedor;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ UpdateProdutoLoja ]
		public bool UpdateProdutoLoja(ProdutoLoja produtoLoja, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ProdutoDAO.UpdateProdutoLoja()";
			string strMsg;
			int intRetorno;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if (produtoLoja == null) return false;
				if ((produtoLoja.fabricante ?? "").Length == 0) return false;
				if ((produtoLoja.produto ?? "").Length == 0) return false;
				if ((produtoLoja.loja ?? "").Length == 0) return false;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateProdutoLoja.Parameters["@fabricante"].Value = produtoLoja.fabricante;
				cmUpdateProdutoLoja.Parameters["@produto"].Value = produtoLoja.produto;
				cmUpdateProdutoLoja.Parameters["@loja"].Value = produtoLoja.loja;
				cmUpdateProdutoLoja.Parameters["@preco_lista"].Value = produtoLoja.preco_lista;
				#endregion

				#region [ Tenta atualizar o registro ]
				try
				{
					intRetorno = _bd.executaNonQuery(ref cmUpdateProdutoLoja);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;
					strMsg = NOME_DESTA_ROTINA + ": exception ao tentar atualizar o preço de lista do produto (" + produtoLoja.fabricante + ")" + produtoLoja.produto + " na tabela de preços da loja " + produtoLoja.loja + " (" + Global.formataMoeda(produtoLoja.preco_lista) + ")!!\r\n" + ex.ToString();
					Global.gravaLogAtividade(strMsg);
					return false;
				}
				#endregion

				if (intRetorno == 0)
				{
					strMsg = NOME_DESTA_ROTINA + ": falha ao tentar atualizar o preço de lista do produto (" + produtoLoja.fabricante + ")" + produtoLoja.produto + " na tabela de preços da loja " + produtoLoja.loja + " (" + Global.formataMoeda(produtoLoja.preco_lista) + ")!!";
					return false;
				}

				strMsg = NOME_DESTA_ROTINA + ": sucesso na atualização do preço de lista do produto (" + produtoLoja.fabricante + ")" + produtoLoja.produto + " na tabela de preços da loja " + produtoLoja.loja + " (" + Global.formataMoeda(produtoLoja.preco_lista) + ")!!";
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return false;
			}
		}
		#endregion

		#endregion
	}
}
