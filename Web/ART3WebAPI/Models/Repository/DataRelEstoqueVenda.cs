using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Domains;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace ART3WebAPI.Models.Repository
{
	#region [ DataRelEstoqueVenda ]
	public class DataRelEstoqueVenda
	{
		#region [ Get ]
		public List<RelEstoqueVendaEntity> Get(Guid? httpRequestId, string usuario, string loja, string filtro_estoque, string filtro_detalhe, string filtro_consolidacao_codigos, string filtro_empresa, string filtro_fabricante, string filtro_produto, string filtro_fabricante_multiplo, string filtro_grupo, string filtro_subgrupo, string filtro_potencia_BTU, string filtro_ciclo, string filtro_posicao_mercado)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DataRelEstoqueVenda.Get()";
			string msg;
			#endregion

			if (filtro_consolidacao_codigos.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_CONSOLIDACAO_CODIGOS.NORMAIS))
			{
				return GetConsolidacaoNormal(httpRequestId, usuario, loja, filtro_estoque, filtro_detalhe, filtro_consolidacao_codigos, filtro_empresa, filtro_fabricante, filtro_produto, filtro_fabricante_multiplo, filtro_grupo, filtro_subgrupo, filtro_potencia_BTU, filtro_ciclo, filtro_posicao_mercado);
			}
			else if (filtro_consolidacao_codigos.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_CONSOLIDACAO_CODIGOS.UNIFICADOS))
			{
				return GetConsolidacaoUnificado(httpRequestId, usuario, loja, filtro_estoque, filtro_detalhe, filtro_consolidacao_codigos, filtro_empresa, filtro_fabricante, filtro_produto, filtro_fabricante_multiplo, filtro_grupo, filtro_subgrupo, filtro_potencia_BTU, filtro_ciclo, filtro_posicao_mercado);
			}
			else
			{
				msg = NOME_DESTA_ROTINA + ": Parâmetro 'filtro_consolidacao_codigos' com valor inválido (" + (filtro_consolidacao_codigos ?? "") + ")!";
				Global.gravaLogAtividade(httpRequestId, msg);
				return null;
			}
		}
		#endregion

		#region [ MontaWhereFiltros ]
		private string MontaWhereFiltros(Guid? httpRequestId, string usuario, string loja, string filtro_estoque, string filtro_detalhe, string filtro_consolidacao_codigos, string filtro_empresa, string filtro_fabricante, string filtro_produto, string filtro_fabricante_multiplo, string filtro_grupo, string filtro_subgrupo, string filtro_potencia_BTU, string filtro_ciclo, string filtro_posicao_mercado)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DataRelEstoqueVenda.MontaWhereFiltros()";
			string msg;
			string[] v;
			StringBuilder sbWhere = new StringBuilder("");
			StringBuilder sbAux = new StringBuilder("");
			#endregion

			try
			{
				#region [ Monta restrições com os filtros ]

				#region [ Fabricante ]
				if ((filtro_fabricante ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (t_ESTOQUE_ITEM.fabricante = '" + filtro_fabricante.Trim() + "')");
				}
				#endregion

				#region [ Produto ]
				if ((filtro_produto ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (t_ESTOQUE_ITEM.produto = '" + filtro_produto.Trim() + "')");
				}
				#endregion

				#region [ Empresa ]
				if ((filtro_empresa ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (t_ESTOQUE.id_nfe_emitente = " + filtro_empresa.Trim() + ")");
				}
				#endregion

				#region [ Fabricante múltiplo ]
				if ((filtro_fabricante_multiplo ?? "").Length > 0)
				{
					v = filtro_fabricante_multiplo.Split('_');
					if (sbAux.Length > 0) sbAux.Clear();
					for (int i = 0; i < v.Length; i++)
					{
						if ((v[i] ?? "").Trim().Length > 0)
						{
							if (sbAux.Length > 0) sbAux.Append(",");
							sbAux.Append("'" + v[i] + "'");
						}
					}

					if (sbAux.Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (t_ESTOQUE_ITEM.fabricante IN (" + sbAux.ToString() + "))");
					}
				}
				#endregion

				#region [ Grupo ]
				if ((filtro_grupo ?? "").Length > 0)
				{
					v = filtro_grupo.Split('_');
					if (sbAux.Length > 0) sbAux.Clear();
					for (int i = 0; i < v.Length; i++)
					{
						if ((v[i] ?? "").Trim().Length > 0)
						{
							if (sbAux.Length > 0) sbAux.Append(",");
							sbAux.Append("'" + v[i] + "'");
						}
					}

					if (sbAux.Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (t_PRODUTO.grupo IN (" + sbAux.ToString() + "))");
					}
				}
				#endregion

				#region [ Subgrupo ]
				if ((filtro_subgrupo ?? "").Length > 0)
				{
					v = filtro_subgrupo.Split('_');
					if (sbAux.Length > 0) sbAux.Clear();
					for (int i = 0; i < v.Length; i++)
					{
						if ((v[i] ?? "").Trim().Length > 0)
						{
							if (sbAux.Length > 0) sbAux.Append(",");
							sbAux.Append("'" + v[i] + "'");
						}
					}

					if (sbAux.Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (t_PRODUTO.subgrupo IN (" + sbAux.ToString() + "))");
					}
				}
				#endregion

				#region [ Potência BTU ]
				if ((filtro_potencia_BTU ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (t_PRODUTO.potencia_BTU = '" + filtro_potencia_BTU.Trim() + "')");
				}
				#endregion

				#region [ Ciclo ]
				if ((filtro_ciclo ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (t_PRODUTO.ciclo = '" + filtro_ciclo.Trim() + "')");
				}
				#endregion

				#region [ Posição mercado ]
				if ((filtro_posicao_mercado ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (t_PRODUTO.posicao_mercado = '" + filtro_posicao_mercado.Trim() + "')");
				}
				#endregion

				#endregion

				return sbWhere.ToString();
			}
			catch (Exception ex)
			{
				msg = NOME_DESTA_ROTINA + ": Exception\n" + ex.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);
				throw new Exception(ex.Message);
			}
		}
		#endregion

		#region [ GetConsolidacaoNormal ]
		private List<RelEstoqueVendaEntity> GetConsolidacaoNormal(Guid? httpRequestId, string usuario, string loja, string filtro_estoque, string filtro_detalhe, string filtro_consolidacao_codigos, string filtro_empresa, string filtro_fabricante, string filtro_produto, string filtro_fabricante_multiplo, string filtro_grupo, string filtro_subgrupo, string filtro_potencia_BTU, string filtro_ciclo, string filtro_posicao_mercado)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DataRelEstoqueVenda.GetConsolidacaoNormal()";
			int idxFabricanteCodigo;
			int idxFabricanteNome;
			int idxFabricanteRazaoSocial;
			int idxProdutoCodigo;
			int idxProdutoDescricao;
			int idxProdutoDescricaoHtml;
			int idxCubagem;
			int idxQtde;
			int idxVlTotal = 0;
			string sWhereFiltros;
			StringBuilder sbSql = new StringBuilder("");
			RelEstoqueVendaEntity relFabricante;
			RelEstoqueVendaItemEntity relProduto;
			List<RelEstoqueVendaEntity> resultado = new List<RelEstoqueVendaEntity>();
			SqlConnection cn;
			SqlCommand cmd;
			SqlDataReader reader;
			#endregion

			#region [ Prepara SQL ]
			sWhereFiltros = MontaWhereFiltros(httpRequestId, usuario, loja, filtro_estoque, filtro_detalhe, filtro_consolidacao_codigos, filtro_empresa, filtro_fabricante, filtro_produto, filtro_fabricante_multiplo, filtro_grupo, filtro_subgrupo, filtro_potencia_BTU, filtro_ciclo, filtro_posicao_mercado);

			sbSql.Append("SELECT" +
							" t_ESTOQUE_ITEM.fabricante" +
							", t_FABRICANTE.nome AS fabricante_nome" +
							", t_FABRICANTE.razao_social AS fabricante_razao_social" +
							", t_ESTOQUE_ITEM.produto" +
							", descricao" +
							", descricao_html" +
							", cubagem" +
							", Sum(qtde-qtde_utilizada) AS saldo");

			if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
			{
				sbSql.Append(", Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total");
			}

			sbSql.Append(" FROM t_ESTOQUE_ITEM" +
							" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" +
							" LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" +
							" LEFT JOIN t_FABRICANTE ON (t_PRODUTO.fabricante=t_FABRICANTE.fabricante)" +
						" WHERE" +
							" ((qtde-qtde_utilizada) > 0)");

			if (sWhereFiltros.Length > 0) sbSql.Append(" AND" + sWhereFiltros);

			sbSql.Append(" GROUP BY" +
							" t_ESTOQUE_ITEM.fabricante" +
							", t_FABRICANTE.nome" +
							", t_FABRICANTE.razao_social" +
							", t_ESTOQUE_ITEM.produto" +
							", descricao" +
							", descricao_html" +
							", cubagem" +
						" ORDER BY" +
							" t_ESTOQUE_ITEM.fabricante" +
							", t_FABRICANTE.nome" +
							", t_FABRICANTE.razao_social" +
							", t_ESTOQUE_ITEM.produto" +
							", descricao" +
							", descricao_html" +
							", cubagem");
			#endregion

			#region [ Abre conexão com o BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			#endregion

			try // Finally: BD.fechaConexao(ref cn)
			{
				#region [ Executa consulta ]
				cmd = new SqlCommand();
				cmd.Connection = cn;
				cmd.CommandText = sbSql.ToString();
				reader = cmd.ExecuteReader();
				#endregion

				try // Finally: reader.Close()
				{
					idxFabricanteCodigo = reader.GetOrdinal("fabricante");
					idxFabricanteNome = reader.GetOrdinal("fabricante_nome");
					idxFabricanteRazaoSocial = reader.GetOrdinal("fabricante_razao_social");
					idxProdutoCodigo = reader.GetOrdinal("produto");
					idxProdutoDescricao = reader.GetOrdinal("descricao");
					idxProdutoDescricaoHtml = reader.GetOrdinal("descricao_html");
					idxCubagem = reader.GetOrdinal("cubagem");
					idxQtde = reader.GetOrdinal("saldo");
					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO)) idxVlTotal = reader.GetOrdinal("preco_total");

					while (reader.Read())
					{
						relProduto = new RelEstoqueVendaItemEntity();
						relProduto.ProdutoCodigo = reader.GetString(idxProdutoCodigo);
						relProduto.ProdutoDescricao = reader.GetString(idxProdutoDescricao);
						relProduto.ProdutoDescricaoHtml = reader.GetString(idxProdutoDescricaoHtml);
						relProduto.Cubagem = reader.GetFloat(idxCubagem);
						relProduto.Qtde = reader.GetInt32(idxQtde);
						if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO)) relProduto.VlTotal = reader.GetDecimal(idxVlTotal);

						try
						{
							relFabricante = resultado.Single(p => p.FabricanteCodigo.Equals(reader.GetString(idxFabricanteCodigo)));
							relFabricante.Produtos.Add(relProduto);
						}
						catch (Exception)
						{
							// Não encontrou entrada referente ao fabricante atual, cria nova entrada para ele
							relFabricante = new RelEstoqueVendaEntity();
							relFabricante.FabricanteCodigo = reader.GetString(idxFabricanteCodigo);
							relFabricante.FabricanteDescricao = reader.IsDBNull(idxFabricanteNome) ? "" : reader.GetString(idxFabricanteNome);
							if (relFabricante.FabricanteDescricao.Trim().Length == 0) relFabricante.FabricanteDescricao = reader.IsDBNull(idxFabricanteRazaoSocial) ? "" : reader.GetString(idxFabricanteRazaoSocial);
							relFabricante.Produtos.Add(relProduto);
							resultado.Add(relFabricante);
						}
					}
				}
				finally
				{
					reader.Close();
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + ex.ToString());
				throw new Exception(ex.Message);
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return resultado;
		}
		#endregion

		#region [ GetConsolidacaoUnificado ]
		private List<RelEstoqueVendaEntity> GetConsolidacaoUnificado(Guid? httpRequestId, string usuario, string loja, string filtro_estoque, string filtro_detalhe, string filtro_consolidacao_codigos, string filtro_empresa, string filtro_fabricante, string filtro_produto, string filtro_fabricante_multiplo, string filtro_grupo, string filtro_subgrupo, string filtro_potencia_BTU, string filtro_ciclo, string filtro_posicao_mercado)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "DataRelEstoqueVenda.GetConsolidacaoUnificado()";
			bool bPularProdutoComposto;
			int qtdeEstoqueVendaComposto;
			decimal vlCusto2Composto;
			int qtdeEstoqueVendaAux;
			double cubagemComposto;
			string sql;
			string sWhereFiltros;
			string[] v;
			StringBuilder sbWhere = new StringBuilder("");
			StringBuilder sbAux = new StringBuilder("");
			StringBuilder sbSql = new StringBuilder("");
			StringBuilder sbSqlQtde = new StringBuilder("");
			StringBuilder sbSqlVlCusto = new StringBuilder("");
			RelEstoqueVendaEntity relFabricante;
			RelEstoqueVendaItemEntity relProduto;
			List<RelEstoqueVendaEntity> resultado = new List<RelEstoqueVendaEntity>();
			ProdutoComposto produtoComposto;
			List<ProdutoComposto> listaProdutoComposto = new List<ProdutoComposto>();
			ProdutoCompostoItem produtoItem;
			List<ProdutoCompostoItem> listaProdutoItem = new List<ProdutoCompostoItem>();
			SqlConnection cn;
			SqlCommand cmd;
			SqlDataReader reader;
			#endregion

			#region [ Abre conexão com o BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			#endregion

			try // Finally: BD.fechaConexao(ref cn)
			{
				#region [ Obtém os produtos compostos ]

				#region [ Prepara SQL ]

				#region [ Filtros ]

				#region [ Fabricante ]
				if ((filtro_fabricante ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (tECPC.fabricante_composto = '" + filtro_fabricante.Trim() + "')");
				}
				#endregion

				#region [ Produto ]
				if ((filtro_produto ?? "").Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(" (tECPC.produto_composto = '" + filtro_produto.Trim() + "')");
				}
				#endregion

				#region [ Fabricante múltiplo ]
				if ((filtro_fabricante_multiplo ?? "").Length > 0)
				{
					v = filtro_fabricante_multiplo.Split('_');
					if (sbAux.Length > 0) sbAux.Clear();
					for (int i = 0; i < v.Length; i++)
					{
						if ((v[i] ?? "").Trim().Length > 0)
						{
							if (sbAux.Length > 0) sbAux.Append(",");
							sbAux.Append("'" + v[i] + "'");
						}
					}

					if (sbAux.Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (tECPC.fabricante_composto IN (" + sbAux.ToString() + "))");
					}
				}
				#endregion

				#region [ Grupo ]
				if ((filtro_grupo ?? "").Length > 0)
				{
					v = filtro_grupo.Split('_');
					if (sbAux.Length > 0) sbAux.Clear();
					for (int i = 0; i < v.Length; i++)
					{
						if ((v[i] ?? "").Trim().Length > 0)
						{
							if (sbAux.Length > 0) sbAux.Append(",");
							sbAux.Append("'" + v[i] + "'");
						}
					}

					if (sbAux.Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (tP.grupo IN (" + sbAux.ToString() + "))");
					}
				}
				#endregion

				#region [ Subgrupo ]
				if ((filtro_subgrupo ?? "").Length > 0)
				{
					v = filtro_subgrupo.Split('_');
					if (sbAux.Length > 0) sbAux.Clear();
					for (int i = 0; i < v.Length; i++)
					{
						if ((v[i] ?? "").Trim().Length > 0)
						{
							if (sbAux.Length > 0) sbAux.Append(",");
							sbAux.Append("'" + v[i] + "'");
						}
					}

					if (sbAux.Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (tP.subgrupo IN (" + sbAux.ToString() + "))");
					}
				}
				#endregion

				#endregion

				sbSql.Append("SELECT" +
								" tECPC.fabricante_composto" +
								", tECPC.produto_composto" +
								", tECPC.descricao AS produto_composto_descricao" +
								", tF.nome AS fabricante_composto_nome" +
							" FROM t_EC_PRODUTO_COMPOSTO tECPC" +
								" LEFT JOIN t_FABRICANTE tF ON (tECPC.fabricante_composto = tF.fabricante)" +
								" LEFT JOIN t_PRODUTO tP ON (tECPC.fabricante_composto = tP.fabricante) AND (tECPC.produto_composto = tP.produto)");

				if (sbWhere.Length > 0) sbSql.Append(" WHERE" + sbWhere.ToString());

				sbSql.Append(" ORDER BY" +
								" tECPC.fabricante_composto" +
								", tECPC.produto_composto");
				#endregion

				#region [ Executa consulta ]
				cmd = new SqlCommand();
				cmd.Connection = cn;
				cmd.CommandText = sbSql.ToString();
				reader = cmd.ExecuteReader();
				#endregion

				#region [ Obtém a relação de produtos compostos ]
				try // Finally: reader.Close()
				{
					int idxFabricanteComposto = reader.GetOrdinal("fabricante_composto");
					int idxFabricanteCompostoNome = reader.GetOrdinal("fabricante_composto_nome");
					int idxProdutoComposto = reader.GetOrdinal("produto_composto");
					int idxProdutoCompostoDescricao = reader.GetOrdinal("produto_composto_descricao");

					while (reader.Read())
					{
						produtoComposto = new ProdutoComposto();
						produtoComposto.fabricante_composto = reader.GetString(idxFabricanteComposto);
						produtoComposto.fabricante_composto_nome = reader.IsDBNull(idxFabricanteCompostoNome) ? "" : reader.GetString(idxFabricanteCompostoNome);
						produtoComposto.produto_composto = reader.GetString(idxProdutoComposto);
						produtoComposto.produto_composto_descricao = reader.GetString(idxProdutoCompostoDescricao);
						listaProdutoComposto.Add(produtoComposto);
					}
				}
				finally
				{
					reader.Close();
				}
				#endregion

				#region [ Para cada produto composto, obtém os dados ]
				foreach (ProdutoComposto prodComposto in listaProdutoComposto)
				{
					bPularProdutoComposto = false;
					qtdeEstoqueVendaComposto = -1;
					vlCusto2Composto = 0;
					cubagemComposto = 0;

					#region [ Obtém a composição do produto composto ]

					#region [ Monta SQL ]
					if (sbSql.Length > 0) sbSql.Clear();
					if (sbWhere.Length > 0) sbWhere.Clear();

					#region [ Filtros ]

					#region [ Fabricante ]
					if ((filtro_fabricante ?? "").Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (fabricante_composto = '" + filtro_fabricante.Trim() + "')");
					}
					#endregion

					#region [ Produto ]
					if ((filtro_produto ?? "").Length > 0)
					{
						if (sbWhere.Length > 0) sbWhere.Append(" AND");
						sbWhere.Append(" (produto_composto = '" + filtro_produto.Trim() + "')");
					}
					#endregion

					#region [ Fabricante múltiplo ]
					if ((filtro_fabricante_multiplo ?? "").Length > 0)
					{
						v = filtro_fabricante_multiplo.Split('_');
						if (sbAux.Length > 0) sbAux.Clear();
						for (int i = 0; i < v.Length; i++)
						{
							if ((v[i] ?? "").Trim().Length > 0)
							{
								if (sbAux.Length > 0) sbAux.Append(",");
								sbAux.Append("'" + v[i] + "'");
							}
						}

						if (sbAux.Length > 0)
						{
							if (sbWhere.Length > 0) sbWhere.Append(" AND");
							sbWhere.Append(" (fabricante_composto IN (" + sbAux.ToString() + "))");
						}
					}
					#endregion

					#endregion

					sbSql.Append("SELECT" +
									" fabricante_item" +
									", produto_item" +
									", qtde" +
								" FROM t_EC_PRODUTO_COMPOSTO_ITEM" +
								" WHERE" +
									" (fabricante_composto = '" + prodComposto.fabricante_composto + "')" +
									" AND (produto_composto = '" + prodComposto.produto_composto + "')");

					if (sbWhere.Length > 0) sbSql.Append(" AND" + sbWhere.ToString());

					sbSql.Append(" ORDER BY" +
									" fabricante_item" +
									", produto_item");
					#endregion

					#region [ Executa consulta ]
					cmd.CommandText = sbSql.ToString();
					reader = cmd.ExecuteReader();
					#endregion

					try // Finally: reader.Close()
					{
						#region [ Armazena a composição do produto composto ]
						if (listaProdutoItem.Count > 0) listaProdutoItem.Clear();

						int idxFabricanteItem = reader.GetOrdinal("fabricante_item");
						int idxProdutoItem = reader.GetOrdinal("produto_item");
						int idxProdutoItemQtde = reader.GetOrdinal("qtde");

						while (reader.Read())
						{
							produtoItem = new ProdutoCompostoItem();
							produtoItem.fabricante_composto = prodComposto.fabricante_composto;
							produtoItem.produto_composto = prodComposto.produto_composto;
							produtoItem.fabricante_item = reader.GetString(idxFabricanteItem);
							produtoItem.produto_item = reader.GetString(idxProdutoItem);
							produtoItem.qtde = reader.GetInt16(idxProdutoItemQtde);
							listaProdutoItem.Add(produtoItem);
						}
						#endregion
					}
					finally
					{
						reader.Close();
					}
					#endregion

					#region [ Analisa cada item da composição ]
					foreach (ProdutoCompostoItem prodItem in listaProdutoItem)
					{
						#region [ Monta SQL ]
						if (sbSql.Length > 0) sbSql.Clear();
						if (sbWhere.Length > 0) sbWhere.Clear();

						#region [ Filtros ]

						#region [ Potência BTU ]
						if ((filtro_potencia_BTU ?? "").Length > 0)
						{
							if (sbWhere.Length > 0) sbWhere.Append(" AND");
							sbWhere.Append(" (tP.potencia_BTU = '" + filtro_potencia_BTU.Trim() + "')");
						}
						#endregion

						#region [ Ciclo ]
						if ((filtro_ciclo ?? "").Length > 0)
						{
							if (sbWhere.Length > 0) sbWhere.Append(" AND");
							sbWhere.Append(" (tP.ciclo = '" + filtro_ciclo.Trim() + "')");
						}
						#endregion

						#region [ Posição mercado ]
						if ((filtro_posicao_mercado ?? "").Length > 0)
						{
							if (sbWhere.Length > 0) sbWhere.Append(" AND");
							sbWhere.Append(" (tP.posicao_mercado = '" + filtro_posicao_mercado.Trim() + "')");
						}
						#endregion

						#endregion

						#region [ Select p/ calcular Qtde ]
						if (sbSqlQtde.Length > 0) sbSqlQtde.Clear();
						sbSqlQtde.Append(" Coalesce(" +
											"(SELECT" +
												" Sum(tEI.qtde-tEI.qtde_utilizada)" +
											" FROM t_ESTOQUE_ITEM tEI" +
												" LEFT JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque)" +
											" WHERE" +
												" (tEI.fabricante=tP.fabricante)" +
												" AND (tEI.produto=tP.produto)" +
												" AND ((tEI.qtde-tEI.qtde_utilizada)>0)");
						if ((filtro_empresa ?? "").Length > 0)
						{
							sbSqlQtde.Append(" AND (tE.id_nfe_emitente = " + filtro_empresa.Trim() + ")");
						}
						sbSqlQtde.Append("), 0)");
						#endregion

						#region [ Select p/ calcular vl_custo2 ]
						if (sbSqlVlCusto.Length > 0) sbSqlVlCusto.Clear();
						sbSqlVlCusto.Append(" Coalesce(" +
												"(SELECT" +
													" Sum((tEI.qtde-tEI.qtde_utilizada) * tEI.vl_custo2)" +
												" FROM t_ESTOQUE_ITEM tEI LEFT JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque)" +
												" WHERE" +
													" (tEI.fabricante=tP.fabricante)" +
													" AND (tEI.produto=tP.produto)" +
													" AND ((tEI.qtde-tEI.qtde_utilizada)>0)");
						if ((filtro_empresa ?? "").Length > 0)
						{
							sbSqlVlCusto.Append(" AND (tE.id_nfe_emitente = " + filtro_empresa.Trim() + ")");
						}
						sbSqlVlCusto.Append("), 0)");
						#endregion

						sbSql.Append("SELECT" +
										" tP.fabricante" +
										", tP.produto" +
										", tP.cubagem" +
										", " + sbSqlQtde.ToString() + " AS qtde_estoque_venda" +
										", " + sbSqlVlCusto.ToString() + " AS vl_custo2" +
									" FROM t_PRODUTO tP" +
									" WHERE" +
										" (tP.fabricante = '" + prodItem.fabricante_item + "')" +
										" AND (tP.produto = '" + prodItem.produto_item + "')");

						if (sbWhere.Length > 0) sbSql.Append(" AND" + sbWhere.ToString());

						sql = "SELECT" +
								" *" +
								" FROM (" + sbSql.ToString() + ") t" +
							" WHERE" +
								" (vl_custo2 > 0)" +
							" ORDER BY" +
								" fabricante" +
								", produto";
						#endregion

						#region [ Executa consulta ]
						cmd.CommandText = sql;
						reader = cmd.ExecuteReader();
						#endregion

						try // Finally: reader.Close()
						{
							int idxFabricante = reader.GetOrdinal("fabricante");
							int idxProduto = reader.GetOrdinal("produto");
							int idxCubagem = reader.GetOrdinal("cubagem");
							int idxQtdeEstoqueVenda = reader.GetOrdinal("qtde_estoque_venda");
							int idxVlCusto2 = reader.GetOrdinal("vl_custo2");

							if (reader == null)
							{
								bPularProdutoComposto = true;
							}
							else if (!reader.HasRows)
							{
								bPularProdutoComposto = true;
							}
							else
							{
								while (reader.Read())
								{
									vlCusto2Composto += (prodItem.qtde * (reader.GetDecimal(idxVlCusto2) / reader.GetInt32(idxQtdeEstoqueVenda)));
									qtdeEstoqueVendaAux = reader.GetInt32(idxQtdeEstoqueVenda) / prodItem.qtde;
									if (!reader.IsDBNull(idxCubagem))
									{
										cubagemComposto += reader.GetFloat(idxCubagem) * prodItem.qtde;
									}
									if (qtdeEstoqueVendaComposto == -1)
									{
										qtdeEstoqueVendaComposto = qtdeEstoqueVendaAux;
									}
									else
									{
										if (qtdeEstoqueVendaAux < qtdeEstoqueVendaComposto)
										{
											qtdeEstoqueVendaComposto = qtdeEstoqueVendaAux;
										}
									}
								}
							}
						}
						finally
						{
							reader.Close();
						}

						if (bPularProdutoComposto) break;
					} // foreach (ProdutoCompostoItem prodItem in listaProdutoItem)
					#endregion

					if ((qtdeEstoqueVendaComposto > 0) && (!bPularProdutoComposto))
					{
						relProduto = new RelEstoqueVendaItemEntity();
						relProduto.ProdutoCodigo = prodComposto.produto_composto;
						relProduto.ProdutoDescricao = prodComposto.produto_composto_descricao;
						relProduto.ProdutoDescricaoHtml = prodComposto.produto_composto_descricao;
						relProduto.Cubagem = cubagemComposto;
						relProduto.Qtde = qtdeEstoqueVendaComposto;
						if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO)) relProduto.VlTotal = vlCusto2Composto * qtdeEstoqueVendaComposto;

						try
						{
							relFabricante = resultado.Single(p => p.FabricanteCodigo.Equals(prodComposto.fabricante_composto.Trim()));
							relFabricante.Produtos.Add(relProduto);
						}
						catch (Exception)
						{
							// Não encontrou entrada referente ao fabricante atual, cria nova entrada para ele
							relFabricante = new RelEstoqueVendaEntity();
							relFabricante.FabricanteCodigo = prodComposto.fabricante_composto;
							relFabricante.FabricanteDescricao = prodComposto.fabricante_composto_nome;
							relFabricante.Produtos.Add(relProduto);
							resultado.Add(relFabricante);
						}
					}
				} // foreach (ProdutoComposto prodComposto in listaProdutoComposto)
				#endregion

				#endregion

				#region [ Obtém os produtos normais ]

				#region [ Monta SQL ]

				if (sbSql.Length > 0) sbSql.Clear();
				if (sbWhere.Length > 0) sbWhere.Clear();

				sWhereFiltros = MontaWhereFiltros(httpRequestId, usuario, loja, filtro_estoque, filtro_detalhe, filtro_consolidacao_codigos, filtro_empresa, filtro_fabricante, filtro_produto, filtro_fabricante_multiplo, filtro_grupo, filtro_subgrupo, filtro_potencia_BTU, filtro_ciclo, filtro_posicao_mercado);

				sbSql.Append("SELECT" +
								" t_ESTOQUE_ITEM.fabricante" +
								", t_FABRICANTE.nome AS fabricante_nome" +
								", t_FABRICANTE.razao_social AS fabricante_razao_social" +
								", t_ESTOQUE_ITEM.produto" +
								", descricao" +
								", descricao_html" +
								", cubagem" +
								", Sum(qtde-qtde_utilizada) AS saldo");

				if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
				{
					sbSql.Append(", Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total");
				}

				sbSql.Append(" FROM t_ESTOQUE_ITEM" +
								" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" +
								" LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" +
								" LEFT JOIN t_FABRICANTE ON (t_PRODUTO.fabricante=t_FABRICANTE.fabricante)" +
							" WHERE" +
								" ((qtde-qtde_utilizada) > 0)" +
								" AND (t_ESTOQUE_ITEM.produto NOT IN" +
									" (" +
									"SELECT produto_item FROM t_EC_PRODUTO_COMPOSTO_ITEM WHERE (fabricante_item=t_ESTOQUE_ITEM.fabricante)" +
									")" +
								")");

				if (sWhereFiltros.Length > 0) sbSql.Append(" AND" + sWhereFiltros);

				sbSql.Append(" GROUP BY" +
								" t_ESTOQUE_ITEM.fabricante" +
								", t_FABRICANTE.nome" +
								", t_FABRICANTE.razao_social" +
								", t_ESTOQUE_ITEM.produto" +
								", descricao" +
								", descricao_html" +
								", cubagem" +
							" ORDER BY" +
								" t_ESTOQUE_ITEM.fabricante" +
								", t_FABRICANTE.nome" +
								", t_FABRICANTE.razao_social" +
								", t_ESTOQUE_ITEM.produto" +
								", descricao" +
								", descricao_html" +
								", cubagem");
				#endregion

				#region [ Executa consulta ]
				cmd.CommandText = sbSql.ToString();
				reader = cmd.ExecuteReader();
				#endregion

				try // Finally: reader.Close()
				{
					int idxFabricanteCodigo = reader.GetOrdinal("fabricante");
					int idxFabricanteNome = reader.GetOrdinal("fabricante_nome");
					int idxFabricanteRazaoSocial = reader.GetOrdinal("fabricante_razao_social");
					int idxProdutoCodigo = reader.GetOrdinal("produto");
					int idxProdutoDescricao = reader.GetOrdinal("descricao");
					int idxProdutoDescricaoHtml = reader.GetOrdinal("descricao_html");
					int idxQtde = reader.GetOrdinal("saldo");
					int idxCubagem = reader.GetOrdinal("cubagem");
					int idxVlTotal = 0;
					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO)) idxVlTotal = reader.GetOrdinal("preco_total");

					while (reader.Read())
					{
						relProduto = new RelEstoqueVendaItemEntity();
						relProduto.ProdutoCodigo = reader.GetString(idxProdutoCodigo);
						relProduto.ProdutoDescricao = reader.GetString(idxProdutoDescricao);
						relProduto.ProdutoDescricaoHtml = reader.GetString(idxProdutoDescricaoHtml);
						relProduto.Cubagem = reader.GetFloat(idxCubagem);
						relProduto.Qtde = reader.GetInt32(idxQtde);
						if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO)) relProduto.VlTotal = reader.GetDecimal(idxVlTotal);

						try
						{
							relFabricante = resultado.Single(p => p.FabricanteCodigo.Equals(reader.GetString(idxFabricanteCodigo)));
							relFabricante.Produtos.Add(relProduto);
						}
						catch (Exception)
						{
							// Não encontrou entrada referente ao fabricante atual, cria nova entrada para ele
							relFabricante = new RelEstoqueVendaEntity();
							relFabricante.FabricanteCodigo = reader.GetString(idxFabricanteCodigo);
							relFabricante.FabricanteDescricao = reader.IsDBNull(idxFabricanteNome) ? "" : reader.GetString(idxFabricanteNome);
							if (relFabricante.FabricanteDescricao.Trim().Length == 0) relFabricante.FabricanteDescricao = reader.IsDBNull(idxFabricanteRazaoSocial) ? "" : reader.GetString(idxFabricanteRazaoSocial);
							relFabricante.Produtos.Add(relProduto);
							resultado.Add(relFabricante);
						}
					}
				}
				finally
				{
					reader.Close();
				}
				#endregion

				#region [ Ordena resultado por fabricante e produto ]
				resultado.Sort((x, y) => x.FabricanteCodigo.CompareTo(y.FabricanteCodigo));

				foreach (RelEstoqueVendaEntity relFabr in resultado)
				{
					relFabr.Produtos.Sort((x, y) => x.ProdutoCodigo.CompareTo(y.ProdutoCodigo));
				}
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + ": " + ex.ToString());
				throw new Exception(ex.Message);
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return resultado;
		}
		#endregion
	}
	#endregion

	#region [ ProdutoComposto ]
	class ProdutoComposto
	{
		public string fabricante_composto { get; set; }
		public string fabricante_composto_nome { get; set; }
		public string produto_composto { get; set; }
		public string produto_composto_descricao { get; set; }
	}
	#endregion

	#region [ ProdutoCompostoItem ]
	class ProdutoCompostoItem
	{
		public string fabricante_composto { get; set; }
		public string produto_composto { get; set; }
		public string fabricante_item { get; set; }
		public string produto_item { get; set; }
		public int qtde { get; set; }
	}
	#endregion
}