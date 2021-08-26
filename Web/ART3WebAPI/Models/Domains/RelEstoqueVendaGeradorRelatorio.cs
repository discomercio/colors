using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using System;
using System.Globalization;
using System.Threading;
using System.Text;

namespace ART3WebAPI.Models.Domains
{
	public class RelEstoqueVendaGeradorRelatorio
	{
		public static Task GeraXLS(List<RelEstoqueVendaEntity> dataRel, string filePath, string usuario, string loja, string filtro_estoque, string filtro_detalhe, string filtro_consolidacao_codigos, string filtro_empresa, string filtro_fabricante, string filtro_produto, string filtro_fabricante_multiplo, string filtro_grupo, string filtro_subgrupo, string filtro_potencia_BTU, string filtro_ciclo, string filtro_posicao_mercado)
		{
			return Task.Run(() =>
			{
				#region [ Declarações ]
				const int COL_INICIAL = 2;
				const int ROW_INICIAL = 2;
				int COL_FINAL;
				int ROW_CABECALHO, ROW_INICIO_DADOS;
				int colFabricanteCodigo = 0, colProdutoCodigo = 0, colProdutoDescricao = 0, colCubagem = 0, colQtde = 0, colVlCustoMedio = 0, colVlCustoTotal = 0;
				int rowTotal, rowUltLinhaDados;
				int nLinha = 0;
				int NumRegistros = 0;
				string cellsIndex;
				string descricao_filtro_detalhe;
				string descricao_filtro_consolidacao_codigos;
				string descricao_filtro_empresa;
				StringBuilder sbAux = new StringBuilder("");
				NFeEmitente emitente;
				#endregion

				#region [ Calcula total de registros ]
				foreach (RelEstoqueVendaEntity relFabricante in dataRel)
				{
					NumRegistros += relFabricante.Produtos.Count;
				}
				#endregion

				#region [ Preparação de campos que descrevem os filtros ]

				#region [ Tipo de Detalhamento ]
				switch (filtro_detalhe)
				{
					case Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.SINTETICO:
						descricao_filtro_detalhe = "Sintético (Sem Custos)";
						break;
					case Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO:
						descricao_filtro_detalhe = "Intermediário (Custos Médios)";
						break;
					default:
						descricao_filtro_detalhe = "Parâmetro Desconhecido";
						break;
				}
				#endregion

				#region [ Tipo de Consulta (consolidação) ]
				switch (filtro_consolidacao_codigos)
				{
					case Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_CONSOLIDACAO_CODIGOS.NORMAIS:
						descricao_filtro_consolidacao_codigos = "Produtos Normais";
						break;
					case Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_CONSOLIDACAO_CODIGOS.UNIFICADOS:
						descricao_filtro_consolidacao_codigos = "Produtos Unificados";
						break;
					default:
						descricao_filtro_consolidacao_codigos = "";
						break;
				}
				#endregion

				#region [ Empresa ]
				if ((filtro_empresa ?? "").Length == 0)
				{
					descricao_filtro_empresa = "N.I.";
				}
				else
				{
					emitente = NFeEmitenteDAO.getNFeEmitenteById((int)Global.converteInteiro(filtro_empresa));
					if (emitente == null)
					{
						descricao_filtro_empresa = "N.I.";
					}
					else
					{
						descricao_filtro_empresa = emitente.apelido;
					}
				}
				#endregion

				#endregion

				using (ExcelPackage pck = new ExcelPackage())
				{
					//Cria uma planilha com nome
					ExcelWorksheet ws = pck.Workbook.Worksheets.Add("RelEstoqueVenda");

					#region [ Configurações gerais ]
					//configurações gerais da planilha
					ws.Cells["A:XFD"].Style.Font.Name = "Arial";
					ws.Cells["A:XFD"].Style.Font.Size = 10;
					ws.View.ShowGridLines = false;
					ws.Column(1).Width = 2;

					colFabricanteCodigo = COL_INICIAL;
					colProdutoCodigo = colFabricanteCodigo + 1;
					colProdutoDescricao = colProdutoCodigo + 1;
					colCubagem = colProdutoDescricao + 1;
					colQtde = colCubagem + 1;
					COL_FINAL = colQtde;
					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
					{
						colVlCustoMedio = colQtde + 1;
						colVlCustoTotal = colVlCustoMedio + 1;
						COL_FINAL = colVlCustoTotal;
					}

					ws.Column(colFabricanteCodigo).Width = 7;
					ws.Column(colProdutoCodigo).Width = 12;
					ws.Column(colProdutoDescricao).Width = 55;
					ws.Column(colCubagem).Width = 14;
					ws.Column(colQtde).Width = 12;
					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
					{
						ws.Column(colVlCustoMedio).Width = 14;
						ws.Column(colVlCustoTotal).Width = 14;
					}

					ws.Row(1).Height = 1;
					#endregion

					#region [ Informativo dos filtros utilizados ]
					nLinha = ROW_INICIAL;
					cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Style.Font.Size = 12;
					ws.Cells[cellsIndex].Value = "Relatório Estoque de Venda";
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Estoque de Interesse: " + Global.getDescricaoEstoque(filtro_estoque);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Tipo de Detalhamento: " + descricao_filtro_detalhe;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Tipo de Consulta: " + descricao_filtro_consolidacao_codigos;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Empresa: " + descricao_filtro_empresa;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Fabricante: " + ((filtro_fabricante ?? "").Length == 0 ? "N.I." : filtro_fabricante);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Produto: " + ((filtro_produto ?? "").Length == 0 ? "N.I." : filtro_produto);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Fabricante(s): " + ((filtro_fabricante_multiplo ?? "").Length == 0 ? "N.I." : filtro_fabricante_multiplo.Replace("_", ", "));
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Grupo(s): " + ((filtro_grupo ?? "").Length == 0 ? "N.I." : filtro_grupo.Replace("_", ", "));
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Subgrupo(s): " + ((filtro_subgrupo ?? "").Length == 0 ? "N.I." : filtro_subgrupo.Replace("_", ", "));
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "BTU/h: " + ((filtro_potencia_BTU ?? "").Length == 0 ? "N.I." : filtro_potencia_BTU);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Ciclo: " + ((filtro_ciclo ?? "").Length == 0 ? "N.I." : filtro_ciclo);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Posição Mercado: " + ((filtro_posicao_mercado ?? "").Length == 0 ? "N.I." : filtro_posicao_mercado);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");
					cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIAL, (COL_INICIAL + 11), nLinha);
					ws.Cells[cellsIndex].Style.Font.Bold = true;
					#endregion

					#region [ Cabeçalho ]
					nLinha++;
					ROW_CABECALHO = nLinha + 1;
					ROW_INICIO_DADOS = ROW_CABECALHO + 1;

					cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_CABECALHO, COL_FINAL, ROW_CABECALHO);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.WrapText = true;
						rng.Style.Font.Bold = true;
						rng.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
						rng.Style.Border.Top.Style = ExcelBorderStyle.Medium;
						rng.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
						rng.Style.Border.Left.Style = ExcelBorderStyle.Medium;
						rng.Style.Border.Right.Style = ExcelBorderStyle.Medium;
						rng.Style.WrapText = true;
					}

					ws.Cells[ROW_CABECALHO, colFabricanteCodigo].Value = "Fabr";
					ws.Cells[ROW_CABECALHO, colFabricanteCodigo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colProdutoCodigo].Value = "Produto";
					ws.Cells[ROW_CABECALHO, colProdutoCodigo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colProdutoDescricao].Value = "Descrição";
					ws.Cells[ROW_CABECALHO, colProdutoDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colCubagem].Value = "Cubagem (Un)";
					ws.Cells[ROW_CABECALHO, colCubagem].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					ws.Cells[ROW_CABECALHO, colQtde].Value = "Qtde";
					ws.Cells[ROW_CABECALHO, colQtde].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
					{
						ws.Cells[ROW_CABECALHO, colVlCustoMedio].Value = "Custo Entrada Unitário Médio";
						ws.Cells[ROW_CABECALHO, colVlCustoMedio].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colVlCustoTotal].Value = "Custo Entrada Total";
						ws.Cells[ROW_CABECALHO, colVlCustoTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					ws.View.FreezePanes((ROW_CABECALHO + 1), 1);
					#endregion

					#region [ Registros ]
					rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;

					#region [ Formatação de colunas de dados ]
					cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIO_DADOS, COL_FINAL, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
						rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					}

					cellsIndex = Excel.RangeAddress(COL_INICIAL, rowUltLinhaDados, COL_FINAL, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
					}

					cellsIndex = Excel.RangeAddress(colFabricanteCodigo, ROW_INICIO_DADOS, colFabricanteCodigo, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "@";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = false;
					}

					cellsIndex = Excel.RangeAddress(colProdutoCodigo, ROW_INICIO_DADOS, colProdutoCodigo, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "@";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = false;
					}

					cellsIndex = Excel.RangeAddress(colProdutoDescricao, ROW_INICIO_DADOS, colProdutoDescricao, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "@";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = true;
					}

					cellsIndex = Excel.RangeAddress(colCubagem, ROW_INICIO_DADOS, colCubagem, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "#,##0.000000";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					cellsIndex = Excel.RangeAddress(colQtde, ROW_INICIO_DADOS, colQtde, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "#,##0";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
					{
						cellsIndex = Excel.RangeAddress(colVlCustoMedio, ROW_INICIO_DADOS, colVlCustoMedio, rowUltLinhaDados);
						using (ExcelRange rng = ws.Cells[cellsIndex])
						{
							rng.Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
							rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						}

						cellsIndex = Excel.RangeAddress(colVlCustoTotal, ROW_INICIO_DADOS, colVlCustoTotal, rowUltLinhaDados);
						using (ExcelRange rng = ws.Cells[cellsIndex])
						{
							rng.Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
							rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						}
					}
					#endregion

					#region [ Preenche linhas de dados ]
					nLinha = ROW_INICIO_DADOS - 1;
					for (int iFab = 0; iFab < dataRel.Count; iFab++)
					{
						for (int iProd = 0; iProd < dataRel[iFab].Produtos.Count; iProd++)
						{
							nLinha++;
							ws.Cells[nLinha, colFabricanteCodigo].Value = dataRel[iFab].FabricanteCodigo;
							ws.Cells[nLinha, colProdutoCodigo].Value = dataRel[iFab].Produtos.ElementAt(iProd).ProdutoCodigo;
							ws.Cells[nLinha, colProdutoDescricao].Value = dataRel[iFab].Produtos.ElementAt(iProd).ProdutoDescricao;
							ws.Cells[nLinha, colCubagem].Value = dataRel[iFab].Produtos.ElementAt(iProd).Cubagem;
							ws.Cells[nLinha, colQtde].Value = dataRel[iFab].Produtos.ElementAt(iProd).Qtde;
							if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
							{
								if (dataRel[iFab].Produtos.ElementAt(iProd).Qtde != 0) ws.Cells[nLinha, colVlCustoMedio].Value = dataRel[iFab].Produtos.ElementAt(iProd).VlTotal / dataRel[iFab].Produtos.ElementAt(iProd).Qtde;
								ws.Cells[nLinha, colVlCustoTotal].Value = dataRel[iFab].Produtos.ElementAt(iProd).VlTotal;
							}
						}
					}
					#endregion

					#region [ Total ]
					rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
					rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
					ws.Cells[rowTotal, COL_INICIAL, rowTotal, COL_FINAL].Style.Font.Bold = true;
					ws.Cells[rowTotal, colProdutoDescricao].Value = "TOTAL";
					ws.Cells[rowTotal, colProdutoDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					ws.Cells[rowTotal, colQtde].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colQtde, rowUltLinhaDados, colQtde));
					ws.Cells[rowTotal, colQtde].Style.Numberformat.Format = "#,##0";
					if (filtro_detalhe.Equals(Global.Cte.Relatorio.RelEstoqueVenda.FILTRO_DETALHE.INTERMEDIARIO))
					{
						ws.Cells[rowTotal, colVlCustoTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colVlCustoTotal, rowUltLinhaDados, colVlCustoTotal));
						ws.Cells[rowTotal, colVlCustoTotal].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
					}
					#endregion

					#endregion

					pck.SaveAs(new FileInfo(filePath));
				} // using (ExcelPackage pck = new ExcelPackage())
			});
		}
	}
}