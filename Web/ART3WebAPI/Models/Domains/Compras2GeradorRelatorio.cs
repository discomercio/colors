using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ART3WebAPI.Models.Entities;
using System;


namespace ART3WebAPI.Models.Domains
{
	public class Compras2GeradorRelatorio
	{
		public static Task GenerateXLS(List<Compras> datasource, string filePath, string tipo_periodo, string dt_inicio, string dt_termino, string fabricante, string produto, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string nf, string dt_nf_inicio, string dt_nf_termino, string visao, string detalhamento)
		{
			return Task.Run(() =>
			{
				const int COL_INICIAL = 2;
				const int ROW_INICIAL = 2;
				int ROW_CABECALHO, ROW_INICIO_DADOS;
				int colFabricante = 0, colProduto = 0, colDescricao = 0, colGrupo = 0, colSubgrupo = 0, colBTU = 0, colCiclo = 0, colNF = 0, colVlReferencia = 0, colMesInicial = 0, colMesAux, colQtdeTotal = 0, colVlTotal = 0;
				int rowTotal, rowUltLinhaDados;
				int totalMeses = 0;
				int nLinha = 0;
				int NumRegistros = datasource.Count;
				DateTime dti, dtf;
				string periodoentrada = "";
				string emissaoNF = "";
				string cellsIndex;

				if (visao.Equals(Global.Cte.Relatorio.Compras2.COD_VISAO_ANALITICA))
				{
					if (!detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_NF))
					{
						if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
						{
							totalMeses = ((Global.converteDdMmYyyyParaDateTime(dt_termino).Year - Global.converteDdMmYyyyParaDateTime(dt_inicio).Year) * 12) + (Global.converteDdMmYyyyParaDateTime(dt_termino).Month - Global.converteDdMmYyyyParaDateTime(dt_inicio).Month) + 1;
						}
						else
						{
							totalMeses = ((Global.converteDdMmYyyyParaDateTime(dt_nf_termino).Year - Global.converteDdMmYyyyParaDateTime(dt_nf_inicio).Year) * 12) + (Global.converteDdMmYyyyParaDateTime(dt_nf_termino).Month - Global.converteDdMmYyyyParaDateTime(dt_nf_inicio).Month) + 1;
						}
					}
				}

				if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
				{
					dti = Global.converteDdMmYyyyParaDateTime(dt_inicio);
					dtf = Global.converteDdMmYyyyParaDateTime(dt_termino);
				}
				else
				{
					dti = Global.converteDdMmYyyyParaDateTime(dt_nf_inicio);
					dtf = Global.converteDdMmYyyyParaDateTime(dt_nf_termino);
				}

				if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
				{
					periodoentrada = "de " + dt_inicio + " a " + dt_termino;
				}

				if (periodoentrada == "") periodoentrada = "N.I";

				if (!string.IsNullOrEmpty(fabricante))
					fabricante = fabricante.Replace("_", ", ");
				else
					fabricante = "N.I";

				if (!string.IsNullOrEmpty(grupo))
					grupo = grupo.Replace("_", ", ");
				else
					grupo = "N.I";

				if (!string.IsNullOrEmpty(subgrupo))
					subgrupo = subgrupo.Replace("_", ", ");
				else
					subgrupo = "N.I";

				if (string.IsNullOrEmpty(produto)) produto = "N.I";

				if (string.IsNullOrEmpty(btu)) btu = "N.I";

				if (string.IsNullOrEmpty(ciclo)) ciclo = "N.I";

				if (string.IsNullOrEmpty(pos_mercado)) pos_mercado = "N.I";

				if (string.IsNullOrEmpty(nf)) nf = "N.I";

				if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_EMISSAO_NF_ENTRADA))
				{
					emissaoNF = "de " + dt_nf_inicio + " a " + dt_nf_termino;
				}

				if (emissaoNF == "") emissaoNF = "N.I";

				using (ExcelPackage pck = new ExcelPackage())
				{
					//Cria uma planilha com nome
					ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Compras 2");

					#region [ Config Gerais ]
					//configurações gerais da planilha
					ws.Cells["A:XFD"].Style.Font.Name = "Arial";
					ws.Cells["A:XFD"].Style.Font.Size = 10;
					ws.View.ShowGridLines = false;
					ws.Column(1).Width = 2;
					if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_FABR))
					{
						colFabricante = COL_INICIAL;
						colDescricao = colFabricante + 1;
						colMesInicial = colDescricao + 1;
						colVlTotal = colMesInicial + totalMeses;
						ws.Column(colFabricante).Width = 6;
						ws.Column(colDescricao).Width = 51;
						ws.Column(colMesInicial).Width = 12;
						ws.Column(colVlTotal).Width = 15;
					}
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_NF))
					{
						colNF = COL_INICIAL;
						colFabricante = colNF + 1;
						colQtdeTotal = colFabricante + 1;
						colVlTotal = colQtdeTotal + 1;
						ws.Column(colNF).Width = 18;
						ws.Column(colFabricante).Width = 58;
						ws.Column(colQtdeTotal).Width = 11;
						ws.Column(colVlTotal).Width = 15;
					}
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_PROD))
					{
						colFabricante = COL_INICIAL;
						colProduto = colFabricante + 1;
						colDescricao = colProduto + 1;
						colGrupo = colDescricao + 1;
						colSubgrupo = colGrupo + 1;
						colBTU = colSubgrupo + 1;
						colCiclo = colBTU + 1;
						colMesInicial = colCiclo + 1;
						colQtdeTotal = colMesInicial + totalMeses;

						ws.Column(colFabricante).Width = 6;
						ws.Column(colProduto).Width = 9;
						ws.Column(colDescricao).Width = 51;
						ws.Column(colGrupo).Width = 8;
						ws.Column(colSubgrupo).Width = 10;
						ws.Column(colBTU).Width = 9;
						ws.Column(colCiclo).Width = 7;
						ws.Column(colMesInicial).Width = 9;
						ws.Column(colQtdeTotal).Width = 9;
					}
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_CUSTO_MEDIO))
					{
						colFabricante = COL_INICIAL;
						colProduto = colFabricante + 1;
						colDescricao = colProduto + 1;
						colGrupo = colDescricao + 1;
						colSubgrupo = colGrupo + 1;
						colBTU = colSubgrupo + 1;
						colCiclo = colBTU + 1;
						colVlReferencia = colCiclo + 1;
						colMesInicial = colVlReferencia + 1;
						colQtdeTotal = colMesInicial + totalMeses;
						colVlTotal = colQtdeTotal + 1;

						ws.Column(colFabricante).Width = 6;
						ws.Column(colProduto).Width = 9;
						ws.Column(colDescricao).Width = 51;
						ws.Column(colGrupo).Width = 8;
						ws.Column(colSubgrupo).Width = 10;
						ws.Column(colBTU).Width = 9;
						ws.Column(colCiclo).Width = 7;
						ws.Column(colVlReferencia).Width = 12;
						ws.Column(colMesInicial).Width = 9;
						ws.Column(colQtdeTotal).Width = 9;
						ws.Column(colVlTotal).Width = 15;
					}
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_CUSTO_INDIVIDUAL))
					{
						colFabricante = COL_INICIAL;
						colProduto = colFabricante + 1;
						colDescricao = colProduto + 1;
						colGrupo = colDescricao + 1;
						colSubgrupo = colGrupo + 1;
						colBTU = colSubgrupo + 1;
						colCiclo = colBTU + 1;
						colVlReferencia = colCiclo + 1;
						colMesInicial = colVlReferencia + 1;
						colQtdeTotal = colMesInicial + totalMeses;
						colVlTotal = colQtdeTotal + 1;

						ws.Column(colFabricante).Width = 6;
						ws.Column(colProduto).Width = 9;
						ws.Column(colDescricao).Width = 51;
						ws.Column(colGrupo).Width = 8;
						ws.Column(colSubgrupo).Width = 10;
						ws.Column(colBTU).Width = 9;
						ws.Column(colCiclo).Width = 7;
						ws.Column(colVlReferencia).Width = 12;
						ws.Column(colMesInicial).Width = 9;
						ws.Column(colQtdeTotal).Width = 9;
						ws.Column(colVlTotal).Width = 15;
					}

					ws.Row(1).Height = 1;
					#endregion

					#region [ Filtro ]
					nLinha = ROW_INICIAL;
					cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Style.Font.Size = 12;
					ws.Cells[cellsIndex].Value = "Compras II";
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					if (tipo_periodo.Equals(Global.Cte.Relatorio.Compras2.COD_CONSULTA_POR_PERIODO_ENTRADA_ESTOQUE))
					{
						ws.Cells[cellsIndex].Value = "Período da Entrada no Estoque: " + periodoentrada;
					}
					else
					{
						ws.Cells[cellsIndex].Value = "Período da Emissão NF Entrada: " + emissaoNF;
					}
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Fabricante(s): " + fabricante;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Grupo(s) de produtos: " + grupo;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Subgrupo(s) de produtos: " + subgrupo;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Produto: " + produto;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "BTU/h: " + btu;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Ciclo: " + ciclo;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Posição Mercado: " + pos_mercado;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Nº Nota Fiscal: " + nf;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Tipo de Detalhamento: " + Global.getDetalhamento(detalhamento);
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

					cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIAL, (COL_INICIAL + 11), nLinha);
					ws.Cells[cellsIndex].Style.Font.Bold = true;

					nLinha++;
					ROW_CABECALHO = nLinha + 1;
					ROW_INICIO_DADOS = ROW_CABECALHO + 1;
					#endregion

					#region [ Cabeçalho ]

					#region [ Sintético por NF ]
					if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_NF))
					{
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_CABECALHO, colVlTotal, ROW_CABECALHO);
						using (ExcelRange rng1 = ws.Cells[cellsIndex])
						{
							rng1.Style.WrapText = true;
							rng1.Style.Font.Bold = true;
							rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
							rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						}

						ws.Cells[ROW_CABECALHO, colVlTotal].Style.Border.Right.Style = ExcelBorderStyle.Medium;
						ws.Cells[ROW_CABECALHO, colNF].Value = "Nº NF";
						ws.Cells[ROW_CABECALHO, colFabricante].Value = "Fabr";
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Value = "Qtde Total";
						ws.Cells[ROW_CABECALHO, colVlTotal].Value = "Valor Total";
						ws.Cells[ROW_CABECALHO, colNF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colVlTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}
					#endregion

					#region [ Sintético por Fabricante ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_FABR))
					{
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_CABECALHO, colVlTotal, ROW_CABECALHO);
						using (ExcelRange rng1 = ws.Cells[cellsIndex])
						{
							rng1.Style.WrapText = true;
							rng1.Style.Font.Bold = true;
							rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
							rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						}

						ws.Cells[ROW_CABECALHO, colVlTotal].Style.Border.Right.Style = ExcelBorderStyle.Medium;
						ws.Cells[ROW_CABECALHO, colFabricante].Value = "Fabr";
						ws.Cells[ROW_CABECALHO, colDescricao].Value = "Descrição";
						ws.Cells[ROW_CABECALHO, colVlTotal].Value = "Valor Total";
						ws.Cells[ROW_CABECALHO, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colVlTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

						for (int i = 1; i <= totalMeses; i++)
						{
							string mes = "", ano = "";
							if (i == 1)
							{
								mes = dti.ToString("MM");
								ano = dti.ToString("yyyy");
							}
							else
							{
								mes = (dti.AddMonths(i - 1)).ToString("MM");
								ano = (dti.AddMonths(i - 1)).ToString("yyyy");
							}

							colMesAux = colMesInicial + (i - 1);
							ws.Column(colMesAux).Width = 15;
							cellsIndex = Excel.CellAddress(colMesAux, ROW_CABECALHO);
							ws.Cells[cellsIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Cells[cellsIndex].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
							//Total
							rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
							rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
							ws.Cells[rowTotal, colMesAux].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colMesAux, rowUltLinhaDados, colMesAux));
							ws.Cells[rowTotal, colMesAux].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						}
					}
					#endregion

					#region [ Sintético por Produto ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_PROD))
					{
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_CABECALHO, colQtdeTotal, ROW_CABECALHO);
						using (ExcelRange rng1 = ws.Cells[cellsIndex])
						{
							rng1.Style.WrapText = true;
							rng1.Style.Font.Bold = true;
							rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
							rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						}

						ws.Cells[ROW_CABECALHO, colQtdeTotal].Style.Border.Right.Style = ExcelBorderStyle.Medium;
						ws.Cells[ROW_CABECALHO, colFabricante].Value = "Fabr";
						ws.Cells[ROW_CABECALHO, colProduto].Value = "Produto";
						ws.Cells[ROW_CABECALHO, colDescricao].Value = "Descrição";
						ws.Cells[ROW_CABECALHO, colGrupo].Value = "Grupo";
						ws.Cells[ROW_CABECALHO, colSubgrupo].Value = "Subgrupo";
						ws.Cells[ROW_CABECALHO, colBTU].Value = "BTU/h";
						ws.Cells[ROW_CABECALHO, colCiclo].Value = "Ciclo";
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Value = " Qtde Total";
						ws.Cells[ROW_CABECALHO, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colProduto].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colGrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colSubgrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colBTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colCiclo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colMesInicial].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

						for (int i = 1; i <= totalMeses; i++)
						{
							string mes = "", ano = "";
							if (i == 1)
							{
								mes = dti.ToString("MM");
								ano = dti.ToString("yyyy");
							}
							else
							{
								mes = (dti.AddMonths(i - 1)).ToString("MM");
								ano = (dti.AddMonths(i - 1)).ToString("yyyy");
							}

							colMesAux = colMesInicial + (i - 1);
							ws.Column(colMesAux).Width = 9;
							ws.Cells[ROW_CABECALHO, colMesAux].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Cells[ROW_CABECALHO, colMesAux].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
							//Total
							rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
							rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
							ws.Cells[rowTotal, colMesAux].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colMesAux, rowUltLinhaDados, colMesAux));
						}
					}
					#endregion

					#region [ Valor Referência Médio ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_CUSTO_MEDIO))
					{
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_CABECALHO, colVlTotal, ROW_CABECALHO);
						using (ExcelRange rng1 = ws.Cells[cellsIndex])
						{
							rng1.Style.WrapText = true;
							rng1.Style.Font.Bold = true;
							rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
							rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						}

						ws.Cells[ROW_CABECALHO, colVlTotal].Style.Border.Right.Style = ExcelBorderStyle.Medium;
						ws.Cells[ROW_CABECALHO, colFabricante].Value = "Fabr";
						ws.Cells[ROW_CABECALHO, colProduto].Value = "Produto";
						ws.Cells[ROW_CABECALHO, colDescricao].Value = "Descrição";
						ws.Cells[ROW_CABECALHO, colGrupo].Value = "Grupo";
						ws.Cells[ROW_CABECALHO, colSubgrupo].Value = "Subgrupo";
						ws.Cells[ROW_CABECALHO, colBTU].Value = "BTU/h";
						ws.Cells[ROW_CABECALHO, colCiclo].Value = "Ciclo";
						ws.Cells[ROW_CABECALHO, colVlReferencia].Value = "Referência Médio";
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Value = "Qtde Total";
						ws.Cells[ROW_CABECALHO, colVlTotal].Value = "Valor Total";
						ws.Cells[ROW_CABECALHO, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colProduto].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colGrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colSubgrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colBTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colCiclo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colVlReferencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colVlTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

						for (int i = 1; i <= totalMeses; i++)
						{
							string mes = "", ano = "";
							if (i == 1)
							{
								mes = dti.ToString("MM");
								ano = dti.ToString("yyyy");
							}
							else
							{
								mes = (dti.AddMonths(i - 1)).ToString("MM");
								ano = (dti.AddMonths(i - 1)).ToString("yyyy");
							}

							colMesAux = colMesInicial + (i - 1);
							ws.Cells[ROW_CABECALHO, colMesAux].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Column(colMesAux).Width = 9;
							ws.Cells[ROW_CABECALHO, colMesAux].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
							//Total
							rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
							rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
							ws.Cells[rowTotal, colMesAux].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colMesAux, rowUltLinhaDados, colMesAux));
						}
					}
					#endregion

					#region [ Valor Referência Individual ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_CUSTO_INDIVIDUAL))
					{
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_CABECALHO, colVlTotal, ROW_CABECALHO);
						using (ExcelRange rng1 = ws.Cells[cellsIndex])
						{
							rng1.Style.WrapText = true;
							rng1.Style.Font.Bold = true;
							rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
							rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
							rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						}

						ws.Cells[ROW_CABECALHO, colVlTotal].Style.Border.Right.Style = ExcelBorderStyle.Medium;
						ws.Cells[ROW_CABECALHO, colFabricante].Value = "Fabr";
						ws.Cells[ROW_CABECALHO, colProduto].Value = "Produto";
						ws.Cells[ROW_CABECALHO, colDescricao].Value = "Descrição";
						ws.Cells[ROW_CABECALHO, colGrupo].Value = "Grupo";
						ws.Cells[ROW_CABECALHO, colSubgrupo].Value = "Subgrupo";
						ws.Cells[ROW_CABECALHO, colBTU].Value = "BTU/h";
						ws.Cells[ROW_CABECALHO, colCiclo].Value = "Ciclo";
						ws.Cells[ROW_CABECALHO, colVlReferencia].Value = "Referência Individual";
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Value = "Qtde Total";
						ws.Cells[ROW_CABECALHO, colVlTotal].Value = "Valor Total";
						ws.Cells[ROW_CABECALHO, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colProduto].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[ROW_CABECALHO, colGrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colSubgrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colBTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colCiclo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						ws.Cells[ROW_CABECALHO, colVlReferencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colQtdeTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[ROW_CABECALHO, colVlTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

						for (int i = 1; i <= totalMeses; i++)
						{
							string mes = "", ano = "";
							if (i == 1)
							{
								mes = dti.ToString("MM");
								ano = dti.ToString("yyyy");
							}
							else
							{
								mes = (dti.AddMonths(i - 1)).ToString("MM");
								ano = (dti.AddMonths(i - 1)).ToString("yyyy");
							}

							colMesAux = colMesInicial + (i - 1);
							ws.Cells[ROW_CABECALHO, colMesAux].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Column(colMesAux).Width = 9;
							ws.Cells[ROW_CABECALHO, colMesAux].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
							//Total
							rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
							rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
							ws.Cells[rowTotal, colMesAux].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colMesAux, rowUltLinhaDados, colMesAux));
						}
					}
					#endregion

					#endregion

					ws.View.FreezePanes((ROW_CABECALHO + 1), 1);

					#region [ Registros ]

					#region [ Sintético por NF ]
					if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_NF))
					{
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados);
						using (ExcelRange rng2 = ws.Cells[cellsIndex])
						{
							rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
							rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
							rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						}

						#region [ Preenche linhas de dados ]
						for (int i = 0; i < NumRegistros; i++)
						{
							ws.Cells[i + ROW_INICIO_DADOS, colNF].Value = datasource.ElementAt(i).NF;
							ws.Cells[i + ROW_INICIO_DADOS, colNF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Value = datasource.ElementAt(i).Fabricante + " - " + datasource.ElementAt(i).FabricanteNome;
							ws.Cells[i + ROW_INICIO_DADOS, colQtdeTotal].Value = datasource.ElementAt(i).Qtde;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Value = datasource.ElementAt(i).Valor;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Style.Numberformat.Format = "###,###,##0.00";
						}
						#endregion

						#region [ Total ]
						rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						ws.Cells[rowTotal, COL_INICIAL, rowTotal, colVlTotal].Style.Font.Bold = true;
						ws.Cells[rowTotal, colQtdeTotal - 1].Value = "TOTAL";
						ws.Cells[rowTotal, colQtdeTotal - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[rowTotal, colQtdeTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colQtdeTotal, rowUltLinhaDados, colQtdeTotal));
						ws.Cells[rowTotal, colVlTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados, colVlTotal));
						ws.Cells[rowTotal, colVlTotal].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						#endregion
					}
					#endregion

					#region [ Sintético por Fabricante ]
					if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_FABR))
					{
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados);
						using (ExcelRange rng2 = ws.Cells[cellsIndex])
						{
							rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
							rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
							rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						}

						#region [ Preenche linhas de dados ]
						for (int i = 0; i < NumRegistros; i++)
						{
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Value = datasource.ElementAt(i).Fabricante;
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Value = datasource.ElementAt(i).FabricanteNome;
							if (visao.Equals(Global.Cte.Relatorio.Compras2.COD_VISAO_ANALITICA))
							{
								for (int j = 0; j < totalMeses; j++)
								{
									colMesAux = colMesInicial + j;
									ws.Cells[i + ROW_INICIO_DADOS, colMesAux].Value = datasource.ElementAt(i).Meses[j];
									ws.Cells[i + ROW_INICIO_DADOS, colMesAux].Style.Numberformat.Format = "###,###,##0.00";
								}
							}
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Value = datasource.ElementAt(i).Valor;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Style.Numberformat.Format = "###,###,##0.00";
						}
						#endregion

						#region [ Total ]
						rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						ws.Cells[rowTotal, COL_INICIAL, rowTotal, colVlTotal].Style.Font.Bold = true;
						ws.Cells[rowTotal, colMesInicial - 1].Value = "TOTAL";
						ws.Cells[rowTotal, colMesInicial - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[rowTotal, colVlTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados, colVlTotal));
						ws.Cells[rowTotal, colVlTotal].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						#endregion
					}
					#endregion

					#region [ Sintético por Produto ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_SINTETICO_PROD))
					{
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIO_DADOS, colQtdeTotal, rowUltLinhaDados);
						using (ExcelRange rng2 = ws.Cells[cellsIndex])
						{
							rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
							rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
							rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						}

						#region [ Preenche linhas de dados ]
						for (int i = 0; i < NumRegistros; i++)
						{
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Value = datasource.ElementAt(i).Fabricante;
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colProduto].Value = datasource.ElementAt(i).Produto;
							ws.Cells[i + ROW_INICIO_DADOS, colProduto].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Value = datasource.ElementAt(i).ProdutoDescricao;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colGrupo].Value = datasource.ElementAt(i).Grupo;
							ws.Cells[i + ROW_INICIO_DADOS, colGrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							ws.Cells[i + ROW_INICIO_DADOS, colSubgrupo].Value = datasource.ElementAt(i).Subgrupo;
							ws.Cells[i + ROW_INICIO_DADOS, colSubgrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							if (datasource.ElementAt(i).Potencia_BTU != 0) ws.Cells[i + ROW_INICIO_DADOS, colBTU].Value = datasource.ElementAt(i).Potencia_BTU;
							ws.Cells[i + ROW_INICIO_DADOS, colBTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Cells[i + ROW_INICIO_DADOS, colBTU].Style.Numberformat.Format = "##,###";
							ws.Cells[i + ROW_INICIO_DADOS, colCiclo].Value = datasource.ElementAt(i).Ciclo;
							ws.Cells[i + ROW_INICIO_DADOS, colCiclo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

							if (visao.Equals(Global.Cte.Relatorio.Compras2.COD_VISAO_ANALITICA))
							{
								for (int j = 0; j < totalMeses; j++)
								{
									colMesAux = colMesInicial + j;
									ws.Cells[i + ROW_INICIO_DADOS, colMesAux].Value = datasource.ElementAt(i).Meses[j];
								}
							}

							ws.Cells[i + ROW_INICIO_DADOS, colQtdeTotal].Value = datasource.ElementAt(i).Qtde;
						}
						#endregion

						#region [ Total ]
						rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						ws.Cells[rowTotal, colFabricante, rowTotal, colQtdeTotal].Style.Font.Bold = true;
						ws.Cells[rowTotal, colMesInicial - 1].Value = "TOTAL";
						ws.Cells[rowTotal, colMesInicial - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[rowTotal, colQtdeTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colQtdeTotal, rowUltLinhaDados, colQtdeTotal));
						for (int k = 0; k < totalMeses; k++)
						{
							colMesAux = colMesInicial + k;
							ws.Cells[rowTotal, colMesAux].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colMesAux, rowUltLinhaDados, colMesAux));
						}
						#endregion
					}
					#endregion

					#region [ Valor Referência Médio ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_CUSTO_MEDIO))
					{
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados);
						using (ExcelRange rng2 = ws.Cells[cellsIndex])
						{
							rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
							rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
							rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						}

						#region [ Preenche linhas de dados ]
						for (int i = 0; i < NumRegistros; i++)
						{
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Value = datasource.ElementAt(i).Fabricante;
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colProduto].Value = datasource.ElementAt(i).Produto;
							ws.Cells[i + ROW_INICIO_DADOS, colProduto].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Value = datasource.ElementAt(i).ProdutoDescricao;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colGrupo].Value = datasource.ElementAt(i).Grupo;
							ws.Cells[i + ROW_INICIO_DADOS, colGrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							ws.Cells[i + ROW_INICIO_DADOS, colSubgrupo].Value = datasource.ElementAt(i).Subgrupo;
							ws.Cells[i + ROW_INICIO_DADOS, colSubgrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							if (datasource.ElementAt(i).Potencia_BTU != 0) ws.Cells[i + ROW_INICIO_DADOS, colBTU].Value = datasource.ElementAt(i).Potencia_BTU;
							ws.Cells[i + ROW_INICIO_DADOS, colBTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Cells[i + ROW_INICIO_DADOS, colBTU].Style.Numberformat.Format = "##,###";
							ws.Cells[i + ROW_INICIO_DADOS, colCiclo].Value = datasource.ElementAt(i).Ciclo;
							ws.Cells[i + ROW_INICIO_DADOS, colCiclo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							ws.Cells[i + ROW_INICIO_DADOS, colVlReferencia].Value = datasource.ElementAt(i).Valor / datasource.ElementAt(i).Qtde;
							ws.Cells[i + ROW_INICIO_DADOS, colVlReferencia].Style.Numberformat.Format = "###,###,##0.00";

							if (visao.Equals(Global.Cte.Relatorio.Compras2.COD_VISAO_ANALITICA))
							{
								for (int j = 0; j < totalMeses; j++)
								{
									colMesAux = colMesInicial + j;
									ws.Cells[i + ROW_INICIO_DADOS, colMesAux].Value = datasource.ElementAt(i).Meses[j];
								}
							}
							ws.Cells[i + ROW_INICIO_DADOS, colQtdeTotal].Value = datasource.ElementAt(i).Qtde;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Value = datasource.ElementAt(i).Valor;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Style.Numberformat.Format = "###,###,##0.00";
						}
						#endregion

						#region [ Total ]
						rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						ws.Cells[rowTotal, COL_INICIAL, rowTotal, colVlTotal].Style.Font.Bold = true;
						ws.Cells[rowTotal, colVlReferencia - 1].Value = "TOTAL";
						ws.Cells[rowTotal, colVlReferencia - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[rowTotal, colQtdeTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colQtdeTotal, rowUltLinhaDados, colQtdeTotal));
						ws.Cells[rowTotal, colVlTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados, colVlTotal));
						ws.Cells[rowTotal, colVlTotal].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						#endregion
					}
					#endregion

					#region [ Valor Referência Individual ]
					else if (detalhamento.Equals(Global.Cte.Relatorio.Compras2.COD_SAIDA_CUSTO_INDIVIDUAL))
					{
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						cellsIndex = Excel.RangeAddress(COL_INICIAL, ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados);
						using (ExcelRange rng2 = ws.Cells[cellsIndex])
						{
							rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
							rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
							rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						}

						#region [ Preenche linhas de dados ]
						for (int i = 0; i < NumRegistros; i++)
						{
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Value = datasource.ElementAt(i).Fabricante;
							ws.Cells[i + ROW_INICIO_DADOS, colFabricante].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colProduto].Value = datasource.ElementAt(i).Produto;
							ws.Cells[i + ROW_INICIO_DADOS, colProduto].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Value = datasource.ElementAt(i).ProdutoDescricao;
							ws.Cells[i + ROW_INICIO_DADOS, colDescricao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
							ws.Cells[i + ROW_INICIO_DADOS, colGrupo].Value = datasource.ElementAt(i).Grupo;
							ws.Cells[i + ROW_INICIO_DADOS, colGrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							ws.Cells[i + ROW_INICIO_DADOS, colSubgrupo].Value = datasource.ElementAt(i).Subgrupo;
							ws.Cells[i + ROW_INICIO_DADOS, colSubgrupo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							if (datasource.ElementAt(i).Potencia_BTU != 0) ws.Cells[i + ROW_INICIO_DADOS, colBTU].Value = datasource.ElementAt(i).Potencia_BTU;
							ws.Cells[i + ROW_INICIO_DADOS, colBTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
							ws.Cells[i + ROW_INICIO_DADOS, colBTU].Style.Numberformat.Format = "##,###";
							ws.Cells[i + ROW_INICIO_DADOS, colCiclo].Value = datasource.ElementAt(i).Ciclo;
							ws.Cells[i + ROW_INICIO_DADOS, colCiclo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
							ws.Cells[i + ROW_INICIO_DADOS, colVlReferencia].Value = datasource.ElementAt(i).Valor;
							ws.Cells[i + ROW_INICIO_DADOS, colVlReferencia].Style.Numberformat.Format = "###,###,##0.00";
							if (visao.Equals(Global.Cte.Relatorio.Compras2.COD_VISAO_ANALITICA))
							{
								for (int j = 0; j < totalMeses; j++)
								{
									colMesAux = colMesInicial + j;
									ws.Cells[i + ROW_INICIO_DADOS, colMesAux].Value = datasource.ElementAt(i).Meses[j];
								}
							}
							ws.Cells[i + ROW_INICIO_DADOS, colQtdeTotal].Value = datasource.ElementAt(i).Qtde;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Value = datasource.ElementAt(i).Valor * datasource.ElementAt(i).Qtde;
							ws.Cells[i + ROW_INICIO_DADOS, colVlTotal].Style.Numberformat.Format = "###,###,##0.00";
						}
						#endregion

						#region [ Total ]
						rowTotal = ROW_INICIO_DADOS + NumRegistros + 1;
						rowUltLinhaDados = ROW_INICIO_DADOS + NumRegistros - 1;
						ws.Cells[rowTotal, COL_INICIAL, rowTotal, colVlTotal].Style.Font.Bold = true;
						ws.Cells[rowTotal, colVlReferencia - 1].Value = "TOTAL";
						ws.Cells[rowTotal, colVlReferencia - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						ws.Cells[rowTotal, colQtdeTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colQtdeTotal, rowUltLinhaDados, colQtdeTotal));
						ws.Cells[rowTotal, colVlTotal].Formula = string.Format("SUM({0})", new ExcelAddress(ROW_INICIO_DADOS, colVlTotal, rowUltLinhaDados, colVlTotal));
						ws.Cells[rowTotal, colVlTotal].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						#endregion
					}
					#endregion

					#endregion

					pck.SaveAs(new FileInfo(filePath));
				}
			});
		}
	}
}