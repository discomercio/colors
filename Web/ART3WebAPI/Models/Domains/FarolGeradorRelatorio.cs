using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ART3WebAPI.Models.Entities;
using System;
using ART3WebAPI.Controllers;

namespace ART3WebAPI.Models.Domains
{
    public class FarolGeradorRelatorio
    {
        #region [ Constantes ]
        private const int LIN_INICIO_REGISTROS = 14;
        private const int LIN_CABECALHO = 13;
        private const int LIN_PERCENTUAL = 12;
        private const int COL_FABRICANTE = 2;
        private const int COL_PRODUTO = COL_FABRICANTE + 1;
        private const int COL_DESCRICAO = COL_PRODUTO + 1;
        private const int COL_GRUPO = COL_DESCRICAO + 1;
        private const int COL_SUBGRUPO = COL_GRUPO + 1;
        private const int COL_BTU = COL_SUBGRUPO + 1;
        private const int COL_CICLO = COL_BTU + 1;
        private const int COL_POSICAO_MERCADO = COL_CICLO + 1;
        private const int COL_VENDA_HISTORICO = COL_POSICAO_MERCADO + 1;
        private const int COL_VENDA_PREVISAO = COL_VENDA_HISTORICO + 1;
        private const int COL_ESTOQUE = COL_VENDA_PREVISAO + 1;
        private const int COL_COMPRADO = COL_ESTOQUE + 1;
        private const int COL_SALDO = COL_COMPRADO + 1;
        private const int COL_CUSTO = COL_SALDO + 1;
        private const int COL_IDEIA = COL_CUSTO + 1;
        private const int COL_VALOR = COL_IDEIA + 1;
        #endregion

        #region [ GenerateXLS versão 2 ]
        public static Task GenerateXLS(List<Farol> datasource, string filePath, string dt_inicio, string dt_termino, string fabricante, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string perc_est_cresc, string loja, string visao)
        {
            return Task.Run(() =>
            {
                #region [ Declarações ]
                string totalVenda = "";
                string totalProjecao = "";
                int cont = 0;
                int totalMeses = 0;
                int totalMesesProjecao = 0;
                int lineAux;
                #endregion

                if (visao.Equals("ANALITICA"))
                {
					totalMeses = ((Global.converteDdMmYyyyParaDateTime(dt_termino).Year - Global.converteDdMmYyyyParaDateTime(dt_inicio).Year) * 12) + (Global.converteDdMmYyyyParaDateTime(dt_termino).Month - Global.converteDdMmYyyyParaDateTime(dt_inicio).Month) + 1;
                    totalMesesProjecao = totalMeses;
                }
                
                int NumRegistros = datasource.Count;
				DateTime dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
				DateTime dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);

                if (!string.IsNullOrEmpty(fabricante))
                    fabricante = fabricante.Replace("_", ", ");
                else
                    fabricante = "N.I";
                if (!string.IsNullOrEmpty(grupo))
                    grupo = grupo.Replace("|", ", ");
                else
                    grupo = "N.I";
                if (!string.IsNullOrEmpty(subgrupo))
                    subgrupo = subgrupo.Replace("|", ", ");
                else
                    subgrupo = "N.I";
                if (string.IsNullOrEmpty(btu))
                    btu = "N.I";
                if (string.IsNullOrEmpty(ciclo))
                    ciclo = "N.I";
                if (string.IsNullOrEmpty(pos_mercado))
                    pos_mercado = "N.I";
                if (string.IsNullOrEmpty(perc_est_cresc))
                    perc_est_cresc = "0";
                if (string.IsNullOrEmpty(loja))
                    loja = "N.I";

                using (ExcelPackage pck = new ExcelPackage())
                {
                    //Cria uma planilha com nome
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Farol Resumido");

                    #region [ Config Gerais ]
                    //configurações gerais da planilha
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.View.ShowGridLines = false;
                    ws.View.FreezePanes(14, 1);
                    
                    ws.Column(1).Width = 2;
                    ws.Column(COL_FABRICANTE).Width = 6;
                    ws.Column(COL_PRODUTO).Width = 9;
                    ws.Column(COL_DESCRICAO).Width = 51;
                    ws.Column(COL_GRUPO).Width = 7;
                    ws.Column(COL_SUBGRUPO).Width = 10;
                    ws.Column(COL_BTU).Width = 11;
                    ws.Column(COL_CICLO).Width = 7;                   
                    ws.Column(COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao).Width = 9;
                    ws.Column(COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao).Width = 9;
                    ws.Column(COL_ESTOQUE + totalMeses + totalMesesProjecao).Width = 10;
                    ws.Column(COL_COMPRADO + totalMeses + totalMesesProjecao).Width = 10;
                    ws.Column(COL_SALDO + totalMeses + totalMesesProjecao).Width = 10;

                    ws.Row(1).Height = 1;

                    ws.Cells[LIN_INICIO_REGISTROS, COL_POSICAO_MERCADO + 1, (NumRegistros + LIN_INICIO_REGISTROS + 1), COL_SALDO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0";

                    #endregion                    

                    #region [ Filtro ]
                    lineAux = 2;
                    ws.Cells["B" + lineAux.ToString()].Style.Font.Size = 12;
                    ws.Cells["B" + lineAux.ToString()].Value = "Farol Resumido";
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Período de vendas: " + dt_inicio + " a " + dt_termino;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Fabricante(s): " + fabricante;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Grupo(s) de produtos: " + grupo;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Subgrupo(s) de produtos: " + subgrupo;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "BTU/h: " + btu;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Ciclo: " + ciclo;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Posição Mercado:" + pos_mercado;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Loja: " + loja;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Visão: " + visao;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    lineAux++; ws.Cells["H" + lineAux.ToString()].Value = "Percentual estimado de crescimento (%): ";
                    ws.Cells["B2:M" + lineAux.ToString()].Style.Font.Bold = true;
                    ws.Cells["H" + lineAux.ToString() + ":I" + lineAux.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    if (visao.Equals("ANALITICA"))
                    {
                        for (int i = 1; i <= totalMesesProjecao; i++)
                        {
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Value = perc_est_cresc;
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.Numberformat.Format = "###.0";
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.Font.Color.SetColor(Color.DarkGreen);
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.Font.Bold = true;
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                    }
                    else
                    {
                        ws.Cells["I12"].Value = perc_est_cresc;
                        ws.Cells["I12"].Style.Numberformat.Format = "###.0";
                        ws.Cells["I12"].Style.Font.Color.SetColor(Color.DarkGreen);
                        ws.Cells["I12"].Style.Font.Bold = true;
                        ws.Cells["I12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    
                    
                    #endregion

                    #region [ Cabeçalho ]
                    using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, COL_SALDO + totalMeses + totalMesesProjecao]])
                    {
                        rng1.Style.WrapText = true;
                        rng1.Style.Font.Bold = true;
                        rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }

                    ws.Cells[LIN_CABECALHO, COL_SALDO + totalMeses + totalMesesProjecao].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    ws.Cells[LIN_CABECALHO, COL_FABRICANTE].Value = "Fabr";
                    ws.Cells[LIN_CABECALHO, COL_PRODUTO].Value = "Produto";
                    ws.Cells[LIN_CABECALHO, COL_DESCRICAO].Value = "Descrição";
                    ws.Cells[LIN_CABECALHO, COL_GRUPO].Value = "Grupo";
                    ws.Cells[LIN_CABECALHO, COL_SUBGRUPO].Value = "Subgrupo";
                    ws.Cells[LIN_CABECALHO, COL_BTU].Value = "BTU/h";
                    ws.Cells[LIN_CABECALHO, COL_CICLO].Value = "Ciclo";

                    for (int i = 1; i <= totalMeses; i++)
                    {
                        
                        string mes = "", ano = "";
                        if (i == 1)
                        {
                            mes = dt1.ToString("MM");
                            ano = dt1.ToString("yyyy");

                        }
                        else
                        {
                            mes = (dt1.AddMonths(i - 1)).ToString("MM");
                            ano = (dt1.AddMonths(i - 1)).ToString("yyyy");
                        }

                        if(i== 1)
                        {
                            //Mês a mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i)));                        
                            // Projeção mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i + 1).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Value = "Proj " + Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i + i)));

                        }
                        else
                        {
                            //Mês a mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + cont)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i + cont).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + cont)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i + cont].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i + cont), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i + cont)));
                            // Projeção mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i + i).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Value = "Proj " + Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i + i)));                  
                        }

                        cont++;

                    }                    

                    ws.Cells[LIN_CABECALHO, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Value = "Total";
                    ws.Cells[LIN_CABECALHO, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Value = "Proj Total";
                    ws.Cells[LIN_CABECALHO, COL_ESTOQUE + totalMeses + totalMesesProjecao].Value = "Estoque";
                    ws.Cells[LIN_CABECALHO, COL_COMPRADO + totalMeses + totalMesesProjecao].Value = "Compra";
                    ws.Cells[LIN_CABECALHO, COL_SALDO + totalMeses + totalMesesProjecao].Value = "Saldo";
                    ws.Cells[LIN_CABECALHO, COL_FABRICANTE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_PRODUTO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_DESCRICAO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_GRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_SUBGRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_BTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_CICLO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_ESTOQUE + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_COMPRADO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_SALDO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    #endregion

                    #region [ Registros ]
                    using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS -1, COL_SALDO + totalMeses + totalMesesProjecao]])
                    {
                        rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    for (int i = 0; i < NumRegistros; i++)
                    {
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_FABRICANTE].Value = datasource.ElementAt(i).Fabricante;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_FABRICANTE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_PRODUTO].Value = datasource.ElementAt(i).Produto;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_PRODUTO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_DESCRICAO].Value = datasource.ElementAt(i).Descricao;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_DESCRICAO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_GRUPO].Value = datasource.ElementAt(i).Grupo;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_GRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SUBGRUPO].Value = datasource.ElementAt(i).Subgrupo;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SUBGRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        if (datasource.ElementAt(i).Potencia_BTU != 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_BTU].Value = datasource.ElementAt(i).Potencia_BTU;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_BTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_BTU].Style.Numberformat.Format = "##,###";
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CICLO].Value = datasource.ElementAt(i).Ciclo;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CICLO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        
                        if (visao.Equals("ANALITICA"))
                        {
                            cont = 0;
                            for (int j = 0; j < totalMeses; j++)
                            {
                                if (j == 0)
                                {
                                    //Registros mês a mês                                   
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].Value = datasource.ElementAt(i).Meses[j];
                                    //Registros mês a mês Projeção 
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].FormulaR1C1 = string.Format("RC[-1]+(R{0}C{1}/100)*RC[-1]", LIN_PERCENTUAL, (COL_CICLO + j + 2));
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].Style.Numberformat.Format = "##,##0";
                                    //Registros Projeção Total e Total Vendas
                                    totalVenda = ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1)].ToString();
                                    totalProjecao = ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].ToString();
                                }
                                else
                                {
                                    //Registros mês a mês
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1 + cont)].Value = datasource.ElementAt(i).Meses[j];
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].Value = datasource.ElementAt(i).Meses[j];
                                    //Registros mês a mês Projeção 
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].FormulaR1C1 = string.Format("RC[-1]+(R{0}C{1}/100)*RC[-1]", LIN_PERCENTUAL, (COL_CICLO + j + 2 + cont));
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].Style.Numberformat.Format = "##,##0";

                                    //Registros Projeção Total Total Vendas
                                    totalVenda = totalVenda + ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1 + cont)].ToString();
                                    totalProjecao = totalProjecao + ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].ToString();

                                }
                                if(totalMesesProjecao > j + 1)
                                {
                                    totalVenda = totalVenda + "+"; 
                                    totalProjecao = totalProjecao + "+";
                                }
                                cont++;
                            }
                        }
                        
                        if(visao.Equals("ANALITICA"))
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Formula = string.Format("=SUM({0})", totalVenda);
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Formula = string.Format("=SUM({0})", totalProjecao);
                        }
                        else
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses].Value = datasource.ElementAt(i).Qtde_vendida;
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses].Formula = string.Format("{0}+(I12/100)*{0}", ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses]);
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses].Style.Numberformat.Format = "##,##0";
                        }                        
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;                        
                        if (datasource.ElementAt(i).Qtde_vendida > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;                        
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0";                       
                        if (datasource.ElementAt(i).Qtde_vendida > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao].Value = datasource.ElementAt(i).Qtde_estoque_venda;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (datasource.ElementAt(i).Qtde_estoque_venda > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao].Value = datasource.ElementAt(i).Farol_qtde_comprada;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (datasource.ElementAt(i).Farol_qtde_comprada > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Formula = string.Format("{0}+{1}-{2}", ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao], ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao], ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao]);
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0;[Red]-##,##0";
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        if (datasource.ElementAt(i).Saldo < 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.Font.Color.SetColor(Color.Red);

                    }
                    #endregion


                    #region [ Total ]
                    ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), COL_FABRICANTE, (NumRegistros + LIN_INICIO_REGISTROS + 1), COL_SALDO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7].Value = "TOTAL";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_ESTOQUE + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_ESTOQUE + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_COMPRADO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_COMPRADO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_SALDO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_SALDO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_SALDO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0;[Red]-##,##0";
                    #endregion

                            pck.SaveAs(new FileInfo(filePath));
                }
            });
        }
        #endregion

        #region[ GenerateXLS versão 3 ]
        public static Task GenerateXLSv3(List<Farol> datasource, string filePath, string opcao_periodo, string dt_inicio, string dt_termino, string fabricante, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string perc_est_cresc, string loja, string visao)
        {
            return Task.Run(() =>
            {
                #region [ Declarações ]
                string sAux;
                string totalVenda = "";
                string totalProjecao = "";
                int cont = 0;
                int totalMeses = 0;
                int totalMesesProjecao = 0;
                int lineAux;
                #endregion

                if (visao.Equals("ANALITICA"))
                {
                    totalMeses = ((Global.converteDdMmYyyyParaDateTime(dt_termino).Year - Global.converteDdMmYyyyParaDateTime(dt_inicio).Year) * 12) + (Global.converteDdMmYyyyParaDateTime(dt_termino).Month - Global.converteDdMmYyyyParaDateTime(dt_inicio).Month) + 1;
                    totalMesesProjecao = totalMeses;
                }

                int NumRegistros = datasource.Count;
                DateTime dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
                DateTime dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);

                if (!string.IsNullOrEmpty(fabricante))
                    fabricante = fabricante.Replace("_", ", ");
                else
                    fabricante = "N.I";
                if (!string.IsNullOrEmpty(grupo))
                    grupo = grupo.Replace("|", ", ");
                else
                    grupo = "N.I";
                if (!string.IsNullOrEmpty(subgrupo))
                    subgrupo = subgrupo.Replace("|", ", ");
                else
                    subgrupo = "N.I";
                if (string.IsNullOrEmpty(btu))
                    btu = "N.I";
                if (string.IsNullOrEmpty(ciclo))
                    ciclo = "N.I";
                if (string.IsNullOrEmpty(pos_mercado))
                    pos_mercado = "N.I";
                if (string.IsNullOrEmpty(perc_est_cresc))
                    perc_est_cresc = "0";
                if (string.IsNullOrEmpty(loja))
                    loja = "N.I";

                using (ExcelPackage pck = new ExcelPackage())
                {
                    //Cria uma planilha com nome
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Farol Resumido");

                    #region [ Config Gerais ]
                    //configurações gerais da planilha
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.View.ShowGridLines = false;
                    ws.View.FreezePanes(14, 1);

                    ws.Column(1).Width = 2;
                    ws.Column(COL_FABRICANTE).Width = 6;
                    ws.Column(COL_PRODUTO).Width = 9;
                    ws.Column(COL_DESCRICAO).Width = 51;
                    ws.Column(COL_GRUPO).Width = 7;
                    ws.Column(COL_SUBGRUPO).Width = 10;
                    ws.Column(COL_BTU).Width = 11;
                    ws.Column(COL_CICLO).Width = 7;
                    ws.Column(COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao).Width = 9;
                    ws.Column(COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao).Width = 9;
                    ws.Column(COL_ESTOQUE + totalMeses + totalMesesProjecao).Width = 10;
                    ws.Column(COL_COMPRADO + totalMeses + totalMesesProjecao).Width = 10;
                    ws.Column(COL_SALDO + totalMeses + totalMesesProjecao).Width = 10;

                    ws.Row(1).Height = 1;

                    ws.Cells[LIN_INICIO_REGISTROS, COL_POSICAO_MERCADO + 1, (NumRegistros + LIN_INICIO_REGISTROS + 1), COL_SALDO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0";

                    #endregion                    

                    #region [ Filtro ]
                    lineAux = 2;
                    ws.Cells["B" + lineAux.ToString()].Style.Font.Size = 12;
                    ws.Cells["B" + lineAux.ToString()].Value = "Farol Resumido (v3)";
                    if (opcao_periodo.Equals(FarolV3Controller.COD_CONSULTA_POR_PERIODO_CADASTRO))
                    {
                        sAux = "Período de vendas: ";
                    }
                    else if (opcao_periodo.Equals(FarolV3Controller.COD_CONSULTA_POR_PERIODO_ENTREGA))
                    {
                        sAux = "Período de entrega: ";
                    }
                    else
                    {
                        sAux = "";
                    }
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = sAux + dt_inicio + " a " + dt_termino;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Fabricante(s): " + fabricante;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Grupo(s) de produtos: " + grupo;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Subgrupo(s) de produtos: " + subgrupo;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "BTU/h: " + btu;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Ciclo: " + ciclo;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Posição Mercado:" + pos_mercado;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Loja: " + loja;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Visão: " + visao;
                    lineAux++; ws.Cells["B" + lineAux.ToString()].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    lineAux++; ws.Cells["H" + lineAux.ToString()].Value = "Percentual estimado de crescimento (%): ";
                    ws.Cells["B2:M" + lineAux.ToString()].Style.Font.Bold = true;
                    ws.Cells["H" + lineAux.ToString() + ":I" + lineAux.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    if (visao.Equals("ANALITICA"))
                    {
                        for (int i = 1; i <= totalMesesProjecao; i++)
                        {
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Value = perc_est_cresc;
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.Numberformat.Format = "###.0";
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.Font.Color.SetColor(Color.DarkGreen);
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.Font.Bold = true;
                            ws.Cells[LIN_PERCENTUAL, (COL_CICLO + i + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                    }
                    else
                    {
                        ws.Cells["I12"].Value = perc_est_cresc;
                        ws.Cells["I12"].Style.Numberformat.Format = "###.0";
                        ws.Cells["I12"].Style.Font.Color.SetColor(Color.DarkGreen);
                        ws.Cells["I12"].Style.Font.Bold = true;
                        ws.Cells["I12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    #endregion

                    #region [ Cabeçalho ]
                    using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, COL_VALOR + totalMeses + totalMesesProjecao]])
                    {
                        rng1.Style.WrapText = true;
                        rng1.Style.Font.Bold = true;
                        rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }

                    ws.Cells[LIN_CABECALHO, COL_VALOR + totalMeses + totalMesesProjecao].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    ws.Cells[LIN_CABECALHO, COL_FABRICANTE].Value = "Fabr";
                    ws.Cells[LIN_CABECALHO, COL_PRODUTO].Value = "Produto";
                    ws.Cells[LIN_CABECALHO, COL_DESCRICAO].Value = "Descrição";
                    ws.Cells[LIN_CABECALHO, COL_GRUPO].Value = "Grupo";
                    ws.Cells[LIN_CABECALHO, COL_SUBGRUPO].Value = "Subgrupo";
                    ws.Cells[LIN_CABECALHO, COL_BTU].Value = "BTU/h";
                    ws.Cells[LIN_CABECALHO, COL_CICLO].Value = "Ciclo";

                    for (int i = 1; i <= totalMeses; i++)
                    {

                        string mes = "", ano = "";
                        if (i == 1)
                        {
                            mes = dt1.ToString("MM");
                            ano = dt1.ToString("yyyy");

                        }
                        else
                        {
                            mes = (dt1.AddMonths(i - 1)).ToString("MM");
                            ano = (dt1.AddMonths(i - 1)).ToString("yyyy");
                        }

                        if (i == 1)
                        {
                            //Mês a mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i)));
                            // Projeção mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i + 1).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Value = "Proj " + Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i + i)));

                        }
                        else
                        {
                            //Mês a mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + cont)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i + cont).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + cont)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i + cont].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i + cont), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i + cont)));
                            // Projeção mês
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(COL_CICLO + i + i).Width = 7.7f;
                            ws.Cells[LIN_CABECALHO, (COL_CICLO + i + i)].Value = "Proj " + Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + i + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (7 + i + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (7 + i + i)));
                        }

                        cont++;
                    }

                    ws.Cells[LIN_CABECALHO, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Value = "Total";
                    ws.Cells[LIN_CABECALHO, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Value = "Proj Total";
                    ws.Cells[LIN_CABECALHO, COL_ESTOQUE + totalMeses + totalMesesProjecao].Value = "Estoque";
                    ws.Cells[LIN_CABECALHO, COL_COMPRADO + totalMeses + totalMesesProjecao].Value = "Compra";
                    ws.Cells[LIN_CABECALHO, COL_SALDO + totalMeses + totalMesesProjecao].Value = "Saldo";
                    ws.Cells[LIN_CABECALHO, COL_CUSTO + totalMeses + totalMesesProjecao].Value = "Custo";
                    ws.Cells[LIN_CABECALHO, COL_IDEIA + totalMeses + totalMesesProjecao].Value = "Ideia";
                    ws.Cells[LIN_CABECALHO, COL_VALOR + totalMeses + totalMesesProjecao].Value = "Valor";
                    ws.Cells[LIN_CABECALHO, COL_FABRICANTE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_PRODUTO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_DESCRICAO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_GRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_SUBGRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_BTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_CICLO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_ESTOQUE + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_COMPRADO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_SALDO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_CUSTO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_IDEIA + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[LIN_CABECALHO, COL_VALOR + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    #endregion

                    #region [ Registros ]
                    using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, COL_VALOR + totalMeses + totalMesesProjecao]])
                    {
                        rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    for (int i = 0; i < NumRegistros; i++)
                    {
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_FABRICANTE].Value = datasource.ElementAt(i).Fabricante;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_FABRICANTE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_PRODUTO].Value = datasource.ElementAt(i).Produto;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_PRODUTO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_DESCRICAO].Value = datasource.ElementAt(i).Descricao;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_DESCRICAO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_GRUPO].Value = datasource.ElementAt(i).Grupo;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_GRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SUBGRUPO].Value = datasource.ElementAt(i).Subgrupo;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SUBGRUPO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        if (datasource.ElementAt(i).Potencia_BTU != 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_BTU].Value = datasource.ElementAt(i).Potencia_BTU;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_BTU].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_BTU].Style.Numberformat.Format = "##,###";
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CICLO].Value = datasource.ElementAt(i).Ciclo;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CICLO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        if (visao.Equals("ANALITICA"))
                        {
                            cont = 0;
                            for (int j = 0; j < totalMeses; j++)
                            {
                                if (j == 0)
                                {
                                    //Registros mês a mês                                   
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].Value = datasource.ElementAt(i).Meses[j];
                                    //Registros mês a mês Projeção 
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].FormulaR1C1 = string.Format("RC[-1]+(R{0}C{1}/100)*RC[-1]", LIN_PERCENTUAL, (COL_CICLO + j + 2));
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].Style.Numberformat.Format = "##,##0";
                                    //Registros Projeção Total e Total Vendas
                                    totalVenda = ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1)].ToString();
                                    totalProjecao = ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2)].ToString();
                                }
                                else
                                {
                                    //Registros mês a mês
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1 + cont)].Value = datasource.ElementAt(i).Meses[j];
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].Value = datasource.ElementAt(i).Meses[j];
                                    //Registros mês a mês Projeção 
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].FormulaR1C1 = string.Format("RC[-1]+(R{0}C{1}/100)*RC[-1]", LIN_PERCENTUAL, (COL_CICLO + j + 2 + cont));
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].Style.Numberformat.Format = "##,##0";

                                    //Registros Projeção Total Total Vendas
                                    totalVenda = totalVenda + ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 1 + cont)].ToString();
                                    totalProjecao = totalProjecao + ws.Cells[i + LIN_INICIO_REGISTROS, (COL_CICLO + j + 2 + j)].ToString();

                                }
                                if (totalMesesProjecao > j + 1)
                                {
                                    totalVenda = totalVenda + "+";
                                    totalProjecao = totalProjecao + "+";
                                }
                                cont++;
                            }
                        }

                        if (visao.Equals("ANALITICA"))
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Formula = string.Format("=SUM({0})", totalVenda);
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Formula = string.Format("=SUM({0})", totalProjecao);
                        }
                        else
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses].Value = datasource.ElementAt(i).Qtde_vendida;
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses].Formula = string.Format("{0}+(I12/100)*{0}", ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses]);
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses].Style.Numberformat.Format = "##,##0";
                        }
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (datasource.ElementAt(i).Qtde_vendida > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0";
                        if (datasource.ElementAt(i).Qtde_vendida > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao].Value = datasource.ElementAt(i).Qtde_estoque_venda;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (datasource.ElementAt(i).Qtde_estoque_venda > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao].Value = datasource.ElementAt(i).Farol_qtde_comprada;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (datasource.ElementAt(i).Farol_qtde_comprada > 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Formula = string.Format("{0}+{1}-{2}", ws.Cells[i + LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao], ws.Cells[i + LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao], ws.Cells[i + LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao]);
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0;[Red]-##,##0";
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_IDEIA + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_IDEIA + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CUSTO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "#,###.00;[Red]-#,###.00";
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CUSTO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CUSTO + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VALOR + totalMeses + totalMesesProjecao].Formula = string.Format("{0}*{1}", ws.Cells[i + LIN_INICIO_REGISTROS, COL_CUSTO + totalMeses + totalMesesProjecao], ws.Cells[i + LIN_INICIO_REGISTROS, COL_IDEIA + totalMeses + totalMesesProjecao]);
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VALOR + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_VALOR + totalMeses + totalMesesProjecao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        if (datasource.ElementAt(i).Saldo < 0) ws.Cells[i + LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao].Style.Font.Color.SetColor(Color.Red);

                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CUSTO + totalMeses + totalMesesProjecao].Value = Global.formataMoeda(datasource.ElementAt(i).Custo);

                    }
                    #endregion


                    #region [ Total ]
                    ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), COL_FABRICANTE, (NumRegistros + LIN_INICIO_REGISTROS + 1), COL_SALDO + totalMeses + totalMesesProjecao].Style.Font.Bold = true;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7].Value = "TOTAL";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_VENDA_HISTORICO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VENDA_PREVISAO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_ESTOQUE + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_ESTOQUE + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_ESTOQUE + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_COMPRADO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_COMPRADO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_COMPRADO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_SALDO + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_SALDO + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_SALDO + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VALOR + totalMeses + totalMesesProjecao].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, COL_VALOR + totalMeses + totalMesesProjecao, NumRegistros + LIN_INICIO_REGISTROS - 1, COL_VALOR + totalMeses + totalMesesProjecao));
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_SALDO + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "##,##0;[Red]-##,##0";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_VALOR + totalMeses + totalMesesProjecao].Style.Numberformat.Format = "#,###.00;[Red]-#,###.00";
                    #endregion

                    pck.SaveAs(new FileInfo(filePath));
                }
            });
        }
        #endregion
    }
}