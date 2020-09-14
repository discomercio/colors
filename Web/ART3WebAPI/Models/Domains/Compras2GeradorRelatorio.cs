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

        #region [ Constantes ]
        private const int LIN_INICIO_REGISTROS = 17;
        private const int LIN_CABECALHO = 16; 
        #endregion
        

        public static Task GenerateXLS(List<Compras> datasource, string filePath, string dt_inicio, string dt_termino, string fabricante,string produto, string grupo, string subgrupo, string btu, string ciclo, string pos_mercado, string nf, string dt_nf_inicio, string dt_nf_termino, string visao, string detalhamento)
        {
            return Task.Run(() =>
            {

                int totalMeses = 0;
                if (visao.Equals("ANALITICA"))
                {
                    if (detalhamento != "SINTETICO_NF")
                    {
                        if (!string.IsNullOrEmpty(dt_inicio))
                        {
                            totalMeses = ((Global.converteDdMmYyyyParaDateTime(dt_termino).Year - Global.converteDdMmYyyyParaDateTime(dt_inicio).Year) * 12) + (Global.converteDdMmYyyyParaDateTime(dt_termino).Month - Global.converteDdMmYyyyParaDateTime(dt_inicio).Month) + 1;
                        }
                        else
                        {
                            totalMeses = ((Global.converteDdMmYyyyParaDateTime(dt_nf_termino).Year - Global.converteDdMmYyyyParaDateTime(dt_nf_inicio).Year) * 12) + (Global.converteDdMmYyyyParaDateTime(dt_nf_termino).Month - Global.converteDdMmYyyyParaDateTime(dt_nf_inicio).Month) + 1;
                        }
                    }
                }

                int NumRegistros = datasource.Count;
                //DateTime dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
                //DateTime dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);
                DateTime dt1, dt2;
                if (!string.IsNullOrEmpty(dt_inicio))
                {
                    dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
                    dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);
                }
                else
                {
                    dt1 = Global.converteDdMmYyyyParaDateTime(dt_nf_inicio);
                    dt2 = Global.converteDdMmYyyyParaDateTime(dt_nf_termino);
                }
            


                string periodoentrada = "";
                if (!string.IsNullOrEmpty(dt_inicio))
                    periodoentrada = dt_inicio;
                if (!string.IsNullOrEmpty(dt_termino))
                {
                    if (!string.IsNullOrEmpty(periodoentrada))
                        periodoentrada = "de " + periodoentrada + " a ";
                    periodoentrada = periodoentrada + dt_termino;
                }
                if (periodoentrada == "")
                    periodoentrada = "N.I";

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
                if (string.IsNullOrEmpty(produto))
                    produto = "N.I";
                if (string.IsNullOrEmpty(btu))
                    btu = "N.I";
                if (string.IsNullOrEmpty(ciclo))
                    ciclo = "N.I";
                if (string.IsNullOrEmpty(pos_mercado))
                    pos_mercado = "N.I";
                if (string.IsNullOrEmpty(nf))
                    nf = "N.I";
                string emissaoNF = "";
                if (!string.IsNullOrEmpty(dt_nf_inicio))
                    emissaoNF = dt_nf_inicio;
                if (!string.IsNullOrEmpty(dt_nf_termino))
                {
                    if (!string.IsNullOrEmpty(emissaoNF))
                            emissaoNF = "de " + emissaoNF + " a ";
                    emissaoNF = emissaoNF + dt_nf_termino;
                }
                if (emissaoNF == "")
                    emissaoNF = "N.I";

                using (ExcelPackage pck = new ExcelPackage())
                {
                    //Cria uma planilha com nome
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Compras 2");

                    #region [ Config Gerais ]
                    //configurações gerais da planilha
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.View.ShowGridLines = false;
                    ws.View.FreezePanes(16, 1);
                    ws.Column(1).Width = 2;
                    if (detalhamento == "SINTETICO_FABR")
                    {
                        ws.Column(2).Width = 6;
                        ws.Column(3).Width = 51;
                        ws.Column(4).Width = 11;
                        ws.Column(4 + totalMeses).Width = 15;
                        ws.Column(5 + totalMeses).Width = 15;
                    }
                    else
                    {
                        if (detalhamento == "SINTETICO_NF")
                        {
                            ws.Column(2).Width = 12;
                            ws.Column(3).Width = 58;
                            ws.Column(4 + totalMeses).Width = 11;
                            ws.Column(5 + totalMeses).Width = 15;
                        }
                        else
                        {
                            ws.Column(2).Width = 6;
                            ws.Column(3).Width = 9;
                            ws.Column(4).Width = 51;
                            ws.Column(5).Width = 11;
                            ws.Column(6 + totalMeses).Width = 11;
                            ws.Column(7 + totalMeses).Width = 15;
                        }
                    }

                    ws.Row(1).Height = 1;

                    #endregion


                    #region [ Filtro ]
                    ws.Cells["B2:M12"].Style.Font.Bold = true;
                    ws.Cells["B2"].Style.Font.Size = 12;
                    ws.Cells["B2"].Value = "Compras II";
                    //ws.Cells["B3"].Value = "Período: " + dt_inicio + " a " + dt_termino;
                    ws.Cells["B3"].Value = "Período: " + periodoentrada;
                    ws.Cells["B4"].Value = "Fabricante(s): " + fabricante;
                    ws.Cells["B5"].Value = "Grupo(s) de produtos: " + grupo;
                    ws.Cells["B6"].Value = "Subgrupo(s) de produtos: " + subgrupo;
                    ws.Cells["B7"].Value = "Produto: " + produto;
                    ws.Cells["B8"].Value = "BTU/h: " + btu;
                    ws.Cells["B9"].Value = "Ciclo: " + ciclo;
                    ws.Cells["B10"].Value = "Posição Mercado: " + pos_mercado;
                    ws.Cells["B11"].Value = "Nº Nota Fiscal: " + nf;
                    ws.Cells["B12"].Value = "Emissão NF Entrada: " + emissaoNF;
                    ws.Cells["B13"].Value = "Tipo de Detalhamento: " + Global.getDetalhamento(detalhamento);
                    ws.Cells["B14"].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    ws.Cells["L15:M15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    #endregion

                    #region [ Cabeçalho ]

                    #region [ Sintético por NF ]
                    if (detalhamento == "SINTETICO_NF")
                    {
                        using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, 5 + totalMeses]])
                        {
                            rng1.Style.WrapText = true;
                            rng1.Style.Font.Bold = true;
                            rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                            rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        
                        ws.Cells[LIN_CABECALHO, 2].Value = "Nº NF";
                        ws.Cells[LIN_CABECALHO, 3].Value = "Fabr";
                        ws.Cells[LIN_CABECALHO, 4 + totalMeses].Value = "Qtde Total";
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Value = "Valor Total";
                        ws.Cells[LIN_CABECALHO, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 4 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    #endregion

                    #region [ Sintético por Fabricante ]
                    if (detalhamento == "SINTETICO_FABR")
                    {

                        using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, 4 + totalMeses]])
                        {
                            rng1.Style.WrapText = true;
                            rng1.Style.Font.Bold = true;
                            rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                            rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                        ws.Cells[LIN_CABECALHO, 4 + totalMeses].Style.Border.Right.Style = ExcelBorderStyle.Medium;

                        ws.Cells[LIN_CABECALHO, 2].Value = "Fabr";
                        ws.Cells[LIN_CABECALHO, 3].Value = "Descrição";
                        ws.Cells[LIN_CABECALHO, 4 + totalMeses].Value = "Valor Total";
                        ws.Cells[LIN_CABECALHO, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 4 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

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

                            ws.Cells[LIN_CABECALHO, (3 + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(3 + i).Width = 15;
                            ws.Cells[LIN_CABECALHO, (3 + i)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 3 + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (3 + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (3 + i)));
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 3 + i].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
                        }
                    } 
                    #endregion

                    #region [ Sintético por Produto ]
                    else if (detalhamento == "SINTETICO_PROD")
                    {

                        using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, 5 + totalMeses]])
                        {
                            rng1.Style.WrapText = true;
                            rng1.Style.Font.Bold = true;
                            rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                            rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Style.Border.Right.Style = ExcelBorderStyle.Medium;

                        ws.Cells[LIN_CABECALHO, 2].Value = "Fabr";
                        ws.Cells[LIN_CABECALHO, 3].Value = "Produto";
                        ws.Cells[LIN_CABECALHO, 4].Value = "Descrição";
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Value = " Qtde Total";
                        ws.Cells[LIN_CABECALHO, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 5 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


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

                            ws.Cells[LIN_CABECALHO, (4 + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(4 + i).Width = 8;
                            ws.Cells[LIN_CABECALHO, (4 + i)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4 + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (4 + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (4 + i)));

                        }

                    } 
                    #endregion

                    #region [ Valor Referência Médio ]
                    else if (detalhamento == "CUSTO_MEDIO")
                    {

                        using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, 7 + totalMeses]])
                        {
                            rng1.Style.WrapText = true;
                            rng1.Style.Font.Bold = true;
                            rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                            rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                        ws.Cells[LIN_CABECALHO, 7 + totalMeses].Style.Border.Right.Style = ExcelBorderStyle.Medium;


                        ws.Cells[LIN_CABECALHO, 2].Value = "Fabr";
                        ws.Cells[LIN_CABECALHO, 3].Value = "Produto";
                        ws.Cells[LIN_CABECALHO, 4].Value = "Descrição";
                        ws.Cells[LIN_CABECALHO, 5].Value = "Referência Médio";
                        ws.Cells[LIN_CABECALHO, 6 + totalMeses].Value = "Total";
                        ws.Cells[LIN_CABECALHO, 7 + totalMeses].Value = "Valor Total";
                        ws.Cells[LIN_CABECALHO, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 6 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 7 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


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

                            ws.Cells[LIN_CABECALHO, (5 + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(5 + i).Width = 8;
                            ws.Cells[LIN_CABECALHO, (5 + i)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 5 + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (5 + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (5 + i)));

                        }

                    } 
                    #endregion

                    #region [ Valor Referência Individual ]
                    else if (detalhamento == "CUSTO_INDIVIDUAL")
                    {

                        using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, 7 + totalMeses]])
                        {
                            rng1.Style.WrapText = true;
                            rng1.Style.Font.Bold = true;
                            rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                            rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                        ws.Cells[LIN_CABECALHO, 7 + totalMeses].Style.Border.Right.Style = ExcelBorderStyle.Medium;



                        ws.Cells[LIN_CABECALHO, 2].Value = "Fabr";
                        ws.Cells[LIN_CABECALHO, 3].Value = "Produto";
                        ws.Cells[LIN_CABECALHO, 4].Value = "Descrição";
                        ws.Cells[LIN_CABECALHO, 5].Value = "Referência Individual";
                        ws.Cells[LIN_CABECALHO, 6 + totalMeses].Value = "Total";
                        ws.Cells[LIN_CABECALHO, 7 + totalMeses].Value = "Valor Total";
                        ws.Cells[LIN_CABECALHO, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[LIN_CABECALHO, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 6 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_CABECALHO, 7 + totalMeses].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;



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

                            ws.Cells[LIN_CABECALHO, (5 + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Column(5 + i).Width = 8;
                            ws.Cells[LIN_CABECALHO, (5 + i)].Value = Global.mesPorExtenso(int.Parse(mes)) + "/" + ano.Substring(2, 2);
                            //Total
                            ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 5 + i].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, (5 + i), NumRegistros + LIN_INICIO_REGISTROS - 1, (5 + i)));

                        }


                    }
                    #endregion

                    #endregion

                    #region [ Registros ]

                    #region [ Sintético por NF ]
                    if (detalhamento == "SINTETICO_NF")
                    {
                        using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, 5 + totalMeses]])
                        {
                            rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        for (int i = 0; i < NumRegistros; i++)
                        {

                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Value = datasource.ElementAt(i).NF;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Value = datasource.ElementAt(i).Fabricante + " - " + Global.getDescFabricante(datasource.ElementAt(i).Fabricante);
                            if (visao.Equals("ANALITICA"))
                            {
                                if (detalhamento != "SINTETICO_NF")
                                {
                                    for (int j = 0; j < totalMeses; j++)
                                    {
                                        ws.Cells[i + LIN_INICIO_REGISTROS, (3 + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                        ws.Cells[i + LIN_INICIO_REGISTROS, (3 + j + 1)].Style.Numberformat.Format = "###,###,##0.00";
                                    } 
                                }
                            }
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4 + totalMeses].Value = datasource.ElementAt(i).Qtde;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5 + totalMeses].Value = datasource.ElementAt(i).Valor;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5 + totalMeses].Style.Numberformat.Format = "###,###,##0.00";

                        }
                        #region [ Total ]

                        ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), 2, (NumRegistros + LIN_INICIO_REGISTROS + 1), 5 + totalMeses].Style.Font.Bold = true;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 3].Value = "TOTAL";
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 4 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 4 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 5 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 5 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 5 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 5 + totalMeses].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
                        #endregion



                    }
                    #endregion

                    #region [ Sintético por Fabricante ]
                    if (detalhamento == "SINTETICO_FABR")
                    {
                        using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, 4 + totalMeses]])
                        {
                            rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        for (int i = 0; i < NumRegistros; i++)
                        {
                            
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Value = datasource.ElementAt(i).Fabricante;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Value = Global.getDescFabricante(datasource.ElementAt(i).Fabricante);
                            if (visao.Equals("ANALITICA"))
                            {
                                for (int j = 0; j < totalMeses; j++)
                                {
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (3 + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (3 + j + 1)].Style.Numberformat.Format = "###,###,##0.00";
                                }
                            }
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4 + totalMeses].Value = datasource.ElementAt(i).Valor;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4 + totalMeses].Style.Numberformat.Format = "###,###,##0.00";

                        }
                        #region [ Total ]

                        ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), 2, (NumRegistros + LIN_INICIO_REGISTROS + 1), 5 + totalMeses].Style.Font.Bold = true;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 3].Value = "TOTAL";
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 4 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 4 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4 + totalMeses].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
                        #endregion



                    } 
                    #endregion

                    #region [ Sintético por Produto ]
                    else if (detalhamento == "SINTETICO_PROD")
                    {
                        using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, 5 + totalMeses]])
                        {
                            rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        for (int i = 0; i < NumRegistros; i++)
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Value = datasource.ElementAt(i).Fabricante;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Value = datasource.ElementAt(i).Produto;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4].Value = Global.getDescProduto(datasource.ElementAt(i).Produto);
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            if (visao.Equals("ANALITICA"))
                            {
                                for (int j = 0; j < totalMeses; j++)
                                {
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (4 + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                }
                            }
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5 + totalMeses].Value = datasource.ElementAt(i).Qtde;                         

                        }

                        #region [ Total ]

                        ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), 2, (NumRegistros + LIN_INICIO_REGISTROS + 1), 5 + totalMeses].Style.Font.Bold = true;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4].Value = "TOTAL";
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 4 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 4 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 5 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 5 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 5 + totalMeses));
                        #endregion

                    } 
                    #endregion

                    #region [ Valor Referência Médio ]
                    else if (detalhamento == "CUSTO_MEDIO")
                    {
                        using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, 7 + totalMeses]])
                        {
                            rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        for (int i = 0; i < NumRegistros; i++)
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Value = datasource.ElementAt(i).Fabricante;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Value = datasource.ElementAt(i).Produto;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4].Value = Global.getDescProduto(datasource.ElementAt(i).Produto);
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5].Value = datasource.ElementAt(i).Valor / datasource.ElementAt(i).Qtde;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5].Style.Numberformat.Format = "###,###,##0.00";
                            ws.Cells[i + LIN_INICIO_REGISTROS, 6].Value = datasource.ElementAt(i).Qtde;


                            if (visao.Equals("ANALITICA"))
                            {
                                for (int j = 0; j < totalMeses; j++)
                                {
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (5 + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                    
                                }
                            }
                            ws.Cells[i + LIN_INICIO_REGISTROS, 6 + totalMeses].Value = datasource.ElementAt(i).Qtde;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 7 + totalMeses].Value = datasource.ElementAt(i).Valor;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 7 + totalMeses].Style.Numberformat.Format = "###,###,##0.00";

                        }

                        #region [ Total ]
                        ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), 2, (NumRegistros + LIN_INICIO_REGISTROS + 1), 7 + totalMeses].Style.Font.Bold = true;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4].Value = "TOTAL";
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 6 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 6 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 6 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 7 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 7 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + totalMeses].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
                        #endregion


                    } 
                    #endregion

                    #region [ Valor Referência Individual ]
                    else if (detalhamento == "CUSTO_INDIVIDUAL")
                    {
                        using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, 7 + totalMeses]])
                        {
                            rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }

                        for (int i = 0; i < NumRegistros; i++)
                        {
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Value = datasource.ElementAt(i).Fabricante;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Value = datasource.ElementAt(i).Produto;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4].Value = Global.getDescProduto(datasource.ElementAt(i).Produto);
                            ws.Cells[i + LIN_INICIO_REGISTROS, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5].Value = datasource.ElementAt(i).Valor;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 5].Style.Numberformat.Format = "###,###,##0.00";
                            if (visao.Equals("ANALITICA"))
                            {
                                for (int j = 0; j < totalMeses; j++)
                                {
                                    ws.Cells[i + LIN_INICIO_REGISTROS, (5 + j + 1)].Value = datasource.ElementAt(i).Meses[j];
                                   
                                }
                            }
                            ws.Cells[i + LIN_INICIO_REGISTROS, 6 + totalMeses].Value = datasource.ElementAt(i).Qtde;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 7 + totalMeses].Value = datasource.ElementAt(i).Valor * datasource.ElementAt(i).Qtde;
                            ws.Cells[i + LIN_INICIO_REGISTROS, 7 + totalMeses].Style.Numberformat.Format = "###,###,##0.00";

                           
                          
                        }

                        #region [ Total ]
                        ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), 2, (NumRegistros + LIN_INICIO_REGISTROS + 1), 7 + totalMeses].Style.Font.Bold = true;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4].Value = "TOTAL";
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 6 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 6 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 6 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + totalMeses].Formula = string.Format("SUM({0})", new ExcelAddress(LIN_INICIO_REGISTROS, 7 + totalMeses, NumRegistros + LIN_INICIO_REGISTROS - 1, 7 + totalMeses));
                        ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, 7 + totalMeses].Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
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