using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ART3WebAPI.Models.Entities;
using System;
using System.Globalization;
using System.Threading;

namespace ART3WebAPI.Models.Domains
{
    public class DevolucaoProdutos2GeradorRelatorio
    {
        #region [ Constantes ]
        private const int LIN_CABECALHO = 13;
        private const int LIN_INICIO_REGISTROS = 14;
        private const int COL_DATA_PEDIDO = 2;
        private const int COL_DATA_DEVOLUCAO = 3;
        private const int COL_DATA_BAIXA = 4;
        private const int COL_QTDE = 5;
        private const int COL_PRODUTO = 6;
        private const int COL_PEDIDO = 7;
        private const int COL_CLIENTE = 8;
        private const int COL_VENDEDOR = 9;
        private const int COL_PARCEIRO = 10;
        private const int COL_MOTIVO = 11;
        #endregion

        public static Task GeraXLS(List<DevolucaoProduto2Entity> lista, string filePath, string dt_devolucao_inicio, string dt_devolucao_termino, string fabricante, string produto, string pedido, string vendedor, string indicador, string captador, string lojas)
        {
            return Task.Run(() =>
            {
                #region [ Declarações ]
                int totalRegistros = lista.Count;
                int totalProdutos = 0;
                int i = 0;
                CultureInfo ci = new CultureInfo("pt-BR");
                #endregion

                if (string.IsNullOrEmpty(fabricante)) fabricante = "todos";
                if (string.IsNullOrEmpty(produto)) produto = "todos";
                if (string.IsNullOrEmpty(pedido)) pedido = "todos";
                if (string.IsNullOrEmpty(vendedor)) vendedor = "todos";
                if (string.IsNullOrEmpty(indicador)) indicador = "todos";
                if (string.IsNullOrEmpty(captador)) captador = "todos";
                if (string.IsNullOrEmpty(lojas))
                    lojas = "todas";
                else
                {
                    lojas = lojas.Replace("_", ", ");
                    lojas = lojas.Replace("-", " a ");
                }

                using (ExcelPackage package = new ExcelPackage())
                {
                    // Cria uma planilha e dá um nome à ela
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("Devolução de Produtos II");

                    #region [ Configurações gerais da planilha ]
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.View.ShowGridLines = false;
                    ws.View.FreezePanes(LIN_INICIO_REGISTROS, 1);
                    ws.Column(1).Width = 2;
                    ws.Column(COL_DATA_PEDIDO).Width = 13;
                    ws.Column(COL_DATA_DEVOLUCAO).Width = 13;
                    ws.Column(COL_DATA_BAIXA).Width = 13;
                    ws.Column(COL_QTDE).Width = 5;
                    ws.Column(COL_PRODUTO).Width = 50;
                    ws.Column(COL_PEDIDO).Width = 12;
                    ws.Column(COL_CLIENTE).Width = 50;
                    ws.Column(COL_VENDEDOR).Width = 15;
                    ws.Column(COL_PARCEIRO).Width = 15;
                    ws.Column(COL_MOTIVO).Width = 90;
                    ws.Row(1).Height = 1;
                    #endregion

                    #region [ Filtro ]
                    ws.Cells["B2:M11"].Style.Font.Bold = true;
                    ws.Cells["B2"].Style.Font.Size = 12;
                    ws.Cells["B2"].Value = "Devolução de Produtos II";
                    ws.Cells["B3"].Value = "Devolvido entre: " + dt_devolucao_inicio + " e " + dt_devolucao_termino;
                    ws.Cells["B4"].Value = "Fabricante: " + fabricante;
                    ws.Cells["B5"].Value = "Produto: " + produto;
                    ws.Cells["B6"].Value = "Pedido: " + pedido;
                    ws.Cells["B7"].Value = "Vendedor: " + vendedor;
                    ws.Cells["B8"].Value = "Indicador: " + indicador;
                    ws.Cells["B9"].Value = "Captador: " + captador;
                    ws.Cells["B10"].Value = "Loja(s): " + lojas;
                    ws.Cells["B11"].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    #endregion
                    
                    #region [ Cabeçalho ]

                    #region [ Configura o estilo do cabeçalho ]
                    using (ExcelRange range = ws.Cells[FromRow: LIN_CABECALHO, FromCol: COL_DATA_PEDIDO, ToRow: LIN_CABECALHO, ToCol: COL_MOTIVO])
                    {
                        range.Style.WrapText = true;
                        range.Style.Font.Bold = true;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }
                    ws.Cells[LIN_CABECALHO, COL_MOTIVO].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    ws.Cells[LIN_CABECALHO, COL_DATA_PEDIDO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_DATA_DEVOLUCAO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_DATA_BAIXA].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[LIN_CABECALHO, COL_QTDE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    #endregion

                    ws.Cells[LIN_CABECALHO, COL_DATA_PEDIDO].Value = "DT Pedido";
                    ws.Cells[LIN_CABECALHO, COL_DATA_DEVOLUCAO].Value = "DT Devol";
                    ws.Cells[LIN_CABECALHO, COL_DATA_BAIXA].Value = "DT Baixa";
                    ws.Cells[LIN_CABECALHO, COL_QTDE].Value = "Qtde";
                    ws.Cells[LIN_CABECALHO, COL_PRODUTO].Value = "Produto";
                    ws.Cells[LIN_CABECALHO, COL_PEDIDO].Value = "Pedido";
                    ws.Cells[LIN_CABECALHO, COL_CLIENTE].Value = "Cliente";
                    ws.Cells[LIN_CABECALHO, COL_VENDEDOR].Value = "Vendedor";
                    ws.Cells[LIN_CABECALHO, COL_PARCEIRO].Value = "Parceiro";
                    ws.Cells[LIN_CABECALHO, COL_MOTIVO].Value = "Motivo";
                    #endregion

                    #region [ Registros ]

                    using (ExcelRange range = ws.Cells[FromRow: LIN_INICIO_REGISTROS, FromCol: COL_DATA_PEDIDO, ToRow: LIN_INICIO_REGISTROS + totalRegistros, ToCol: COL_MOTIVO])
                    {
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    foreach (DevolucaoProduto2Entity item in lista)
                    {
                        #region [ Data Pedido ]
                        if (item.DataPedido != DateTime.MinValue)
                        {
                            ws.Cells[LIN_INICIO_REGISTROS + i, COL_DATA_PEDIDO].Value = Global.formataDataDdMmYyyyComSeparador(item.DataPedido);
                            ws.Cells[LIN_INICIO_REGISTROS + i, COL_DATA_PEDIDO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        #endregion

                        #region [ Data Devolução ]
                        if (item.DataDevolvido != DateTime.MinValue)
                        {
                            ws.Cells[LIN_INICIO_REGISTROS + i, COL_DATA_DEVOLUCAO].Value = Global.formataDataDdMmYyyyComSeparador(item.DataDevolvido);
                            ws.Cells[LIN_INICIO_REGISTROS + i, COL_DATA_DEVOLUCAO].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        #endregion

                        #region [ Data Baixa ]
                        if (item.DataBaixa != DateTime.MinValue)
                        {
                            ws.Cells[LIN_INICIO_REGISTROS + i, COL_DATA_BAIXA].Value = Global.formataDataDdMmYyyyComSeparador(item.DataBaixa);
                            ws.Cells[LIN_INICIO_REGISTROS + i, COL_DATA_BAIXA].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        #endregion

                        #region [ Qtde ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_QTDE].Value = item.Qtde;
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_QTDE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_QTDE].Style.Font.Bold = true;
                        totalProdutos += item.Qtde;
                        #endregion

                        #region [ Produto ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_PRODUTO].Value = "(" + item.Fabricante + ")" +
                        item.Produto + " - " + item.Descricao;
                        #endregion

                        #region [ Pedido ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_PEDIDO].Value = item.Pedido;
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_PEDIDO].Style.Font.Bold = true;
                        #endregion

                        #region [ Cliente ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_CLIENTE].Value = item.Cliente;
                        #endregion

                        #region [ Vendedor ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_VENDEDOR].Value = item.Vendedor;
                        #endregion

                        #region [ Parceiro ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_PARCEIRO].Value = item.Indicador;
                        #endregion

                        #region [ Motivo ]
                        ws.Cells[LIN_INICIO_REGISTROS + i, COL_MOTIVO].Value = item.Motivo;
                        #endregion

                        i++;
                    }
                    #endregion

                    #region [ Total produtos ]
                    ws.Cells[(LIN_INICIO_REGISTROS + totalRegistros + 1), COL_DATA_BAIXA].Style.Font.Bold = true;
                    ws.Cells[(LIN_INICIO_REGISTROS + totalRegistros + 1), COL_DATA_BAIXA].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[(LIN_INICIO_REGISTROS + totalRegistros + 1), COL_DATA_BAIXA].Value = "TOTAL:";
                    ws.Cells[(LIN_INICIO_REGISTROS + totalRegistros + 1), COL_QTDE].Style.Font.Bold = true;
                    ws.Cells[(LIN_INICIO_REGISTROS + totalRegistros + 1), COL_QTDE].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[(LIN_INICIO_REGISTROS + totalRegistros + 1), COL_QTDE].Value = totalProdutos;
                    #endregion

                    package.SaveAs(new FileInfo(filePath));
                }
            });
        }
    }
}