using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using ART3WebAPI.Models.Domains;
using System;

namespace ART3WebAPI.Models.Domains
{
    public class EstOcorrenciasGeradorRelatorio
    {
        #region [ Constantes ]
        private const int LIN_INICIO_REGISTROS = 14;
        private const int LIN_CABECALHO = 13;
        private const int COL_Pedido = 2;
        private const int COL_NF = 3;
        private const int COL_Transportadora = 4;
        private const int COL_Ocorrencia = 5;
        private const int COL_TipoOcorrencia = 6;
        #endregion

        public static Task GenerateXLS(List<Ocorrencias> datasource, string filePath, string dt_inicio, string dt_termino, string motivo_ocorrencia, string tp_ocorrencia, string transportadora, string vendedor, string indicador, string UF, string loja)
        {
            return Task.Run(() =>
            {
                int NumRegistros = datasource.Count;
                DateTime dt1 = Global.converteDdMmYyyyParaDateTime(dt_inicio);
                DateTime dt2 = Global.converteDdMmYyyyParaDateTime(dt_termino);

                if (!string.IsNullOrEmpty(motivo_ocorrencia))
                    motivo_ocorrencia = BD.obtem_descricao_tabela_t_codigo_descricao(Global.Cte.GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA, motivo_ocorrencia);
                else
                    motivo_ocorrencia = "Todos";
                if (string.IsNullOrEmpty(UF))
                    UF = "Todos";
                if (string.IsNullOrEmpty(tp_ocorrencia))
                    tp_ocorrencia = "Todos";
                else
                    tp_ocorrencia = BD.obtem_descricao_tabela_t_codigo_descricao(Global.Cte.GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA, tp_ocorrencia);
                if ((string.IsNullOrEmpty(transportadora)) || (transportadora == "0"))
                    transportadora = "Todos";
                if (string.IsNullOrEmpty(vendedor))
                    vendedor = "Todos";
                if (string.IsNullOrEmpty(indicador))
                    indicador = "Todos";
                if (string.IsNullOrEmpty(loja))
                    loja = "Todas";
                string[] v_loja = loja.Split('_');
                loja = "";
                for (int i = 0; i < v_loja.Length; i++)
                {
                    loja = loja + v_loja[i];
                    if (i != v_loja.Length - 1)
                    {
                        loja = loja + ", ";
                    }
                }
                using (ExcelPackage pck = new ExcelPackage())
                {
                    //Cria uma planilha com nome
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Estatísticas de Ocorrências ");

                    #region [ Config Gerais ]
                    //configurações gerais da planilha
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.View.ShowGridLines = false;


                    ws.Column(1).Width = 2;
                    ws.Column(COL_Pedido).Width = 15;
                    ws.Column(COL_NF).Width = 15;
                    ws.Column(COL_Transportadora).Width = 20;
                    ws.Column(COL_Ocorrencia).Width = 40;
                    ws.Column(COL_TipoOcorrencia).Width = 40;
 
                    ws.Row(1).Height = 1;                
                    #endregion


                    #region [ Filtro ]
                    ws.Cells["B2:M12"].Style.Font.Bold = true;
                    ws.Cells["B2"].Style.Font.Size = 12;
                    ws.Cells["B2"].Value = "Estatísticas de Ocorrências ";
                    ws.Cells["B3"].Value = "Período da Ocorrência: " + dt_inicio + " a " + dt_termino;
                    ws.Cells["B4"].Value = "Motivo Abertura: " + motivo_ocorrencia;
                    ws.Cells["B5"].Value = "Tipo de Ocorrência: " + tp_ocorrencia;
                    ws.Cells["B6"].Value = "Transportadora: " + transportadora;
                    ws.Cells["B7"].Value = "Vendedor: " + vendedor;
                    ws.Cells["B8"].Value = "Indicador: " + indicador;
                    ws.Cells["B9"].Value = "UF: " + UF;                    
                    ws.Cells["B10"].Value = "Loja(s): " + loja;
                    ws.Cells["B11"].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    #endregion

                    #region [ Cabeçalho ]
                    using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, COL_TipoOcorrencia]])
                    {
                        rng1.Style.WrapText = true;
                        rng1.Style.Font.Bold = true;
                        rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }

                    ws.Cells[LIN_CABECALHO, COL_TipoOcorrencia].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    ws.Cells[LIN_CABECALHO, COL_Pedido].Value = "Pedido";
                    ws.Cells[LIN_CABECALHO, COL_NF].Value = "NF";
                    ws.Cells[LIN_CABECALHO, COL_Transportadora].Value = "Transportadora";
                    ws.Cells[LIN_CABECALHO, COL_Ocorrencia].Value = "Ocorrência";
                    ws.Cells[LIN_CABECALHO, COL_TipoOcorrencia].Value = "Tipo de Ocorrência";


                    ws.Cells[LIN_CABECALHO, COL_Pedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_NF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Transportadora].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Ocorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_TipoOcorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    #endregion

                    #region [ Registros ]
                    using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, COL_TipoOcorrencia]])
                    {
                        rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    for (int i = 0; i < NumRegistros; i++)
                    {
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Pedido].Value = datasource.ElementAt(i).Pedido;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Pedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_NF].Value = datasource.ElementAt(i).NF;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_NF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Transportadora].Value = datasource.ElementAt(i).Transportadora;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Transportadora].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Ocorrencia].Value = datasource.ElementAt(i).Ocorrencia;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Ocorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_TipoOcorrencia].Value = datasource.ElementAt(i).TipoOcorrencia;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_TipoOcorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_TipoOcorrencia].Style.WrapText = true;

                    }
                    #endregion


                    #region [ Total ]

                    ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), COL_Pedido, (NumRegistros + LIN_INICIO_REGISTROS + 1), COL_TipoOcorrencia].Style.Font.Bold = true;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_Pedido].Value = "TOTAL:";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_Pedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_NF].Value = NumRegistros + " Ocorrência(s)";
                    #endregion

                    pck.SaveAs(new FileInfo(filePath));
                }
            });
        }
    }
}