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
    public class OcorrenciasGeradorRelatorio
    {
        #region [ Constantes ]
        private const int LIN_INICIO_REGISTROS = 8;
        private const int LIN_CABECALHO = 7;
        private const int COL_Loja = 2;
        private const int COL_CD = 3;
        private const int COL_Pedido = 4;
        private const int COL_NF = 5;
        private const int COL_Cliente = 6;
        private const int COL_UF = 7;
        private const int COL_Cidade = 8;
        private const int COL_Transportadora = 9;
        private const int COL_Contato = 10;
        private const int COL_Telefone = 11;
        private const int COL_Ocorrencia = 12;
        private const int COL_TipoOcorrencia = 13;
        private const int COL_Status = 14;
        #endregion

        public static Task GenerateXLS(List<OcorrenciasStatus> datasource, string filePath, string oc_status, string loja, string transportadora)
        {
            return Task.Run(() =>
            {
                int NumRegistros = datasource.Count;

                if (string.IsNullOrEmpty(oc_status))
                    oc_status = "Ambos";
                if ((string.IsNullOrEmpty(transportadora)) || (transportadora == "0"))
                    transportadora = "Todos";
                if (string.IsNullOrEmpty(loja))
                    loja = "Todas";
                using (ExcelPackage pck = new ExcelPackage())
                {
                    //Cria uma planilha com nome
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Ocorrências ");

                    #region [ Config Gerais ]
                    //configurações gerais da planilha
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.View.ShowGridLines = false;


                    ws.Column(1).Width = 2;
                    ws.Column(COL_Loja).Width = 10;
                    ws.Column(COL_CD).Width = 10;
                    ws.Column(COL_Pedido).Width = 10;
                    ws.Column(COL_NF).Width = 10;
                    ws.Column(COL_Cliente).Width = 40;
					ws.Column(COL_UF).Width = 10;
					ws.Column(COL_Cidade).Width = 40;
					ws.Column(COL_Transportadora).Width = 30;
					ws.Column(COL_Contato).Width = 30;
                    ws.Column(COL_Telefone).Width = 20;
					ws.Column(COL_Ocorrencia).Width = 40;
                    ws.Column(COL_TipoOcorrencia).Width = 20;
					ws.Column(COL_Status).Width = 20;
 
                    ws.Row(1).Height = 1;                
                    #endregion


                    #region [ Filtro ]
                    ws.Cells["B2:M12"].Style.Font.Bold = true;
                    ws.Cells["B2"].Style.Font.Size = 12;
                    ws.Cells["B2"].Value = "Ocorrências ";
                    ws.Cells["B3"].Value = "Transportadora: " + transportadora;
                    ws.Cells["B4"].Value = "Loja: " + loja;
                    ws.Cells["B5"].Value = "Emissão: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    #endregion

                    #region [ Cabeçalho ]
                    using (ExcelRange rng1 = ws.Cells["B" + (LIN_INICIO_REGISTROS - 1) + ":" + ws.Cells[LIN_INICIO_REGISTROS - 1, COL_Status]])
                    {
                        rng1.Style.WrapText = true;
                        rng1.Style.Font.Bold = true;
                        rng1.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        rng1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        rng1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }

                    ws.Cells[LIN_CABECALHO, COL_Status].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    ws.Cells[LIN_CABECALHO, COL_Loja].Value = "Loja";
                    ws.Cells[LIN_CABECALHO, COL_CD].Value = "CD";
                    ws.Cells[LIN_CABECALHO, COL_Pedido].Value = "Pedido";
                    ws.Cells[LIN_CABECALHO, COL_NF].Value = "NF";
                    ws.Cells[LIN_CABECALHO, COL_Cliente].Value = "Cliente";
                    ws.Cells[LIN_CABECALHO, COL_UF].Value = "Estado";
                    ws.Cells[LIN_CABECALHO, COL_Cidade].Value = "Cidade";
                    ws.Cells[LIN_CABECALHO, COL_Transportadora].Value = "Transportadora";
                    ws.Cells[LIN_CABECALHO, COL_Contato].Value = "Contato";
                    ws.Cells[LIN_CABECALHO, COL_Telefone].Value = "Telefone";
                    ws.Cells[LIN_CABECALHO, COL_Ocorrencia].Value = "Ocorrência";
                    ws.Cells[LIN_CABECALHO, COL_TipoOcorrencia].Value = "Tipo de Ocorrência";
					ws.Cells[LIN_CABECALHO, COL_Status].Value = "Status";                    


                    ws.Cells[LIN_CABECALHO, COL_Loja].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_CD].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Pedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_NF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Cliente].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_UF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Cidade].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Transportadora].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Contato].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Telefone].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Ocorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_TipoOcorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[LIN_CABECALHO, COL_Status].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    #endregion

                    #region [ Registros ]
                    using (ExcelRange rng2 = ws.Cells["B" + LIN_INICIO_REGISTROS + ":" + ws.Cells[NumRegistros + LIN_INICIO_REGISTROS - 1, COL_Status]])
                    {
                        rng2.Style.Font.Bold = false;
                        rng2.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        rng2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        rng2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    for (int i = 0; i < NumRegistros; i++)
                    {
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Loja].Value = datasource.ElementAt(i).Loja;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Loja].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CD].Value = datasource.ElementAt(i).CD;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_CD].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Pedido].Value = datasource.ElementAt(i).Pedido;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Pedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_NF].Value = datasource.ElementAt(i).NF;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_NF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Cliente].Value = datasource.ElementAt(i).Cliente;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Cliente].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						ws.Cells[i + LIN_INICIO_REGISTROS, COL_Cliente].Style.WrapText = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_UF].Value = datasource.ElementAt(i).UF;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_UF].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Cidade].Value = datasource.ElementAt(i).Cidade;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Cidade].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Transportadora].Value = datasource.ElementAt(i).Transportadora;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Transportadora].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Transportadora].Style.WrapText = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Contato].Value = datasource.ElementAt(i).Contato;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Contato].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Contato].Style.WrapText = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Telefone].Value = datasource.ElementAt(i).Telefone;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Telefone].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Ocorrencia].Value = datasource.ElementAt(i).Ocorrencia;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Ocorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Ocorrencia].Style.WrapText = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_TipoOcorrencia].Value = datasource.ElementAt(i).TipoOcorrencia;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_TipoOcorrencia].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_TipoOcorrencia].Style.WrapText = true;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Status].Value = datasource.ElementAt(i).Status;
                        ws.Cells[i + LIN_INICIO_REGISTROS, COL_Status].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        
                    }
                    #endregion


                    #region [ Total ]

                    ws.Cells[(NumRegistros + LIN_INICIO_REGISTROS + 1), COL_Loja, (NumRegistros + LIN_INICIO_REGISTROS + 1), COL_Status].Style.Font.Bold = true;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_Loja].Value = "TOTAL:";
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_Loja].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[NumRegistros + LIN_INICIO_REGISTROS + 1, COL_CD].Value = NumRegistros + " Ocorrência(s)";
                    #endregion

                    pck.SaveAs(new FileInfo(filePath));
                }
            });
        }
    }
}