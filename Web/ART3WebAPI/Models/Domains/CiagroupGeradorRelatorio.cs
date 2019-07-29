using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;

namespace ART3WebAPI.Models.Domains
{
    public class CiagroupGeradorRelatorio
    {

        public static Task GenerateCSV(List<Indicador> datasource, string filePath)
        {
            return Task.Run(() =>
            {
                Encoding encode = Encoding.GetEncoding("Windows-1252");
                using (StreamWriter sw = new StreamWriter(new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Read), encode))
                {
                    
                    string delimitador = ";";
                    int length = datasource.Count;
                    StringBuilder sb = new StringBuilder();
                    
                    //cria cabeçalho
                    sb.Append("Nome" + delimitador);
                    sb.Append("TipoPessoa(F/J)" + delimitador);
                    sb.Append("CPF/CNPJ" + delimitador);
                    sb.Append("Banco" + delimitador);
                    sb.Append("Agencia" + delimitador);
                    sb.Append("DigitoAgencia" + delimitador);
                    sb.Append("CodOperação" + delimitador);
                    sb.Append("Conta" + delimitador);
                    sb.Append("DigitoConta" + delimitador);
                    sb.Append("Tipoconta(CC/CP)" + delimitador);
                    sb.Append("Valor" + delimitador);
                    sb.Append("CodigoCliente" + delimitador);

                    sw.WriteLine(sb);

                    for (int i = 0; i < length; i++)
                    {
                        
                        sw.WriteLine(datasource.ElementAt(i).Nome.ToUpper() + delimitador + 
                            datasource.ElementAt(i).TipoPessoa + delimitador +
                            Global.formataCnpjCpf(datasource.ElementAt(i).CpfCnpj) + delimitador +
                            datasource.ElementAt(i).Banco + delimitador +
                            datasource.ElementAt(i).Agencia + delimitador +
                            datasource.ElementAt(i).DigitoAgencia + delimitador +
                            datasource.ElementAt(i).Operacao + delimitador +
                            datasource.ElementAt(i).Conta + delimitador +
                            datasource.ElementAt(i).DigitoConta + delimitador +
                            datasource.ElementAt(i).TipoContaCSV + delimitador +
                            Global.formataMoeda(datasource.ElementAt(i).Valor) + delimitador + 
                            "" + delimitador);
                    }

                }

            });
        }


        #region [ GenerateXLS ]
        public static Task GenerateXLS(List<Indicador> datasource, string filePath)
        {
            return Task.Run(() =>
            {
                
                using (ExcelPackage pck = new ExcelPackage())
                {
                    DataCiagroup e = new DataCiagroup();
                                        
                    //Cria uma planilha com nome
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Relação de Favorecidos");

                    //configurações gerais da planilha
                    ws.Cells["A:XFD"].Style.Font.Name = "Arial";
                    ws.Cells["A:XFD"].Style.Font.Size = 10;
                    ws.Cells["A:XFD"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.View.ShowGridLines = false;

                    ws.Cells["B2:L3"].Merge = true;
                    ws.Cells["B4:L4"].Merge = true;
                    ws.Cells["C6:D6"].Merge = true;
                    ws.Cells["E6:I6"].Merge = true;
                    ws.Cells["J6:L6"].Merge = true;
                    ws.Cells["A7:A8"].Merge = true;
                    ws.Cells["B7:B8"].Merge = true;
                    ws.Cells["C7:D8"].Merge = true;
                    ws.Cells["E7:I8"].Merge = true;
                    ws.Cells["J7:L7"].Merge = true;
                    ws.Cells["J8:L8"].Merge = true;
                    ws.Cells["A10:I10"].Merge = true;
                    ws.Cells["J10:L10"].Merge = true;
                    ws.Cells["B11:I11"].Merge = true;
                    ws.Cells["B12:I12"].Merge = true;
                    ws.Cells["B13:I13"].Merge = true;
                    ws.Cells["B14:I14"].Merge = true;
                    ws.Cells["J11:L11"].Merge = true;
                    ws.Cells["J12:L12"].Merge = true;
                    ws.Cells["J13:L13"].Merge = true;
                    ws.Cells["J14:L14"].Merge = true;
                    ws.Cells["J15:L15"].Merge = true;
                    ws.Cells["B15:C15"].Merge = true;
                    ws.Cells["E15:I15"].Merge = true;
                    ws.Cells["A17:K17"].Merge = true;
                    ws.Cells["A18:F18"].Merge = true;
                    ws.Cells["I18:K18"].Merge = true;
                    ws.Cells["A19:K19"].Merge = true;

                    ws.Column(1).Width = 19;
                    ws.Column(2).Width = 57;
                    ws.Column(3).Width = 13;
                    ws.Column(4).Width = 25;
                    ws.Column(5).Width = 9;
                    ws.Column(6).Width = 10;
                    ws.Column(7).Width = 4;
                    ws.Column(8).Width = 10;
                    ws.Column(9).Width = 14;
                    ws.Column(10).Width = 4;
                    ws.Column(11).Width = 17;
                    ws.Column(12).Width = 33;

                    ws.Row(1).Height = 14;
                    ws.Row(2).Height = 19;
                    ws.Row(3).Height = 22;
                    ws.Row(4).Height = 23;
                    ws.Row(5).Height = 14;
                    ws.Row(6).Height = 23;
                    ws.Row(7).Height = 23;
                    ws.Row(8).Height = 23;
                    ws.Row(9).Height = 14;
                    ws.Row(10).Height = 23;
                    ws.Row(11).Height = 23;
                    ws.Row(12).Height = 23;
                    ws.Row(13).Height = 23;
                    ws.Row(14).Height = 23;
                    ws.Row(15).Height = 23;
                    ws.Row(16).Height = 14;
                    ws.Row(17).Height = 27;
                    ws.Row(18).Height = 27;
                    ws.Row(19).Height = 27;
                    ws.Row(20).Height = 14;
                    ws.Row(21).Height = 45;

                    #region[ primeiro bloco ]
                    ws.Cells["A2"].Value = "RAZÃO SOCIAL";
                    ws.Cells["B2"].Value = e.empresa.RazaoSocial;
                    ws.Cells["A4"].Value = "CNPJ:";
                    ws.Cells["B4"].Value = Global.formataCnpjCpf(e.empresa.Cnpj);
                    ws.Cells["A2:A4"].Style.Font.Size = 12;
                    ws.Cells["A2:L4"].Style.Font.Bold = true;
                    ws.Cells["A2:L4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["B2"].Style.Font.Size = 16;
                    ws.Cells["B4"].Style.Font.Size = 14;
                    ws.Cells["B2:L3"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A2:A3"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A4"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["B4:L4"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A7:A8"].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    #endregion

                    #region [ segundo bloco ]
                    using (ExcelRange serie2 = ws.Cells["A11:L15"])
                    {
                        serie2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    }
                    ws.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["B7:B8"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["C7:D8"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["E7:I8"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["C6:D6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["E6:I6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["J7:L8"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["J6:L6"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A6:L6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["A6:L6"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    ws.Cells["A6:L6"].Style.Font.Size = 11;
                    ws.Cells["A6:L8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["A6:L8"].Style.Font.Bold = true;
                    ws.Cells["A6"].Value = "NOTA FISCAL Nº";
                    ws.Cells["B6"].Value = "PRODUTO";
                    ws.Cells["C6"].Value = "DATA DO PEDIDO";
                    ws.Cells["E6"].Value = "VENCT° N FISCAL";
                    ws.Cells["J6"].Value = "INSERIR NO CORPO DA NOTA FISCAL";
                    ws.Cells["A7:I8"].Style.Font.Size = 12;
                    ws.Cells["J6:L8"].Style.Font.Size = 10;
                    ws.Cells["J8:L8"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                    #endregion

                    #region [ terceiro bloco ]
                    ws.Cells["A10:L15"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A10:I10"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["J10:L10"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A11:A15"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["B11:I15"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A10:L10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["A10:L10"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    ws.Cells["A11:A15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["A11:A15"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    ws.Cells["A10:L15"].Style.Font.Size = 10;
                    ws.Cells["A10:L10"].Style.Font.Bold = true;
                    ws.Cells["A11:A15"].Style.Font.Bold = true;
                    ws.Cells["A10:L10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["A11:A15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["A10"].Value = "ENDEREÇO DE ENTREGA NF";
                    ws.Cells["J10"].Value = "OBSERVAÇÕES";
                    ws.Cells["A11"].Value = "NOME FANTASIA";
                    ws.Cells["B11"].Value = e.empresa.NomeFantasia;
                    ws.Cells["A12"].Value = "ENDEREÇO";
                    ws.Cells["B12"].Value = e.empresa.Endereco;
                    ws.Cells["A13"].Value = "CIDADE / UF / CEP";
                    ws.Cells["B13"].Value = e.empresa.Cidade + " - " + e.empresa.Uf + ", " + e.empresa.Cep;
                    ws.Cells["A14"].Value = "CONTATO";
                    ws.Cells["B14"].Value = e.empresa.Contato;
                    ws.Cells["A15"].Value = "EMAIL";
                    ws.Cells["B15"].Value = e.empresa.Email;
                    ws.Cells["D15"].Value = "TELEFONE";
                    ws.Cells["E15"].Value = e.empresa.Telefone;
                    ws.Cells["D15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["D15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["D15"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    ws.Cells["D15"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells["D15"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ws.Cells["B11:L14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["B15:C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["E15:L15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells["D15:E15"].Style.Font.Bold = true;
                    #endregion

                    #region [ quarto bloco ]
                    ws.Cells["A17:L17"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A18:L18"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A19:L19"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    ws.Cells["A17:L19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["A17:L19"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    ws.Cells["A17:L19"].Style.Font.Bold = true;
                    ws.Cells["a17:l19"].Style.Font.Size = 12;
                    ws.Cells["A17"].Value = "VALOR DO PEDIDO";
                    ws.Cells["A18"].Value = "TAXA ADM";
                    ws.Cells["I18"].Value = e.empresa.TaxaAdm;
                    ws.Cells["A19"].Value = "TOTAL DA NF";
                    #endregion

                    ws.Cells[21, 2].Value = "NOME COMPLETO";
                    ws.Cells[21, 3].Value = "Tipo (1-CPF / 2 -CNPJ)";
                    ws.Cells[21, 4].Value = "CPF / CNPJ";
                    ws.Cells[21, 5].Value = "BANCO";
                    ws.Cells[21, 6].Value = "AGÊNCIA";
                    ws.Cells[21, 7].Value = "DV";
                    ws.Cells[21, 8].Value = "Oper";
                    ws.Cells[21, 9].Value = "CONTA";
                    ws.Cells[21, 10].Value = "DV";
                    ws.Cells[21, 11].Value = "TIPO CONTA (C - Corrente / P - Poupança)";
                    ws.Cells[21, 12].Value = "VALOR";
                    ws.Cells[21, 8].AddComment("Oper: Operação - Para Contas da Caixa Econômica Federal (Banco 104)", "Me");

                    for (int i = 0; i < datasource.Count(); i++)
                    {
                        ws.Row(i + 22).Height = 16;
                        ws.Cells[i + 22, 2].Value = datasource.ElementAt(i).Nome.ToUpper();
                        ws.Cells[i + 22, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 3].Value = datasource.ElementAt(i).TipoDocumento;
                        ws.Cells[i + 22, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 4].Value = Global.formataCnpjCpf(datasource.ElementAt(i).CpfCnpj);
                        ws.Cells[i + 22, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 5].Value = datasource.ElementAt(i).Banco;
                        ws.Cells[i + 22, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 6].Value = datasource.ElementAt(i).Agencia;
                        ws.Cells[i + 22, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 7].Value = datasource.ElementAt(i).DigitoAgencia;
                        ws.Cells[i + 22, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 8].Value = datasource.ElementAt(i).Operacao;
                        ws.Cells[i + 22, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 9].Value = datasource.ElementAt(i).Conta;
                        ws.Cells[i + 22, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 10].Value = datasource.ElementAt(i).DigitoConta;
                        ws.Cells[i + 22, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 11].Value = datasource.ElementAt(i).TipoConta;
                        ws.Cells[i + 22, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[i + 22, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        ws.Cells[i + 22, 12].Value = datasource.ElementAt(i).Valor;
                        ws.Cells[i + 22, 12].Style.Numberformat.Format = "#,###.00";
                        ws.Cells[i + 22, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    //cabeçalho da relação de favorecidos
                    using (ExcelRange serie1 = ws.Cells["B21:L21"])
                    {
                        serie1.Style.Font.Bold = true;
                        serie1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        serie1.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        serie1.Style.Font.Color.SetColor(Color.Black);
                        serie1.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        serie1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        serie1.Style.WrapText = true;
                        serie1.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        serie1.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }
                    ws.Cells["B21"].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                    pck.SaveAs(new FileInfo(filePath));
                }
            });
        }
        #endregion
    }
}