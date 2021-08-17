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
using System.Text;

namespace ART3WebAPI.Models.Domains
{
	public class RelPreDevolucaoGeradorRelatorio
	{
		public static Task GeraXLS(List<RelPreDevolucaoEntity> dataRel, string filePath, string usuario, string loja, string filtro_status, string filtro_data_inicio, string filtro_data_termino, string filtro_lojas)
		{
			return Task.Run(() =>
			{
				#region [ Declarações ]
				const int COL_INICIAL = 2;
				const int ROW_INICIAL = 2;
				int COL_FINAL;
				int ROW_CABECALHO, ROW_INICIO_DADOS;
				int colDtCadastro = 0, colIdDevolucao = 0, colLoja = 0, colPedido = 0, colVendedor = 0, colIndicador = 0, colCliente = 0, colTransportadora = 0, colMotivo = 0, colVlPedido = 0, colVlDevolucao = 0, colStatus = 0;
				int rowUltLinhaDados;
				int nLinha = 0;
				int NumRegistros = dataRel.Count;
				string cellsIndex;
				string descricao_filtro_status;
				string descricao_filtro_loja = "";
				string[] vLojas;
				string[] vAux;
				StringBuilder sbAux = new StringBuilder("");
				DateTime dt_inicio = DateTime.MinValue;
				DateTime dt_termino = DateTime.MinValue;
				#endregion

				#region [ Preparação de campos que descrevem os filtros ]

				#region [ Período ]
				if ((filtro_data_inicio ?? "").Length > 0) dt_inicio = Global.converteDdMmYyyyParaDateTime(filtro_data_inicio);
				if ((filtro_data_termino ?? "").Length > 0) dt_termino = Global.converteDdMmYyyyParaDateTime(filtro_data_termino);
				#endregion

				#region [ Status ]
				switch (filtro_status)
				{
					case "CADASTRADA":
						descricao_filtro_status = "Cadastrada";
						break;
					case "EM_ANDAMENTO":
						descricao_filtro_status = "Em Andamento";
						break;
					case "MERCADORIA_RECEBIDA":
						descricao_filtro_status = "Mercadoria Recebida";
						break;
					case "FINALIZADA":
						descricao_filtro_status = "Finalizada";
						break;
					case "REPROVADA":
						descricao_filtro_status = "Reprovada";
						break;
					case "CANCELADA":
						descricao_filtro_status = "Cancelada";
						break;
					case "TODOS":
						descricao_filtro_status = "Todos";
						break;
					default:
						descricao_filtro_status = "Parâmetro Desconhecido";
						break;
				}
				#endregion

				#region [ Loja ]
				// Se o relatório está sendo solicitado através do módulo Loja, assegura que ocorra a filtragem pela loja
				if (((filtro_lojas ?? "").Length == 0) && ((loja ?? "").Length > 0)) filtro_lojas = loja;

				if ((filtro_lojas ?? "").Length > 0)
				{
					filtro_lojas = filtro_lojas.Replace('_', ',');
					vLojas = filtro_lojas.Split(',');

					for (int i = 0; i < vLojas.Length; i++)
					{
						if (vLojas.Length == 0) continue;

						if (vLojas[i].Trim().Length > 0)
						{
							vAux = vLojas[i].Split('-');
							if (vAux.Length == 1)
							{
								if (descricao_filtro_loja.Length > 0) descricao_filtro_loja += ", ";
								descricao_filtro_loja += vLojas[i].Trim();
							}
							else
							{
								sbAux.Clear();
								if (descricao_filtro_loja.Length > 0) descricao_filtro_loja += ", ";
								descricao_filtro_loja += (vAux[0].Trim().Length == 0 ? "N.I." : vAux[0].Trim()) + " a " + (vAux[1].Trim().Length == 0 ? "N.I." : vAux[1].Trim());
							}
						}
					}
				}
				#endregion

				#endregion

				using (ExcelPackage pck = new ExcelPackage())
				{
					//Cria uma planilha com nome
					ExcelWorksheet ws = pck.Workbook.Worksheets.Add("RelPreDevolucao");

					#region [ Configurações gerais ]
					//configurações gerais da planilha
					ws.Cells["A:XFD"].Style.Font.Name = "Arial";
					ws.Cells["A:XFD"].Style.Font.Size = 10;
					ws.View.ShowGridLines = false;
					ws.Column(1).Width = 2;

					colDtCadastro = COL_INICIAL;
					colIdDevolucao = colDtCadastro + 1;
					colLoja = colIdDevolucao + 1;
					colPedido = colLoja + 1;
					colVendedor = colPedido + 1;
					colIndicador = colVendedor + 1;
					colCliente = colIndicador + 1;
					colTransportadora = colCliente + 1;
					colMotivo = colTransportadora + 1;
					colVlPedido = colMotivo + 1;
					colVlDevolucao = colVlPedido + 1;
					colStatus = colVlDevolucao + 1;
					COL_FINAL = colStatus;

					ws.Column(colDtCadastro).Width = 11;
					ws.Column(colIdDevolucao).Width = 9;
					ws.Column(colLoja).Width = 6;
					ws.Column(colPedido).Width = 12;
					ws.Column(colVendedor).Width = 14;
					ws.Column(colIndicador).Width = 18;
					ws.Column(colCliente).Width = 30;
					ws.Column(colTransportadora).Width = 20;
					ws.Column(colMotivo).Width = 30;
					ws.Column(colVlPedido).Width = 14;
					ws.Column(colVlDevolucao).Width = 14;
					ws.Column(colStatus).Width = 20;

					ws.Row(1).Height = 1;
					#endregion

					#region [ Informativo dos filtros utilizados ]
					nLinha = ROW_INICIAL;
					cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Style.Font.Size = 12;
					ws.Cells[cellsIndex].Value = "Relatório de Pré-Devoluções";
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Período: " + (dt_inicio == DateTime.MinValue ? "N.I." : Global.formataDataDdMmYyyyComSeparador(dt_inicio)) + " a " + (dt_termino == DateTime.MinValue ? "N.I." : Global.formataDataDdMmYyyyComSeparador(dt_termino));
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Status da Pré-Devolução: " + descricao_filtro_status;
					nLinha++; cellsIndex = Excel.CellAddress(COL_INICIAL, nLinha);
					ws.Cells[cellsIndex].Value = "Loja: " + (descricao_filtro_loja.Length == 0 ? "N.I." : descricao_filtro_loja);
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

					ws.Cells[ROW_CABECALHO, colDtCadastro].Value = "DT Cadastro";
					ws.Cells[ROW_CABECALHO, colDtCadastro].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					ws.Cells[ROW_CABECALHO, colIdDevolucao].Value = "ID Devol";
					ws.Cells[ROW_CABECALHO, colIdDevolucao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					ws.Cells[ROW_CABECALHO, colLoja].Value = "Loja";
					ws.Cells[ROW_CABECALHO, colLoja].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					ws.Cells[ROW_CABECALHO, colPedido].Value = "Pedido";
					ws.Cells[ROW_CABECALHO, colPedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colVendedor].Value = "Vendedor";
					ws.Cells[ROW_CABECALHO, colVendedor].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colIndicador].Value = "Indicador";
					ws.Cells[ROW_CABECALHO, colIndicador].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colCliente].Value = "Cliente";
					ws.Cells[ROW_CABECALHO, colCliente].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colTransportadora].Value = "Transp";
					ws.Cells[ROW_CABECALHO, colTransportadora].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colMotivo].Value = "Motivo";
					ws.Cells[ROW_CABECALHO, colMotivo].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					ws.Cells[ROW_CABECALHO, colVlPedido].Value = "VL Pedido";
					ws.Cells[ROW_CABECALHO, colVlPedido].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					ws.Cells[ROW_CABECALHO, colVlDevolucao].Value = "VL Devolução";
					ws.Cells[ROW_CABECALHO, colVlDevolucao].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					ws.Cells[ROW_CABECALHO, colStatus].Value = "Status";
					ws.Cells[ROW_CABECALHO, colStatus].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

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

					cellsIndex = Excel.RangeAddress(colDtCadastro, ROW_INICIO_DADOS, colDtCadastro, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "dd/mm/yyyy";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						rng.Style.WrapText = false;
					}

					cellsIndex = Excel.RangeAddress(colIdDevolucao, ROW_INICIO_DADOS, colIdDevolucao, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					}

					cellsIndex = Excel.RangeAddress(colLoja, ROW_INICIO_DADOS, colLoja, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					}

					cellsIndex = Excel.RangeAddress(colPedido, ROW_INICIO_DADOS, colPedido, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
					}

					cellsIndex = Excel.RangeAddress(colVendedor, ROW_INICIO_DADOS, colVendedor, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = true;
					}

					cellsIndex = Excel.RangeAddress(colIndicador, ROW_INICIO_DADOS, colIndicador, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = true;
					}

					cellsIndex = Excel.RangeAddress(colCliente, ROW_INICIO_DADOS, colCliente, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = true;
					}

					cellsIndex = Excel.RangeAddress(colTransportadora, ROW_INICIO_DADOS, colTransportadora, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = true;
					}

					cellsIndex = Excel.RangeAddress(colMotivo, ROW_INICIO_DADOS, colMotivo, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
						rng.Style.WrapText = true;
					}

					cellsIndex = Excel.RangeAddress(colVlPedido, ROW_INICIO_DADOS, colVlPedido, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					cellsIndex = Excel.RangeAddress(colVlDevolucao, ROW_INICIO_DADOS, colVlDevolucao, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.Numberformat.Format = "###,###,##0.00;[Red]-###,###,##0.00";
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					}

					cellsIndex = Excel.RangeAddress(colStatus, ROW_INICIO_DADOS, colStatus, rowUltLinhaDados);
					using (ExcelRange rng = ws.Cells[cellsIndex])
					{
						rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
						rng.Style.WrapText = true;
					}
					#endregion

					#region [ Preenche linhas de dados ]
					for (int i = 0; i < NumRegistros; i++)
					{
						ws.Cells[i + ROW_INICIO_DADOS, colDtCadastro].Value = dataRel.ElementAt(i).dt_hr_cadastro;
						ws.Cells[i + ROW_INICIO_DADOS, colIdDevolucao].Value = dataRel.ElementAt(i).id_devolucao;
						ws.Cells[i + ROW_INICIO_DADOS, colLoja].Value = dataRel.ElementAt(i).loja;
						ws.Cells[i + ROW_INICIO_DADOS, colPedido].Value = dataRel.ElementAt(i).pedido;
						ws.Cells[i + ROW_INICIO_DADOS, colVendedor].Value = dataRel.ElementAt(i).vendedor;
						ws.Cells[i + ROW_INICIO_DADOS, colIndicador].Value = dataRel.ElementAt(i).indicador;
						ws.Cells[i + ROW_INICIO_DADOS, colCliente].Value = dataRel.ElementAt(i).cliente_nome;
						ws.Cells[i + ROW_INICIO_DADOS, colTransportadora].Value = dataRel.ElementAt(i).transportadora_id;
						ws.Cells[i + ROW_INICIO_DADOS, colMotivo].Value = dataRel.ElementAt(i).descricao_devolucao_motivo;
						ws.Cells[i + ROW_INICIO_DADOS, colVlPedido].Value = dataRel.ElementAt(i).vl_pedido;
						ws.Cells[i + ROW_INICIO_DADOS, colVlDevolucao].Value = dataRel.ElementAt(i).vl_devolucao;
						ws.Cells[i + ROW_INICIO_DADOS, colStatus].Value = Global.getDescricaoStPedidoDevolucao(dataRel.ElementAt(i).status) + "\n(" + Global.formataDataDdMmYyyyHhMmComSeparador(dataRel.ElementAt(i).status_data_hora) + ")";
					}
					#endregion

					#endregion

					pck.SaveAs(new FileInfo(filePath));
				} // using (ExcelPackage pck = new ExcelPackage())
			});
		}
	}
}