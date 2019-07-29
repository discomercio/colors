#region [ using ]
using System;
using System.Data;
using System.Drawing;
using System.Media;
using System.Reflection;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FTitulosConciliacao : FModelo
	{
		#region [ Atributos ]
		#region [ Diversos ]
		private bool _atualizacaoAutomaticaPesquisaEmAndamento = false;
		private DataSet _dsBoletosComInfoPagtoDivergentes;
		DateTime _data_final_periodo;
		FTrataConciliacao _fTrataConciliacao;
		#endregion
		#endregion

		#region [ Construtor ]
		public FTitulosConciliacao()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos ]

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			int intQtdeRegistros = 0;
			int id_serasa_arq_conciliacao_input;
			int qtdeRegProcessado;
			int percProgressoAtual;
			int percProgressoAnterior;
			String msg_erro;
			String strMsgProgresso;
			DateTime data_final_periodo;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			if (_atualizacaoAutomaticaPesquisaEmAndamento) return false;

			_atualizacaoAutomaticaPesquisaEmAndamento = true;
			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
				if (!ConciliacaoTituloDAO.obtemArquivoRetornoParaTratamento(out id_serasa_arq_conciliacao_input, out data_final_periodo, out msg_erro))
				{
					avisoErro(msg_erro);
					gridDados.Rows.Clear();
					lblTotalizacaoRegistros.Text = "";
					return false;
				}

				_data_final_periodo = data_final_periodo;

				dtbConsulta = ConciliacaoTituloDAO.selecionaBoletosParaTratamento(id_serasa_arq_conciliacao_input);
				if (dtbConsulta.Rows.Count == 0)
				{
					aviso("Não há conciliação a ser tratada!!");
					gridDados.Rows.Clear();
					lblTotalizacaoRegistros.Text = "";
					return false;
				}

				#region [ Exibição dos dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					gridDados.SuspendLayout();

					#region [ Carrega os dados no grid ]
					gridDados.Rows.Clear();
					if (dtbConsulta.Rows.Count > 0) gridDados.Rows.Add(dtbConsulta.Rows.Count);
					qtdeRegProcessado = 0;
					percProgressoAnterior = 0;
					for (int i = 0; i < dtbConsulta.Rows.Count; i++)
					{
						qtdeRegProcessado++;
						percProgressoAtual = 100 * qtdeRegProcessado / dtbConsulta.Rows.Count;
						if (percProgressoAtual != percProgressoAnterior)
						{
							percProgressoAnterior = percProgressoAtual;
							strMsgProgresso = "carregando dados no grid para exibição: " + percProgressoAtual.ToString() + "%";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							Application.DoEvents();
						}

						rowConsulta = dtbConsulta.Rows[i];
						String numTitulo = BD.readToString(rowConsulta["num_titulo_estendido"]).Trim();

						gridDados.Rows[i].Cells["id"].Value = BD.readToString(rowConsulta["id"]);
						gridDados.Rows[i].Cells["nosso_numero"].Value = numTitulo.Substring(0, numTitulo.Length - 1) + "-" + numTitulo.Substring(numTitulo.Length - 1);
						gridDados.Rows[i].Cells["dt_emissao"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_data_emissao"]));
						gridDados.Rows[i].Cells["dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_data_vencto_original"]));
						gridDados.Rows[i].Cells["vl_titulo"].Value = Global.formataMoeda(BD.readToDecimal(rowConsulta["vl_valor_titulo_original"]));
						if (rowConsulta["dt_data_pagto_editado"] == DBNull.Value)
						{
							gridDados.Rows[i].Cells["dt_pagto"].Value = "";
						}
						else
						{
							gridDados.Rows[i].Cells["dt_pagto"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_data_pagto_editado"]));
						}

						intQtdeRegistros++;
					}
					#endregion

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < gridDados.Rows.Count; i++)
					{
						if (gridDados.Rows[i].Selected) gridDados.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					gridDados.ResumeLayout();
				}
				#endregion

				#region [ Exibe totalização ]
				lblTotalizacaoRegistros.Text = Global.formataInteiro(intQtdeRegistros);
				#endregion

				gridDados.Focus();

				// Feedback da conclusão da pesquisa
				SystemSounds.Exclamation.Play();

				return true;
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return false;
			}
			finally
			{
				_atualizacaoAutomaticaPesquisaEmAndamento = false;
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ executaPesquisaBoletosComInfoPagtoDivergentes ]
		private bool executaPesquisaBoletosComInfoPagtoDivergentes(out bool blnOcorreuErro)
		{
			blnOcorreuErro = false;

			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				_dsBoletosComInfoPagtoDivergentes = ConciliacaoTituloDAO.selecionaBoletosComInfoPagtoDivergentes();
				if (_dsBoletosComInfoPagtoDivergentes.Tables["DtbBoleto"].Rows.Count == 0)
				{
					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Ocorreu um erro ao executar o método executaPesquisaBoletosComInfoPagtoDivergentes()!! " + ex.StackTrace);
				blnOcorreuErro = true;
				return false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ geraPlanilhaExcel ]
		private bool geraPlanilhaExcel()
		{
			#region [ Declarações ]
			const int MAX_LINHAS_EXCEL = 65536;
			const String FN_LISTAGEM = "Arial";
			const int FS_LISTAGEM = 8;
			const int FS_CABECALHO = 8;
			bool blnExcelSuportaUseSystemSeparators = false;
			bool blnExcelSuportaDecimalDataType = false;
			bool blnFlag;
			String strMsg;
			String strAux;
			String strExcelDecimalSeparator = "";
			String strExcelThousandsSeparator = "";
			String strTexto;
			int intQtdeRegistros = 0;
			int intPrimeiraLinhaDados = 0;
			int intUltimaLinhaDados = 0;
			int iNumLinha = 1;
			int iOffSetArray = 2;
			int iXlDadosMinIndex;
			int iXlDadosMaxIndex;
			int iXlMargemEsq;
			int iXlId;
			int iXlCNPJ;
			int iXlNumBoleto;
			int iXlDtEmissao;
			int iXlDtVencto;
			int iXlVlBoleto;
			int iXlDtPagto;
			int iXlDescricao;
			DateTime dtEmissao;
			DateTime dtVencto;
			DateTime dtPagto;
			Decimal vlBoleto;
			object oXL;
			object oWBs;
			object oWB;
			object oWS;
			object oWindow;
			object oWindows;
			object oPageSetup;
			object oStyles;
			object oStyle;
			object oFont;
			object oBorders;
			object oBorder;
			object oCells;
			object oCell;
			object oColumns;
			object oColumn;
			object oRows;
			object oRow;
			object oRange;
			object oApplication;
			String[] vDados;
			#endregion

			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "gerando planilha Excel");

				#region [ Cria instância do Excel ]
				try
				{
					oXL = ExcelAutomation.CriaInstanciaExcel();
				}
				catch (Exception ex)
				{
					strMsg = "Falha ao acionar o Excel!!\nVerifique se o Excel está instalado!!\n\n" + ex.ToString();
					avisoErro(strMsg);
					return false;
				}
				#endregion

				#region [ Inicializa planilha ]
				ExcelAutomation.SetProperty(oXL, "Visible", true);
				ExcelAutomation.SetProperty(oXL, "DisplayAlerts", false);
				ExcelAutomation.SetProperty(oXL, "SheetsInNewWorkbook", 1);
				oWBs = ExcelAutomation.GetProperty(oXL, "Workbooks");
				oWB = ExcelAutomation.InvokeMethod(oWBs, "Add", Missing.Value);
				oWindows = ExcelAutomation.GetProperty(oWB, "Windows");
				oWindow = ExcelAutomation.GetProperty(oWindows, "Item", 1);
				ExcelAutomation.SetProperty(oWindow, "DisplayGridlines", false);
				ExcelAutomation.SetProperty(oWindow, "DisplayHeadings", true);
				ExcelAutomation.SetProperty(oWindow, "WindowState", ExcelAutomation.XlWindowState.xlMaximized);
				oWS = ExcelAutomation.GetProperty(oWB, "ActiveSheet");
				oPageSetup = ExcelAutomation.GetProperty(oWS, "PageSetup");
				ExcelAutomation.SetProperty(oPageSetup, "PaperSize", ExcelAutomation.XlPaperSize.xlPaperA4);
				ExcelAutomation.SetProperty(oPageSetup, "Orientation", ExcelAutomation.XlPageOrientation.xlLandscape);
				ExcelAutomation.SetProperty(oPageSetup, "LeftMargin", 2);
				ExcelAutomation.SetProperty(oPageSetup, "RightMargin", 2);
				ExcelAutomation.SetProperty(oPageSetup, "TopMargin", 15);
				ExcelAutomation.SetProperty(oPageSetup, "BottomMargin", 15);
				ExcelAutomation.SetProperty(oPageSetup, "HeaderMargin", 5);
				ExcelAutomation.SetProperty(oPageSetup, "FooterMargin", 5);
				ExcelAutomation.SetProperty(oPageSetup, "CenterHorizontally", true);
				ExcelAutomation.SetProperty(oPageSetup, "CenterVertically", false);
				oStyles = ExcelAutomation.GetProperty(oWB, "Styles");
				oStyle = ExcelAutomation.GetProperty(oStyles, "Item", "Normal");
				ExcelAutomation.SetProperty(oStyle, "IncludeNumber", true);
				ExcelAutomation.SetProperty(oStyle, "IncludeFont", true);
				ExcelAutomation.SetProperty(oStyle, "IncludeAlignment", true);
				ExcelAutomation.SetProperty(oStyle, "IncludeBorder", true);
				ExcelAutomation.SetProperty(oStyle, "IncludePatterns", true);
				ExcelAutomation.SetProperty(oStyle, "IncludeProtection", true);
				ExcelAutomation.SetProperty(oStyle, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				ExcelAutomation.SetProperty(oStyle, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignTop);
				ExcelAutomation.SetProperty(oStyle, "WrapText", false);
				ExcelAutomation.SetProperty(oStyle, "IndentLevel", 0);
				ExcelAutomation.SetProperty(oStyle, "ShrinkToFit", false);
				oFont = ExcelAutomation.GetProperty(oStyle, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Bold", false);
				ExcelAutomation.SetProperty(oFont, "Italic", false);
				ExcelAutomation.SetProperty(oFont, "Underline", ExcelAutomation.XlUnderlineStyle.xlUnderlineStyleNone);
				ExcelAutomation.SetProperty(oFont, "Strikethrough", false);
				ExcelAutomation.SetProperty(oFont, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				oCells = ExcelAutomation.GetProperty(oWS, "Cells");
				ExcelAutomation.SetProperty(oCells, "Style", "Normal");
				ExcelAutomation.SetProperty(oCells, "NumberFormat", "@");
				ExcelAutomation.SetProperty(oWS, "DisplayPageBreaks", false);
				ExcelAutomation.SetProperty(oWS, "Name", "Checagem de Títulos Vencidos");
				ExcelAutomation.SetProperty(oXL, "DisplayAlerts", true);
				ExcelAutomation.SetProperty(oXL, "UserControl", true);
				#endregion

				#region [ Verifica se o Excel suporta o tipo 'decimal' ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", 1, 1);
				try
				{
					ExcelAutomation.SetProperty(oCell, "Value", (decimal)0.5);
					blnExcelSuportaDecimalDataType = true;
				}
				catch (Exception)
				{
					blnExcelSuportaDecimalDataType = false;
				}
				finally
				{
					ExcelAutomation.SetProperty(oCell, "Value", null);
				}
				#endregion

				#region [ Índices que definem a posição das colunas ]
				iXlMargemEsq = 1;
				iXlId = iXlMargemEsq + 1;
				iXlCNPJ = iXlId + 2;
				iXlDtEmissao = iXlCNPJ + 2;
				iXlDtVencto = iXlDtEmissao + 2;
				iXlDtPagto = iXlDtVencto + 2;
				iXlVlBoleto = iXlDtPagto + 2;
				iXlDescricao = iXlVlBoleto + 2;
				iXlNumBoleto = iXlVlBoleto + 2;
				#endregion

				#region [ Array usado p/ transferir dados p/ o Excel ]
				iXlDadosMinIndex = iXlMargemEsq + 1;
				iXlDadosMaxIndex = iXlNumBoleto;
				vDados = new string[(iXlDadosMaxIndex - iXlDadosMinIndex + 1)];
				#endregion

				#region [ Configura largura das colunas ]
				oColumns = ExcelAutomation.GetProperty(oWS, "Columns");
				// Margem
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlMargemEsq, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// ID
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlId, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 13.5);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlId + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// CNPJ
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlCNPJ, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 26.5);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlCNPJ + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Número Boleto
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlNumBoleto, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 19);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlNumBoleto + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Data Emissao
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtEmissao, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 18.67);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtEmissao + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Data Vencimento
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtVencto, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 18.67);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtVencto + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Valor Boleto
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlBoleto, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 18.17);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlBoleto + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				//Data Pagamento
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtPagto, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 18.67);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtPagto + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				//Descrição
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDescricao, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 19);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDescricao + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				#endregion

				#region [ Linha usada como margem superior ]
				oRows = ExcelAutomation.GetProperty(oWS, "Rows");
				oRow = ExcelAutomation.GetProperty(oRows, "Item", iNumLinha, Missing.Value);
				ExcelAutomation.SetProperty(oRow, "RowHeight", 5);
				iNumLinha++;
				#endregion

				#region [ Cabeçalho do relatório ]

				oCells = ExcelAutomation.GetProperty(oWS, "Cells");

				#region [ Título do relatório ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlId);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", 14);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Checagem de Títulos Vencidos");
				#endregion

				#region [ Data/hora da emissão ]
				strTexto = "Gerado em: " + Global.formataDataDdMmYyyyHhMmComSeparador(DateTime.Now);
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlBoleto);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Nova Linha ]
				iNumLinha++;
				#endregion

				#endregion

				#region [ Bordas dos títulos das colunas ]
				iNumLinha++;
				oRow = ExcelAutomation.GetProperty(oRows, "Item", iNumLinha, Missing.Value);
				ExcelAutomation.SetProperty(oRow, "RowHeight", 4);
				iNumLinha++;
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
				oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeTop);
				ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
				ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
				ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
				ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
				ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
				ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				#endregion

				#region [ Título das colunas ]

				#region [ ID ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlId);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "ID");
				#endregion

				#region [ CNPJ ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCNPJ);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "CNPJ");
				#endregion

				#region [ Número boleto ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNumBoleto);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Número Boleto");
				#endregion

				#region [ Data Emissão ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtEmissao);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Data Emissão");
				#endregion

				#region [ Data Vencimento ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtVencto);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Data Vencimento");
				#endregion

				#region [ Valor ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlBoleto);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Valor Boleto");
				#endregion

				#region [ Data Pagamento ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtPagto);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Data Pagamento");
				#endregion

				#region [ Descrição ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricao);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Descrição");
				#endregion

				#endregion

				#region [ Obtém separador decimal usado pelo Excel ]
				oApplication = ExcelAutomation.GetProperty(oXL, "Application");
				try
				{
					blnFlag = (bool)ExcelAutomation.GetProperty(oApplication, "UseSystemSeparators");
					if (blnFlag)
					{
						System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
						strExcelDecimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
						strExcelThousandsSeparator = ci.NumberFormat.NumberGroupSeparator;
					}
					else
					{
						strExcelDecimalSeparator = (string)ExcelAutomation.GetProperty(oApplication, "DecimalSeparator");
						strExcelThousandsSeparator = (string)ExcelAutomation.GetProperty(oApplication, "ThousandsSeparator");
					}

					blnExcelSuportaUseSystemSeparators = true;
				}
				catch (Exception)
				{
					blnExcelSuportaUseSystemSeparators = false;
				}

				if (!blnExcelSuportaUseSystemSeparators || (strExcelDecimalSeparator.Length == 0) || (strExcelThousandsSeparator.Length == 0))
				{
					System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
					strExcelDecimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
					strExcelThousandsSeparator = ci.NumberFormat.NumberGroupSeparator;
				}
				#endregion

				#region [ Formatação/alinhamento das colunas ]

				#region [ Coluna: ID]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlId) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlId) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				#endregion

				#region [ Coluna: CNPJ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlCNPJ) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlCNPJ) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				#endregion

				#region [ Coluna: Número Boleto ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlNumBoleto) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlNumBoleto) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				#endregion

				#region [ Coluna: Data Emissão ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDtEmissao) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDtEmissao) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormatLocal", "dd/mm/aaaa");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Coluna: Data Vencto ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDtVencto) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDtVencto) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormatLocal", "dd/mm/aaaa");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Coluna: Valor Boleto ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlVlBoleto) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlBoleto) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormatLocal", "#" + strExcelThousandsSeparator + "##0" + strExcelDecimalSeparator + "00");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
				#endregion

				#region [ Coluna: Data Pagamento ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDtPagto) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDtPagto) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormatLocal", "dd/mm/aaaa");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Coluna: Descrição ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDescricao) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDescricao) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				#endregion

				#endregion

				#region [ Laço para listagem ]
				foreach (DataRow row in _dsBoletosComInfoPagtoDivergentes.Tables["DtbBoleto"].Rows)
				{
					intQtdeRegistros++;
					iNumLinha++;
					if (intPrimeiraLinhaDados == 0) intPrimeiraLinhaDados = iNumLinha;
					intUltimaLinhaDados = iNumLinha;

					#region [ Transfere dados para o Excel (campos texto) ]

					#region [ ID ]
					vDados[iXlId - iOffSetArray] = BD.readToString(row["id"]);
					#endregion

					#region [ CNPJ ]
					vDados[iXlCNPJ - iOffSetArray] = Global.formataCnpjCpf(BD.readToString(row["cnpj_cliente"]));
					#endregion

					#region [ Número Boleto ]
					strAux = BD.readToString(row["num_titulo_estendido"]).Trim();
					vDados[iXlNumBoleto - iOffSetArray] = strAux.Insert(strAux.Length - 1, "-");
					#endregion

					#region [ Descrição ]
					vDados[iXlDescricao - iOffSetArray] = BD.readToString(row["descricao"]).Trim();
					#endregion

					#region [ Transfere dados do vetor p/ o Excel ]
					strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
					oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
					ExcelAutomation.SetProperty(oRange, "Value2", vDados);
					#endregion

					#endregion

					#region [ Transfere dados para o Excel (campos datetime) ]

					#region [ Data Emissão ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtEmissao);
					dtEmissao = BD.readToDateTime(row["dt_data_emissao"]);
					if (dtEmissao != DateTime.MinValue) ExcelAutomation.SetProperty(oCell, "Value", dtEmissao);
					#endregion

					#region [ Data Vencto ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtVencto);
					dtVencto = BD.readToDateTime(row["dt_data_vencto_original"]);
					if (dtVencto != DateTime.MinValue) ExcelAutomation.SetProperty(oCell, "Value", dtVencto);
					#endregion

					#region [ Data Pagto ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtPagto);
					dtPagto = BD.readToDateTime(row["dt_competencia"]);
					if (dtPagto != DateTime.MinValue) ExcelAutomation.SetProperty(oCell, "Value", dtPagto);
					#endregion

					#endregion

					#region [ Transfere dados para o Excel (campos numéricos) ]

					#region [ Valor Boleto ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlBoleto);
					vlBoleto = BD.readToDecimal(row["vl_valor_titulo_original"]);
					if (blnExcelSuportaDecimalDataType)
					{
						ExcelAutomation.SetProperty(oCell, "Value", vlBoleto);
					}
					else
					{
						ExcelAutomation.SetProperty(oCell, "Value", (double)vlBoleto);
					}
					#endregion

					#endregion

					#region [ Borda inferior da linha ]
					oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
					oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
					ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlDot);
					ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlHairline);
					ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
					#endregion
				}
				#endregion

				#region [ Linha com os totais ]

				#region [ Nova Linha ]
				iNumLinha++;
				#endregion

				#region [ Borda ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
				oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeTop);
				ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
				ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
				ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				#endregion

				#region [ Total de registros ]
				strAux = "TOTAL: " + intQtdeRegistros.ToString() + " registro(s)";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCNPJ);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "Value", strAux);
				#endregion

				#region [ Soma dos valores dos boletos ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlBoleto);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				strAux = "=SOMA(" + Global.excel_converte_numeracao_digito_para_letra(iXlVlBoleto) + intPrimeiraLinhaDados.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlBoleto) + intUltimaLinhaDados.ToString() + ")";
				ExcelAutomation.SetProperty(oCell, "FormulaLocal", strAux);
				#endregion

				#endregion

				// Feedback da conclusão da rotina
				SystemSounds.Exclamation.Play();

				return true;
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataOcorrenciaSelecionada ]
		private void trataOcorrenciaSelecionada()
		{
			#region [ Declarações ]
			int id = 0;
			String numBoleto = "";
			DateTime dtEmissao = DateTime.MinValue;
			DateTime dtVencto = DateTime.MinValue;
			Decimal vlTitulo = 0;
			DateTime dtPagto = DateTime.MinValue;

			DateTime dtVenctoCorrigido;
			Decimal vlTituloCorrigido;
			DateTime dtPagtoCorrigido;
			bool blnTituloExcluido = false;

			DialogResult drResultado;
			bool blnSucesso = false;
			DataGridViewRow rowGridSelecionado = null;
			#endregion

			#region [ Consistência ]
			if (gridDados.SelectedRows.Count == 0)
			{
				avisoErro("Nenhum registro foi selecionado!!");
				return;
			}

			if (gridDados.SelectedRows.Count > 1)
			{
				avisoErro("Não é permitida a seleção de múltiplos registros!!");
				return;
			}
			#endregion

			try
			{
				#region [ Obtém dados a serem editados do registro selecionado ]
				foreach (DataGridViewRow item in gridDados.SelectedRows)
				{
					rowGridSelecionado = item;
					id = Convert.ToInt32(item.Cells["id"].Value);
					numBoleto = item.Cells["nosso_numero"].Value.ToString();
					dtEmissao = Global.converteDdMmYyyyParaDateTime(item.Cells["dt_emissao"].Value.ToString());
					dtVencto = Global.converteDdMmYyyyParaDateTime(item.Cells["dt_vencto"].Value.ToString());
					vlTitulo = Global.converteNumeroDecimal(item.Cells["vl_titulo"].Value.ToString());
					if (item.Cells["dt_pagto"].Value.ToString().Trim().Length > 0)
					{
						dtPagto = Global.converteDdMmYyParaDateTime(item.Cells["dt_pagto"].Value.ToString());
					}
					else
					{
						dtPagto = DateTime.MinValue;
					}
				}
				#endregion

				#region [ Exibe painel p/ tratar o título em conciliação ]
				_fTrataConciliacao = new FTrataConciliacao(id,
														  numBoleto,
														  dtEmissao,
														  dtVencto,
														  vlTitulo,
														  dtPagto,
														  _data_final_periodo);

				_fTrataConciliacao.StartPosition = FormStartPosition.Manual;
				_fTrataConciliacao.Left = this.Left + (this.Width - _fTrataConciliacao.Width) / 2;
				_fTrataConciliacao.Top = this.Top + (this.Height - _fTrataConciliacao.Height) / 2;
				drResultado = _fTrataConciliacao.ShowDialog();
				if (drResultado != DialogResult.OK) return;
				#endregion

				#region [ Altera os valores do titulo em conciliação ]
				dtVenctoCorrigido = _fTrataConciliacao._dtVenctoCorrigido;
				vlTituloCorrigido = _fTrataConciliacao._vlTituloCorrigido;
				dtPagtoCorrigido = _fTrataConciliacao._dtPagtoCorrigido;
				blnTituloExcluido = _fTrataConciliacao._blnTituloExcluido;

				BD.iniciaTransacao();
				try
				{
					if (!ConciliacaoTituloDAO.trataTitulo(dtVenctoCorrigido,
														  vlTituloCorrigido,
														  dtPagtoCorrigido,
														  id,
														  blnTituloExcluido))
					{
						throw new Exception("Falha na tentativa de gravação do tratamento do título!!");
					}

					blnSucesso = true;
				}
				finally
				{
					if (blnSucesso)
					{
						BD.commitTransacao();
					}
					else
					{
						BD.rollbackTransacao();
					}
				}
				#endregion

				#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
				if (rowGridSelecionado != null)
				{
					if (dtPagtoCorrigido != DateTime.MinValue)
					{
						rowGridSelecionado.Cells["dt_pagto"].Value = Global.formataDataDdMmYyyyComSeparador(dtPagtoCorrigido);
					}
					else
					{
						rowGridSelecionado.Cells["dt_pagto"].Value = "";
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				avisoErro(ex.ToString());
			}
		}
		#endregion

		#region [ trataBotaoGeraPlanilha ]
		private void trataBotaoGeraPlanilha()
		{
			#region [ Declaração ]
			bool blnPossuiBoletosComInfoPagtoDivergentes = false;
			bool blnOcorreuErro;
			#endregion
			try
			{
				blnPossuiBoletosComInfoPagtoDivergentes = executaPesquisaBoletosComInfoPagtoDivergentes(out blnOcorreuErro);
				if (blnOcorreuErro)
				{
					throw new Exception("Ocorreu um erro ao consultar o banco de dados!!");
				}

				if (blnPossuiBoletosComInfoPagtoDivergentes)
				{
					geraPlanilhaExcel();
				}
				else
				{
					aviso("Não foi encontrado nenhum título que necessite de checagem manual!!");
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Ocorreu um erro ao executar o método trataBotaoGeraPlanilha()!! " + ex.StackTrace);
				avisoErro(ex.ToString());
			}
		}
		#endregion

		#region [ trataBotaoLocalizar ]
		private void trataBotaoLocalizar()
		{
			#region [ Declaração ]
			String strId;
			int intRowIndex;
			#endregion

			#region [ Consistência ]
			if (gridDados.Rows.Count == 0)
			{
				aviso("Não há dados a procurar!!");
				return;
			}

			if (txtIdSearch.Text.Length == 0)
			{
				aviso("Informe o ID do registro a ser localizado!!");
				return;
			}
			#endregion

			strId = txtIdSearch.Text.Trim();
			intRowIndex = -1;
			foreach (DataGridViewRow row in gridDados.Rows)
			{
				if (row.Cells[0].Value.ToString().Equals(strId))
				{
					intRowIndex = row.Index;
					break;
				}
			}

			if (intRowIndex == -1)
			{
				aviso("ID não localizado!!");
			}
			else
			{
				gridDados.Rows[intRowIndex].Selected = true;
				gridDados.FirstDisplayedScrollingRowIndex = intRowIndex;
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FTitulosConciliacao ]
		#region [ FTitulosConciliacao_FormClosing ]
		private void FTitulosConciliacao_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain._fMain.Location = this.Location;
			FMain._fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FTitulosConciliacao_KeyDown ]
		private void FTitulosConciliacao_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				e.SuppressKeyPress = true;
				executaPesquisa();
				return;
			}
		}
		#endregion

		#region [ FTitulosConciliacao_Shown ]
		private void FTitulosConciliacao_Shown(object sender, EventArgs e)
		{
			lblTotalizacaoRegistros.Text = "";
		}
		#endregion
		#endregion

		#region [ gridDados ]
		#region [ gridDados_KeyDown ]
		private void gridDados_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				trataOcorrenciaSelecionada();
				return;
			}
		}
		#endregion

		#region [ gridDados_DoubleClick ]
		private void gridDados_DoubleClick(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion
		#endregion

		#region [ Botões / Menu ]

		#region [ Pesquisar ]
		#region [ btnPesquisar_Click ]
		private void btnPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#region [ menuOcorrenciaPesquisar_Click ]
		private void menuOcorrenciaPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion
		#endregion

		#region [ Tratar Ocorrência ]
		#region [ btnOcorrenciaTratar_Click ]
		private void btnOcorrenciaTratar_Click(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion

		#region [ menuOcorrenciaTratar_Click ]
		private void menuOcorrenciaTratar_Click(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion
		#endregion

		#region [ Gera Excel ]
		#region [ btnExcel_Click ]
		private void btnExcel_Click(object sender, EventArgs e)
		{
			trataBotaoGeraPlanilha();
		}
		#endregion
		#endregion

		#region [ Procurar Registro ]
		#region [ btnLocalizar_Click ]
		private void btnLocalizar_Click(object sender, EventArgs e)
		{
			trataBotaoLocalizar();
		}
		#endregion
		#endregion

		#region [ txtIdSearch ]
		
		#region [ txtIdSearch_KeyDown ]
		private void txtIdSearch_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				trataBotaoLocalizar();
				return;
			}
		}
		#endregion

		#region [ txtIdSearch_Enter ]
		private void txtIdSearch_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtIdSearch_KeyPress ]
		private void txtIdSearch_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#endregion

		#endregion
	}
}
