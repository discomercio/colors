using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADM2
{
	public partial class FAtualizaPlanilhaEstoque : FModelo
	{
		#region [ Atributos ]
		private bool _InicializacaoOk;
		public bool inicializacaoOk
		{
			get { return _InicializacaoOk; }
		}

		private bool _OcorreuExceptionNaInicializacao;
		public bool ocorreuExceptionNaInicializacao
		{
			get { return _OcorreuExceptionNaInicializacao; }
		}

		public readonly string _AtualizaPlanilhaEstoqueExcelVisible = Global.GetConfigurationValue("AtualizaPlanilhaEstoqueExcelVisible");
		public readonly string _AtualizaCorFundoCellQtdeEstoque = Global.GetConfigurationValue("AtualizaCorFundoCellQtdeEstoque");

		private string _tituloBoxDisplayInformativo = "Mensagens Informativas";
		private int _qtdeMsgDisplayInformativo = 0;
		private string _tituloBoxDisplayErro = "Mensagens de Erro";
		private int _qtdeMsgDisplayErro = 0;
		#endregion

		#region [ Construtor ]
		public FAtualizaPlanilhaEstoque()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ adicionaDisplay ]
		private void adicionaDisplay(String mensagem)
		{
			String strMensagem;
			_qtdeMsgDisplayInformativo++;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			foreach (string linha in strMensagem.Split('\n'))
			{
				lbMensagem.Items.Add(linha);
			}
			lbMensagem.SelectedIndex = lbMensagem.Items.Count - 1;
			gboxMensagensInformativas.Text = _tituloBoxDisplayInformativo + "  (" + _qtdeMsgDisplayInformativo.ToString() + ")";
			Global.gravaLogAtividade(mensagem);
		}
		#endregion

		#region [ adicionaErro ]
		private void adicionaErro(String mensagem)
		{
			String strMensagem;
			_qtdeMsgDisplayErro++;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			foreach (string linha in strMensagem.Split('\n'))
			{
				lbErro.Items.Add(linha);
			}
			lbErro.SelectedIndex = lbErro.Items.Count - 1;
			gboxMsgErro.Text = _tituloBoxDisplayErro + "  (" + _qtdeMsgDisplayErro.ToString() + ")";
			Global.gravaLogAtividade("ERRO: " + mensagem);
		}
		#endregion

		#region [ pathPlanilhaEstoqueValorDefault ]
		private String pathPlanilhaEstoqueValorDefault()
		{
			String strResp = "";

			try
			{
				strResp = Path.GetPathRoot(Application.StartupPath);
			}
			catch (Exception)
			{
				strResp = "";
			}

			if (strResp.Length == 0) strResp = @"\";
			if (Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque))
				{
					strResp = Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNamePlanilhaEstoqueValorDefault ]
		private String fileNamePlanilhaEstoqueValorDefault()
		{
			String strResp = "";

			if ((Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque ?? "").Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque + "\\" + Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque))
				{
					strResp = Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque;
				}
			}
			return strResp;
		}
		#endregion

		#region [ limpaCamposMensagem ]
		private void limpaCamposMensagem()
		{
			lbMensagem.Items.Clear();
			_qtdeMsgDisplayInformativo = 0;
			gboxMensagensInformativas.Text = _tituloBoxDisplayInformativo;

			lbErro.Items.Clear();
			_qtdeMsgDisplayErro = 0;
			gboxMsgErro.Text = _tituloBoxDisplayErro;
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Planilha de estoque ]
			if (txtPlanilhaEstoque.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar a planilha de estoque que será processada!!");
				return false;
			}
			if (!File.Exists(txtPlanilhaEstoque.Text))
			{
				avisoErro("A planilha de estoque informada não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoAbrePlanilhaEstoque ]
		private void trataBotaoAbrePlanilhaEstoque()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "trataBotaoAbrePlanilhaEstoque()";
			string strNomeArqPlanilhaEstoque;
			#endregion

			strNomeArqPlanilhaEstoque = txtPlanilhaEstoque.Text.Trim();

			if (strNomeArqPlanilhaEstoque.Length == 0) return;

			if (!File.Exists(strNomeArqPlanilhaEstoque))
			{
				avisoErro("Arquivo '" + Path.GetFileName(strNomeArqPlanilhaEstoque) + "' não foi encontrado!!");
				return;
			}

			try
			{
				Process.Start(strNomeArqPlanilhaEstoque);
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar abrir planilha '" + Path.GetFileName(strNomeArqPlanilhaEstoque) + "'\r\n" + ex.ToString());
			}
		}
		#endregion

		#region [ trataBotaoSelecionaPlanilhaEstoque ]
		private void trataBotaoSelecionaPlanilhaEstoque()
		{
			#region [ Declarações ]
			DialogResult dr;
			#endregion

			try
			{
				openFileDialogCtrl.InitialDirectory = pathPlanilhaEstoqueValorDefault();
				openFileDialogCtrl.FileName = "";
				dr = openFileDialogCtrl.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialogCtrl.FileName.Length > 0) && (txtPlanilhaEstoque.Text.Length > 0))
				{
					if (openFileDialogCtrl.FileName.ToUpper().Equals(txtPlanilhaEstoque.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos de mensagens ]
				limpaCamposMensagem();
				#endregion

				txtPlanilhaEstoque.Text = openFileDialogCtrl.FileName;
				Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque = Path.GetDirectoryName(openFileDialogCtrl.FileName);
				Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque = Path.GetFileName(openFileDialogCtrl.FileName);
			}
			catch (Exception ex)
			{
				info(ModoExibicaoMensagemRodape.Normal);
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoAtualizaPlanilhaEstoque ]
		private void trataBotaoAtualizaPlanilhaEstoque()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FAtualizaPlanilhaEstoque.trataBotaoAtualizaPlanilhaEstoque";
			const double COR_FUNDO_INFO_NOT_FOUND = 12566463d;
			const double COR_FUNDO_LARANJA = 9420794d;
			const double COR_FUNDO_VERDE = 10213316d;
			const double COR_FUNDO_AZUL = 14136213d;
			const double COR_FUNDO_ROXO = 13082801d;
			string strAux;
			string strMsg;
			string strMsgErro;
			string strMsgErroAux;
			string strMsgErroLog = "";
			string strValue;
			string strNomeArqPlanilhaEstoque;
			string strPathPlanilhaEstoqueBackup;
			string strNomeArqPlanilhaEstoqueBackup;
			string strTituloEncontrado;
			string strTituloEncontradoUppercase;
			string strCelQtdeEstoqueVenda;
			string strCelVlCustoIntermediario;
			string strCellPrecoLista;
			StringBuilder sbMsg;
			StringBuilder sbMsgErro;
			int iNumLinha;
			int iXlDadosMinIndex;
			int iXlDadosMaxIndex;
			int iColIndex;
			int qtdeColObrigatoriaEncontrada;
			int qtdeLinhasVaziasConsecutivas;
			int qtdeLinhaDadosPlanilhaEstoque;
			int qtdeIsVendavelTrue;
			int qtdeProdutoCompostoItem;
			int qtdeEstoqueAcumulado;
			decimal vlNovoCustoIntermediario;
			bool blnTitulosOk;
			bool blnAdicionaErro;
			bool blnAchou;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			object oXL = null;
			object oWBs = null;
			object oWB = null;
			object oWSs = null;
			object oWS = null;
			object oRange = null;
			object[,] oRangeValue = null;
			object oCellQtdeEstoqueVenda = null;
			object oCellVlCustoIntermediario = null;
			object oCellPrecoLista = null;
			object oCellInteriorQtdeEstoqueVenda = null;
			List<PlanilhaEstoqueColumn> vPlanilhaEstoqueColunasObrigatorias;
			PlanilhaEstoqueHeader planilhaEstoqueHeader;
			PlanilhaEstoqueLinha linhaPlanilhaEstoque;
			ProdutoEstoqueVenda produtoEstoqueVenda;
			ProdutoEstoqueVendaLojaConsolidado produtoConsolidado;
			ProdutoLoja produtoLoja;
			ProdutoUnificadoPrecoLista produtoUnificadoPrecoLista;
			ProdutoUnificadoPrecoLista produtoConsolidadoPrecoLista;
			List<ProdutoCompostoItem> vProdutoCompostoItem;
			Log log = new Log();
			Process[] processosAnterior;
			Process[] processosAtual;
			#endregion

			#region [ Observações Importantes ]
			// Observações:
			// ============
			// 1) Todas as referências ao Excel devem ser devidamentes desalocadas, senão o processo do Excel não é encerrado ao final.
			//    Se uma variável for ser reutilizada para acessar outro objeto, como é o caso de 'range' por exemplo, antes de atribuir as novas referências, as anteriores devem ser desalocadas.
			//    Os comandos para desalocar as referências foram encapsuladas na rotina ExcelAutomation.NAR(), seguindo orientações do artigo https://support.microsoft.com/en-us/kb/317109
			// 2) O comando p/ maximizar a janela do Excel deve ser evitado porque senão o processo do Excel não é encerrado ao final, mesmo executando os comandos p/ desalocar as referências:
			//    ExcelAutomation.SetProperty(oXL, "WindowState", ExcelAutomation.XlWindowState.xlMaximized);
			// 3) Após realizar alterações nesta rotina, deve-se verificar se o processo do Excel está sendo encerrado ao final ou se está ficando pendente.
			#endregion

			limpaCamposMensagem();

			#region [ Obtém nome do arquivo da planilha de estoque ]
			strNomeArqPlanilhaEstoque = txtPlanilhaEstoque.Text;
			#endregion

			#region [ Consistências ]
			if (strNomeArqPlanilhaEstoque.Length == 0)
			{
				strMsgErro = "É necessário selecionar a planilha do estoque a ser processada!!";
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}

			if (!File.Exists(strNomeArqPlanilhaEstoque))
			{
				strMsgErro = "O arquivo da planilha do estoque não existe!!\r\n" + strNomeArqPlanilhaEstoque;
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}

			if (Global.IsFileLocked(strNomeArqPlanilhaEstoque))
			{
				strMsgErro = "A planilha do estoque '" + Path.GetFileName(strNomeArqPlanilhaEstoque) + "' está aberta e em uso!!\r\nNão é possível prosseguir com o processamento!!";
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma o processamento da planilha?"))
			{
				adicionaErro("Operação cancelada!");
				return;
			}
			#endregion

			#region [ Memoriza os processos de Excel em execução ]
			processosAnterior = Process.GetProcessesByName("EXCEL");
			sbMsg = new StringBuilder("");
			foreach (Process procAnterior in processosAnterior)
			{
				if (sbMsg.Length > 0) sbMsg.Append(", ");
				sbMsg.Append(procAnterior.Id.ToString());
			}
			strMsg = NOME_DESTA_ROTINA + ": processos do Excel em execução antes do início do processamento (PID = " + (sbMsg.Length == 0 ? "(nenhum)" : sbMsg.ToString()) + ")";
			Global.gravaLogAtividade(strMsg);
			#endregion

			#region [ Inicialização do processamento ]
			dtInicioProcessamento = DateTime.Now;
			strMsg = "Início do processamento\r\n" +
					"        Planilha do estoque: " + strNomeArqPlanilhaEstoque;
			adicionaDisplay(strMsg);
			#endregion

			#region [ Backup da planilha do estoque ]
			info(ModoExibicaoMensagemRodape.EmExecucao, "Criando backup da planilha do estoque");
			strPathPlanilhaEstoqueBackup = Path.GetDirectoryName(strNomeArqPlanilhaEstoque) + "\\Backup";
			try
			{
				if (!Directory.Exists(strPathPlanilhaEstoqueBackup))
				{
					Directory.CreateDirectory(strPathPlanilhaEstoqueBackup);
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha ao tentar criar o diretório para armazenar o backup da planilha do estoque!!\r\n" + ex.ToString());
			}

			// Se não conseguiu criar o diretório, grava os arquivos no mesmo diretório em que está a planilha do estoque
			if (!Directory.Exists(strPathPlanilhaEstoqueBackup)) strPathPlanilhaEstoqueBackup = Path.GetDirectoryName(strNomeArqPlanilhaEstoque);

			strNomeArqPlanilhaEstoqueBackup = strPathPlanilhaEstoqueBackup + "\\" +
												Path.GetFileNameWithoutExtension(strNomeArqPlanilhaEstoque) +
												"_" + Global.formataDataYyyyMmDdComSeparador(DateTime.Now, "-") +
												"_" + Global.formataHoraHhMmSsComSimbolo(DateTime.Now) +
												Path.GetExtension(strNomeArqPlanilhaEstoque);
			File.Copy(strNomeArqPlanilhaEstoque, strNomeArqPlanilhaEstoqueBackup);
			if (!File.Exists(strNomeArqPlanilhaEstoqueBackup))
			{
				adicionaErro("Falha ao tentar criar a cópia de backup do arquivo da planilha do estoque");
			}
			else
			{
				strMsg = "Backup da planilha do estoque realizado com sucesso: " + strNomeArqPlanilhaEstoqueBackup;
				adicionaDisplay(strMsg);
			}
			#endregion

			try // Try-Catch-Finally
			{
				try // Try-Finally
				{
					strMsg = "Inicialização do processo do Excel para processamento da planilha do estoque";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Atualização dos dados na planilha do estoque");

					#region [ Instancia o Excel ]
					try
					{
						oXL = ExcelAutomation.CriaInstanciaExcel();
					}
					catch (Exception ex)
					{
						strMsg = "Falha ao acionar o Excel!!\nVerifique se o Excel está instalado!!\n\n" + ex.ToString();
						adicionaErro(strMsg);
						avisoErro(strMsg);
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\r\n" + ex.ToString());
						return;
					}
					#endregion

					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.Visible, (_AtualizaPlanilhaEstoqueExcelVisible.Equals("0") ? false : true));
					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.DisplayAlerts, false);
					oWBs = ExcelAutomation.GetProperty(oXL, ExcelAutomation.PropertyType.Workbooks);
					ExcelAutomation.InvokeMethod(oWBs, ExcelAutomation.MethodType.Open, strNomeArqPlanilhaEstoque);
					oWB = ExcelAutomation.GetProperty(oWBs, ExcelAutomation.PropertyType.Item, 1);
					oWSs = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.Worksheets, null);
					oWS = ExcelAutomation.GetProperty(oWSs, ExcelAutomation.PropertyType.Item, 1);
					ExcelAutomation.InvokeMethod(oWS, ExcelAutomation.MethodType.Select, null);

					#region [ Obtém a linha dos títulos da planilha ]
					// Analisa somente até a coluna 'Custo Médio', os campos seguintes são desconsiderados por serem campos calculados ou de uso específico do usuário
					iXlDadosMinIndex = 1;
					iXlDadosMaxIndex = 64;

					iNumLinha = 1;
					strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
					oRange = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strAux);
					oRangeValue = (object[,])ExcelAutomation.GetProperty(oRange, ExcelAutomation.PropertyType.Value);
					#endregion

					#region [ Colunas obrigatórias ]

					#region [ Obtém a posição das colunas obrigatórias ]
					vPlanilhaEstoqueColunasObrigatorias = new List<PlanilhaEstoqueColumn>();
					planilhaEstoqueHeader = new PlanilhaEstoqueHeader();

					planilhaEstoqueHeader.Sku.ColTitleEsperado = "C.UNIF";
					vPlanilhaEstoqueColunasObrigatorias.Add(planilhaEstoqueHeader.Sku);

					planilhaEstoqueHeader.ProdutoDescricao.ColTitleEsperado = "Produto";
					vPlanilhaEstoqueColunasObrigatorias.Add(planilhaEstoqueHeader.ProdutoDescricao);

					planilhaEstoqueHeader.QtdeEstoque.ColTitleEsperado = "Qtd";
					vPlanilhaEstoqueColunasObrigatorias.Add(planilhaEstoqueHeader.QtdeEstoque);

					planilhaEstoqueHeader.ValorCustoMedio.ColTitleEsperado = "Custo Médio";
					vPlanilhaEstoqueColunasObrigatorias.Add(planilhaEstoqueHeader.ValorCustoMedio);

					planilhaEstoqueHeader.PrecoLista.ColTitleEsperado = "Valor AV";
					vPlanilhaEstoqueColunasObrigatorias.Add(planilhaEstoqueHeader.PrecoLista);

					qtdeColObrigatoriaEncontrada = 0;
					iColIndex = oRangeValue.GetLowerBound(1) - 1;
					while (true)
					{
						iColIndex++;
						if (iColIndex > oRangeValue.GetUpperBound(1)) break;

						if (oRangeValue[1, iColIndex] == null) continue;

						strTituloEncontrado = oRangeValue[1, iColIndex].ToString().Trim();
						strTituloEncontradoUppercase = strTituloEncontrado.ToUpper();

						foreach (var colObrigatoria in vPlanilhaEstoqueColunasObrigatorias)
						{
							if (strTituloEncontradoUppercase.Equals(colObrigatoria.ColTitleEsperado.ToUpper()))
							{
								qtdeColObrigatoriaEncontrada++;
								colObrigatoria.ColIndex = iColIndex;
								colObrigatoria.ColTitle = strTituloEncontrado;
								break;
							}
						}

						// Já localizou todas as colunas obrigatórias?
						if (qtdeColObrigatoriaEncontrada == vPlanilhaEstoqueColunasObrigatorias.Count) break;
					} // while (true)
					#endregion

					#region [ Verifica se alguma coluna obrigatória não foi localizada ]
					blnTitulosOk = true;
					sbMsgErro = new StringBuilder("");
					foreach (var colObrigatoria in vPlanilhaEstoqueColunasObrigatorias)
					{
						if (colObrigatoria.ColIndex == 0)
						{
							blnTitulosOk = false;
							strMsg = "Coluna '" + colObrigatoria.ColTitleEsperado + "' não encontrada";
							sbMsgErro.AppendLine(strMsg);
						}
					}

					if (!blnTitulosOk)
					{
						strMsgErro = "A planilha '" + Path.GetFileName(strNomeArqPlanilhaEstoque) + "' não possui os títulos corretos para as colunas!!\r\nVerifique se a planilha correta foi selecionada!!\r\n\r\n" + sbMsgErro.ToString();
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#endregion

					#region [ Consolida dados da planilha de estoque ]
					strMsg = "Iniciando processamento de atualização da planilha do estoque";
					adicionaDisplay(strMsg);

					qtdeLinhasVaziasConsecutivas = 0;
					qtdeLinhaDadosPlanilhaEstoque = 0;
					iXlDadosMaxIndex = 0;
					foreach (var item in vPlanilhaEstoqueColunasObrigatorias)
					{
						if (item.ColIndex > iXlDadosMaxIndex) iXlDadosMaxIndex = item.ColIndex;
					}

					while (true)
					{
						if (qtdeLinhasVaziasConsecutivas >= 10) break;

						#region [ Recupera array de objetos com linha de dados da planilha ]
						iNumLinha++;
						strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
						// Desaloca as referências anteriores, caso contrário, o processo do Excel não é encerrado ao final
						ExcelAutomation.NAR(oRange);
						oRange = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strAux);
						oRangeValue = (object[,])ExcelAutomation.GetProperty(oRange, ExcelAutomation.PropertyType.Value);
						if (oRangeValue[1, planilhaEstoqueHeader.Sku.ColIndex] == null)
						{
							qtdeLinhasVaziasConsecutivas++;
							continue;
						}

						if (oRangeValue[1, planilhaEstoqueHeader.Sku.ColIndex].ToString().Trim().Length == 0)
						{
							qtdeLinhasVaziasConsecutivas++;
							continue;
						}
						#endregion

						qtdeLinhaDadosPlanilhaEstoque++;

						// Encontrou linha com dados, portanto, zera contador de linhas vazias consecutivas
						qtdeLinhasVaziasConsecutivas = 0;
						strMsg = "Processamento da " + qtdeLinhaDadosPlanilhaEstoque.ToString() + "ª linha de dados da planilha do estoque: SKU " + oRangeValue[1, planilhaEstoqueHeader.Sku.ColIndex].ToString();
						adicionaDisplay(strMsg);
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);

						#region [ Carrega dados para objeto da classe LinhaPlanilhaEstoque ]
						linhaPlanilhaEstoque = new PlanilhaEstoqueLinha();

						#region [ SKU ]
						linhaPlanilhaEstoque.Sku = oRangeValue[1, planilhaEstoqueHeader.Sku.ColIndex].ToString().Trim();
						linhaPlanilhaEstoque.SkuFormatado = Global.normalizaCodigoProduto(linhaPlanilhaEstoque.Sku);
						#endregion

						#region [ Descrição do produto ]
						if (oRangeValue[1, planilhaEstoqueHeader.ProdutoDescricao.ColIndex] != null)
						{
							linhaPlanilhaEstoque.ProdutoDescricao = oRangeValue[1, planilhaEstoqueHeader.ProdutoDescricao.ColIndex].ToString().Trim();
						}
						#endregion

						#region [ Quantidade no estoque ]
						if (oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaEstoque.QtdeEstoque = (double)oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex];
								}
								else if (oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									linhaPlanilhaEstoque.QtdeEstoque = (int)oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaEstoque.QtdeEstoque = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = "Falha ao ler o conteúdo da célula '" + planilhaEstoqueHeader.QtdeEstoque.ColTitle + "' do SKU " + linhaPlanilhaEstoque.Sku + " (Type = " + oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaEstoqueHeader.QtdeEstoque.ColIndex].ToString() + ")";
								adicionaErro(strMsg);
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
							}
						}
						#endregion

						#region [ Valor do custo médio ]
						if (oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaEstoque.ValorCustoMedio = (double)oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex];
								}
								else if (oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaEstoque.ValorCustoMedio = (double)(decimal)oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex];

								}
								else
								{
									strValue = oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaEstoque.ValorCustoMedio = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								blnAdicionaErro = true;
								if (oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel exibe #DIV/0!, #NÚM! ou #REF!, situações em que o valor da célula é do tipo System.Int32 e valores como -2146826281, -2146826252, etc
									if ((int)oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex] < 0) blnAdicionaErro = false;
								}
								if (oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel contém anotações do usuário
									if (Global.contagemLetras(oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].ToString()) > 3) blnAdicionaErro = false;
								}
								if (blnAdicionaErro)
								{
									strMsg = "Falha ao ler o conteúdo da célula '" + planilhaEstoqueHeader.ValorCustoMedio.ColTitle + "' do SKU " + linhaPlanilhaEstoque.Sku + " (Type = " + oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaEstoqueHeader.ValorCustoMedio.ColIndex].ToString() + ")";
									adicionaErro(strMsg);
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
								}
							}
						}
						#endregion

						#region [ Valor AV (Preço de Lista) ]
						if (oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaEstoque.PrecoLista = (double)oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex];
								}
								else if (oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaEstoque.PrecoLista = (double)(decimal)oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaEstoque.PrecoLista = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								blnAdicionaErro = true;
								if (oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel exibe #DIV/0!, #NÚM! ou #REF!, situações em que o valor da célula é do tipo System.Int32 e valores como -2146826281, -2146826252, etc
									if ((int)oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex] < 0) blnAdicionaErro = false;
								}
								if (oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel contém anotações do usuário
									if (Global.contagemLetras(oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].ToString()) > 3) blnAdicionaErro = false;
								}
								if (blnAdicionaErro)
								{
									strMsg = "Falha ao ler o conteúdo da célula '" + planilhaEstoqueHeader.PrecoLista.ColTitle + "' do SKU " + linhaPlanilhaEstoque.Sku + " (Type = " + oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaEstoqueHeader.PrecoLista.ColIndex].ToString() + ")";
									adicionaErro(strMsg);
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
								}
							}
						}
						#endregion

						#endregion

						#region [ Obtém quantidade disponível no estoque, preço de custo médio e preço de lista ]
						produtoConsolidado = new ProdutoEstoqueVendaLojaConsolidado();
						produtoConsolidadoPrecoLista = new ProdutoUnificadoPrecoLista();

						vlNovoCustoIntermediario = 0m;

						foreach (var ambiente in FMain.contextoBD.Ambientes)
						{
							produtoUnificadoPrecoLista = ambiente.produtoDAO.GetProdutoUnificadoPrecoLista(linhaPlanilhaEstoque.SkuFormatado, ambiente.NumeroLoja, out strMsgErroAux);
							if (produtoUnificadoPrecoLista != null)
							{
								if ((produtoUnificadoPrecoLista.produto.Trim().Length > 0) && (produtoUnificadoPrecoLista.isCadastrado))
								{
									if (produtoConsolidadoPrecoLista.produto.Trim().Length == 0)
									{
										produtoConsolidadoPrecoLista.fabricante = produtoUnificadoPrecoLista.fabricante;
										produtoConsolidadoPrecoLista.produto = produtoUnificadoPrecoLista.produto;
										produtoConsolidadoPrecoLista.loja = produtoUnificadoPrecoLista.loja;
									}

									// Importante: um produto pode não estar cadastrado como produto comum ou produto composto em algum dos ambientes
									produtoConsolidadoPrecoLista.isCadastrado = (produtoConsolidadoPrecoLista.isCadastrado || produtoUnificadoPrecoLista.isCadastrado);
									produtoConsolidadoPrecoLista.isComposto = (produtoConsolidadoPrecoLista.isComposto || produtoUnificadoPrecoLista.isComposto);

									// Caso o produto não esteja cadastrado em algum dos ambientes, a descrição pode ser vazia ou um caracter '.', etc
									if ((produtoUnificadoPrecoLista.descricao ?? "").Length > 0)
									{
										if (produtoUnificadoPrecoLista.descricao.Length > produtoConsolidadoPrecoLista.descricao.Length) produtoConsolidadoPrecoLista.descricao = produtoUnificadoPrecoLista.descricao;
									}

									// Usa o maior preço de lista entre os ambientes
									if (produtoUnificadoPrecoLista.preco_lista > produtoConsolidadoPrecoLista.preco_lista) produtoConsolidadoPrecoLista.preco_lista = produtoUnificadoPrecoLista.preco_lista;
								}
							}

							produtoEstoqueVenda = ambiente.produtoDAO.GetProdutoUnificadoCustoIntermediario(linhaPlanilhaEstoque.SkuFormatado, out strMsgErroAux);
							if (produtoEstoqueVenda != null)
							{
								if ((produtoEstoqueVenda.produto.Trim().Length > 0) && (produtoEstoqueVenda.isCadastrado))
								{
									if (produtoConsolidado.produto.Trim().Length == 0)
									{
										produtoConsolidado.fabricante = produtoEstoqueVenda.fabricante;
										produtoConsolidado.produto = produtoEstoqueVenda.produto;
									}

									// Importante: um produto pode não estar cadastrado como produto comum ou produto composto em algum dos ambientes
									produtoConsolidado.isCadastrado = (produtoConsolidado.isCadastrado || produtoEstoqueVenda.isCadastrado);
									produtoConsolidado.isComposto = (produtoConsolidado.isComposto || produtoEstoqueVenda.isComposto);

									// Caso o produto não esteja cadastrado em algum dos ambientes, a descrição pode ser vazia ou um caracter '.', etc
									if ((produtoEstoqueVenda.descricao ?? "").Length > 0)
									{
										if (produtoEstoqueVenda.descricao.Length > produtoConsolidado.descricao.Length) produtoConsolidado.descricao = produtoEstoqueVenda.descricao;
									}

									qtdeEstoqueAcumulado = produtoConsolidado.qtdeEstoqueVenda + produtoEstoqueVenda.qtdeEstoqueVenda;
									if (qtdeEstoqueAcumulado == 0m)
									{
										vlNovoCustoIntermediario = 0m;
									}
									else
									{
										vlNovoCustoIntermediario = ((produtoConsolidado.qtdeEstoqueVenda * produtoConsolidado.vlCustoIntermediario) + (produtoEstoqueVenda.qtdeEstoqueVenda * produtoEstoqueVenda.vlCustoIntermediario)) / qtdeEstoqueAcumulado;
									}
									produtoConsolidado.vlCustoIntermediario = vlNovoCustoIntermediario;
									produtoConsolidado.qtdeEstoqueVenda += produtoEstoqueVenda.qtdeEstoqueVenda;

									if (produtoEstoqueVenda.qtdeEstoqueVenda > 0) produtoConsolidado.QtdeAmbientesComEstoqueDisponivel++;
									produtoConsolidado.listaEstoqueVendaAmbiente.Add(new EstoqueVendaAmbiente(produtoEstoqueVenda.qtdeEstoqueVenda, produtoEstoqueVenda.vlCustoIntermediario, ambiente.NomeAmbiente));
									Global.gravaLogAtividade("[" + ambiente.NomeAmbiente + "] SKU " + linhaPlanilhaEstoque.SkuFormatado + ": Qtde estoque = " + Global.formataInteiro(produtoEstoqueVenda.qtdeEstoqueVenda) + ", Custo intermediário = " + Global.formataMoeda(produtoEstoqueVenda.vlCustoIntermediario));

									#region [ Verifica se o produto é 'vendável' ]
									if (produtoEstoqueVenda.isComposto)
									{
										qtdeIsVendavelTrue = 0;
										qtdeProdutoCompostoItem = 0;
										vProdutoCompostoItem = ambiente.produtoDAO.GetProdutoCompostoItem(produtoEstoqueVenda.fabricante, produtoEstoqueVenda.produto, out strMsgErroAux);
										if (vProdutoCompostoItem != null)
										{
											foreach (var item in vProdutoCompostoItem)
											{
												produtoLoja = ambiente.produtoDAO.GetProdutoLoja(item.fabricante_item, item.produto_item, ambiente.NumeroLoja, out strMsgErroAux);
												if (produtoLoja != null)
												{
													if ((produtoLoja.produto ?? "").Length > 0)
													{
														qtdeProdutoCompostoItem++;
														if (produtoLoja.vendavel.ToUpper().Equals("S")) qtdeIsVendavelTrue++;
													}
												}
											}
										}
										// Todos os itens da composição precisam estar como 'vendável'
										if ((qtdeIsVendavelTrue == qtdeProdutoCompostoItem) && (qtdeProdutoCompostoItem > 0))
										{
											// Se o produto estiver como 'vendável' em pelo menos um dos ambientes, ele será considerado 'vendável'
											produtoConsolidado.isVendavel = true;
										}
									}
									else
									{
										produtoLoja = ambiente.produtoDAO.GetProdutoLoja(produtoConsolidado.fabricante, produtoConsolidado.produto, ambiente.NumeroLoja, out strMsgErroAux);
										if (produtoLoja != null)
										{
											if ((produtoLoja.produto ?? "").Length > 0)
											{
												if (produtoLoja.vendavel.ToUpper().Equals("S"))
												{
													// Se o produto estiver como 'vendável' em pelo menos um dos ambientes, ele será considerado 'vendável'
													produtoConsolidado.isVendavel = true;
												}
											}
										}
									}
									#endregion
								}
							}
						} // foreach (var ambiente in FMain.contextoBD.Ambientes)

						if (produtoConsolidado.isCadastrado) produtoConsolidado.vlCustoIntermediario = Global.arredondaParaMonetario(produtoConsolidado.vlCustoIntermediario);
						#endregion

						#region [ Atualiza dados na planilha da quantidade no estoque e do custo médio ]
						// 18/01/2019: por determinação do Carlos, ignorar o flag que indica se o produto é vendável ou não
						if (produtoConsolidado.isCadastrado && produtoConsolidadoPrecoLista.isCadastrado)
						{
							#region [ Produto cadastrado ]

							#region [ Qtde estoque venda ]
							strCelQtdeEstoqueVenda = Global.excel_converte_numeracao_digito_para_letra(planilhaEstoqueHeader.QtdeEstoque.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellQtdeEstoqueVenda);
							oCellQtdeEstoqueVenda = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelQtdeEstoqueVenda);
							ExcelAutomation.NAR(oCellInteriorQtdeEstoqueVenda);
							oCellInteriorQtdeEstoqueVenda = ExcelAutomation.GetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Interior);
							// Atualiza valor
							ExcelAutomation.SetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Value, (double)produtoConsolidado.qtdeEstoqueVenda);

							if (!_AtualizaCorFundoCellQtdeEstoque.Equals("0"))
							{
								if ((produtoConsolidado.qtdeEstoqueVenda == 0) || (produtoConsolidado.QtdeAmbientesComEstoqueDisponivel == 0))
								{
									#region [ Cor de fundo vazio ]
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlNone);
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0);
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0);
									#endregion
								}
								else if (produtoConsolidado.QtdeAmbientesComEstoqueDisponivel == 1)
								{
									#region [ Define cor de fundo de forma a identificar o único ambiente que possui estoque disponível ]
									foreach (EstoqueVendaAmbiente item in produtoConsolidado.listaEstoqueVendaAmbiente)
									{
										if (item.qtdeEstoqueVenda > 0)
										{
											if (item.nomeAmbiente.ToUpper().Equals("OLD01"))
											{
												#region [ Fundo alaranjado ]
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Color, COR_FUNDO_LARANJA);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0d);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
												#endregion
											}
											else if (item.nomeAmbiente.ToUpper().Equals("DIS"))
											{
												#region [ Fundo esverdeado ]
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Color, COR_FUNDO_VERDE);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0d);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
												#endregion
											}
											else
											{
												#region [ Ambiente desconhecido: fundo roxo ]
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Color, COR_FUNDO_ROXO);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0d);
												ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
												#endregion
											}

											break;
										}
									}
									#endregion
								}
								else
								{
									#region [ Define cor de fundo indicando que há disponibilidade de estoque em mais do que um ambiente ]
									#region [ Fundo azulado ]
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Color, COR_FUNDO_AZUL);
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0d);
									ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
									#endregion
									#endregion
								}
							} // if (!_AtualizaCorFundoCellQtdeEstoque.Equals("0"))
							#endregion

							#region [ Custo médio ]
							strCelVlCustoIntermediario = Global.excel_converte_numeracao_digito_para_letra(planilhaEstoqueHeader.ValorCustoMedio.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellVlCustoIntermediario);
							oCellVlCustoIntermediario = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelVlCustoIntermediario);
							if (produtoConsolidado.qtdeEstoqueVenda > 0)
							{
								ExcelAutomation.SetProperty(oCellVlCustoIntermediario, ExcelAutomation.PropertyType.Value, (double)produtoConsolidado.vlCustoIntermediario);
							}
							else
							{
								ExcelAutomation.SetProperty(oCellVlCustoIntermediario, ExcelAutomation.PropertyType.Value, 0d);
							}
							#endregion

							#region [ Preço de lista ]
							strCellPrecoLista = Global.excel_converte_numeracao_digito_para_letra(planilhaEstoqueHeader.PrecoLista.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellPrecoLista);
							oCellPrecoLista = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCellPrecoLista);
							if (produtoConsolidado.qtdeEstoqueVenda > 0)
							{
								ExcelAutomation.SetProperty(oCellPrecoLista, ExcelAutomation.PropertyType.Value, (double)produtoConsolidadoPrecoLista.preco_lista);
							}
							else
							{
								ExcelAutomation.SetProperty(oCellPrecoLista, ExcelAutomation.PropertyType.Value, 0d);
							}
							#endregion

							#endregion
						}
						else
						{
							#region [ Produto não cadastrado ]
							// Se o produto não está cadastrado/disponível no sistema, zera o estoque na planilha
							strCelQtdeEstoqueVenda = Global.excel_converte_numeracao_digito_para_letra(planilhaEstoqueHeader.QtdeEstoque.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellQtdeEstoqueVenda);
							oCellQtdeEstoqueVenda = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelQtdeEstoqueVenda);
							ExcelAutomation.NAR(oCellInteriorQtdeEstoqueVenda);
							oCellInteriorQtdeEstoqueVenda = ExcelAutomation.GetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Interior);
							// Atualiza valor: zero
							ExcelAutomation.SetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Value, 0d);

							if (!_AtualizaCorFundoCellQtdeEstoque.Equals("0"))
							{
								#region [ Altera cor de fundo p/ indicar que não está disponível ]
								ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
								ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
								ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
								ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0d);
								ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
								#endregion
							} // if (!_AtualizaCorFundoCellQtdeEstoque.Equals("0"))
							#endregion

							#region [ Custo médio ]
							strCelVlCustoIntermediario = Global.excel_converte_numeracao_digito_para_letra(planilhaEstoqueHeader.ValorCustoMedio.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellVlCustoIntermediario);
							oCellVlCustoIntermediario = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelVlCustoIntermediario);
							ExcelAutomation.SetProperty(oCellVlCustoIntermediario, ExcelAutomation.PropertyType.Value, 0d);
							#endregion

							#region [ Preço de lista ]
							strCellPrecoLista = Global.excel_converte_numeracao_digito_para_letra(planilhaEstoqueHeader.PrecoLista.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellPrecoLista);
							oCellPrecoLista = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCellPrecoLista);
							ExcelAutomation.SetProperty(oCellPrecoLista, ExcelAutomation.PropertyType.Value, 0d);
							#endregion
						}
						#endregion
					} // while (true)

					strMsg = qtdeLinhaDadosPlanilhaEstoque.ToString() + " linhas de dados processadas na planilha de estoque";
					adicionaDisplay(strMsg);
					#endregion

					#region [ Salva planilha de estoque ]
					strMsg = "Salvando planilha de estoque";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Salvando planilha de estoque");

					ExcelAutomation.InvokeMethod(oWB, ExcelAutomation.MethodType.Save, null);
					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.DisplayAlerts, true);
					#endregion
				}
				finally
				{
					strMsg = "Fechamento da planilha do estoque e encerramento do processo do Excel";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Fechamento da planilha do estoque e encerramento do processo do Excel");
					try
					{
						ExcelAutomation.NAR(oCellInteriorQtdeEstoqueVenda);
						ExcelAutomation.NAR(oCellQtdeEstoqueVenda);
						ExcelAutomation.NAR(oCellVlCustoIntermediario);
						ExcelAutomation.NAR(oCellPrecoLista);

						ExcelAutomation.NAR(oRange);

						ExcelAutomation.NAR(oWS);
						ExcelAutomation.NAR(oWSs);

						if (oWB != null)
						{
							ExcelAutomation.InvokeMethod(oWB, ExcelAutomation.MethodType.Close, 0);
						}

						ExcelAutomation.NAR(oWB);
						ExcelAutomation.NAR(oWBs);

						if (oXL != null)
						{
							ExcelAutomation.InvokeMethod(oXL, ExcelAutomation.MethodType.Quit, null);
							Thread.Sleep(2000);
						}
						ExcelAutomation.NAR(oXL);
						Thread.Sleep(1000);
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(ex.ToString());
					}
				}

				#region [ Verifica se algum processo do Excel ficou pendente ]
				processosAtual = Process.GetProcessesByName("EXCEL");
				sbMsg = new StringBuilder("");
				foreach (Process procAtual in processosAtual)
				{
					if (sbMsg.Length > 0) sbMsg.Append(", ");
					sbMsg.Append(procAtual.Id.ToString());
				}
				strMsg = NOME_DESTA_ROTINA + ": processos do Excel em execução após a finalização do processamento (PID = " + (sbMsg.Length == 0 ? "(nenhum)" : sbMsg.ToString()) + ")";
				Global.gravaLogAtividade(strMsg);

				foreach (Process procAtual in processosAtual)
				{
					blnAchou = false;
					foreach (Process procAnterior in processosAnterior)
					{
						if (procAtual.Id == procAnterior.Id)
						{
							blnAchou = true;
							break;
						}
					}

					if (!blnAchou)
					{
						#region [ Encerra processo ]
						try
						{
							procAtual.Kill();
							Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": processo do Excel encerrado forçadamente (PID=" + procAtual.Id.ToString() + ")");
						}
						catch (Exception ex)
						{
							Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar encerrar forçadamente o processo PID=" + procAtual.Id.ToString() + "\r\n" + ex.ToString());
						}
						#endregion
					}
				}
				#endregion

				tsDuracaoProcessamento = DateTime.Now - dtInicioProcessamento;

				#region [ Grava log ]
				strMsg = "Sucesso no processamento da planilha de estoque (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + "): planilha de estoque = " + strNomeArqPlanilhaEstoque + " (" + qtdeLinhaDadosPlanilhaEstoque.ToString() + " linhas de dados)";
				Global.gravaLogAtividade(strMsg);
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.PROCESSAMENTO_PLANILHA_ESTOQUE;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				#region [ Mensagem de sucesso ]
				info(ModoExibicaoMensagemRodape.Normal);
				strMsg = "Processamento concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!!";
				adicionaDisplay(strMsg);
				strMsg = "Processamento concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!!" +
						"\r\n\r\nDeseja abrir a planilha '" + Path.GetFileName(strNomeArqPlanilhaEstoque) + "' agora?";
				if (confirma(strMsg))
				{
					try
					{
						Process.Start(strNomeArqPlanilhaEstoque);
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar abrir planilha '" + Path.GetFileName(strNomeArqPlanilhaEstoque) + "'\r\n" + ex.ToString());
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				adicionaErro(ex.Message);
				avisoErro(ex.ToString());
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FAtualizaPlanilhaEstoque ]

		#region [ FAtualizaPlanilhaEstoque_Load ]
		private void FAtualizaPlanilhaEstoque_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				txtPlanilhaEstoque.Text = "";
				limpaCamposMensagem();
				blnSucesso = true;
			}
			catch (Exception ex)
			{
				_OcorreuExceptionNaInicializacao = true;
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				if (!blnSucesso) Close();
			}
		}
		#endregion

		#region [ FAtualizaPlanilhaEstoque_Shown ]
		private void FAtualizaPlanilhaEstoque_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			string strFileNameArquivoPlanilhaEstoque;
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					strFileNameArquivoPlanilhaEstoque = pathPlanilhaEstoqueValorDefault() + "\\" + fileNamePlanilhaEstoqueValorDefault();
					if (File.Exists(strFileNameArquivoPlanilhaEstoque)) txtPlanilhaEstoque.Text = strFileNameArquivoPlanilhaEstoque;

					openFileDialogCtrl.InitialDirectory = pathPlanilhaEstoqueValorDefault();
					openFileDialogCtrl.FileName = "";

					_InicializacaoOk = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				_OcorreuExceptionNaInicializacao = true;
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
				// Se não inicializou corretamente, assegura-se de que o painel será fechado
				if (!_InicializacaoOk) Close();
			}
		}
		#endregion

		#region [ FAtualizaPlanilhaEstoque_FormClosing ]
		private void FAtualizaPlanilhaEstoque_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaPlanilhaEstoque ]

		#region [ btnSelecionaPlanilhaEstoque_Click ]
		private void btnSelecionaPlanilhaEstoque_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaPlanilhaEstoque();
		}
		#endregion

		#endregion

		#region [ txtPlanilhaEstoque ]

		#region [ txtPlanilhaEstoque_Enter ]
		private void txtPlanilhaEstoque_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtPlanilhaEstoque_DoubleClick ]
		private void txtPlanilhaEstoque_DoubleClick(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#endregion

		#region [ lbMensagem ]

		#region [ lbMensagem_DoubleClick ]
		private void lbMensagem_DoubleClick(object sender, EventArgs e)
		{
			if (lbMensagem.Items.Count == 0) return;
			if (lbMensagem.SelectedIndex < 0) return;
			aviso(lbMensagem.Items[lbMensagem.SelectedIndex].ToString());
		}
		#endregion

		#endregion

		#region [ lbErro ]

		#region [ lbErro_DoubleClick ]
		private void lbErro_DoubleClick(object sender, EventArgs e)
		{
			if (lbErro.Items.Count == 0) return;
			if (lbErro.SelectedIndex < 0) return;
			aviso(lbErro.Items[lbErro.SelectedIndex].ToString());
		}
		#endregion

		#endregion

		#region [ btnAbrePlanilhaEstoque ]

		#region [ btnAbrePlanilhaEstoque_Click ]
		private void btnAbrePlanilhaEstoque_Click(object sender, EventArgs e)
		{
			trataBotaoAbrePlanilhaEstoque();
		}
		#endregion

		#endregion

		#region [ btnAtualizaPlanilhaEstoque ]

		#region [ btnAtualizaPlanilhaEstoque_Click ]
		private void btnAtualizaPlanilhaEstoque_Click(object sender, EventArgs e)
		{
			trataBotaoAtualizaPlanilhaEstoque();
		}
		#endregion

		#endregion

		#endregion
	}
}
