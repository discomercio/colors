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

namespace ConsolidadorXlsEC
{
	public partial class FAtualizaPrecosSistema : FModelo
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

		public readonly string _AtualizaPrecosSistemaPlanilhaControleExcelVisible = Global.GetConfigurationValue("AtualizaPrecosSistemaPlanilhaControleExcelVisible");

		private string _tituloBoxDisplayInformativo = "Mensagens Informativas";
		private int _qtdeMsgDisplayInformativo = 0;
		private string _tituloBoxDisplayErro = "Mensagens de Erro";
		private int _qtdeMsgDisplayErro = 0;
		#endregion

		#region [ Construtor ]
		public FAtualizaPrecosSistema()
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

		#region [ pathPlanilhaControleValorDefault ]
		private String pathPlanilhaControleValorDefault()
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
			if (Global.Usuario.Defaults.FAtualizaPrecosSistema.pathArquivoPlanilhaControle.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FAtualizaPrecosSistema.pathArquivoPlanilhaControle))
				{
					strResp = Global.Usuario.Defaults.FAtualizaPrecosSistema.pathArquivoPlanilhaControle;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNamePlanilhaControleValorDefault ]
		private String fileNamePlanilhaControleValorDefault()
		{
			String strResp = "";

			if ((Global.Usuario.Defaults.FAtualizaPrecosSistema.fileNameArquivoPlanilhaControle ?? "").Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FAtualizaPrecosSistema.pathArquivoPlanilhaControle + "\\" + Global.Usuario.Defaults.FAtualizaPrecosSistema.fileNameArquivoPlanilhaControle))
				{
					strResp = Global.Usuario.Defaults.FAtualizaPrecosSistema.fileNameArquivoPlanilhaControle;
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
			#region [ Planilha de controle ]
			if (txtPlanilhaControle.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar a planilha de controle que será processada!!");
				return false;
			}
			if (!File.Exists(txtPlanilhaControle.Text))
			{
				avisoErro("A planilha de controle informada não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoAbrePlanilhaControle ]
		private void trataBotaoAbrePlanilhaControle()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "trataBotaoAbrePlanilhaControle()";
			string strNomeArqPlanilhaControle;
			#endregion

			strNomeArqPlanilhaControle = txtPlanilhaControle.Text.Trim();

			if (strNomeArqPlanilhaControle.Length == 0) return;

			if (!File.Exists(strNomeArqPlanilhaControle))
			{
				avisoErro("Arquivo '" + Path.GetFileName(strNomeArqPlanilhaControle) + "' não foi encontrado!!");
				return;
			}

			try
			{
				Process.Start(strNomeArqPlanilhaControle);
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar abrir planilha '" + Path.GetFileName(strNomeArqPlanilhaControle) + "'\r\n" + ex.ToString());
			}
		}
		#endregion

		#region [ trataBotaoSelecionaPlanilhaControle ]
		private void trataBotaoSelecionaPlanilhaControle()
		{
			#region [ Declarações ]
			DialogResult dr;
			#endregion

			try
			{
				openFileDialogCtrl.InitialDirectory = pathPlanilhaControleValorDefault();
				openFileDialogCtrl.FileName = fileNamePlanilhaControleValorDefault();
				dr = openFileDialogCtrl.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialogCtrl.FileName.Length > 0) && (txtPlanilhaControle.Text.Length > 0))
				{
					if (openFileDialogCtrl.FileName.ToUpper().Equals(txtPlanilhaControle.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos de mensagens ]
				limpaCamposMensagem();
				#endregion

				txtPlanilhaControle.Text = openFileDialogCtrl.FileName;
				Global.Usuario.Defaults.FAtualizaPrecosSistema.pathArquivoPlanilhaControle = Path.GetDirectoryName(openFileDialogCtrl.FileName);
				Global.Usuario.Defaults.FAtualizaPrecosSistema.fileNameArquivoPlanilhaControle = Path.GetFileName(openFileDialogCtrl.FileName);
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

		#region [ trataBotaoAtualizaPrecosSistema ]
		private void trataBotaoAtualizaPrecosSistema()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FAtualizaPrecosSistema.trataBotaoAtualizaPrecosSistema";
			const string NOME_WORKSHEET_CONTROLE = "Arclube";
			const int QTDE_PARCELAMENTO_PRAZO = 10;
			string MARGEM_MSG_NIVEL_2 = new string(' ', 8);
			int iNumLinha;
			int iXlDadosMinIndex;
			int iXlDadosMaxIndex;
			int iColIndex;
			int qtdeColObrigatoriaEncontrada;
			int qtdeLinhasVaziasConsecutivas;
			int qtdeLinhaDadosPlanilhaControle;
			decimal vlAVista;
			bool blnTitulosOk;
			bool blnAchou;
			bool blnAdicionaErro;
			bool blnSucesso;
			string strAux;
			string strMsg;
			string strMsgErro = "";
			string strMsgErroAux;
			string strMsgErroLog = "";
			string strNomeArqPlanilhaControle;
			string strValue;
			string strTituloEncontrado;
			string strTituloEncontradoUppercase;
			string strCodigoProdutoPrecoEmAtualizacao;
			string strFabricante;
			StringBuilder sbMsg;
			StringBuilder sbMsgErro;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			Log log = new Log();
			Process[] processosAnterior;
			Process[] processosAtual;
			object oXL = null;
			object oWBs = null;
			object oWB = null;
			object oWSs = null;
			object oWS = null;
			object oRange = null;
			object[,] oRangeValue = null;
			List<PlanilhaControleColumn> vPlanilhaControleColunasObrigatorias;
			PlanilhaControleHeader planilhaControleHeader;
			PlanilhaControleLinha linhaPlanilhaControle;
			ProdutoCadastroBasico produtoCadastroBasico;
			ProdutoLoja produtoLoja;
			ProdutoLoja produtoLojaNovo;
			PercentualCustoFinanceiroFornecedor percentualCustoFinanceiro;
			List<ProdutoCompostoItem> vProdutoCompostoItem;
			ProdutoCompostoCalculaPrecoLista prodCompCalculo;
			List<string> listaFabricanteErroTabelaCustoFinanceiro = new List<string>();
			List<string> listaProdutoNaoVendavel;
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

			#region [ Obtém nome do arquivo da planilha de controle e da planilha da ferramenta de preços ]
			strNomeArqPlanilhaControle = txtPlanilhaControle.Text;
			#endregion

			#region [ Consistências ]
			if (strNomeArqPlanilhaControle.Length == 0)
			{
				strMsgErro = "É necessário selecionar a planilha de controle a ser processada!!";
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}

			if (!File.Exists(strNomeArqPlanilhaControle))
			{
				strMsgErro = "O arquivo da planilha de controle não existe!!\r\n" + strNomeArqPlanilhaControle;
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}

			if (Global.IsFileLocked(strNomeArqPlanilhaControle))
			{
				strMsgErro = "A planilha de controle '" + Path.GetFileName(strNomeArqPlanilhaControle) + "' está aberta e em uso!!\r\nNão é possível prosseguir com o processamento!!";
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma a execução da atualização de preços no sistema?"))
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
					MARGEM_MSG_NIVEL_2 + "Planilha de controle: " + strNomeArqPlanilhaControle;
			adicionaDisplay(strMsg);
			#endregion

			try
			{
				#region [ Processa planilha de controle ]
				try // Finally
				{
					strMsg = "Inicialização do processo do Excel para processamento da planilha de controle";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Atualização dos preços no sistema");

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

					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.Visible, (_AtualizaPrecosSistemaPlanilhaControleExcelVisible.Equals("0") ? false : true));
					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.DisplayAlerts, false);
					oWBs = ExcelAutomation.GetProperty(oXL, ExcelAutomation.PropertyType.Workbooks);
					ExcelAutomation.InvokeMethod(oWBs, ExcelAutomation.MethodType.Open, strNomeArqPlanilhaControle);
					oWB = ExcelAutomation.GetProperty(oWBs, ExcelAutomation.PropertyType.Item, 1);
					oWSs = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.Worksheets, null);
					oWS = ExcelAutomation.GetProperty(oWSs, ExcelAutomation.PropertyType.Item, NOME_WORKSHEET_CONTROLE);
					ExcelAutomation.InvokeMethod(oWS, ExcelAutomation.MethodType.Select, null);

					#region [ Verifica se a planilha que está ativa realmente é a planilha de controle ]
					// Desaloca as referências anteriores, caso contrário, o processo do Excel não é encerrado ao final
					ExcelAutomation.NAR(oWS);
					oWS = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.ActiveSheet);
					strValue = (String)ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Name);
					if (!strValue.ToUpper().Equals(NOME_WORKSHEET_CONTROLE.ToUpper()))
					{
						strMsg = "Falha ao tentar tornar ativa a planilha de controle (" + NOME_WORKSHEET_CONTROLE + ")\r\nPlanilha ativa: " + strValue;
						adicionaErro(strMsg);
						avisoErro(strMsg);
						return;
					}
					#endregion

					#region [ Obtém a linha dos títulos da planilha ]
					// Analisa somente até a coluna 'Média Mercado', os campos seguintes são desconsiderados por serem campos calculados ou de uso específico da equipe Arclube
					iXlDadosMinIndex = 1;
					iXlDadosMaxIndex = 64;

					iNumLinha = 1;
					strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
					oRange = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strAux);
					oRangeValue = (object[,])ExcelAutomation.GetProperty(oRange, ExcelAutomation.PropertyType.Value);
					#endregion

					#region [ Colunas obrigatórias ]

					#region [ Obtém a posição das colunas obrigatórias ]
					vPlanilhaControleColunasObrigatorias = new List<PlanilhaControleColumn>();
					planilhaControleHeader = new PlanilhaControleHeader();

					planilhaControleHeader.Sku.ColTitleEsperado = "SKU";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.Sku);

					planilhaControleHeader.Asterisco.ColTitleEsperado = "*";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.Asterisco);

					planilhaControleHeader.ProdutoDescricao.ColTitleEsperado = "Nome";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.ProdutoDescricao);

					planilhaControleHeader.QtdeEstoque.ColTitleEsperado = "Quantidade";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.QtdeEstoque);

					planilhaControleHeader.ValorCustoMedio.ColTitleEsperado = "Custo Médio";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.ValorCustoMedio);

					planilhaControleHeader.ValorMedioMercado.ColTitleEsperado = "Média Mercado";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.ValorMedioMercado);

					planilhaControleHeader.ValorMinimoMercado.ColTitleEsperado = "Mínimo Mercado";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.ValorMinimoMercado);

					planilhaControleHeader.PrecoFinal.ColTitleEsperado = "Preço Final";
					vPlanilhaControleColunasObrigatorias.Add(planilhaControleHeader.PrecoFinal);

					qtdeColObrigatoriaEncontrada = 0;
					iColIndex = oRangeValue.GetLowerBound(1) - 1;
					while (true)
					{
						iColIndex++;
						if (iColIndex > oRangeValue.GetUpperBound(1)) break;

						if (oRangeValue[1, iColIndex] == null) continue;

						strTituloEncontrado = oRangeValue[1, iColIndex].ToString().Trim();
						strTituloEncontradoUppercase = strTituloEncontrado.ToUpper();

						foreach (var colObrigatoria in vPlanilhaControleColunasObrigatorias)
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
						if (qtdeColObrigatoriaEncontrada == vPlanilhaControleColunasObrigatorias.Count) break;
					} // while (true)
					#endregion

					#region [ Verifica se alguma coluna obrigatória não foi localizada ]
					blnTitulosOk = true;
					sbMsgErro = new StringBuilder("");
					foreach (var colObrigatoria in vPlanilhaControleColunasObrigatorias)
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
						strMsgErro = "A planilha '" + Path.GetFileName(strNomeArqPlanilhaControle) + "' não possui os títulos corretos para as colunas!!\r\nVerifique se a planilha correta foi selecionada!!\r\n\r\n" + sbMsgErro.ToString();
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#endregion

					#region [ Verifica existência dos dados de custo financeiro do fornecedor ]
					foreach (var ambiente in FMain.contextoBD.Ambientes)
					{
						if (ambiente.tabelaPercCustoFinanceiroFornecedor == null)
						{
							ambiente.tabelaPercCustoFinanceiroFornecedor = ambiente.produtoDAO.GetTabelaPercentualCustoFinanceiroFornecedor(out strMsgErroAux);
						}

						if (ambiente.tabelaPercCustoFinanceiroFornecedor == null)
						{
							strMsg = "Não há dados de percentual de custo financeiro do fornecedor no ambiente '" + ambiente.NomeAmbiente + "', portanto, todos os produtos da planilha serão atualizados no sistema com o preço a prazo!";
							adicionaErro(strMsg);
						}
					}
					#endregion

					#region [ Memoriza a tabela de preço original da loja ]
					// Os valores originais são memorizados para permitir o cálculo da proporção correto para os produtos componentes de um produto composto.
					// Sem a memorização dos valores originais, ocorre um problema quando um produto componente é usado por mais de um produto composto: após atualizar
					// o primeiro produto composto que cause alteração no valor do produto componente, ao calcular a proporção dos itens para o segundo produto composto que
					// contenha o mesmo produto componente, as proporções estarão incorretas.
					foreach (var ambiente in FMain.contextoBD.Ambientes)
					{
						ambiente.tabelaProdutoLojaOriginal = ambiente.produtoDAO.GetTabelaPrecoLoja(ambiente.NumeroLojaArclube, out strMsgErroAux);
					}
					#endregion

					#region [ Atualiza preços no sistema ]
					strMsg = "Iniciando atualização de preços no sistema";
					adicionaDisplay(strMsg);

					qtdeLinhasVaziasConsecutivas = 0;
					qtdeLinhaDadosPlanilhaControle = 0;
					iXlDadosMaxIndex = 0;
					foreach (var item in vPlanilhaControleColunasObrigatorias)
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
						if (oRangeValue[1, planilhaControleHeader.Sku.ColIndex] == null)
						{
							qtdeLinhasVaziasConsecutivas++;
							continue;
						}

						if (oRangeValue[1, planilhaControleHeader.Sku.ColIndex].ToString().Trim().Length == 0)
						{
							qtdeLinhasVaziasConsecutivas++;
							continue;
						}
						#endregion

						qtdeLinhaDadosPlanilhaControle++;
						// Encontrou linha com dados, portanto, zera contador de linhas vazias consecutivas
						qtdeLinhasVaziasConsecutivas = 0;
						strMsg = "Processamento da " + qtdeLinhaDadosPlanilhaControle.ToString() + "ª linha de dados da planilha de controle: SKU " + oRangeValue[1, planilhaControleHeader.Sku.ColIndex].ToString();
						adicionaDisplay(strMsg);
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);

						#region [ Carrega dados para objeto da classe LinhaPlanilhaControle ]
						linhaPlanilhaControle = new PlanilhaControleLinha();

						#region [ SKU ]
						linhaPlanilhaControle.Sku = oRangeValue[1, planilhaControleHeader.Sku.ColIndex].ToString().Trim();
						linhaPlanilhaControle.SkuFormatado = Global.normalizaCodigoProduto(linhaPlanilhaControle.Sku);
						#endregion

						#region [ Asterisco ]
						if (oRangeValue[1, planilhaControleHeader.Asterisco.ColIndex] != null)
						{
							linhaPlanilhaControle.Asterisco = oRangeValue[1, planilhaControleHeader.Asterisco.ColIndex].ToString().Trim();
						}
						#endregion

						#region [ Descrição do produto ]
						if (oRangeValue[1, planilhaControleHeader.ProdutoDescricao.ColIndex] != null)
						{
							linhaPlanilhaControle.ProdutoDescricao = oRangeValue[1, planilhaControleHeader.ProdutoDescricao.ColIndex].ToString().Trim();
						}
						#endregion

						#region [ Quantidade no estoque ]
						if (oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaControle.QtdeEstoque = (double)oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex];
								}
								else if (oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									linhaPlanilhaControle.QtdeEstoque = (int)oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaControle.QtdeEstoque = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = "Falha ao ler o conteúdo da célula '" + planilhaControleHeader.QtdeEstoque.ColTitle + "' do SKU " + linhaPlanilhaControle.Sku + " (Type = " + oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaControleHeader.QtdeEstoque.ColIndex].ToString() + ")";
								adicionaErro(strMsg);
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
							}
						}
						#endregion

						#region [ Valor do custo médio ]
						if (oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaControle.ValorCustoMedio = (double)oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex];
								}
								else if (oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaControle.ValorCustoMedio = (double)(decimal)oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaControle.ValorCustoMedio = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								blnAdicionaErro = true;
								if (oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel exibe #DIV/0!, #NÚM! ou #REF!, situações em que o valor da célula é do tipo System.Int32 e valores como -2146826281, -2146826252, etc
									if ((int)oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex] < 0) blnAdicionaErro = false;
								}
								if (oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel contém anotações do usuário
									if (Global.contagemLetras(oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].ToString()) > 3) blnAdicionaErro = false;
								}
								if (blnAdicionaErro)
								{
									strMsg = "Falha ao ler o conteúdo da célula '" + planilhaControleHeader.ValorCustoMedio.ColTitle + "' do SKU " + linhaPlanilhaControle.Sku + " (Type = " + oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaControleHeader.ValorCustoMedio.ColIndex].ToString() + ")";
									adicionaErro(strMsg);
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
								}
							}
						}
						#endregion

						#region [ Valor do preço médio do mercado ]
						if (oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaControle.ValorMedioMercado = (double)oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex];
								}
								else if (oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaControle.ValorMedioMercado = (double)(decimal)oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaControle.ValorMedioMercado = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								blnAdicionaErro = true;
								if (oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel exibe #DIV/0!, #NÚM! ou #REF!, situações em que o valor da célula é do tipo System.Int32 e valores como -2146826281, -2146826252, etc
									if ((int)oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex] < 0) blnAdicionaErro = false;
								}
								if (oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel contém anotações do usuário
									if (Global.contagemLetras(oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].ToString()) > 3) blnAdicionaErro = false;
								}
								if (blnAdicionaErro)
								{
									strMsg = "Falha ao ler o conteúdo da célula '" + planilhaControleHeader.ValorMedioMercado.ColTitle + "' do SKU " + linhaPlanilhaControle.Sku + " (Type = " + oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaControleHeader.ValorMedioMercado.ColIndex].ToString() + ")";
									adicionaErro(strMsg);
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
								}
							}
						}
						#endregion

						#region [ Valor mínimo do mercado ]
						if (oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaControle.ValorMinimoMercado = (double)oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex];
								}
								else if (oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaControle.ValorMinimoMercado = (double)(decimal)oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaControle.ValorMinimoMercado = (double)Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								blnAdicionaErro = true;
								if (oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel exibe #DIV/0!, #NÚM! ou #REF!, situações em que o valor da célula é do tipo System.Int32 e valores como -2146826281, -2146826252, etc
									if ((int)oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex] < 0) blnAdicionaErro = false;
								}
								if (oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel contém anotações do usuário
									if (Global.contagemLetras(oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].ToString()) > 3) blnAdicionaErro = false;
								}
								if (blnAdicionaErro)
								{
									strMsg = "Falha ao ler o conteúdo da célula '" + planilhaControleHeader.ValorMinimoMercado.ColTitle + "' do SKU " + linhaPlanilhaControle.Sku + " (Type = " + oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaControleHeader.ValorMinimoMercado.ColIndex].ToString() + ")";
									adicionaErro(strMsg);
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
								}
							}
						}
						#endregion

						#region [ Preço final ]
						if (oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].GetType().FullName.ToUpper().Equals("System.Double".ToUpper()))
								{
									linhaPlanilhaControle.PrecoFinal = (decimal)(double)oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex];
								}
								else if (oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaControle.PrecoFinal = (decimal)oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex];
								}
								else
								{
									linhaPlanilhaControle.PrecoFinal = Global.converteNumeroDecimal(oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].ToString());
								}
							}
							catch (Exception ex)
							{
								blnAdicionaErro = true;
								if (oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel exibe #DIV/0!, #NÚM! ou #REF!, situações em que o valor da célula é do tipo System.Int32 e valores como -2146826281, -2146826252, etc
									if ((int)oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex] < 0) blnAdicionaErro = false;
								}
								if (oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									// Não registra o erro quando a célula do Excel contém anotações do usuário
									if (Global.contagemLetras(oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].ToString()) > 3) blnAdicionaErro = false;
								}
								if (blnAdicionaErro)
								{
									strMsg = "Falha ao ler o conteúdo da célula '" + planilhaControleHeader.PrecoFinal.ColTitle + "' do SKU " + linhaPlanilhaControle.Sku + " (Type = " + oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaControleHeader.PrecoFinal.ColIndex].ToString() + ")";
									adicionaErro(strMsg);
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
								}
							}
						}
						#endregion

						#endregion

						if (linhaPlanilhaControle.PrecoFinal == null)
						{
							strMsg = "Produto " + linhaPlanilhaControle.SkuFormatado + " será ignorado na atualização de preços porque não foi possível ler o valor do preço final da planilha!";
							adicionaErro(strMsg);
						}
						else if (linhaPlanilhaControle.PrecoFinal == 0m)
						{
							strMsg = "Produto " + linhaPlanilhaControle.SkuFormatado + " será ignorado na atualização de preços porque o valor do preço final na planilha está zerado!";
							adicionaErro(strMsg);
						}
						else
						{
							#region [ Atualiza o preço em todos os ambientes ]
							foreach (var ambiente in FMain.contextoBD.Ambientes)
							{
								produtoCadastroBasico = ambiente.produtoDAO.GetProdutoCadastroBasico(linhaPlanilhaControle.SkuFormatado, out strMsgErroAux);
								if (produtoCadastroBasico != null)
								{
									if (!produtoCadastroBasico.isCadastrado)
									{
										strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + linhaPlanilhaControle.SkuFormatado + " não encontrado no cadastro de produtos!";
										adicionaDisplay(strMsg);
									}
									else
									{
										if (!produtoCadastroBasico.isComposto)
										{
											#region [ Produto normal ]
											try
											{
												produtoLoja = ambiente.tabelaProdutoLojaOriginal.Single(p => p.loja.Equals(ambiente.NumeroLojaArclube) && p.fabricante.Equals(produtoCadastroBasico.fabricante) && p.produto.Equals(produtoCadastroBasico.produto));
											}
											catch (Exception)
											{
												produtoLoja = null;
												strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + produtoCadastroBasico.produto + " não foi localizado na tabela de preços da loja " + ambiente.NumeroLojaArclube;
												adicionaDisplay(strMsg);
											}

											if (produtoLoja == null)
											{
												strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + produtoCadastroBasico.produto + " não foi encontrado na tabela de preços da loja " + ambiente.NumeroLojaArclube;
												adicionaDisplay(strMsg);
											}
											else
											{
												// Obtém o percentual do custo financeiro do fabricante
												if (ambiente.tabelaPercCustoFinanceiroFornecedor == null)
												{
													vlAVista = (decimal)linhaPlanilhaControle.PrecoFinal;
												}
												else
												{
													// Calcula o preço à vista, já que o preço na planilha é referente ao preço a prazo em 10x
													try
													{
														percentualCustoFinanceiro = ambiente.tabelaPercCustoFinanceiroFornecedor.Single(
																		p => p.fabricante.Equals(produtoCadastroBasico.fabricante)
																			&& p.tipo_parcelamento.Equals(Global.Cte.PercentualCustoFinanceiroFornecedor.TipoParcelamento.SEM_ENTRADA)
																			&& (p.qtde_parcelas == QTDE_PARCELAMENTO_PRAZO));
													}
													catch (Exception)
													{
														percentualCustoFinanceiro = null;
														Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": falha ao tentar localizar o percentual de custo financeiro do fabricante " + produtoCadastroBasico.fabricante);
													}

													if (percentualCustoFinanceiro == null)
													{
														vlAVista = (decimal)linhaPlanilhaControle.PrecoFinal;
														strFabricante = listaFabricanteErroTabelaCustoFinanceiro.SingleOrDefault(
																		p => p.Equals(produtoCadastroBasico.fabricante));
														if ((strFabricante ?? "").Length == 0)
														{
															listaFabricanteErroTabelaCustoFinanceiro.Add(produtoCadastroBasico.fabricante);
															strMsg = "Não foi localizado o coeficiente de custo financeiro do fornecedor " + produtoCadastroBasico.fabricante + " no ambiente '" + ambiente.NomeAmbiente + "' para o parcelamento em " + QTDE_PARCELAMENTO_PRAZO.ToString() + " vezes, portanto, o preço a prazo será usado como sendo o preço à vista!";
															adicionaErro(strMsg);
														}
													}
													else
													{
														if (percentualCustoFinanceiro.coeficiente > 0d)
														{
															vlAVista = (decimal)((double)linhaPlanilhaControle.PrecoFinal / percentualCustoFinanceiro.coeficiente);
														}
														else
														{
															vlAVista = (decimal)linhaPlanilhaControle.PrecoFinal;
															strFabricante = listaFabricanteErroTabelaCustoFinanceiro.SingleOrDefault(
																		p => p.Equals(produtoCadastroBasico.fabricante));
															if ((strFabricante ?? "").Length == 0)
															{
																listaFabricanteErroTabelaCustoFinanceiro.Add(produtoCadastroBasico.fabricante);
																strMsg = "Não foi encontrado o coeficiente de custo financeiro do fornecedor " + produtoCadastroBasico.fabricante + " no ambiente '" + ambiente.NomeAmbiente + "' para o parcelamento em " + QTDE_PARCELAMENTO_PRAZO.ToString() + " vezes, portanto, o preço a prazo será usado como sendo o preço à vista!";
																adicionaErro(strMsg);
															}
														}
													}
												}

												vlAVista = Global.arredondaParaMonetario(vlAVista);

												if (produtoLoja.preco_lista == vlAVista)
												{
													strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + produtoLoja.produto + ": preço de lista não teve alteração (" + Global.formataMoeda(vlAVista) + ")";
													adicionaDisplay(strMsg);
												}
												else
												{
													#region [ Atualiza o preço no BD ]
													produtoLojaNovo = new ProdutoLoja();
													produtoLojaNovo.fabricante = produtoLoja.fabricante;
													produtoLojaNovo.produto = produtoLoja.produto;
													produtoLojaNovo.loja = produtoLoja.loja;
													produtoLojaNovo.preco_lista = vlAVista;
													if (ambiente.produtoDAO.UpdateProdutoLoja(produtoLojaNovo, out strMsgErroAux))
													{
														strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + produtoLojaNovo.produto + ": preço de lista atualizado de " + Global.formataMoeda(produtoLoja.preco_lista) + " para " + Global.formataMoeda(produtoLojaNovo.preco_lista);
														adicionaDisplay(strMsg);
													}
													else
													{
														strMsg = "[" + ambiente.NomeAmbiente + "] Produto " + produtoLojaNovo.produto + ": falha ao tentar atualizar preço de lista de " + Global.formataMoeda(produtoLoja.preco_lista) + " para " + Global.formataMoeda(produtoLojaNovo.preco_lista) + (strMsgErroAux.Length > 0 ? "\r\n" + MARGEM_MSG_NIVEL_2 + strMsgErroAux : "");
														adicionaErro(strMsg);
													}
													#endregion
												}
											}
											#endregion
										}
										else
										{
											#region [ Produto composto ]
											vProdutoCompostoItem = ambiente.produtoDAO.GetProdutoCompostoItem(produtoCadastroBasico.fabricante, produtoCadastroBasico.produto, out strMsgErroAux);
											if (vProdutoCompostoItem == null)
											{
												strMsg = "[" + ambiente.NomeAmbiente + "] Falha ao tentar recuperar os dados da composição do produto composto " + produtoCadastroBasico.produto +
														((strMsgErroAux.Length > 0) ? "\r\n" + MARGEM_MSG_NIVEL_2 + strMsgErroAux : "");
												adicionaErro(strMsg);
											}
											else
											{
												if (vProdutoCompostoItem.Count == 0)
												{
													strMsg = "[" + ambiente.NomeAmbiente + "] Não há dados no retorno da consulta para recuperar os detalhes da composição do produto composto " + produtoCadastroBasico.produto;
													adicionaErro(strMsg);
												}
												else
												{
													#region [ Calcula o valor do preço de lista a ser atribuído a cada item da composição a partir da proporção dos valores cadastrados atualmente no BD ]
													prodCompCalculo = new ProdutoCompostoCalculaPrecoLista();
													prodCompCalculo.preco_lista_a_prazo_total_novo = (decimal)linhaPlanilhaControle.PrecoFinal;

													foreach (ProdutoCompostoItem item in vProdutoCompostoItem)
													{
														prodCompCalculo.Itens.Add(new ProdutoCompostoItemCalculoPrecoLista(item.fabricante_composto, item.produto_composto, item.fabricante_item, item.produto_item, item.qtde));
													}

													listaProdutoNaoVendavel = new List<string>();
													foreach (ProdutoCompostoItemCalculoPrecoLista item in prodCompCalculo.Itens)
													{
														try
														{
															produtoLoja = ambiente.tabelaProdutoLojaOriginal.Single(p => p.loja.Equals(ambiente.NumeroLojaArclube) && p.fabricante.Equals(item.fabricante_item) && p.produto.Equals(item.produto_item));
														}
														catch (Exception)
														{
															produtoLoja = null;
															strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + item.produto_item + " (SKU " + produtoCadastroBasico.produto + ") não foi localizado na tabela de preços da loja " + ambiente.NumeroLojaArclube;
															adicionaDisplay(strMsg);
															prodCompCalculo.ocorreuErro = true;
															prodCompCalculo.mensagem_erro = strMsg;
															break;
														}

														if (produtoLoja == null)
														{
															strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + item.produto_item + " (SKU " + produtoCadastroBasico.produto + ") não foi encontrado na tabela de preços da loja " + ambiente.NumeroLojaArclube;
															adicionaDisplay(strMsg);
															prodCompCalculo.ocorreuErro = true;
															prodCompCalculo.mensagem_erro = strMsg;
															break;
														}
														else if (!produtoLoja.vendavel.ToUpper().Equals("S"))
														{
															listaProdutoNaoVendavel.Add(produtoLoja.produto);
														}
														else
														{
															item.preco_lista_a_vista_atual = produtoLoja.preco_lista;
															prodCompCalculo.preco_lista_a_vista_total_atual += item.qtde * item.preco_lista_a_vista_atual;
														}
													}

													if (!prodCompCalculo.ocorreuErro)
													{
														if ((prodCompCalculo.preco_lista_a_vista_total_atual == 0m) && (listaProdutoNaoVendavel.Count > 0))
														{
															foreach (string codigoProduto in listaProdutoNaoVendavel)
															{
																strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + codigoProduto + " (SKU " + produtoCadastroBasico.produto + ") está com preço zero e consta como 'não-vendável' na tabela de preços da loja " + ambiente.NumeroLojaArclube;
																adicionaDisplay(strMsg);
															}
															continue;
														}
													}

													if (!prodCompCalculo.ocorreuErro)
													{
														if (prodCompCalculo.preco_lista_a_vista_total_atual == 0m)
														{
															strMsg = "[" + ambiente.NomeAmbiente + "] O produto composto " + produtoCadastroBasico.produto + " resulta em um preço de lista zero na tabela da loja " + ambiente.NumeroLojaArclube + ", portanto, não é possível calcular o novo preço de lista dos itens da composição!!";
															prodCompCalculo.ocorreuErro = true;
															prodCompCalculo.mensagem_erro = strMsg;
														}
													}

													// Em caso de erro, interrompe o processamento p/ este ambiente e segue p/ o próximo
													if (prodCompCalculo.ocorreuErro)
													{
														strMsg = "[" + ambiente.NomeAmbiente + "] Atualização de preço dos itens da composição do SKU " + produtoCadastroBasico.produto + " foi cancelada devido ao erro:\r\n" + MARGEM_MSG_NIVEL_2 + prodCompCalculo.mensagem_erro;
														adicionaErro(strMsg);
														continue;
													}

													#region [ Calcula a razão do preço de lista e calcula o preço de lista à vista ]
													foreach (ProdutoCompostoItemCalculoPrecoLista item in prodCompCalculo.Itens)
													{
														item.razaoPrecoListaTotalPorUnidade = (double)(item.preco_lista_a_vista_atual / prodCompCalculo.preco_lista_a_vista_total_atual);
														item.preco_lista_a_prazo_novo = (decimal)((double)prodCompCalculo.preco_lista_a_prazo_total_novo * item.razaoPrecoListaTotalPorUnidade);

														if (ambiente.tabelaPercCustoFinanceiroFornecedor == null)
														{
															vlAVista = item.preco_lista_a_prazo_novo;
														}
														else
														{
															try
															{
																// A planilha contém o preço a prazo e o banco de dados armazena o valor do preço à vista
																percentualCustoFinanceiro = ambiente.tabelaPercCustoFinanceiroFornecedor.Single(
																				p => p.fabricante.Equals(item.fabricante_item)
																				&& p.tipo_parcelamento.Equals(Global.Cte.PercentualCustoFinanceiroFornecedor.TipoParcelamento.SEM_ENTRADA)
																				&& (p.qtde_parcelas == QTDE_PARCELAMENTO_PRAZO));
															}
															catch (Exception)
															{
																percentualCustoFinanceiro = null;
																Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": falha ao tentar localizar o percentual de custo financeiro do fabricante " + item.fabricante_item);
															}

															if (percentualCustoFinanceiro == null)
															{
																vlAVista = item.preco_lista_a_prazo_novo;
																strFabricante = listaFabricanteErroTabelaCustoFinanceiro.SingleOrDefault(
																			p => p.Equals(item.fabricante_item));
																if ((strFabricante ?? "").Length == 0)
																{
																	listaFabricanteErroTabelaCustoFinanceiro.Add(item.fabricante_item);
																	strMsg = "Não foi localizado o coeficiente de custo financeiro do fornecedor " + item.fabricante_item + " no ambiente '" + ambiente.NomeAmbiente + "' para o parcelamento em " + QTDE_PARCELAMENTO_PRAZO.ToString() + " vezes, portanto, o preço a prazo será usado como sendo o preço à vista!";
																	adicionaErro(strMsg);
																}
															}
															else
															{
																if (percentualCustoFinanceiro.coeficiente > 0d)
																{
																	vlAVista = (decimal)((double)item.preco_lista_a_prazo_novo / percentualCustoFinanceiro.coeficiente);
																}
																else
																{
																	vlAVista = item.preco_lista_a_prazo_novo;
																	strFabricante = listaFabricanteErroTabelaCustoFinanceiro.SingleOrDefault(
																				p => p.Equals(item.fabricante_item));
																	if ((strFabricante ?? "").Length == 0)
																	{
																		listaFabricanteErroTabelaCustoFinanceiro.Add(item.fabricante_item);
																		strMsg = "Não foi encontrado o coeficiente de custo financeiro do fornecedor " + item.fabricante_item + " no ambiente '" + ambiente.NomeAmbiente + "' para o parcelamento em " + QTDE_PARCELAMENTO_PRAZO.ToString() + " vezes, portanto, o preço a prazo será usado como sendo o preço à vista!";
																		adicionaErro(strMsg);
																	}
																}
															}
														}

														vlAVista = Global.arredondaParaMonetario(vlAVista);
														item.preco_lista_a_vista_novo = vlAVista;
													} // foreach (ProdutoCompostoItemCalculoPrecoLista item in prodCompCalculo.Itens)
													#endregion

													#endregion

													#region [ Atualiza o preço de lista no BD para cada item da composição ]
													blnSucesso = false;
													ambiente.BD.iniciaTransacao();
													strCodigoProdutoPrecoEmAtualizacao = "";
													try
													{
														foreach (ProdutoCompostoItemCalculoPrecoLista item in prodCompCalculo.Itens)
														{
															strCodigoProdutoPrecoEmAtualizacao = item.produto_item;
															if (item.preco_lista_a_vista_atual == item.preco_lista_a_vista_novo)
															{
																blnSucesso = true;
																strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + item.produto_item + " (SKU " + produtoCadastroBasico.produto + "): preço de lista não teve alteração (" + Global.formataMoeda(item.preco_lista_a_vista_novo) + ")";
																adicionaDisplay(strMsg);
															}
															else
															{
																produtoLojaNovo = new ProdutoLoja();
																produtoLojaNovo.fabricante = item.fabricante_item;
																produtoLojaNovo.produto = item.produto_item;
																produtoLojaNovo.loja = ambiente.NumeroLojaArclube;
																produtoLojaNovo.preco_lista = item.preco_lista_a_vista_novo;
																blnSucesso = ambiente.produtoDAO.UpdateProdutoLoja(produtoLojaNovo, out strMsgErroAux);
																if (blnSucesso)
																{
																	strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto " + produtoLojaNovo.produto + " (SKU " + produtoCadastroBasico.produto + "): preço de lista atualizado de " + Global.formataMoeda(item.preco_lista_a_vista_atual) + " para " + Global.formataMoeda(produtoLojaNovo.preco_lista);
																	adicionaDisplay(strMsg);
																}
																else
																{
																	strMsg = "[" + ambiente.NomeAmbiente + "] Produto " + produtoLojaNovo.produto + " (SKU " + produtoCadastroBasico.produto + "): falha ao tentar atualizar preço de lista de " + Global.formataMoeda(item.preco_lista_a_vista_atual) + " para " + Global.formataMoeda(produtoLojaNovo.preco_lista) + (strMsgErroAux.Length > 0 ? "\r\n" + MARGEM_MSG_NIVEL_2 + strMsgErroAux : "");
																	adicionaErro(strMsg);
																	// Se houve erro, interrompe a atualização dos preços dos itens da composição e o rollback irá retornar os valores originais
																	break;
																}
															}
														}
													}
													catch (Exception ex)
													{
														blnSucesso = false;
														strMsg = "Falha ao tentar atualizar o preço do produto " + strCodigoProdutoPrecoEmAtualizacao + " (SKU " + produtoCadastroBasico.produto + "): " + ex.Message;
														adicionaErro(strMsg);
														Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar atualizar o preço do produto " + strCodigoProdutoPrecoEmAtualizacao + " (produto composto: " + produtoCadastroBasico.produto + ")\r\n" + ex.ToString());
													}
													finally
													{
														if (blnSucesso)
														{
															ambiente.BD.commitTransacao();
														}
														else
														{
															ambiente.BD.rollbackTransacao();
														}
													}

													if (blnSucesso)
													{
														strMsg = MARGEM_MSG_NIVEL_2 + "[" + ambiente.NomeAmbiente + "] Produto composto " + produtoCadastroBasico.produto + ": preço de lista de todos os itens da composição foram alterados com sucesso!";
														adicionaDisplay(strMsg);
													}
													else
													{
														strMsg = "[" + ambiente.NomeAmbiente + "] Produto composto " + produtoCadastroBasico.produto + ": ocorreu erro na atualização do preço de lista dos itens da composição e os valores originais foram mantidos!";
														adicionaErro(strMsg);
													}
													#endregion
												}
											}
											#endregion
										}
									} // if (!produtoCadastroBasico.isCadastrado)
								} // if (produtoCadastroBasico != null)
							} // foreach (var ambiente in FMain.contextoBD.Ambientes)
							#endregion
						}
					} // while (true)

					strMsg = qtdeLinhaDadosPlanilhaControle.ToString() + " linhas de dados processadas na planilha de controle";
					adicionaDisplay(strMsg);
					#endregion
				}
				finally
				{
					strMsg = "Fechamento da planilha de controle e encerramento do processo do Excel";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Fechamento da planilha de controle e encerramento do processo do Excel");
					try
					{
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
				#endregion

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
				strMsg = "Sucesso na atualização de preços no sistema (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + "): planilha de controle = " + strNomeArqPlanilhaControle + " (" + qtdeLinhaDadosPlanilhaControle.ToString() + " linhas de dados)";
				Global.gravaLogAtividade(strMsg);
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.CXLSEC.LogOperacao.ATUALIZA_PRECOS_SISTEMA;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				#region [ Mensagem de sucesso ]
				info(ModoExibicaoMensagemRodape.Normal);
				strMsg = "Processamento concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!!";
				adicionaDisplay(strMsg);
				aviso(strMsg);
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

		#region [ FAtualizaPrecosSistema ]

		#region [ FAtualizaPrecosSistema_Load ]
		private void FAtualizaPrecosSistema_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				txtPlanilhaControle.Text = "";
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

		#region [ FAtualizaPrecosSistema_Shown ]
		private void FAtualizaPrecosSistema_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			string strFileNameArquivoPlanilhaControle;
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					strFileNameArquivoPlanilhaControle = pathPlanilhaControleValorDefault() + "\\" + fileNamePlanilhaControleValorDefault();
					if (File.Exists(strFileNameArquivoPlanilhaControle)) txtPlanilhaControle.Text = strFileNameArquivoPlanilhaControle;

					openFileDialogCtrl.InitialDirectory = pathPlanilhaControleValorDefault();
					openFileDialogCtrl.FileName = fileNamePlanilhaControleValorDefault();

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

		#region [ FAtualizaPrecosSistema_FormClosing ]
		private void FAtualizaPrecosSistema_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaPlanilhaControle ]

		#region [ btnSelecionaPlanilhaControle_Click ]
		private void btnSelecionaPlanilhaControle_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaPlanilhaControle();
		}
		#endregion

		#endregion

		#region [ btnAtualizaPrecos ]

		#region [ btnAtualizaPrecos_Click ]
		private void btnAtualizaPrecos_Click(object sender, EventArgs e)
		{
			trataBotaoAtualizaPrecosSistema();
		}
		#endregion

		#endregion

		#region [ txtPlanilhaControle ]

		#region [ txtPlanilhaControle_Enter ]
		private void txtPlanilhaControle_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtPlanilhaControle_DoubleClick ]
		private void txtPlanilhaControle_DoubleClick(object sender, EventArgs e)
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

		#region [ btnAbrePlanilhaControle ]

		#region [ btnAbrePlanilhaControle_Click ]
		private void btnAbrePlanilhaControle_Click(object sender, EventArgs e)
		{
			trataBotaoAbrePlanilhaControle();
		}
		#endregion

		#endregion

		#endregion
	}
}
