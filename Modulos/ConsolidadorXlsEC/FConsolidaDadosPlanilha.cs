#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
#endregion

namespace ConsolidadorXlsEC
{
	public partial class FConsolidaDadosPlanilha : FModelo
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

		public readonly string _ConsolidaDadosPlanilhaControleExcelVisible = Global.GetConfigurationValue("ConsolidaDadosPlanilhaControleExcelVisible");
		public readonly string _ConsolidaDadosPlanilhaPrecosExcelVisible = Global.GetConfigurationValue("ConsolidaDadosPlanilhaPrecosExcelVisible");

		private string _tituloBoxDisplayInformativo = "Mensagens Informativas";
		private int _qtdeMsgDisplayInformativo = 0;
		private string _tituloBoxDisplayErro = "Mensagens de Erro";
		private int _qtdeMsgDisplayErro = 0;
		#endregion

		#region [ Construtor ]
		public FConsolidaDadosPlanilha()
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
			if (Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaControle.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaControle))
				{
					strResp = Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaControle;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNamePlanilhaControleValorDefault ]
		private String fileNamePlanilhaControleValorDefault()
		{
			String strResp = "";

			if ((Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaControle ?? "").Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaControle + "\\" + Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaControle))
				{
					strResp = Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaControle;
				}
			}
			return strResp;
		}
		#endregion

		#region [ pathPlanilhaFerramentaPrecosValorDefault ]
		private String pathPlanilhaFerramentaPrecosValorDefault()
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
			if (Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaFerramentaPrecos.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaFerramentaPrecos))
				{
					strResp = Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaFerramentaPrecos;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNamePlanilhaFerramentaPrecosValorDefault ]
		private String fileNamePlanilhaFerramentaPrecosValorDefault()
		{
			String strResp = "";

			if (Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaFerramentaPrecos.Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaFerramentaPrecos + "\\" + Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaFerramentaPrecos))
				{
					strResp = Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaFerramentaPrecos;
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

			#region [ Planilha da ferramenta de preços ]
			if (txtPlanilhaFerramentaPrecos.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar a planilha da ferramenta de preços que será processada!!");
				return false;
			}
			if (!File.Exists(txtPlanilhaFerramentaPrecos.Text))
			{
				avisoErro("A planilha da ferramenta de preços informada não existe!!");
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

		#region [ trataBotaoAbrePlanilhaPrecos ]
		private void trataBotaoAbrePlanilhaPrecos()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "trataBotaoAbrePlanilhaPrecos()";
			string strNomeArqPlanilhaFerramentaPrecos;
			#endregion

			strNomeArqPlanilhaFerramentaPrecos = txtPlanilhaFerramentaPrecos.Text.Trim();

			if (strNomeArqPlanilhaFerramentaPrecos.Length == 0) return;

			if (!File.Exists(strNomeArqPlanilhaFerramentaPrecos))
			{
				avisoErro("Arquivo '" + Path.GetFileName(strNomeArqPlanilhaFerramentaPrecos) + "' não foi encontrado!!");
				return;
			}

			try
			{
				Process.Start(strNomeArqPlanilhaFerramentaPrecos);
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar abrir planilha '" + Path.GetFileName(strNomeArqPlanilhaFerramentaPrecos) + "'\r\n" + ex.ToString());
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
				Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaControle = Path.GetDirectoryName(openFileDialogCtrl.FileName);
				Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaControle = Path.GetFileName(openFileDialogCtrl.FileName);
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

		#region [ trataBotaoSelecionaPlanilhaPrecos ]
		private void trataBotaoSelecionaPlanilhaPrecos()
		{
			#region [ Declarações ]
			DialogResult dr;
			#endregion

			try
			{
				openFileDialogPrecos.InitialDirectory = pathPlanilhaFerramentaPrecosValorDefault();
				openFileDialogPrecos.FileName = fileNamePlanilhaFerramentaPrecosValorDefault();
				dr = openFileDialogPrecos.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialogPrecos.FileName.Length > 0) && (txtPlanilhaFerramentaPrecos.Text.Length > 0))
				{
					if (openFileDialogPrecos.FileName.ToUpper().Equals(txtPlanilhaFerramentaPrecos.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos de mensagens ]
				limpaCamposMensagem();
				#endregion

				txtPlanilhaFerramentaPrecos.Text = openFileDialogPrecos.FileName;
				Global.Usuario.Defaults.FConsolidaDadosPlanilha.pathArquivoPlanilhaFerramentaPrecos = Path.GetDirectoryName(openFileDialogPrecos.FileName);
				Global.Usuario.Defaults.FConsolidaDadosPlanilha.fileNameArquivoPlanilhaFerramentaPrecos = Path.GetFileName(openFileDialogPrecos.FileName);
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

		#region [ trataBotaoConsolidaPlanilha ]
		private void trataBotaoConsolidaPlanilha()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FConsolidaDadosPlanilha.trataBotaoConsolidaPlanilha";
			const string NOME_WORKSHEET_CONTROLE = "Arclube";
			const string NOME_WORKSHEET_PRECOS = "undefined";
			const string NOME_WORKSHEET_PRECOS_V2 = "Produtos";
			const double MAX_VALOR_MARGEM_ERRO_PRECO = 0.01d;
			const double COR_FUNDO_REALCE_ALTERACAO = 65535d; // 9420794d;
			const double COR_FUNDO_INFO_NOT_FOUND = 12566463d;
			const double COR_FUNDO_LARANJA = 9420794d;
			const double COR_FUNDO_VERDE = 10213316d;
			const double COR_FUNDO_AZUL = 14136213d;
			const double COR_FUNDO_ROXO = 13082801d;
			int iNumLinha;
			int iXlDadosMinIndex;
			int iXlDadosMaxIndex;
			int iColIndex;
			int qtdeLinhasVaziasConsecutivas;
			int qtdeLinhaDadosPlanilhaPrecos;
			int qtdeLinhaDadosPlanilhaControle;
			int qtdeIsVendavelTrue;
			int qtdeProdutoCompostoItem;
			int qtdeConcorrentesPrecoMedio;
			int qtdeColObrigatoriaEncontrada;
			int qtdeEstoqueAcumulado;
			decimal vlNovoCustoIntermediario;
			decimal? vlPrecoMinimoConcorrente;
			decimal? vlPrecoMedioConcorrente;
			decimal vlAcumuladoTotalPrecoMedioConcorrente;
			decimal? vlAux;
			bool blnTitulosOk;
			bool blnAchou;
			bool blnAdicionaErro;
			string strMsgErro = "";
			string strMsgErroAux;
			string strMsgErroLog = "";
			string strNomeArqPlanilhaControle;
			string strNomeArqPlanilhaFerramentaPrecos;
			string strPathPlanilhaControleBackup;
			string strNomeArqPlanilhaControleBackup;
			string strMsg;
			string strValue;
			string strTituloEncontrado;
			string strTituloEncontradoUppercase;
			string strAux;
			string strCelQtdeEstoqueVenda;
			string strCelVlCustoIntermediario;
			string strCelPrecoMedioMercado;
			string strCelPrecoMinimoMercado;
			string strNomeConcorrentePrecoMinimo;
			string strCellCodigo;
			string strCellDescricao;
			StringBuilder sbProdutosNaoEncontrados;
			StringBuilder sbMsg;
			StringBuilder sbMsgErro;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			object oXL = null;
			object oWBs = null;
			object oWB = null;
			object oWSs = null;
			object oWS = null;
			object oCellQtdeEstoqueVenda = null;
			object oCellVlCustoIntermediario = null;
			object oCellPrecoMedioMercado = null;
			object oCellPrecoMinimoMercado = null;
			object oCellInteriorQtdeEstoqueVenda = null;
			object oCellInteriorPrecoMedioMercado = null;
			object oCellInteriorPrecoMinimoMercado = null;
			object oRange = null;
			object[,] oRangeValue = null;
			PlanilhaPrecosHeader planilhaPrecosHeader;
			List<PlanilhaPrecosColumn> vPlanilhaPrecosColunasObrigatorias;
			List<PlanilhaControleColumn> vPlanilhaControleColunasObrigatorias;
			PlanilhaPrecosColumn colunaPrecoConcorrente;
			PlanilhaPrecosPrecoConcorrente precoConcorrente;
			PlanilhaPrecosLinha linhaPlanilhaPrecos;
			List<PlanilhaPrecosLinha> vLinhaPlanilhaPrecos;
			PlanilhaControleHeader planilhaControleHeader;
			PlanilhaControleLinha linhaPlanilhaControle;
			List<PlanilhaPrecosLinha> vProdutosNaoEncontrados;
			ProdutoEstoqueVenda produtoEstoqueVenda;
			ProdutoLoja produtoLoja;
			ProdutoEstoqueVendaLojaConsolidado produtoConsolidado;
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

			#region [ Obtém nome do arquivo da planilha de controle e da planilha da ferramenta de preços ]
			strNomeArqPlanilhaControle = txtPlanilhaControle.Text;
			strNomeArqPlanilhaFerramentaPrecos = txtPlanilhaFerramentaPrecos.Text;
			#endregion

			#region [ Consistências ]
			if (strNomeArqPlanilhaControle.Length == 0)
			{
				strMsgErro = "É necessário selecionar a planilha de controle a ser processada!!";
				adicionaErro(strMsgErro);
				avisoErro(strMsgErro);
				return;
			}

			if (strNomeArqPlanilhaFerramentaPrecos.Length == 0)
			{
				strMsgErro = "É necessário selecionar a planilha da ferramenta de preços a ser processada!!";
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

			if (!File.Exists(strNomeArqPlanilhaFerramentaPrecos))
			{
				strMsgErro = "O arquivo da planilha da ferramenta de preços não existe!!\r\n" + strNomeArqPlanilhaFerramentaPrecos;
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
			if (!confirma("Confirma o processamento das planilhas?"))
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
					"        Planilha de controle: " + strNomeArqPlanilhaControle + "\r\n" +
					"        Planilha de preços: " + strNomeArqPlanilhaFerramentaPrecos;
			adicionaDisplay(strMsg);
			#endregion

			#region [ Backup da planilha de controle ]
			info(ModoExibicaoMensagemRodape.EmExecucao, "Criando backup da planilha de controle");
			strPathPlanilhaControleBackup = Path.GetDirectoryName(strNomeArqPlanilhaControle) + "\\Backup";
			try
			{
				if (!Directory.Exists(strPathPlanilhaControleBackup))
				{
					Directory.CreateDirectory(strPathPlanilhaControleBackup);
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha ao tentar criar o diretório para armazenar o backup da planilha de controle!!\r\n" + ex.ToString());
			}

			// Se não conseguiu criar o diretório, grava os arquivos no mesmo diretório em que está a planilha de controle
			if (!Directory.Exists(strPathPlanilhaControleBackup)) strPathPlanilhaControleBackup = Path.GetDirectoryName(strNomeArqPlanilhaControle);

			strNomeArqPlanilhaControleBackup = strPathPlanilhaControleBackup + "\\" +
												Path.GetFileNameWithoutExtension(strNomeArqPlanilhaControle) +
												"_" + Global.formataDataYyyyMmDdComSeparador(DateTime.Now, "-") +
												"_" + Global.formataHoraHhMmSsComSimbolo(DateTime.Now) +
												Path.GetExtension(strNomeArqPlanilhaControle);
			File.Copy(strNomeArqPlanilhaControle, strNomeArqPlanilhaControleBackup);
			if (!File.Exists(strNomeArqPlanilhaControleBackup))
			{
				adicionaErro("Falha ao tentar criar a cópia de backup do arquivo da planilha de controle");
			}
			else
			{
				strMsg = "Backup da planilha de controle realizado com sucesso: " + strNomeArqPlanilhaControleBackup;
				adicionaDisplay(strMsg);
			}
			#endregion

			try
			{
				#region [ Planilha de preços da ferramenta ]
				try // Finally
				{
					strMsg = "Inicialização do processo do Excel para leitura da planilha de preços";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Carregando dados da planilha de preços");

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

					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.Visible, (_ConsolidaDadosPlanilhaPrecosExcelVisible.Equals("0") ? false : true));
					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.DisplayAlerts, false);
					oWBs = ExcelAutomation.GetProperty(oXL, ExcelAutomation.PropertyType.Workbooks);
					ExcelAutomation.InvokeMethod(oWBs, ExcelAutomation.MethodType.Open, strNomeArqPlanilhaFerramentaPrecos);
					oWB = ExcelAutomation.GetProperty(oWBs, ExcelAutomation.PropertyType.Item, 1);
					oWSs = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.Worksheets, null);
					oWS = ExcelAutomation.GetProperty(oWSs, ExcelAutomation.PropertyType.Item, 1);
					ExcelAutomation.InvokeMethod(oWS, ExcelAutomation.MethodType.Select, null);

					#region [ Verifica se a planilha que está ativa realmente é a planilha de preços ]
					// Desaloca as referências anteriores, caso contrário, o processo do Excel não é encerrado ao final
					ExcelAutomation.NAR(oWS);
					oWS = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.ActiveSheet);
					strValue = (String)ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Name);
					if ((!strValue.ToUpper().Equals(NOME_WORKSHEET_PRECOS.ToUpper())) && (!strValue.ToUpper().Equals(NOME_WORKSHEET_PRECOS_V2.ToUpper())))
					{
						strMsg = "Falha ao tentar tornar ativa a planilha de preços (" + NOME_WORKSHEET_PRECOS + " ou " + NOME_WORKSHEET_PRECOS_V2 + ")\r\nPlanilha ativa: " + strValue;
						adicionaErro(strMsg);
						avisoErro(strMsg);
						return;
					}
					#endregion

					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.DisplayAlerts, true);

					#region [ Obtém a linha dos títulos da planilha ]
					iXlDadosMinIndex = 1;
					iXlDadosMaxIndex = 64;

					iNumLinha = 1;
					strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
					oRange = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strAux);
					oRangeValue = (object[,])ExcelAutomation.GetProperty(oRange, ExcelAutomation.PropertyType.Value);
					#endregion

					#region [ Colunas obrigatórias ]

					#region [ Obtém a posição das colunas obrigatórias ]
					vPlanilhaPrecosColunasObrigatorias = new List<PlanilhaPrecosColumn>();
					planilhaPrecosHeader = new PlanilhaPrecosHeader();

					planilhaPrecosHeader.Codigo.ColTitleEsperado = "Cod";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.Codigo);

					planilhaPrecosHeader.ProdutoDescricao.ColTitleEsperado = "Produto";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.ProdutoDescricao);

					planilhaPrecosHeader.QtdeLojasConcorrentes.ColTitleEsperado = "Lojas";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.QtdeLojasConcorrentes);

					planilhaPrecosHeader.Status.ColTitleEsperado = "Status";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.Status);

					planilhaPrecosHeader.SeuPreco.ColTitleEsperado = "Seu Preço";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.SeuPreco);

					planilhaPrecosHeader.Diferenca.ColTitleEsperado = "Diferença";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.Diferenca);

					planilhaPrecosHeader.Regra.ColTitleEsperado = "Regra";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.Regra);

					planilhaPrecosHeader.Sugestao.ColTitleEsperado = "Sugestão";
					vPlanilhaPrecosColunasObrigatorias.Add(planilhaPrecosHeader.Sugestao);

					qtdeColObrigatoriaEncontrada = 0;
					iColIndex = oRangeValue.GetLowerBound(1) - 1;
					while (true)
					{
						iColIndex++;

						if (iColIndex > oRangeValue.GetUpperBound(1)) break;

						if (oRangeValue[1, iColIndex] == null) continue;

						strTituloEncontrado = oRangeValue[1, iColIndex].ToString().Trim();
						strTituloEncontradoUppercase = strTituloEncontrado.ToUpper();

						foreach (var colObrigatoria in vPlanilhaPrecosColunasObrigatorias)
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
						if (qtdeColObrigatoriaEncontrada == vPlanilhaPrecosColunasObrigatorias.Count) break;
					} // while (true)
					#endregion

					#region [ Verifica se alguma coluna obrigatória não foi localizada ]
					blnTitulosOk = true;
					sbMsgErro = new StringBuilder("");
					foreach (var colObrigatoria in vPlanilhaPrecosColunasObrigatorias)
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
						strMsgErro = "A planilha '" + Path.GetFileName(strNomeArqPlanilhaFerramentaPrecos) + "' não possui os títulos corretos para as colunas!!\r\nVerifique se a planilha correta foi selecionada!!\r\n\r\n" + sbMsgErro.ToString();
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#endregion

					#region [ Colunas com preços dos concorrentes ]
					// As colunas de preços dos concorrentes estão entre as colunas 'Seu Preço' e 'Diferença'
					iColIndex = planilhaPrecosHeader.SeuPreco.ColIndex;
					while (true)
					{
						iColIndex++;

						if (iColIndex >= planilhaPrecosHeader.Diferenca.ColIndex) break;

						strTituloEncontrado = oRangeValue[1, iColIndex].ToString().Trim();
						strTituloEncontradoUppercase = strTituloEncontrado.ToUpper();

						colunaPrecoConcorrente = new PlanilhaPrecosColumn();
						colunaPrecoConcorrente.ColIndex = iColIndex;
						colunaPrecoConcorrente.ColTitle = strTituloEncontrado;

						planilhaPrecosHeader.ColunasPrecoConcorrente.Add(colunaPrecoConcorrente);
					}
					#endregion

					#region [ Carrega os dados da planilha para lista em memória ]
					strMsg = "Iniciando leitura das linhas de dados da planilha de preços";
					adicionaDisplay(strMsg);

					vLinhaPlanilhaPrecos = new List<PlanilhaPrecosLinha>();
					qtdeLinhasVaziasConsecutivas = 0;
					qtdeLinhaDadosPlanilhaPrecos = 0;
					iXlDadosMaxIndex = 0;
					foreach (var item in vPlanilhaPrecosColunasObrigatorias)
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

						// IMPORTANTE: Produtos que ainda não foram associados a um SKU na ferramenta de comparação de preços estarão c/ a célula 'Cod' vazia
						strCellCodigo = "";
						strCellDescricao = "";
						if (oRangeValue[1, planilhaPrecosHeader.Codigo.ColIndex] != null) strCellCodigo = oRangeValue[1, planilhaPrecosHeader.Codigo.ColIndex].ToString().Trim();
						if (oRangeValue[1, planilhaPrecosHeader.ProdutoDescricao.ColIndex] != null) strCellDescricao = oRangeValue[1, planilhaPrecosHeader.ProdutoDescricao.ColIndex].ToString().Trim();

						if ((strCellCodigo.Length == 0) && (strCellDescricao.Length == 0))
						{
							qtdeLinhasVaziasConsecutivas++;
							continue;
						}

						// Se ainda não possui um SKU associado, não há o que processar, segue p/ o próximo produto, mas sem incrementar o contador de linhas vazias consecutivas
						if (strCellCodigo.Length == 0) continue;
						#endregion

						qtdeLinhaDadosPlanilhaPrecos++;
						// Encontrou linha com dados, portanto, zera contador de linhas vazias consecutivas
						qtdeLinhasVaziasConsecutivas = 0;
						strMsg = "Leitura da " + qtdeLinhaDadosPlanilhaPrecos.ToString() + "ª linha de dados da planilha de preços: SKU " + oRangeValue[1, planilhaPrecosHeader.Codigo.ColIndex].ToString();
						adicionaDisplay(strMsg);
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);

						#region [ Carrega dados para objeto da classe LinhaPlanilhaPrecos e adiciona na lista ]
						linhaPlanilhaPrecos = new PlanilhaPrecosLinha();

						#region [ Código ]
						if (oRangeValue[1, planilhaPrecosHeader.Codigo.ColIndex] != null)
						{
							linhaPlanilhaPrecos.Codigo = oRangeValue[1, planilhaPrecosHeader.Codigo.ColIndex].ToString().Trim();
							linhaPlanilhaPrecos.CodigoFormatado = Global.normalizaCodigoProduto(linhaPlanilhaPrecos.Codigo);
						}
						#endregion

						#region [ Descrição do produto ]
						if (oRangeValue[1, planilhaPrecosHeader.ProdutoDescricao.ColIndex] != null)
						{
							linhaPlanilhaPrecos.ProdutoDescricao = oRangeValue[1, planilhaPrecosHeader.ProdutoDescricao.ColIndex].ToString().Trim();
						}
						#endregion

						#region [ Qtde de lojas ]
						if (oRangeValue[1, planilhaPrecosHeader.QtdeLojasConcorrentes.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaPrecosHeader.QtdeLojasConcorrentes.ColIndex].GetType().FullName.ToUpper().Equals("System.Int32".ToUpper()))
								{
									linhaPlanilhaPrecos.QtdeLojas = (int)oRangeValue[1, planilhaPrecosHeader.QtdeLojasConcorrentes.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaPrecosHeader.QtdeLojasConcorrentes.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaPrecos.QtdeLojas = (int)Global.converteInteiro(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = "Falha ao ler o conteúdo da célula '" + planilhaPrecosHeader.QtdeLojasConcorrentes.ColTitle + "' do SKU " + linhaPlanilhaPrecos.Codigo + " (Type = " + oRangeValue[1, planilhaPrecosHeader.QtdeLojasConcorrentes.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaPrecosHeader.QtdeLojasConcorrentes.ColIndex].ToString() + ")";
								adicionaErro(strMsg);
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
							}
						}
						#endregion

						#region [ Status ]
						if (oRangeValue[1, planilhaPrecosHeader.Status.ColIndex] != null)
						{
							linhaPlanilhaPrecos.Status = oRangeValue[1, planilhaPrecosHeader.Status.ColIndex].ToString().Trim();
						}
						#endregion

						#region [ Diferença ]
						if (oRangeValue[1, planilhaPrecosHeader.Diferenca.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaPrecosHeader.Diferenca.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaPrecos.Diferenca = (decimal)oRangeValue[1, planilhaPrecosHeader.Diferenca.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaPrecosHeader.Diferenca.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaPrecos.Diferenca = Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = "Falha ao ler o conteúdo da célula '" + planilhaPrecosHeader.Diferenca.ColTitle + "' do SKU " + linhaPlanilhaPrecos.Codigo + " (Type = " + oRangeValue[1, planilhaPrecosHeader.Diferenca.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaPrecosHeader.Diferenca.ColIndex].ToString() + ")";
								adicionaErro(strMsg);
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
							}
						}
						#endregion

						#region [ Regra ]
						if (oRangeValue[1, planilhaPrecosHeader.Regra.ColIndex] != null)
						{
							linhaPlanilhaPrecos.Regra = oRangeValue[1, planilhaPrecosHeader.Regra.ColIndex].ToString().Trim();
						}
						#endregion

						#region [ Preço sugestão ]
						if (oRangeValue[1, planilhaPrecosHeader.Sugestao.ColIndex] != null)
						{
							try
							{
								if (oRangeValue[1, planilhaPrecosHeader.Sugestao.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
								{
									linhaPlanilhaPrecos.PrecoSugestao = (decimal)oRangeValue[1, planilhaPrecosHeader.Sugestao.ColIndex];
								}
								else
								{
									strValue = oRangeValue[1, planilhaPrecosHeader.Sugestao.ColIndex].ToString().Trim();
									if (strValue.Length > 0)
									{
										linhaPlanilhaPrecos.PrecoSugestao = Global.converteNumeroDecimal(strValue);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = "Falha ao ler o conteúdo da célula '" + planilhaPrecosHeader.Sugestao.ColTitle + "' do SKU " + linhaPlanilhaPrecos.Codigo + " (Type = " + oRangeValue[1, planilhaPrecosHeader.Sugestao.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, planilhaPrecosHeader.Sugestao.ColIndex].ToString() + ")";
								adicionaErro(strMsg);
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
							}
						}
						#endregion

						#region [ Preços dos concorrentes ]
						foreach (var colConcorrente in planilhaPrecosHeader.ColunasPrecoConcorrente)
						{
							precoConcorrente = new PlanilhaPrecosPrecoConcorrente();
							precoConcorrente.ColIndex = colConcorrente.ColIndex;
							precoConcorrente.ColTitle = colConcorrente.ColTitle;
							precoConcorrente.ColTitleEsperado = colConcorrente.ColTitleEsperado;
							if (oRangeValue[1, colConcorrente.ColIndex] != null)
							{
								precoConcorrente.CellValue = oRangeValue[1, colConcorrente.ColIndex].ToString().Trim();
								if (!oRangeValue[1, colConcorrente.ColIndex].GetType().FullName.ToUpper().Equals("System.String".ToUpper()))
								{
									vlAux = null;
									try
									{
										if (oRangeValue[1, colConcorrente.ColIndex].GetType().FullName.ToUpper().Equals("System.Decimal".ToUpper()))
										{
											vlAux = (decimal)oRangeValue[1, colConcorrente.ColIndex];
										}
										else
										{
											strValue = oRangeValue[1, colConcorrente.ColIndex].ToString().Trim();
											if (strValue.Length > 0)
											{
												vlAux = Global.converteNumeroDecimal(strValue);
											}
										}
									}
									catch (Exception ex)
									{
										vlAux = null;
										strMsg = "Falha ao ler o conteúdo da célula '" + colConcorrente.ColTitle + "' do SKU " + linhaPlanilhaPrecos.Codigo + " (Type = " + oRangeValue[1, colConcorrente.ColIndex].GetType().FullName + ", Value = " + oRangeValue[1, colConcorrente.ColIndex].ToString() + ")";
										adicionaErro(strMsg);
										Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + strMsg + "\r\n" + ex.ToString());
									}
									precoConcorrente.Preco = vlAux;
								}
							}

							linhaPlanilhaPrecos.ColunasPrecoConcorrente.Add(precoConcorrente);
						}
						#endregion

						#region [ Calcula preço mínimo e médio dos concorrentes ]
						strNomeConcorrentePrecoMinimo = "";
						vlPrecoMinimoConcorrente = null;
						vlPrecoMedioConcorrente = null;
						qtdeConcorrentesPrecoMedio = 0;
						vlAcumuladoTotalPrecoMedioConcorrente = 0m;
						foreach (var concorrente in linhaPlanilhaPrecos.ColunasPrecoConcorrente)
						{
							if (concorrente.Preco != null)
							{
								#region [ Preço mínimo ]
								if (vlPrecoMinimoConcorrente == null)
								{
									vlPrecoMinimoConcorrente = concorrente.Preco;
									strNomeConcorrentePrecoMinimo = concorrente.ColTitle;
								}
								else
								{
									if (concorrente.Preco < vlPrecoMinimoConcorrente)
									{
										vlPrecoMinimoConcorrente = concorrente.Preco;
										strNomeConcorrentePrecoMinimo = concorrente.ColTitle;
									}
								}
								#endregion

								#region [ Preço médio ]
								qtdeConcorrentesPrecoMedio++;
								vlAcumuladoTotalPrecoMedioConcorrente += (decimal)concorrente.Preco;
								#endregion
							}
						}

						if (qtdeConcorrentesPrecoMedio > 0)
						{
							vlPrecoMedioConcorrente = vlAcumuladoTotalPrecoMedioConcorrente / qtdeConcorrentesPrecoMedio;
							vlPrecoMedioConcorrente = Global.arredondaParaMonetario((decimal)vlPrecoMedioConcorrente);
						}
						#endregion

						linhaPlanilhaPrecos.NomeConcorrentePrecoMinimo = strNomeConcorrentePrecoMinimo;
						linhaPlanilhaPrecos.PrecoMinimo = vlPrecoMinimoConcorrente;
						linhaPlanilhaPrecos.PrecoMedio = vlPrecoMedioConcorrente;

						vLinhaPlanilhaPrecos.Add(linhaPlanilhaPrecos);
						#endregion
					}

					if (vLinhaPlanilhaPrecos.Count == 0)
					{
						strMsgErro = "A planilha de preços não possui dados!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}

					strMsg = qtdeLinhaDadosPlanilhaPrecos.ToString() + " linhas de dados carregadas da planilha de preços";
					adicionaDisplay(strMsg);
					#endregion
				}
				finally
				{
					strMsg = "Fechamento da planilha de preços e encerramento do processo do Excel";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Fechamento da planilha de preços e encerramento do processo do Excel");
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

				#region [ Planilha de controle ]
				try // Finally
				{
					strMsg = "Inicialização do processo do Excel para processamento da planilha de controle";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Consolidação dos dados na planilha de controle");

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

					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.Visible, (_ConsolidaDadosPlanilhaControleExcelVisible.Equals("0") ? false : true));
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

					#region [ Consolida dados da planilha de controle ]
					strMsg = "Iniciando consolidação dos dados da planilha de controle";
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

						#endregion

						#region [ Obtém quantidade disponível no estoque e preço de custo médio ]
						produtoConsolidado = new ProdutoEstoqueVendaLojaConsolidado();
						vlNovoCustoIntermediario = 0m;
						foreach (var ambiente in FMain.contextoBD.Ambientes)
						{
							produtoEstoqueVenda = ambiente.produtoDAO.GetProdutoUnificadoCustoIntermediario(linhaPlanilhaControle.SkuFormatado, out strMsgErroAux);
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
									Global.gravaLogAtividade("[" + ambiente.NomeAmbiente + "] SKU " + linhaPlanilhaControle.SkuFormatado + ": Qtde estoque = " + Global.formataInteiro(produtoEstoqueVenda.qtdeEstoqueVenda) + ", Custo intermediário = " + Global.formataMoeda(produtoEstoqueVenda.vlCustoIntermediario));

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
												produtoLoja = ambiente.produtoDAO.GetProdutoLoja(item.fabricante_item, item.produto_item, ambiente.NumeroLojaArclube, out strMsgErroAux);
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
										produtoLoja = ambiente.produtoDAO.GetProdutoLoja(produtoConsolidado.fabricante, produtoConsolidado.produto, ambiente.NumeroLojaArclube, out strMsgErroAux);
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
						if (produtoConsolidado.isCadastrado && produtoConsolidado.isVendavel)
						{
							#region [ Produto cadastrado e vendável ]
							
							#region [ Qtde estoque venda ]
							strCelQtdeEstoqueVenda = Global.excel_converte_numeracao_digito_para_letra(planilhaControleHeader.QtdeEstoque.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellQtdeEstoqueVenda);
							oCellQtdeEstoqueVenda = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelQtdeEstoqueVenda);
							ExcelAutomation.NAR(oCellInteriorQtdeEstoqueVenda);
							oCellInteriorQtdeEstoqueVenda = ExcelAutomation.GetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Interior);
							// Atualiza valor
							ExcelAutomation.SetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Value, (double)produtoConsolidado.qtdeEstoqueVenda);

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
									if (item.qtdeEstoqueVenda>0)
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
							#endregion

							#region [ Custo médio ]
							// Se o custo médio for zero (estoque zerado), não atualiza a célula p/ preservar o último valor conhecido, caso contrário, o valor zerado seria
							// atualizado no Magento e, consequentemente, o produto seria exibido como indiponível e c/ preço zero.
							if ((double)produtoConsolidado.vlCustoIntermediario > 0d)
							{
								strCelVlCustoIntermediario = Global.excel_converte_numeracao_digito_para_letra(planilhaControleHeader.ValorCustoMedio.ColIndex) + iNumLinha.ToString();
								ExcelAutomation.NAR(oCellVlCustoIntermediario);
								oCellVlCustoIntermediario = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelVlCustoIntermediario);
								ExcelAutomation.SetProperty(oCellVlCustoIntermediario, ExcelAutomation.PropertyType.Value, (double)produtoConsolidado.vlCustoIntermediario);
							}
							#endregion

							#endregion
						}
						else
						{
							#region [ Produto não cadastrado e/ou não vendável ]
							// Se o produto não está cadastrado/disponível no sistema, zera o estoque na planilha
							strCelQtdeEstoqueVenda = Global.excel_converte_numeracao_digito_para_letra(planilhaControleHeader.QtdeEstoque.ColIndex) + iNumLinha.ToString();
							ExcelAutomation.NAR(oCellQtdeEstoqueVenda);
							oCellQtdeEstoqueVenda = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelQtdeEstoqueVenda);
							ExcelAutomation.NAR(oCellInteriorQtdeEstoqueVenda);
							oCellInteriorQtdeEstoqueVenda = ExcelAutomation.GetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Interior);
							// Atualiza valor: zero
							ExcelAutomation.SetProperty(oCellQtdeEstoqueVenda, ExcelAutomation.PropertyType.Value, 0d);

							#region [ Altera cor de fundo p/ indicar que não está disponível ]
							ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
							ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
							ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
							ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.TintAndShade, 0d);
							ExcelAutomation.SetProperty(oCellInteriorQtdeEstoqueVenda, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
							#endregion
							#endregion
						}
						#endregion

						#region [ Analisa se houve alteração no preço médio de mercado e/ou preço mínimo de mercado ]

						#region [ Célula do preço médio de mercado ]
						strCelPrecoMedioMercado = Global.excel_converte_numeracao_digito_para_letra(planilhaControleHeader.ValorMedioMercado.ColIndex) + iNumLinha.ToString();
						ExcelAutomation.NAR(oCellPrecoMedioMercado);
						oCellPrecoMedioMercado = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelPrecoMedioMercado);
						ExcelAutomation.NAR(oCellInteriorPrecoMedioMercado);
						oCellInteriorPrecoMedioMercado = ExcelAutomation.GetProperty(oCellPrecoMedioMercado, ExcelAutomation.PropertyType.Interior);
						#endregion

						#region [ Célula do preço mínimo de mercado ]
						strCelPrecoMinimoMercado = Global.excel_converte_numeracao_digito_para_letra(planilhaControleHeader.ValorMinimoMercado.ColIndex) + iNumLinha.ToString();
						ExcelAutomation.NAR(oCellPrecoMinimoMercado);
						oCellPrecoMinimoMercado = ExcelAutomation.GetProperty(oWS, ExcelAutomation.PropertyType.Range, strCelPrecoMinimoMercado);
						ExcelAutomation.NAR(oCellInteriorPrecoMinimoMercado);
						oCellInteriorPrecoMinimoMercado = ExcelAutomation.GetProperty(oCellPrecoMinimoMercado, ExcelAutomation.PropertyType.Interior);
						#endregion

						try
						{
							linhaPlanilhaPrecos = vLinhaPlanilhaPrecos.Single(p => p.CodigoFormatado.Equals(linhaPlanilhaControle.SkuFormatado));
						}
						catch (Exception)
						{
							// O método Single() lança uma exception se houver 0 (zero) ou mais do que 1 elemento no resultado
							linhaPlanilhaPrecos = null;
						}

						if (linhaPlanilhaPrecos == null)
						{
							#region [ Preço médio de mercado ]
							if (linhaPlanilhaControle.ValorMedioMercado != null)
							{
								// Limpa o valor da célula
								ExcelAutomation.SetProperty(oCellPrecoMedioMercado, ExcelAutomation.PropertyType.Value, null);
							}

							#region [ Altera cor da célula: fundo indicando informação ausente na planilha de preços ]
							ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
							#endregion

							#endregion

							#region [ Preço mínimo de mercado ]
							if (linhaPlanilhaControle.ValorMinimoMercado != null)
							{
								// Limpa o valor da célula
								ExcelAutomation.SetProperty(oCellPrecoMinimoMercado, ExcelAutomation.PropertyType.Value, null);
							}

							#region [ Altera cor da célula: fundo indicando informação ausente na planilha de preços ]
							ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
							ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
							#endregion

							#endregion
						}
						else
						{
							linhaPlanilhaPrecos.ProcessadoStatus = true;

							#region [ Preço médio de mercado ]
							if ((linhaPlanilhaPrecos.PrecoMedio == null) && (linhaPlanilhaControle.ValorMedioMercado == null))
							{
								#region [ Altera cor da célula: fundo indicando informação ausente na planilha de preços ]
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
								#endregion
							}
							else if ((linhaPlanilhaPrecos.PrecoMedio == null) || (linhaPlanilhaControle.ValorMedioMercado == null))
							{
								if (linhaPlanilhaPrecos.PrecoMedio == null)
								{
									#region [ Informação ausente na planilha de preços ]
									// Limpa o valor da célula
									ExcelAutomation.SetProperty(oCellPrecoMedioMercado, ExcelAutomation.PropertyType.Value, null);

									#region [ Altera cor da célula: fundo indicando informação ausente na planilha de preços ]
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
									#endregion
									#endregion
								}
								else
								{
									#region [ Atualiza c/ o valor da planilha de preços ]
									// Altera o valor
									ExcelAutomation.SetProperty(oCellPrecoMedioMercado, ExcelAutomation.PropertyType.Value, (double)linhaPlanilhaPrecos.PrecoMedio);

									#region [ Altera cor da célula: fundo indicando que o valor exibido é diferente que o da versão anterior ]
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_REALCE_ALTERACAO);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
									#endregion
									#endregion
								}
							}
							else if (Math.Abs((double)linhaPlanilhaPrecos.PrecoMedio - (double)linhaPlanilhaControle.ValorMedioMercado) <= MAX_VALOR_MARGEM_ERRO_PRECO)
							{
								#region [ Altera cor da célula: fundo vazio indicando que o valor permanece igual ]
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlNone);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.TintAndShade, 0);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0);
								#endregion
							}
							else
							{
								#region [ Altera cor da célula: fundo indicando que o valor exibido é diferente que o da versão anterior ]
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_REALCE_ALTERACAO);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMedioMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
								// Altera o valor
								ExcelAutomation.SetProperty(oCellPrecoMedioMercado, ExcelAutomation.PropertyType.Value, (double)linhaPlanilhaPrecos.PrecoMedio);
								#endregion
							}
							#endregion

							#region [ Preço mínimo de mercado ]
							if ((linhaPlanilhaPrecos.PrecoMinimo == null) && (linhaPlanilhaControle.ValorMinimoMercado == null))
							{
								#region [ Altera cor da célula: fundo indicando informação ausente na planilha de preços ]
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
								#endregion
							}
							else if ((linhaPlanilhaPrecos.PrecoMinimo == null) || (linhaPlanilhaControle.ValorMinimoMercado == null))
							{
								if (linhaPlanilhaPrecos.PrecoMinimo == null)
								{
									#region [ Informação ausente na planilha de preços ]
									// Limpa o valor da célula
									ExcelAutomation.SetProperty(oCellPrecoMinimoMercado, ExcelAutomation.PropertyType.Value, null);

									#region [ fundo indicando informação ausente na planilha de preços ]
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_INFO_NOT_FOUND);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
									#endregion
									#endregion
								}
								else
								{
									#region [ Atualiza c/ o valor da planilha de preços ]
									// Altera o valor
									ExcelAutomation.SetProperty(oCellPrecoMinimoMercado, ExcelAutomation.PropertyType.Value, (double)linhaPlanilhaPrecos.PrecoMinimo);

									#region [ fundo indicando que o valor exibido é diferente que o da versão anterior ]
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_REALCE_ALTERACAO);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
									ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
									#endregion
									#endregion
								}
							}
							else if (Math.Abs((double)linhaPlanilhaPrecos.PrecoMinimo - (double)linhaPlanilhaControle.ValorMinimoMercado) <= MAX_VALOR_MARGEM_ERRO_PRECO)
							{
								#region [ Altera cor da célula: fundo vazio indicando que o valor permanece igual ]
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlNone);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.TintAndShade, 0);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0);
								#endregion
							}
							else
							{
								#region [ Altera cor da célula: fundo indicando que o valor exibido é diferente que o da versão anterior ]
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.Color, COR_FUNDO_REALCE_ALTERACAO);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.TintAndShade, 0d);
								ExcelAutomation.SetProperty(oCellInteriorPrecoMinimoMercado, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
								// Altera o valor
								ExcelAutomation.SetProperty(oCellPrecoMinimoMercado, ExcelAutomation.PropertyType.Value, (double)linhaPlanilhaPrecos.PrecoMinimo);
								#endregion
							}
							#endregion
						}
						#endregion
					} // while (true)

					strMsg = qtdeLinhaDadosPlanilhaControle.ToString() + " linhas de dados processadas na planilha de controle";
					adicionaDisplay(strMsg);
					#endregion

					#region [ Salva planilha de controle ]
					strMsg = "Salvando planilha de controle";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Salvando planilha de controle");

					ExcelAutomation.InvokeMethod(oWB, ExcelAutomation.MethodType.Save, null);
					ExcelAutomation.SetProperty(oXL, ExcelAutomation.PropertyType.DisplayAlerts, true);
					#endregion
				}
				finally
				{
					strMsg = "Fechamento da planilha de controle e encerramento do processo do Excel";
					adicionaDisplay(strMsg);
					info(ModoExibicaoMensagemRodape.EmExecucao, "Fechamento da planilha de controle e encerramento do processo do Excel");
					try
					{
						ExcelAutomation.NAR(oCellInteriorQtdeEstoqueVenda);
						ExcelAutomation.NAR(oCellInteriorPrecoMedioMercado);
						ExcelAutomation.NAR(oCellInteriorPrecoMinimoMercado);
						ExcelAutomation.NAR(oCellQtdeEstoqueVenda);
						ExcelAutomation.NAR(oCellVlCustoIntermediario);
						ExcelAutomation.NAR(oCellPrecoMedioMercado);
						ExcelAutomation.NAR(oCellPrecoMinimoMercado);

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

				#region [ Verifica se na planilha de preços haviam códigos de produtos que não existem na planilha de controle ]
				sbProdutosNaoEncontrados = new StringBuilder("");
				try
				{
					vProdutosNaoEncontrados = vLinhaPlanilhaPrecos.FindAll(p => p.ProcessadoStatus == false);
				}
				catch (Exception)
				{
					vProdutosNaoEncontrados = null;
				}

				if (vProdutosNaoEncontrados != null)
				{
					if (vProdutosNaoEncontrados.Count > 0)
					{
						foreach (var item in vProdutosNaoEncontrados)
						{
							strMsg = "        " + item.Codigo + " - " + item.ProdutoDescricao;
							sbProdutosNaoEncontrados.AppendLine(strMsg);
						}

						strMsg = "ATENÇÃO!!\r\nOs seguintes produtos da planilha de preços não foram encontrados na planilha de controle:\r\n" + sbProdutosNaoEncontrados.ToString();
						adicionaDisplay(strMsg);
						aviso(strMsg);
					}
				}
				#endregion

				#region [ Grava log ]
				strMsg = "Sucesso no processamento da planilha de exportação de produtos do e-commerce (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + "): planilha de preços = " + strNomeArqPlanilhaFerramentaPrecos + " (" + qtdeLinhaDadosPlanilhaPrecos.ToString() + " linhas de dados), planilha de controle = " + strNomeArqPlanilhaControle + " (" + qtdeLinhaDadosPlanilhaControle.ToString() + " linhas de dados)";
				Global.gravaLogAtividade(strMsg);
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.CXLSEC.LogOperacao.PROCESSAMENTO_PLANILHA;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				#region [ Mensagem de sucesso ]
				info(ModoExibicaoMensagemRodape.Normal);
				strMsg = "Processamento concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!!";
				adicionaDisplay(strMsg);
				strMsg = "Processamento concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!!" +
						"\r\n\r\nDeseja abrir a planilha '" + Path.GetFileName(strNomeArqPlanilhaControle) + "' agora?";
				if (confirma(strMsg))
				{
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

		#region [ FConsolidaDadosPlanilha ]

		#region [ FConsolidaDadosPlanilha_Load ]
		private void FConsolidaDadosPlanilha_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				txtPlanilhaControle.Text = "";
				txtPlanilhaFerramentaPrecos.Text = "";
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

		#region [ FConsolidaDadosPlanilha_Shown ]
		private void FConsolidaDadosPlanilha_Shown(object sender, EventArgs e)
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
					openFileDialogPrecos.InitialDirectory = pathPlanilhaFerramentaPrecosValorDefault();
					openFileDialogPrecos.FileName = fileNamePlanilhaFerramentaPrecosValorDefault();

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

		#region [ FConsolidaDadosPlanilha_FormClosing ]
		private void FConsolidaDadosPlanilha_FormClosing(object sender, FormClosingEventArgs e)
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

		#region [ btnSelecionaPlanilhaFerramentaPrecos ]

		#region [ btnSelecionaPlanilhaFerramentaPrecos_Click ]
		private void btnSelecionaPlanilhaFerramentaPrecos_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaPlanilhaPrecos();
		}
		#endregion

		#endregion

		#region [ btnConsolidaPlanilha ]

		#region [ btnConsolidaPlanilha_Click ]
		private void btnConsolidaPlanilha_Click(object sender, EventArgs e)
		{
			trataBotaoConsolidaPlanilha();
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

		#region [ txtPlanilhaFerramentaPrecos ]
		#region [ txtPlanilhaFerramentaPrecos_Enter ]
		private void txtPlanilhaFerramentaPrecos_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtPlanilhaFerramentaPrecos_DoubleClick ]
		private void txtPlanilhaFerramentaPrecos_DoubleClick(object sender, EventArgs e)
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

		#region [ btnAbrePlanilhaPrecos ]

		#region [ btnAbrePlanilhaPrecos_Click ]
		private void btnAbrePlanilhaPrecos_Click(object sender, EventArgs e)
		{
			trataBotaoAbrePlanilhaPrecos();
		}
		#endregion

		#endregion

		#endregion
	}
}
