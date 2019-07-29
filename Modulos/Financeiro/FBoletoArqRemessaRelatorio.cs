#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	public partial class FBoletoArqRemessaRelatorio : Financeiro.FModelo
	{
		#region [ Atributos ]

		#region [ Diversos ]
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

		String[] _linhasArqRemessa;
		LinhaHeaderArquivoRemessa _linhaHeader = new LinhaHeaderArquivoRemessa();
		LinhaTraillerArquivoRemessa _linhaTrailler = new LinhaTraillerArquivoRemessa();
		BoletoCedente _boletoCedente;
		#endregion

		#region [ Controle da impressão ]
		private int _intImpressaoArqRemessaIdxLinha = 0;
		private int _intImpressaoNumPagina = 0;
		private String _strImpressaoDataEmissao;
		private int _intQtdeTotalRegistros;
		private decimal _vlTotalRegistros;
		Impressao impressao;
		const String NOME_FONTE_DEFAULT = "Courier New";
		Font fonteTitulo;
		Font fonteListagem;
		Font fonteDataEmissao;
		Font fonteNumPagina;
		Font fonteAtual;
		Brush brushPadrao;
		Pen penTracoTitulo;
		Pen penTracoPontilhado;
		float cxInicio;
		float cxFim;
		float cyInicio;
		float cyFim;
		float cyRodapeNumPagina;
		float larguraUtil;
		float alturaUtil;

		#region [ Colunas Listagem ]
		float ixNomeSacado;
		float wxNomeSacado;
		float ixNumInscricaoSacado;
		float wxNumInscricaoSacado;
		float ixEndereco;
		float wxEndereco;
		float ixLoja;
		float wxLoja;
		float ixPedido;
		float wxPedido;
		float ixNumeroDocumento;
		float wxNumeroDocumento;
		float ixDtVencto;
		float wxDtVencto;
		float ixVlTitulo;
		float wxVlTitulo;
		float ESPACAMENTO_COLUNAS;
		#endregion

		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoArqRemessaRelatorio()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ printPreviewListagem ]
		private void printPreviewListagem()
		{
			#region [ Consistência ]
			if (_linhasArqRemessa == null)
			{
				avisoErro("É necessário selecionar um arquivo de remessa!!");
				return;
			}
			if (_linhasArqRemessa.Length <= 2)
			{
				avisoErro("Arquivo de remessa selecionado não possui dados!!");
				return;
			}
			#endregion

			prnPreviewConsulta.WindowState = FormWindowState.Maximized;
			prnPreviewConsulta.MinimizeBox = true;
			prnPreviewConsulta.Text = Global.Cte.Aplicativo.M_ID + " - Visualização da Impressão";
			prnPreviewConsulta.PrintPreviewControl.Zoom = 1;
			prnPreviewConsulta.PrintPreviewControl.AutoZoom = true;
			prnPreviewConsulta.FormBorderStyle = FormBorderStyle.Sizable;
			prnPreviewConsulta.ShowDialog();
		}
		#endregion

		#region [ imprimeListagem ]
		private void imprimeListagem()
		{
			#region [ Consistência ]
			if (_linhasArqRemessa == null)
			{
				avisoErro("É necessário selecionar um arquivo de remessa!!");
				return;
			}
			if (_linhasArqRemessa.Length <= 2)
			{
				avisoErro("Arquivo de remessa selecionado não possui dados!!");
				return;
			}
			#endregion

			prnDocConsulta.Print();
		}
		#endregion

		#region [ printerDialog ]
		private void printerDialog()
		{
			prnDialogConsulta.ShowDialog();
		}
		#endregion

		#region [ pathBoletoArquivoRemessaValorDefault ]
		private String pathBoletoArquivoRemessaValorDefault()
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
			if (Global.Usuario.Defaults.pathBoletoArquivoRemessaRelatorio.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.pathBoletoArquivoRemessaRelatorio))
				{
					strResp = Global.Usuario.Defaults.pathBoletoArquivoRemessaRelatorio;
				}
			}
			return strResp;
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtArqRemessa.Text = "";
			grdBoletos.Rows.Clear();
			lblTotalRegistros.Text = "";
			lblTotalizacaoValor.Text = "";
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Arquivo de Remessa ]
			if (txtArqRemessa.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo de remessa que será carregado!!");
				return false;
			}
			if (!File.Exists(txtArqRemessa.Text))
			{
				avisoErro("O arquivo de remessa informado não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ carregaGridBoletos ]
		private void carregaGridBoletos()
		{
			#region [ Declarações ]
			String[] v;
			String strIdBoletoItem;
			String strPedido;
			String strLoja;
			String strAux;
			int idBoletoItem;
			List<String> listaPedidoLoja;
			LinhaRegistroTipo1ArquivoRemessa linhaRegistro = new LinhaRegistroTipo1ArquivoRemessa();
			int intLinhaGrid = 0;
			decimal vlTitulo;
			decimal vlTotal = 0m;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
			try
			{
				#region [ Limpa campos ]
				grdBoletos.Rows.Clear();
				lblTotalRegistros.Text = "";
				lblTotalizacaoValor.Text = "";
				#endregion

				#region [ Consistência ]
				if (_linhasArqRemessa == null)
				{
					avisoErro("É necessário selecionar um arquivo de remessa!!");
					return;
				}
				if (_linhasArqRemessa.Length <= 2)
				{
					avisoErro("Arquivo de remessa selecionado não possui dados!!");
					return;
				}
				#endregion

				#region [ Consistência do header e trailler ]
				if ((!_linhaHeader.identificacaoRegistro.valor.Equals("0")) ||
					(!_linhaHeader.identificacaoArquivoRemessa.valor.Equals("1")) ||
					(!_linhaHeader.literalRemessa.valor.ToUpper().Equals("REMESSA")))
				{
					avisoErro("Arquivo de remessa com header inválido!!");
					return;
				}

				if (!_linhaTrailler.identificacaoRegistro.valor.Equals("9"))
				{
					avisoErro("Arquivo de remessa com trailler inválido!!");
					return;
				}
				#endregion

				#region [ Preenche grid ]
				if (_linhasArqRemessa.Length > 2) grdBoletos.Rows.Add((_linhasArqRemessa.Length - 2) / 2);
				for (int i = 1; i < (_linhasArqRemessa.Length - 1); i += 2)
				{
					linhaRegistro.CarregaDados(_linhasArqRemessa[i]);

					vlTitulo = Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor);
					vlTotal += vlTitulo;

					grdBoletos.Rows[intLinhaGrid].Cells["colSacado"].Value = linhaRegistro.nomeSacado.valor.Trim() + " (" + Global.formataCnpjCpf(linhaRegistro.numInscricaoSacado.valor) + ")";
					grdBoletos.Rows[intLinhaGrid].Cells["colEndereco"].Value = linhaRegistro.enderecoCompleto.valor.Trim() + " - CEP: " + Global.formataCep(linhaRegistro.cep.valor.Trim() + linhaRegistro.sufixoCep.valor.Trim());
					grdBoletos.Rows[intLinhaGrid].Cells["colNumeroDocumento"].Value = linhaRegistro.numDocumento.valor;
					grdBoletos.Rows[intLinhaGrid].Cells["colDataVencto"].Value = Global.formataDataCampoArquivoDdMmYyParaDDMMYYYYComSeparador(linhaRegistro.dataVenctoTitulo.valor);
					grdBoletos.Rows[intLinhaGrid].Cells["colValorTitulo"].Value = Global.formataMoeda(vlTitulo);

					strIdBoletoItem = "";
					if (linhaRegistro.numControleParticipante.valor.IndexOf('=') != -1)
					{
						v = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = v[1];
					}
					grdBoletos.Rows[intLinhaGrid].Cells["colIdBoletoItem"].Value = strIdBoletoItem;

					strPedido = "";
					strLoja = "";
					idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
					listaPedidoLoja = BoletoDAO.obtemBoletoInformacaoPedidoLoja(idBoletoItem);
					for (int j = 0; j < listaPedidoLoja.Count; j++)
					{
						strAux = listaPedidoLoja[j];
						if (strAux == null) continue;
						if (strAux.Length == 0) continue;
						v = strAux.Split('=');
						if (strPedido.Length > 0) strPedido += ", ";
						strPedido += v[0];
						if (strLoja.Length > 0) strLoja += ", ";
						strLoja += v[1];
					}
					grdBoletos.Rows[intLinhaGrid].Cells["colLoja"].Value = strLoja;
					grdBoletos.Rows[intLinhaGrid].Cells["colPedido"].Value = strPedido;

					intLinhaGrid++;
				}
				#endregion

				#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
				for (int i = 0; i < grdBoletos.Rows.Count; i++)
				{
					if (grdBoletos.Rows[i].Selected)
					{
						grdBoletos.Rows[i].Selected = false;
						break;
					}
				}
				#endregion

				lblTotalRegistros.Text = Global.formataInteiro(intLinhaGrid);
				lblTotalizacaoValor.Text = Global.formataMoeda(vlTotal);

				Global.Usuario.Defaults.pathBoletoArquivoRemessaRelatorio = Path.GetDirectoryName(txtArqRemessa.Text);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoSelecionaArqRemessa ]
		private void trataBotaoSelecionaArqRemessa()
		{
			#region [ Declarações ]
			DialogResult dr;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			try
			{
				openFileDialog.InitialDirectory = pathBoletoArquivoRemessaValorDefault();
				dr = openFileDialog.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialog.FileName.Length > 0) && (txtArqRemessa.Text.Length > 0))
				{
					if (openFileDialog.FileName.ToUpper().Equals(txtArqRemessa.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos ]
				grdBoletos.Rows.Clear();
				lblTotalRegistros.Text = "";
				lblTotalizacaoValor.Text = "";
				#endregion

				txtArqRemessa.Text = openFileDialog.FileName;

				#region [ Carrega dados do arquivo em array ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "lendo dados do arquivo");
				_linhasArqRemessa = File.ReadAllLines(txtArqRemessa.Text, encode);
				#endregion

				#region [ Consistência ]
				if (_linhasArqRemessa == null)
				{
					avisoErro("É necessário selecionar um arquivo de remessa!!");
					return;
				}
				if (_linhasArqRemessa.Length <= 2)
				{
					avisoErro("Arquivo de remessa selecionado não possui dados!!");
					return;
				}
				#endregion

				#region [ Carrega header/trailler ]
				_linhaHeader.CarregaDados(_linhasArqRemessa[0]);
				_linhaTrailler.CarregaDados(_linhasArqRemessa[_linhasArqRemessa.Length - 1]);
				_boletoCedente = BoletoCedenteDAO.getBoletoCedenteByCodigoEmpresa(_linhaHeader.codigoEmpresa.valor);
				#endregion

				carregaGridBoletos();
				grdBoletos.Focus();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoListagem ]
		private void trataBotaoListagem()
		{
			if (rbSaidaVisualizacao.Checked)
				printPreviewListagem();
			else if (rbSaidaImpressora.Checked)
				imprimeListagem();
			else
				avisoErro("Selecione o tipo de saída: Impressora ou Print Preview");
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoArqRemessaRelatorio ]

		#region [ FBoletoArqRemessaRelatorio_Load ]
		private void FBoletoArqRemessaRelatorio_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCampos();

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

		#region [ FBoletoArqRemessaRelatorio_Shown ]
		private void FBoletoArqRemessaRelatorio_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String nomeOpcaoSaidaDefault;
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Opção de saída ]
					if (Global.Usuario.Defaults.relatorioArqRemessaTipoSaida.Trim().Length > 0)
					{
						nomeOpcaoSaidaDefault = Global.Usuario.Defaults.relatorioArqRemessaTipoSaida;
						if (rbSaidaImpressora.Name.Equals(nomeOpcaoSaidaDefault))
							rbSaidaImpressora.Checked = true;
						else if (rbSaidaVisualizacao.Name.Equals(nomeOpcaoSaidaDefault))
							rbSaidaVisualizacao.Checked = true;
					}
					#endregion

					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion
					
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

		#region [ FBoletoArqRemessaRelatorio_FormClosing ]
		private void FBoletoArqRemessaRelatorio_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Atualiza opção default ]
			if (rbSaidaImpressora.Checked)
				Global.Usuario.Defaults.relatorioArqRemessaTipoSaida = rbSaidaImpressora.Name;
			else if (rbSaidaVisualizacao.Checked)
				Global.Usuario.Defaults.relatorioArqRemessaTipoSaida = rbSaidaVisualizacao.Name;
			#endregion

			#region [ Exibe o form principal ]
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
			#endregion
		}
		#endregion

		#endregion

		#region [ btnSelecionaArqRemessa ]

		#region [ btnSelecionaArqRemessa_Click ]
		private void btnSelecionaArqRemessa_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArqRemessa();
		}
		#endregion

		#endregion

		#region [ txtArqRemessa ]

		#region [ txtArqRemessa_Enter ]
		private void txtArqRemessa_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtArqRemessa_DoubleClick ]
		private void txtArqRemessa_DoubleClick(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#endregion

		#region [ btnListagemArqRemessa ]

		#region [ btnListagemArqRemessa_Click ]
		private void btnListagemArqRemessa_Click(object sender, EventArgs e)
		{
			trataBotaoListagem();
		}
		#endregion

		#endregion

		#region [ btnPrinterDialog ]

		#region [ btnPrinterDialog_Click ]
		private void btnPrinterDialog_Click(object sender, EventArgs e)
		{
			printerDialog();
		}
		#endregion

		#endregion

		#region [ Impressão ]

		#region [ prnDocConsulta_QueryPageSettings ]
		private void prnDocConsulta_QueryPageSettings(object sender, System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			executaQueryPageSettingsListagem(ref sender, ref e);
		}
		#endregion

		#region [ prnDocConsulta_BeginPrint ]
		private void prnDocConsulta_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
		{
			executaBeginPrintListagem(ref sender, ref e);
		}
		#endregion

		#region [ prnDocConsulta_PrintPage ]
		private void prnDocConsulta_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			executaPrintPageListagem(ref sender, ref e);
		}
		#endregion

		#region [ executaQueryPageSettingsListagem ]
		private void executaQueryPageSettingsListagem(ref object sender, ref System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			e.PageSettings.Landscape = true;
		}
		#endregion

		#region [ executaBeginPrintListagem ]
		private void executaBeginPrintListagem(ref object sender, ref System.Drawing.Printing.PrintEventArgs e)
		{
			#region [ Consistência ]
			if (_linhasArqRemessa == null)
			{
				e.Cancel = true;
				return;
			}
			if (_linhasArqRemessa.Length <= 2)
			{
				e.Cancel = true;
				return;
			}
			#endregion

			_intImpressaoArqRemessaIdxLinha = 1;  // Índice zero é o header e o último índice é o trailler
			_intImpressaoNumPagina = 0;
			_strImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
			_intQtdeTotalRegistros = 0;
			_vlTotalRegistros = 0m;

			prnDocConsulta.DefaultPageSettings.Landscape = true;

			impressao = new Impressao(prnDocConsulta.DefaultPageSettings.Landscape);

			#region [ Prepara elementos de impressão ]
			fonteTitulo = new Font(NOME_FONTE_DEFAULT, 10f, FontStyle.Bold);
			fonteListagem = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteDataEmissao = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteNumPagina = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			brushPadrao = new SolidBrush(Color.Black);
			penTracoTitulo = new Pen(brushPadrao, .5f);
			penTracoPontilhado = Impressao.criaPenTracoPontilhado();
			#endregion
		}
		#endregion

		#region [ executaPrintPageListagem ]
		private void executaPrintPageListagem(ref object sender, ref System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			float cx;
			float cy;
			float hMax;
			RectangleF r;
			int idBoletoItem;
			String[] v;
			String strAux;
			String strTexto;
			String strIdBoletoItem;
			String strPedido;
			String strLoja;
			int intLinhasImpressasNestaPagina = 0;
			decimal vlTitulo;
			LinhaRegistroTipo1ArquivoRemessa linhaRegistro = new LinhaRegistroTipo1ArquivoRemessa();
			List<String> listaPedidoLoja;
			#endregion

			#region [ Consistência ]
			if (_linhasArqRemessa == null)
			{
				e.Cancel = true;
				return;
			}
			if (_linhasArqRemessa.Length <= 2)
			{
				e.Cancel = true;
				return;
			}
			#endregion

			#region [ Contador de página ]
			_intImpressaoNumPagina++;
			#endregion

			e.Graphics.PageUnit = GraphicsUnit.Millimeter;
			if (_intImpressaoNumPagina == 1)
			{
				#region [ Medidas do papel ]
				prnDocConsulta.DocumentName = "Listagem do arquivo de remessa";
				cxInicio = impressao.getLeftMarginInMm(e);
				larguraUtil = impressao.getWidthInMm(e);
				cxFim = cxInicio + larguraUtil;
				cyInicio = impressao.getTopMarginInMm(e);
				alturaUtil = impressao.getHeightInMm(e);
				cyFim = cyInicio + alturaUtil;
				cyRodapeNumPagina = cyFim - fonteNumPagina.GetHeight(e.Graphics) - 1;
				#endregion

				#region [ Layout das colunas da listagem ]
				ESPACAMENTO_COLUNAS = 2f;
				wxNomeSacado = 50f;
				wxNumInscricaoSacado = 30f;
				wxLoja = 10f;
				wxPedido = 15f;
				wxNumeroDocumento = 19f;
				wxDtVencto = 16f;
				wxVlTitulo = 23f;
				wxEndereco = larguraUtil
							 - wxNomeSacado            // A 1ª coluna não tem espaçamento
							 - ESPACAMENTO_COLUNAS - wxNumInscricaoSacado
							 - ESPACAMENTO_COLUNAS     // Espaçamento da própria coluna "Endereço"
							 - ESPACAMENTO_COLUNAS - wxLoja
							 - ESPACAMENTO_COLUNAS - wxPedido
							 - ESPACAMENTO_COLUNAS - wxNumeroDocumento
							 - ESPACAMENTO_COLUNAS - wxDtVencto
							 - ESPACAMENTO_COLUNAS - wxVlTitulo;
				ixNomeSacado = cxInicio;
				ixNumInscricaoSacado = ixNomeSacado + wxNomeSacado + ESPACAMENTO_COLUNAS;
				ixEndereco = ixNumInscricaoSacado + wxNumInscricaoSacado + ESPACAMENTO_COLUNAS;
				ixLoja = ixEndereco + wxEndereco + ESPACAMENTO_COLUNAS;
				ixPedido = ixLoja + wxLoja + ESPACAMENTO_COLUNAS;
				ixNumeroDocumento = ixPedido + wxPedido + ESPACAMENTO_COLUNAS;
				ixDtVencto = ixNumeroDocumento + wxNumeroDocumento + ESPACAMENTO_COLUNAS;
				ixVlTitulo = ixDtVencto + wxDtVencto + ESPACAMENTO_COLUNAS;
				#endregion
			}

			cy = cyInicio;

			#region [ Título ]
			strTexto = "LISTAGEM DO ARQUIVO DE REMESSA";
			fonteAtual = fonteTitulo;
			cx = cxInicio + (larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			cy += 1f;
			#endregion

			#region [ Informações no cabeçalho ]

			#region [ Nome do arquivo ]
			strTexto = "Arquivo de remessa: " + txtArqRemessa.Text;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Empresa cedente / Carteira ]
			strTexto = "Empresa Cedente...: " + _linhaHeader.nomeEmpresa.valor;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil / 2);
			strTexto = "Carteira.....: " + _boletoCedente.carteira + "  -  Agência: " + _boletoCedente.agencia + '-' + _boletoCedente.digito_agencia + "  -  Conta: " + _boletoCedente.conta + '-' + _boletoCedente.digito_conta;
			fonteAtual = fonteListagem;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Código da empresa / Data no banco e Data do crédito ]
			strTexto = "Código da Empresa.: " + _boletoCedente.codigo_empresa + "  -  Nº sequencial de remessa: " + _linhaHeader.numSequencialRemessa.valor;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil / 2);
			strTexto = "Data da gravação do arquivo: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataGravacaoArquivo.valor));
			fonteAtual = fonteListagem;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Títulos da listagem ]
			cy += .5f;
			fonteAtual = fonteListagem;
			strTexto = "NOME SACADO";
			cx = ixNomeSacado;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Nº INSCR SACADO";
			cx = ixNumInscricaoSacado;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "ENDEREÇO";
			cx = ixEndereco;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "LOJA";
			cx = ixLoja;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "PEDIDO";
			cx = ixPedido;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Nº DOCUMENTO";
			cx = ixNumeroDocumento;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "DT VENCTO";
			cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "VL TÍTULO";
			cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cy += fonteAtual.GetHeight(e.Graphics);
			cy += .5f;
			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Laço para listagem ]
			while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
				   (_intImpressaoArqRemessaIdxLinha < (_linhasArqRemessa.Length - 1)))
			{
				linhaRegistro.CarregaDados(_linhasArqRemessa[_intImpressaoArqRemessaIdxLinha]);

				#region [ Consulta BD para obter pedido+loja ]
				strPedido = "";
				strLoja = "";
				strIdBoletoItem = "";
				if (linhaRegistro.numControleParticipante.valor.IndexOf('=') != -1)
				{
					v = linhaRegistro.numControleParticipante.valor.Split('=');
					strIdBoletoItem = v[1];
				}
				if (strIdBoletoItem.Length > 0)
				{
					idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
					listaPedidoLoja = BoletoDAO.obtemBoletoInformacaoPedidoLoja(idBoletoItem);
					for (int i = 0; i < listaPedidoLoja.Count; i++)
					{
						strAux = listaPedidoLoja[i];
						if (strAux == null) continue;
						if (strAux.Length == 0) continue;
						v = strAux.Split('=');
						if (strPedido.Length > 0) strPedido += ", ";
						strPedido += v[0];
						if (strLoja.Length > 0) strLoja += ", ";
						strLoja += v[1];
					}
				}
				#endregion

				hMax = fonteListagem.GetHeight(e.Graphics);

				#region [ Nome Sacado ]
				strTexto = linhaRegistro.nomeSacado.valor.Trim();
				while ((e.Graphics.MeasureString(strTexto, fonteAtual).Width > wxNomeSacado) && (strTexto.Length > 0))
				{
					strTexto = strTexto.Substring(0, strTexto.Length - 1);
				}
				cx = ixNomeSacado;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Nº Inscrição Sacado ]
				strTexto = Global.formataCnpjCpf(linhaRegistro.numInscricaoSacado.valor);
				cx = ixNumInscricaoSacado;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Endereço ]
				strTexto = linhaRegistro.enderecoCompleto.valor.Trim() + " - CEP: " + Global.formataCep(linhaRegistro.cep.valor.Trim() + linhaRegistro.sufixoCep.valor.Trim());
				while ((e.Graphics.MeasureString(strTexto, fonteAtual).Width > wxEndereco) && (strTexto.Length > 0))
				{
					strTexto = strTexto.Substring(0, strTexto.Length - 1);
				}
				cx = ixEndereco;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Loja ]
				cx = ixLoja;
				r = new RectangleF(ixLoja, cy, wxLoja, 20);
				strTexto = strLoja;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxLoja).Height);
				#endregion

				#region [ Pedido ]
				cx = ixPedido;
				r = new RectangleF(ixPedido, cy, wxPedido, 20);
				strTexto = strPedido;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxPedido).Height);
				#endregion

				#region [ Nº Documento ]
				strTexto = linhaRegistro.numDocumento.valor;
				cx = ixNumeroDocumento;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Data do vencimento ]
				strTexto = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
				cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Valor do título ]
				vlTitulo = Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor);
				strTexto = Global.formataMoeda(vlTitulo);
				cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				cy += hMax;

				_intQtdeTotalRegistros++;
				_vlTotalRegistros += vlTitulo;

				intLinhasImpressasNestaPagina++;
				_intImpressaoArqRemessaIdxLinha += 2; // Lembre-se: sempre há um registro tipo 1 e um tipo 2 juntos

				#region [ Na última linha não imprime o tracejado ]
				if (_intImpressaoArqRemessaIdxLinha < (_linhasArqRemessa.Length - 1))
				{
					#region [ Traço pontilhado ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoPontilhado, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion
				}
				#endregion
			}
			#endregion

			#region [ Tem mais páginas para imprimir? ]
			if (_intImpressaoArqRemessaIdxLinha < (_linhasArqRemessa.Length - 1))
			{
				e.HasMorePages = true;
			}
			else
			{
				e.HasMorePages = false;

				#region [ Há espaço suficiente? ]
				if ((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 10))
				{
					if (intLinhasImpressasNestaPagina > 0)
					{
						#region [ Traço ]
						cy += 1f;
						e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
						cy += 1f;
						#endregion
					}
					else cy += .5f;

					#region [ Imprime os totais ]
					fonteAtual = fonteListagem;
					cx = cxInicio;
					strTexto = "TOTAL";
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataInteiro(_intQtdeTotalRegistros) + " registro(s)";
					cx = ixNumInscricaoSacado;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(_vlTotalRegistros);
					cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion
				}
				else e.HasMorePages = true;
				#endregion
			}
			#endregion

			#region [ Imprime nº página ]
			strTexto = "Página: " + _intImpressaoNumPagina.ToString().PadLeft(2, ' ');
			fonteAtual = fonteNumPagina;
			cy = cyRodapeNumPagina;
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion
		}
		#endregion

		#endregion

		#endregion
	}
}
