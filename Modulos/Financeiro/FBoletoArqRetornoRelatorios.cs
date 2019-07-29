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
	public partial class FBoletoArqRetornoRelatorios : Financeiro.FModelo
	{
		#region [ Enum ]
		private enum eRelatorioArqRetorno
		{
			NENHUM = 0,
			LISTAGEM_TODAS_OCORRENCIAS = 1,
			LISTAGEM_REGISTROS_REJEITADOS = 2
		}
		#endregion

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

		String[] _linhasArqRetorno;
		List<TipoLinhaResumoRelArqRetorno> _listaLinhasResumoArqRetorno;
		List<TipoLinhaResumoRelArqRetorno> _listaTotalizacaoOcorrencia;
		LinhaHeaderArquivoRetorno _linhaHeader = new LinhaHeaderArquivoRetorno();
		LinhaTraillerArquivoRetorno _linhaTrailler = new LinhaTraillerArquivoRetorno();
		BoletoCedente _boletoCedente;
		#endregion

		#region [ Controle da impressão ]
		eRelatorioArqRetorno _opcaoRelatorioSelecionado = eRelatorioArqRetorno.NENHUM;
		private int _intImpressaoArqRetornoIdxLinha = 0;
		private int _intImpressaoResumoArqRetornoIdxLinha = 0;
		private int _intImpressaoNumPagina = 0;
		private int _intQtdeRejeicoes;
		private decimal _vlTotalRejeicoes;
		private String _strImpressaoDataEmissao;
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
		float ixOcorrencia;
		float wxOcorrencia;
		float ixNossoNumero;
		float wxNossoNumero;
		float ixNumeroTitulo;
		float wxNumeroTitulo;
		float ixNomeSacado;
		float wxNomeSacado;
		float ixDtPagto;
		float wxDtPagto;
		float ixDtVencto;
		float wxDtVencto;
		float ixVlTitulo;
		float wxVlTitulo;
		float ixVlPago;
		float wxVlPago;
		float ixVlOscilacao;
		float wxVlOscilacao;
		float ixBcoCobr;
		float wxBcoCobr;
		float ixAgCobr;
		float wxAgCobr;
		float ixMotivosOcorrencia;
		float wxMotivosOcorrencia;
		float ixIdentifCedente;
		float wxIdentifCedente;
		float ESPACAMENTO_COLUNAS;
		#endregion

		#region [ Colunas do resumo ]
		float ixResumoOcorrencia;
		float wxResumoOcorrencia;
		float ixResumoQtdeTitulos;
		float wxResumoQtdeTitulos;
		float ixResumoVlTotal;
		float wxResumoVlTotal;
		#endregion

		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoArqRetornoRelatorios()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ printPreviewListagemTodasOcorrencias ]
		private void printPreviewListagemTodasOcorrencias()
		{
			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				avisoErro("É necessário selecionar um arquivo de retorno!!");
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
			{
				avisoErro("Arquivo de retorno selecionado não possui dados!!");
				return;
			}
			#endregion

			_opcaoRelatorioSelecionado = eRelatorioArqRetorno.LISTAGEM_TODAS_OCORRENCIAS;

			prnPreviewConsulta.WindowState = FormWindowState.Maximized;
			prnPreviewConsulta.MinimizeBox = true;
			prnPreviewConsulta.Text = Global.Cte.Aplicativo.M_ID + " - Visualização da Impressão";
			prnPreviewConsulta.PrintPreviewControl.Zoom = 1;
			prnPreviewConsulta.PrintPreviewControl.AutoZoom = true;
			prnPreviewConsulta.FormBorderStyle = FormBorderStyle.Sizable;
			prnPreviewConsulta.ShowDialog();
		}
		#endregion

		#region [ imprimeListagemTodasOcorrencias ]
		private void imprimeListagemTodasOcorrencias()
		{
			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				avisoErro("É necessário selecionar um arquivo de retorno!!");
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
			{
				avisoErro("Arquivo de retorno selecionado não possui dados!!");
				return;
			}
			#endregion

			_opcaoRelatorioSelecionado = eRelatorioArqRetorno.LISTAGEM_TODAS_OCORRENCIAS;

			prnDocConsulta.Print();
		}
		#endregion

		#region [ printPreviewListagemRegistrosRejeitados ]
		private void printPreviewListagemRegistrosRejeitados()
		{
			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				avisoErro("É necessário selecionar um arquivo de retorno!!");
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
			{
				avisoErro("Arquivo de retorno selecionado não possui dados!!");
				return;
			}
			#endregion

			_opcaoRelatorioSelecionado = eRelatorioArqRetorno.LISTAGEM_REGISTROS_REJEITADOS;

			prnPreviewConsulta.WindowState = FormWindowState.Maximized;
			prnPreviewConsulta.MinimizeBox = true;
			prnPreviewConsulta.Text = Global.Cte.Aplicativo.M_ID + " - Visualização da Impressão";
			prnPreviewConsulta.PrintPreviewControl.Zoom = 1;
			prnPreviewConsulta.PrintPreviewControl.AutoZoom = true;
			prnPreviewConsulta.FormBorderStyle = FormBorderStyle.Sizable;
			prnPreviewConsulta.ShowDialog();
		}
		#endregion

		#region [ imprimeListagemRegistrosRejeitados ]
		private void imprimeListagemRegistrosRejeitados()
		{
			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				avisoErro("É necessário selecionar um arquivo de retorno!!");
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
			{
				avisoErro("Arquivo de retorno selecionado não possui dados!!");
				return;
			}
			#endregion

			_opcaoRelatorioSelecionado = eRelatorioArqRetorno.LISTAGEM_REGISTROS_REJEITADOS;

			prnDocConsulta.Print();
		}
		#endregion

		#region [ printerDialog ]
		private void printerDialog()
		{
			prnDialogConsulta.ShowDialog();
		}
		#endregion

		#region [ pathBoletoArquivoRetornoValorDefault ]
		private String pathBoletoArquivoRetornoValorDefault()
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
			if (Global.Usuario.Defaults.pathBoletoArquivoRetornoRelatorios.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.pathBoletoArquivoRetornoRelatorios))
				{
					strResp = Global.Usuario.Defaults.pathBoletoArquivoRetornoRelatorios;
				}
			}
			return strResp;
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtArqRetorno.Text = "";
			grdBoletos.Rows.Clear();
			lblTotalRegistros.Text = "";
			lblCedenteNome.Text = "";
			lblCedenteCarteira.Text = "";
			lblCedenteAgencia.Text = "";
			lblCedenteConta.Text = "";
			lblCedenteCodigoEmpresa.Text = "";
			lblCedenteNumAvisoBancario.Text = "";
			lblCedenteDataBanco.Text = "";
			lblCedenteDataCredito.Text = "";
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Arquivo de Retorno ]
			if (txtArqRetorno.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo de retorno que será carregado!!");
				return false;
			}
			if (!File.Exists(txtArqRetorno.Text))
			{
				avisoErro("O arquivo de retorno informado não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ preencheCamposInfArqRetorno ]
		private void preencheCamposInfArqRetorno()
		{
			lblCedenteNome.Text = _linhaHeader.nomeEmpresa.valor;
			lblCedenteCarteira.Text = _boletoCedente.carteira;
			lblCedenteAgencia.Text = _boletoCedente.agencia + '-' + _boletoCedente.digito_agencia;
			lblCedenteConta.Text = _boletoCedente.conta + '-' + _boletoCedente.digito_conta;
			lblCedenteCodigoEmpresa.Text = _boletoCedente.codigo_empresa;
			lblCedenteNumAvisoBancario.Text = _linhaHeader.numAvisoBancario.valor;
			lblCedenteDataBanco.Text = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataGravacaoArquivo.valor));
			lblCedenteDataCredito.Text = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataCredito.valor));
		}
		#endregion

		#region [ carregaGridBoletos ]
		private void carregaGridBoletos()
		{
			#region [ Declarações ]
			String[] v;
			String strDescricaoMotivoOcorrencia;
			String strIdBoletoItem;
			LinhaRegistroTipo1ArquivoRetorno linhaRegistro = new LinhaRegistroTipo1ArquivoRetorno();
			int intLinhaGrid = 0;
			List<Global.TipoDescricaoMotivoOcorrencia> listaMotivoOcorrencia;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
			try
			{
				#region [ Limpa campos ]
				grdBoletos.Rows.Clear();
				lblTotalRegistros.Text = "";
				#endregion

				#region [ Consistência ]
				if (_linhasArqRetorno == null)
				{
					avisoErro("É necessário selecionar um arquivo de retorno!!");
					return;
				}
				if (_linhasArqRetorno.Length <= 2)
				{
					avisoErro("Arquivo de retorno selecionado não possui dados!!");
					return;
				}
				#endregion

				#region [ Consistência do header e trailler ]
				if ((!_linhaHeader.identificacaoRegistro.valor.Equals("0")) ||
					(!_linhaHeader.identificacaoArquivoRetorno.valor.Equals("2")) ||
					(!_linhaHeader.literalRetorno.valor.ToUpper().Equals("RETORNO")))
				{
					avisoErro("Arquivo de retorno com header inválido!!");
					return;
				}

				if ((!_linhaTrailler.identificacaoRegistro.valor.Equals("9")) ||
					(!_linhaTrailler.identificacaoRetorno.valor.Equals("2")) ||
					(!_linhaTrailler.codigoBanco.valor.Equals("237")))
				{
					avisoErro("Arquivo de retorno com trailler inválido!!");
					return;
				}
				#endregion

				#region [ Preenche grid ]
				if (_linhasArqRetorno.Length > 2) grdBoletos.Rows.Add(_linhasArqRetorno.Length - 2);
				for (int i = 1; i < (_linhasArqRetorno.Length - 1); i++)
				{
					linhaRegistro.CarregaDados(_linhasArqRetorno[i]);
					grdBoletos.Rows[intLinhaGrid].Cells["numeroDocumento"].Value = linhaRegistro.numeroDocumento.valor;
					grdBoletos.Rows[intLinhaGrid].Cells["dataVenctoTitulo"].Value = Global.formataDataCampoArquivoDdMmYyParaDDMMYYYYComSeparador(linhaRegistro.dataVenctoTitulo.valor);
					grdBoletos.Rows[intLinhaGrid].Cells["valorTitulo"].Value = Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor));
					grdBoletos.Rows[intLinhaGrid].Cells["identificacaoOcorrencia"].Value = linhaRegistro.identificacaoOcorrencia.valor + " - " + Global.decodificaIdentificacaoOcorrencia(linhaRegistro.identificacaoOcorrencia.valor);

					#region [ Motivos da ocorrência ]
					strDescricaoMotivoOcorrencia = "";
					if (linhaRegistro.identificacaoOcorrencia.valor.Equals("19"))
					{
						strDescricaoMotivoOcorrencia = linhaRegistro.motivoCodigoOcorrencia19 + " - " + Global.decodificaMotivoOcorrencia19(linhaRegistro.motivoCodigoOcorrencia19.valor);
					}
					else
					{
						listaMotivoOcorrencia = Global.decodificaMotivoOcorrencia(linhaRegistro.identificacaoOcorrencia.valor, linhaRegistro.motivosRejeicoes.valor);
						for (int j = 0; j < listaMotivoOcorrencia.Count; j++)
						{
							if (listaMotivoOcorrencia[j].descricaoMotivoOcorrencia.Length > 0)
							{
								if (strDescricaoMotivoOcorrencia.Length > 0) strDescricaoMotivoOcorrencia += "\n";
								strDescricaoMotivoOcorrencia += listaMotivoOcorrencia[j].motivoOcorrencia + " - " + listaMotivoOcorrencia[j].descricaoMotivoOcorrencia;
							}
						}
					}
					grdBoletos.Rows[intLinhaGrid].Cells["motivosRejeicoes"].Value = strDescricaoMotivoOcorrencia;

					strIdBoletoItem = "";
					if (linhaRegistro.numControleParticipante.valor.IndexOf('=') != -1)
					{
						v = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = v[1];
					}
					grdBoletos.Rows[intLinhaGrid].Cells["id_boleto_item"].Value = strIdBoletoItem;
					#endregion

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

				Global.Usuario.Defaults.pathBoletoArquivoRetornoRelatorios = Path.GetDirectoryName(txtArqRetorno.Text);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoSelecionaArqRetorno ]
		private void trataBotaoSelecionaArqRetorno()
		{
			#region[ Declarações ]
			DialogResult dr;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			try
			{
				openFileDialog.InitialDirectory = pathBoletoArquivoRetornoValorDefault();
				dr = openFileDialog.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialog.FileName.Length > 0) && (txtArqRetorno.Text.Length > 0))
				{
					if (openFileDialog.FileName.ToUpper().Equals(txtArqRetorno.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos ]
				limpaCampos();
				#endregion

				txtArqRetorno.Text = openFileDialog.FileName;

				#region [ Carrega dados do arquivo em array ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "lendo dados do arquivo");
				_linhasArqRetorno = File.ReadAllLines(txtArqRetorno.Text, encode);
				#endregion

				#region [ Consistência ]
				if (_linhasArqRetorno == null)
				{
					avisoErro("É necessário selecionar um arquivo de retorno!!");
					return;
				}
				if (_linhasArqRetorno.Length <= 2)
				{
					avisoErro("Arquivo de retorno selecionado não possui dados!!");
					return;
				}
				#endregion

				#region [ Carrega header/trailler ]
				_linhaHeader.CarregaDados(_linhasArqRetorno[0]);
				_linhaTrailler.CarregaDados(_linhasArqRetorno[_linhasArqRetorno.Length - 1]);
				_boletoCedente = BoletoCedenteDAO.getBoletoCedenteByCodigoEmpresa(_linhaHeader.codigoEmpresa.valor);
				#endregion

				preencheCamposInfArqRetorno();
				carregaGridBoletos();
				grdBoletos.Focus();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoListagemTodasOcorrencias ]
		private void trataBotaoListagemTodasOcorrencias()
		{
			if (rbSaidaVisualizacao.Checked)
				printPreviewListagemTodasOcorrencias();
			else if (rbSaidaImpressora.Checked)
				imprimeListagemTodasOcorrencias();
			else
				avisoErro("Selecione o tipo de saída: Impressora ou Print Preview");
		}
		#endregion

		#region [ trataBotaoListagemRegistrosRejeitados ]
		private void trataBotaoListagemRegistrosRejeitados()
		{
			if (rbSaidaVisualizacao.Checked)
				printPreviewListagemRegistrosRejeitados();
			else if (rbSaidaImpressora.Checked)
				imprimeListagemRegistrosRejeitados();
			else
				avisoErro("Selecione o tipo de saída: Impressora ou Print Preview");
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoArqRetornoRelatorios ]

		#region [ FBoletoArqRetornoRelatorios_Load ]
		private void FBoletoArqRetornoRelatorios_Load(object sender, EventArgs e)
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

		#region [ FBoletoArqRetornoRelatorios_Shown ]
		private void FBoletoArqRetornoRelatorios_Shown(object sender, EventArgs e)
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
					if (Global.Usuario.Defaults.relatorioArqRetornoTipoSaida.Trim().Length > 0)
					{
						nomeOpcaoSaidaDefault = Global.Usuario.Defaults.relatorioArqRetornoTipoSaida;
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

		#region [ FBoletoArqRetornoRelatorios_FormClosing ]
		private void FBoletoArqRetornoRelatorios_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Atualiza opção default ]
			if (rbSaidaImpressora.Checked)
				Global.Usuario.Defaults.relatorioArqRetornoTipoSaida = rbSaidaImpressora.Name;
			else if (rbSaidaVisualizacao.Checked)
				Global.Usuario.Defaults.relatorioArqRetornoTipoSaida = rbSaidaVisualizacao.Name;
			#endregion

			#region [ Exibe o form principal ]
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
			#endregion
		}
		#endregion

		#endregion

		#region [ btnSelecionaArqRetorno ]

		#region [ btnSelecionaArqRetorno_Click ]
		private void btnSelecionaArqRetorno_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArqRetorno();
		}
		#endregion

		#endregion

		#region [ txtArqRetorno ]

		#region [ txtArqRetorno_Enter ]
		private void txtArqRetorno_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtArqRetorno_DoubleClick ]
		private void txtArqRetorno_DoubleClick(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#endregion

		#region [ btnListagemTodasOcorrencias ]

		#region [ btnListagemTodasOcorrencias_Click ]
		private void btnListagemTodasOcorrencias_Click(object sender, EventArgs e)
		{
			trataBotaoListagemTodasOcorrencias();
		}
		#endregion

		#endregion

		#region [ btnListagemRegistrosRejeitados ]

		#region [ btnListagemRegistrosRejeitados_Click ]
		private void btnListagemRegistrosRejeitados_Click(object sender, EventArgs e)
		{
			trataBotaoListagemRegistrosRejeitados();
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
			if (_opcaoRelatorioSelecionado == eRelatorioArqRetorno.LISTAGEM_TODAS_OCORRENCIAS)
				executaQueryPageSettingsListagemTodasOcorrencias(ref sender, ref e);
			else if (_opcaoRelatorioSelecionado == eRelatorioArqRetorno.LISTAGEM_REGISTROS_REJEITADOS)
				executaQueryPageSettingsListagemRegistrosRejeitados(ref sender, ref e);
		}
		#endregion

		#region [ prnDocConsulta_BeginPrint ]
		private void prnDocConsulta_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
		{
			if (_opcaoRelatorioSelecionado == eRelatorioArqRetorno.LISTAGEM_TODAS_OCORRENCIAS)
				executaBeginPrintListagemTodasOcorrencias(ref sender, ref e);
			else if (_opcaoRelatorioSelecionado == eRelatorioArqRetorno.LISTAGEM_REGISTROS_REJEITADOS)
				executaBeginPrintListagemRegistrosRejeitados(ref sender, ref e);
		}
		#endregion

		#region [ prnDocConsulta_PrintPage ]
		private void prnDocConsulta_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			if (_opcaoRelatorioSelecionado == eRelatorioArqRetorno.LISTAGEM_TODAS_OCORRENCIAS)
				executaPrintPageListagemTodasOcorrencias(ref sender, ref e);
			else if (_opcaoRelatorioSelecionado == eRelatorioArqRetorno.LISTAGEM_REGISTROS_REJEITADOS)
				executaPrintPageListagemRegistrosRejeitados(ref sender, ref e);
		}
		#endregion

		#region [ executaQueryPageSettingsListagemTodasOcorrencias ]
		private void executaQueryPageSettingsListagemTodasOcorrencias(ref object sender, ref System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			e.PageSettings.Landscape = true;
		}
		#endregion

		#region [ executaQueryPageSettingsListagemRegistrosRejeitados ]
		private void executaQueryPageSettingsListagemRegistrosRejeitados(ref object sender, ref System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			e.PageSettings.Landscape = true;
		}
		#endregion

		#region [ executaBeginPrintListagemTodasOcorrencias ]
		private void executaBeginPrintListagemTodasOcorrencias(ref object sender, ref System.Drawing.Printing.PrintEventArgs e)
		{
			#region [ Declarações ]
			bool blnAchou;
			decimal vlAux;
			String strIdentificacaoOcorrencia;
			LinhaRegistroTipo1ArquivoRetorno linhaRegistro = new LinhaRegistroTipo1ArquivoRetorno();
			TipoLinhaResumoRelArqRetorno itemTotalizacaoOcorrencia = null;
			TipoLinhaResumoRelArqRetorno itemTotalTarifasRegCobranca = new TipoLinhaResumoRelArqRetorno("TOTAL TARIFAS REGISTRO COBRANÇA", 0, 0m);
			TipoLinhaResumoRelArqRetorno itemTotalTarifasProtesto = new TipoLinhaResumoRelArqRetorno("TOTAL TARIFAS DE PROTESTO", 0, 0m);
			TipoLinhaResumoRelArqRetorno itemTotalTarifasCustasProtesto = new TipoLinhaResumoRelArqRetorno("TOTAL TARIFAS CUSTAS PROTESTO", 0, 0m);
			TipoLinhaResumoRelArqRetorno itemTotalTitulosPagosBradesco = new TipoLinhaResumoRelArqRetorno("TÍTULOS PAGOS NO BRADESCO", 0, 0m);
			TipoLinhaResumoRelArqRetorno itemTotalTitulosPagosOutrosBancos = new TipoLinhaResumoRelArqRetorno("TÍTULOS PAGOS EM OUTROS BANCOS", 0, 0m);
			#endregion

			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				e.Cancel = true;
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
			{
				e.Cancel = true;
				return;
			}
			#endregion

			prnDocConsulta.DocumentName = "Listagem de todas as ocorrências do arquivo de retorno";

			_intImpressaoArqRetornoIdxLinha = 1;  // Índice zero é o header e o último índice é o trailler
			_intImpressaoResumoArqRetornoIdxLinha = 0;
			_intImpressaoNumPagina = 0;
			_strImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);

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

			#region [ Totaliza quantidade/valor das ocorrências ]
			_listaTotalizacaoOcorrencia = new List<TipoLinhaResumoRelArqRetorno>();

			for (int i = 1; i < _linhasArqRetorno.Length - 1; i++)
			{
				linhaRegistro.CarregaDados(_linhasArqRetorno[i]);
				strIdentificacaoOcorrencia = linhaRegistro.identificacaoOcorrencia.valor;
				blnAchou = false;
				for (int j = _listaTotalizacaoOcorrencia.Count - 1; j >= 0; j--)
				{
					if (_listaTotalizacaoOcorrencia[j].descricao.Equals(strIdentificacaoOcorrencia))
					{
						blnAchou = true;
						itemTotalizacaoOcorrencia = _listaTotalizacaoOcorrencia[j];
						break;
					}
				}

				if (!blnAchou)
				{
					itemTotalizacaoOcorrencia = new TipoLinhaResumoRelArqRetorno(strIdentificacaoOcorrencia, 0, 0m);
					_listaTotalizacaoOcorrencia.Add(itemTotalizacaoOcorrencia);
				}

				itemTotalizacaoOcorrencia.qtdeTotal++;

				if (strIdentificacaoOcorrencia.Equals("06") || 
					strIdentificacaoOcorrencia.Equals("15") ||
					strIdentificacaoOcorrencia.Equals("16"))
				{
					itemTotalizacaoOcorrencia.vlTotal += Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor);
				}
				else
				{
					itemTotalizacaoOcorrencia.vlTotal += Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor);
				}

				#region [ Outras totalizações ]
				if (strIdentificacaoOcorrencia.Equals("02"))
				{
					vlAux = Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor);
					if (vlAux > 0)
					{
						itemTotalTarifasRegCobranca.qtdeTotal++;
						itemTotalTarifasRegCobranca.vlTotal += vlAux;
					}
				}
				else if (strIdentificacaoOcorrencia.Equals("28"))
				{
					if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "04"))
					{
						itemTotalTarifasProtesto.qtdeTotal++;
						vlAux = Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor);
						itemTotalTarifasProtesto.vlTotal += vlAux;
					}
					else if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "08"))
					{
						itemTotalTarifasCustasProtesto.qtdeTotal++;
						vlAux = Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor);
						itemTotalTarifasCustasProtesto.vlTotal += vlAux;
					}
				}
				else if (strIdentificacaoOcorrencia.Equals("06") ||
						 strIdentificacaoOcorrencia.Equals("15") ||
						 strIdentificacaoOcorrencia.Equals("17"))
				{
					if (linhaRegistro.bancoCobrador.valor.Equals("237"))
					{
						itemTotalTitulosPagosBradesco.qtdeTotal++;
						vlAux = Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor);
						itemTotalTitulosPagosBradesco.vlTotal += vlAux;
					}
					else
					{
						itemTotalTitulosPagosOutrosBancos.qtdeTotal++;
						vlAux = Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor);
						itemTotalTitulosPagosOutrosBancos.vlTotal += vlAux;
					}
				}
				#endregion
			}
			#endregion

			#region [ Prepara dados p/ impressão do resumo do arquivo de retorno ]
			_listaLinhasResumoArqRetorno = new List<TipoLinhaResumoRelArqRetorno>();
			_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
							"TOTAL EM COBRANÇA BANCÁRIA",
							(int)Global.converteInteiro(_linhaTrailler.qtdeTitulosEmCobranca.valor),
							Global.decodificaCampoMonetario(_linhaTrailler.valorTotalEmCobranca.valor)));
			_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
							"CONFIRMAÇÃO DE ENTRADAS",
							(int)Global.converteInteiro(_linhaTrailler.qtdeRegsOcorrencia02ConfirmacaoEntradas.valor),
							Global.decodificaCampoMonetario(_linhaTrailler.valorRegsOcorrencia02ConfirmacaoEntradas.valor)));
			_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
							"LIQUIDADO EM COBRANÇA",
							(int)Global.converteInteiro(_linhaTrailler.qtdeRegsOcorrencia06Liquidacao.valor),
							Global.decodificaCampoMonetario(_linhaTrailler.valorRegsOcorrencia06Liquidacao.valor)));
			_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
							"TÍTULOS BAIXADOS",
							(int)Global.converteInteiro(_linhaTrailler.qtdeRegsOcorrencia09e10TitulosBaixados.valor),
							Global.decodificaCampoMonetario(_linhaTrailler.valorRegsOcorrencia09e10TitulosBaixados.valor)));
			_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
							"VENCIMENTO ALTERADO",
							(int)Global.converteInteiro(_linhaTrailler.qtdeRegsOcorrencia14VenctoAlterado.valor),
							Global.decodificaCampoMonetario(_linhaTrailler.valorRegsOcorrencia14VenctoAlterado.valor)));
			_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
							"CONF. INSTRUÇÃO PROTESTO",
							(int)Global.converteInteiro(_linhaTrailler.qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto.valor),
							Global.decodificaCampoMonetario(_linhaTrailler.valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto.valor)));
			_listaLinhasResumoArqRetorno.Add(new TipoLinhaResumoRelArqRetorno("", 0, 0m));
			for (int i = 0; i < _listaTotalizacaoOcorrencia.Count; i++)
			{
				_listaLinhasResumoArqRetorno.Add(
					new TipoLinhaResumoRelArqRetorno(
						"(" + _listaTotalizacaoOcorrencia[i].descricao + ") " +
							Global.decodificaIdentificacaoOcorrencia(_listaTotalizacaoOcorrencia[i].descricao).ToUpper(),
						_listaTotalizacaoOcorrencia[i].qtdeTotal,
						_listaTotalizacaoOcorrencia[i].vlTotal)
					);
			}
			_listaLinhasResumoArqRetorno.Add(new TipoLinhaResumoRelArqRetorno("", 0, 0m));
			_listaLinhasResumoArqRetorno.Add(itemTotalTarifasRegCobranca);
			_listaLinhasResumoArqRetorno.Add(itemTotalTarifasProtesto);
			_listaLinhasResumoArqRetorno.Add(itemTotalTarifasCustasProtesto);
			_listaLinhasResumoArqRetorno.Add(new TipoLinhaResumoRelArqRetorno("", 0, 0m));
			_listaLinhasResumoArqRetorno.Add(itemTotalTitulosPagosBradesco);
			_listaLinhasResumoArqRetorno.Add(itemTotalTitulosPagosOutrosBancos);
			#endregion
		}
		#endregion

		#region [ executaPrintPageListagemTodasOcorrencias ]
		private void executaPrintPageListagemTodasOcorrencias(ref object sender, ref System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			float cx;
			float cy;
			String strTexto;
			String strIdBoletoItem;
			String[] vId;
			int intLinhasImpressasNestaPagina = 0;
			int idBoletoItem;
			LinhaRegistroTipo1ArquivoRetorno linhaRegistro = new LinhaRegistroTipo1ArquivoRetorno();
			decimal vlTitulo;
			decimal vlPago;
			decimal vlOscilacao;
			DsDataSource.DtbFinBoletoRow rowBoleto;
			#endregion

			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				e.Cancel = true;
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
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
				prnDocConsulta.DocumentName = "Listagem de todas as ocorrências do arquivo de retorno";
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
				wxOcorrencia = 7f;
				wxNossoNumero = 22f;
				wxNumeroTitulo = 19f;
				wxDtPagto = 16f;
				wxDtVencto = 16f;
				wxVlTitulo = 23f;
				wxVlPago = 23f;
				wxVlOscilacao = 23f;
				wxBcoCobr = 7f;
				wxAgCobr = 9f;
				wxMotivosOcorrencia = 24f;
				wxNomeSacado = larguraUtil
								- wxOcorrencia                         // A 1ª coluna não tem espaçamento
								- ESPACAMENTO_COLUNAS - wxNossoNumero
								- ESPACAMENTO_COLUNAS - wxNumeroTitulo
								- ESPACAMENTO_COLUNAS                  // Espaçamento da própria coluna "Nome Sacado"
								- ESPACAMENTO_COLUNAS - wxDtPagto
								- ESPACAMENTO_COLUNAS - wxDtVencto
								- ESPACAMENTO_COLUNAS - wxVlTitulo
								- ESPACAMENTO_COLUNAS - wxVlPago
								- ESPACAMENTO_COLUNAS - wxVlOscilacao
								- ESPACAMENTO_COLUNAS - wxBcoCobr
								- ESPACAMENTO_COLUNAS - wxAgCobr
								- ESPACAMENTO_COLUNAS - wxMotivosOcorrencia;
				ixOcorrencia = cxInicio;
				ixNossoNumero = ixOcorrencia + wxOcorrencia + ESPACAMENTO_COLUNAS;
				ixNumeroTitulo = ixNossoNumero + wxNossoNumero + ESPACAMENTO_COLUNAS;
				ixNomeSacado = ixNumeroTitulo + wxNumeroTitulo + ESPACAMENTO_COLUNAS;
				ixDtPagto = ixNomeSacado + wxNomeSacado + ESPACAMENTO_COLUNAS;
				ixDtVencto = ixDtPagto + wxDtPagto + ESPACAMENTO_COLUNAS;
				ixVlTitulo = ixDtVencto + wxDtVencto + ESPACAMENTO_COLUNAS;
				ixVlPago = ixVlTitulo + wxVlTitulo + ESPACAMENTO_COLUNAS;
				ixVlOscilacao = ixVlPago + wxVlPago + ESPACAMENTO_COLUNAS;
				ixBcoCobr = ixVlOscilacao + wxVlOscilacao + ESPACAMENTO_COLUNAS;
				ixAgCobr = ixBcoCobr + wxBcoCobr + ESPACAMENTO_COLUNAS;
				ixMotivosOcorrencia = ixAgCobr + wxAgCobr + ESPACAMENTO_COLUNAS;
				#endregion

				#region [ Layout das colunas do resumo ]
				ixResumoOcorrencia = cxInicio;
				for (int i = 0; i < _listaTotalizacaoOcorrencia.Count; i++)
				{
					strTexto = _listaTotalizacaoOcorrencia[i].descricao;
					strTexto += " - " + Global.decodificaIdentificacaoOcorrencia(strTexto);
					wxResumoOcorrencia = Math.Max(wxResumoOcorrencia, e.Graphics.MeasureString(strTexto, fonteListagem).Width);
				}
				wxResumoOcorrencia = Math.Max(wxResumoOcorrencia + 2f, 60f);
				ixResumoQtdeTitulos = ixResumoOcorrencia + wxResumoOcorrencia + ESPACAMENTO_COLUNAS;
				wxResumoQtdeTitulos = 25f;
				ixResumoVlTotal = ixResumoQtdeTitulos + wxResumoQtdeTitulos + ESPACAMENTO_COLUNAS;
				wxResumoVlTotal = 35f;
				#endregion
			}

			cy = cyInicio;

			#region [ Título ]
			strTexto = "LISTAGEM DE TODAS AS OCORRÊNCIAS DO ARQUIVO DE RETORNO";
			fonteAtual = fonteTitulo;
			cx = cxInicio + (larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			cy += 1f;
			#endregion

			#region [ Informações no cabeçalho ]

			#region [ Nome do arquivo ]
			strTexto = "Arquivo de retorno: " + txtArqRetorno.Text;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Empresa cedente / Carteira]
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

			#region [ Código da empresa / Data no banco e Data do crédito]
			strTexto = "Código da Empresa.: " + _boletoCedente.codigo_empresa + "  -  Nº aviso bancário: " + _linhaHeader.numAvisoBancario.valor;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil / 2);
			strTexto = "Data no Banco: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataGravacaoArquivo.valor)) + "  -  Data do Crédito: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataCredito.valor));
			fonteAtual = fonteListagem;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#endregion

			#region [ Imprime a seção da listagem dos registros? ]
			if (_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1))
			{
				cy += .5f;
				e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
				cy += .5f;

				#region [ Títulos da listagem ]
				cy += .5f;
				fonteAtual = fonteListagem;
				strTexto = "OCO";
				cx = ixOcorrencia + (wxOcorrencia - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "NOSSO NÚMERO";
				cx = ixNossoNumero;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "Nº TÍTULO";
				cx = ixNumeroTitulo;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "NOME SACADO";
				cx = ixNomeSacado;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "DT PAGTO";
				cx = ixDtPagto + (wxDtPagto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "DT VENCTO";
				cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "VL TÍTULO";
				cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "VL PAGO";
				cx = ixVlPago + wxVlPago - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "OSCILAÇÃO";
				cx = ixVlOscilacao + wxVlOscilacao - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "BCO";
				cx = ixBcoCobr + (wxBcoCobr - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "AGÊNC";
				cx = ixAgCobr + (wxAgCobr - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "MOTIVOS";
				cx = ixMotivosOcorrencia;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				cy += fonteAtual.GetHeight(e.Graphics);
				cy += .5f;
				#endregion

				cy += .5f;
				e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
				cy += .5f;

				#region [ Laço para listagem ]
				while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
					   (_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1)))
				{
					linhaRegistro.CarregaDados(_linhasArqRetorno[_intImpressaoArqRetornoIdxLinha]);

					#region [ Ocorrência ]
					strTexto = linhaRegistro.identificacaoOcorrencia.valor;
					cx = ixOcorrencia + (wxOcorrencia - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Nosso Número ]
					strTexto = Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor);
					cx = ixNossoNumero;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Número do título ]
					strTexto = linhaRegistro.numeroDocumento.valor;
					cx = ixNumeroTitulo;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Nome do sacado ]

					#region [ Possui nº controle do participante (t_FIN_BOLETO_ITEM.id)? ]
					idBoletoItem = 0;
					if (linhaRegistro.numControleParticipante.valor.Trim().Length > 0)
					{
						if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") != -1)
						{
							vId = linhaRegistro.numControleParticipante.valor.Split('=');
							strIdBoletoItem = vId[1];
							if (strIdBoletoItem != null)
							{
								if (strIdBoletoItem.Trim().Length > 0) idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
							}
						}
					}
					#endregion

					#region [ Pesquisa por Nosso Número / Número Documento ]
					if (idBoletoItem > 0)
					{
						rowBoleto = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					}
					else if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoleto = BoletoDAO.obtemRegistroPrincipalBoletoByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
					}
					else if (linhaRegistro.numeroDocumento.valor.Trim().Length > 0)
					{
						rowBoleto = BoletoDAO.obtemRegistroPrincipalBoletoByNumeroDocumento(_boletoCedente.id, linhaRegistro.numeroDocumento.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
					}
					else
					{
						rowBoleto = null;
					}
					#endregion

					strTexto = (rowBoleto != null ? rowBoleto.nome_sacado : " ");
					while ((e.Graphics.MeasureString(strTexto, fonteAtual).Width > wxNomeSacado) && (strTexto.Length > 0))
					{
						strTexto = strTexto.Substring(0, strTexto.Length - 1);
					}
					cx = ixNomeSacado;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Data do pagamento ]
					strTexto = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor));
					cx = ixDtPagto + (wxDtPagto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
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

					#region [ Valor pago ]
					vlPago = Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor);
					strTexto = Global.formataMoeda(vlPago);
					cx = ixVlPago + wxVlPago - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Valor Oscilação ]
					vlOscilacao = (vlPago != 0 ? vlPago - vlTitulo : 0);
					strTexto = Global.formataMoeda(vlOscilacao);
					cx = ixVlOscilacao + wxVlOscilacao - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Banco Cobrador ]
					strTexto = linhaRegistro.bancoCobrador.valor;
					cx = ixBcoCobr + (wxBcoCobr - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Agencia Cobradora ]
					strTexto = linhaRegistro.agenciaCobradora.valor;
					cx = ixAgCobr + (wxAgCobr - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					#region [ Motivos da ocorrência ]
					if (linhaRegistro.identificacaoOcorrencia.Equals("19"))
					{
						strTexto = linhaRegistro.motivoCodigoOcorrencia19.valor;
					}
					else
					{
						strTexto = linhaRegistro.motivosRejeicoes.valor;
						strTexto = strTexto.Insert(8, " ");
						strTexto = strTexto.Insert(6, " ");
						strTexto = strTexto.Insert(4, " ");
						strTexto = strTexto.Insert(2, " ");
					}
					cx = ixMotivosOcorrencia;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion

					cy += fonteAtual.GetHeight(e.Graphics);

					intLinhasImpressasNestaPagina++;
					_intImpressaoArqRetornoIdxLinha++;

					#region [ Na última linha não imprime o tracejado ]
					if (_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1))
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
			}
			#endregion

			#region [ Resumo do arquivo de retorno ]
			if (!(_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1)) &&
				 (_intImpressaoResumoArqRetornoIdxLinha < _listaLinhasResumoArqRetorno.Count))
			{
				if (intLinhasImpressasNestaPagina > 0) cy += 5f;
				
				if ((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 20))
				{
					#region [ Título: Resumo do arquivo de retorno ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;

					strTexto = "RESUMO DO ARQUIVO DE RETORNO";
					cx = cxInicio + (larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					cy += fonteAtual.GetHeight(e.Graphics);
					cy += .5f;

					cy += .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion

					#region [ Títulos do resumo ]
					cy += .5f;
					fonteAtual = fonteListagem;
					strTexto = "DESCRIÇÃO / OCORRÊNCIA";
					cx = ixResumoOcorrencia;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = "Nº TÍTULOS";
					cx = ixResumoQtdeTitulos + wxResumoQtdeTitulos - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = "VALORES TOTAIS";
					cx = ixResumoVlTotal + wxResumoVlTotal - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					cy += fonteAtual.GetHeight(e.Graphics);
					cy += .5f;

					cy += .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, ixResumoVlTotal + wxResumoVlTotal, cy);
					cy += .5f;
					#endregion

					#region [ Laço para listagem ]
					while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
						   (_intImpressaoResumoArqRetornoIdxLinha < _listaLinhasResumoArqRetorno.Count))
					{
						#region [ Descrição/Ocorrência ]
						strTexto = _listaLinhasResumoArqRetorno[_intImpressaoResumoArqRetornoIdxLinha].descricao;
						cx = cxInicio;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						#endregion

						#region [ Nº Títulos ]
						if (_listaLinhasResumoArqRetorno[_intImpressaoResumoArqRetornoIdxLinha].descricao.Trim().Length > 0)
						{
							strTexto = Global.formataInteiro(_listaLinhasResumoArqRetorno[_intImpressaoResumoArqRetornoIdxLinha].qtdeTotal);
							cx = ixResumoQtdeTitulos + wxResumoQtdeTitulos - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
							e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						}
						#endregion

						#region [ Valor Total ]
						if (_listaLinhasResumoArqRetorno[_intImpressaoResumoArqRetornoIdxLinha].descricao.Trim().Length > 0)
						{
							strTexto = Global.formataMoeda(_listaLinhasResumoArqRetorno[_intImpressaoResumoArqRetornoIdxLinha].vlTotal);
							cx = ixResumoVlTotal + wxResumoVlTotal - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
							e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						}
						#endregion

						cy += fonteAtual.GetHeight(e.Graphics);

						intLinhasImpressasNestaPagina++;
						_intImpressaoResumoArqRetornoIdxLinha++;
					}
					#endregion
				}
			}
			#endregion

			#region [ Imprime nº página ]
			strTexto = "Página: " + _intImpressaoNumPagina.ToString().PadLeft(2, ' ');
			fonteAtual = fonteNumPagina;
			cy = cyRodapeNumPagina;
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Tem mais páginas para imprimir? ]
			if ((_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1)) ||
			    (_intImpressaoResumoArqRetornoIdxLinha < _listaLinhasResumoArqRetorno.Count))
			{
				e.HasMorePages = true;
			}
			else
			{
				e.HasMorePages = false;
			}
			#endregion
		}
		#endregion

		#region [ executaBeginPrintListagemRegistrosRejeitados ]
		private void executaBeginPrintListagemRegistrosRejeitados(ref object sender, ref System.Drawing.Printing.PrintEventArgs e)
		{
			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				e.Cancel = true;
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
			{
				e.Cancel = true;
				return;
			}
			#endregion

			prnDocConsulta.DocumentName = "Listagem de registros rejeitados no processamento do retorno";

			_intImpressaoArqRetornoIdxLinha = 1;  // Índice zero é o header e o último índice é o trailler
			_intImpressaoNumPagina = 0;
			_strImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
			_intQtdeRejeicoes = 0;
			_vlTotalRejeicoes = 0;

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

		#region [ executaPrintPageListagemRegistrosRejeitados ]
		private void executaPrintPageListagemRegistrosRejeitados(ref object sender, ref System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			float cx;
			float cy;
			float hMax;
			decimal vlTitulo;
			int idBoletoItem;
			bool blnImprimeRegistro;
			String strIdentificacaoOcorrencia;
			String strIdBoletoItem;
			String[] vId;
			RectangleF r;
			String strTexto;
			String strDescricaoMotivoOcorrencia;
			int intLinhasImpressasNestaPagina = 0;
			LinhaRegistroTipo1ArquivoRetorno linhaRegistro = new LinhaRegistroTipo1ArquivoRetorno();
			DsDataSource.DtbFinBoletoRow rowBoleto;
			List<Global.TipoDescricaoMotivoOcorrencia> listaMotivoOcorrencia;
			#endregion

			#region [ Consistência ]
			if (_linhasArqRetorno == null)
			{
				e.Cancel = true;
				return;
			}
			if (_linhasArqRetorno.Length <= 2)
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
				prnDocConsulta.DocumentName = "Listagem de registros rejeitados no processamento do retorno";
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
				wxOcorrencia = 7f;
				wxIdentifCedente = 40f;
				wxNossoNumero = 22f;
				wxNumeroTitulo = 19f;
				wxDtVencto = 16f;
				wxVlTitulo = 23f;
				wxMotivosOcorrencia = 75f;
				wxNomeSacado = larguraUtil
								- wxOcorrencia                         // A 1ª coluna não tem espaçamento
								- ESPACAMENTO_COLUNAS - wxIdentifCedente
								- ESPACAMENTO_COLUNAS - wxNossoNumero
								- ESPACAMENTO_COLUNAS - wxNumeroTitulo
								- ESPACAMENTO_COLUNAS - wxDtVencto
								- ESPACAMENTO_COLUNAS - wxVlTitulo
								- ESPACAMENTO_COLUNAS                  // Espaçamento da própria coluna "Nome Sacado"
								- ESPACAMENTO_COLUNAS - wxMotivosOcorrencia;
				ixOcorrencia = cxInicio;
				ixIdentifCedente = ixOcorrencia + wxOcorrencia + ESPACAMENTO_COLUNAS;
				ixNossoNumero = ixIdentifCedente + wxIdentifCedente + ESPACAMENTO_COLUNAS;
				ixNumeroTitulo = ixNossoNumero + wxNossoNumero + ESPACAMENTO_COLUNAS;
				ixDtVencto = ixNumeroTitulo + wxNumeroTitulo + ESPACAMENTO_COLUNAS;
				ixVlTitulo = ixDtVencto + wxDtVencto + ESPACAMENTO_COLUNAS;
				ixNomeSacado = ixVlTitulo + wxVlTitulo + ESPACAMENTO_COLUNAS;
				ixMotivosOcorrencia = ixNomeSacado + wxNomeSacado + ESPACAMENTO_COLUNAS;
				#endregion
			}

			cy = cyInicio;

			#region [ Título ]
			strTexto = "REGISTROS REJEITADOS NO PROCESSAMENTO DO RETORNO";
			fonteAtual = fonteTitulo;
			cx = cxInicio + (larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			cy += 1f;
			#endregion

			#region [ Informações no cabeçalho ]

			#region [ Nome do arquivo ]
			strTexto = "Arquivo de retorno: " + txtArqRetorno.Text;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Empresa cedente / Carteira]
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

			#region [ Código da empresa / Data no banco e Data do crédito]
			strTexto = "Código da Empresa.: " + _boletoCedente.codigo_empresa + "  -  Nº aviso bancário: " + _linhaHeader.numAvisoBancario.valor;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil / 2);
			strTexto = "Data no Banco: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataGravacaoArquivo.valor)) + "  -  Data do Crédito: " + Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_linhaHeader.dataCredito.valor));
			fonteAtual = fonteListagem;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#endregion

			#region [ Imprime a seção da listagem dos registros? ]
			if (_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1))
			{
				cy += .5f;
				e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
				cy += .5f;

				#region [ Títulos da listagem ]
				cy += .5f;
				fonteAtual = fonteListagem;
				strTexto = "OCO";
				cx = ixOcorrencia + (wxOcorrencia - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "IDEN. NO BANCO";
				cx = ixIdentifCedente;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "NOSSO NÚMERO";
				cx = ixNossoNumero;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "Nº TÍTULO";
				cx = ixNumeroTitulo;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "DT VENCTO";
				cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "VL TÍTULO";
				cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "NOME SACADO";
				cx = ixNomeSacado;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = "MOTIVO";
				cx = ixMotivosOcorrencia;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				cy += fonteAtual.GetHeight(e.Graphics);
				cy += .5f;
				#endregion

				cy += .5f;
				e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
				cy += .5f;

				#region [ Laço para listagem ]
				while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
					   (_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1)))
				{
					linhaRegistro.CarregaDados(_linhasArqRetorno[_intImpressaoArqRetornoIdxLinha]);

					hMax = fonteAtual.GetHeight(e.Graphics);

					#region [ Verifica se é uma rejeição e se deve ser impressa ]
					strIdentificacaoOcorrencia = linhaRegistro.identificacaoOcorrencia.valor;
					blnImprimeRegistro = true;
					if (strIdentificacaoOcorrencia.Equals("02")
							||
						strIdentificacaoOcorrencia.Equals("06")
							||
						strIdentificacaoOcorrencia.Equals("12")
							||
						strIdentificacaoOcorrencia.Equals("13")
							||
						strIdentificacaoOcorrencia.Equals("14")
							||
						strIdentificacaoOcorrencia.Equals("15")
							||
						strIdentificacaoOcorrencia.Equals("16")
						)
					{
						blnImprimeRegistro = false;
					}
					#endregion

					if (blnImprimeRegistro)
					{
						#region [ Imprime tracejado para separar da linha anterior? ]
						if (intLinhasImpressasNestaPagina > 0)
						{
							#region [ Traço pontilhado ]
							cy += .5f;
							e.Graphics.DrawLine(penTracoPontilhado, cxInicio, cy, cxFim, cy);
							cy += .5f;
							#endregion
						}
						#endregion

						#region [ Ocorrência ]
						strTexto = linhaRegistro.identificacaoOcorrencia.valor;
						cx = ixOcorrencia + (wxOcorrencia - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						#endregion

						#region [ Iden. no Banco ]
						strTexto = linhaRegistro.identifCedenteCarteira.valor + " / " + linhaRegistro.identifCedenteAgencia + " / " + linhaRegistro.identifCedenteCtaCorrente + '-' + linhaRegistro.identifCedenteDigitoCtaCorrente;
						cx = ixIdentifCedente;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						#endregion

						#region [ Nosso Número ]
						strTexto = Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor);
						cx = ixNossoNumero;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						#endregion

						#region [ Número do título ]
						strTexto = linhaRegistro.numeroDocumento.valor;
						cx = ixNumeroTitulo;
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

						#region [ Nome do sacado ]

						#region [ Possui nº controle do participante (t_FIN_BOLETO_ITEM.id)? ]
						idBoletoItem = 0;
						if (linhaRegistro.numControleParticipante.valor.Trim().Length > 0)
						{
							if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") != -1)
							{
								vId = linhaRegistro.numControleParticipante.valor.Split('=');
								strIdBoletoItem = vId[1];
								if (strIdBoletoItem != null)
								{
									if (strIdBoletoItem.Trim().Length > 0) idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
								}
							}
						}
						#endregion

						#region [ Pesquisa por Nosso Número / Número Documento ]
						if (idBoletoItem > 0)
						{
							rowBoleto = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
						}
						else if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
						{
							rowBoleto = BoletoDAO.obtemRegistroPrincipalBoletoByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
						}
						else if (linhaRegistro.numeroDocumento.valor.Trim().Length > 0)
						{
							rowBoleto = BoletoDAO.obtemRegistroPrincipalBoletoByNumeroDocumento(_boletoCedente.id, linhaRegistro.numeroDocumento.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
						}
						else
						{
							rowBoleto = null;
						}
						#endregion

						strTexto = (rowBoleto != null ? rowBoleto.nome_sacado : " ");
						while ((e.Graphics.MeasureString(strTexto, fonteAtual).Width > wxNomeSacado) && (strTexto.Length > 0))
						{
							strTexto = strTexto.Substring(0, strTexto.Length - 1);
						}
						cx = ixNomeSacado;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						#endregion

						#region [ Motivo ]
						strDescricaoMotivoOcorrencia = Global.decodificaIdentificacaoOcorrencia(linhaRegistro.identificacaoOcorrencia.valor);
						if (linhaRegistro.identificacaoOcorrencia.valor.Equals("19"))
						{
							if (strDescricaoMotivoOcorrencia.Length > 0) strDescricaoMotivoOcorrencia += "\n";
							strDescricaoMotivoOcorrencia = linhaRegistro.motivoCodigoOcorrencia19 + " - " + Global.decodificaMotivoOcorrencia19(linhaRegistro.motivoCodigoOcorrencia19.valor);
						}
						else
						{
							listaMotivoOcorrencia = Global.decodificaMotivoOcorrencia(linhaRegistro.identificacaoOcorrencia.valor, linhaRegistro.motivosRejeicoes.valor);
							for (int j = 0; j < listaMotivoOcorrencia.Count; j++)
							{
								if (listaMotivoOcorrencia[j].descricaoMotivoOcorrencia.Length > 0)
								{
									if (strDescricaoMotivoOcorrencia.Length > 0) strDescricaoMotivoOcorrencia += "\n";
									strDescricaoMotivoOcorrencia += listaMotivoOcorrencia[j].motivoOcorrencia + " - " + listaMotivoOcorrencia[j].descricaoMotivoOcorrencia;
								}
							}
						}

						strTexto = strDescricaoMotivoOcorrencia;
						cx = ixMotivosOcorrencia;
						r = new RectangleF(ixMotivosOcorrencia, cy, wxMotivosOcorrencia, 20);
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
						hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxMotivosOcorrencia).Height);
						#endregion

						cy += hMax;

						intLinhasImpressasNestaPagina++;
						_intQtdeRejeicoes++;
						_vlTotalRejeicoes += vlTitulo;
					}

					_intImpressaoArqRetornoIdxLinha++;
				}
				#endregion
			}
			#endregion

			#region [ Tem mais páginas para imprimir? ]
			if ((_intImpressaoArqRetornoIdxLinha < (_linhasArqRetorno.Length - 1)))
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
					strTexto = "Número de títulos lidos: " + Global.formataInteiro(_intQtdeRejeicoes);
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = "Valor total títulos em " + Global.Cte.Etc.SIMBOLO_MONETARIO + ": " + Global.formataMoeda(_vlTotalRejeicoes).PadLeft(12, ' ');
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

	#region [ Classe: TipoLinhaResumoRelArqRetorno ]
	class TipoLinhaResumoRelArqRetorno
	{
		#region [ Getters ]
		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
		}

		private int _qtdeTotal = 0;
		public int qtdeTotal
		{
			get { return _qtdeTotal; }
			set { _qtdeTotal = value; }
		}

		private decimal _vlTotal = 0m;
		public decimal vlTotal
		{
			get { return _vlTotal; }
			set { _vlTotal = value; }
		}
		#endregion

		#region [ Construtor ]
		public TipoLinhaResumoRelArqRetorno(String descricao, int qtdeTotal, decimal vlTotal)
		{
			_descricao = descricao;
			_qtdeTotal = qtdeTotal;
			_vlTotal = vlTotal;
		}
		#endregion
	}
	#endregion
}
