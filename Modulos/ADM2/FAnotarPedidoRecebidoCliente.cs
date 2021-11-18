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
	public partial class FAnotarPedidoRecebidoCliente : FModelo
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

		private string _tituloBoxDisplayInformativo = "Mensagens Informativas";
		private int _qtdeMsgDisplayInformativo = 0;
		private string _tituloBoxDisplayErro = "Mensagens de Erro";
		private int _qtdeMsgDisplayErro = 0;
		List<RastreioPedidoRecebidoCliente> _vRastreio;

		private string _tituloPainelOriginal;
		Parametro paramProcDtPrevEntregaTransp;
		#endregion

		#region [ Constantes ]
		// Obs: a coluna 'ColVisibleOrdenacaoPadrao' é a coluna visível usada p/ poder ser clicada e fazer a ordenação conforme o padrão inicial, sendo que as células dessa coluna ficam vazias.
		// E a coluna 'ColHiddenValorOrdenacaoPadrao' é a coluna invisível que possui os dados usados p/ a ordenação padrão.
		const string GRID_COL_VISIBLE_ORDENACAO_PADRAO = "ColVisibleOrdenacaoPadrao";
		const string GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO = "ColHiddenValorOrdenacaoPadrao";
		const string GRID_COL_NF = "NF";
		const string GRID_COL_HIDDEN_NF = "ColHiddenNF";
		const string GRID_COL_DESTINATARIO = "Destinatario";
		const string GRID_COL_DESTINO = "Destino";
		const string GRID_COL_SITUACAO = "Situacao";
		const string GRID_COL_DETALHE = "Detalhe";
		const string GRID_COL_DATA_ENTREGA = "DataEntrega";
		const string GRID_COL_HIDDEN_DATA_ENTREGA = "ColHiddenDataEntrega";
		const string GRID_COL_PREVISAO_ENTREGA = "PrevisaoEntrega";
		const string GRID_COL_HIDDEN_PREVISAO_ENTREGA = "ColHiddenPrevisaoEntrega";
		const string GRID_COL_HIDDEN_GUID = "ColHiddenGuid";
		const string GRID_COL_STATUS = "Status";
		const string GRID_COL_MENSAGEM = "Mensagem";
		const int MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS = 500;
		#endregion

		#region [ Construtor ]
		public FAnotarPedidoRecebidoCliente()
		{
			InitializeComponent();

			paramProcDtPrevEntregaTransp = FMain.contextoBD.AmbienteBase.parametroDAO.getParametro(Global.Cte.ADM2.ID_T_PARAMETRO.ADM2_PROCESSAR_DATA_PREVISAO_ENTREGA_TRANSPORTADORA);
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

		#region [ pathArquivoRastreioValorDefault ]
		private String pathArquivoRastreioValorDefault()
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
			if (Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio))
				{
					strResp = Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNameArquivoRastreioValorDefault ]
		private String fileNameArquivoRastreioValorDefault()
		{
			String strResp = "";

			if ((Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio ?? "").Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio + "\\" + Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio))
				{
					strResp = Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio;
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

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			limpaCamposMensagem();
			lblTotalRegistros.Text = "";
			lblQtdeRegErro.Text = "";
			lblQtdeRegApto.Text = "";
			lblQtdeAtualizFalha.Text = "";
			lblQtdeAtualizSucesso.Text = "";
			grid.Rows.Clear();
			for (int i = 0; i < grid.Columns.Count; i++)
			{
				grid.Columns[i].HeaderCell.SortGlyphDirection = SortOrder.None;
			}
			_vRastreio = new List<RastreioPedidoRecebidoCliente>();
		}
		#endregion

		#region [ trataBotaoSelecionaArquivoRastreio ]
		private void trataBotaoSelecionaArquivoRastreio()
		{
			#region [ Declarações ]
			DialogResult dr;
			#endregion

			try
			{
				openFileDialogCtrl.InitialDirectory = pathArquivoRastreioValorDefault();
				openFileDialogCtrl.FileName = "";
				dr = openFileDialogCtrl.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				//if ((openFileDialogCtrl.FileName.Length > 0) && (txtArquivoRastreio.Text.Length > 0))
				//{
				//    if (openFileDialogCtrl.FileName.ToUpper().Equals(txtArquivoRastreio.Text.ToUpper())) return;
				//}
				#endregion

				#region [ Limpa campos ]
				limpaCampos();
				#endregion

				txtArquivoRastreio.Text = openFileDialogCtrl.FileName;
				Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio = Path.GetDirectoryName(openFileDialogCtrl.FileName);
				Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio = Path.GetFileName(openFileDialogCtrl.FileName);

				carregaDadosArquivoRastreio();
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

		#region [ trataBotaoConfirma ]
		private void trataBotaoConfirma()
		{
			executaAnotaPedidoRecebidoCliente();
		}
		#endregion

		#region [ carregaDadosArquivoRastreio ]
		private void carregaDadosArquivoRastreio()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FAnotaPedidoRecebidoCliente.carregaDadosArquivoRastreio";
			int qtdeLinhaDadosArquivo = 0;
			int percProgresso;
			int percProgressoAnterior;
			int qtdeRegErro = 0;
			int qtdeRegApto = 0;
			long lngMiliSegundosDecorridos;
			bool blnCnpjCpfOk;
			bool blnHaLinhasStatusDesconhecido = false;
			bool blnDevolucao;
			string MARGEM_MSG_NIVEL_2 = new string(' ', 8);
			string sYYYYMMDDBrancos;
			string sOrdenacao;
			string sPedidos;
			string sNFBrancos;
			string sNF;
			string sDatePart;
			string sTimePart;
			string sDataHora;
			string sDia;
			string sMes;
			string sAno;
			string sHora;
			string sMinuto;
			string sSegundo;
			string strNomeArquivo;
			string strMsg;
			string strMsgErro;
			string strMsgProgresso;
			string linhaHeader;
			string sCpfLeftPaddedZero;
			string sCpfRightPaddedZero;
			string sOcorrenciaSituacao;
			string sOcorrenciaDetalhe;
			string[] linhasCSV;
			string[] camposHeader;
			string[] camposCSV;
			string[] v;
			StringBuilder sbErro;
			HeaderRastreioPedidoRecebidoCliente header = new HeaderRastreioPedidoRecebidoCliente();
			RastreioPedidoRecebidoCliente rastreio;
			List<RastreioPedidoRecebidoCliente> vRastreioOrdenado;
			DateTime dtHrUltProgresso;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			Pedido pedido;
			List<Pedido> listaPedidos;
			NFeImagem nFeImagem;
			List<NFeImagem> listaNFeImagem;
			#endregion

			try
			{
				#region [ Obtém o nome do arquivo ]
				strNomeArquivo = txtArquivoRastreio.Text.Trim();
				#endregion

				#region [ Consistências ]
				if (strNomeArquivo.Length == 0)
				{
					strMsgErro = "É necessário selecionar o arquivo com os dados de rastreio!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if (!File.Exists(strNomeArquivo))
				{
					strMsgErro = "O arquivo NÃO existe!\r\n" + strNomeArquivo;
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if (Global.IsFileLocked(strNomeArquivo))
				{
					strMsgErro = "O arquivo '" + Path.GetFileName(strNomeArquivo) + "' está aberto e em uso!\r\nNão é possível prosseguir com o processamento!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Inicialização do processamento ]
				dtInicioProcessamento = DateTime.Now;
				strMsg = "Início do processamento\r\n" +
						MARGEM_MSG_NIVEL_2 + "Arquivo: " + strNomeArquivo;
				adicionaDisplay(strMsg);
				#endregion

				#region [ Carrega dados do arquivo de rastreio ]
				try
				{
					#region [ Lê dados do arquivo ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Lendo dados do arquivo de rastreio");
					linhasCSV = File.ReadAllLines(strNomeArquivo, encode);
					adicionaDisplay("Registros para processar: " + Global.formataInteiro(linhasCSV.Length - 1));
					#endregion

					#region [ Verifica linha com títulos ]
					linhaHeader = linhasCSV[0];
					camposHeader = linhaHeader.Split(';');
					for (int i = 0; i < camposHeader.Length; i++)
					{
						foreach (var item in header.listaCamposHeader)
						{
							if (camposHeader[i].Equals(item.tituloColuna))
							{
								item.indexColuna = i;
								break;
							}
						}
					}

					sbErro = new StringBuilder("");
					foreach (var item in header.listaCamposHeader)
					{
						if (item.indexColuna == null)
						{
							sbErro.AppendLine("Não foi encontrada a coluna '" + item.tituloColuna + "'!");
						}
					}

					if (sbErro.Length > 0)
					{
						strMsgErro = "Falha ao analisar o header do arquivo '" + Path.GetFileName(strNomeArquivo) + "'\r\n" + sbErro.ToString() + "\r\nNão é possível prosseguir com o processamento!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#region [ Verifica se possui linha de dados ]
					if (linhasCSV.Length <= 1)
					{
						strMsgErro = "Arquivo '" + Path.GetFileName(strNomeArquivo) + "' não possui dados!\r\nNão é possível prosseguir com o processamento!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#region [ Carrega dados em uma lista ]
					percProgressoAnterior = -1;
					dtHrUltProgresso = DateTime.MinValue;
					// Ignora a primeira linha que é a do header
					for (int i = 1; i < linhasCSV.Length; i++)
					{
						if (linhasCSV[i].Trim().Length == 0) continue;

						lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
						percProgresso = 100 * i / (linhasCSV.Length - 1);
						if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
						{
							strMsgProgresso = "Analisando linhas do arquivo: " + i.ToString() + " / " + (linhasCSV.Length - 1).ToString() + "   (" + percProgresso.ToString() + "%)";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							percProgressoAnterior = percProgresso;
							Application.DoEvents();
							// Atualiza o horário após o DoEvents() para que o intervalo entre os DoEvents() não seja inferior ao mínimo definido
							dtHrUltProgresso = DateTime.Now;
						}

						qtdeLinhaDadosArquivo++;
						camposCSV = linhasCSV[i].Split(';');

						rastreio = new RastreioPedidoRecebidoCliente();

						#region [ Carrega dados do arquivo (da forma como vieram) ]
						rastreio.dadosRaw.CnpjCpfRemetente = camposCSV[(int)header.CnpjCpfRemetente.indexColuna].Trim();
						rastreio.dadosRaw.Remetente = camposCSV[(int)header.Remetente.indexColuna].Trim();
						rastreio.dadosRaw.CnpjCpfDestinatario = camposCSV[(int)header.CnpjCpfDestinatario.indexColuna].Trim();
						rastreio.dadosRaw.Destinatario = camposCSV[(int)header.Destinatario.indexColuna].Trim();
						rastreio.dadosRaw.CTRC = camposCSV[(int)header.CTRC.indexColuna].Trim();
						rastreio.dadosRaw.NF = camposCSV[(int)header.NF.indexColuna].Trim();
						rastreio.dadosRaw.NroPedido = camposCSV[(int)header.NroPedido.indexColuna].Trim();
						rastreio.dadosRaw.DataInclusao = camposCSV[(int)header.DataInclusao.indexColuna].Trim();
						rastreio.dadosRaw.CidadeDestino = camposCSV[(int)header.CidadeDestino.indexColuna].Trim();
						rastreio.dadosRaw.UfDestino = camposCSV[(int)header.UfDestino.indexColuna].Trim();
						rastreio.dadosRaw.Unidade = camposCSV[(int)header.Unidade.indexColuna].Trim();
						rastreio.dadosRaw.DataHoraOcorrencia = camposCSV[(int)header.DataHoraOcorrencia.indexColuna].Trim();
						rastreio.dadosRaw.Situacao = camposCSV[(int)header.Situacao.indexColuna].Trim();
						rastreio.dadosRaw.Empresa = camposCSV[(int)header.Empresa.indexColuna].Trim();
						rastreio.dadosRaw.Detalhe = camposCSV[(int)header.Detalhe.indexColuna].Trim();
						rastreio.dadosRaw.DataEntrega = camposCSV[(int)header.DataEntrega.indexColuna].Trim();
						rastreio.dadosRaw.PrevisaoEntrega = camposCSV[(int)header.PrevisaoEntrega.indexColuna].Trim();
						#endregion

						#region [ Normaliza os dados ]
						rastreio.dadosNormalizado.CnpjCpfRemetente = Global.digitos(rastreio.dadosRaw.CnpjCpfRemetente);
						rastreio.dadosNormalizado.Remetente = Texto.iniciaisEmMaiusculas(rastreio.dadosRaw.Remetente.Trim());
						rastreio.dadosNormalizado.CnpjCpfDestinatario = Global.digitos(rastreio.dadosRaw.CnpjCpfDestinatario);
						rastreio.dadosNormalizado.Destinatario = Texto.iniciaisEmMaiusculas(rastreio.dadosRaw.Destinatario.Trim());
						rastreio.dadosNormalizado.CTRC = rastreio.dadosRaw.CTRC.Trim();

						#region [ NF ]
						if (rastreio.dadosRaw.NF.Trim().Length > 0)
						{
							v = rastreio.dadosRaw.NF.Split(' ');
							if (v.Length >= 2)
							{
								// Tratamento somente se a NF estiver no formato: [Nº Série] [Espaço Branco] [Nº NF]
								rastreio.dadosNormalizado.SerieNF = v[0].Trim();
								rastreio.dadosNormalizado.numSerieNF = (int)Global.converteInteiro(rastreio.dadosNormalizado.SerieNF);
								rastreio.dadosNormalizado.NF = v[1].Trim();
								rastreio.dadosNormalizado.numNF = (int)Global.converteInteiro(rastreio.dadosNormalizado.NF);
							}
						}
						#endregion

						rastreio.dadosNormalizado.NroPedido = rastreio.dadosRaw.NroPedido.Trim();

						#region [ Data Inclusão ]
						if (rastreio.dadosRaw.DataInclusao.Trim().Length > 0)
						{
							v = rastreio.dadosRaw.DataInclusao.Split('/');
							if (v.Length >= 3)
							{
								sDia = v[0].Trim();
								sMes = v[1].Trim();
								sAno = v[2].Trim();
								if (sDia.Length == 1) sDia = '0' + sDia;
								if (sMes.Length == 1) sMes = '0' + sMes;
								if (sAno.Length == 2) sAno = "20" + sAno;
								rastreio.dadosNormalizado.DataInclusao = sDia + '/' + sMes + '/' + sAno;
								rastreio.dadosNormalizado.dtDataInclusao = Global.converteDdMmYyyyParaDateTime(rastreio.dadosNormalizado.DataInclusao);
							}
						}
						#endregion

						rastreio.dadosNormalizado.CidadeDestino = Texto.iniciaisEmMaiusculas(rastreio.dadosRaw.CidadeDestino.Trim());
						rastreio.dadosNormalizado.UfDestino = rastreio.dadosRaw.UfDestino.Trim().ToUpper();
						rastreio.dadosNormalizado.Unidade = Texto.iniciaisEmMaiusculas(rastreio.dadosRaw.Unidade.Trim());

						#region [ Data/hora da ocorrência ]
						if (rastreio.dadosRaw.DataHoraOcorrencia.Trim().Length > 0)
						{
							sDia = "";
							sMes = "";
							sAno = "";
							sHora = "";
							sMinuto = "";
							sSegundo = "";
							sDatePart = "";
							sTimePart = "";
							sDataHora = "";
							// Separa data da hora
							if (rastreio.dadosRaw.DataHoraOcorrencia.Contains(' '))
							{
								v = rastreio.dadosRaw.DataHoraOcorrencia.Split(' ');
								if (v.Length >= 2)
								{
									sDatePart = v[0].Trim();
									sTimePart = v[1].Trim();
								}
							}
							else if (rastreio.dadosRaw.DataHoraOcorrencia.Contains('/'))
							{
								sDatePart = rastreio.dadosRaw.DataHoraOcorrencia.Trim();
							}
							else if (rastreio.dadosRaw.DataHoraOcorrencia.Contains(':'))
							{
								sTimePart = rastreio.dadosRaw.DataHoraOcorrencia.Trim();
							}

							if (sDatePart.Length > 0)
							{
								v = sDatePart.Split('/');
								if (v.Length >= 3)
								{
									sDia = v[0].Trim();
									sMes = v[1].Trim();
									sAno = v[2].Trim();
									if (sDia.Length == 1) sDia = '0' + sDia;
									if (sMes.Length == 1) sMes = '0' + sMes;
									if (sAno.Length == 2) sAno = "20" + sAno;
								}
							}

							if (sTimePart.Length > 0)
							{
								v = sTimePart.Split(':');
								if (v.Length >= 2)
								{
									sHora = v[0].Trim();
									sMinuto = v[1].Trim();
									if (v.Length >= 3) sSegundo = v[2].Trim();
									if (sHora.Length == 1) sHora = '0' + sHora;
									if (sMinuto.Length == 1) sMinuto = '0' + sMinuto;
									while (sSegundo.Length < 2)
									{
										sSegundo = '0' + sSegundo;
									}
								}
							}

							if (sAno.Length > 0)
							{
								rastreio.dadosNormalizado.DataHoraOcorrencia = sDia + '/' + sMes + '/' + sAno;
								sDataHora = sAno + '-' + sMes + '-' + sDia;

								if (sHora.Length > 0)
								{
									if (rastreio.dadosNormalizado.DataHoraOcorrencia.Length > 0) rastreio.dadosNormalizado.DataHoraOcorrencia += ' ';
									rastreio.dadosNormalizado.DataHoraOcorrencia += sHora + ':' + sMinuto + ':' + sSegundo;
									if (sDataHora.Length > 0) sDataHora += ' ';
									sDataHora += sHora + ':' + sMinuto + ':' + sSegundo;
								}

								rastreio.dadosNormalizado.dtDataHoraOcorrencia = Global.converteYyyyMmDdHhMmSsParaDateTime(sDataHora);
							}
						}
						#endregion

						rastreio.dadosNormalizado.Situacao = Texto.iniciaisEmMaiusculas(rastreio.dadosRaw.Situacao.Trim());
						rastreio.dadosNormalizado.Empresa = rastreio.dadosRaw.Empresa.Trim();
						rastreio.dadosNormalizado.Detalhe = Texto.iniciaisEmMaiusculas(rastreio.dadosRaw.Detalhe.Trim());

						#region [ Data Entrega ]
						if (rastreio.dadosRaw.DataEntrega.Trim().Length > 0)
						{
							v = rastreio.dadosRaw.DataEntrega.Split('/');
							if (v.Length >= 3)
							{
								sDia = v[0].Trim();
								sMes = v[1].Trim();
								sAno = v[2].Trim();
								if (sDia.Length == 1) sDia = '0' + sDia;
								if (sMes.Length == 1) sMes = '0' + sMes;
								if (sAno.Length == 2) sAno = "20" + sAno;
								rastreio.dadosNormalizado.DataEntrega = sDia + '/' + sMes + '/' + sAno;
								rastreio.dadosNormalizado.dtDataEntrega = Global.converteDdMmYyyyParaDateTime(rastreio.dadosNormalizado.DataEntrega);
							}
						}
						#endregion

						#region [ Previsão de Entrega ]
						if (rastreio.dadosRaw.PrevisaoEntrega.Trim().Length > 0)
						{
							v = rastreio.dadosRaw.PrevisaoEntrega.Split('/');
							if (v.Length >= 3)
							{
								sDia = v[0].Trim();
								sMes = v[1].Trim();
								sAno = v[2].Trim();
								if (sDia.Length == 1) sDia = '0' + sDia;
								if (sMes.Length == 1) sMes = '0' + sMes;
								if (sAno.Length == 2) sAno = "20" + sAno;
								rastreio.dadosNormalizado.PrevisaoEntrega = sDia + '/' + sMes + '/' + sAno;
								rastreio.dadosNormalizado.dtPrevisaoEntrega = Global.converteDdMmYyyyParaDateTime(rastreio.dadosNormalizado.PrevisaoEntrega);
							}
						}
						#endregion

						#endregion

						_vRastreio.Add(rastreio);
					}
					#endregion
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return;
				}
				#endregion

				#region [ Gera um identificador único para cada linha ]
				for (int iv = 0; iv < _vRastreio.Count; iv++)
				{
					_vRastreio[iv].processo.Guid = Guid.NewGuid().ToString();
				}
				#endregion

				#region [ Analisa as ocorrências que podem ser de devolução ]
				strMsgProgresso = "Analisando existência de registros de devolução";
				info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
				try
				{
					for (int iv = 0; iv < _vRastreio.Count; iv++)
					{
						if (paramProcDtPrevEntregaTransp.campo_inteiro == 0)
						{
							if (_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length == 0) continue;
						}
						else
						{
							if ((_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length == 0) && (_vRastreio[iv].dadosNormalizado.PrevisaoEntrega.Trim().Length == 0)) continue;
						}

						blnDevolucao = false;
						sOcorrenciaSituacao = Global.filtraAcentuacao(_vRastreio[iv].dadosRaw.Situacao.Trim().ToUpper());
						sOcorrenciaDetalhe = Global.filtraAcentuacao(_vRastreio[iv].dadosRaw.Detalhe.Trim().ToUpper());
						if (sOcorrenciaSituacao.Contains("DEVOLUC") || sOcorrenciaSituacao.Contains("DEVOLV")) blnDevolucao = true;
						if (sOcorrenciaDetalhe.Contains("DEVOLUC") || sOcorrenciaDetalhe.Contains("DEVOLV")) blnDevolucao = true;
						if (blnDevolucao)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_DEVOLUCAO;
							_vRastreio[iv].processo.MensagemErro = "Esta ocorrência não será processada por haver indícios de se tratar de uma devolução";

							#region [ Se houver outras ocorrências para a mesma NF neste arquivo, bloqueia o processamento ]
							for (int jv = 0; jv < _vRastreio.Count; jv++)
							{
								// É este próprio registro
								if (_vRastreio[iv].processo.Guid.Equals(_vRastreio[jv].processo.Guid)) continue;
								// Já está com algum outro erro
								if (_vRastreio[jv].processo.Status != eRastreioPedidoRecebidoClienteProcessoStatus.StatusInicial) continue;

								if ((_vRastreio[iv].dadosNormalizado.CnpjCpfRemetente.Equals(_vRastreio[jv].dadosNormalizado.CnpjCpfRemetente))
									&& (_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario.Equals(_vRastreio[jv].dadosNormalizado.CnpjCpfDestinatario))
									&& (_vRastreio[iv].dadosNormalizado.numSerieNF == _vRastreio[jv].dadosNormalizado.numSerieNF)
									&& (_vRastreio[iv].dadosNormalizado.numNF == _vRastreio[jv].dadosNormalizado.numNF))
								{
									_vRastreio[jv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
									_vRastreio[jv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_DEVOLUCAO;
									_vRastreio[jv].processo.MensagemErro = "Esta ocorrência não será processada por estar relacionada com outra ocorrência deste arquivo que possui indícios de se tratar de uma devolução";
								}
							}
							#endregion
						}
					}
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return;
				}
				finally
				{
					info(ModoExibicaoMensagemRodape.Normal);
				}
				#endregion

				#region [ Analisa se há duplicidade ]
				if (paramProcDtPrevEntregaTransp.campo_inteiro == 0)
				{
					strMsgProgresso = "Analisando existência de registros em duplicidade";
					info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
					try
					{
						for (int iv = 0; iv < _vRastreio.Count; iv++)
						{
							if (_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length == 0) continue;

							#region [ Verifica se linha está repetida ]
							for (int jv = 0; jv < _vRastreio.Count; jv++)
							{
								if (_vRastreio[jv].dadosNormalizado.DataEntrega.Trim().Length == 0) continue;

								if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.StatusInicial)
									&& (_vRastreio[jv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.StatusInicial)
									&& (!_vRastreio[iv].processo.Guid.Equals(_vRastreio[jv].processo.Guid))
									&& (_vRastreio[iv].dadosNormalizado.CnpjCpfRemetente.Equals(_vRastreio[jv].dadosNormalizado.CnpjCpfRemetente))
									&& (_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario.Equals(_vRastreio[jv].dadosNormalizado.CnpjCpfDestinatario))
									&& (_vRastreio[iv].dadosNormalizado.numSerieNF == _vRastreio[jv].dadosNormalizado.numSerieNF)
									&& (_vRastreio[iv].dadosNormalizado.numNF == _vRastreio[jv].dadosNormalizado.numNF)
									&& (_vRastreio[iv].dadosNormalizado.DataEntrega.Equals(_vRastreio[jv].dadosNormalizado.DataEntrega)))
								{
									_vRastreio[jv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
									_vRastreio[jv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_REPETIDA;
									_vRastreio[jv].processo.MensagemErro = "Esta ocorrência será ignorada por já existir outra igual no mesmo arquivo";
								}
							}
							#endregion
						}
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(ex.ToString());
						adicionaErro(ex.Message);
						avisoErro(ex.ToString());
						return;
					}
					finally
					{
						info(ModoExibicaoMensagemRodape.Normal);
					}
				}
				#endregion

				#region [ Pesquisa as NFs no BD ]
				try
				{
					percProgressoAnterior = -1;
					dtHrUltProgresso = DateTime.MinValue;
					for (int iv = 0; iv < _vRastreio.Count; iv++)
					{
						lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
						percProgresso = 100 * (iv + 1) / _vRastreio.Count;
						if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
						{
							strMsgProgresso = "Consultando informações no banco de dados: linha " + (iv + 1).ToString() + " / " + _vRastreio.Count.ToString() + "   (" + percProgresso.ToString() + "%)";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							percProgressoAnterior = percProgresso;
							Application.DoEvents();
							// Atualiza o horário após o DoEvents() para que o intervalo entre os DoEvents() não seja inferior ao mínimo definido
							dtHrUltProgresso = DateTime.Now;
						}

						#region [ Registro já reprovado por consistência anterior? ]
						if (_vRastreio[iv].processo.Status != eRastreioPedidoRecebidoClienteProcessoStatus.StatusInicial)
						{
							continue;
						}
						#endregion

						#region [ Consistências ]

						#region [ Há nº NF? ]
						if (_vRastreio[iv].dadosRaw.NF.Trim().Length == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.NUMERO_NF_NAO_INFORMADO;
							_vRastreio[iv].processo.MensagemErro = "NF não informada";
							continue;
						}
						#endregion

						#region [ Nº NF em formato válido? ]
						// Considera que o arquivo deve informar no formato: [Nº Série] [Espaço Branco] [Nº NF]
						if ((_vRastreio[iv].dadosRaw.NF.Trim().Length > 0) && (_vRastreio[iv].dadosNormalizado.NF.Trim().Length == 0))
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.NUMERO_NF_FORMATO_INVALIDO;
							_vRastreio[iv].processo.MensagemErro = "NF foi informada em formato inválido (" + _vRastreio[iv].dadosRaw.NF.Trim() + ")";
							continue;
						}
						#endregion

						#region [ Nº NF válido? ]
						// Considera que o arquivo deve informar no formato: [Nº Série] [Espaço Branco] [Nº NF]
						if (_vRastreio[iv].dadosNormalizado.numNF == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.NUMERO_NF_FORMATO_INVALIDO;
							_vRastreio[iv].processo.MensagemErro = "NF informada é inválida (" + _vRastreio[iv].dadosRaw.NF.Trim() + ")";
							continue;
						}
						#endregion

						#region [ Nº série da NF válido? ]
						// Considera que o arquivo deve informar no formato: [Nº Série] [Espaço Branco] [Nº NF]
						if (_vRastreio[iv].dadosNormalizado.numSerieNF == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.NUMERO_NF_FORMATO_INVALIDO;
							_vRastreio[iv].processo.MensagemErro = "Nº série da NF informada é inválido (" + _vRastreio[iv].dadosRaw.NF.Trim() + ")";
							continue;
						}
						#endregion

						if (paramProcDtPrevEntregaTransp.campo_inteiro == 0)
						{
							#region [ Verifica campo 'Situacao' ]
							if (!_vRastreio[iv].dadosNormalizado.Situacao.Trim().ToUpper().Equals("MERCADORIA ENTREGUE"))
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_COM_SITUACAO_INVALIDA;
								_vRastreio[iv].processo.MensagemErro = "Ocorrência informa situação diferente de 'Mercadoria Entregue'";
								continue;
							}
							#endregion

							#region [ Ocorrência informa data de recebimento pelo cliente? ]
							if (_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length == 0)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_SEM_DATA_RECEBIMENTO;
								_vRastreio[iv].processo.MensagemErro = "Ocorrência não informa data de recebimento";
								continue;
							}
							#endregion
						}
						else
						{
							#region [ Ocorrência informa data de recebimento pelo cliente ou data de previsão de entrega? ]
							if ((_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length == 0) && (_vRastreio[iv].dadosNormalizado.PrevisaoEntrega.Trim().Length == 0))
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_SEM_DATA_RECEBIMENTO_E_SEM_DATA_PREVISAO_ENTREGA;
								_vRastreio[iv].processo.MensagemErro = "Ocorrência não informa data de recebimento ou data de previsão de entrega";
								continue;
							}
							#endregion
						}

						#region [ Tenta localizar o pedido através da NF ]
						listaPedidos = FMain.contextoBD.AmbienteBase.pedidoDAO.getPedidoByNF(_vRastreio[iv].dadosNormalizado.CnpjCpfRemetente, _vRastreio[iv].dadosNormalizado.numSerieNF, _vRastreio[iv].dadosNormalizado.numNF, flagNaoCarregarItens: true);

						if (listaPedidos.Count == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.PEDIDO_NAO_LOCALIZADO_POR_NF;
							_vRastreio[iv].processo.MensagemErro = "Pedido não localizado através do nº NF";
							continue;
						}
						#endregion

						#region [ Há mais de um pedido encontrado? ]
						if (listaPedidos.Count > 1)
						{
							sPedidos = "";
							for (int i = 0; i < listaPedidos.Count; i++)
							{
								if (sPedidos.Length > 0) sPedidos += ", ";
								sPedidos += listaPedidos[i].pedido;
							}
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.MULTIPLOS_PEDIDOS_LOCALIZADOS_PARA_NF;
							_vRastreio[iv].processo.MensagemErro = listaPedidos.Count().ToString() + " pedidos localizados para a NF: " + sPedidos;
							continue;
						}
						#endregion

						pedido = listaPedidos[0];

						#region [ Memoriza dados do pedido usados no processamento ]
						_vRastreio[iv].processo.Pedido = pedido.pedido;
						_vRastreio[iv].processo.marketplace_codigo_origem = pedido.marketplace_codigo_origem;
						_vRastreio[iv].processo.MarketplacePedidoRecebidoRegistrarStatus = pedido.MarketplacePedidoRecebidoRegistrarStatus;
						_vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal = pedido.PrevisaoEntregaTranspData;
						#endregion

						#region [ Verifica se já consta como recebido pelo cliente ]
						if ((paramProcDtPrevEntregaTransp.campo_inteiro == 0) || (_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length > 0))
						{
							if (pedido.pedidoRecebidoStatus == Global.Cte.StPedidoRecebido.COD_ST_PEDIDO_RECEBIDO_SIM)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.PEDIDO_RECEBIDO_JA_REGISTRADO;
								_vRastreio[iv].processo.MensagemErro = "Pedido " + pedido.pedido + " já consta como recebido em " + Global.formataDataDdMmYyyyComSeparador(pedido.pedidoRecebidoData);
								continue;
							}
						}
						#endregion

						#region [ Verifica status do campo 'st_entrega' ]
						if (!pedido.st_entrega.Equals(Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.PEDIDO_ST_ENTREGA_INVALIDO;
							_vRastreio[iv].processo.MensagemErro = "Pedido " + pedido.pedido + " possui status inválido: " + Global.stEntregaPedidoDescricao(pedido.st_entrega).ToUpper();
							continue;
						}
						#endregion

						#region [ Pedido possui transportadora cadastrada? ]
						if (pedido.transportadora_id.Trim().Length == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.PEDIDO_SEM_TRANSPORTADORA_CADASTRADA;
							_vRastreio[iv].processo.MensagemErro = "Pedido " + pedido.pedido + " não possui nenhuma transportadora cadastrada";
							continue;
						}
						#endregion

						#region [ Verifica se a data de recebimento está coerente com a data 'entregue_data' ]
						if ((paramProcDtPrevEntregaTransp.campo_inteiro == 0) || (_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length > 0))
						{
							if (_vRastreio[iv].dadosNormalizado.dtDataEntrega < pedido.entregue_data)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.DATA_RECEBIMENTO_ANTERIOR_DATA_PEDIDO_ENTREGUE;
								_vRastreio[iv].processo.MensagemErro = "Pedido " + pedido.pedido + ": a data de recebimento informada (" + _vRastreio[iv].dadosNormalizado.DataEntrega + ") é anterior à data do pedido entregue (" + Global.formataDataDdMmYyyyComSeparador(pedido.entregue_data) + ")";
								continue;
							}
						}
						#endregion

						#region [ Verifica se CPF/CNPJ do cliente confere com o que consta no sistema p/ a NFe emitida ]
						try
						{
							listaNFeImagem = FMain.contextoBD.AmbienteBase.nfeDAO.getNFeImagemByNF(_vRastreio[iv].dadosNormalizado.CnpjCpfRemetente, _vRastreio[iv].dadosNormalizado.numSerieNF, _vRastreio[iv].dadosNormalizado.numNF);
						}
						catch (Exception)
						{
							listaNFeImagem = null;
						}

						#region [ Nenhuma NFe localizada no sistema com esse número ]
						if (listaNFeImagem == null)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.NUMERO_NF_NAO_ENCONTRADO;
							_vRastreio[iv].processo.MensagemErro = "NFe " + _vRastreio[iv].dadosNormalizado.NF + " não foi encontrada no sistema!";
							continue;
						}

						if (listaNFeImagem.Count == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
							_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.NUMERO_NF_NAO_ENCONTRADO;
							_vRastreio[iv].processo.MensagemErro = "NFe " + _vRastreio[iv].dadosNormalizado.NF + " não foi localizada no sistema!";
							continue;
						}
						#endregion

						// Analisa o registro da emissão mais recente (se houver mais de um registro, a lista é retornada em ordem decrescente)
						nFeImagem = listaNFeImagem[0];

						if (Global.digitos(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario).Length == Global.Cte.Etc.TAMANHO_CPF)
						{
							if (!Global.digitos(nFeImagem.dest__CPF).Equals(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario))
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.CNPJ_CPF_DIVERGENTE;
								_vRastreio[iv].processo.MensagemErro = "A NFe " + _vRastreio[iv].dadosNormalizado.NF + " foi informada como sendo do cliente " + Global.formataCnpjCpf(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario) + " mas no sistema consta que foi emitida para " + Global.formataCnpjCpf(nFeImagem.dest__CPF);
								continue;
							}
						}
						else
						{
							if (!Global.digitos(nFeImagem.dest__CNPJ).Equals(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario))
							{
								blnCnpjCpfOk = false;
								#region [ Verifica se é a situação em que o CPF do cliente é informado formatado como CNPJ (ex: 002.449.961/72 informado como 00.000.244/9961-72) ]
								sCpfLeftPaddedZero = Global.digitos(nFeImagem.dest__CPF);
								sCpfRightPaddedZero = Global.digitos(nFeImagem.dest__CPF);
								while (sCpfLeftPaddedZero.Length < _vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario.Length)
								{
									sCpfLeftPaddedZero = '0' + sCpfLeftPaddedZero;
								}
								while (sCpfRightPaddedZero.Length < _vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario.Length)
								{
									sCpfRightPaddedZero = sCpfRightPaddedZero + '0';
								}
								if (sCpfLeftPaddedZero.Equals(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario) || sCpfRightPaddedZero.Equals(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario)) blnCnpjCpfOk = true;
								#endregion

								if (!blnCnpjCpfOk)
								{
									_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
									_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.CNPJ_CPF_DIVERGENTE;
									_vRastreio[iv].processo.MensagemErro = "A NFe " + _vRastreio[iv].dadosNormalizado.NF + " foi informada como sendo do cliente " + Global.formataCnpjCpf(_vRastreio[iv].dadosNormalizado.CnpjCpfDestinatario) + " mas no sistema consta que foi emitida para " + Global.formataCnpjCpf(nFeImagem.dest__CNPJ);
									continue;
								}
							}
						}
						#endregion

						#endregion

						// Se chegou até este ponto, está apto para registrar os dados de pedido recebido pelo cliente e/ou a data de previsão de entrega
						if (paramProcDtPrevEntregaTransp.campo_inteiro == 0)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente;
						}
						else
						{
							if ((_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length > 0) && (_vRastreio[iv].dadosNormalizado.PrevisaoEntrega.Trim().Length > 0))
							{
								if (_vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega != _vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal)
								{
									_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega;
								}
								else
								{
									_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente;
								}
							}
							else if (_vRastreio[iv].dadosNormalizado.DataEntrega.Trim().Length > 0)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente;
							}
							else if ((_vRastreio[iv].dadosNormalizado.PrevisaoEntrega.Trim().Length > 0) && (_vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega != _vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal))
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarDataPrevisaoEntrega;
							}
							else
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia;
								_vRastreio[iv].processo.CodigoErro = eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_SEM_DATA_RECEBIMENTO_E_SEM_DATA_PREVISAO_ENTREGA;
								_vRastreio[iv].processo.MensagemErro = "Ocorrência não informa data de recebimento ou alteração na previsão de entrega";
							}
						}
					} //for (int iv = 0; iv < vRastreio.Count; iv++)
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return;
				}
				finally
				{
					info(ModoExibicaoMensagemRodape.Normal);
				}
				#endregion

				#region [ Ordena a lista ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "Ordenando a listagem");
				for (int iv = 0; iv < _vRastreio.Count; iv++)
				{
					if (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia)
					{
						#region [ 1ª posição: linhas com erros (meramente informativas, não há ação por parte do usuário) ]
						sOrdenacao = Global.normalizaCodigo("1", 2) + Global.normalizaCodigo(((int)_vRastreio[iv].processo.CodigoErro).ToString(), 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numSerieNF.ToString(), 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente) || (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega))
					{
						#region [ 2ª posição: pedidos recebidos pelo cliente ]
						sOrdenacao = Global.normalizaCodigo("2", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numSerieNF.ToString(), 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else if (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarDataPrevisaoEntrega)
					{
						#region [ 3ª posição: registrar data de previsão de entrega ]
						sOrdenacao = Global.normalizaCodigo("3", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numSerieNF.ToString(), 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else
					{
						#region [ Situação desconhecida ]
						blnHaLinhasStatusDesconhecido = true;
						sOrdenacao = Global.normalizaCodigo("4", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numSerieNF.ToString(), 3) + Global.normalizaCodigo(_vRastreio[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}

					_vRastreio[iv].processo.campoOrdenacao = sOrdenacao;
				}

				vRastreioOrdenado = _vRastreio.OrderBy(o => o.processo.campoOrdenacao).ToList();
				#endregion

				#region [ Preenche o grid ]
				try
				{
					grid.SuspendLayout();
					grid.Rows.Add(vRastreioOrdenado.Count);

					#region [ Mantém a exibição do grid sem nenhuma linha selecionada enquanto os dados são carregados ]
					for (int i = 0; i < grid.Rows.Count; i++)
					{
						if (grid.Rows[i].Selected) grid.Rows[i].Selected = false;
					}
					#endregion

					try
					{
						sYYYYMMDDBrancos = new string(' ', Global.formataDataYyyyMmDdComSeparador(DateTime.Now).Length);
						sNFBrancos = Global.normalizaCodigo("0", 9);

						percProgressoAnterior = -1;
						dtHrUltProgresso = DateTime.MinValue;
						for (int iv = 0; iv < vRastreioOrdenado.Count; iv++)
						{
							lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
							percProgresso = 100 * (iv + 1) / vRastreioOrdenado.Count;
							if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
							{
								strMsgProgresso = "Carregando dados no grid: linha " + (iv + 1).ToString() + " / " + vRastreioOrdenado.Count.ToString() + "   (" + percProgresso.ToString() + "%)";
								info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
								percProgressoAnterior = percProgresso;
								Application.DoEvents();
								// Atualiza o horário após o DoEvents() para que o intervalo entre os DoEvents() não seja inferior ao mínimo definido
								dtHrUltProgresso = DateTime.Now;
							}

							grid.Rows[iv].Cells[GRID_COL_HIDDEN_GUID].Value = vRastreioOrdenado[iv].processo.Guid;
							grid.Rows[iv].Cells[GRID_COL_VISIBLE_ORDENACAO_PADRAO].Value = (iv + 1).ToString() + ".";
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value = iv;
							grid.Rows[iv].Cells[GRID_COL_NF].Value = vRastreioOrdenado[iv].dadosNormalizado.NF;
							sNF = Global.normalizaCodigo(vRastreioOrdenado[iv].dadosNormalizado.NF, 9);
							if (sNF.Length == 0) sNF = sNFBrancos;
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_NF].Value = sNF + ' ' + Global.normalizaCodigo(iv.ToString(), 6);
							grid.Rows[iv].Cells[GRID_COL_DESTINATARIO].Value = vRastreioOrdenado[iv].dadosNormalizado.Destinatario;
							grid.Rows[iv].Cells[GRID_COL_DESTINO].Value = vRastreioOrdenado[iv].dadosNormalizado.UfDestino + " / " + vRastreioOrdenado[iv].dadosNormalizado.CidadeDestino;
							grid.Rows[iv].Cells[GRID_COL_SITUACAO].Value = vRastreioOrdenado[iv].dadosNormalizado.Situacao;
							grid.Rows[iv].Cells[GRID_COL_DETALHE].Value = vRastreioOrdenado[iv].dadosNormalizado.Detalhe;
							grid.Rows[iv].Cells[GRID_COL_DATA_ENTREGA].Value = vRastreioOrdenado[iv].dadosNormalizado.DataEntrega;
							sDatePart = Global.formataDataYyyyMmDdComSeparador(vRastreioOrdenado[iv].dadosNormalizado.dtDataEntrega);
							if (sDatePart.Length == 0) sDatePart = sYYYYMMDDBrancos;
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_DATA_ENTREGA].Value = sDatePart + ' ' + Global.normalizaCodigo(iv.ToString(), 6);
							grid.Rows[iv].Cells[GRID_COL_PREVISAO_ENTREGA].Value = vRastreioOrdenado[iv].dadosNormalizado.PrevisaoEntrega;
							sDatePart = Global.formataDataYyyyMmDdComSeparador(vRastreioOrdenado[iv].dadosNormalizado.dtPrevisaoEntrega);
							if (sDatePart.Length == 0) sDatePart = sYYYYMMDDBrancos;
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_PREVISAO_ENTREGA].Value = sDatePart + ' ' + Global.normalizaCodigo(iv.ToString(), 6);

							if (vRastreioOrdenado[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.ErroInconsistencia)
							{
								qtdeRegErro++;
								grid.Rows[iv].Cells[GRID_COL_STATUS].Value = "ERRO";
								grid.Rows[iv].Cells[GRID_COL_MENSAGEM].Value = vRastreioOrdenado[iv].processo.MensagemErro;
								if ((vRastreioOrdenado[iv].processo.CodigoErro == eRastreioPedidoRecebidoClienteProcessoCodigoErro.OCORRENCIA_COM_SITUACAO_INVALIDA)
									|| (vRastreioOrdenado[iv].processo.CodigoErro == eRastreioPedidoRecebidoClienteProcessoCodigoErro.PEDIDO_RECEBIDO_JA_REGISTRADO))
								{
									grid.Rows[iv].DefaultCellStyle.ForeColor = Color.DarkViolet;
								}
								else
								{
									grid.Rows[iv].DefaultCellStyle.ForeColor = Color.Red;
								}
							}
							else if ((vRastreioOrdenado[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente)
									|| (vRastreioOrdenado[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega)
									|| (vRastreioOrdenado[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarDataPrevisaoEntrega)
									)
							{
								qtdeRegApto++;
							}
							else
							{
								// SITUAÇÃO DESCONHECIDA: ALTERA A COR P/ CHAMAR A ATENÇÃO PARA A SITUAÇÃO
								grid.Rows[iv].DefaultCellStyle.BackColor = Color.DeepPink;
							}
						}

						#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
						for (int i = 0; i < grid.Rows.Count; i++)
						{
							if (grid.Rows[i].Selected) grid.Rows[i].Selected = false;
						}
						#endregion
					}
					finally
					{
						//Exibe o grid sem nenhuma linha pré-selecionada
						grid.ClearSelection();

						grid.ResumeLayout();
					}
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return;
				}
				#endregion

				lblTotalRegistros.Text = Global.formataInteiro(_vRastreio.Count);
				lblQtdeRegErro.Text = Global.formataInteiro(qtdeRegErro);
				lblQtdeRegApto.Text = Global.formataInteiro(qtdeRegApto);

				tsDuracaoProcessamento = DateTime.Now - dtInicioProcessamento;

				#region [ Mensagem de sucesso ]
				info(ModoExibicaoMensagemRodape.Normal);
				strMsg = "Leitura do arquivo concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!";
				adicionaDisplay(strMsg);
				aviso(strMsg);
				#endregion

				if (blnHaLinhasStatusDesconhecido) aviso("ATENÇÃO: há linha(s) com status desconhecido!\nFavor informar o suporte técnico sobre essa situação!");

				grid.Focus();
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				adicionaErro(ex.Message);
				avisoErro(ex.ToString());
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ executaAnotaPedidoRecebidoCliente ]
		private void executaAnotaPedidoRecebidoCliente()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FAnotaPedidoRecebidoCliente.executaAnotaPedidoRecebidoCliente";
			int qtdeRegistrosParaAtualizar = 0;
			int qtdeRegistrosAtualizados = 0;
			int qtdeRegistrosAtualizadosSucesso = 0;
			int qtdeRegistrosAtualizadosFalha = 0;
			int qtdeRegistrosAtualizadosSucessoUpdatePedidoRecebidoData = 0;
			int qtdeRegistrosAtualizadosFalhaUpdatePedidoRecebidoData = 0;
			int qtdeRegistrosAtualizadosSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = 0;
			int qtdeRegistrosAtualizadosFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = 0;
			int qtdeRegistrosAtualizadosSucessoUpdatePrevisaoEntregaTranspData = 0;
			int qtdeRegistrosAtualizadosFalhaUpdatePrevisaoEntregaTranspData = 0;
			int percProgresso;
			int percProgressoAnterior;
			long lngMiliSegundosDecorridos;
			bool blnFalhaAtualizacao;
			bool blnUpdatePedidoRecebidoData;
			bool blnUpdateMarketplacePedidoRecebidoRegistrarDataRecebido;
			bool blnUpdatePrevisaoEntregaTranspData;
			bool blnRegistroLiberadoProcessamento;
			eRastreioPedidoRecebidoClienteProcessoStatus statusProcessoOriginal;
			string strMsg;
			string strMsgProgresso;
			string strMsgErroLog = "";
			string msg_erro;
			DateTime dtHrUltProgresso;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			StringBuilder sbLogSucessoUpdatePedidoRecebidoData = new StringBuilder("");
			StringBuilder sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = new StringBuilder("");
			StringBuilder sbLogSucessoUpdatePrevisaoEntregaTranspData = new StringBuilder("");
			StringBuilder sbLogFalhaUpdatePedidoRecebidoData = new StringBuilder("");
			StringBuilder sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = new StringBuilder("");
			StringBuilder sbLogFalhaUpdatePrevisaoEntregaTranspData = new StringBuilder("");
			Log log = new Log();
			#endregion

			try
			{
				for (int iv = 0; iv < _vRastreio.Count; iv++)
				{
					if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente)
						|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega)
						|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarDataPrevisaoEntrega))
					{
						qtdeRegistrosParaAtualizar++;
					}
				}

				if (qtdeRegistrosParaAtualizar == 0)
				{
					avisoErro("Não há nenhum registro para ser atualizado no banco de dados!");
					return;
				}

				#region [ Solicita confirmação antes de executar a operação ]
				strMsg = "Confirma a atualização no banco de dados?";
				if (!confirma(strMsg)) return;
				#endregion

				#region [ Inicialização do processamento ]
				dtInicioProcessamento = DateTime.Now;
				strMsg = "Início da atualização no banco de dados";
				adicionaDisplay(strMsg);
				strMsgProgresso = "Atualizando pedidos no banco de dados";
				info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
				#endregion

				percProgressoAnterior = -1;
				dtHrUltProgresso = DateTime.MinValue;
				for (int iv = 0; iv < _vRastreio.Count; iv++)
				{
					blnRegistroLiberadoProcessamento = false;
					blnFalhaAtualizacao = false;
					statusProcessoOriginal = _vRastreio[iv].processo.Status;

					#region [ Registra pedido recebido pelo cliente ]
					if ((statusProcessoOriginal == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente)
						|| (statusProcessoOriginal == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega))
					{
						#region [ Executa atualização no banco de dados ]
						blnRegistroLiberadoProcessamento = true;
						blnUpdatePedidoRecebidoData = FMain.contextoBD.AmbienteBase.anotarPedidoRecebidoClienteDAO.UpdatePedidoRecebidoData(_vRastreio[iv].processo.Pedido, _vRastreio[iv].dadosNormalizado.dtDataEntrega, Global.Usuario.usuario, out msg_erro);
						if (blnUpdatePedidoRecebidoData)
						{
							qtdeRegistrosAtualizadosSucessoUpdatePedidoRecebidoData++;
							if (sbLogSucessoUpdatePedidoRecebidoData.Length > 0) sbLogSucessoUpdatePedidoRecebidoData.Append(", ");
							strMsg = _vRastreio[iv].processo.Pedido
									+ " (" + (_vRastreio[iv].dadosNormalizado.dtDataEntrega == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].dadosNormalizado.dtDataEntrega)) + ")";
							sbLogSucessoUpdatePedidoRecebidoData.Append(strMsg);
						}
						else
						{
							if (sbLogFalhaUpdatePedidoRecebidoData.Length > 0) sbLogFalhaUpdatePedidoRecebidoData.Append(", ");
							strMsg = _vRastreio[iv].processo.Pedido
									+ " (" + (_vRastreio[iv].dadosNormalizado.dtDataEntrega == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].dadosNormalizado.dtDataEntrega)) + ")";
							sbLogFalhaUpdatePedidoRecebidoData.Append(strMsg);

							blnFalhaAtualizacao = true;
							qtdeRegistrosAtualizadosFalhaUpdatePedidoRecebidoData++;
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente;
							strMsg = "Falha ao tentar atualizar o pedido " + _vRastreio[iv].processo.Pedido + " (NF: " + _vRastreio[iv].dadosNormalizado.NF + "): " + msg_erro;
							_vRastreio[iv].processo.MensagemErro = strMsg;
							adicionaErro(strMsg);
						}

						if (blnUpdatePedidoRecebidoData)
						{
							if ((_vRastreio[iv].processo.marketplace_codigo_origem.Trim().Length > 0) && (_vRastreio[iv].processo.MarketplacePedidoRecebidoRegistrarStatus == 0))
							{
								blnUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = FMain.contextoBD.AmbienteBase.anotarPedidoRecebidoClienteDAO.UpdateMarketplacePedidoRecebidoRegistrarDataRecebido(_vRastreio[iv].processo.Pedido, _vRastreio[iv].dadosNormalizado.dtDataEntrega, Global.Usuario.usuario, out msg_erro);
								if (blnUpdateMarketplacePedidoRecebidoRegistrarDataRecebido)
								{
									qtdeRegistrosAtualizadosSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido++;
									if (sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(", ");
									strMsg = _vRastreio[iv].processo.Pedido
											+ " (" + (_vRastreio[iv].dadosNormalizado.dtDataEntrega == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].dadosNormalizado.dtDataEntrega)) + ")";
									sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(strMsg);
								}
								else
								{
									if (sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(", ");
									strMsg = _vRastreio[iv].processo.Pedido
											+ " (" + (_vRastreio[iv].dadosNormalizado.dtDataEntrega == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].dadosNormalizado.dtDataEntrega)) + ")";
									sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(strMsg);

									blnFalhaAtualizacao = true;
									qtdeRegistrosAtualizadosFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido++;
									_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente;
									strMsg = "Falha ao tentar atualizar o pedido " + _vRastreio[iv].processo.Pedido + " (NF: " + _vRastreio[iv].dadosNormalizado.NF + "): " + msg_erro;
									_vRastreio[iv].processo.MensagemErro = strMsg;
									adicionaErro(strMsg);
								}
							}
						}
						#endregion

						if (!blnFalhaAtualizacao)
						{
							_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente;
						}
					}
					#endregion

					#region [ Registra data de previsão de entrega da transportadora ]
					if ((statusProcessoOriginal == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarDataPrevisaoEntrega)
						|| (statusProcessoOriginal == eRastreioPedidoRecebidoClienteProcessoStatus.LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega))
					{
						#region [ Executa atualização no banco de dados ]
						blnRegistroLiberadoProcessamento = true;
						blnUpdatePrevisaoEntregaTranspData = FMain.contextoBD.AmbienteBase.anotarPedidoRecebidoClienteDAO.UpdatePrevisaoEntregaTranspData(_vRastreio[iv].processo.Pedido, _vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega, _vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal, Global.Usuario.usuario, out msg_erro);
						if (blnUpdatePrevisaoEntregaTranspData)
						{
							qtdeRegistrosAtualizadosSucessoUpdatePrevisaoEntregaTranspData++;
							if (sbLogSucessoUpdatePrevisaoEntregaTranspData.Length > 0) sbLogSucessoUpdatePrevisaoEntregaTranspData.Append(", ");
							strMsg = _vRastreio[iv].processo.Pedido
									+ " ("
									+ (_vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal))
									+ " => "
									+ (_vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega))
									+ ")";
							sbLogSucessoUpdatePrevisaoEntregaTranspData.Append(strMsg);
							if (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega;
							}
							else if (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega;
							}
							else
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroDataPrevisaoEntrega;
							}
						}
						else
						{
							if (sbLogFalhaUpdatePrevisaoEntregaTranspData.Length > 0) sbLogFalhaUpdatePrevisaoEntregaTranspData.Append(", ");
							strMsg = _vRastreio[iv].processo.Pedido
									+ " ("
									+ (_vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].processo.PrevisaoEntregaTranspDataOriginal))
									+ " => "
									+ (_vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(_vRastreio[iv].dadosNormalizado.dtPrevisaoEntrega))
									+ ")";
							sbLogFalhaUpdatePrevisaoEntregaTranspData.Append(strMsg);

							blnFalhaAtualizacao = true;
							qtdeRegistrosAtualizadosFalhaUpdatePrevisaoEntregaTranspData++;
							if (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega;
							}
							else if (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente)
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega;
							}
							else
							{
								_vRastreio[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroDataPrevisaoEntrega;
							}
							strMsg = "Falha ao tentar atualizar o pedido " + _vRastreio[iv].processo.Pedido + " (NF: " + _vRastreio[iv].dadosNormalizado.NF + "): " + msg_erro;
							if (_vRastreio[iv].processo.MensagemErro.Trim().Length > 0) _vRastreio[iv].processo.MensagemErro += "\r\n";
							_vRastreio[iv].processo.MensagemErro += strMsg;
							adicionaErro(strMsg);
						}
						#endregion
					}
					#endregion

					// Registro ignorado para processamento
					if (!blnRegistroLiberadoProcessamento) continue;

					#region [ Progresso ]
					qtdeRegistrosAtualizados++;

					lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
					percProgresso = 100 * qtdeRegistrosAtualizados / qtdeRegistrosParaAtualizar;
					if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
					{
						strMsgProgresso = "Atualizando pedidos no banco de dados: " + qtdeRegistrosAtualizados.ToString() + " / " + qtdeRegistrosParaAtualizar.ToString() + "   (" + percProgresso.ToString() + "%)";
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
						percProgressoAnterior = percProgresso;
						Application.DoEvents();
						// Atualiza o horário após o DoEvents() para que o intervalo entre os DoEvents() não seja inferior ao mínimo definido
						dtHrUltProgresso = DateTime.Now;
					}
					#endregion

					#region [ Atualiza status no grid ]
					for (int jv = 0; jv < grid.Rows.Count; jv++)
					{
						if (grid.Rows[jv].Cells[GRID_COL_HIDDEN_GUID].Value.ToString().Equals(_vRastreio[iv].processo.Guid))
						{
							if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente)
								|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroDataPrevisaoEntrega)
								|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega))
							{
								grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "OK";
								grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.Green;
							}
							else if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente)
									|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroDataPrevisaoEntrega)
									|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega))
							{
								grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "FALHA";
								grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.Red;
								grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Value = _vRastreio[iv].processo.MensagemErro;
								grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Style.ForeColor = Color.Red;
							}
							else
							{
								// Resultado misto
								grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "";
								grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.OrangeRed;
								grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Value = _vRastreio[iv].processo.MensagemErro;
								grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Style.ForeColor = Color.OrangeRed;

								#region [ Pedido recebido pelo cliente ]
								if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente)
									|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega)
									|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega))
								{
									grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "OK (pedido recebido)";
								}
								else if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente)
										|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega)
										|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega))
								{
									grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "FALHA (pedido recebido)";
								}
								#endregion

								#region [ Data de previsão de entrega ]
								if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroDataPrevisaoEntrega)
									|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega)
									|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega))
								{
									if (grid.Rows[jv].Cells[GRID_COL_STATUS].Value.ToString().Length > 0) grid.Rows[jv].Cells[GRID_COL_STATUS].Value += "\r\n";
									grid.Rows[jv].Cells[GRID_COL_STATUS].Value += "OK (previsão entrega)";
								}
								else if ((_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroDataPrevisaoEntrega)
										|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.FalhaRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega)
										|| (_vRastreio[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.SucessoRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega))
								{
									if (grid.Rows[jv].Cells[GRID_COL_STATUS].Value.ToString().Length > 0) grid.Rows[jv].Cells[GRID_COL_STATUS].Value += "\r\n";
									grid.Rows[jv].Cells[GRID_COL_STATUS].Value += "FALHA (previsão entrega)";
								}
								#endregion
							}

							break;
						}
					}
					#endregion

					if (blnFalhaAtualizacao)
					{
						qtdeRegistrosAtualizadosFalha++;
						// Prossegue para o próximo registro
						continue;
					}

					qtdeRegistrosAtualizadosSucesso++;
					strMsg = "Sucesso na atualização do pedido " + _vRastreio[iv].processo.Pedido + " (NF: " + _vRastreio[iv].dadosNormalizado.NF + ")";
					adicionaDisplay(strMsg);
				} //for (int iv = 0; iv < _vRastreio.Count; iv++)

				lblQtdeAtualizSucesso.Text = Global.formataInteiro(qtdeRegistrosAtualizadosSucesso);
				lblQtdeAtualizFalha.Text = Global.formataInteiro(qtdeRegistrosAtualizadosFalha);

				#region [ Grava o log ]
				strMsg = "[Módulo ADM2] Operação 'Anotar Pedidos Recebidos pelo Cliente':";
				if (sbLogSucessoUpdatePedidoRecebidoData.Length > 0) strMsg += "\nSucesso (campo 'PedidoRecebidoData') [" + Global.formataInteiro(qtdeRegistrosAtualizadosSucessoUpdatePedidoRecebidoData) + " pedidos]: " + sbLogSucessoUpdatePedidoRecebidoData.ToString();
				if (sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) strMsg += "\nSucesso (campo 'MarketplacePedidoRecebidoRegistrarDataRecebido') [" + Global.formataInteiro(qtdeRegistrosAtualizadosSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido) + " pedidos]: " + sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.ToString();
				if (sbLogSucessoUpdatePrevisaoEntregaTranspData.Length > 0) strMsg += "\nSucesso (campo 'PrevisaoEntregaTranspData') [" + Global.formataInteiro(qtdeRegistrosAtualizadosSucessoUpdatePrevisaoEntregaTranspData) + " pedidos]: " + sbLogSucessoUpdatePrevisaoEntregaTranspData.ToString();
				if (sbLogFalhaUpdatePedidoRecebidoData.Length > 0) strMsg += "\nFalha (campo 'PedidoRecebidoData') [" + Global.formataInteiro(qtdeRegistrosAtualizadosFalhaUpdatePedidoRecebidoData) + " pedidos]: " + sbLogFalhaUpdatePedidoRecebidoData.ToString();
				if (sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) strMsg += "\nFalha (campo 'MarketplacePedidoRecebidoRegistrarDataRecebido') [" + Global.formataInteiro(qtdeRegistrosAtualizadosFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido) + " pedidos]: " + sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.ToString();
				if (sbLogFalhaUpdatePrevisaoEntregaTranspData.Length > 0) strMsg += "\nFalha (campo 'PrevisaoEntregaTranspData') [" + Global.formataInteiro(qtdeRegistrosAtualizadosFalhaUpdatePrevisaoEntregaTranspData) + " pedidos]: " + sbLogFalhaUpdatePrevisaoEntregaTranspData.ToString();
				strMsg += "\nArquivo processado: " + txtArquivoRastreio.Text.Trim() + " (contendo " + Global.formataInteiro(_vRastreio.Count) + " registros)";
				log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_PEDIDO_RECEBIDO_VIA_ADM2;
				log.usuario = Global.Usuario.usuario;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				tsDuracaoProcessamento = DateTime.Now - dtInicioProcessamento;

				#region [ Mensagem de sucesso ]
				info(ModoExibicaoMensagemRodape.Normal);
				strMsg = "Atualização no banco de dados concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!";
				adicionaDisplay(strMsg);
				aviso(strMsg);
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				adicionaErro(ex.Message);
				avisoErro(ex.ToString());
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FAnotarPedidoRecebidoCliente ]

		#region [ FAnotarPedidoRecebidoCliente_Load ]
		private void FAnotarPedidoRecebidoCliente_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				_tituloPainelOriginal = lblTituloPainel.Text;
				if (paramProcDtPrevEntregaTransp.campo_inteiro == 1)
				{
					lblTituloPainel.Text += "/Previsão Entrega";
				}

				txtArquivoRastreio.Text = "";
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

		#region [ FAnotarPedidoRecebidoCliente_Shown ]
		private void FAnotarPedidoRecebidoCliente_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					openFileDialogCtrl.InitialDirectory = pathArquivoRastreioValorDefault();
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

		#region [ FAnotarPedidoRecebidoCliente_FormClosing ]
		private void FAnotarPedidoRecebidoCliente_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaArquivoRastreio ]

		#region [ btnSelecionaArquivoRastreio_Click ]
		private void btnSelecionaArquivoRastreio_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArquivoRastreio();
		}
		#endregion

		#endregion

		#region [ grid ]

		#region [ grid_SortCompare ]
		private void grid_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			#region [ Declarações ]
			string sValue1;
			string sValue2;
			#endregion

			switch (e.Column.Name)
			{
				case GRID_COL_NF:
					sValue1 = grid.Rows[e.RowIndex1].Cells[GRID_COL_HIDDEN_NF].Value.ToString();
					sValue2 = grid.Rows[e.RowIndex2].Cells[GRID_COL_HIDDEN_NF].Value.ToString();
					e.SortResult = String.Compare(sValue1, sValue2);
					e.Handled = true;
					break;
				case GRID_COL_DATA_ENTREGA:
					sValue1 = grid.Rows[e.RowIndex1].Cells[GRID_COL_HIDDEN_DATA_ENTREGA].Value.ToString();
					sValue2 = grid.Rows[e.RowIndex2].Cells[GRID_COL_HIDDEN_DATA_ENTREGA].Value.ToString();
					e.SortResult = String.Compare(sValue1, sValue2);
					e.Handled = true;
					break;
				case GRID_COL_PREVISAO_ENTREGA:
					sValue1 = grid.Rows[e.RowIndex1].Cells[GRID_COL_HIDDEN_PREVISAO_ENTREGA].Value.ToString();
					sValue2 = grid.Rows[e.RowIndex2].Cells[GRID_COL_HIDDEN_PREVISAO_ENTREGA].Value.ToString();
					e.SortResult = String.Compare(sValue1, sValue2);
					e.Handled = true;
					break;
				case GRID_COL_VISIBLE_ORDENACAO_PADRAO:
					// Obs: a coluna 'ColVisibleOrdenacaoPadrao' é a coluna visível usada p/ poder ser clicada e fazer a ordenação conforme o padrão inicial, sendo que as células dessa coluna ficam vazias.
					// E a coluna 'ColHiddenValorOrdenacaoPadrao' é a coluna invisível que possui os dados usados p/ a ordenação padrão.
					sValue1 = grid.Rows[e.RowIndex1].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value.ToString();
					sValue2 = grid.Rows[e.RowIndex2].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value.ToString();
					e.SortResult = Global.converteInteiro(sValue1).CompareTo(Global.converteInteiro(sValue2));
					e.Handled = true;
					break;
				default:
					break;
			}
		}
		#endregion

		#endregion

		#region [ btnConfirma ]

		#region [ btnConfirma_Click ]
		private void btnConfirma_Click(object sender, EventArgs e)
		{
			trataBotaoConfirma();
		}
		#endregion

		#endregion

		#endregion
	}
}
