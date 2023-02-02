using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADM2
{
	public partial class FAnotarFretePedidoViaCsv : FModelo
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
		List<FretePedidoViaCsv> _vFrete;
		#endregion

		#region [ Constantes ]
		// Obs: a coluna 'ColVisibleOrdenacaoPadrao' é a coluna visível usada p/ poder ser clicada e fazer a ordenação conforme o padrão inicial, sendo que as células dessa coluna ficam vazias.
		// E a coluna 'ColHiddenValorOrdenacaoPadrao' é a coluna invisível que possui os dados usados p/ a ordenação padrão.
		const string GRID_COL_VISIBLE_ORDENACAO_PADRAO = "ColVisibleOrdenacaoPadrao";
		const string GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO = "ColHiddenValorOrdenacaoPadrao";
		const string GRID_COL_NF = "NF";
		const string GRID_COL_HIDDEN_NF = "ColHiddenNF";
		const string GRID_COL_HIDDEN_GUID = "ColHiddenGuid";
		const string GRID_COL_REMETENTE = "Remetente";
		const string GRID_COL_TRANSPORTADORA_CSV = "TransportadoraCsv";
		const string GRID_COL_TRANSPORTADORA_PEDIDO = "TransportadoraPedido";
		const string GRID_COL_VL_FRETE = "VlFrete";
		const string GRID_COL_HIDDEN_VL_FRETE = "ColHiddenVlFrete";
		const string GRID_COL_TIPO_FRETE = "TipoFrete";
		const string GRID_COL_HIDDEN_TIPO_FRETE = "ColHiddenTipoFrete";
		const string GRID_COL_PEDIDO = "Pedido";
		const string GRID_COL_STATUS = "Status";
		const string GRID_COL_MENSAGEM = "Mensagem";
		const int MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS = 500;
		#endregion

		#region [ Construtor ]
		public FAnotarFretePedidoViaCsv()
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

		#region [ pathArquivoCsvValorDefault ]
		private string pathArquivoCsvValorDefault()
		{
			string strResp = "";

			try
			{
				strResp = Path.GetPathRoot(Application.StartupPath);
			}
			catch (Exception)
			{
				strResp = "";
			}

			if (strResp.Length == 0) strResp = @"\";
			if (Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.pathArquivoCsv.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.pathArquivoCsv))
				{
					strResp = Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.pathArquivoCsv;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNameArquivoCsvValorDefault ]
		private string fileNameArquivoCsvValorDefault()
		{
			string strResp = "";

			if ((Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.fileNameArquivoCsv ?? "").Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.pathArquivoCsv + "\\" + Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.fileNameArquivoCsv))
				{
					strResp = Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.fileNameArquivoCsv;
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
			_vFrete = new List<FretePedidoViaCsv>();
		}
		#endregion

		#region [ trataBotaoSelecionaArquivoCsv ]
		private void trataBotaoSelecionaArquivoCsv()
		{
			#region [ Declarações ]
			DialogResult dr;
			#endregion

			try
			{
				openFileDialogCtrl.InitialDirectory = pathArquivoCsvValorDefault();
				openFileDialogCtrl.FileName = "";
				dr = openFileDialogCtrl.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				//if ((openFileDialogCtrl.FileName.Length > 0) && (txtArquivoCsv.Text.Length > 0))
				//{
				//	if (openFileDialogCtrl.FileName.ToUpper().Equals(txtArquivoCsv.Text.ToUpper())) return;
				//}
				#endregion

				#region [ Limpa campos ]
				limpaCampos();
				#endregion

				txtArquivoCsv.Text = openFileDialogCtrl.FileName;
				Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.pathArquivoCsv = Path.GetDirectoryName(openFileDialogCtrl.FileName);
				Global.Usuario.Defaults.FAnotarFretePedidoViaCsv.fileNameArquivoCsv = Path.GetFileName(openFileDialogCtrl.FileName);

				carregaDadosArquivoCsv();
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
			executaAnotaFretePedidoViaCsv();
		}
		#endregion

		#region [ carregaDadosArquivoCsv ]
		private void carregaDadosArquivoCsv()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FAnotarFretePedidoViaCsv.carregaDadosArquivoCsv";
			string MARGEM_MSG_NIVEL_2 = new string(' ', 8);
			string MARGEM_LINHA_OBS = new string(' ', 4);
			bool blnAchou;
			bool blnHaLinhasStatusDesconhecido = false;
			int qtdeLinhaDadosArquivo = 0;
			int percProgresso;
			int percProgressoAnterior;
			int qtdeRegErro = 0;
			int qtdeRegApto = 0;
			int linhaFreteCadastrado;
			long lngMiliSegundosDecorridos;
			string sNFBrancos;
			string sNF;
			string sVlFrete;
			string strMsgErroParam;
			string strMsgErro;
			string strMsg;
			string strMsgProgresso;
			string strNomeArquivo;
			string linhaHeader;
			string sPedidos;
			string sOrdenacao;
			string[] linhasCSV;
			string[] camposHeader;
			string[] camposCSV;
			string[] v;
			StringBuilder sbErro;
			StringBuilder sbAux;
			HeaderFretePedidoViaCsv header = new HeaderFretePedidoViaCsv();
			FretePedidoViaCsv frete;
			List<FretePedidoViaCsv> vFreteOrdenado;
			DateTime dtHrUltProgresso;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			List<CodigoDescricao> listaCodigoDescricao;
			CodigoDescricaoTipoFrete codigoDescricaoTipoFrete;
			List<CodigoDescricaoTipoFrete> listaCodigoDescricaoTipoFrete = new List<CodigoDescricaoTipoFrete>();
			List<NFeEmitente> listaNFeEmitente;
			List<Transportadora> listaTransportadora;
			Pedido pedido;
			List<Pedido> listaPedidos;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			try
			{
				#region [ Obtém o nome do arquivo ]
				strNomeArquivo = txtArquivoCsv.Text.Trim();
				#endregion

				#region [ Consistências ]
				if (strNomeArquivo.Length == 0)
				{
					strMsgErro = "É necessário selecionar o arquivo com os dados de frete!";
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

				#region [ Carrega dados do arquivo de frete ]
				try
				{
					#region [ Lê dados do arquivo ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Lendo dados do arquivo CSV");
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

					#region [ Carrega lista de tipos de frete configurados no sistema ]
					listaCodigoDescricao = FMain.contextoBD.AmbienteBase.geralDAO.getCodigoDescricaoByGrupo(Global.Cte.GruposCodigoDescricao.ID_GRUPO__PEDIDO_TIPO_FRETE, Global.eFiltroFlagStInativo.FLAG_IGNORADO, out strMsgErroParam);
					if ((listaCodigoDescricao == null))
					{
						strMsgErro = "Não foi possível localizar os tipos de frete configurados no sistema!";
						if ((strMsgErroParam ?? "").Trim().Length > 0) strMsgErro += "\n\n" + strMsgErroParam;
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}

					foreach (CodigoDescricao item in listaCodigoDescricao)
					{
						codigoDescricaoTipoFrete = new CodigoDescricaoTipoFrete();
						codigoDescricaoTipoFrete.grupo = item.grupo;
						codigoDescricaoTipoFrete.codigo = item.codigo;
						codigoDescricaoTipoFrete.descricao = item.descricao;
						codigoDescricaoTipoFrete.ordenacao = item.ordenacao;
						if ((item.parametro_2_campo_texto ?? "").Trim().Length > 0)
						{
							v = item.parametro_2_campo_texto.Split('|');
							foreach (string sCod in v)
							{
								if ((sCod ?? "").Trim().Length == 0) continue;
								codigoDescricaoTipoFrete.listaCodigosAceitosCsv.Add(sCod.Trim());
							}
						}
						listaCodigoDescricaoTipoFrete.Add(codigoDescricaoTipoFrete);
					}
					#endregion

					#region [ Carrega lista de emitentes de NFe ]
					listaNFeEmitente = FMain.contextoBD.AmbienteBase.geralDAO.getAllNFeEmitenteByCnpj(out strMsgErroParam);
					#endregion

					#region [ Carrega lista de transportadoras ]
					listaTransportadora = FMain.contextoBD.AmbienteBase.geralDAO.getAllTransportadora(out strMsgErroParam);
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

						frete = new FretePedidoViaCsv();

						#region [ Carrega dados do arquivo (da forma como vieram) ]
						frete.dadosRaw.linhaArquivoCsv = (i + 1);
						frete.dadosRaw.CnpjRemetente = camposCSV[(int)header.CnpjRemetente.indexColuna].Trim();
						frete.dadosRaw.NF = camposCSV[(int)header.NF.indexColuna].Trim();
						frete.dadosRaw.ValorFrete = camposCSV[(int)header.ValorFrete.indexColuna].Trim();
						frete.dadosRaw.TransportadoraCsv = camposCSV[(int)header.TransportadoraCsv.indexColuna].Trim();
						frete.dadosRaw.TipoFrete = camposCSV[(int)header.TipoFrete.indexColuna].Trim();
						#endregion

						#region [ Normaliza os dados ]
						frete.dadosNormalizado.CnpjRemetente = Global.digitos(frete.dadosRaw.CnpjRemetente);

						#region [ NF ]
						frete.dadosNormalizado.NF = Global.digitos(frete.dadosRaw.NF.Trim());
						frete.dadosNormalizado.numNF = (int)Global.converteInteiro(frete.dadosNormalizado.NF);
						#endregion

						#region [ Valor do frete ]
						frete.dadosNormalizado.ValorFrete = frete.dadosRaw.ValorFrete.Replace("R$", "").Trim();
						frete.dadosNormalizado.vlFrete = Global.converteNumeroDecimal(frete.dadosNormalizado.ValorFrete, ',');
						#endregion

						frete.dadosNormalizado.TransportadoraCsv = frete.dadosRaw.TransportadoraCsv.Trim().ToUpper();

						#region [ Tipo de frete ]
						frete.dadosNormalizado.TipoFrete = frete.dadosRaw.TipoFrete.Trim();
						blnAchou = false;
						foreach (CodigoDescricaoTipoFrete codDescTipoFrete in listaCodigoDescricaoTipoFrete)
						{
							foreach (string codAceito in codDescTipoFrete.listaCodigosAceitosCsv)
							{
								if (frete.dadosRaw.TipoFrete.Trim().ToUpper().Equals(codAceito.ToUpper()))
								{
									frete.dadosNormalizado.TipoFreteCodigoSistema = codDescTipoFrete.codigo;
									frete.dadosNormalizado.TipoFreteDescricaoSistema = codDescTipoFrete.descricao;
									frete.dadosNormalizado.TipoFreteOrdenacaoSistema = codDescTipoFrete.ordenacao;
									blnAchou = true;
									break;
								}
							}
							if (blnAchou) break;
						}
						#endregion

						#endregion

						_vFrete.Add(frete);
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
				for (int iv = 0; iv < _vFrete.Count; iv++)
				{
					_vFrete[iv].processo.Guid = Guid.NewGuid().ToString();
				}
				#endregion

				#region [ Pesquisa as NFs e demais dados no BD ]
				try
				{
					percProgressoAnterior = -1;
					dtHrUltProgresso = DateTime.MinValue;
					for (int iv = 0; iv < _vFrete.Count; iv++)
					{
						lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
						percProgresso = 100 * (iv + 1) / _vFrete.Count;
						if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
						{
							strMsgProgresso = "Consultando informações no banco de dados: linha " + (iv + 1).ToString() + " / " + _vFrete.Count.ToString() + "   (" + percProgresso.ToString() + "%)";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							percProgressoAnterior = percProgresso;
							Application.DoEvents();
							// Atualiza o horário após o DoEvents() para que o intervalo entre os DoEvents() não seja inferior ao mínimo definido
							dtHrUltProgresso = DateTime.Now;
						}

						#region [ Registro já reprovado por consistência anterior? ]
						if (_vFrete[iv].processo.Status != eFretePedidoViaCsvProcessoStatus.StatusInicial)
						{
							continue;
						}
						#endregion

						#region [ Consistências (regras que desqualificam o registro) ]

						#region [ Há nº NF? ]
						if (!_vFrete[iv].dadosRaw.hasNF)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.NUMERO_NF_NAO_INFORMADO;
							_vFrete[iv].processo.MensagemErro = "NF não informada";
							continue;
						}
						#endregion

						#region [ Nº NF válido? ]
						if (_vFrete[iv].dadosNormalizado.numNF == 0)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.NUMERO_NF_FORMATO_INVALIDO;
							_vFrete[iv].processo.MensagemErro = "NF informada é inválida (" + _vFrete[iv].dadosRaw.NF.Trim() + ")";
							continue;
						}
						#endregion

						#region [ CNPJ remetente válido? ]
						blnAchou = false;
						foreach (NFeEmitente emitente in listaNFeEmitente)
						{
							if (_vFrete[iv].dadosNormalizado.CnpjRemetente.Equals(emitente.cnpj))
							{
								blnAchou = true;
								break;
							}
						}

						if (!blnAchou)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.CNPJ_REMETENTE_DESCONHECIDO;
							_vFrete[iv].processo.MensagemErro = "CNPJ remetente desconhecido (" + _vFrete[iv].dadosRaw.CnpjRemetente.Trim() + ")";
							continue;
						}
						#endregion

						#region [ Tenta localizar o pedido através da NF ]
						// Como o arquivo não informa a série da NF, assume-se que é 1
						_vFrete[iv].dadosNormalizado.serieNF = 1;
						listaPedidos = FMain.contextoBD.AmbienteBase.pedidoDAO.getPedidoByNF(_vFrete[iv].dadosNormalizado.CnpjRemetente, _vFrete[iv].dadosNormalizado.serieNF, _vFrete[iv].dadosNormalizado.numNF, flagNaoCarregarItens: true);

						if (listaPedidos.Count == 0)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.PEDIDO_NAO_LOCALIZADO_POR_NF;
							_vFrete[iv].processo.MensagemErro = "Pedido não localizado através do nº NF";
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
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.MULTIPLOS_PEDIDOS_LOCALIZADOS_PARA_NF;
							_vFrete[iv].processo.MensagemErro = listaPedidos.Count().ToString() + " pedidos localizados para a NF: " + sPedidos;
							continue;
						}
						#endregion

						pedido = listaPedidos[0];

						#region [ Memoriza dados do pedido usados no processamento ]
						_vFrete[iv].processo.pedido = pedido;
						#endregion

						#region [ Transportadora válida? ]
						blnAchou = false;
						foreach (Transportadora transportadora in listaTransportadora)
						{
							if (_vFrete[iv].dadosNormalizado.TransportadoraCsv.ToUpper().Equals(transportadora.id))
							{
								blnAchou = true;
								_vFrete[iv].dadosNormalizado.TransportadoraCnpjCsv = Global.digitos(transportadora.cnpj);
								break;
							}
						}

						if (!blnAchou)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.TRANSPORTADORA_DESCONHECIDA;
							_vFrete[iv].processo.MensagemErro = "Transportadora desconhecida (" + _vFrete[iv].dadosRaw.TransportadoraCsv.Trim() + ")";
							continue;
						}
						#endregion

						#region [ Tipo de frete válido? ]
						if (_vFrete[iv].dadosNormalizado.TipoFreteCodigoSistema.Trim().Length == 0)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.TIPO_FRETE_DESCONHECIDO;
							_vFrete[iv].processo.MensagemErro = "Tipo de frete desconhecido (" + _vFrete[iv].dadosRaw.TipoFrete.Trim() + ")";
							continue;
						}
						#endregion

						#region [ Valor de frete válido? ]
						if (_vFrete[iv].dadosNormalizado.vlFrete <= 0)
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.VALOR_FRETE_INVALIDO;
							_vFrete[iv].processo.MensagemErro = "Valor de frete inválido (" + _vFrete[iv].dadosRaw.ValorFrete.Trim() + ")";
							continue;
						}
						#endregion

						#region [ Verifica status do campo 'st_entrega' ]
						if (pedido.st_entrega.Equals(Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
							_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.PEDIDO_ST_ENTREGA_INVALIDO;
							_vFrete[iv].processo.MensagemErro = "Pedido " + pedido.pedido + " possui status inválido: " + Global.stEntregaPedidoDescricao(pedido.st_entrega).ToUpper();
							continue;
						}
						#endregion

						#region [ Frete já registrado? ]
						if (pedido.listaPedidoFrete != null)
						{
							blnAchou = false;
							foreach (PedidoFrete freteCadastrado in pedido.listaPedidoFrete)
							{
								if ((freteCadastrado.vl_frete == _vFrete[iv].dadosNormalizado.vlFrete)
									&& (freteCadastrado.transportadora_id.Equals(_vFrete[iv].dadosNormalizado.TransportadoraCsv))
									&& (freteCadastrado.codigo_tipo_frete.Equals(_vFrete[iv].dadosNormalizado.TipoFreteCodigoSistema)))
								{
									blnAchou = true;
									_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.ErroInconsistencia;
									_vFrete[iv].processo.CodigoErro = eFretePedidoViaCsvProcessoCodigoErro.FRETE_JA_REGISTRADO;
									_vFrete[iv].processo.MensagemErro = "Frete já registrado anteriormente em " + Global.formataDataDdMmYyyyHhMmComSeparador(freteCadastrado.dt_hr_cadastro) + " por '" + freteCadastrado.usuario_cadastro + "'";
									break;
								}
							}
							if (blnAchou) continue;
						}
						#endregion

						#endregion

						// Por precaução, verifica se o registro está com erro desqualificante antes de prosseguir
						if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.ErroInconsistencia) continue;

						#region [ Consistências auxiliares (regras que não desqualificam o registro) ]

						#region [ Transportadora no CSV é diferente da transportadora no pedido? ]
						// Observação: essa divergência é permitida pois ocorre em casos de redespacho, por exemplo
						if (!_vFrete[iv].dadosNormalizado.TransportadoraCsv.ToUpper().Equals(pedido.transportadora_id.ToUpper()))
						{
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido;
							if (_vFrete[iv].processo.MensagemInformativa.Length > 0) _vFrete[iv].processo.MensagemInformativa += "\n";
							_vFrete[iv].processo.MensagemInformativa += "A transportadora no CSV é diferente da transportadora no pedido!";
						}
						#endregion

						#region [ Pedido já possui registrado um frete do mesmo tipo, mas c/ valor e/ou transportadora diferente? ]
						if (pedido.listaPedidoFrete != null)
						{
							foreach (PedidoFrete freteCadastrado in pedido.listaPedidoFrete)
							{
								if (freteCadastrado.codigo_tipo_frete.Equals(_vFrete[iv].dadosNormalizado.TipoFreteCodigoSistema))
								{
									_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido;
									if (_vFrete[iv].processo.MensagemInformativa.Length > 0) _vFrete[iv].processo.MensagemInformativa += "\n";
									_vFrete[iv].processo.MensagemInformativa += "Já consta frete do mesmo tipo no pedido: "
										+ Global.formataMoedaComSimbolo(freteCadastrado.vl_frete)
										+ " de '" + freteCadastrado.transportadora_id + "'"
										+ " cadastrado em " + Global.formataDataDdMmYyyyHhMmComSeparador(freteCadastrado.dt_hr_cadastro)
										+ " por '" + freteCadastrado.usuario_cadastro + "'";
								}
							}
						}
						#endregion

						// Não interrompe o processamento neste ponto, prossegue p/ incluir eventuais observações que possam existir
						#endregion

						#region [ Verificações auxiliares (regras que não desqualificam o registro) ]
						// Verifica se há observações a serem exibidas para que o operador fique atento a eventuais casos excepcionais

						#region [ Há dados de frete registrados anteriormente no pedido? ]
						if (pedido.listaPedidoFrete != null)
						{
							sbAux = new StringBuilder("");
							linhaFreteCadastrado = 0;
							foreach (PedidoFrete freteCadastrado in pedido.listaPedidoFrete)
							{
								if (freteCadastrado.vl_frete != 0)
								{
									linhaFreteCadastrado++;
									strMsg = MARGEM_LINHA_OBS + "(" + linhaFreteCadastrado.ToString() + ") "
											+ Global.formataMoedaComSimbolo(freteCadastrado.vl_frete)
											+ " de '" + freteCadastrado.transportadora_id + "'"
											+ " em " + Global.formataDataDdMmYyyyHhMmComSeparador(freteCadastrado.dt_hr_cadastro)
											+ " por '" + freteCadastrado.usuario_cadastro + "'"
											+ " do tipo '" + freteCadastrado.descricao_tipo_frete + "'";
									if (sbAux.Length > 0) sbAux.AppendLine("");
									sbAux.Append(strMsg);
								}
							}

							if (sbAux.Length > 0)
							{
								if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.StatusInicial) _vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.LiberadoComObsParaRegistrarFretePedido;
								strMsg = "Fretes registrados anteriormente no pedido:" + "\n" + sbAux.ToString();
								if (_vFrete[iv].processo.MensagemInformativa.Length > 0) _vFrete[iv].processo.MensagemInformativa += "\n";
								_vFrete[iv].processo.MensagemInformativa += strMsg;
							}
						}
						#endregion

						#endregion

						// Se chegou até este ponto, está apto para registrar o frete no pedido (com ressalvas)
						if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido) continue;

						// Se chegou até este ponto, está apto para registrar o frete no pedido (com observações sendo apontadas)
						if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComObsParaRegistrarFretePedido) continue;

						// Se chegou até este ponto, está apto para registrar o frete no pedido (sem ressalvas)
						_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.LiberadoParaRegistrarFretePedido;
					} // for (int iv = 0; iv < _vFrete.Count; iv++)
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
				for (int iv = 0; iv < _vFrete.Count; iv++)
				{
					if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.ErroInconsistencia)
					{
						#region [ 1ª posição: linhas com erros (meramente informativas, não há ação por parte do usuário) ]
						sOrdenacao = Global.normalizaCodigo("1", 2) + Global.normalizaCodigo(((int)_vFrete[iv].processo.CodigoErro).ToString(), 3) + Global.normalizaCodigo(_vFrete[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido)
					{
						#region [ 2ª posição: registrar frete no pedido (com ressalvas) ]
						sOrdenacao = Global.normalizaCodigo("2", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vFrete[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComObsParaRegistrarFretePedido)
					{
						#region [ 3ª posição: registrar frete no pedido (com observações) ]
						sOrdenacao = Global.normalizaCodigo("3", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vFrete[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoParaRegistrarFretePedido)
					{
						#region [ 4ª posição: registrar frete no pedido ]
						sOrdenacao = Global.normalizaCodigo("4", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vFrete[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}
					else
					{
						#region [ Situação desconhecida ]
						blnHaLinhasStatusDesconhecido = true;
						sOrdenacao = Global.normalizaCodigo("5", 2) + Global.normalizaCodigo("0", 3) + Global.normalizaCodigo(_vFrete[iv].dadosNormalizado.numNF.ToString(), 9);
						#endregion
					}

					_vFrete[iv].processo.campoOrdenacao = sOrdenacao;
				}

				vFreteOrdenado = _vFrete.OrderBy(o => o.processo.campoOrdenacao).ToList();
				#endregion

				#region [ Preenche o grid ]
				try
				{
					grid.SuspendLayout();
					grid.Rows.Add(vFreteOrdenado.Count);

					#region [ Mantém a exibição do grid sem nenhuma linha selecionada enquanto os dados são carregados ]
					for (int i = 0; i < grid.Rows.Count; i++)
					{
						if (grid.Rows[i].Selected) grid.Rows[i].Selected = false;
					}
					#endregion

					try
					{
						sNFBrancos = Global.normalizaCodigo("0", 9);

						percProgressoAnterior = -1;
						dtHrUltProgresso = DateTime.MinValue;

						for (int iv = 0; iv < vFreteOrdenado.Count; iv++)
						{
							lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
							percProgresso = 100 * (iv + 1) / vFreteOrdenado.Count;
							if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
							{
								strMsgProgresso = "Carregando dados no grid: linha " + (iv + 1).ToString() + " / " + vFreteOrdenado.Count.ToString() + "   (" + percProgresso.ToString() + "%)";
								info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
								percProgressoAnterior = percProgresso;
								Application.DoEvents();
								// Atualiza o horário após o DoEvents() para que o intervalo entre os DoEvents() não seja inferior ao mínimo definido
								dtHrUltProgresso = DateTime.Now;
							}

							grid.Rows[iv].Cells[GRID_COL_HIDDEN_GUID].Value = vFreteOrdenado[iv].processo.Guid;
							grid.Rows[iv].Cells[GRID_COL_VISIBLE_ORDENACAO_PADRAO].Value = (iv + 1).ToString() + ".";
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value = iv;

							grid.Rows[iv].Cells[GRID_COL_NF].Value = vFreteOrdenado[iv].dadosNormalizado.NF;
							sNF = Global.normalizaCodigo(vFreteOrdenado[iv].dadosNormalizado.NF, 9);
							if (sNF.Length == 0) sNF = sNFBrancos;
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_NF].Value = sNF + ' ' + Global.normalizaCodigo(iv.ToString(), 6);

							grid.Rows[iv].Cells[GRID_COL_REMETENTE].Value = Global.formataCnpjCpf(vFreteOrdenado[iv].dadosNormalizado.CnpjRemetente);
							grid.Rows[iv].Cells[GRID_COL_TRANSPORTADORA_CSV].Value = vFreteOrdenado[iv].dadosNormalizado.TransportadoraCsv;

							if (vFreteOrdenado[iv].processo.pedido != null)
							{
								grid.Rows[iv].Cells[GRID_COL_PEDIDO].Value = vFreteOrdenado[iv].processo.pedido.pedido;
								grid.Rows[iv].Cells[GRID_COL_TRANSPORTADORA_PEDIDO].Value = vFreteOrdenado[iv].processo.pedido.transportadora_id.ToUpper();
							}

							sVlFrete = Global.formataMoeda(vFreteOrdenado[iv].dadosNormalizado.vlFrete);
							grid.Rows[iv].Cells[GRID_COL_VL_FRETE].Value = sVlFrete;
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_VL_FRETE].Value = Global.normalizaCodigo(Global.digitos(sVlFrete), 18) + ' ' + Global.normalizaCodigo(iv.ToString(), 6);

							grid.Rows[iv].Cells[GRID_COL_TIPO_FRETE].Value = vFreteOrdenado[iv].dadosNormalizado.TipoFreteDescricaoSistema;
							grid.Rows[iv].Cells[GRID_COL_HIDDEN_TIPO_FRETE].Value = Global.normalizaCodigo(vFreteOrdenado[iv].dadosNormalizado.TipoFreteOrdenacaoSistema.ToString(), 12) + ' ' + Global.normalizaCodigo(iv.ToString(), 6);

							if (vFreteOrdenado[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.ErroInconsistencia)
							{
								qtdeRegErro++;
								grid.Rows[iv].Cells[GRID_COL_STATUS].Value = "ERRO";
								grid.Rows[iv].Cells[GRID_COL_MENSAGEM].Value = vFreteOrdenado[iv].processo.MensagemErro;
								grid.Rows[iv].DefaultCellStyle.ForeColor = Color.Red;
							}
							else if (vFreteOrdenado[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido)
							{
								qtdeRegApto++;
								grid.Rows[iv].Cells[GRID_COL_STATUS].Value = "ATENÇÃO";
								grid.Rows[iv].Cells[GRID_COL_MENSAGEM].Value = vFreteOrdenado[iv].processo.MensagemInformativa;
								grid.Rows[iv].DefaultCellStyle.ForeColor = Color.DarkViolet;
							}
							else if (vFreteOrdenado[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComObsParaRegistrarFretePedido)
							{
								qtdeRegApto++;
								grid.Rows[iv].Cells[GRID_COL_STATUS].Value = "OBS";
								grid.Rows[iv].Cells[GRID_COL_MENSAGEM].Value = vFreteOrdenado[iv].processo.MensagemInformativa;
								grid.Rows[iv].DefaultCellStyle.ForeColor = Color.Black;
							}
							else if (vFreteOrdenado[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoParaRegistrarFretePedido)
							{
								qtdeRegApto++;
								grid.Rows[iv].DefaultCellStyle.ForeColor = Color.Black;
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

				lblTotalRegistros.Text = Global.formataInteiro(_vFrete.Count);
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

		#region [ executaAnotaFretePedidoViaCsv ]
		private void executaAnotaFretePedidoViaCsv()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FAnotarFretePedidoViaCsv.executaAnotaFretePedidoViaCsv";
			int qtdeRegistrosParaInsertTotal = 0;
			int qtdeRegistrosParaInsertComRessalvas = 0;
			int qtdeRegistrosProcessados = 0;
			int qtdeRegistrosInsertSucesso = 0;
			int qtdeRegistrosInsertFalha = 0;
			int percProgresso;
			int percProgressoAnterior;
			long lngMiliSegundosDecorridos;
			bool blnFalhaInsert;
			bool blnRegistroLiberadoProcessamento;
			bool blnInsert;
			eFretePedidoViaCsvProcessoStatus statusProcessoOriginal;
			string strMsg;
			string strMsgProgresso;
			string strMsgErroLog = "";
			string msg_erro;
			StringBuilder sbLogSucesso = new StringBuilder("");
			StringBuilder sbLogFalha = new StringBuilder("");
			DateTime dtHrUltProgresso;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			PedidoFrete pedidoFrete;
			Log log = new Log();
			#endregion

			try
			{
				for (int iv = 0; iv < _vFrete.Count; iv++)
				{
					if ((_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoParaRegistrarFretePedido)
						|| (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido)
						|| (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComObsParaRegistrarFretePedido))

					{
						qtdeRegistrosParaInsertTotal++;
					}

					if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido)
					{
						qtdeRegistrosParaInsertComRessalvas++;
					}
				}

				if (qtdeRegistrosParaInsertTotal == 0)
				{
					avisoErro("Não há nenhum registro para ser atualizado no banco de dados!");
					return;
				}

				#region [ Solicita confirmação antes de executar a operação ]
				strMsg = "Confirma a atualização no banco de dados?";
				if (qtdeRegistrosParaInsertComRessalvas > 0) strMsg += "\nHá " + Global.formataInteiro(qtdeRegistrosParaInsertComRessalvas) + " registros com status de atenção!";
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
				for (int iv = 0; iv < _vFrete.Count; iv++)
				{
					blnRegistroLiberadoProcessamento = false;
					blnFalhaInsert = false;
					statusProcessoOriginal = _vFrete[iv].processo.Status;

					#region [ Registra o frete no pedido ]
					if ((statusProcessoOriginal == eFretePedidoViaCsvProcessoStatus.LiberadoParaRegistrarFretePedido)
						|| (statusProcessoOriginal == eFretePedidoViaCsvProcessoStatus.LiberadoComRessalvasParaRegistrarFretePedido)
						|| (statusProcessoOriginal == eFretePedidoViaCsvProcessoStatus.LiberadoComObsParaRegistrarFretePedido))
					{
						#region [ Prepara dados para gravação ]
						pedidoFrete = new PedidoFrete();
						pedidoFrete.pedido = _vFrete[iv].processo.pedido.pedido;
						pedidoFrete.codigo_tipo_frete = _vFrete[iv].dadosNormalizado.TipoFreteCodigoSistema;
						pedidoFrete.vl_frete = _vFrete[iv].dadosNormalizado.vlFrete;
						pedidoFrete.transportadora_id = _vFrete[iv].dadosNormalizado.TransportadoraCsv;
						pedidoFrete.transportadora_cnpj = _vFrete[iv].dadosNormalizado.TransportadoraCnpjCsv;
						pedidoFrete.id_nfe_emitente = _vFrete[iv].processo.pedido.id_nfe_emitente;
						pedidoFrete.serie_NF = _vFrete[iv].dadosNormalizado.serieNF;
						pedidoFrete.numero_NF = _vFrete[iv].dadosNormalizado.numNF;
						pedidoFrete.tipo_preenchimento = Global.Cte.FIN.T_PEDIDO_FRETE__TIPO_PREENCHIMENTO.ANOTACAO_VIA_CSV_ADM2;
						pedidoFrete.vl_NF = _vFrete[iv].processo.pedido.vl_total_NF_calculado_deste_pedido;
						pedidoFrete.emissor_cnpj = _vFrete[iv].dadosNormalizado.CnpjRemetente;
						pedidoFrete.id_editrp_arq_input_linha_processada_n1 = _vFrete[iv].dadosRaw.linhaArquivoCsv;
						#endregion

						#region [ Executa atualização no banco de dados ]
						blnRegistroLiberadoProcessamento = true;
						blnInsert = FMain.contextoBD.AmbienteBase.pedidoFreteDAO.InsertPedidoFreteViaCsv(pedidoFrete, Global.Usuario.usuario, out msg_erro);
						if (blnInsert)
						{
							qtdeRegistrosInsertSucesso++;
							strMsg = _vFrete[iv].processo.pedido.pedido
									+ " (NF=" + _vFrete[iv].dadosNormalizado.NF
									+ ", vl_frete=" + Global.formataMoeda(_vFrete[iv].dadosNormalizado.vlFrete)
									+ ", codigo_tipo_frete=" + _vFrete[iv].dadosNormalizado.TipoFreteCodigoSistema
									+ ", transportadora_id=" + _vFrete[iv].dadosNormalizado.TransportadoraCsv + ")";
							if (sbLogSucesso.Length > 0) sbLogSucesso.Append(", ");
							sbLogSucesso.Append(strMsg);
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.SucessoRegistroFretePedido;
						}
						else
						{
							strMsg = _vFrete[iv].processo.pedido.pedido
										+ " (NF=" + _vFrete[iv].dadosNormalizado.NF
										+ ", vl_frete=" + Global.formataMoeda(_vFrete[iv].dadosNormalizado.vlFrete)
										+ ", codigo_tipo_frete=" + _vFrete[iv].dadosNormalizado.TipoFreteCodigoSistema
										+ ", transportadora_id=" + _vFrete[iv].dadosNormalizado.TransportadoraCsv + ")";
							if (sbLogFalha.Length > 0) sbLogFalha.Append(", ");
							sbLogFalha.Append(strMsg);

							blnFalhaInsert = true;
							qtdeRegistrosInsertFalha++;
							_vFrete[iv].processo.Status = eFretePedidoViaCsvProcessoStatus.FalhaRegistroFretePedido;
							strMsg = "Falha ao tentar anotar o frete no pedido " + _vFrete[iv].processo.pedido.pedido
									+ " (NF: " + _vFrete[iv].dadosNormalizado.NF
									+ ", frete: " + Global.formataMoeda(_vFrete[iv].dadosNormalizado.vlFrete)
									+ ", tipo de frete: " + _vFrete[iv].dadosNormalizado.TipoFrete
									+ ", transportadora: " + _vFrete[iv].dadosNormalizado.TransportadoraCsv + ")";
							_vFrete[iv].processo.MensagemErro = strMsg;
							adicionaErro(strMsg);
						}
						#endregion
					}
					#endregion

					// Registro ignorado para processamento
					if (!blnRegistroLiberadoProcessamento) continue;

					#region [ Progresso ]
					qtdeRegistrosProcessados++;

					lngMiliSegundosDecorridos = Global.calculaTimeSpanMiliSegundos(DateTime.Now - dtHrUltProgresso);
					percProgresso = 100 * qtdeRegistrosProcessados / qtdeRegistrosParaInsertTotal;
					if ((percProgressoAnterior != percProgresso) && (lngMiliSegundosDecorridos >= MIN_INTERVALO_DOEVENTS_EM_MILISEGUNDOS))
					{
						strMsgProgresso = "Atualizando banco de dados: " + qtdeRegistrosProcessados.ToString() + " / " + qtdeRegistrosParaInsertTotal.ToString() + "   (" + percProgresso.ToString() + "%)";
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
						if (grid.Rows[jv].Cells[GRID_COL_HIDDEN_GUID].Value.ToString().Equals(_vFrete[iv].processo.Guid))
						{
							if (_vFrete[iv].processo.Status == eFretePedidoViaCsvProcessoStatus.SucessoRegistroFretePedido)
							{
								grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "OK";
								grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.Green;
							}
							else
							{
								grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "FALHA";
								grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.Red;
								grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Value = _vFrete[iv].processo.MensagemErro;
								grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Style.ForeColor = Color.Red;
							}

							break;
						}
					}
					#endregion

					if (blnFalhaInsert)
					{
						// Prossegue para o próximo registro
						continue;
					}

					strMsg = "Sucesso na anotação do frete no pedido " + _vFrete[iv].processo.pedido.pedido
							+ " (NF: " + _vFrete[iv].dadosNormalizado.NF
								+ ", frete: " + Global.formataMoeda(_vFrete[iv].dadosNormalizado.vlFrete)
								+ ", tipo de frete: " + _vFrete[iv].dadosNormalizado.TipoFrete
								+ ", transportadora: " + _vFrete[iv].dadosNormalizado.TransportadoraCsv + ")";
					adicionaDisplay(strMsg);
				} // for (int iv = 0; iv < _vFrete.Count; iv++)

				lblQtdeAtualizSucesso.Text = Global.formataInteiro(qtdeRegistrosInsertSucesso);
				lblQtdeAtualizFalha.Text = Global.formataInteiro(qtdeRegistrosInsertFalha);

				tsDuracaoProcessamento = DateTime.Now - dtInicioProcessamento;

				#region [ Grava o log ]
				strMsg = "[Módulo ADM2] Operação 'Anotar Frete no Pedido via CSV':"
						+ "\nSucesso: " + Global.formataInteiro(qtdeRegistrosInsertSucesso) + " registro(s) (" + (sbLogSucesso.Length > 0 ? sbLogSucesso.ToString() : "vazio") + ")"
						+ "\nFalha: " + Global.formataInteiro(qtdeRegistrosInsertFalha) + " registro(s) (" + (sbLogFalha.Length > 0 ? sbLogFalha.ToString() : "vazio") + ")"
						+ "\nDuração do processamento: " + Global.formataDuracaoHMS(tsDuracaoProcessamento)
						+ "\nArquivo processado: " + txtArquivoCsv.Text.Trim() + " (contendo " + Global.formataInteiro(_vFrete.Count) + " registros)";
				log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_ANOTA_FRETE_PEDIDO_VIA_CSV;
				log.usuario = Global.Usuario.usuario;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

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

		#region [ FAnotarFretePedidoViaCsv ]

		#region [ FAnotarFretePedidoViaCsv_Load ]
		private void FAnotarFretePedidoViaCsv_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				txtArquivoCsv.Text = "";
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

		#region [ FAnotarFretePedidoViaCsv_Shown ]
		private void FAnotarFretePedidoViaCsv_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					openFileDialogCtrl.InitialDirectory = pathArquivoCsvValorDefault();
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

		#region [ FAnotarFretePedidoViaCsv_FormClosing ]
		private void FAnotarFretePedidoViaCsv_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (txtArquivoCsv.Text.Length > 0)
			{
				if (!confirma("Sair do painel?"))
				{
					e.Cancel = true;
					return;
				}
			}

			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaArquivoCsv ]

		#region [ btnSelecionaArquivoCsv_Click ]
		private void btnSelecionaArquivoCsv_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArquivoCsv();
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
				case GRID_COL_VL_FRETE:
					sValue1 = grid.Rows[e.RowIndex1].Cells[GRID_COL_HIDDEN_VL_FRETE].Value.ToString();
					sValue2 = grid.Rows[e.RowIndex2].Cells[GRID_COL_HIDDEN_VL_FRETE].Value.ToString();
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
