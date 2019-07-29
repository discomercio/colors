#region [ using ]
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FArqRetorno : FModelo
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

		ArqRemessaRetorno _arqRemessaRetorno;
		ArqRemessaRetorno.MensagemProcessamento _msgProcRemessa;
		ArqRemessaRetorno.MensagemProcessamento _msgProcRemessaTotalizador;
		ArqRemessaRetorno.RelatorioTotalizacao _relTotal;
		ArqRemessaRetorno.TabelaErro _tabErro;
		ArqRemessa.LinhaHeader _linhaHeader;
		ArqRemessa.DetalheTitulo _detalheTitulo;
		ArqRemessaRetorno.RegistroDetalhe _regDetalhe;
		ArqRemessaRetorno.TotalizadorCliente _totCliente;
		ArqRemessaRetorno.TotalizadorPagamento _totPagto;
		#endregion

		#region [ Constantes ]
		private const int ST_PROCESSAMENTO_EM_ANDAMENTO = 1;
		private const int ST_PROCESSAMENTO_SUCESSO = 2;
		private const int ST_PROCESSAMENTO_FALHA = 3;
		private const int ZERO_SEGUNDO = 0;
		private const int ST_RETORNO_SERASA_NAO_PROCESSADO = 0;
		private const int ST_RETORNO_SERASA_PROCESSADO = 1;
		private const int ST_PROCESSADO_SERASA_FALHA = 0;
		private const int ST_PROCESSADO_SERASA_SUCESSO = 1;
		#endregion

		#region [ Construtor ]
		public FArqRetorno()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ adicionaDisplay ]
		private void adicionaDisplay(String mensagem)
		{
			String strMensagem;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			lbMensagem.Items.Add(strMensagem.Replace('\n', ' '));
			lbMensagem.SelectedIndex = lbMensagem.Items.Count - 1;
			Global.gravaLogAtividade(mensagem);
		}
		#endregion

		#region [ adicionaErro ]
		private void adicionaErro(String mensagem)
		{
			String strMensagem;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			lbErro.Items.Add(strMensagem.Replace('\n', ' '));
			lbErro.SelectedIndex = lbErro.Items.Count - 1;
			Global.gravaLogAtividade("ERRO: " + mensagem);
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
			if (Global.Usuario.Defaults.FArqRetorno.pathTituloArquivoRetorno.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FArqRetorno.pathTituloArquivoRetorno))
				{
					strResp = Global.Usuario.Defaults.FArqRetorno.pathTituloArquivoRetorno;
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
			lbMensagem.Items.Clear();
			lbErro.Items.Clear();
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

		#region [ verificaSeArquivoJaFoiCarregadoAntes ]
		private bool verificaSeArquivoJaFoiCarregadoAntes(string[] linhasArquivo)
		{
			#region [ Declaração ]
			string linha_header = "";
			string s_data_inicial;
			string s_data_final;
			string cnpj_empresa;
			bool arquivoJaCarregado = true;
			#endregion

			for (int i = 0; i < linhasArquivo.Length; i++)
			{
				if (Texto.leftStr(linhasArquivo[i], 2).Equals("00"))
				{
					linha_header = linhasArquivo[i];
					break;
				}
			}

			if (linha_header.Trim().Length == 0) throw new Exception("Arquivo com formato inválido: não foi encontrado o header do arquivo!!");

			s_data_inicial = linha_header.Substring(36, 8);
			s_data_final = linha_header.Substring(44, 8);
			cnpj_empresa = linha_header.Substring(22, 14);

			#region [ Verifica se o arquivo selecionado já foi carregado anteriormente ]
			arquivoJaCarregado = ArqRetornoDAO.verificaSeArquivoJaFoiCarregadoAntes(Path.GetFileName(txtArqRetorno.Text), s_data_inicial, s_data_final, cnpj_empresa);
			#endregion

			return arquivoJaCarregado;
		}
		#endregion

		#region [ carregaGridBoletos ]
		private void carregaGridBoletos()
		{
			#region [ Declarações ]
			String[] linhasArqRetorno;
			String strNomeArquivoCompleto;
			String strDtInicio;
			String strDtFim;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			String linha;
			List<ArqRemessaRetorno.RelatorioTotalizacao> listaRelTotal = new List<ArqRemessaRetorno.RelatorioTotalizacao>();
			#endregion

			strNomeArquivoCompleto = txtArqRetorno.Text;

			info(ModoExibicaoMensagemRodape.EmExecucao, "lendo o arquivo de retorno");
			try
			{
				#region [ Limpa campos ]
				grdBoletos.Rows.Clear();
				lblTotalRegistros.Text = "";
				#endregion

				#region [ Carrega dados do arquivo em array ]
				linhasArqRetorno = File.ReadAllLines(strNomeArquivoCompleto, encode);
				#endregion

				#region [ Processa Mensagem Processamento Remessa ]
				linha = linhasArqRetorno[0];
				_msgProcRemessa = new ArqRemessaRetorno.MensagemProcessamento();
				if (Texto.leftStr(linha, 2).Equals("77")) _msgProcRemessa.carrega(linha);
				#endregion

				#region [ Processa Texto Relatorio Tot. Remessa ]
				int i = 1;
				for (; i < linhasArqRetorno.Length; i++)
				{
					linha = linhasArqRetorno[i];
					if (!Texto.leftStr(linha, 2).Equals("85")) break;
					_relTotal = new ArqRemessaRetorno.RelatorioTotalizacao();
					_relTotal.carrega(linha);
					listaRelTotal.Add(_relTotal);
				}
				#endregion

				#region [ Processa a Tabela de Erros ]
				for (; i < linhasArqRetorno.Length; i++)
				{
					linha = linhasArqRetorno[i];
					if (!Texto.leftStr(linha, 2).Equals("88")) break;
					_tabErro = new ArqRemessaRetorno.TabelaErro();
					_tabErro.carrega(linha);
					_arqRemessaRetorno.adicionaRegistroErro(_tabErro);
				}
				#endregion

				#region [ Header do arquivo de remessa ]
				if (i < linhasArqRetorno.Length)
				{
					linha = linhasArqRetorno[i];
					if (Texto.leftStr(linha, 2).Equals("00"))
					{
						DateTime dtInicio = DateTime.MinValue;
						strDtInicio = linha.Substring(36, 8);
						if (strDtInicio.Equals("CONCILIA"))
						{
							throw new Exception("O arquivo selecionado é de conciliação!!");
						}
						strDtInicio = Global.digitos(strDtInicio);
						if (strDtInicio.Length == 8) dtInicio = Global.converteYyyyMmDdSemSeparadorParaDateTime(strDtInicio);

						DateTime dtFim = DateTime.MinValue;
						strDtFim = linha.Substring(44, 8);
						strDtFim = Global.digitos(strDtFim);
						if (strDtFim.Length == 8) dtFim = Global.converteYyyyMmDdSemSeparadorParaDateTime(strDtFim);

						_linhaHeader = new ArqRemessa.LinhaHeader(dtInicio, dtFim);
						i++;
					}
				}
				#endregion

				#region [ Processa registros da remessa ]
				for (; i < linhasArqRetorno.Length; i++)
				{
					linha = linhasArqRetorno[i];

					#region [  Registro de tempo de relacionamento com cliente? ]
					if (Texto.leftStr(linha, 2).Equals("01"))
					{
						if (linha.Substring(16, 2).Equals("01")) continue;
					}
					#endregion

					// Chegou na mensagem do totalizador?
					if (Texto.leftStr(linha, 2).Equals("77")) break;

					// Chegou no 2º header?
					if (Texto.leftStr(linha, 2).Equals("00")) break;

					// Chegou no totalizador de pagamentos?
					if (Texto.leftStr(linha, 2).Equals("05")) break;

					// Alcançou a seção de totalizador de clientes?
					if (Texto.leftStr(linha, 2).Equals("01"))
					{
						if (linha.Length <= 8) break;
						if (linha.Substring(8, 1).Equals(" ")) break;
					}

					// Acabaram os registros da remessa?
					if (!Texto.leftStr(linha, 2).Equals("01")) break;

					#region [ Armazena registro da remessa ]
					_detalheTitulo = new ArqRemessa.DetalheTitulo();
					_detalheTitulo.carrega(linha);

					String erros = "";
					if (linha.Length > 130)
					{
						if (linha.Length >= 220)
						{
							erros = linha.Substring(130, 90);
						}
						else
						{
							erros = linha.Substring(130);
						}
					}

					_regDetalhe = new ArqRemessaRetorno.RegistroDetalhe(_detalheTitulo, erros);
					_arqRemessaRetorno.adicionaRegistroDetalhe(_regDetalhe);
					#endregion
				}
				#endregion

				#region [ Mensagem do totalizador? ]
				_msgProcRemessaTotalizador = new ArqRemessaRetorno.MensagemProcessamento();
				if (i < linhasArqRetorno.Length)
				{
					linha = linhasArqRetorno[i];
					if (Texto.leftStr(linha, 2).Equals("77"))
					{
						_msgProcRemessaTotalizador.carrega(linha);
						i++;
					}
				}
				#endregion

				#region [ 2º Header? ]
				if (i < linhasArqRetorno.Length)
				{
					linha = linhasArqRetorno[i];
					if (Texto.leftStr(linha, 2).Equals("00"))
					{
						i++;
					}
				}
				#endregion

				#region [ Processa Totalizador de Clientes e Pagamentos ]
				for (; i < linhasArqRetorno.Length; i++)
				{
					linha = linhasArqRetorno[i];
					String idLinha = Texto.leftStr(linha, 2);

					#region [ Totalizador de Clientes ]
					if (idLinha.Equals("01"))
					{
						_totCliente = new ArqRemessaRetorno.TotalizadorCliente();
						_totCliente.carrega(linha);
					}
					#endregion

					#region [ Totalizador de Pagamentos ]
					if (idLinha.Equals("05"))
					{
						_totPagto = new ArqRemessaRetorno.TotalizadorPagamento();
						_totPagto.carrega(linha);
					}
					#endregion
				}
				#endregion

				#region [ Preenche grid ]
				if (_arqRemessaRetorno.registrosDetalhe.Count > 0) grdBoletos.Rows.Add(_arqRemessaRetorno.registrosDetalhe.Count);

				i = 0;
				foreach (ArqRemessaRetorno.RegistroDetalhe registro in _arqRemessaRetorno.registrosDetalhe)
				{
					String numBoleto = registro.linhaDetalhe.numeroTitulo.Trim();
					String numBoletoFormatado = numBoleto.Insert(numBoleto.Length - 1, "-");
					grdBoletos.Rows[i].Cells["numero_boleto"].Value = numBoletoFormatado;
					grdBoletos.Rows[i].Cells["data_emissao"].Value = Global.formataDataDdMmYyyyComSeparador(registro.linhaDetalhe.dataEmissao);
					grdBoletos.Rows[i].Cells["data_vencimento"].Value = Global.formataDataDdMmYyyyComSeparador(registro.linhaDetalhe.dataVencimento);
					if (registro.linhaDetalhe.valorTitulo == 99999999999.99m)
					{
						grdBoletos.Rows[i].Cells["valor"].Value = "";
					}
					else
					{
						grdBoletos.Rows[i].Cells["valor"].Value = Global.formataMoeda(registro.linhaDetalhe.valorTitulo);
					}

					if (registro.linhaDetalhe.dataPagamento == DateTime.MinValue)
					{
						grdBoletos.Rows[i].Cells["data_pagamento"].Value = "";
					}
					else
					{
						grdBoletos.Rows[i].Cells["data_pagamento"].Value = Global.formataDataDdMmYyyyComSeparador(registro.linhaDetalhe.dataPagamento);
					}

					String linhaCodErros = registro.erros.Trim();
					if (linhaCodErros.Length > 0)
					{
						grdBoletos.Rows[i].Cells["ocorrencia"].Value = _arqRemessaRetorno.getMsgErro(linhaCodErros.Substring(0, 3)); //mostra a primeira ocorrencia
					}
					else
					{
						grdBoletos.Rows[i].Cells["ocorrencia"].Value = "";
					}

					i++;
				}
				#endregion

				lblTotalRegistros.Text = Global.formataInteiro(i);
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
			#region [ Declarações ]
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
				Global.Usuario.Defaults.FArqRetorno.pathTituloArquivoRetorno = Path.GetDirectoryName(openFileDialog.FileName);

				#region [ Carrega dados do arquivo em array ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "lendo dados do arquivo");
				string[] _linhasArqRetorno = File.ReadAllLines(txtArqRetorno.Text, encode);
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

				#region [ Consistência do Header ]
				String linhaHeader = _linhasArqRetorno[0];
				String arquivoId = linhaHeader.Substring(0, 2);
				if (!arquivoId.Equals("77"))
				{
					avisoErro("O arquivo possui header inválido!!");
					return;
				}
				#endregion

				#region [ Verifica se o usuário carregou o arquivo anteriormente ]
				if (verificaSeArquivoJaFoiCarregadoAntes(_linhasArqRetorno))
				{
					avisoErro("O arquivo selecionado já foi carregado anteriormente!!");
					txtArqRetorno.Text = "";
					return;
				}
				#endregion

				#region [ Prepara um novo objeto para carregar o arquivo de retorno ]
				_arqRemessaRetorno = new ArqRemessaRetorno();
				#endregion

				carregaGridBoletos();
				grdBoletos.Focus();
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

		#region [ trataBotaoCarregaArqRetorno ]
		private void trataBotaoCarregaArqRetorno()
		{
			#region [ Declarações ]
			int idArqRetornoNormal = 0;
			bool blnGerouNsu;
			bool blnSucesso;
			String strMsgErro = "";
			String strNomeArquivoCompleto;
			String strNomeArquivoCompletoRenomeado;
			String strNomeArquivoCompletoRenomeadoAux;
			String[] linhasArqRetorno;
			DateTime dtInicioProcessamento;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			#region [ Obtém nome do arquivo de retorno ]
			strNomeArquivoCompleto = txtArqRetorno.Text;
			#endregion

			#region [ Consistência ]
			if (strNomeArquivoCompleto.Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo de retorno a ser carregado!!");
				return;
			}
			if (!File.Exists(strNomeArquivoCompleto))
			{
				avisoErro("O arquivo de retorno selecionado não existe!!\n\n" + strNomeArquivoCompleto);
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma a carga do arquivo de retorno?")) return;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "lendo o arquivo de retorno");
			try
			{
				dtInicioProcessamento = DateTime.Now;

				#region [ Carrega dados do arquivo em array ]
				adicionaDisplay("Leitura dos registros do arquivo de retorno " + Path.GetFileName(strNomeArquivoCompleto));
				linhasArqRetorno = File.ReadAllLines(strNomeArquivoCompleto, encode);
				adicionaDisplay("Registros para processar: " + Global.formataInteiro(linhasArqRetorno.Length - 2));
				#endregion

				#region [ Verifica se o usuário carregou o arquivo anteriormente ]
				if (verificaSeArquivoJaFoiCarregadoAntes(linhasArqRetorno))
				{
					avisoErro("O arquivo selecionado já foi carregado anteriormente!!");
					txtArqRetorno.Text = "";
					return;
				}
				#endregion

				blnSucesso = false;
				try
				{
					BD.iniciaTransacao();

					info(ModoExibicaoMensagemRodape.EmExecucao, "processando o arquivo de retorno");

					#region [ Gera o NSU para o novo registro que será gravado em t_SERASA_ARQ_RETORNO_NORMAL ]
					blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_ARQ_RETORNO_NORMAL, ref idArqRetornoNormal, ref strMsgErro);
					if (!blnGerouNsu)
					{
						throw new Exception("Falha ao tentar gerar o NSU para o registro de histórico de arquivos de retorno!!\n" + strMsgErro);
					}
					#endregion

					#region [ Insere Registro na Tabela t_SERASA_ARQ_RETORNO_NORMAL ]
					if (!ArqRetornoDAO.insere(idArqRetornoNormal,
											  DateTime.Now,
											  DateTime.Now,
											  Global.Usuario.usuario,
											  ArqRemessa.LinhaHeader.CNPJ_EMPRESA_CONVENIADA,
											  Global.formataDataYyyyMmDdSemSeparador(_linhaHeader.dataInicio),
											  _linhaHeader.dataInicio,
											  Global.formataDataYyyyMmDdSemSeparador(_linhaHeader.dataFim),
											  _linhaHeader.dataFim,
											  ArqRemessa.LinhaHeader.PERIODICIDADE_REMESSA,
											  ArqRemessa.LinhaHeader.RESERVADO_SERASA,
											  ArqRemessa.LinhaHeader.ID_GRUPO_RELATO_SEGMENTO,
											  ArqRemessa.LinhaHeader.ID_VERSAO_LAYOUT,
											  ArqRemessa.LinhaHeader.NUM_VERSAO_LAYOUT,
											  _arqRemessaRetorno.getTotalRegistros(),
											  _arqRemessaRetorno.getTotalRegistrosSemRejeicao(),
											  _arqRemessaRetorno.getTotalRegistrosRejeitados(),
											  ZERO_SEGUNDO,
											  Path.GetFileName(strNomeArquivoCompleto),
											  Path.GetDirectoryName(strNomeArquivoCompleto),
											  ST_PROCESSAMENTO_EM_ANDAMENTO,
											  null))
					{
						throw new Exception("Falha na criação de registro para o arquivo de retorno!!");
					}
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					strMsgErro = ex.ToString();
					blnSucesso = false;
				}

				if (!blnSucesso)
				{
					BD.rollbackTransacao();

					strMsgErro = "Falha na carga do arquivo de retorno!!\n\n" + strMsgErro;
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
				}
				else
				{
					BD.commitTransacao();

					try
					{
						BD.iniciaTransacao();

						int idRetNormalTitulo = 0;

						foreach (ArqRemessaRetorno.RegistroDetalhe registro in _arqRemessaRetorno.registrosDetalhe)
						{
							blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_RETORNO_NORMAL_TITULO, ref idRetNormalTitulo, ref strMsgErro);
							if (!blnGerouNsu)
							{
								throw new Exception("Falha ao tentar gerar o NSU para o registro do título do arquivo de retorno!!\n" + strMsgErro);
							}

							String dtPagto = "";
							if (registro.linhaDetalhe.dataPagamento == DateTime.MinValue)
							{
								dtPagto = new String(' ', 8);
							}
							else
							{
								dtPagto = Global.formataDataYyyyMmDdSemSeparador(registro.linhaDetalhe.dataPagamento);
							}

							if (!RetNormalTituloDAO.insere(idRetNormalTitulo,
													  idArqRetornoNormal,
													  ArqRemessa.DetalheTitulo.ID,
													  registro.linhaDetalhe.cnpjSacado,
													  ArqRemessa.DetalheTitulo.TIPO_DADOS,
													  registro.linhaDetalhe.numeroTitulo.Substring(0, 10),
													  Global.formataDataYyyyMmDdSemSeparador(registro.linhaDetalhe.dataEmissao),
													  Global.formataMoedaSemSeparador(registro.linhaDetalhe.valorTitulo, 13),
													  Global.formataDataYyyyMmDdSemSeparador(registro.linhaDetalhe.dataVencimento),
													  dtPagto,
													  "#D",
													  registro.linhaDetalhe.numeroTitulo,
													  registro.erros))
							{
								throw new Exception("Falha na criação de registro para o título do arquivo de retorno!!");
							}

							if (!TituloMovimentoDAO.atualizaTituloAposRetornoArquivo(ST_RETORNO_SERASA_PROCESSADO,
																					idArqRetornoNormal,
																					ST_PROCESSADO_SERASA_FALHA,
																					registro.erros,
																					registro.linhaDetalhe.numeroTitulo))
							{
								throw new Exception("Falha na atualização do status do titulo do arquivo de retorno!!");
							}
						}

						int idTabErros = 0;
						foreach (ArqRemessaRetorno.TabelaErro erro in _arqRemessaRetorno.registrosErro)
						{
							blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_RETORNO_NORMAL_TAB_ERROS, ref idTabErros, ref strMsgErro);
							if (!blnGerouNsu)
							{
								throw new Exception("Falha ao tentar gerar o NSU para o registro de código de erros do arquivo de retorno!!\n" + strMsgErro);
							}

							if (!TabErrosDAO.insere(idTabErros,
											   idArqRetornoNormal,
											   erro.numeroMsg,
											   erro.descricao))
							{
								throw new Exception("Falha na criação de registro para o código de erro do arquivo de retorno!!");
							}
						}

						if (!ArqRetornoDAO.atualizaDuracaoProcessamento(DateTime.Now.Subtract(dtInicioProcessamento).Seconds,
																							  idArqRetornoNormal))
						{
							throw new Exception("Falha na atualização da duração de processamento!!");
						}

						if (!ArqRetornoDAO.atualizaStatusProcessamento(ST_PROCESSAMENTO_SUCESSO,
																	null,
																	idArqRetornoNormal))
						{
							throw new Exception("Falha na atualização do status final do processamento do arquivo de retorno!!");
						}

						blnSucesso = true;
					}
					catch (Exception e)
					{
						Global.gravaLogAtividade(e.ToString());
						strMsgErro = e.ToString();
						blnSucesso = false;
					}

					if (blnSucesso)
					{
						BD.commitTransacao();

						adicionaDisplay("Arquivo de retorno carregado com sucesso!!");

						#region [ Renomeia o arquivo de retorno ]
						strNomeArquivoCompletoRenomeado = strNomeArquivoCompleto + ".PRC";
						if (File.Exists(strNomeArquivoCompletoRenomeado))
						{
							int intFileIndex = 0;
							while (File.Exists(strNomeArquivoCompletoRenomeado))
							{
								intFileIndex++;
								strNomeArquivoCompletoRenomeadoAux = strNomeArquivoCompletoRenomeado + ".OLD." + intFileIndex.ToString().PadLeft(3, '0');
								if (!File.Exists(strNomeArquivoCompletoRenomeadoAux)) File.Move(strNomeArquivoCompletoRenomeado, strNomeArquivoCompletoRenomeadoAux);
							}
						}
						File.Move(strNomeArquivoCompleto, strNomeArquivoCompletoRenomeado);
						#endregion

						adicionaDisplay("Arquivo de retorno renomeado para " + Path.GetFileName(strNomeArquivoCompletoRenomeado));

						info(ModoExibicaoMensagemRodape.Normal);
						aviso("Arquivo de retorno carregado com sucesso!!\n\n" + strNomeArquivoCompleto);
					}
					else
					{
						BD.rollbackTransacao();

						try
						{
							BD.iniciaTransacao();
							if (!ArqRetornoDAO.atualizaStatusProcessamento(ST_PROCESSAMENTO_FALHA,
																			"Falha no processamento do arquivo",
																			idArqRetornoNormal))
							{
								throw new Exception("Falha na atualização do status de erro do arquivo de retorno!!");
							}
							BD.commitTransacao();
						}
						catch (Exception e)
						{
							Global.gravaLogAtividade(e.ToString());
							BD.rollbackTransacao();
						}

						strMsgErro = "Falha na carga do arquivo de retorno!!\n\n" + strMsgErro;
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
					}
				}
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
		#region [ FArqRetorno ]
		#region [ FArqRetorno_Load ]
		private void FArqRetorno_Load(object sender, EventArgs e)
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

		#region [ FArqRetorno_Shown ]
		private void FArqRetorno_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					openFileDialog.InitialDirectory = pathBoletoArquivoRetornoValorDefault();

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

		#region [ FArqRetorno_FormClosing ]
		private void FArqRetorno_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain._fMain.Location = this.Location;
			FMain._fMain.Visible = true;
			this.Visible = false;
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

		#region [ btnCarregaArqRetorno ]
		#region [ btnCarregaArqRetorno_Click ]
		private void btnCarregaArqRetorno_Click(object sender, EventArgs e)
		{
			trataBotaoCarregaArqRetorno();
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
		#endregion
	}
}
