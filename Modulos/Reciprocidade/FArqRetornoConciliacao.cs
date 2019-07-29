#region [ using ]
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FArqRetornoConciliacao : FModelo
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

		List<ArqRemessa.DetalheTitulo> _titulos;
		ArqRemessa.DetalheTitulo _detalheTitulo;
		String _periodoFinalHeader;
		#endregion

		#region [ Constantes ]
		private const int ST_PROCESSAMENTO_EM_ANDAMENTO = 1;
		private const int ST_PROCESSAMENTO_SUCESSO = 2;
		private const int ST_PROCESSAMENTO_FALHA = 3;
		private const int ZERO_SEGUNDO = 0;
		private const int ST_RETORNO_SERASA_NAO_PROCESSADO = 0;
		private const int ST_RETORNO_SERASA_PROCESSADO = 1;
		private const int ST_TITULO_TRATADO_MANUAL_NAO = 0;
		private const int ST_TITULO_TRATADO_MANUAL_SIM = 1;
		private const int ID_ARQ_CONC_OUTPUT_ZERO = 0;
		private const int VL_TITULO_EDITADO_ZERO = 0;
		private const int ST_ENVIADO_SERASA_NAO_ENVIADO = 0;
		#endregion

		#region [ Construtor ]
		public FArqRetornoConciliacao()
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
			if (Global.Usuario.Defaults.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao))
				{
					strResp = Global.Usuario.Defaults.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao;
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
			string header = linhasArquivo[0];
			string s_data_final = header.Substring(44, 8);
			string cnpj_empresa = header.Substring(22, 14);
			bool arquivoJaCarregado = true;
			#endregion

			#region [ Verifica se o arquivo selecionado já foi carregado anteriormente ]
			arquivoJaCarregado = ArqConciliacaoInputDAO.verificaSeArquivoJaFoiCarregadoAntes(Path.GetFileName(txtArqRetorno.Text), s_data_final, cnpj_empresa);
			#endregion

			return arquivoJaCarregado;
		}
		#endregion

		#region [ carregaGridBoletos ]
		private void carregaGridBoletos()
		{
			#region [ Declarações ]
			int qtdeRegProcessado;
			int percProgressoAtual;
			int percProgressoAnterior;
			String[] linhasArqConciliacao;
			String strMsgProgresso;
			String strNomeArquivoCompleto;
			int intLinhaGrid = 0;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			String linha;
			#endregion

			strNomeArquivoCompleto = txtArqRetorno.Text;

			info(ModoExibicaoMensagemRodape.EmExecucao, "lendo o arquivo de conciliação");
			try
			{
				#region [ Limpa campos ]
				grdBoletos.Rows.Clear();
				lblTotalRegistros.Text = "";
				#endregion

				#region [ Carrega dados do arquivo em array ]
				linhasArqConciliacao = File.ReadAllLines(strNomeArquivoCompleto, encode);
				#endregion

				#region [ Consiste o header do arquivo ]
				linha = linhasArqConciliacao[0];
				String strConcilia = linha.Substring(36, 8);
				if (!strConcilia.Equals("CONCILIA"))
				{
					avisoErro("O arquivo selecionado não é de conciliação!!");
					return;
				}
				#endregion

				#region [ Verifica se o usuário carregou o arquivo anteriormente ]
				if (verificaSeArquivoJaFoiCarregadoAntes(linhasArqConciliacao))
				{
					avisoErro("O arquivo selecionado já foi carregado anteriormente!!");
					txtArqRetorno.Text = "";
					return;
				}
				#endregion

				#region [ Obtem o periodo final do header ]
				_periodoFinalHeader = linha.Substring(44, 8);
				#endregion

				#region [ Processa registros da remessa ]
				qtdeRegProcessado = 0;
				percProgressoAnterior = 0;
				for (int i = 1; i < linhasArqConciliacao.Length; i++)
				{
					qtdeRegProcessado++;
					percProgressoAtual = 100 * qtdeRegProcessado / linhasArqConciliacao.Length;
					if (percProgressoAtual != percProgressoAnterior)
					{
						percProgressoAnterior = percProgressoAtual;
						strMsgProgresso = "Processando dados do arquivo: " + percProgressoAtual.ToString() + "%";
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
						Application.DoEvents();
					}

					linha = linhasArqConciliacao[i];

					if ((linha.Substring(0, 2).Equals("99")))
					{
						break; //Alcançou o trailler do arquivo
					}

					_detalheTitulo = new ArqRemessa.DetalheTitulo();
					_detalheTitulo.carrega(linha);
					_titulos.Add(_detalheTitulo);
				}
				#endregion

				#region [ Preenche grid ]
				if (_titulos.Count > 0) grdBoletos.Rows.Add(_titulos.Count);

				qtdeRegProcessado = 0;
				percProgressoAnterior = 0;
				foreach (ArqRemessa.DetalheTitulo titulo in _titulos)
				{
					qtdeRegProcessado++;
					percProgressoAtual = 100 * qtdeRegProcessado / _titulos.Count;
					if (percProgressoAtual != percProgressoAnterior)
					{
						percProgressoAnterior = percProgressoAtual;
						strMsgProgresso = "Carregando dados no grid: " + percProgressoAtual.ToString() + "%";
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
						Application.DoEvents();
					}

					String numBoleto = titulo.numeroTitulo.Trim();
					String numBoletoFormatado = numBoleto.Insert(numBoleto.Length - 1, "-");
					grdBoletos.Rows[intLinhaGrid].Cells["numero_boleto"].Value = numBoletoFormatado;
					grdBoletos.Rows[intLinhaGrid].Cells["data_emissao"].Value = Global.formataDataDdMmYyyyComSeparador(titulo.dataEmissao);
					grdBoletos.Rows[intLinhaGrid].Cells["data_vencimento"].Value = Global.formataDataDdMmYyyyComSeparador(titulo.dataVencimento);
					grdBoletos.Rows[intLinhaGrid].Cells["valor"].Value = Global.formataMoeda(titulo.valorTitulo);

					if (titulo.dataPagamento == DateTime.MinValue)
					{
						grdBoletos.Rows[intLinhaGrid].Cells["data_pagamento"].Value = "";
					}
					else
					{
						grdBoletos.Rows[intLinhaGrid].Cells["data_pagamento"].Value = Global.formataDataDdMmYyyyComSeparador(titulo.dataPagamento);
					}

					intLinhaGrid++;
				}
				#endregion

				#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
				for (int i = 0; i < grdBoletos.Rows.Count; i++)
				{
					if (grdBoletos.Rows[i].Selected) grdBoletos.Rows[i].Selected = false;
				}
				#endregion

				lblTotalRegistros.Text = Global.formataInteiro(intLinhaGrid);
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
				Global.Usuario.Defaults.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao = Path.GetDirectoryName(openFileDialog.FileName);

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

				#region [ Prepara um novo objeto para carregar o arquivo de retorno ]
				_titulos = new List<ArqRemessa.DetalheTitulo>();
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
			int idArqConciliacaoInput = 0;
			int qtdeRegProcessado;
			bool blnGerouNsu;
			bool blnSucesso;
			int percProgressoAtual;
			int percProgressoAnterior;
			String strMsgErro = "";
			String strMsgProgresso;
			String strNomeArquivoCompleto;
			String strNomeArquivoCompletoRenomeado;
			String strNomeArquivoCompletoRenomeadoAux;
			String[] linhasArqRetorno;
			String linhaHeader = "";
			String linhaTrailler = "";
			String cnpjEmpresa = "";
			DateTime dtInicioProcessamento;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			#region [ Obtém nome do arquivo de conciliação ]
			strNomeArquivoCompleto = txtArqRetorno.Text;
			#endregion

			#region [ Consistência ]
			if (strNomeArquivoCompleto.Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo de conciliação a ser carregado!!");
				return;
			}
			if (!File.Exists(strNomeArquivoCompleto))
			{
				avisoErro("O arquivo de conciliação selecionado não existe!!\n\n" + strNomeArquivoCompleto);
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma a carga do arquivo de conciliação?")) return;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "lendo o arquivo de conciliação");
			try
			{
				dtInicioProcessamento = DateTime.Now;

				#region [ Carrega dados do arquivo em array ]
				adicionaDisplay("Leitura dos registros do arquivo de conciliação " + Path.GetFileName(strNomeArquivoCompleto));
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

				#region [ Obtém as linhas do header e do trailler ]
				if (Texto.leftStr(linhasArqRetorno[0], 2).Equals("00"))
				{
					linhaHeader = linhasArqRetorno[0];
				}

				if (Texto.leftStr(linhasArqRetorno[linhasArqRetorno.Length - 1], 2).Equals("99"))
				{
					linhaTrailler = linhasArqRetorno[linhasArqRetorno.Length - 1];
				}

				if ((linhaHeader.Length == 0) || (linhaTrailler.Length == 0))
				{
					for (int i = 0; i < linhasArqRetorno.Length; i++)
					{
						if (linhaHeader.Length == 0)
						{
							if (Texto.leftStr(linhasArqRetorno[i], 2).Equals("00")) linhaHeader = linhasArqRetorno[i];
						}

						if (linhaTrailler.Length == 0)
						{
							if (Texto.leftStr(linhasArqRetorno[i], 2).Equals("99")) linhaTrailler = linhasArqRetorno[i];
						}

						if ((linhaHeader.Length > 0) && (linhaTrailler.Length > 0)) break;
					}
				}
				#endregion

				#region [ CNPJ da empresa conveniada ]
				if (linhaHeader.Length > 0)
				{
					cnpjEmpresa = linhaHeader.Substring(22, 14);
				}
				#endregion

				blnSucesso = false;
				try
				{
					BD.iniciaTransacao();

					info(ModoExibicaoMensagemRodape.EmExecucao, "processando o arquivo de conciliação");

					#region [ Gera o NSU para o novo registro que será gravado em t_SERASA_ARQ_CONCILIACAO_INPUT ]
					blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_ARQ_CONCILIACAO_INPUT, ref idArqConciliacaoInput, ref strMsgErro);
					if (!blnGerouNsu)
					{
						throw new Exception("Falha ao tentar gerar o NSU para o registro da carga do arquivo de conciliação!!\n" + strMsgErro);
					}
					#endregion

					#region [ Insere Registro na Tabela t_SERASA_ARQ_CONCILIACAO_INPUT ]
					if (!ArqConciliacaoInputDAO.insere(idArqConciliacaoInput,
														DateTime.Now,
														DateTime.Now,
														Global.Usuario.usuario,
														_periodoFinalHeader,
														_titulos.Count,
														ZERO_SEGUNDO,
														cnpjEmpresa,
														Path.GetFileName(strNomeArquivoCompleto),
														Path.GetDirectoryName(strNomeArquivoCompleto),
														ST_PROCESSAMENTO_EM_ANDAMENTO,
														linhaHeader,
														linhaTrailler,
														null))
					{
						throw new Exception("Falha na criação de registro para o arquivo de conciliação!!");
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

					strMsgErro = "Falha na carga do arquivo de conciliação!!\n\n" + strMsgErro;
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
				}
				else
				{
					BD.commitTransacao();

					try
					{
						BD.iniciaTransacao();
						int idConciliacaoTitulo = 0;

						qtdeRegProcessado = 0;
						percProgressoAnterior = 0;
						foreach (ArqRemessa.DetalheTitulo titulo in _titulos)
						{
							qtdeRegProcessado++;
							percProgressoAtual = 100 * qtdeRegProcessado / _titulos.Count;
							if (percProgressoAtual != percProgressoAnterior)
							{
								percProgressoAnterior = percProgressoAtual;
								strMsgProgresso = "Processando registros: " + percProgressoAtual.ToString() + "%";
								info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
								Application.DoEvents();
							}

							blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_CONCILIACAO_TITULO, ref idConciliacaoTitulo, ref strMsgErro);
							if (!blnGerouNsu)
							{
								throw new Exception("Falha ao tentar gerar o NSU para o registro do título do arquivo de conciliação!!\n" + strMsgErro);
							}

							String dtPagto = "";
							if (titulo.dataPagamento == DateTime.MinValue)
							{
								dtPagto = new String(' ', 8);
							}
							else
							{
								dtPagto = Global.formataDataYyyyMmDdSemSeparador(titulo.dataPagamento);
							}

							if (!ConciliacaoTituloDAO.insere(idConciliacaoTitulo,
															  idArqConciliacaoInput,
															  ST_TITULO_TRATADO_MANUAL_NAO,
															  DateTime.MinValue,
															  DateTime.MinValue,
															  null,
															  ST_ENVIADO_SERASA_NAO_ENVIADO,
															  ID_ARQ_CONC_OUTPUT_ZERO,
															  ArqRemessa.DetalheTitulo.ID,
															  titulo.cnpjSacado,
															  ArqRemessa.DetalheTitulo.TIPO_DADOS,
															  titulo.numeroTitulo.Substring(0, 10),
															  Global.formataDataYyyyMmDdSemSeparador(titulo.dataEmissao),
															  titulo.dataEmissao,
															  Global.formataMoedaSemSeparador(titulo.valorTitulo, 13),
															  titulo.valorTitulo,
															  null,
															  VL_TITULO_EDITADO_ZERO,
															  Global.formataDataYyyyMmDdSemSeparador(titulo.dataVencimento),
															  titulo.dataVencimento,
															  null,
															  DateTime.MinValue,
															  dtPagto,
															  titulo.dataPagamento,
															  null,
															  DateTime.MinValue,
															  "#D",
															  titulo.numeroTitulo))
							{
								throw new Exception("Falha na criação de registro para o título do arquivo de conciliação!!");
							}
						}

						if (!ArqConciliacaoInputDAO.atualizaDuracaoProcessamento(DateTime.Now.Subtract(dtInicioProcessamento).Seconds,
																				 idArqConciliacaoInput))
						{
							throw new Exception("Falha na atualização da duração do processamento!!");
						}

						if (!ArqConciliacaoInputDAO.atualizaStatusProcessamento(ST_PROCESSAMENTO_SUCESSO,
																				null,
																				idArqConciliacaoInput))
						{
							throw new Exception("Falha na atualização do status final do processamento do arquivo de conciliação!!");
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
						adicionaDisplay("Arquivo de conciliação carregado com sucesso!!");

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

						adicionaDisplay("Arquivo de conciliação renomeado para " + Path.GetFileName(strNomeArquivoCompletoRenomeado));
						info(ModoExibicaoMensagemRodape.Normal);
						aviso("Arquivo de conciliação carregado com sucesso!!\n\n" + strNomeArquivoCompleto);
					}
					else
					{
						BD.rollbackTransacao();

						try
						{
							BD.iniciaTransacao();
							if (!ArqConciliacaoInputDAO.atualizaStatusProcessamento(ST_PROCESSAMENTO_FALHA,
																					"Falha no processamento do arquivo de conciliação!",
																					idArqConciliacaoInput))
							{
								throw new Exception("Falha na atualização do status de erro do arquivo de conciliação!!");
							}
							BD.commitTransacao();
						}
						catch (Exception e)
						{
							Global.gravaLogAtividade(e.ToString());
							BD.rollbackTransacao();
						}

						strMsgErro = "Falha na carga do arquivo de conciliação!!\n\n" + strMsgErro;
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
		#region [ FArqRetornoConciliacao ]
		#region [ FArqRetornoConciliacao_Load ]
		private void FArqRetornoConciliacao_Load(object sender, EventArgs e)
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

		#region [ FArqRetornoConciliacao_Shown ]
		private void FArqRetornoConciliacao_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
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

		#region [ FArqRetornoConciliacao_FormClosing ]
		private void FArqRetornoConciliacao_FormClosing(object sender, FormClosingEventArgs e)
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
