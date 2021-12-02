#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Media;
using System.Text;
using System.Windows.Forms;
#endregion

namespace ADM2
{
	#region [ Classe: FIbpt ]
	public partial class FIbpt : ADM2.FModelo
	{
		#region [ Constantes ]
		const String GRID_COL_CODIGO = "colCodigo";
		const String GRID_COL_EX = "colEX";
		const String GRID_COL_TABELA = "colTabela";
		const String GRID_COL_ALIQ_NAC = "colAliqNac";
		const String GRID_COL_ALIQ_IMP = "colAliqImp";
		const String GRID_COL_DESCRICAO = "colDescricao";
		#endregion

		#region [ Atributos ]
		private bool _emProcessamento = false;
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

		String _nomeArquivoSelecionado = "";
		String[] _linhasDadosArquivo;
		LinhaHeaderArquivoIbptCsv _linhaHeader = new LinhaHeaderArquivoIbptCsv();
		int _qtdeLinhasNcm = 0;
		int _qtdeLinhasNbs = 0;
		int _qtdeLinhasLC116 = 0;
		#endregion

		#region [ Construtor ]
		public FIbpt()
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

		#region [ pathArquivoValorDefault ]
		private String pathArquivoValorDefault()
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
			if (Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv))
				{
					strResp = Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv;
				}
			}
			return strResp;
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtArquivo.Text = "";
			grdDados.Rows.Clear();
			lblVersaoArquivo.Text = "";
			lblQtdeRegsNcm.Text = "";
			lblQtdeRegsNbs.Text = "";
			lblQtdeRegsLC116.Text = "";
			lblQtdeRegsTotal.Text = "";
			lblTotalRegistros.Text = "";
			lbMensagem.Items.Clear();
			lbErro.Items.Clear();
			lblUltVersaoArqCarregadaBd.Text = "";

			preencheLblUltVersaoArqIbptCsvCarregadoBd();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Arquivo ]
			if (txtArquivo.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo que será carregado!!");
				return false;
			}
			if (!File.Exists(txtArquivo.Text))
			{
				avisoErro("O arquivo informado não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ preencheLblUltVersaoArqIbptCsvCarregadoBd ]
		private void preencheLblUltVersaoArqIbptCsvCarregadoBd()
		{
			#region [ Declarações ]
			String strMsgException;
			Parametro parametro;
			#endregion

			#region [ Obtém e exibe a versão e data do último arquivo carregado ]
			try
			{
				parametro = FMain.contextoBD.AmbienteBase.parametroDAO.getParametro(Global.Cte.ADM2.ID_T_PARAMETRO.VERSAO_ULT_ARQ_IBPT_CSV_CARREGADO);
			}
			catch (Exception ex)
			{
				strMsgException = ex.Message;
				parametro = null;
			}

			if (parametro == null)
			{
				lblUltVersaoArqCarregadaBd.Text = "(nenhum arquivo carregado)";
			}
			else
			{
				lblUltVersaoArqCarregadaBd.Text = parametro.campo_texto + "   (por " + parametro.usuario_ult_atualizacao + " em " + Global.formataDataDdMmYyyyHhMmComSeparador(parametro.dt_hr_ult_atualizacao) + ")";
			}
			#endregion
		}
		#endregion

		#region [ preencheCamposInfArquivo ]
		private void preencheCamposInfArquivo()
		{
			#region [ Declarações ]
			int intQtdeNcm = 0;
			int intQtdeNbs = 0;
			int intQtdeLC116 = 0;
			int intQtdeTotal = 0;
			LinhaDadosArquivoIbptCsv linhaDados = new LinhaDadosArquivoIbptCsv();
			#endregion

			lblVersaoArquivo.Text = _linhaHeader.versao;

			for (int i = 1; i < _linhasDadosArquivo.Length; i++)
			{
				if (_linhasDadosArquivo[i] == null) continue;
				if (_linhasDadosArquivo[i].Trim().Length == 0) continue;

				intQtdeTotal++;
				linhaDados.carregaDados(_linhasDadosArquivo[i]);
				if (linhaDados.tabela.Trim().Equals("0"))
				{
					intQtdeNcm++;
				}
				else if (linhaDados.tabela.Trim().Equals("1"))
				{
					intQtdeNbs++;
				}
				else if (linhaDados.tabela.Trim().Equals("2"))
				{
					intQtdeLC116++;
				}
			}

			_qtdeLinhasNcm = intQtdeNcm;
			_qtdeLinhasNbs = intQtdeNbs;
			_qtdeLinhasLC116 = intQtdeLC116;

			lblQtdeRegsNcm.Text = Global.formataInteiro(intQtdeNcm);
			lblQtdeRegsNbs.Text = Global.formataInteiro(intQtdeNbs);
			lblQtdeRegsLC116.Text = Global.formataInteiro(intQtdeLC116);
			lblQtdeRegsTotal.Text = Global.formataInteiro(intQtdeTotal);

			#region [ Exibe a versão e data do último arquivo carregado ]
			preencheLblUltVersaoArqIbptCsvCarregadoBd();
			#endregion
		}
		#endregion

		#region [ carregaGrid ]
		private void carregaGrid()
		{
			#region [ Declarações ]
			int intLinhaGrid = 0;
			String strMsg;
			LinhaDadosArquivoIbptCsv linhaDados = new LinhaDadosArquivoIbptCsv();
			#endregion

			try
			{
				_emProcessamento = true;

				#region [ Limpa campos ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "preparando o grid para exibição dos dados");
				grdDados.Rows.Clear();
				lblTotalRegistros.Text = "";
				if (_linhasDadosArquivo.Length > 1) grdDados.Rows.Add(_linhasDadosArquivo.Length - 1);
				Application.DoEvents();
				#endregion

				#region [ Preenche grid ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "carregando os dados no grid");
				for (int i = 1; i < _linhasDadosArquivo.Length; i++)
				{
					if (_linhasDadosArquivo[i] == null) continue;
					if (_linhasDadosArquivo[i].Trim().Length == 0) continue;

					linhaDados.carregaDados(_linhasDadosArquivo[i]);
					grdDados.Rows[intLinhaGrid].Cells[GRID_COL_CODIGO].Value = linhaDados.codigo;
					grdDados.Rows[intLinhaGrid].Cells[GRID_COL_EX].Value = linhaDados.ex;
					if (linhaDados.tabela.Trim().Equals("0"))
					{
						grdDados.Rows[intLinhaGrid].Cells[GRID_COL_TABELA].Value = "NCM";
					}
					else if (linhaDados.tabela.Trim().Equals("1"))
					{
						grdDados.Rows[intLinhaGrid].Cells[GRID_COL_TABELA].Value = "NBS";
					}
					else if (linhaDados.tabela.Trim().Equals("2"))
					{
						grdDados.Rows[intLinhaGrid].Cells[GRID_COL_TABELA].Value = "LC 116";
					}

					grdDados.Rows[intLinhaGrid].Cells[GRID_COL_DESCRICAO].Value = linhaDados.descricao;
					grdDados.Rows[intLinhaGrid].Cells[GRID_COL_ALIQ_NAC].Value = linhaDados.aliqNac;
					grdDados.Rows[intLinhaGrid].Cells[GRID_COL_ALIQ_IMP].Value = linhaDados.aliqImp;
					intLinhaGrid++;

					#region [ Exibe progresso ]
					if ((intLinhaGrid % 10) == 0)
					{
						strMsg = "Registros carregados no grid: " + Global.formataInteiro(intLinhaGrid) + "   (" + (100 * intLinhaGrid / (_qtdeLinhasNcm + _qtdeLinhasNbs + _qtdeLinhasLC116)).ToString() + "%)";
						info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);
						Application.DoEvents();
					}
					#endregion
				}
				#endregion

				lblTotalRegistros.Text = Global.formataInteiro(intLinhaGrid);
			}
			finally
			{
				_emProcessamento = false;
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoSelecionaArquivo ]
		private void trataBotaoSelecionaArquivo()
		{
			#region [ Declarações ]
			String strMsgErro;
			String strNomeArquivo;
			DialogResult dr;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			try
			{
				_emProcessamento = true;

				openFileDialogIbpt.InitialDirectory = pathArquivoValorDefault();
				dr = openFileDialogIbpt.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialogIbpt.FileName.Length > 0) && (txtArquivo.Text.Length > 0))
				{
					if (openFileDialogIbpt.FileName.ToUpper().Equals(txtArquivo.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos ]
				limpaCampos();
				#endregion

				strNomeArquivo = openFileDialogIbpt.FileName.Trim();
				if (!File.Exists(strNomeArquivo))
				{
					avisoErro("Arquivo selecionado não existe!!\n" + strNomeArquivo);
					return;
				}

				txtArquivo.Text = strNomeArquivo;
				_nomeArquivoSelecionado = strNomeArquivo;

				#region [ Carrega dados do arquivo em array ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "lendo dados do arquivo");
				_linhasDadosArquivo = File.ReadAllLines(_nomeArquivoSelecionado, encode);
				#endregion

				#region [ Consistência ]
				if (_linhasDadosArquivo == null)
				{
					avisoErro("É necessário selecionar um arquivo com conteúdo válido!!");
					return;
				}
				if (_linhasDadosArquivo.Length <= 1)
				{
					avisoErro("Arquivo selecionado não possui dados!!");
					return;
				}
				#endregion

				#region [ Carrega header ]
				_linhaHeader.carregaDados(_linhasDadosArquivo[0]);
				#endregion

				#region [ Consistência do header ]
				if ((_linhaHeader.codigo != "codigo") ||
					(_linhaHeader.ex != "ex") ||
					(_linhaHeader.tabela != "tabela") ||
					(_linhaHeader.descricao != "descricao") ||
					(_linhaHeader.aliqNac != "aliqNac") ||
					(_linhaHeader.aliqImp != "aliqImp") ||
					(_linhaHeader.versao.Trim().Length == 0))
				{
					strMsgErro = "Arquivo com header inválido!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				preencheCamposInfArquivo();
				carregaGrid();
				grdDados.Focus();

				Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv = Path.GetDirectoryName(_nomeArquivoSelecionado);
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
				_emProcessamento = false;
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoCarregaArquivo ]
		private void trataBotaoCarregaArquivo()
		{
			#region [ Declarações ]
			String strAux;
			String strMsg;
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			bool blnSucesso;
			long lngDuracaoProcessamentoEmSeg = 0;
			int intQtdeRegistrosGravadosTotal = 0;
			int intQtdeRegistrosGravadosNcm = 0;
			int intQtdeRegistrosGravadosNbs = 0;
			int intQtdeRegistrosGravadosLC116 = 0;
			int intQtdeRegistrosGravadosTotalUltAtualizProgresso = 0;
			DateTime dtInicioProcessamento;
			LinhaDadosArquivoIbptCsv linhaDados = new LinhaDadosArquivoIbptCsv();
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			Log log = new Log();
			Log logAutorizacao;
			Parametro parametro;
			#endregion

			try
			{
				_emProcessamento = true;

				#region [ Consistência do header ]
				if ((_linhaHeader.codigo != "codigo") ||
					(_linhaHeader.ex != "ex") ||
					(_linhaHeader.tabela != "tabela") ||
					(_linhaHeader.descricao != "descricao") ||
					(_linhaHeader.aliqNac != "aliqNac") ||
					(_linhaHeader.aliqImp != "aliqImp") ||
					(_linhaHeader.versao.Trim().Length == 0))
				{
					strMsgErro = "Arquivo com header inválido!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Confirmação ]
				if (!confirma("Confirma a carga do arquivo selecionado?")) return;
				#endregion

				#region [ Verifica se outro usuário pode estar executando este mesmo processamento ]
				if (FMain.contextoBD.AmbienteBase.ibptDadosDAO.isTabelaTemporariaCriada())
				{
					#region [ Exibe alerta e solicita autorização ]
					strAux = "ATENÇÃO!!" + "\n" +
							"O sistema detectou que está ocorrendo uma das seguintes situações:" + "\n" +
							"1) O processamento anterior não foi finalizado corretamente devido a algum erro." + "\n" +
							"2) Há outro usuário executando a carga do arquivo do IBPT neste momento." + "\n" +
							"\n" +
							"Certifique-se de que não há outro usuário executando este mesmo processamento, pois senão haverá o risco dos dados serem corrompidos!!" + "\n" +
							"Se houver dúvida, cancele esta operação e aguarde alguns minutos antes de tentar novamente!!" + "\n" +
							"\n" +
							"Caso deseje prosseguir com a carga do arquivo do IBPT, digite a sua senha para confirmar que autoriza a operação!!";
					fAutorizacao = new FAutorizacao(strAux);
					drAutorizacao = fAutorizacao.ShowDialog();
					if (drAutorizacao != DialogResult.OK)
					{
						strMsg = "Operação não confirmada!!\nA carga do arquivo do IBPT foi cancelada!!";
						adicionaErro(strMsg);
						avisoErro(strMsg);
						return;
					}
					if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
					{
						strMsg = "Senha inválida!!\nA carga do arquivo do IBPT foi cancelada!!";
						adicionaErro(strMsg);
						avisoErro(strMsg);
						return;
					}
					#endregion

					#region [ Grava log sobre a autorização dada ]
					logAutorizacao = new Log();
					logAutorizacao.usuario = Global.Usuario.usuario;
					logAutorizacao.operacao = Global.Cte.ADM2.LogOperacao.IBPT_CARGA_ARQUIVO_CSV_AUTORIZACAO;
					logAutorizacao.complemento = "O usuário digitou a senha para confirmar que deseja prosseguir com a operação após a exibição do alerta informando a existência de uma das seguintes situações: (1) O processamento anterior não foi finalizado corretamente devido a algum erro; (2) Há outro usuário executando a carga do arquivo do IBPT neste momento.";
					FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, logAutorizacao, ref strMsgErroLog);
					#endregion

					this.Refresh();
				}
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "carregando o arquivo selecionado para o BD");
				try
				{
					dtInicioProcessamento = DateTime.Now;

					#region [ Carrega dados do arquivo em array ]
					adicionaDisplay("Registros para processar: " + Global.formataInteiro(_qtdeLinhasNcm + _qtdeLinhasNbs + _qtdeLinhasLC116));
					#endregion

					blnSucesso = false;
					try
					{
						FMain.contextoBD.AmbienteBase.BD.iniciaTransacao();

						FMain.contextoBD.AmbienteBase.ibptDadosDAO.criaTabelaTemporaria();
						try
						{
							#region [ Grava cada uma das linhas de dados na tabela temporária ]
							for (int i = 1; i < _linhasDadosArquivo.Length; i++)
							{
								linhaDados.carregaDados(_linhasDadosArquivo[i]);

								if (linhaDados.tabela.Trim().Equals("0") || linhaDados.tabela.Trim().Equals("1") || linhaDados.tabela.Trim().Equals("2"))
								{
									if (!FMain.contextoBD.AmbienteBase.ibptDadosDAO.insereTabelaTemporaria(Global.Usuario.usuario, linhaDados, ref strDescricaoLog, ref strMsgErro))
									{
										throw new Exception("Falha ao tentar gravar dados do arquivo do IBPT na tabela temporária (linha de dados do arquivo: " + i.ToString() + ")!!\n" + strMsgErro);
									}
								}

								intQtdeRegistrosGravadosTotal++;

								if (linhaDados.tabela.Trim().Equals("0"))
								{
									intQtdeRegistrosGravadosNcm++;
								}
								else if (linhaDados.tabela.Trim().Equals("1"))
								{
									intQtdeRegistrosGravadosNbs++;
								}
								else if (linhaDados.tabela.Trim().Equals("2"))
								{
									intQtdeRegistrosGravadosLC116++;
								}

								#region [ Exibe progresso ]
								if ((intQtdeRegistrosGravadosTotal % 10) == 0)
								{
									intQtdeRegistrosGravadosTotalUltAtualizProgresso = intQtdeRegistrosGravadosTotal;
									strMsg = "Registros gravados: " + Global.formataInteiro(intQtdeRegistrosGravadosTotal) + "   (" + (100 * intQtdeRegistrosGravadosTotal / (_qtdeLinhasNcm + _qtdeLinhasNbs + _qtdeLinhasLC116)).ToString() + "%)";
									adicionaDisplay(strMsg);
									info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);
									Application.DoEvents();
								}
								#endregion
							}
							#endregion

							#region [ Exibe progresso (final) ]
							if (intQtdeRegistrosGravadosTotalUltAtualizProgresso != intQtdeRegistrosGravadosTotal)
							{
								strMsg = "Registros gravados: " + Global.formataInteiro(intQtdeRegistrosGravadosTotal) + "   (" + (100 * intQtdeRegistrosGravadosTotal / (_qtdeLinhasNcm + _qtdeLinhasNbs + _qtdeLinhasLC116)).ToString() + "%)";
								adicionaDisplay(strMsg);
								info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);
							}
							#endregion

							#region [ Transfere dados da tabela temporária para a tabela de produção ]
							info(ModoExibicaoMensagemRodape.EmExecucao, "transferindo dados da tabela temporária para a tabela de produção");
							adicionaDisplay("Transferindo dados da tabela temporária para a tabela de produção");
							if (!FMain.contextoBD.AmbienteBase.ibptDadosDAO.transfereDadosTabelaTemporariaParaTabelaProducao(ref strMsgErro))
							{
								throw new Exception("Falha ao transferir os dados da tabela temporária para a tabela de produção!!\n" + strMsgErro);
							}
							#endregion

							#region [ Grava a versão do arquivo do IBPT na tabela de controle ]
							parametro = new Parametro();
							parametro.id = Global.Cte.ADM2.ID_T_PARAMETRO.VERSAO_ULT_ARQ_IBPT_CSV_CARREGADO;
							parametro.campo_texto = _linhaHeader.versao.Trim();
							FMain.contextoBD.AmbienteBase.parametroDAO.salva(Global.Usuario.usuario, parametro, out strDescricaoLog, out strMsg);
							#endregion
						}
						finally
						{
							FMain.contextoBD.AmbienteBase.ibptDadosDAO.dropTabelaTemporaria();
						}

						lngDuracaoProcessamentoEmSeg = Global.calculaTimeSpanSegundos(DateTime.Now - dtInicioProcessamento);
						blnSucesso = true;
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(ex.ToString());
						strMsgErro = ex.ToString();
						blnSucesso = false;
					}

					if (blnSucesso)
					{
						FMain.contextoBD.AmbienteBase.BD.commitTransacao();

						adicionaDisplay("Arquivo carregado com sucesso!!");

						#region [ Grava o log no BD ]
						strDescricaoLog = "Sucesso na carga do arquivo do IBPT (versão: " + _linhaHeader.versao + "): " + _nomeArquivoSelecionado + " (duração do processamento: " + lngDuracaoProcessamentoEmSeg.ToString() + " segundos; total de registros gravados: " + Global.formataInteiro(intQtdeRegistrosGravadosTotal) + "; registros de NCM gravados: " + Global.formataInteiro(intQtdeRegistrosGravadosNcm) + "; registros de NBS gravados: " + Global.formataInteiro(intQtdeRegistrosGravadosNbs) + "; registros de LC 116 gravados: " + Global.formataInteiro(intQtdeRegistrosGravadosLC116) + ")";
						log.usuario = Global.Usuario.usuario;
						log.operacao = Global.Cte.ADM2.LogOperacao.IBPT_CARGA_ARQUIVO_CSV;
						log.complemento = strDescricaoLog;
						FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
						#endregion

						Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv = Path.GetDirectoryName(_nomeArquivoSelecionado);

						preencheLblUltVersaoArqIbptCsvCarregadoBd();

						info(ModoExibicaoMensagemRodape.Normal);
						aviso("Arquivo carregado com sucesso!!\n\n" + _nomeArquivoSelecionado);
					}
					else
					{
						FMain.contextoBD.AmbienteBase.BD.rollbackTransacao();

						#region [ Grava o log no BD ]
						strDescricaoLog = "Falha na carga do arquivo do IBPT (versão: " + _linhaHeader.versao + "): " + _nomeArquivoSelecionado + " (" + strMsgErro + ")";
						log.usuario = Global.Usuario.usuario;
						log.operacao = Global.Cte.ADM2.LogOperacao.IBPT_CARGA_ARQUIVO_CSV;
						log.complemento = strDescricaoLog;
						FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
						#endregion

						strMsgErro = "Falha na carga do arquivo!!\n\n" + strMsgErro;
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
					}
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
				}
			}
			finally
			{
				_emProcessamento = false;
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FIbpt ]

		#region [ FIbpt_Load ]
		private void FIbpt_Load(object sender, EventArgs e)
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

		#region [ FIbpt_Shown ]
		private void FIbpt_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Ajusta layout do header do grid ]
					grdDados.Columns[GRID_COL_CODIGO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdDados.Columns[GRID_COL_EX].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdDados.Columns[GRID_COL_TABELA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdDados.Columns[GRID_COL_ALIQ_NAC].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdDados.Columns[GRID_COL_ALIQ_IMP].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdDados.Columns[GRID_COL_DESCRICAO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
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

		#region [ FIbpt_FormClosing ]
		private void FIbpt_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (_emProcessamento)
			{
				SystemSounds.Exclamation.Play();
				e.Cancel = true;
				return;
			}

			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaArquivo ]

		#region [ btnSelecionaArquivo_Click ]
		private void btnSelecionaArquivo_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArquivo();
		}
		#endregion

		#endregion

		#region [ btnCarregaArquivo ]

		#region [ btnCarregaArquivo_Click ]
		private void btnCarregaArquivo_Click(object sender, EventArgs e)
		{
			trataBotaoCarregaArquivo();
		}
		#endregion

		#endregion

		#region [ txtArquivo ]

		#region [ txtArquivo_Enter ]
		private void txtArquivo_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtArquivo_DoubleClick ]
		private void txtArquivo_DoubleClick(object sender, EventArgs e)
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

		#region [ grdDados ]

		#region [ grdDados_SortCompare ]
		private void grdDados_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			#region [ Declarações ]
			bool blnOrdenarPeloCodigo = false;
			String s1;
			String s2;
			#endregion

			if (e.Column.Name.Equals(GRID_COL_CODIGO))
			{
				s1 = e.CellValue1.ToString();
				s2 = e.CellValue2.ToString();
				if (s1.Length != s2.Length)
				{
					e.SortResult = s1.Length - s2.Length;
				}
				else
				{
					e.SortResult = (int)Global.converteInteiro(s1) - (int)Global.converteInteiro(s2);
				}
				e.Handled = true;
			}
			else
			{
				s1 = e.CellValue1.ToString();
				s2 = e.CellValue2.ToString();
				if (e.Column.Name.Equals(GRID_COL_EX))
				{
					if (!s1.Equals(s2))
					{ e.SortResult = (int)Global.converteInteiro(s1) - (int)Global.converteInteiro(s2); }
					else
					{ blnOrdenarPeloCodigo = true; }
					e.Handled = true;
				}
				else if (e.Column.Name.Equals(GRID_COL_TABELA))
				{
					if (!s1.Equals(s2))
					{ e.SortResult = String.Compare(s1, s2); }
					else { blnOrdenarPeloCodigo = true; }
					e.Handled = true;
				}
				else if (e.Column.Name.Equals(GRID_COL_ALIQ_NAC))
				{
					if (!s1.Equals(s2))
					{ e.SortResult = System.Decimal.Compare(Global.converteNumeroDecimal(s1), Global.converteNumeroDecimal(s2)); }
					else { blnOrdenarPeloCodigo = true; }
					e.Handled = true;
				}
				else if (e.Column.Name.Equals(GRID_COL_ALIQ_IMP))
				{
					if (!s1.Equals(s2))
					{ e.SortResult = System.Decimal.Compare(Global.converteNumeroDecimal(s1), Global.converteNumeroDecimal(s2)); }
					else { blnOrdenarPeloCodigo = true; }
					e.Handled = true;
				}
				else if (e.Column.Name.Equals(GRID_COL_DESCRICAO))
				{
					if (!s1.Equals(s2))
					{ e.SortResult = String.Compare(s1, s2); }
					else { blnOrdenarPeloCodigo = true; }
					e.Handled = true;
				}
			}

			if (blnOrdenarPeloCodigo)
			{
				s1 = grdDados.Rows[e.RowIndex1].Cells[GRID_COL_CODIGO].Value.ToString();
				s2 = grdDados.Rows[e.RowIndex2].Cells[GRID_COL_CODIGO].Value.ToString();
				if (s1.Length != s2.Length)
				{
					e.SortResult = s1.Length - s2.Length;
				}
				else
				{
					e.SortResult = (int)Global.converteInteiro(s1) - (int)Global.converteInteiro(s2);
				}
			}
		}
		#endregion

		#endregion

		#endregion
	}
	#endregion
}
