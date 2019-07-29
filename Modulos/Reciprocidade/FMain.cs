#region [ using ]
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
#endregion

namespace Reciprocidade
{
	public partial class FMain : FModelo
	{
		#region[ Atributos ]
		public static FMain _fMain;
		private bool _InicializacaoOk;
		private String REGISTRY_PATH_FORM_OPTIONS;
		FArqRemessa _fArqRemessa;
		FArqRetorno _fArqRetorno;
		FSerasaOcorrencias _fSerasaOcorrencias;
		FArqRemessaRetificacao _fArqRemessaRetificacao;
		FArqRetornoConciliacao _fArqRetornoConciliacao;
		FTitulosConciliacao _fTitulosConciliacao;
		FArqRemessaConciliacao _fArqRemessaConciliacao;
		#endregion

		#region [ Construtor ]
		public FMain()
		{
			InitializeComponent();
			_fMain = this;
			if (Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa.Length > 0)
			{
				if (!Directory.Exists(Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa)) Directory.CreateDirectory(Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa);
			}
			REGISTRY_PATH_FORM_OPTIONS = Global.RegistryApp.REGISTRY_BASE_PATH + "\\" + this.Name;
			if (!Directory.Exists(Global.Cte.LogAtividade.PathLogAtividade)) Directory.CreateDirectory(Global.Cte.LogAtividade.PathLogAtividade);
		}
		#endregion

		#region [ Métodos Privados ]
		#region[ iniciaBancoDados ]
		/// <summary>
		/// Inicializa os objetos de acesso ao banco de dados e se conecta ao servidor.
		/// </summary>
		/// <returns>
		/// True: sucesso
		/// False: falha
		/// </returns>
		private bool iniciaBancoDados(ref String strMsgErro)
		{
			strMsgErro = "";
			try
			{
				BD.abreConexao();
				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ trataBotaoArqRemessa ]
		private void trataBotaoArqRemessa()
		{
			_fArqRemessa = new FArqRemessa();
			_fArqRemessa.Location = this.Location;
			_fArqRemessa.Show();
			this.Visible = false;
		}
		#endregion

		#region [ trataBotaoArqRetorno ]
		private void trataBotaoArqRetorno()
		{
			_fArqRetorno = new FArqRetorno();
			_fArqRetorno.Location = this.Location;
			_fArqRetorno.Show();
			this.Visible = false;
		}
		#endregion

		#region [ trataBotaoOcorrencias ]
		private void trataBotaoOcorrencias()
		{
			_fSerasaOcorrencias = new FSerasaOcorrencias();
			_fSerasaOcorrencias.Location = this.Location;
			_fSerasaOcorrencias.Show();
			this.Visible = false;
		}
		#endregion

		#region [ trataBotaoArqRemessaRetificacao ]
		private void trataBotaoArqRemessaRetificacao()
		{
			_fArqRemessaRetificacao = new FArqRemessaRetificacao();
			_fArqRemessaRetificacao.Location = this.Location;
			_fArqRemessaRetificacao.Show();
			this.Visible = false;
		}
		#endregion

		#region [ trataBotaoArqRetornoConciliacao ]
		private void trataBotaoArqRetornoConciliacao()
		{
			_fArqRetornoConciliacao = new FArqRetornoConciliacao();
			_fArqRetornoConciliacao.Location = this.Location;
			_fArqRetornoConciliacao.Show();
			this.Visible = false;
		}
		#endregion

		#region [ trataBotaoOcorrenciasConciliacao ]
		private void trataBotaoOcorrenciasConciliacao()
		{
			_fTitulosConciliacao = new FTitulosConciliacao();
			_fTitulosConciliacao.Location = this.Location;
			_fTitulosConciliacao.Show();
			this.Visible = false;
		}
		#endregion

		#region [ trataBotaoArqRemessaConciliacao ]
		private void trataBotaoArqRemessaConciliacao()
		{
			_fArqRemessaConciliacao = new FArqRemessaConciliacao();
			_fArqRemessaConciliacao.Location = this.Location;
			_fArqRemessaConciliacao.Show();
			this.Visible = false;
		}
		#endregion
		#endregion

		#region [ Eventos ]
		#region [ FMain ]
		#region [ FMain_Shown ]
		private void FMain_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strSenhaDescriptografada = "";
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strUltimoUsuario;
			String strTop;
			String strLeft;
			String strMsg;
			int intTop;
			int intLeft;
			bool blnRestauraPosicaoAnterior;
			bool blnValidacaoUsuarioOk;
			Color? cor;
			DateTime dtHrServidor;
			UsuarioDAO usuarioDAO;
			VersaoModulo versaoModulo;
			FLogin fLogin = new FLogin();
			DialogResult drLogin;
			Log log = new Log();
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Registry: posição do form na execução anterior ]
					RegistryKey regKey = Global.RegistryApp.criaRegistryKey(REGISTRY_PATH_FORM_OPTIONS);
					strTop = (String)regKey.GetValue(Global.RegistryApp.Chaves.top);
					intTop = (int)Global.converteInteiro(strTop);
					if (intTop < 0) intTop = 1;
					strLeft = (String)regKey.GetValue(Global.RegistryApp.Chaves.left);
					intLeft = (int)Global.converteInteiro(strLeft);
					if (intLeft < 0) intLeft = 1;

					blnRestauraPosicaoAnterior = true;
					if ((strTop == null) || (strLeft == null)) blnRestauraPosicaoAnterior = false;
					if (intTop > Screen.PrimaryScreen.WorkingArea.Height - 100) blnRestauraPosicaoAnterior = false;
					if (intLeft > Screen.PrimaryScreen.WorkingArea.Width - 100) blnRestauraPosicaoAnterior = false;

					if (blnRestauraPosicaoAnterior)
					{
						this.StartPosition = FormStartPosition.Manual;
						this.Top = intTop;
						this.Left = intLeft;
					}
					#endregion

#if (HOMOLOGACAO)
					this.Text += "  (Versão de HOMOLOGAÇÃO)";
					if (!confirma("Versão exclusiva para o ambiente de HOMOLOGAÇÃO!!\nContinua assim mesmo?"))
					{
						Close();
						return;
					}
#elif (DESENVOLVIMENTO)
					this.Text += "  (Versão de DESENVOLVIMENTO)";
					if (!confirma("Versão exclusiva de DESENVOLVIMENTO!!\nContinua assim mesmo?"))
					{
						Close();
						return;
					}
#elif (PRODUCAO)
					// NOP
#else
					this.Text += "  (Versão DESCONHECIDA)";
					avisoErro("Versão DESCONHECIDA!!\nNão é possível continuar!!");
					Close();
					return;
#endif

					#region [ Verifica se parâmetros do arquivo de configuração foram lidos corretamente ]
					if ((Global.Cte.SerasaReciprocidade.CNPJ_EMPRESA_CONVENIADA ?? "").Trim().Length == 0)
					{
						avisoErro("Falha ao ler do arquivo de configuração o CNPJ da empresa conveniada com a Serasa!!");
						Close();
						return;
					}

					if (!Global.isCnpjOk(Global.Cte.SerasaReciprocidade.CNPJ_EMPRESA_CONVENIADA))
					{
						avisoErro("O CNPJ da empresa conveniada com a Serasa obtido do arquivo de configuração é inválido: " + Global.Cte.SerasaReciprocidade.CNPJ_EMPRESA_CONVENIADA);
						Close();
						return;
					}
					#endregion

					#region [ Registry: dados da sessão anterior ]
					strUltimoUsuario = (String)regKey.GetValue(Global.RegistryApp.Chaves.usuario, "");
					Global.Usuario.Defaults.FArqRetorno.pathTituloArquivoRetorno = (String)regKey.GetValue(Global.RegistryApp.Chaves.FArqRetorno.pathTituloArquivoRetorno, "");
					Global.Usuario.Defaults.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao = (String)regKey.GetValue(Global.RegistryApp.Chaves.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao, "");
					Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa = (String)regKey.GetValue(Global.RegistryApp.Chaves.FArqRemessa.pathTituloArquivoRemessa, "");
					#endregion

					#region [ Login do usuário ]
					// Laço para obter dados corretos na tela de login
					// Permanece no laço até digitar um usuário/senha correto ou o usuário cancelar
					FLogin.usuario = strUltimoUsuario;
					do
					{
						blnValidacaoUsuarioOk = true;

						#region [ Obtém login do usuário ]
						fLogin.Location = new Point(this.Location.X + (this.Size.Width - fLogin.Size.Width) / 2, this.Location.Y + (this.Size.Height - fLogin.Size.Height) / 2);
						drLogin = fLogin.ShowDialog();
						// O usuário cancelou o login
						if (drLogin != DialogResult.OK)
						{
							avisoErro("Login cancelado!!");
							Close();
							return;
						}
						#endregion

						try
						{
							#region[ Inicia Banco de Dados ]
							info(ModoExibicaoMensagemRodape.EmExecucao, "conectando com o banco de dados");
							if (!iniciaBancoDados(ref strMsgErro))
							{
								avisoErro("Falha ao conectar com o banco de dados!!\n\n" + strMsgErro);
								Close();
								return;
							}
							#endregion

							#region [ Validação do usuário ]
							info(ModoExibicaoMensagemRodape.EmExecucao, "validando usuário");
							Global.Usuario.usuario = FLogin.usuario;
							Global.Usuario.senhaDigitada = FLogin.senha;

							#region [ Obtém dados no BD ]
							usuarioDAO = new UsuarioDAO(Global.Usuario.usuario, ref Global.Acesso.listaOperacoesPermitidas);
							Global.Usuario.usuario = usuarioDAO.usuario;
							Global.Usuario.nome = usuarioDAO.nome;
							Global.Usuario.senhaCriptografada = usuarioDAO.datastamp;
							// Descriptografa a senha
							if (!CriptoHex.decodificaDado(Global.Usuario.senhaCriptografada, ref strSenhaDescriptografada)) strSenhaDescriptografada = "";
							Global.Usuario.senhaDescriptografada = strSenhaDescriptografada;
							Global.Usuario.cadastrado = usuarioDAO.cadastrado;
							Global.Usuario.bloqueado = usuarioDAO.bloqueado;
							Global.Usuario.senhaExpirada = usuarioDAO.senhaExpirada;
							Global.Usuario.fin_email_remetente = usuarioDAO.fin_email_remetente;
							Global.Usuario.fin_display_name_remetente = usuarioDAO.fin_display_name_remetente;
							Global.Usuario.fin_servidor_smtp_endereco = usuarioDAO.fin_servidor_smtp;
							Global.Usuario.fin_servidor_smtp_porta = usuarioDAO.fin_servidor_smtp_porta;
							Global.Usuario.fin_usuario_smtp = usuarioDAO.fin_usuario_smtp;
							// Descriptografa a senha
							Global.Usuario.fin_senha_smtp = Criptografia.Descriptografa(usuarioDAO.fin_senha_smtp);
							#endregion

							#region [ Usuário não cadastrado ]
							if (blnValidacaoUsuarioOk)
							{
								// Não cadastrado
								if (!Global.Usuario.cadastrado)
								{
									avisoErro("Usuário não cadastrado!!\n\n" + strMsgErro);
									blnValidacaoUsuarioOk = false;
								}
							}
							#endregion

							#region [ Acesso bloqueado ]
							if (blnValidacaoUsuarioOk)
							{
								// Acesso bloqueado
								if (Global.Usuario.bloqueado)
								{
									avisoErro("Acesso bloqueado!!\n\n" + strMsgErro);
									blnValidacaoUsuarioOk = false;
								}
							}
							#endregion

							#region [ Senha expirada ]
							if (blnValidacaoUsuarioOk)
							{
								// Senha expirada
								if (Global.Usuario.senhaExpirada)
								{
									avisoErro("Senha expirada!!\n\n" + strMsgErro);
									blnValidacaoUsuarioOk = false;
								}
							}
							#endregion

							#region [ Senha incorreta ]
							if (blnValidacaoUsuarioOk)
							{
								// Senha incorreta
								if (!Global.Usuario.senhaDescriptografada.ToUpper().Equals(Global.Usuario.senhaDigitada.ToUpper()))
								{
									avisoErro("Senha inválida!!\n\n" + strMsgErro);
									blnValidacaoUsuarioOk = false;
								}
							}
							#endregion

							#region [ Permissão de acesso ao módulo ]
							if (blnValidacaoUsuarioOk)
							{
								// Permissão de acesso ao módulo
								if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_SERASA_RECIPROCIDADE_ACESSO_AO_MODULO))
								{
									avisoErro("Nível de acesso insuficiente!!\n\n" + strMsgErro);
									blnValidacaoUsuarioOk = false;
								}
							}
							#endregion

							#endregion
						}
						finally
						{
							info(ModoExibicaoMensagemRodape.Normal);
						}
					} while (!blnValidacaoUsuarioOk);
					#endregion

					#region [ Inicializa construtores estáticos ]
					LogDAO.inicializaConstrutorEstatico();
					UsuarioDAO.inicializaConstrutorEstatico();
					ParametroDAO.inicializaConstrutorEstatico();
					//ComumDAO.inicializaConstrutorEstatico();
					#endregion

					#region [ Verifica data/hora da máquina local ]
					dtHrServidor = BD.obtemDataHoraServidor();
					if (dtHrServidor != DateTime.MinValue)
					{
						if (Math.Abs(Global.calculaTimeSpanMinutos(DateTime.Now - dtHrServidor)) > 90)
						{
							strMsg = "O relógio desta máquina está defasado com relação ao servidor além do limite máximo tolerado:" +
									 "\n\n" +
									 "Horário no servidor: " + Global.formataDataDdMmYyyyHhMmSsComSeparador(dtHrServidor) +
									 "\n" +
									 "Horário nesta máquina: " + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) +
									 "\n" +
									 "Defasagem: " + Math.Abs(Global.calculaTimeSpanMinutos(DateTime.Now - dtHrServidor)).ToString() + " minutos" +
									 "\n\n" +
									 "O programa será fechado!!" +
									 "\n" +
									 "Ajuste o relógio manualmente antes de tentar novamente!!";
							Global.gravaLogAtividade(strMsg);
							avisoErro(strMsg);
							Close();
							return;
						}
					}
					#endregion

					#region [ Armazena a data/hora de início ]
					Global.dtHrInicioRefRelogioServidor = dtHrServidor;
					Global.dtHrInicioRefRelogioLocal = DateTime.Now;
					#endregion

					#region [ Apaga os arquivos de log de atividade antigos ]
					Global.executaManutencaoArqLogAtividade();
					#endregion

					#region [ Grava no arquivo de log o início do aplicativo ]
					string linhaSeparadora = new string('=', 150);
					Global.gravaLogAtividade(linhaSeparadora);
					Global.gravaLogAtividade("Iniciado: " + Global.Cte.Aplicativo.M_ID);
					Global.gravaLogAtividade("Usuário: " + Global.Usuario.usuario + (Global.Usuario.usuario.ToUpper().Equals(Global.Usuario.nome.ToUpper()) ? "" : " - " + Global.Usuario.nome));
					Global.gravaLogAtividade(linhaSeparadora);
					#endregion

					#region [ Validação da versão deste programa ]
					versaoModulo = BD.getVersaoModulo("SerasaRecipr", out strMsgErro);
					if (versaoModulo == null)
					{
						strMsgErro = "Falha ao tentar obter no banco de dados o número da versão em produção deste aplicativo!!\n" + strMsgErro;
						Global.gravaLogAtividade(strMsgErro);
						avisoErro(strMsgErro);
						Close();
						return;
					}

					if (!versaoModulo.versao.Equals(Global.Cte.Aplicativo.VERSAO_NUMERO))
					{
						strMsgErro = "Versão inválida do aplicativo!!\n\nVersão deste programa: " + Global.Cte.Aplicativo.VERSAO_NUMERO + "\nVersão permitida: " + versaoModulo.versao;
						Global.gravaLogAtividade(strMsgErro);
						avisoErro(strMsgErro);
						Close();
						return;
					}
					#endregion

					#region [ Cor de fundo padrão cadastrado no BD ]
					if (versaoModulo.cor_fundo_padrao != null)
					{
						if (versaoModulo.cor_fundo_padrao.Trim().Length > 0)
						{
							cor = Global.converteColorFromHtml(versaoModulo.cor_fundo_padrao);
							if (cor != null)
							{
								if (cor != Global.BackColorPainelPadrao)
								{
									Global.BackColorPainelPadrao = (Color)cor;
									for (int i = 0; i < Application.OpenForms.Count; i++)
									{
										Application.OpenForms[i].BackColor = (Color)cor;
									}

									#region [ Salva a cor padrão indicada no BD no arquivo de configuração ]
									Global.setBackColorToAppConfig(versaoModulo.cor_fundo_padrao);
									#endregion
								}
							}
						}
					}
					#endregion

					#region [ Log de logon realizado gravado no BD ]
					log.usuario = Global.Usuario.usuario;
					log.operacao = Global.Cte.Reciprocidade.LogOperacao.LOGON;
					log.complemento = "Logon realizado na máquina=" +
										System.Environment.MachineName +
										"; OS=" + System.Environment.OSVersion.VersionString +
										"; OS Version=" + System.Environment.OSVersion.Version +
										"; OS SP=" + System.Environment.OSVersion.ServicePack +
										"; Processor Count=" + System.Environment.ProcessorCount.ToString() +
										"; Windows User Name=" + System.Environment.UserName;
					LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
					#endregion

					_InicializacaoOk = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
				// Se não inicializou corretamente, assegura-se de que o programa será terminado
				if (!_InicializacaoOk) Application.Exit();
			}
		}
		#endregion

		#region [ FMain_FormClosing ]
		private void FMain_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Declarações ]
			Log log;
			String strMsgErroLog = "";
			#endregion

			if (_InicializacaoOk)
			{
				#region [ Memoriza no registry ]
				RegistryKey regKey = Global.RegistryApp.criaRegistryKey(REGISTRY_PATH_FORM_OPTIONS);
				regKey.SetValue(Global.RegistryApp.Chaves.top, this.Top.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.left, this.Left.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.usuario, Global.Usuario.usuario);
				regKey.SetValue(Global.RegistryApp.Chaves.FArqRetorno.pathTituloArquivoRetorno, Global.Usuario.Defaults.FArqRetorno.pathTituloArquivoRetorno);
				regKey.SetValue(Global.RegistryApp.Chaves.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao, Global.Usuario.Defaults.FArqRetornoConciliacao.pathTituloArquivoRetornoConciliacao);
				regKey.SetValue(Global.RegistryApp.Chaves.FArqRemessa.pathTituloArquivoRemessa, Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa);
				#endregion

				#region [ Log em arquivo ]
				Global.gravaLogAtividade("Término do programa");
				Global.gravaLogAtividade(null);
				Global.gravaLogAtividade(null);
				#endregion

				#region [ Log de logoff realizado gravado no BD ]
				log = new Log();
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.Reciprocidade.LogOperacao.LOGOFF;
				log.complemento = "Logoff após " + Global.formataDuracaoHMS(DateTime.Now - Global.dtHrInicioRefRelogioLocal);
				LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion
			}
			BD.fechaConexao();
		}
		#endregion
		#endregion

		#region [ Botao Remessa Normal ]
		private void btnGeraArqRemessa_Click(object sender, EventArgs e)
		{
			trataBotaoArqRemessa();
		}
		#endregion

		#region [ Botoes Retorno ]
		private void btnCarregaArqRetorno_Click(object sender, EventArgs e)
		{
			trataBotaoArqRetorno();
		}

		private void btnTrataOcorrencias_Click(object sender, EventArgs e)
		{
			trataBotaoOcorrencias();
		}

		private void btnGeraArqRemessaRetificacao_Click(object sender, EventArgs e)
		{
			trataBotaoArqRemessaRetificacao();
		}
		#endregion

		#region [ Botoes Conciliacao ]
		private void btnCarregaArqRetornoConciliacao_Click(object sender, EventArgs e)
		{
			trataBotaoArqRetornoConciliacao();
		}

		private void btnTrataOcorrenciasConciliacao_Click(object sender, EventArgs e)
		{
			trataBotaoOcorrenciasConciliacao();
		}

		private void btnGeraArqRemessaConciliacao_Click(object sender, EventArgs e)
		{
			trataBotaoArqRemessaConciliacao();
		}
		#endregion
		#endregion
	}
}
