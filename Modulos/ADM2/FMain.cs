#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Configuration;
#endregion

namespace ADM2
{
	public partial class FMain : ADM2.FModelo
	{
		#region[ Atributos ]
		public static FMain fMain;
		public static ContextoBD contextoBD = new ContextoBD();
		private String REGISTRY_PATH_FORM_OPTIONS;
		FIbpt fIbpt;
		FAtualizaPlanilhaEstoque fAtualizaPlanilhaEstoque;
		FAnotarPedidoRecebidoCliente fAnotarPedidoRecebidoCliente;
		private bool _InicializacaoOk;
		private bool _OcorreuExceptionNaInicializacao = false;
		#endregion

		#region [ Construtor ]
		public FMain()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ADM2.FMain.Constructor()";
			int qtdeAmbientesBD;
			int indiceAmbienteBDBase = 0;
			string strNomeAmbiente;
			string strServidorBanco;
			string strNomeBanco;
			string strLoginBanco;
			string strSenhaBancoCriptografada;
			string strNumeroLoja;
			string msgErro;
			string strMsg;
			string strParametro;
			string strValue;
			AmbienteBD ambiente;
			#endregion

			InitializeComponent();
			fMain = this;
			REGISTRY_PATH_FORM_OPTIONS = Global.RegistryApp.REGISTRY_BASE_PATH + "\\" + this.Name;
			if (!Directory.Exists(Global.Cte.LogAtividade.PathLogAtividade)) Directory.CreateDirectory(Global.Cte.LogAtividade.PathLogAtividade);

			try
			{
				#region [ Conexões com os BD's ]

				#region [ Parâmetro: QtdeAmbientesBD ]
				strParametro = "QtdeAmbientesBD";
				strValue = Global.GetConfigurationValue(strParametro);
				if ((strValue ?? "").Trim().Length == 0)
				{
					msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
					throw new Exception(msgErro);
				}

				qtdeAmbientesBD = (int)Global.converteInteiro(strValue);
				if (qtdeAmbientesBD <= 0)
				{
					msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': quantidade inválida (" + strValue + ")!";
					throw new Exception(msgErro);
				}
				#endregion

				#region [ Parâmetro: IndiceAmbienteBDBase ]
				strParametro = "IndiceAmbienteBDBase";
				strValue = Global.GetConfigurationValue(strParametro);
				if ((strValue ?? "").Trim().Length == 0)
				{
					msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
					throw new Exception(msgErro);
				}

				// O índice adicionado no sufixo dos parâmetros inicia em 1
				indiceAmbienteBDBase = (int)Global.converteInteiro(strValue);
				if (indiceAmbienteBDBase <= 0)
				{
					msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': índice inválido (" + strValue + ")!";
					throw new Exception(msgErro);
				}

				if (indiceAmbienteBDBase > qtdeAmbientesBD)
				{
					msgErro = NOME_DESTA_ROTINA + " - Parâmetro '" + strParametro + "' informa um índice maior que a quantidade de ambientes (índice: " + indiceAmbienteBDBase.ToString() + ", qtde ambientes: " + qtdeAmbientesBD.ToString() + ")!";
					throw new Exception(msgErro);
				}
				#endregion

				#region [ Parâmetros de conexão ao BD ]
				for (int i = 1; i <= qtdeAmbientesBD; i++)
				{
					strParametro = "NomeAmbiente" + i.ToString();
					strNomeAmbiente = ConfigurationManager.ConnectionStrings[strParametro].ConnectionString;
					if ((strNomeAmbiente ?? "").Trim().Length == 0)
					{
						msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
						throw new Exception(msgErro);
					}

					strParametro = "ServidorBanco" + i.ToString();
					strServidorBanco = ConfigurationManager.ConnectionStrings[strParametro].ConnectionString;
					if ((strServidorBanco ?? "").Trim().Length == 0)
					{
						msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
						throw new Exception(msgErro);
					}

					strParametro = "NomeBanco" + i.ToString();
					strNomeBanco = ConfigurationManager.ConnectionStrings[strParametro].ConnectionString;
					if ((strNomeBanco ?? "").Trim().Length == 0)
					{
						msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
						throw new Exception(msgErro);
					}

					strParametro = "LoginBanco" + i.ToString();
					strLoginBanco = ConfigurationManager.ConnectionStrings[strParametro].ConnectionString;
					if ((strLoginBanco ?? "").Trim().Length == 0)
					{
						msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
						throw new Exception(msgErro);
					}

					strParametro = "SenhaBanco" + i.ToString();
					strSenhaBancoCriptografada = ConfigurationManager.ConnectionStrings[strParametro].ConnectionString;
					if ((strSenhaBancoCriptografada ?? "").Trim().Length == 0)
					{
						msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
						throw new Exception(msgErro);
					}

					strParametro = "NumeroLoja" + i.ToString();
					strNumeroLoja = ConfigurationManager.ConnectionStrings[strParametro].ConnectionString;
					if ((strNumeroLoja ?? "").Trim().Length == 0)
					{
						msgErro = NOME_DESTA_ROTINA + " - Falha na leitura do parâmetro '" + strParametro + "': não há valor configurado!";
						throw new Exception(msgErro);
					}

					ambiente = new AmbienteBD(strNomeAmbiente, strServidorBanco, strNomeBanco, strLoginBanco, strSenhaBancoCriptografada, strNumeroLoja);

					contextoBD.Ambientes.Add(ambiente);
					if (indiceAmbienteBDBase == i) contextoBD.AmbienteBase = ambiente;
				}

				#endregion

				if (contextoBD.AmbienteBase == null) throw new Exception("Nenhum ambiente de banco de dados foi definido para ser usado como ambiente base!");
				#endregion
			}
			catch (Exception ex)
			{
				msgErro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\r\n" + msgErro);
				strMsg = "Falha ao iniciar módulo!!\r\n\r\n" + ex.Message;
				avisoErro(strMsg);
				_OcorreuExceptionNaInicializacao = true;
			}
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ inicializaConstrutoresEstaticosUnitsDAO ]
		private static void inicializaConstrutoresEstaticosUnitsDAO()
		{
			LogDAO.inicializaConstrutorEstatico();
			UsuarioDAO.inicializaConstrutorEstatico();
			IbptDadosDAO.inicializaConstrutorEstatico();
			ParametroDAO.inicializaConstrutorEstatico();
			ComboDAO.inicializaConstrutorEstatico();
			PedidoDAO.inicializaConstrutorEstatico();
			ProdutoDAO.inicializaConstrutorEstatico();
		}
		#endregion

		#region [ reinicializaObjetosUnitsDAO ]
		private static void reinicializaObjetosUnitsDAO()
		{
			try
			{
				foreach (var item in contextoBD.Ambientes)
				{
					item.logDAO.inicializaObjetos();
					item.usuarioDAO.inicializaObjetos();
					item.ibptDadosDAO.inicializaObjetos();
					item.parametroDAO.inicializaObjetos();
					item.pedidoDAO.inicializaObjetos();
					item.produtoDAO.inicializaObjetos();
				}

				Global.gravaLogAtividade("Sucesso ao reinicializar os objetos das units de acesso ao Banco de Dados!!");
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao reinicializar os objetos das units de acesso ao Banco de Dados!!\n" + ex.Message);
			}
		}
		#endregion

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
			#region [ Declarações ]
			string msg_erro_completo;
			string msg_erro_resumido;
			#endregion

			strMsgErro = "";
			try
			{
				foreach (var item in contextoBD.Ambientes)
				{
					item.BD.abreConexao(out msg_erro_completo, out msg_erro_resumido);
				}

				// IMPORTANTE: o método abreConexao() faz com que a conexão seja aberta usando um novo objeto SqlConnection
				// Portanto, é fundamental recriar os objetos de acesso ao BD p/ que a propriedade Connection do SqlCommand esteja referenciando o mesmo SqlConnection,
				// caso contrário, ao executar uma operação de atualização dentro de uma transação irá ocorrer o erro:
				//		System.InvalidOperationException: A transação não está associada à conexão atual ou já foi concluída.
				// O problema foi percebido apenas com transações, pois as consultas continuaram funcionando normalmente.
				reinicializaObjetosUnitsDAO();

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ reiniciaBancoDados ]
		public static bool reiniciaBancoDados()
		{
			#region [ Declarações ]
			String strMsgErroLog = "";
			Log log = new Log();
			#endregion

			Global.gravaLogAtividade("Início da tentativa de reconectar com o Banco de Dados!!");

			#region [ Tenta fechar a conexão anterior ]
			try
			{
				foreach (var item in contextoBD.Ambientes)
				{
					if (item.BD != null)
					{
						if (item.BD.cnConexao.State != ConnectionState.Closed) item.BD.cnConexao.Close();
					}
				}
			}
			catch (Exception)
			{
				// NOP
			}
			#endregion

			#region [ Tenta abrir nova conexão ]
			try
			{
				foreach (var item in contextoBD.Ambientes)
				{
					item.BD.cnConexao = item.BD.getNovaConexao();
					Global.gravaLogAtividade("Sucesso ao estabelecer nova conexão com o banco de dados (ambiente: " + item.BD.NomeAmbiente + ")");
				}

				#region [ Reinicializa objetos ]
				reinicializaObjetosUnitsDAO();
				#endregion

				Global.gravaLogAtividade("Sucesso ao reconectar com o Banco de Dados (processo concluído)!!");

				#region [ Grava log no BD ]
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.RECONEXAO_BD;
				log.complemento = "Sucesso ao reconectar com o Banco de Dados";
				contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				return true;
			}
			catch (Exception)
			{
				Global.gravaLogAtividade("Falha ao tentar reconectar com o Banco de Dados!!");
				return false;
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoIbptCarregaArqCsv ]
		private void trataBotaoIbptCarregaArqCsv()
		{
			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!contextoBD.AmbienteBase.BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return;
				}
			}
			#endregion

			fIbpt = new FIbpt();
			fIbpt.Location = this.Location;
			fIbpt.Show();
			if (!fIbpt.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ trataBotaoAtualizaPlanilhaEstoque ]
		private void trataBotaoAtualizaPlanilhaEstoque()
		{
			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!contextoBD.AmbienteBase.BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return;
				}
			}
			#endregion

			#region [ Permissão de acesso ]
			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_ADM2_APP_ADMINISTRATIVO_ACESSO_AO_MODULO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			#endregion

			fAtualizaPlanilhaEstoque = new FAtualizaPlanilhaEstoque();
			fAtualizaPlanilhaEstoque.Location = this.Location;
			fAtualizaPlanilhaEstoque.Show();
			if (!fAtualizaPlanilhaEstoque.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ trataBotaoAnotarPedidosRecebidosCliente ]
		private void trataBotaoAnotarPedidosRecebidosCliente()
		{
			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!contextoBD.AmbienteBase.BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return;
				}
			}
			#endregion

			#region [ Permissão de acesso ]
			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_ADM2_ANOTAR_PEDIDOS_RECEBIDOS_PELO_CLIENTE))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			#endregion

			fAnotarPedidoRecebidoCliente = new FAnotarPedidoRecebidoCliente();
			fAnotarPedidoRecebidoCliente.Location = this.Location;
			fAnotarPedidoRecebidoCliente.Show();
			if (!fAnotarPedidoRecebidoCliente.ocorreuExceptionNaInicializacao) this.Visible = false;
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
				if (_OcorreuExceptionNaInicializacao) return;

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

					#region [ Registry: dados da sessão anterior ]
					strUltimoUsuario = (String)regKey.GetValue(Global.RegistryApp.Chaves.usuario, "");
					Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv = (String)regKey.GetValue(Global.RegistryApp.Chaves.FIbpt.pathIbptArquivoCsv, "");
					Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque = (String)regKey.GetValue(Global.RegistryApp.Chaves.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque, "");
					Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque = (String)regKey.GetValue(Global.RegistryApp.Chaves.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque, "");
					Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio = (String)regKey.GetValue(Global.RegistryApp.Chaves.FAnotarPedidoRecebidoCliente.pathArquivoRastreio, "");
					Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio = (String)regKey.GetValue(Global.RegistryApp.Chaves.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio, "");
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
							// No construtor de FMain, ao inicializar os ambientes de banco de dados, as conexões já foram abertas,
							// portanto, não é necessário fazer a chamada ao método iniciaBancoDados()
							#endregion

							#region [ Validação do usuário ]
							info(ModoExibicaoMensagemRodape.EmExecucao, "validando usuário");
							Global.Usuario.usuario = FLogin.usuario;
							Global.Usuario.senhaDigitada = FLogin.senha;

							#region [ Obtém dados no BD ]
							usuarioDAO = new UsuarioDAO(ref contextoBD.AmbienteBase.BD, Global.Usuario.usuario);
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
							Global.Acesso.listaOperacoesPermitidas = usuarioDAO.listaOperacoesPermitidas;
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
								if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_ADM2_APP_ADMINISTRATIVO_ACESSO_AO_MODULO))
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
					inicializaConstrutoresEstaticosUnitsDAO();
					#endregion

					#region [ Verifica data/hora da máquina local ]
					dtHrServidor = contextoBD.AmbienteBase.BD.obtemDataHoraServidor();
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
					versaoModulo = contextoBD.AmbienteBase.BD.getVersaoModulo("ADM2", out strMsgErro);
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
					log.operacao = Global.Cte.ADM2.LogOperacao.LOGON;
					log.complemento = "Logon realizado na máquina=" +
										System.Environment.MachineName +
										"; OS=" + System.Environment.OSVersion.VersionString +
										"; OS Version=" + System.Environment.OSVersion.Version +
										"; OS SP=" + System.Environment.OSVersion.ServicePack +
										"; Processor Count=" + System.Environment.ProcessorCount.ToString() +
										"; Windows User Name=" + System.Environment.UserName;
					contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
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
				regKey.SetValue(Global.RegistryApp.Chaves.FIbpt.pathIbptArquivoCsv, Global.Usuario.Defaults.FIbpt.pathIbptArquivoCsv);
				regKey.SetValue(Global.RegistryApp.Chaves.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque, Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.pathArquivoPlanilhaEstoque);
				regKey.SetValue(Global.RegistryApp.Chaves.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque, Global.Usuario.Defaults.FAtualizaPlanilhaEstoque.fileNameArquivoPlanilhaEstoque);
				regKey.SetValue(Global.RegistryApp.Chaves.FAnotarPedidoRecebidoCliente.pathArquivoRastreio, Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.pathArquivoRastreio);
				regKey.SetValue(Global.RegistryApp.Chaves.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio, Global.Usuario.Defaults.FAnotarPedidoRecebidoCliente.fileNameArquivoRastreio);
				#endregion

				#region [ Log em arquivo ]
				Global.gravaLogAtividade("Término do programa");
				Global.gravaLogAtividade(null);
				Global.gravaLogAtividade(null);
				#endregion

				#region [ Log de logoff realizado gravado no BD ]
				log = new Log();
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.LOGOFF;
				log.complemento = "Logoff após " + Global.formataDuracaoHMS(DateTime.Now - Global.dtHrInicioRefRelogioLocal);
				contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion
			}

			foreach (var item in contextoBD.Ambientes)
			{
				item.BD.fechaConexao();
			}
		}
		#endregion

		#endregion

		#region [ btnIbptCarregaArqCsv ]

		#region [ btnIbptCarregaArqCsv_Click ]
		private void btnIbptCarregaArqCsv_Click(object sender, EventArgs e)
		{
			trataBotaoIbptCarregaArqCsv();
		}
		#endregion

		#endregion

		#region [ btnAtualizarPlanilhaEstoque ]

		#region [ btnAtualizarPlanilhaEstoque_Click ]
		private void btnAtualizarPlanilhaEstoque_Click(object sender, EventArgs e)
		{
			trataBotaoAtualizaPlanilhaEstoque();
		}
		#endregion

		#endregion

		#region [ btnAnotarPedidosRecebidosCliente ]

		#region [ btnAnotarPedidosRecebidosCliente_Click ]
		private void btnAnotarPedidosRecebidosCliente_Click(object sender, EventArgs e)
		{
			trataBotaoAnotarPedidosRecebidosCliente();
		}
		#endregion

		#endregion

		#endregion
	}
}
