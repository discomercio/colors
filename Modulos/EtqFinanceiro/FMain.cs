#region [ using ]
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
#endregion

namespace EtqFinanceiro
{
    public partial class FMain : FModelo
    {
        #region [ Atributos ]
        private bool _inicializacaoOk;
        FEtiquetaDepositos fEtiquetasDepositos;
		FEtiquetaDepositosDesc fEtiquetasDepositosDesc;
		public static FMain fMain;
        #endregion

        #region [ Construtor ]
        public FMain()
        {
            InitializeComponent();

            fMain = this;
            if (!Directory.Exists(Global.Cte.LogAtividade.PathLogAtividade)) Directory.CreateDirectory(Global.Cte.LogAtividade.PathLogAtividade);
        }
        #endregion

        #region [ Métodos Privados ]

        #region [ iniciaBancoDados ]
        private bool iniciaBancoDados(ref string strMsgErro)
        {
            strMsgErro = "";
            try
            {
                BD.abreConexao();
                return true;
            }
            catch (Exception ex)
            {
                strMsgErro = ex.Message.ToString();
                return false;
            }
        }
        #endregion

        #region [ trataBotaoImprimirEtiquetasFin ]
        private void trataBotaoImprimirEtiquetasDepositos()
        {
            #region [ Verifica se a conexão c/ o BD está ok ]
            if (!BD.isConexaoOk())
            {
                if (!FMain.reiniciaBancoDados())
                {
                    avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
                    return;
                }
            }
            #endregion

            fEtiquetasDepositos = new FEtiquetaDepositos();
            fEtiquetasDepositos.Location = this.Location;
            fEtiquetasDepositos.Show();
            if (!fEtiquetasDepositos.ocorreuExceptionNaInicializacao) this.Visible = false;
        }
		#endregion

		#region [ trataBotaoImprimirEtiquetasFinDesc ]
		private void trataBotaoImprimirEtiquetasDepositosDesc()
		{
			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return;
				}
			}
			#endregion

			fEtiquetasDepositosDesc = new FEtiquetaDepositosDesc();
			fEtiquetasDepositosDesc.Location = this.Location;
			fEtiquetasDepositosDesc.Show();
			if (!fEtiquetasDepositosDesc.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#endregion

		#region [ Métodos Públicos ]

		#region [ reiniciaBandoDados ]
		public static bool reiniciaBancoDados()
        {
            #region [ Declarações ]
            string strMsgErroLog = "";
            Log log = new Log();
            #endregion

            Global.gravaLogAtividade("Início da tentativa de reconectar com o Bando de Dados!!");

            #region [ Tenta fechar a conexão anterior ]
            try
            {
                if (BD.cnConexao != null)
                {
                    if (BD.cnConexao.State != ConnectionState.Closed) BD.cnConexao.Close();
                }
            }
            catch (Exception)
            {
                //NOP
            }
            #endregion

            #region [ Tenta abrir nova conexão ]
            try
            {
                BD.cnConexao = BD.abreNovaConexao();
                Global.gravaLogAtividade("Sucesso ao estabelecer nova conexão!!");

                #region [ Reinicializa construtores estáticos ]
                LogDAO.inicializaConstrutorEstatico();
                UsuarioDAO.inicializaConstrutorEstatico();
                #endregion

                Global.gravaLogAtividade("Sucesso ao reconectar com o Bando de Dados (processo concluído)!!");

                #region [ Grava log no BD ]
                log.usuario = Global.Usuario.usuario;
                log.operacao = Global.Cte.EtqFinanceiro.LogOperacao.RECONEXAO_BD;
                log.complemento = "Sucesso ao reconectar com o Bando de Dados";
                LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
                #endregion

                return true;
            }
            catch (Exception)
            {
                Global.gravaLogAtividade("Falha ao tentar reconectar com o Bando de Dados!!");
                return false;
            }
            #endregion
        }
        #endregion

        #endregion

        #region [ Eventos ]

        #region [ FMain ]

        #region [ FMain_Shown ]
        private void FMain_Shown(object sender, EventArgs e)
        {
            #region [ Declarações ]
            string strSenhaDescriptografada = "";
            string strMsgErro = "";
            string strMsgErroLog = "";
            string strMsg;
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
                #region [ Executa rotinas de inicialização ]
                if (!_inicializacaoOk)
                {
#if (DESENVOLVIMENTO)
                    this.Text += " (Versão de DESENVOLVIMENTO)";
                    if (!confirma("Versão exclusiva de DESENVOLVIMENTO!!\nContinua assim mesmo?"))
                    {
                        Close();
                        return;
                    }
#elif (HOMOLOGACAO)
                    this.Text += " (Versão de HOMOLOGAÇÃO)";
                    if (!confirma("Versão exclusiva para o ambiente de HOMOLOGAÇÃO!!\nContinua assim mesmo?"))
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

                    #region [ Autenticação do usuário ]
                    do
                    {
                        blnValidacaoUsuarioOk = true;

                        #region [ Obtém login do usuário ]
                        fLogin.Location = new Point(this.Location.X + (this.Size.Width - fLogin.Size.Width) / 2, this.Location.Y + (this.Size.Height - fLogin.Size.Height) / 2);
                        drLogin = fLogin.ShowDialog();
                        if(drLogin != DialogResult.OK)
                        {
                            avisoErro("Login cancelado!!");
                            Close();
                            return;
                        }
                        #endregion

                        try
                        {
                            #region [ Inicia Banco de Dados ]
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
                            //Descriptografa a senha
                            if (!CriptoHex.decodificaDado(Global.Usuario.senhaCriptografada, ref strSenhaDescriptografada)) strSenhaDescriptografada = "";
                            Global.Usuario.senhaDescriptografada = strSenhaDescriptografada;
                            Global.Usuario.cadastrado = usuarioDAO.cadastrado;
                            Global.Usuario.bloqueado = usuarioDAO.bloqueado;
                            Global.Usuario.senhaExpirada = usuarioDAO.senhaExpirada;
                            #endregion

                            #region [ Usuário não cadastrado ]
                            if (blnValidacaoUsuarioOk)
                            {
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
                                if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_ETQFIN_APP_ETIQUETA_FIN_ACESSO_AO_MODULO))
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
                    #endregion

                    #region [ Verifica data/hora da máquina local ]
                    dtHrServidor = BD.obtemDataHoraServidor();
                    if (dtHrServidor != DateTime.MinValue)
                    {
                        if (Math.Abs(Global.calculaTimeSpanMinutos(DateTime.Now - dtHrServidor)) > 90)
                        {
                            strMsg = "O relógio desta máquina está defasado com relação ao servidor além do limite máximo tolerado:" +
                                     "\n\n" +
                                     "Horário do servidor: " + Global.formataDataDdMmYyyyHhMmSsComSeparador(dtHrServidor) +
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
                    versaoModulo = BD.getVersaoModulo("EtqFin", out strMsgErro);
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
                    log.operacao = Global.Cte.EtqFinanceiro.LogOperacao.LOGON;
                    log.complemento = "Logon realizado na máquina=" +
                                        System.Environment.MachineName +
                                        "; OS=" + System.Environment.OSVersion.VersionString +
                                        "; OS Version=" + System.Environment.OSVersion.Version +
                                        "; OS SP=" + System.Environment.OSVersion.ServicePack +
                                        "; Processor Count=" + System.Environment.ProcessorCount.ToString() +
                                        "; Windows User Name=" + System.Environment.UserName;
                    LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
                    #endregion

                    _inicializacaoOk = true;
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
                if (!_inicializacaoOk) Application.Exit();
            }
        }
        #endregion

        #region [ FMain_FormClosing ]
        private void FMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            #region [ Declarações ]
            Log log;
            string strMsgErroLog = "";
            #endregion

            if (_inicializacaoOk)
            {
                #region [ Grava log em arquivo ]
                Global.gravaLogAtividade("Término do programa");
                Global.gravaLogAtividade(null);
                Global.gravaLogAtividade(null);
                #endregion

                #region [ Log de logoff realizado gravado no BD ]
                log = new Log();
                log.usuario = Global.Usuario.usuario;
                log.operacao = Global.Cte.EtqFinanceiro.LogOperacao.LOGOFF;
                log.complemento = "Logoff após " + Global.formataDuracaoHMS(DateTime.Now - Global.dtHrInicioRefRelogioLocal);
                LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
                #endregion
            }
            BD.fechaConexao();
        }
        #endregion

        #endregion

        #region [ btnImprimirEtiquetasFin ]

        #region [ btnImprimirEtiquetasFin_Click ]
        private void btnImprimirEtiquetasFin_Click(object sender, EventArgs e)
        {
            trataBotaoImprimirEtiquetasDepositos();
        }
		#endregion

		#endregion

		#region [ btnImprimirEtiquetasFinDesc ]

		#region [ btnImprimirEtiquetasFinDesc_Click ]
		private void btnImprimirEtiquetasFinDesc_Click(object sender, EventArgs e)
		{
			trataBotaoImprimirEtiquetasDepositosDesc();
		}
		#endregion

		#endregion

		#endregion
	}
}
