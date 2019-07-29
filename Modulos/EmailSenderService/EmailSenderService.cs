using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EmailSenderService
{
	public partial class EmailSenderService : ServiceBase
	{
		#region [ Atributos ]
		// Singleton
		public static readonly EmailSenderService emailSenderService = new EmailSenderService();
		private static readonly Thread _thrEnviaEmail = new Thread(new ThreadStart(rotinaThreadEnviaEmail));
		#endregion

		#region [ Getters/Setters ]
		public static EmailSenderService getInstance()
		{
			// Singleton
			return emailSenderService;
		}

		private static bool _isFinalizacaoExecutada = false;
		public static bool isFinalizacaoExecutada
		{
			get { return EmailSenderService._isFinalizacaoExecutada; }
			set { EmailSenderService._isFinalizacaoExecutada = value; }
		}

		private static bool _isOnStopAcionado = false;
		public static bool isOnStopAcionado
		{
			get { return _isOnStopAcionado; }
			set { _isOnStopAcionado = value; }
		}

		private static bool _isOnShutdownAcionado = false;
		public static bool isOnShutdownAcionado
		{
			get { return _isOnShutdownAcionado; }
			set { _isOnShutdownAcionado = value; }
		}

		private static bool _isThreadEnviaEmailIniciada = false;
		public static bool isThreadEnviaEmailIniciada
		{
			get { return _isThreadEnviaEmailIniciada; }
			set { _isThreadEnviaEmailIniciada = value; }
		}

		private static bool _isThreadEnviaEmailEncerrada = false;
		public static bool isThreadEnviaEmailEncerrada
		{
			get { return _isThreadEnviaEmailEncerrada; }
			set { _isThreadEnviaEmailEncerrada = value; }
		}

		private static DateTime _dtHrInicioServico = DateTime.Now;
		public static DateTime dtHrInicioServico
		{
			get { return EmailSenderService._dtHrInicioServico; }
		}
		#endregion

		#region [ Construtor ]
		// Construtor private devido ao Singleton
		private EmailSenderService()
		{
			InitializeComponent();
		}
		#endregion

		#region [ OnStart ]
		protected override void OnStart(string[] args)
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "OnStart()";
			#endregion

			/* Observação: Não tratar 'Exceptions', pois ao iniciar o serviço, a ocorrência
			 * =========== de um exception será automaticamente registrada no event viewer e
			 * irá fazer com que o SCM (Service Control Manager) perceba que o serviço falhou
			 * ao iniciar!!
			 */

			Global.gravaEventLog(strNomeDestaRotina + "\r\n" + "Método OnStart() acionado\r\nAmbiente: " + BD.strDescricaoAmbiente.ToUpper(), EventLogEntryType.Information);

			executaThreadEnviaEmail();
		}
		#endregion

		#region [ OnStop ]
		protected override void OnStop()
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "OnStop()";
			String strMsg;
			#endregion

			try
			{
				isOnStopAcionado = true;
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + "Método OnStop() acionado", EventLogEntryType.Information);

				finalizaExecucao();
			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Error);
			}
		}
		#endregion

		#region [ OnShutdown ]
		protected override void OnShutdown()
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "OnShutdown()";
			String strMsg;
			#endregion

			try
			{
				isOnShutdownAcionado = true;
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + "Método OnShutdown() acionado", EventLogEntryType.Information);

				base.OnShutdown();
				finalizaExecucao();
			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Error);
			}
		}
		#endregion

		#region [ inicializaConstrutoresEstaticosUnitsDAO ]
		private static bool inicializaConstrutoresEstaticosUnitsDAO()
		{
			try
			{
				ComumDAO.inicializaConstrutorEstatico();
				EmailSndSvcDAO.inicializaConstrutorEstatico();
				LogDAO.inicializaConstrutorEstatico();
				LogErroDAO.inicializaConstrutorEstatico();
				MensagemDAO.inicializaConstrutorEstatico();
				ParametroDAO.inicializaConstrutorEstatico();
				RemetenteDAO.inicializaConstrutorEstatico();
				return true;
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao inicializar os objetos estáticos das units de acesso ao Banco de Dados!!\n" + ex.Message);
				return false;
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
		private static bool iniciaBancoDados(ref String strMsgErroCompleto, ref String strMsgErroResumido)
		{
			strMsgErroCompleto = "";
			strMsgErroResumido = "";
			try
			{
				BD.abreConexao();
				return true;
			}
			catch (Exception ex)
			{
				strMsgErroCompleto = ex.ToString();
				strMsgErroResumido = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ finalizaExecucao ]
		private static void finalizaExecucao()
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "finalizaExecucao()";
			String strMsg;
			#endregion

			#region [ Rotina de finalização já foi chamada? ]
			/* É necessário garantir que ServiceBase.Stop() é chamada apenas 1 vez,
			 * pois cada chamada aciona o método OnStop(), o que causaria um loop infinito!!
			 */
			if (isFinalizacaoExecutada) return;
			isFinalizacaoExecutada = true;
			#endregion

			try
			{
				aguardaEncerramento();
				BD.fechaConexao();
			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Error);
			}
			finally
			{
				emailSenderService.Stop();
			}
		}
		#endregion

		#region [ aguardaEncerramento ]
		private static void aguardaEncerramento()
		{
			#region [ Declarações ]
			DateTime dtHrInicioEspera;
			#endregion

			if (_thrEnviaEmail == null) return;

			dtHrInicioEspera = DateTime.Now;
			while (isThreadEnviaEmailIniciada && (!isThreadEnviaEmailEncerrada))
			{
				Thread.Sleep(500);
				// Timeout?
				if (dtHrInicioEspera.AddSeconds(3 * 60) < DateTime.Now)
				{
					_thrEnviaEmail.Abort();
					break;
				}
			}
		}
		#endregion

		#region [ ProcessaSleep ]
		private static void ProcessaSleep(double tempoEmMiliSegundos)
		{
			#region [ Declarações ]
			DateTime dtHrInicio = DateTime.Now;
			#endregion

			while (dtHrInicio.AddMilliseconds(tempoEmMiliSegundos) > DateTime.Now)
			{
				#region [ Serviço deve parar? ]
				if (isOnShutdownAcionado) break;
				if (isOnStopAcionado) break;
				#endregion

				Thread.Sleep(500);
			}
		}
		#endregion

		#region [ executaThreadEnviaEmail ]
		static void executaThreadEnviaEmail()
		{
			/* Observação: Não tratar 'Exceptions', pois ao iniciar o serviço, a ocorrência
			 * =========== de um exception será automaticamente registrada no event viewer e
			 * irá fazer com que o SCM (Service Control Manager) perceba que o serviço falhou
			 * ao iniciar!!
			 */

			if (isThreadEnviaEmailIniciada) return;
			isThreadEnviaEmailIniciada = true;

			_thrEnviaEmail.IsBackground = true;
			_thrEnviaEmail.Priority = ThreadPriority.Normal;
			_thrEnviaEmail.Start();
		}
		#endregion

		#region [ rotinaThreadEnviaEmail ]
		static void rotinaThreadEnviaEmail()
		{
			#region [ Declarações ]
			const String strNomeDestaRotina = "rotinaThreadEnviaEmail()";
			int intDuracaoPausaInicializacaoEmSegundos;
			String strMsg;
			String strMsgErro = "";
			String strMsgErroCompleto = "";
			String strMsgErroResumido = "";
			String strNumeroVersaoProducao = "";
			DateTime dtHrUltProcClientesEmAtraso = DateTime.MinValue;
			DateTime dtHrUltVerificacaoProcClientesEmAtraso = DateTime.MinValue;
			DateTime dtHrUltCargaArqRetornoBoleto = DateTime.MinValue;
			DateTime dtHrUltManutencaoArqLogAtividade = DateTime.MinValue;
			DateTime dtHrUltSinalVida = DateTime.MinValue;
			DateTime dtHrInicioPausaInicializacao;
			int intQtdeTentativas = 0;
			long lngSegundosDecorridos;
			#endregion

			try
			{
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + "Thread de envio de e-mails iniciada!", EventLogEntryType.Information);

				#region [ Aguarda alguns instantes antes de tentar conectar ao BD ]
				intDuracaoPausaInicializacaoEmSegundos = 120;
				strMsg = "Início da contagem da pausa de inicialização: " + intDuracaoPausaInicializacaoEmSegundos.ToString() + " segundos (" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ")";
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Information);
				dtHrInicioPausaInicializacao = DateTime.Now;
				while (true)
				{
					#region [ Serviço deve parar? ]
					if (isOnShutdownAcionado) return;
					if (isOnStopAcionado) return;
					#endregion

					lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioPausaInicializacao);
					if (lngSegundosDecorridos > intDuracaoPausaInicializacaoEmSegundos) break;
					Thread.Sleep(1000);
				} // while (true)
				strMsg = "Fim da contagem da pausa de inicialização (" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ")";
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Information);
				#endregion

				#region [ Conecta ao BD ]
				// Como o serviço do SQL Server pode estar na mesma máquina e, ao reiniciar a máquina,
				// o banco de dados pode ainda não estar pronto para receber conexões, são feitas 
				// algumas tentativas antes de desistir.
				while (true)
				{
					#region [ Serviço deve parar? ]
					if (isOnShutdownAcionado) return;
					if (isOnStopAcionado) return;
					#endregion

					intQtdeTentativas++;
					if (iniciaBancoDados(ref strMsgErroCompleto, ref strMsgErroResumido))
					{
						break;
					}
					else
					{
						if (intQtdeTentativas >= 5)
						{
							#region [ Não conseguiu conectar ao BD, então pára o serviço! ]
							strMsg = "Falha ao conectar ao banco de dados (Server=" + BD.strServidor + ", BD=" + BD.strNomeBancoDados + ")!!\r\n" + strMsgErroResumido;
							throw new Exception(strMsg);
							#endregion
						}
						else
						{
							strMsg = "Falha ao conectar ao BD (tentativa " + intQtdeTentativas.ToString() + "): " + strMsgErroResumido;
							Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Information);
							strMsg = strNomeDestaRotina + ": " + "Falha ao conectar ao BD (tentativa " + intQtdeTentativas.ToString() + "): " + strMsgErroCompleto;
							Global.gravaLogAtividade(strMsg);
						}
						ProcessaSleep(60 * 1000);
					}
				} // while (true)

				strMsg = "Conectado ao banco de dados: Server=" + BD.strServidor + ", BD=" + BD.strNomeBancoDados;
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Information);
				#endregion

				#region [ Inicializa construtores estáticos ]
				inicializaConstrutoresEstaticosUnitsDAO();
				#endregion

				#region [ Validação da versão deste programa ]
				if (!BD.obtemNumeroVersaoProducaoModuloEmailSndSvc(out strNumeroVersaoProducao, out strMsgErro))
				{
					#region [ Falha ao obter nº da versão ]
					strMsg = "Falha ao tentar obter no banco de dados o número da versão em produção deste aplicativo!!\n" + strMsgErro;
					throw new Exception(strMsg);
					#endregion
				}

				if (!strNumeroVersaoProducao.Equals(Global.Cte.Aplicativo.VERSAO_NUMERO))
				{
					#region [ Versão inválida! ]
					strMsg = "Versão inválida do aplicativo!!\n\nVersão deste programa: " + Global.Cte.Aplicativo.VERSAO_NUMERO + "\nVersão permitida: " + strNumeroVersaoProducao;
					throw new Exception(strMsg);
					#endregion
				}
				#endregion

				#region [ Grava log no BD informando que serviço foi iniciado com sucesso ]
				strMsg = "Serviço do Windows iniciado '" + Global.Cte.Aplicativo.ID_SISTEMA_EMAILSENDER + "' (" + Global.Cte.Aplicativo.VERSAO + ")";
				ComumDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_EMAILSENDERSERVICE_INICIADO, strMsg, out strMsgErro);
				#endregion

				#region [ Log informando o valor dos parâmetros ]
				strMsg = "Parâmetros em uso:" +
						"\r\nEmailSndSvc_FlagHabilitacao = " + EmailManager.flagRotinaEnvioEmailHabilitado.ToString() +
						"\r\nEmailSndSvc_PeriodoSuspensao = " + EmailManager.periodoSuspensao +
						"\r\nEmailSndSvc_IntervaloMinEmSegundosEntreMsgs = " + EmailManager.intervaloMinEmSegundosEntreMsgs.ToString() +
						"\r\nEmailSndSvc_IntervaloEmSegundosAposCicloOcioso = " + EmailManager.intervaloEmSegundosAposCicloOcioso.ToString() +
						"\r\nEmailSndSvc_QtdeMaxTentativas = " + EmailManager.qtdeMaxTentativas.ToString() +
						"\r\nEmailSndSvc_IntervaloMinEmSegundos_Tentativa_1_2 = " + EmailManager.intervaloMinEmSegundos_Tentativa_1_2.ToString() +
						"\r\nEmailSndSvc_IntervaloMinEmSegundos_Tentativa_2_3 = " + EmailManager.intervaloMinEmSegundos_Tentativa_2_3.ToString() +
						"\r\nEmailSndSvc_IntervaloMinEmSegundos_Tentativa_Demais = " + EmailManager.intervaloMinEmSegundos_Tentativa_Demais.ToString();
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Information);
				#endregion

				#region [ Memória: recupera do BD data/hora da última execução das diversas rotinas ]
				dtHrUltManutencaoArqLogAtividade = ComumDAO.getCampoDataTabelaParametro(Global.Cte.EMAILSND.ID_T_PARAMETRO.DT_HR_ULT_MANUTENCAO_ARQ_LOG_ATIVIDADE);
				#endregion

				#region [ Laço permanente de execução ]
				try
				{
					while (true)
					{
						try
						{
							#region [ Serviço deve parar? ]
							if (isOnShutdownAcionado) return;
							if (isOnStopAcionado) return;
							#endregion

							#region [ Sinaliza periodicamente no log em arquivo que o serviço está em execução ]
							if (((DateTime.Now.Minute % 10) == 0) && (Global.calculaTimeSpanMinutos(DateTime.Now - dtHrUltSinalVida) >= 9))
							{
								dtHrUltSinalVida = DateTime.Now;
								Global.gravaLogAtividade("Mensagem informativa: serviço ativo e em execução há " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicioServico) + " (iniciado em " + Global.formataDataDdMmYyyyHhMmSsComSeparador(dtHrInicioServico) + ")");
							}
							#endregion

							#region [ Apaga os arquivos de log de atividade antigos? ]
							try
							{
								if ((dtHrUltManutencaoArqLogAtividade == DateTime.MinValue) ||
									((DateTime.Now.Hour == 1) && (DateTime.Now.Minute >= 20) && (dtHrUltManutencaoArqLogAtividade.DayOfYear != DateTime.Now.DayOfYear)))
								{
									Global.executaManutencaoArqLogAtividade(out strMsgErro);
									dtHrUltManutencaoArqLogAtividade = DateTime.Now;
									ComumDAO.setCampoDataTabelaParametro(Global.Cte.EMAILSND.ID_T_PARAMETRO.DT_HR_ULT_MANUTENCAO_ARQ_LOG_ATIVIDADE, dtHrUltManutencaoArqLogAtividade);
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							EmailManager.processaMensagens();
						}
						catch (Exception ex)
						{
							strMsg = ex.ToString();
							Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Error);
						}

						ProcessaSleep(1000);
					} // while (true)
				}
				finally
				{
					#region [ Grava log no BD informando que serviço foi encerrado ]
					strMsg = "Serviço do Windows encerrado '" + Global.Cte.Aplicativo.ID_SISTEMA_EMAILSENDER;
					ComumDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_EMAILSENDERSERVICE_ENCERRADO, strMsg, out strMsgErro);
					#endregion
				}
				#endregion

			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + strMsg, EventLogEntryType.Error);
				isThreadEnviaEmailEncerrada = true;
				finalizaExecucao();
			}
			finally
			{
				Global.gravaEventLog(strNomeDestaRotina + "\r\n" + "Thread de envio de e-mail encerrada!", EventLogEntryType.Information);
				isThreadEnviaEmailEncerrada = true;
			}
		}
		#endregion
	}
}
