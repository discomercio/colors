#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
#endregion

namespace FinanceiroService
{
	public partial class FinanceiroService : ServiceBase
	{
		#region [ Atributos ]
		// Singleton
		public static readonly FinanceiroService financeiroService = new FinanceiroService();
		private static readonly Thread _thrManutencao = new Thread(new ThreadStart(rotinaThreadManutencao));
		private static DateTime _dtHrUltReinicializaObjetosEstaticosUnitsDAO = DateTime.MinValue;
		#endregion

		#region [ Getters/Setters ]
		public static FinanceiroService getInstance()
		{
			// Singleton
			return financeiroService;
		}

		private static bool _isFinalizacaoExecutada = false;
		public static bool isFinalizacaoExecutada
		{
			get { return FinanceiroService._isFinalizacaoExecutada; }
			set { FinanceiroService._isFinalizacaoExecutada = value; }
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

		private static bool _isThreadManutencaoIniciada = false;
		public static bool isThreadManutencaoIniciada
		{
			get { return _isThreadManutencaoIniciada; }
			set { _isThreadManutencaoIniciada = value; }
		}

		private static bool _isThreadManutencaoEncerrada = false;
		public static bool isThreadManutencaoEncerrada
		{
			get { return _isThreadManutencaoEncerrada; }
			set { _isThreadManutencaoEncerrada = value; }
		}

		private static DateTime _dtHrInicioServico = DateTime.Now;
		public static DateTime dtHrInicioServico
		{
			get { return FinanceiroService._dtHrInicioServico; }
		}
		#endregion

		#region [ Construtor ]
		// Construtor private devido ao Singleton
		private FinanceiroService()
		{
			InitializeComponent();
		}
		#endregion

		#region [ OnStart ]
		protected override void OnStart(string[] args)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "OnStart()";
			#endregion

			/* Observação: Não tratar 'Exceptions', pois ao iniciar o serviço, a ocorrência
			 * =========== de um exception será automaticamente registrada no event viewer e
			 * irá fazer com que o SCM (Service Control Manager) perceba que o serviço falhou
			 * ao iniciar!!
			 */

			Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + "Método OnStart() acionado\r\nAmbiente: " + BD.strDescricaoAmbiente.ToUpper(), EventLogEntryType.Information);

			executaThreadManutencao();
		}
		#endregion

		#region [ OnStop ]
		protected override void OnStop()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "OnStop()";
			String strMsg;
			#endregion

			try
			{
				isOnStopAcionado = true;
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + "Método OnStop() acionado", EventLogEntryType.Information);

				finalizaExecucao();
			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
			}
		}
		#endregion

		#region [ OnShutdown ]
		protected override void OnShutdown()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "OnShutdown()";
			String strMsg;
			#endregion

			try
			{
				isOnShutdownAcionado = true;
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + "Método OnShutdown() acionado", EventLogEntryType.Information);

				base.OnShutdown();
				finalizaExecucao();
			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
			}
		}
		#endregion

		#region [ inicializaConstrutoresEstaticosUnitsDAO ]
		private static bool inicializaConstrutoresEstaticosUnitsDAO()
		{
			try
			{
				GeralDAO.inicializaConstrutorEstatico();
				EstoqueDAO.inicializaConstrutorEstatico();
				PedidoDAO.inicializaConstrutorEstatico();
				ClearsaleDAO.inicializaConstrutorEstatico();
				BraspagDAO.inicializaConstrutorEstatico();
				EmailSndSvcDAO.inicializaConstrutorEstatico();
				EmailCtrlDAO.inicializaConstrutorEstatico();
				ClienteDAO.inicializaConstrutorEstatico();
				LancamentoFluxoCaixaDAO.inicializaConstrutorEstatico();
				FinLogDAO.inicializaConstrutorEstatico();
				PlanoContasDAO.inicializaConstrutorEstatico();
				_dtHrUltReinicializaObjetosEstaticosUnitsDAO = DateTime.Now;
				return true;
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao inicializar os objetos estáticos das units de acesso ao Banco de Dados!!\n" + ex.Message);
				return false;
			}
		}
		#endregion

		#region [ reinicializaObjetosEstaticosUnitsDAO ]
		private static bool reinicializaObjetosEstaticosUnitsDAO()
		{
			try
			{
				GeralDAO.inicializaObjetosEstaticos();
				EstoqueDAO.inicializaObjetosEstaticos();
				PedidoDAO.inicializaObjetosEstaticos();
				ClearsaleDAO.inicializaObjetosEstaticos();
				BraspagDAO.inicializaObjetosEstaticos();
				EmailSndSvcDAO.inicializaObjetosEstaticos();
				EmailCtrlDAO.inicializaObjetosEstaticos();
				ClienteDAO.inicializaObjetosEstaticos();
				LancamentoFluxoCaixaDAO.inicializaObjetosEstaticos();
				FinLogDAO.inicializaObjetosEstaticos();
				PlanoContasDAO.inicializaObjetosEstaticos();
				Global.gravaLogAtividade("Sucesso ao reinicializar os objetos estáticos das units de acesso ao Banco de Dados!!");
				_dtHrUltReinicializaObjetosEstaticosUnitsDAO = DateTime.Now;
				return true;
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao reinicializar os objetos estáticos das units de acesso ao Banco de Dados!!\n" + ex.Message);
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

		#region [ reiniciaBancoDados ]
		public static bool reiniciaBancoDados()
		{
			#region [ Declarações ]
			String strMsgErroLog = "";
			#endregion

			Global.gravaLogAtividade("Início da tentativa de reconectar com o Banco de Dados!!");

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
				// NOP
			}
			#endregion

			#region [ Tenta abrir nova conexão ]
			try
			{
				BD.cnConexao = BD.abreNovaConexao();
				Global.gravaLogAtividade("Sucesso ao estabelecer nova conexão!!");
				reinicializaObjetosEstaticosUnitsDAO();
				Global.gravaLogAtividade("Sucesso ao reconectar com o Banco de Dados (processo concluído)!!");

				#region [ Grava log no BD ]
				GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_RECONEXAO_BD, "Sucesso ao reconectar com o Banco de Dados", out strMsgErroLog);
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

		#region [ finalizaExecucao ]
		private static void finalizaExecucao()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "finalizaExecucao()";
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
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
			}
			finally
			{
				financeiroService.Stop();
			}
		}
		#endregion

		#region [ aguardaEncerramento ]
		private static void aguardaEncerramento()
		{
			#region [ Declarações ]
			DateTime dtHrInicioEspera;
			#endregion

			if (_thrManutencao == null) return;

			dtHrInicioEspera = DateTime.Now;
			while (isThreadManutencaoIniciada && (!isThreadManutencaoEncerrada))
			{
				Thread.Sleep(500);
				// Timeout?
				if (dtHrInicioEspera.AddSeconds(3 * 60) < DateTime.Now)
				{
					_thrManutencao.Abort();
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

		#region [ executaThreadManutencao ]
		static void executaThreadManutencao()
		{
			/* Observação: Não tratar 'Exceptions', pois ao iniciar o serviço, a ocorrência
			 * =========== de um exception será automaticamente registrada no event viewer e
			 * irá fazer com que o SCM (Service Control Manager) perceba que o serviço falhou
			 * ao iniciar!!
			 */

			if (isThreadManutencaoIniciada) return;
			isThreadManutencaoIniciada = true;

			_thrManutencao.IsBackground = true;
			_thrManutencao.Priority = ThreadPriority.Normal;
			_thrManutencao.Start();
		}
		#endregion

		#region [ rotinaThreadManutencao ]
		static void rotinaThreadManutencao()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "rotinaThreadManutencao()";
			const int INTERVALO_ENTRE_ATUALIZACAO_LEITURA_PARAMETRO_EM_SEG = 5 * 60;
			int qtdeClearsalePedidosNovosEnviados;
			int qtdeClearsalePedidosFalhaEnvio;
			int qtdeClearsalePedidosResultadoProcessado;
			int qtdeEstornosPendentesVerificados;
			int qtdeEstornosConfirmados;
			int qtdeEstornosAbortados;
			int intDuracaoPausaInicializacaoEmSegundos;
			int id_emailsndsvc_mensagem;
			List<int> listaIdNfeEmitente;
			bool blnEmailAlertaEnviado;
			String strParametro;
			String strDestinatario;
			String strAux;
			String strMsg;
			String strMsgErro = "";
			String strMsgErroAux;
			String strMsgErroCompleto = "";
			String strMsgErroResumido = "";
			String strMsgInfo;
			String strMsgInfoCancelAutoPedidos;
			String strMsgInfoBpCsAntifraudeClearsale;
			String strMsgInfoEstornosPendentes;
			String strMsgInformativa;
			String strLogFalha;
			StringBuilder sbMsgParametros = new StringBuilder("");
			String strSubject;
			String strBody;
			string[] vAux;
			DateTime dtHrInicioFinanceiroService = DateTime.Now;
			DateTime dtHrUltProcClientesEmAtraso = DateTime.MinValue;
			DateTime dtHrUltVerificacaoProcClientesEmAtraso = DateTime.MinValue;
			DateTime dtHrUltCargaArqRetornoBoleto = DateTime.MinValue;
			DateTime dtHrUltManutencaoArqLogAtividade = DateTime.MinValue;
			DateTime dtHrUltManutencaoBdLogAntigo = DateTime.MinValue;
			DateTime dtHrUltCancelamentoAutomaticoPedidos = DateTime.MinValue;
			DateTime dtHrUltProcessamentoBpCsAntifraudeClearsale = DateTime.MinValue;
			DateTime dtHrUltProcBpCsBraspagAtualizaStatusTransacoesPendentes = DateTime.MinValue;
			DateTime dtHrUltProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto = DateTime.MinValue;
			DateTime dtHrUltProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto = DateTime.MinValue;
			DateTime dtHrUltProcEnviarEmailAlertaPedidoNovoAnaliseCredito = DateTime.MinValue;
			DateTime dtHrUltProcWebhookBraspag = DateTime.MinValue;
			DateTime dtHrUltProcEstornosPendentes = DateTime.MinValue;
			DateTime dtHrUltConsultaExecucaoSolicitadaProcProdutosVendidosSemPresencaEstoque = DateTime.MinValue;
			DateTime dtHrUltSinalVida = DateTime.MinValue;
			DateTime dtHrInicioPausaInicializacao;
			DateTime dtHrInicioProcessamento;
			DateTime dtHrUltLeituraParametro;
			DateTime dtHrUltMsgInatividadeBpCs = DateTime.MinValue;
			DateTime dtHrUltMsgInatividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito = DateTime.MinValue;
			DateTime dtHrUltBpCsTransicaoPeriodoInatividade = DateTime.MinValue;
			DateTime dtHrUltProcEnvioEmailAlertaPedidoNovoAnaliseCreditoTransicaoPeriodoInatividade = DateTime.MinValue;
			DateTime dtHrUltProcWebhookBraspagTransicaoPeriodoInatividade = DateTime.MinValue;
			DateTime dtHrUltProcEstornosPendentesTransicaoPeriodoInatividade = DateTime.MinValue;
			DateTime dtHrUltMsgInatividadeProcWebhookBraspag = DateTime.MinValue;
			DateTime dtHrUltMsgInatividadeProcEstornosPendentes = DateTime.MinValue;
			DateTime dtHrUltVerificacaoConexaoBd = DateTime.MinValue;
			DateTime dtHrUltLimpezaSessionToken = DateTime.MinValue;
			DateTime dtHrUltUploadFileManutencaoArquivos = DateTime.MinValue;
			bool blnFlag;
			bool blnFlagPeriodoInatividadeAux;
			bool blnPeriodoAtividadeAux;
			bool blnPeriodoAtividadeBpCs = true;
			bool blnPeriodoAtividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito = true;
			bool blnPeriodoAtividadeProcWebhookBraspag = true;
			bool blnPeriodoAtividadeProcEstornosPendentes = true;
			int intQtdeTentativas = 0;
			int intParametro;
			long lngSegundosDecorridos;
			long lngDuracaoProcessamentoEmSegundos;
			TimeSpan tsParametro;
			TimeSpan tsParametroInicioAux;
			TimeSpan tsParametroTerminoAux;
			TimeSpan tsParametroHorarioAux;
			RegistroTabelaParametro parametro;
			VersaoModulo versaoModulo;
			FinSvcLog svcLog;
			PlanoContasConta planoContasConta;
			List<NfeEmitente> listaNfeEmitente;
			#endregion

			try
			{
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + "Thread de manutenção iniciada!", EventLogEntryType.Information);

				#region [ Aguarda alguns instantes antes de tentar conectar ao BD ]
				intDuracaoPausaInicializacaoEmSegundos = 120;
				strMsg = "Início da contagem da pausa de inicialização: " + intDuracaoPausaInicializacaoEmSegundos.ToString() + " segundos (" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ")";
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
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
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
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
						if (intQtdeTentativas >= 10)
						{
							#region [ Não conseguiu conectar ao BD, então pára o serviço! ]
							strMsg = "Falha ao conectar ao banco de dados (Server=" + BD.strServidor + ", BD=" + BD.strNomeBancoDados + ")!!\r\n" + strMsgErroResumido;
							throw new Exception(strMsg);
							#endregion
						}
						else
						{
							strMsg = "Falha ao conectar ao BD (tentativa " + intQtdeTentativas.ToString() + "): " + strMsgErroResumido;
							Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
							strMsg = NOME_DESTA_ROTINA + ": " + "Falha ao conectar ao BD (tentativa " + intQtdeTentativas.ToString() + "): " + strMsgErroCompleto;
							Global.gravaLogAtividade(strMsg);
						}
						ProcessaSleep(60 * 1000);
					}
				} // while (true)

				strMsg = "Conectado ao banco de dados: Server=" + BD.strServidor + ", BD=" + BD.strNomeBancoDados;
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
				#endregion

				#region [ Inicializa construtores estáticos ]
				if (inicializaConstrutoresEstaticosUnitsDAO())
				{
					Global.gravaLogAtividade("Sucesso na inicialização dos objetos estáticos das units de acesso ao Banco de Dados!!");
				}
				else
				{
					Global.gravaLogAtividade("Falha na inicialização dos objetos estáticos das units de acesso ao Banco de Dados!!");
				}
				#endregion

				#region [ Validação da versão deste programa ]
				versaoModulo = BD.getVersaoModulo("FINSVC", out strMsgErro);
				if (versaoModulo == null)
				{
					#region [ Falha ao obter nº da versão ]
					strMsg = "Falha ao tentar obter no banco de dados o número da versão em produção deste aplicativo!!\n" + strMsgErro;
					throw new Exception(strMsg);
					#endregion
				}

				Global.Cte.Aplicativo.IDENTIFICADOR_AMBIENTE_OWNER = versaoModulo.identificador_ambiente;

				if (!versaoModulo.versao.Equals(Global.Cte.Aplicativo.VERSAO_NUMERO))
				{
					#region [ Versão inválida! ]
					strMsg = "Versão inválida do aplicativo!!\n\nVersão deste programa: " + Global.Cte.Aplicativo.VERSAO_NUMERO + "\nVersão permitida: " + versaoModulo.versao;
					throw new Exception(strMsg);
					#endregion
				}
				#endregion

				#region [ Parâmetros de acesso ao ambiente da Braspag estão ok? ]
				if ((Global.Cte.Braspag.WS_ENDERECO_PAGADOR_TRANSACTION ?? "").Trim().Length == 0)
				{
					strMsg = "Endereço do web service da Braspag (Pagador Transaction) não está definido!!";
					throw new Exception(strMsg);
				}

				if ((Global.Cte.Braspag.WS_ENDERECO_PAGADOR_QUERY ?? "").Trim().Length == 0)
				{
					strMsg = "Endereço do web service da Braspag (Pagador Query) não está definido!!";
					throw new Exception(strMsg);
				}
				#endregion

				#region [ Parâmetros de acesso ao ambiente da Clearsale estão ok? ]
				if ((Global.Cte.Clearsale.CS_ENTITY_CODE ?? "").Trim().Length == 0)
				{
					strMsg = "O 'Entity Code' da Clearsale não está definido!!";
					throw new Exception(strMsg);
				}

				if ((Global.Cte.Clearsale.WS_CS_ENDERECO_SERVICE ?? "").Trim().Length == 0)
				{
					strMsg = "Endereço do web service da Clearsale (Service) não está definido!!";
					throw new Exception(strMsg);
				}

				if ((Global.Cte.Clearsale.WS_CS_ENDERECO_EXTENDED_SERVICE ?? "").Trim().Length == 0)
				{
					strMsg = "Endereço do web service da Clearsale (Extended Service) não está definido!!";
					throw new Exception(strMsg);
				}
				#endregion

				#region [ Braspag Webhook: dados do plano de contas p/ gravar lançamento do boleto de e-commerce ]
				// Completa os dados do plano de contas obtendo o código do grupo a partir do cadastro do plano de contas no banco de dados
				foreach (var item in Global.Parametros.Braspag.webhookBraspagPlanoContasBoletoECList)
				{
					planoContasConta = PlanoContasDAO.getPlanoContasContaById(item.id_plano_contas_conta, out strMsgErroAux);
					if (planoContasConta != null) item.id_plano_contas_grupo = planoContasConta.id_plano_contas_grupo;
				}
				#endregion

				#region [ Log informativo: parâmetros Braspag/Clearsale ]
				strMsg = "Parâmetros do sistema (" + Global.Cte.Aplicativo.AMBIENTE_EXECUCAO + ")";
				sbMsgParametros.AppendLine(strMsg);
				strMsg = "Braspag (Pagador Transaction): " + Global.Cte.Braspag.WS_ENDERECO_PAGADOR_TRANSACTION;
				sbMsgParametros.AppendLine(strMsg);
				strMsg = "Braspag (Pagador Query): " + Global.Cte.Braspag.WS_ENDERECO_PAGADOR_QUERY;
				sbMsgParametros.AppendLine(strMsg);
				strMsg = "Clearsale (Entity Code): " + Global.Cte.Clearsale.CS_ENTITY_CODE;
				sbMsgParametros.AppendLine(strMsg);
				strMsg = "Clearsale (Service): " + Global.Cte.Clearsale.WS_CS_ENDERECO_SERVICE;
				sbMsgParametros.AppendLine(strMsg);
				strMsg = "Clearsale (Extended Service): " + Global.Cte.Clearsale.WS_CS_ENDERECO_EXTENDED_SERVICE;
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Log informativo: endereços de remetente/destinatário dos emails de alerta ]
				strMsg = "E-mails de alerta do sistema:" +
						"\r\n    RemetenteMsgAlertaSistema=" + Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA +
						"\r\n    DestinatarioMsgAlertaSistema=" + Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA;
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Memória: recupera do BD data/hora da última execução das diversas rotinas ]
				dtHrUltProcClientesEmAtraso = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROC_CLIENTES_EM_ATRASO);
				dtHrUltManutencaoArqLogAtividade = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_MANUTENCAO_ARQ_LOG_ATIVIDADE);
				dtHrUltManutencaoBdLogAntigo = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_MANUTENCAO_BD_LOG_ANTIGO);
				dtHrUltCancelamentoAutomaticoPedidos = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_CANCELAMENTO_AUTOMATICO_PEDIDOS);
				dtHrUltProcessamentoBpCsAntifraudeClearsale = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_ANTIFRAUDE_CLEARSALE);
				dtHrUltProcBpCsBraspagAtualizaStatusTransacoesPendentes = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_ATUALIZA_STATUS_TRANSACOES_PENDENTES);
				dtHrUltProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO);
				dtHrUltProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO);
				dtHrUltProcEnviarEmailAlertaPedidoNovoAnaliseCredito = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO);
				dtHrUltProcWebhookBraspag = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_WEBHOOK_BRASPAG);
				dtHrUltProcEstornosPendentes = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_ESTORNOS_PENDENTES);
				dtHrUltLimpezaSessionToken = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_LIMPEZA_SESSION_TOKEN);
				dtHrUltUploadFileManutencaoArquivos = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_UPLOAD_FILE_MANUTENCAO_ARQUIVOS);
				dtHrUltConsultaExecucaoSolicitadaProcProdutosVendidosSemPresencaEstoque = GeralDAO.getCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_CONSULTA_EXECUCAO_SOLICITADA_PROC_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE);
				#endregion

				#region [ Leitura de parâmetros do BD ]

				#region [ Limpeza de Session Token ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.LIMPEZA_SESSION_TOKEN_HORARIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.SessionToken_Limpeza_Horario = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_LIMPEZA_SESSION_TOKEN);
				blnFlag = (intParametro != 0) ? true : false;
				if (Global.Parametros.Geral.SessionToken_Limpeza_Horario == TimeSpan.MinValue) blnFlag = false;
				Global.Parametros.Geral.SessionToken_Limpeza_FlagHabilitacao = blnFlag;
				strMsg = "Status da rotina de limpeza de session token: " +
						(blnFlag ? "ativado" : "desativado") +
						" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.SessionToken_Limpeza_Horario, "(nenhum)") + ")";
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Manutenção de arquivos salvos no servidor através da WebAPI (UploadFile) ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.UPLOAD_FILE_MANUTENCAO_ARQUIVOS_HORARIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_UPLOAD_FILE_MANUTENCAO_ARQUIVOS);
				blnFlag = (intParametro != 0) ? true : false;
				if (Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario == TimeSpan.MinValue) blnFlag = false;
				Global.Parametros.Geral.UploadFile_ManutencaoArquivos_FlagHabilitacao = blnFlag;
				strMsg = "Status da rotina de manutenção de arquivos salvos no servidor através da WebAPI (UploadFile): " +
						(blnFlag ? "ativado" : "desativado") +
						" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario, "(nenhum)") + ")";
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Flag de habilitação da rotina de cancelamento automático de pedidos ]
				strAux = "";
				for (int i = 0; i < Global.Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas.Count; i++)
				{
					if (Global.Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas[i] > 0)
					{
						if (strAux.Length > 0) strAux += ", ";
						strAux += Global.Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas[i].ToString();
					}
				}
				if (strAux.Length == 0) strAux = "(nenhuma)";
				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_CANCELAMENTO_AUTOMATICO_PEDIDOS);
				blnFlag = (intParametro != 0) ? true : false;
				Global.Parametros.Geral.ExecutarCancelamentoAutomaticoPedidos = blnFlag;
				strMsg = "Status da rotina de cancelamento automático de pedidos: " +
						(blnFlag ? "ativado" : "desativado") +
						" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.HorarioCancelamentoAutomaticoPedidos, "(nenhum)") + ", lojas ignoradas: " + strAux + ")";
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Flag de habilitação do processamento dos produtos vendidos sem presença no estoque ]
				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE);
				blnFlag = (intParametro != 0) ? true : false;
				Global.Parametros.Geral.ProcessamentoProdutosVendidosSemPresencaEstoque_FlagHabilitacao = blnFlag;
				strMsg = "Flag de habilitação do processamento dos produtos vendidos sem presença no estoque: " +
						(blnFlag ? "ativado" : "desativado");
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Tempo (em segundos) entre verificações se há solicitação de execução do processamento de produtos vendidos sem presença no estoque ]
				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.CONSULTA_EXECUCAO_SOLICITADA_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE_EM_SEG, Global.Parametros.Geral.ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg);
				Global.Parametros.Geral.ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg = intParametro;
				strMsg = "Parâmetro: tempo entre verificações se há solicitação de execução do processamento de produtos vendidos sem presença no estoque (em seg) = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Atualização do status de transações Braspag pendentes ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_PROCESSAMENTO_ATUALIZA_STATUS_TRANSACOES_PENDENTES_HORARIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_BRASPAG_ATUALIZA_STATUS_TRANSACOES_PENDENTES);
				blnFlag = (intParametro != 0) ? true : false;
				if (Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario == TimeSpan.MinValue) blnFlag = false;
				Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagAtualizaStatusTransacoesPendentes = blnFlag;
				strMsg = "Status da rotina de atualização de status das transações Braspag pendentes: " +
						(blnFlag ? "ativado" : "desativado") +
						" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario, "(nenhum)") + ")";
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Processamento das requisições de estorno pendentes (Getnet) ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS);
				if ((strParametro ?? "").Trim().Length == 0) strParametro = Global.Parametros.Geral.DESTINATARIO_PADRAO_MSG_ALERTA_SISTEMA;
				Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS = strParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES);
				blnFlag = (intParametro != 0) ? true : false;
				Global.Parametros.Braspag.ExecutarProcessamentoBpCsEstornosPendentes = blnFlag;
				strMsg = "Processamento da verificação de estornos pendentes (Braspag): " +
						(blnFlag ? "ativado" : "desativado") +
						", destinatário(s): " + (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS.Length > 0 ? Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS : "(nenhum)");
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS, Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS);
				Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS = intParametro;
				strMsg = "Parâmetro (Braspag): prazo máximo para verificação dos estornos pendentes (em dias) = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_TEMPO_ENTRE_PROCESSAMENTO_EM_SEG, Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg);
				Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg = intParametro;
				strMsg = "Parâmetro (Braspag): tempo entre processamento de estornos pendentes (em seg) = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				#region [ Parâmetros do período de inatividade do processamento de estornos pendentes ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE_HORARIO_INICIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio = tsParametro;

				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE_HORARIO_TERMINO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE);
				blnFlag = (intParametro != 0) ? true : false;
				if ((Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio == TimeSpan.MinValue)
					||
					(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino == TimeSpan.MinValue))
				{
					blnFlag = false;
				}
				Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao = blnFlag;

				if (Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao)
				{
					strMsg = "Suspender processamento de estornos pendentes durante o período de inatividade (" +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio, "(nenhum)") +
						" às " +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino, "(nenhum)") +
						"): " + (blnFlag ? "ativado" : "desativado");
				}
				else
				{
					strMsg = "Suspender processamento de estornos pendentes durante o período de inatividade: " + (blnFlag ? "ativado" : "desativado");
				}

				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#endregion

				#region [ Envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO_HORARIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario = tsParametro;

				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO);
				Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = strParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO);
				blnFlag = (intParametro != 0) ? true : false;
				if (Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario == TimeSpan.MinValue) blnFlag = false;
				if (Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO.Length == 0) blnFlag = false;
				Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagEnviarEmailAlertaTransacoesPendentesProxCancelAuto = blnFlag;
				strMsg = "Envio de email de alerta sobre transações pendentes da Braspag próximas do cancelamento automático: " +
						(blnFlag ? "ativado" : "desativado") +
						" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario, "(nenhum)") + ")" +
						", destinatário(s): " + (Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO.Length > 0 ? Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO : "(nenhum)");
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Captura transação pendente devido prazo final de cancelamento automático ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO_HORARIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario = tsParametro;

				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO);
				Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = strParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO);
				blnFlag = (intParametro != 0) ? true : false;
				if (Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario == TimeSpan.MinValue) blnFlag = false;
				if (Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO.Length == 0) blnFlag = false;
				Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto = blnFlag;
				strMsg = "Processamento da captura automática de transações pendentes devido ao prazo final do cancelamento automático: " +
						(blnFlag ? "ativado" : "desativado") +
						" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario, "(nenhum)") + ")" +
						", destinatário(s): " + (Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO.Length > 0 ? Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO : "(nenhum)");
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#region [ Envio de email de alerta sobre novo pedido cadastrado aguardando tratamento da análise de crédito ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO);
				Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = strParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO);
				blnFlag = (intParametro != 0) ? true : false;
				if (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO.Length == 0) blnFlag = false;
				Global.Parametros.Geral.ExecutarProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito = blnFlag;
				strMsg = "Envio de email de alerta sobre novo pedido cadastrado aguardando tratamento da análise de crédito: " +
						(blnFlag ? "ativado" : "desativado") +
						", destinatário(s): " + (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO.Length > 0 ? Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO : "(nenhum)");
				sbMsgParametros.AppendLine(strMsg);

				#region [ Parâmetros do período de inatividade do processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE_HORARIO_INICIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio = tsParametro;

				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE_HORARIO_TERMINO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE);
				blnFlag = (intParametro != 0) ? true : false;
				if ((Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio == TimeSpan.MinValue)
					||
					(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino == TimeSpan.MinValue))
				{
					blnFlag = false;
				}
				Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao = blnFlag;

				if (Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao)
				{
					strMsg = "Suspender processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito durante o período de inatividade (" +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio, "(nenhum)") +
						" às " +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino, "(nenhum)") +
						"): " + (blnFlag ? "ativado" : "desativado");
				}
				else
				{
					strMsg = "Suspender processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito durante o período de inatividade: " + (blnFlag ? "ativado" : "desativado");
				}
				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#endregion

				#region [ Processamento dos dados recebidos pelo Webhook Braspag ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG);
				if ((strParametro ?? "").Trim().Length == 0) strParametro = Global.Parametros.Geral.DESTINATARIO_PADRAO_MSG_ALERTA_SISTEMA;
				Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG = strParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG);
				blnFlag = (intParametro != 0) ? true : false;
				Global.Parametros.Geral.ExecutarProcessamentoWebhookBraspag = blnFlag;
				strMsg = "Processamento dos dados recebidos pelo Webhook Braspag: " +
						(blnFlag ? "ativado" : "desativado") +
						", destinatário(s): " + (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG.Length > 0 ? Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG : "(nenhum)");
				sbMsgParametros.AppendLine(strMsg);

				#region [ Parâmetros do período de inatividade do processamento dos dados recebidos pelo Webhook Braspag ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE_HORARIO_INICIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio = tsParametro;

				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE_HORARIO_TERMINO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE);
				blnFlag = (intParametro != 0) ? true : false;
				if ((Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio == TimeSpan.MinValue)
					||
					(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino == TimeSpan.MinValue))
				{
					blnFlag = false;
				}
				Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao = blnFlag;

				if (Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao)
				{
					strMsg = "Suspender processamento dos dados recebidos pelo Webhook Braspag durante o período de inatividade (" +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio, "(nenhum)") +
						" às " +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino, "(nenhum)") +
						"): " + (blnFlag ? "ativado" : "desativado");
				}
				else
				{
					strMsg = "Suspender processamento dos dados recebidos pelo Webhook Braspag durante o período de inatividade: " + (blnFlag ? "ativado" : "desativado");
				}
				sbMsgParametros.AppendLine(strMsg);

				#region [ Parâmetros Braspag MerchantId para serem usados no processamento dos dados recebidos pelo Webhook ]
				strMsg = "Parâmetros Braspag MerchantId para ser usado no processamento dos dados recebidos pelo Webhook:";
				sbMsgParametros.AppendLine(strMsg);
				foreach (var item in Global.Parametros.Braspag.webhookBraspagMerchantIdList)
				{
					strMsg = "    " + item.Empresa + " = " + item.MerchantId;
					sbMsgParametros.AppendLine(strMsg);
				}
				#endregion

				#region [ Parâmetros de plano de contas para gravação de lançamentos no fluxo de caixa devido aos boletos de e-commerce (Webhook Braspag) ]
				strMsg = "Parâmetros de plano de contas para gravação de lançamentos no fluxo de caixa dos boletos de e-commerce (Webhook Braspag):";
				sbMsgParametros.AppendLine(strMsg);
				foreach (var item in Global.Parametros.Braspag.webhookBraspagPlanoContasBoletoECList)
				{
					strMsg = "    " + item.Empresa + ": id_conta_corrente=" + item.id_conta_corrente.ToString() + ", id_plano_contas_empresa=" + item.id_plano_contas_empresa.ToString() + ", id_plano_contas_conta=" + item.id_plano_contas_conta.ToString();
					sbMsgParametros.AppendLine(strMsg);
				}
				#endregion

				#endregion

				#endregion

				#region [ Parâmetros para processamento Clearsale ]
				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_ANTIFRAUDE_CLEARSALE);
				blnFlag = (intParametro != 0) ? true : false;
				Global.Parametros.Clearsale.ExecutarProcessamentoBpCsAntifraudeClearsale = blnFlag;
				strMsg = "Status da rotina de processamento de transações com a Clearsale: " + (blnFlag ? "ativado" : "desativado");
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_MAX_TENTATIVAS_TX_ANTIFRAUDE, Global.Parametros.Clearsale.MaxTentativasEnvioTransacao);
				Global.Parametros.Clearsale.MaxTentativasEnvioTransacao = intParametro;
				strMsg = "Parâmetro (Clearsale): quantidade máxima de tentativas de envio da transação = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_TEMPO_MIN_ENTRE_TENTATIVAS_EM_SEG, Global.Parametros.Clearsale.TempoMinEntreTentativasEmSeg);
				Global.Parametros.Clearsale.TempoMinEntreTentativasEmSeg = intParametro;
				strMsg = "Parâmetro (Clearsale): tempo mínimo entre tentativas de envio da transação (em seg) = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_TEMPO_ENTRE_PROCESSAMENTO_EM_SEG, Global.Parametros.Clearsale.TempoEntreProcessamentoEmSeg);
				Global.Parametros.Clearsale.TempoEntreProcessamentoEmSeg = intParametro;
				strMsg = "Parâmetro (Clearsale): tempo entre processamento de transações (em seg) = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_TEMPO_MAX_CLIENTE_TOTALIZAR_PAGTO_EM_SEG, Global.Parametros.Clearsale.TempoMaxClienteTotalizarPagtoEmSeg);
				Global.Parametros.Clearsale.TempoMaxClienteTotalizarPagtoEmSeg = intParametro;
				strMsg = "Parâmetro (Clearsale): tempo máximo para o cliente realizar as transações de pagamento até totalizar o valor esperado antes de enviar a transação (em seg) = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_MAX_QTDE_FALHAS_CONSECUTIVAS_METODO_GETRETURNANALYSIS, Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis);
				Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis = intParametro;
				strMsg = "Parâmetro (Clearsale): quantidade máxima de falhas consecutivas na chamada ao método GetReturnAnalysis() antes de enviar mensagem de alerta = " + intParametro.ToString();
				sbMsgParametros.AppendLine(strMsg);

				#region [ Parâmetros do período de inatividade do processamento c/ a Clearsale ]
				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE_HORARIO_INICIO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio = tsParametro;

				strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE_HORARIO_TERMINO);
				tsParametro = Global.converteHhMmParaTimeSpan(strParametro);
				Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino = tsParametro;

				intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE);
				blnFlag = (intParametro != 0) ? true : false;
				if ((Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio == TimeSpan.MinValue)
					||
					(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino == TimeSpan.MinValue))
				{
					blnFlag = false;
				}
				Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao = blnFlag;

				if (Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao)
				{
					strMsg = "Suspender processamento com a Clearsale durante o período de inatividade (" +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio, "(nenhum)") +
						" às " +
						Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino, "(nenhum)") +
						"): " + (blnFlag ? "ativado" : "desativado");
				}
				else
				{
					strMsg = "Suspender processamento com a Clearsale durante o período de inatividade: " + (blnFlag ? "ativado" : "desativado");
				}

				sbMsgParametros.AppendLine(strMsg);
				#endregion

				#endregion

				// Log informativo com os parâmetros
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + sbMsgParametros.ToString(), EventLogEntryType.Information);

				dtHrUltLeituraParametro = DateTime.Now;
				#endregion

				#region [ Grava log no BD informando que serviço foi iniciado com sucesso (tabela de log geral) ]
				strMsg = "Serviço do Windows iniciado '" + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + "' (" + Global.Cte.Aplicativo.VERSAO + ")" +
						"\r\n" +
						"Ambiente de " + BD.strDescricaoAmbiente + " (Banco de dados: Server=" + BD.strServidor + ", BD=" + BD.strNomeBancoDados + ")" +
						"\r\n" +
						sbMsgParametros.ToString();
				GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_FINANCEIROSERVICE_INICIADO, strMsg, out strMsgErro);
				#endregion

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = strMsg;
				GeralDAO.gravaFinSvcLog(svcLog, out strMsgErroAux);
				#endregion

				#region [ Envia email de alerta ]
				// Somente é possível inserir o email na fila de envio após a conexão ao BD estar disponível
				strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Serviço " + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + " iniciado às " + Global.formataDataDdMmYyyyHhMmSsComSeparador(dtHrInicioFinanceiroService);
				strBody = "Mensagem de " + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + " (" + Global.Cte.Aplicativo.VERSAO + ")" + ": o serviço foi iniciado às " + Global.formataDataDdMmYyyyHhMmSsComSeparador(dtHrInicioFinanceiroService) + " no ambiente de " + BD.strDescricaoAmbiente + " (Banco de dados: Server=" + BD.strServidor + ", BD=" + BD.strNomeBancoDados + ")" +
							"\r\n\r\n" +
							sbMsgParametros.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
					Global.gravaLogAtividade(strMsg);
				}
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

							#region [ Periodicamente verifica se conexão com o BD está saudável ]
							lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltVerificacaoConexaoBd);
							if (lngSegundosDecorridos >= 60)
							{
								dtHrUltVerificacaoConexaoBd = DateTime.Now;
								if (BD.isConexaoOk())
								{
									// NOP
								}
								else
								{
									Global.gravaLogAtividade("Status da conexão com o BD: FALHA");
									if (!reiniciaBancoDados())
									{
										Global.gravaLogAtividade("Falha ao tentar reiniciar conexão com o BD");
										ProcessaSleep(60 * 1000);
										continue;
									}
									else
									{
										Global.gravaLogAtividade("Conexão com o BD foi reiniciada com sucesso!!");
									}
								}
							}
							#endregion

							#region [ Periodicamente reinicializa os 'prepared statements' usados com o BD ]
							// O objetivo é minimizar o risco do serviço ficar inoperante devido a algum problema com esses objetos
							lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - _dtHrUltReinicializaObjetosEstaticosUnitsDAO);
							if (lngSegundosDecorridos >= (60 * 60))
							{
								Global.gravaLogAtividade("Reinicialização preventiva automática dos objetos estáticos das units de acesso ao Banco de Dados!!");
								if (reinicializaObjetosEstaticosUnitsDAO())
								{
									_dtHrUltReinicializaObjetosEstaticosUnitsDAO = DateTime.Now;
									Global.gravaLogAtividade("Sucesso na reinicialização preventiva automática dos objetos estáticos das units de acesso ao Banco de Dados!!");
								}
								else
								{
									Global.gravaLogAtividade("Falha na reinicialização preventiva automática dos objetos estáticos das units de acesso ao Banco de Dados!!");
								}
								// Se a conexão não estiver ok, volta ao início do laço p/ reprocessar a rotina de reconexão automática
								if (!BD.isConexaoOk()) continue;
							}
							#endregion

							#region [ Periodicamente atualiza a leitura dos parâmetros ]
							lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltLeituraParametro);
							if (lngSegundosDecorridos >= INTERVALO_ENTRE_ATUALIZACAO_LEITURA_PARAMETRO_EM_SEG)
							{
								#region [ Parâmetro: flag de habilitação da rotina de limpeza de session token ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.LIMPEZA_SESSION_TOKEN_HORARIO);
								tsParametroHorarioAux = Global.converteHhMmParaTimeSpan(strParametro);
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_LIMPEZA_SESSION_TOKEN);
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Geral.SessionToken_Limpeza_FlagHabilitacao != blnFlag)
									||
									(tsParametroHorarioAux != Global.Parametros.Geral.SessionToken_Limpeza_Horario))
								{
									Global.Parametros.Geral.SessionToken_Limpeza_FlagHabilitacao = blnFlag;
									Global.Parametros.Geral.SessionToken_Limpeza_Horario = tsParametroHorarioAux;
									strMsg = "Rotina de limpeza de session token (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.SessionToken_Limpeza_Horario, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação da rotina de manutenção de arquivos (UploadFile) ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.UPLOAD_FILE_MANUTENCAO_ARQUIVOS_HORARIO);
								tsParametroHorarioAux = Global.converteHhMmParaTimeSpan(strParametro);
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_UPLOAD_FILE_MANUTENCAO_ARQUIVOS);
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Geral.UploadFile_ManutencaoArquivos_FlagHabilitacao != blnFlag)
									||
									(tsParametroHorarioAux != Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario))
								{
									Global.Parametros.Geral.UploadFile_ManutencaoArquivos_FlagHabilitacao = blnFlag;
									Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario = tsParametroHorarioAux;
									strMsg = "Rotina de manutenção de arquivos salvos no servidor através da WebAPI (UploadFile) (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação do cancelamento automático de pedidos ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_CANCELAMENTO_AUTOMATICO_PEDIDOS);
								blnFlag = (intParametro != 0 ? true : false);
								if (Global.Parametros.Geral.ExecutarCancelamentoAutomaticoPedidos != blnFlag)
								{
									Global.Parametros.Geral.ExecutarCancelamentoAutomaticoPedidos = blnFlag;
									strMsg = "Rotina de cancelamento automático de pedidos (alteração de status): " + (blnFlag ? "ativado" : "desativado");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação do processamento dos produtos vendidos sem presença no estoque ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE);
								blnFlag = (intParametro != 0 ? true : false);
								if (Global.Parametros.Geral.ProcessamentoProdutosVendidosSemPresencaEstoque_FlagHabilitacao != blnFlag)
								{
									Global.Parametros.Geral.ProcessamentoProdutosVendidosSemPresencaEstoque_FlagHabilitacao = blnFlag;
									strMsg = "Rotina de processamento dos produtos vendidos sem presença no estoque (alteração de status): " + (blnFlag ? "ativado" : "desativado");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Tempo (em segundos) entre verificações se há solicitação de execução do processamento de produtos vendidos sem presença no estoque ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.CONSULTA_EXECUCAO_SOLICITADA_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE_EM_SEG, Global.Parametros.Geral.ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg);
								if (Global.Parametros.Geral.ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg != intParametro)
								{
									Global.Parametros.Geral.ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg = intParametro;
									strMsg = "Parâmetro: alteração do tempo entre verificações se há solicitação de execução do processamento de produtos vendidos sem presença no estoque (em seg) = " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação da rotina de atualização de status das transações Braspag pendentes ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_PROCESSAMENTO_ATUALIZA_STATUS_TRANSACOES_PENDENTES_HORARIO);
								tsParametroHorarioAux = Global.converteHhMmParaTimeSpan(strParametro);
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_BRASPAG_ATUALIZA_STATUS_TRANSACOES_PENDENTES);
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagAtualizaStatusTransacoesPendentes != blnFlag)
									||
									(tsParametroHorarioAux != Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario))
								{
									Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagAtualizaStatusTransacoesPendentes = blnFlag;
									Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario = tsParametroHorarioAux;
									strMsg = "Rotina de atualização de status das transações Braspag pendentes (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação da rotina de envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO_HORARIO);
								tsParametroHorarioAux = Global.converteHhMmParaTimeSpan(strParametro);
								strDestinatario = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO);
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO);
								if (tsParametroHorarioAux == TimeSpan.MinValue) intParametro = 0;
								if (strDestinatario.Length == 0) intParametro = 0;
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagEnviarEmailAlertaTransacoesPendentesProxCancelAuto != blnFlag)
									||
									(!strDestinatario.Equals(Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO))
									||
									(tsParametroHorarioAux != Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario))
								{
									Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagEnviarEmailAlertaTransacoesPendentesProxCancelAuto = blnFlag;
									Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario = tsParametroHorarioAux;
									Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = strDestinatario;
									strMsg = "Rotina de envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario, "(nenhum)") + ")" +
											", destinatário(s): " + (Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO.Length > 0 ? Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO : "(nenhum)");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação da rotina de captura automática de transação pendente devido ao prazo final de cancelamento automático ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO_HORARIO);
								tsParametroHorarioAux = Global.converteHhMmParaTimeSpan(strParametro);
								strDestinatario = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_BRASPAG_DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO);
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO);
								if (tsParametroHorarioAux == TimeSpan.MinValue) intParametro = 0;
								if (strDestinatario.Length == 0) intParametro = 0;
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto != blnFlag)
									||
									(!strDestinatario.Equals(Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO))
									||
									(tsParametroHorarioAux != Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario))
								{
									Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto = blnFlag;
									Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario = tsParametroHorarioAux;
									Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = strDestinatario;
									strMsg = "Rotina de captura automática de transação pendente devido ao prazo final de cancelamento automático (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											" (horário programado: " + Global.formataTimeSpanHorario(Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario, "(nenhum)") + ")" +
											", destinatário(s): " + (Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO.Length > 0 ? Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO : "(nenhum)");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: flag de habilitação da rotina de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito ]
								strDestinatario = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO);
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO);
								if (strDestinatario.Length == 0) intParametro = 0;
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Geral.ExecutarProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito != blnFlag)
									||
									(!strDestinatario.Equals(Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO)))
								{
									Global.Parametros.Geral.ExecutarProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito = blnFlag;
									Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = strDestinatario;
									strMsg = "Rotina de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											", destinatário(s): " + (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO.Length > 0 ? Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO : "(nenhum)");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetros do período de inatividade do processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE_HORARIO_INICIO);
								tsParametroInicioAux = Global.converteHhMmParaTimeSpan(strParametro);

								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE_HORARIO_TERMINO);
								tsParametroTerminoAux = Global.converteHhMmParaTimeSpan(strParametro);

								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE);
								blnFlagPeriodoInatividadeAux = (intParametro != 0) ? true : false;

								if ((tsParametroInicioAux == TimeSpan.MinValue) || (tsParametroTerminoAux == TimeSpan.MinValue)) blnFlagPeriodoInatividadeAux = false;

								if ((blnFlagPeriodoInatividadeAux != Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao)
									||
									(tsParametroInicioAux != Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio)
									||
									(tsParametroTerminoAux != Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino)
								)
								{
									strMsg = "Período de inatividade do processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito alterado de " +
											(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao ? "'habilitado'" : "'desabilitado'") +
											" (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")" +
											" para " +
											(blnFlagPeriodoInatividadeAux ? "'habilitado'" : "'desabilitado'") +
											" (" + Global.formataTimeSpanHorario(tsParametroInicioAux, "(nenhum)") + " às " + Global.formataTimeSpanHorario(tsParametroTerminoAux, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);

									Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao = blnFlagPeriodoInatividadeAux;
									Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio = tsParametroInicioAux;
									Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino = tsParametroTerminoAux;
								}
								#endregion

								#region [ Parâmetro: flag de habilitação do processamento dos dados recebidos pelo Webhook Braspag ]
								strDestinatario = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG);
								if ((strDestinatario ?? "").Trim().Length == 0) strDestinatario = Global.Parametros.Geral.DESTINATARIO_PADRAO_MSG_ALERTA_SISTEMA;
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG);
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Geral.ExecutarProcessamentoWebhookBraspag != blnFlag)
									||
									(!strDestinatario.Equals(Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG)))
								{
									Global.Parametros.Geral.ExecutarProcessamentoWebhookBraspag = blnFlag;
									Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG = strDestinatario;
									strMsg = "Processamento dos dados recebidos pelo Webhook Braspag (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											", destinatário(s): " + (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG.Length > 0 ? Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG : "(nenhum)");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetros do período de inatividade do processamento de dados recebidos pelo Webhook Braspag ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE_HORARIO_INICIO);
								tsParametroInicioAux = Global.converteHhMmParaTimeSpan(strParametro);

								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE_HORARIO_TERMINO);
								tsParametroTerminoAux = Global.converteHhMmParaTimeSpan(strParametro);

								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE);
								blnFlagPeriodoInatividadeAux = (intParametro != 0) ? true : false;

								if ((tsParametroInicioAux == TimeSpan.MinValue) || (tsParametroTerminoAux == TimeSpan.MinValue)) blnFlagPeriodoInatividadeAux = false;

								if ((blnFlagPeriodoInatividadeAux != Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao)
									||
									(tsParametroInicioAux != Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio)
									||
									(tsParametroTerminoAux != Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino)
								)
								{
									strMsg = "Período de inatividade do processamento de dados recebidos pelo Webhook Braspag alterado de " +
											(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao ? "'habilitado'" : "'desabilitado'") +
											" (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")" +
											" para " +
											(blnFlagPeriodoInatividadeAux ? "'habilitado'" : "'desabilitado'") +
											" (" + Global.formataTimeSpanHorario(tsParametroInicioAux, "(nenhum)") + " às " + Global.formataTimeSpanHorario(tsParametroTerminoAux, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);

									Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao = blnFlagPeriodoInatividadeAux;
									Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio = tsParametroInicioAux;
									Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino = tsParametroTerminoAux;
								}
								#endregion

								#region [ Parâmetro: (Clearsale) flag de habilitação do processamento antifraude ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_ANTIFRAUDE_CLEARSALE);
								blnFlag = (intParametro != 0 ? true : false);
								if (Global.Parametros.Clearsale.ExecutarProcessamentoBpCsAntifraudeClearsale != blnFlag)
								{
									Global.Parametros.Clearsale.ExecutarProcessamentoBpCsAntifraudeClearsale = blnFlag;
									strMsg = "Rotina de processamento de transações com a Clearsale (alteração de status): " + (blnFlag ? "ativado" : "desativado");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: (Clearsale) quantidade máxima de tentativas de envio da transação ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_MAX_TENTATIVAS_TX_ANTIFRAUDE, Global.Parametros.Clearsale.MaxTentativasEnvioTransacao);
								if (Global.Parametros.Clearsale.MaxTentativasEnvioTransacao != intParametro)
								{
									Global.Parametros.Clearsale.MaxTentativasEnvioTransacao = intParametro;
									strMsg = "Parâmetro (Clearsale): alteração da quantidade máxima de tentativas de envio da transação = " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: (Clearsale) intervalo mínimo entre tentativas ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_TEMPO_MIN_ENTRE_TENTATIVAS_EM_SEG, Global.Parametros.Clearsale.TempoMinEntreTentativasEmSeg);
								if (Global.Parametros.Clearsale.TempoMinEntreTentativasEmSeg != intParametro)
								{
									Global.Parametros.Clearsale.TempoMinEntreTentativasEmSeg = intParametro;
									strMsg = "Parâmetro (Clearsale): alteração do tempo mínimo entre tentativas de envio da transação (em seg) = " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: (Clearsale) intervalo entre cada execução do processamento ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_TEMPO_ENTRE_PROCESSAMENTO_EM_SEG, Global.Parametros.Clearsale.TempoEntreProcessamentoEmSeg);
								if (Global.Parametros.Clearsale.TempoEntreProcessamentoEmSeg != intParametro)
								{
									Global.Parametros.Clearsale.TempoEntreProcessamentoEmSeg = intParametro;
									strMsg = "Parâmetro (Clearsale): alteração do tempo entre processamento de transações (em seg) = " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: (Clearsale) tempo máximo que o cliente tem para totalizar o pagamento antes de enviar a transação para a Clearsale ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_TEMPO_MAX_CLIENTE_TOTALIZAR_PAGTO_EM_SEG, Global.Parametros.Clearsale.TempoMaxClienteTotalizarPagtoEmSeg);
								if (Global.Parametros.Clearsale.TempoMaxClienteTotalizarPagtoEmSeg != intParametro)
								{
									Global.Parametros.Clearsale.TempoMaxClienteTotalizarPagtoEmSeg = intParametro;
									strMsg = "Parâmetro (Clearsale): alteração do tempo máximo para o cliente realizar as transações de pagamento até totalizar o valor esperado antes de enviar a transação (em seg) = " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro: (Clearsale) quantidade máxima de falhas consecutivas na chamada ao método Clearsale GetReturnAnalysis ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_CLEARSALE_MAX_QTDE_FALHAS_CONSECUTIVAS_METODO_GETRETURNANALYSIS, Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis);
								if (Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis != intParametro)
								{
									Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis = intParametro;
									strMsg = "Parâmetro (Clearsale): alteração na quantidade máxima de falhas consecutivas na chamada ao método GetReturnAnalysis() antes de enviar mensagem de alerta = " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetros do período de inatividade do processamento com a Clearsale ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE_HORARIO_INICIO);
								tsParametroInicioAux = Global.converteHhMmParaTimeSpan(strParametro);

								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE_HORARIO_TERMINO);
								tsParametroTerminoAux = Global.converteHhMmParaTimeSpan(strParametro);

								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE);
								blnFlagPeriodoInatividadeAux = (intParametro != 0) ? true : false;

								if ((tsParametroInicioAux == TimeSpan.MinValue) || (tsParametroTerminoAux == TimeSpan.MinValue)) blnFlagPeriodoInatividadeAux = false;

								if ((blnFlagPeriodoInatividadeAux != Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao)
									||
									(tsParametroInicioAux != Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio)
									||
									(tsParametroTerminoAux != Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino)
								)
								{
									strMsg = "Período de inatividade com a Clearsale alterado de " +
										(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao ? "'habilitado'" : "'desabilitado'") +
										" (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")" +
										" para " +
										(blnFlagPeriodoInatividadeAux ? "'habilitado'" : "'desabilitado'") +
										" (" + Global.formataTimeSpanHorario(tsParametroInicioAux, "(nenhum)") + " às " + Global.formataTimeSpanHorario(tsParametroTerminoAux, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);

									Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao = blnFlagPeriodoInatividadeAux;
									Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio = tsParametroInicioAux;
									Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino = tsParametroTerminoAux;
								}
								#endregion

								#region [ Parâmetro: (Braspag) flag de habilitação do processamento de estornos pendentes ]
								strDestinatario = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS);
								if ((strDestinatario ?? "").Trim().Length == 0) strDestinatario = Global.Parametros.Geral.DESTINATARIO_PADRAO_MSG_ALERTA_SISTEMA;
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES);
								blnFlag = (intParametro != 0 ? true : false);
								if ((Global.Parametros.Braspag.ExecutarProcessamentoBpCsEstornosPendentes != blnFlag)
									||
									(!strDestinatario.Equals(Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS)))
								{
									Global.Parametros.Braspag.ExecutarProcessamentoBpCsEstornosPendentes = blnFlag;
									Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS = strDestinatario;
									strMsg = "Processamento da verificação de estornos pendentes (alteração da configuração): " +
											(blnFlag ? "ativado" : "desativado") +
											", destinatário(s): " + (Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS.Length > 0 ? Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS : "(nenhum)");
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}
								#endregion

								#region [ Parâmetro (Braspag): prazo máximo para verificação dos estornos pendentes ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS, Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS);
								if (Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS != intParametro)
								{
									strMsg = "Parâmetro (Braspag): alteração do prazo máximo de verificação dos estornos pendentes (em dias) de " + Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS.ToString() + " para " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
									Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS = intParametro;
								}
								#endregion

								#region [ Parâmetro: (Braspag) intervalo entre cada execução do processamento de estornos pendentes ]
								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_TEMPO_ENTRE_PROCESSAMENTO_EM_SEG, Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg);
								if (Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg != intParametro)
								{
									strMsg = "Parâmetro (Braspag): alteração do tempo entre processamento de estornos pendentes (em seg) de " + Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg.ToString() + " para " + intParametro.ToString();
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
									Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg = intParametro;
								}
								#endregion

								#region [ Parâmetros do período de inatividade do processamento de estornos pendentes ]
								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE_HORARIO_INICIO);
								tsParametroInicioAux = Global.converteHhMmParaTimeSpan(strParametro);

								strParametro = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE_HORARIO_TERMINO);
								tsParametroTerminoAux = Global.converteHhMmParaTimeSpan(strParametro);

								intParametro = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE);
								blnFlagPeriodoInatividadeAux = (intParametro != 0) ? true : false;

								if ((tsParametroInicioAux == TimeSpan.MinValue) || (tsParametroTerminoAux == TimeSpan.MinValue)) blnFlagPeriodoInatividadeAux = false;

								if ((blnFlagPeriodoInatividadeAux != Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao)
									||
									(tsParametroInicioAux != Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio)
									||
									(tsParametroTerminoAux != Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino)
								)
								{
									strMsg = "Período de inatividade do processamento de estornos pendentes alterado de " +
										(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao ? "'habilitado'" : "'desabilitado'") +
										" (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")" +
										" para " +
										(blnFlagPeriodoInatividadeAux ? "'habilitado'" : "'desabilitado'") +
										" (" + Global.formataTimeSpanHorario(tsParametroInicioAux, "(nenhum)") + " às " + Global.formataTimeSpanHorario(tsParametroTerminoAux, "(nenhum)") + ")";
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);

									Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao = blnFlagPeriodoInatividadeAux;
									Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio = tsParametroInicioAux;
									Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino = tsParametroTerminoAux;
								}
								#endregion

								dtHrUltLeituraParametro = DateTime.Now;
							}
							#endregion

							#region [ Apaga os arquivos de log de atividade antigos? ]
							try
							{
								if (
										(DateTime.Now.TimeOfDay >= Global.Parametros.Geral.HorarioManutencaoArqLogAtividade)
										&&
										(dtHrUltManutencaoArqLogAtividade.DayOfYear != DateTime.Now.DayOfYear)
									)
								{
									Global.executaManutencaoArqLogAtividade(out strMsgErro);
									dtHrUltManutencaoArqLogAtividade = DateTime.Now;
									GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_MANUTENCAO_ARQ_LOG_ATIVIDADE, dtHrUltManutencaoArqLogAtividade);
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nManutenção de arquivos antigos de log atividade\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Apaga registros de log antigos no BD? ]
							try
							{
								if (
										(DateTime.Now.TimeOfDay >= Global.Parametros.Geral.HorarioManutencaoBdLogAntigo)
										&&
										(dtHrUltManutencaoBdLogAntigo.DayOfYear != DateTime.Now.DayOfYear)
									)
								{
									GeralDAO.executaManutencaoBdLogAntigo(out strMsgErro);
									dtHrUltManutencaoBdLogAntigo = DateTime.Now;
									GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_MANUTENCAO_BD_LOG_ANTIGO, dtHrUltManutencaoBdLogAntigo);
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nManutenção de registros antigos de log no BD\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Executa a limpeza de session token? ]
							try
							{
								#region [ Executa a limpeza de session token ]
								if (Global.Parametros.Geral.SessionToken_Limpeza_FlagHabilitacao)
								{
									if (
											(DateTime.Now.TimeOfDay >= Global.Parametros.Geral.SessionToken_Limpeza_Horario)
											&&
											(dtHrUltLimpezaSessionToken.DayOfYear != DateTime.Now.DayOfYear)
									)
									{
										dtHrInicioProcessamento = DateTime.Now;
										strMsg = "Início da execução da limpeza de session token";
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
										if (GeralDAO.executaLimpezaSessionToken(out strMsgInfo, out strMsgErro))
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Sucesso na execução da limpeza de session token (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInfo;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_LIMPEZA_SESSION_TOKEN, strMsg, out strMsgErro);
										}
										else
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Falha na execução da limpeza de session token (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_LIMPEZA_SESSION_TOKEN, strMsg, out strMsgErro);

											#region [ Envia email de alerta ]
											strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução da limpeza de session token [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = strMsg;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
												Global.gravaLogAtividade(strMsg);
											}
											#endregion
										}

										dtHrUltLimpezaSessionToken = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_LIMPEZA_SESSION_TOKEN, dtHrUltLimpezaSessionToken);
									}
								}
								#endregion
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nLimpeza de session token\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Manutenção de arquivos salvos no servidor através da WebAPI (UploadFile)? ]
							try
							{
								#region [ Executa a manutenção de arquivos salvos no servidor através da WebAPI (UploadFile) ]
								if (Global.Parametros.Geral.UploadFile_ManutencaoArquivos_FlagHabilitacao)
								{
									if (
										(DateTime.Now.TimeOfDay >= Global.Parametros.Geral.UploadFile_ManutencaoArquivos_Horario)
										&&
										(dtHrUltUploadFileManutencaoArquivos.DayOfYear != DateTime.Now.DayOfYear)
									)
									{
										dtHrInicioProcessamento = DateTime.Now;
										strMsg = "Início da manutenção dos arquivos salvos no servidor através da WebAPI (UploadFile)";
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
										if (GeralDAO.executaManutencaoArquivosUploadFile(out strMsgInfo, out strLogFalha, out strMsgErro))
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Sucesso na manutenção dos arquivos salvos no servidor através da WebAPI (UploadFile) (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInfo;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_MANUTENCAO_ARQUIVOS_UPLOAD_FILE, strMsg, out strMsgErro);

											#region [ Envia mensagem sobre falhas inesperadas? ]
											if (strLogFalha.Length > 0)
											{
												strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falhas inesperadas ocorridas durante a execução da manutenção dos arquivos salvos no servidor através da WebAPI (UploadFile) [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
												strBody = strLogFalha;
												if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
												{
													strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
													Global.gravaLogAtividade(strMsg);
												}
											}
											#endregion
										}
										else
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Falha na execução da manutenção dos arquivos salvos no servidor através da WebAPI (UploadFile) (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
											if (strLogFalha.Length > 0) strMsg += "\nFalhas inesperadas ocorridas durante o processamento:\n" + strLogFalha;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_MANUTENCAO_ARQUIVOS_UPLOAD_FILE, strMsg, out strMsgErro);

											#region [ Envia email de alerta ]
											strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução da manutenção dos arquivos salvos no servidor através da WebAPI (UploadFile) [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = strMsg;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
												Global.gravaLogAtividade(strMsg);
											}
											#endregion
										}

										dtHrUltUploadFileManutencaoArquivos = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_UPLOAD_FILE_MANUTENCAO_ARQUIVOS, dtHrUltUploadFileManutencaoArquivos);
									}
								}
								#endregion
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nManutenção de arquivos salvos no servidor através da WebAPI (UploadFile)\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Executa o cancelamento automático de pedidos? ]
							try
							{
								#region [ Executa a rotina de cancelamento automático de pedidos ]
								if (Global.Parametros.Geral.ExecutarCancelamentoAutomaticoPedidos)
								{
									if (
											(DateTime.Now.TimeOfDay >= Global.Parametros.Geral.HorarioCancelamentoAutomaticoPedidos)
											&&
											(dtHrUltCancelamentoAutomaticoPedidos.DayOfYear != DateTime.Now.DayOfYear)
										)
									{
										dtHrInicioProcessamento = DateTime.Now;
										strMsg = "Início da execução do cancelamento automático de pedidos";
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
										if (PedidoDAO.executaCancelamentoAutomaticoPedidos(out strMsgInfoCancelAutoPedidos, out strMsgErro))
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Sucesso na execução do cancelamento automático de pedidos (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInfoCancelAutoPedidos;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_CANCELAMENTO_AUTOMATICO_PEDIDO, strMsg, out strMsgErro);
										}
										else
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Falha na execução do cancelamento automático de pedidos (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_CANCELAMENTO_AUTOMATICO_PEDIDO, strMsg, out strMsgErro);

											#region [ Envia email de alerta ]
											strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução do cancelamento automático de pedidos [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = strMsg;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
												Global.gravaLogAtividade(strMsg);
											}
											#endregion
										}

										dtHrUltCancelamentoAutomaticoPedidos = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_CANCELAMENTO_AUTOMATICO_PEDIDOS, dtHrUltCancelamentoAutomaticoPedidos);
									}
								}
								#endregion
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nCancelamento automático de pedidos\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Executa o processamento dos produtos vendidos sem presença no estoque? ]
							try
							{
								#region [ Executa a rotina de processamento dos produtos vendidos sem presença no estoque ]
								if (Global.Parametros.Geral.ProcessamentoProdutosVendidosSemPresencaEstoque_FlagHabilitacao)
								{
									// Verifica se já passou o intervalo de tempo de espera desde a última verificação
									lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltConsultaExecucaoSolicitadaProcProdutosVendidosSemPresencaEstoque);
									if (lngSegundosDecorridos >= Global.Parametros.Geral.ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg)
									{
										// IMPORTANTE: essa rotina é acionada sob demanda, ou seja, não é executada em intervalos regulares
										// =========== A sinalização para que a rotina seja executada é feita através de uma flag definida em parâmetro.
										// Logo após a execução, a flag é desligada.
										parametro = GeralDAO.getRegistroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_EXECUCAO_SOLICITADA_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE);
										if (parametro != null)
										{
											// A execução foi solicitada
											if (parametro.campo_inteiro == 1)
											{
												try // Finally
												{
													// Para o parâmetro usado na solicitação da execução do processamento dos produtos vendidos sem presença no estoque,
													// os campos possuem os seguintes significados:
													//		campo_inteiro: flag que sinaliza a solicitação de execução do processamento (1 = execução solicitada)
													//		campo_texto: Relação dos códigos de id_nfe_emitente para as quais o processamento deve ser realizado (se estiver vazio, indica que deve ser realizado para todos os códigos de id_nfe_emitente ativos; se houver mais de um, separar com vírgula, ponto e vírgula ou caractere pipe, sem espaços em branco)
													//		campo_2_texto: operação/método/página que solicitou o processamento
													//		dt_hr_ult_atualizacao: data/hora em que a solicitação foi realizada
													//		usuario_ult_atualizacao: usuário que acionou a operação que acarretou na solicitação do processamento
													strMsg = "Execução do processamento dos produtos vendidos sem presença no estoque devido à solicitação requisitada para o(s) código(s) de id_nfe_emitente " +
															(parametro.campo_texto.Length == 0 ? "(todos)" : parametro.campo_texto) + " em " +
															Global.formataDataDdMmYyyyHhMmSsComSeparador(parametro.dt_hr_ult_atualizacao) +
															" por '" + parametro.usuario_ult_atualizacao + "'" +
															" (" + parametro.campo_2_texto + ")";
													Global.gravaLogAtividade(strMsg);

													dtHrInicioProcessamento = DateTime.Now;
													listaIdNfeEmitente = new List<int>();

													if (parametro.campo_texto.Trim().Length == 0)
													{
														#region [ Obtém a relação de todos os códigos de id_nfe_emitente ativos ]
														// Se o campo estiver vazio, significa que o processamento deve ser realizado para todos os códigos de
														// id_nfe_emitente que estiverem ativos.
														listaNfeEmitente = GeralDAO.getListaNfeEmitente(Global.eOpcaoFiltroStAtivo.SELECIONAR_SOMENTE_ATIVOS);
														if (listaNfeEmitente == null)
														{
															strMsg = "O processamento dos produtos vendidos sem presença no estoque não será realizado porque houve falha ao tentar obter os códigos de id_nfe_emitente ativos!";
															Global.gravaLogAtividade(strMsg);
														}
														else
														{
															strAux = "";
															foreach (NfeEmitente emitente in listaNfeEmitente)
															{
																if ((emitente.st_ativo == 1) && (emitente.st_habilitado_ctrl_estoque == 1))
																{
																	if (strAux.Length > 0) strAux += ", ";
																	strAux += emitente.id.ToString();
																	listaIdNfeEmitente.Add(emitente.id);
																}
															}
															strMsg = "O processamento dos produtos vendidos sem presença no estoque será realizado para os seguintes códigos de id_nfe_emitente ativos no sistema: " + strAux;
															Global.gravaLogAtividade(strMsg);
														}
														#endregion
													}
													else
													{
														#region [ Obtém e normaliza a lista de códigos de id_nfe_emitente p/ a qual o processamento foi solicitado ]
														strParametro = parametro.campo_texto;
														while (strParametro.Contains(" ")) strParametro = strParametro.Replace(" ", "");
														if (strParametro.Contains(";")) strParametro = strParametro.Replace(";", ",");
														if (strParametro.Contains("|")) strParametro = strParametro.Replace("|", ",");
														if (strParametro.Contains(","))
														{
															vAux = strParametro.Split(',');
															foreach (string sIdNfeEmitente in vAux)
															{
																if (sIdNfeEmitente.Trim().Length == 0) continue;
																// Verifica antes se o conteúdo informado é um texto que representa um número inteiro
																if (Global.digitos(sIdNfeEmitente).Equals(sIdNfeEmitente))
																{
																	listaIdNfeEmitente.Add((int)Global.converteInteiro(sIdNfeEmitente));
																}
																else
																{
																	strMsg = "O processamento dos produtos vendidos sem presença no estoque irá ignorar o código de id_nfe_emitente que está em formato inválido: " + sIdNfeEmitente;
																	Global.gravaLogAtividade(strMsg);
																}
															}
														}
														else
														{
															// Verifica antes se o conteúdo informado é um texto que representa um número inteiro
															if (Global.digitos(strParametro).Equals(strParametro))
															{
																listaIdNfeEmitente.Add((int)Global.converteInteiro(strParametro));
															}
															else
															{
																strMsg = "O processamento dos produtos vendidos sem presença no estoque irá ignorar o código de id_nfe_emitente que está em formato inválido: " + strParametro;
																Global.gravaLogAtividade(strMsg);
															}
														}
														#endregion
													}

													#region [ Executa o processamento ]
													if (listaIdNfeEmitente.Count == 0)
													{
														strMsg = "O processamento dos produtos vendidos sem presença no estoque não será realizado porque a solicitação não informou em formato válido a relação de códigos de id_nfe_emitente para a qual o processamento deveria ser realizado (" + (parametro.campo_texto.Length == 0 ? "(todos)" : parametro.campo_texto) + ")!";
														Global.gravaLogAtividade(strMsg);
													}
													else
													{
														if (PedidoDAO.executaProcessamentoProdutosVendidosSemPresencaEstoque(listaIdNfeEmitente, out strMsgErro))
														{
															#region [ Tratamento para sucesso no processamento ]
															lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
															strMsg = "Sucesso na execução do processamento dos produtos vendidos sem presença no estoque (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " +
																	"id_nfe_emitente = " + (parametro.campo_texto.Length == 0 ? "(todos)" : parametro.campo_texto) +
																	"; data/hora da solicitação = " + Global.formataDataDdMmYyyyHhMmSsComSeparador(parametro.dt_hr_ult_atualizacao) +
																	"; usuário = " + parametro.usuario_ult_atualizacao +
																	"; operação = " + parametro.campo_2_texto;
															Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
															GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_FINANCEIROSERVICE_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE, strMsg, out strMsgErro);
															#endregion
														}
														else
														{
															#region [ Tratamento para erro no processamento ]
															lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
															strMsg = "Falha na execução do processamento dos produtos vendidos sem presença no estoque (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " +
																	"id_nfe_emitente = " + (parametro.campo_texto.Length == 0 ? "(todos)" : parametro.campo_texto) +
																	"; data/hora da solicitação = " + Global.formataDataDdMmYyyyHhMmSsComSeparador(parametro.dt_hr_ult_atualizacao) +
																	"; usuário = " + parametro.usuario_ult_atualizacao +
																	"; operação = " + parametro.campo_2_texto +
																	"\r\n" +
																	"\r\n" +
																	strMsgErro;
															Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
															GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_FINANCEIROSERVICE_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE, strMsg, out strMsgErro);

															#region [ Envia email de alerta ]
															strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução do processamento dos produtos vendidos sem presença no estoque [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
															strBody = strMsg;
															if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
															{
																strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
																Global.gravaLogAtividade(strMsg);
															}
															#endregion
															#endregion
														}
													}
													#endregion
												}
												finally
												{
													#region [ Limpa/reseta o parâmetro usado para a solicitação de execução ]
													GeralDAO.resetRegistroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.FLAG_EXECUCAO_SOLICITADA_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE);
													#endregion
												}
											}
										}

										dtHrUltConsultaExecucaoSolicitadaProcProdutosVendidosSemPresencaEstoque = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_CONSULTA_EXECUCAO_SOLICITADA_PROC_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE, dtHrUltConsultaExecucaoSolicitadaProcProdutosVendidosSemPresencaEstoque);
									}
								}
								#endregion
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nProcessamento dos produtos vendidos sem presença no estoque\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Braspag: atualiza o status das transações de dias anteriores (cancelamento automático pela administradora por não terem sido capturadas no prazo) ]
							try
							{
								if (Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagAtualizaStatusTransacoesPendentes)
								{
									if (
											(DateTime.Now.TimeOfDay >= Global.Parametros.Braspag.FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario)
											&&
											(dtHrUltProcBpCsBraspagAtualizaStatusTransacoesPendentes.DayOfYear != DateTime.Now.DayOfYear)
										)
									{
										dtHrInicioProcessamento = DateTime.Now;
										strMsg = "Início da execução da rotina de atualização de status das transações Braspag pendentes";
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
										if (Braspag.executaProcessamentoAtualizaStatusTransacoesPendentes(out strMsgInformativa, out strMsgErro))
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Sucesso na execução da rotina de atualização de status das transações Braspag pendentes (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInformativa;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_ATUALIZACAO_STATUS_TR_PENDENTES, strMsg, out strMsgErro);
										}
										else
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Falha na execução da rotina de atualização de status das transações Braspag pendentes (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_ATUALIZACAO_STATUS_TR_PENDENTES, strMsg, out strMsgErro);

											#region [ Envia email de alerta ]
											strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução da rotina de atualização de status das transações Braspag pendentes [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = strMsg;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
												Global.gravaLogAtividade(strMsg);
											}
											#endregion
										}

										dtHrUltProcBpCsBraspagAtualizaStatusTransacoesPendentes = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_ATUALIZA_STATUS_TRANSACOES_PENDENTES, dtHrUltProcBpCsBraspagAtualizaStatusTransacoesPendentes);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nAtualização de status das transações Braspag pendentes\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Braspag: processamento dos estornos pendentes (Getnet) ]
							try
							{
								blnPeriodoAtividadeAux = true;
								if (Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao)
								{
									if (Global.isHorarioDentroIntervalo(DateTime.Now.TimeOfDay, Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio, Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino))
									{
										blnPeriodoAtividadeAux = false;
									}
								}

								if ((blnPeriodoAtividadeProcEstornosPendentes != blnPeriodoAtividadeAux) || (dtHrUltProcEstornosPendentesTransicaoPeriodoInatividade == DateTime.MinValue))
								{
									dtHrUltProcEstornosPendentesTransicaoPeriodoInatividade = DateTime.Now;
									blnPeriodoAtividadeProcEstornosPendentes = blnPeriodoAtividadeAux;
									if (blnPeriodoAtividadeProcEstornosPendentes)
									{
										strMsg = "O processamento dos estornos pendentes entrou no período de atividade (período de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									else
									{
										strMsg = "O processamento dos estornos pendentes entrou no período de inatividade (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}

								if (blnPeriodoAtividadeProcEstornosPendentes)
								{
									#region [ Executa o processamento em intervalos regulares ]
									if (Global.Parametros.Braspag.ExecutarProcessamentoBpCsEstornosPendentes)
									{
										// Executa nas seguintes condições:
										//	1) A rotina nunca foi executada.
										//	2) Já passou o intervalo de tempo de espera desde a última execução
										lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltProcEstornosPendentes);
										if (lngSegundosDecorridos >= Global.Parametros.Braspag.TempoEntreProcessamentoEstornosPendentesEmSeg)
										{
											dtHrInicioProcessamento = DateTime.Now;
											strMsg = "Início do processamento de estornos pendentes";
											Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsg);
											if (Braspag.executaProcessamentoEstornosPendentes(out qtdeEstornosPendentesVerificados, out qtdeEstornosConfirmados, out qtdeEstornosAbortados, out strMsgInfoEstornosPendentes, out strMsgErro))
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Sucesso no processamento de estornos pendentes (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInfoEstornosPendentes;
												if ((qtdeEstornosConfirmados > 0) || (qtdeEstornosAbortados > 0))
												{
													Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
													GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_ESTORNOS_PENDENTES, strMsg, out strMsgErro);
												}
												else
												{
													Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsg);
												}
											}
											else
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Falha no processamento de estornos pendentes (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
												Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
												GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_ESTORNOS_PENDENTES, strMsg, out strMsgErro);

												#region [ Envia email de alerta ]
												strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha no processamento de estornos pendentes [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
												strBody = strMsg;
												if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
												{
													strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
													Global.gravaLogAtividade(strMsg);
												}
												#endregion
											}

											dtHrUltProcEstornosPendentes = DateTime.Now;
											GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_ESTORNOS_PENDENTES, dtHrUltProcEstornosPendentes);
										}
									}
									#endregion
								}
								else
								{
									if (Global.calculaTimeSpanMinutos(DateTime.Now - dtHrUltMsgInatividadeProcEstornosPendentes) >= 30)
									{
										dtHrUltMsgInatividadeProcEstornosPendentes = DateTime.Now;
										strMsg = "Processamento dos estornos pendentes está suspenso devido ao horário de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino, "(nenhum)");
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nAtualização de status das requisições de estorno pendentes\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Consulta as transações próximas do prazo final de captura e envia email de alerta p/ supervisor ]
							try
							{
								if (Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagEnviarEmailAlertaTransacoesPendentesProxCancelAuto)
								{
									if (
											(DateTime.Now.TimeOfDay >= Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario)
											&&
											(dtHrUltProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto.DayOfYear != DateTime.Now.DayOfYear)
										)
									{
										dtHrInicioProcessamento = DateTime.Now;
										strMsg = "Início da execução da rotina de envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático";
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
										if (Braspag.executaProcessamentoEnviarEmailAlertaTransacoesPendentesProxCancelAuto(out strMsgInformativa, out strMsgErro))
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Sucesso na execução da rotina de envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInformativa;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO, strMsg, out strMsgErro);
										}
										else
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Falha na execução da rotina de envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO, strMsg, out strMsgErro);

											#region [ Envia email de alerta ]
											strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução da rotina de envio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = strMsg;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
												Global.gravaLogAtividade(strMsg);
											}
											#endregion
										}

										dtHrUltProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO, dtHrUltProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nEnvio de email de alerta sobre transações pendentes com a Braspag próximas do cancelamento automático\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Processamento das transações antifraude Clearsale (pagamentos via Braspag/Clearsale) ]
							try
							{
								blnPeriodoAtividadeAux = true;
								if (Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao)
								{
									if (Global.isHorarioDentroIntervalo(DateTime.Now.TimeOfDay, Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio, Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino))
									{
										blnPeriodoAtividadeAux = false;
									}
								}

								if ((blnPeriodoAtividadeBpCs != blnPeriodoAtividadeAux) || (dtHrUltBpCsTransicaoPeriodoInatividade == DateTime.MinValue))
								{
									dtHrUltBpCsTransicaoPeriodoInatividade = DateTime.Now;
									blnPeriodoAtividadeBpCs = blnPeriodoAtividadeAux;
									if (blnPeriodoAtividadeBpCs)
									{
										strMsg = "O processamento com a Clearsale entrou no período de atividade (período de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									else
									{
										strMsg = "O processamento com a Clearsale entrou no período de inatividade (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}

								if (blnPeriodoAtividadeBpCs)
								{
									#region [ Executa o processamento em intervalos regulares ]
									if (Global.Parametros.Clearsale.ExecutarProcessamentoBpCsAntifraudeClearsale)
									{
										// Executa nas seguintes condições:
										//	1) A rotina nunca foi executada.
										//	2) Já passou o intervalo de tempo de espera desde a última execução
										lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltProcessamentoBpCsAntifraudeClearsale);
										if (lngSegundosDecorridos >= Global.Parametros.Clearsale.TempoEntreProcessamentoEmSeg)
										{
											dtHrInicioProcessamento = DateTime.Now;
											strMsg = "Início do processamento de transações com a Clearsale";
											Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsg);

											if (Clearsale.executaProcessamentoAntifraude(out qtdeClearsalePedidosNovosEnviados, out qtdeClearsalePedidosFalhaEnvio, out qtdeClearsalePedidosResultadoProcessado, out strMsgInfoBpCsAntifraudeClearsale, out strMsgErro))
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Sucesso no processamento de transações com a Clearsale (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInfoBpCsAntifraudeClearsale;
												if ((qtdeClearsalePedidosNovosEnviados > 0) || (qtdeClearsalePedidosFalhaEnvio > 0) || (qtdeClearsalePedidosResultadoProcessado > 0))
												{
													Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
													GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_ANTIFRAUDE_CLEARSALE, strMsg, out strMsgErro);
												}
												else
												{
													Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsg);
												}
											}
											else
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Falha no processamento de transações com a Clearsale (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
												Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
												GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_ANTIFRAUDE_CLEARSALE, strMsg, out strMsgErro);

												#region [ Envia email de alerta ]
												strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha no processamento de transações com a Clearsale [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
												strBody = strMsg;
												if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
												{
													strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
													Global.gravaLogAtividade(strMsg);
												}
												#endregion
											}

											dtHrUltProcessamentoBpCsAntifraudeClearsale = DateTime.Now;
											GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_ANTIFRAUDE_CLEARSALE, dtHrUltProcessamentoBpCsAntifraudeClearsale);
										}
									}
									#endregion
								}
								else
								{
									if (Global.calculaTimeSpanMinutos(DateTime.Now - dtHrUltMsgInatividadeBpCs) >= 30)
									{
										dtHrUltMsgInatividadeBpCs = DateTime.Now;
										strMsg = "Processamento das transações antifraude Clearsale estão suspensas devido ao horário de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino, "(nenhum)");
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nProcessamento de transações antifraude Clearsale\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Consulta pedido novo aguardando tratamento da análise de crédito para envio de email de alerta ]
							try
							{
								blnPeriodoAtividadeAux = true;
								if (Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao)
								{
									if (Global.isHorarioDentroIntervalo(DateTime.Now.TimeOfDay, Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio, Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino))
									{
										blnPeriodoAtividadeAux = false;
									}
								}

								if ((blnPeriodoAtividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito != blnPeriodoAtividadeAux) || (dtHrUltProcEnvioEmailAlertaPedidoNovoAnaliseCreditoTransicaoPeriodoInatividade == DateTime.MinValue))
								{
									dtHrUltProcEnvioEmailAlertaPedidoNovoAnaliseCreditoTransicaoPeriodoInatividade = DateTime.Now;
									blnPeriodoAtividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito = blnPeriodoAtividadeAux;
									if (blnPeriodoAtividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito)
									{
										strMsg = "O processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito entrou no período de atividade (período de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									else
									{
										strMsg = "O processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito entrou no período de inatividade (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}

								if (blnPeriodoAtividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito)
								{
									#region [ Executa o processamento em intervalos regulares ]
									if (Global.Parametros.Geral.ExecutarProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito)
									{
										// Executa nas seguintes condições:
										//	1) A rotina nunca foi executada.
										//	2) Já passou o intervalo de tempo de espera desde a última execução
										lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltProcEnviarEmailAlertaPedidoNovoAnaliseCredito);
										if (lngSegundosDecorridos >= Global.Parametros.Geral.TempoEntreProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCreditoEmSeg)
										{
											dtHrInicioProcessamento = DateTime.Now;
											if (Geral.executaProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito(out blnEmailAlertaEnviado, out strMsgInformativa, out strMsgErro))
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Sucesso na execução da rotina de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInformativa;
												if (blnEmailAlertaEnviado)
												{
													Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
													GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO, strMsg, out strMsgErro);
												}
												else
												{
													Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsg);
												}
											}
											else
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Falha na execução da rotina de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
												Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
												GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO, strMsg, out strMsgErro);

												#region [ Envia email de alerta ]
												strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução da rotina de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
												strBody = strMsg;
												if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
												{
													strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
													Global.gravaLogAtividade(strMsg);
												}
												#endregion
											}

											dtHrUltProcEnviarEmailAlertaPedidoNovoAnaliseCredito = DateTime.Now;
											GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO, dtHrUltProcEnviarEmailAlertaPedidoNovoAnaliseCredito);
										}
									}
									#endregion
								}
								else
								{
									if (Global.calculaTimeSpanMinutos(DateTime.Now - dtHrUltMsgInatividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito) >= 30)
									{
										dtHrUltMsgInatividadeProcEnvioEmailAlertaPedidoNovoAnaliseCredito = DateTime.Now;
										strMsg = "Processamento de envio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito está suspenso devido ao horário de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino, "(nenhum)");
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nEnvio de email de alerta sobre pedido novo aguardando tratamento da análise de crédito\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Processa captura automática de transação pendente devido prazo final de cancelamento automático ]
							try
							{
								if (Global.Parametros.Braspag.ExecutarProcessamentoBpCsBraspagCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto)
								{
									if (
											(DateTime.Now.TimeOfDay >= Global.Parametros.Braspag.FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario)
											&&
											(dtHrUltProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto.DayOfYear != DateTime.Now.DayOfYear)
										)
									{
										dtHrInicioProcessamento = DateTime.Now;
										strMsg = "Início da execução da rotina de captura automática de transação pendente devido ao prazo final de cancelamento automático";
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
										if (Braspag.executaCapturaTransacoesPendentesPrazoFinalCancelAuto(out strMsgInformativa, out strMsgErro))
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Sucesso na execução da rotina de captura automática de transação pendente devido ao prazo final de cancelamento automático (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInformativa;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO, strMsg, out strMsgErro);
										}
										else
										{
											lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
											strMsg = "Falha na execução da rotina de captura automática de transação pendente devido ao prazo final de cancelamento automático (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
											Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
											GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO, strMsg, out strMsgErro);

											#region [ Envia email de alerta ]
											strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução da rotina de captura automática de transação pendente devido ao prazo final de cancelamento automático [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = strMsg;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
												Global.gravaLogAtividade(strMsg);
											}
											#endregion
										}

										dtHrUltProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto = DateTime.Now;
										GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO, dtHrUltProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nProcessamento da captura automática de transação pendente devido ao prazo final de cancelamento automático\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Processamento dos dados recebidos pelo Webhook Braspag ]
							try
							{
								blnPeriodoAtividadeAux = true;
								if (Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao)
								{
									if (Global.isHorarioDentroIntervalo(DateTime.Now.TimeOfDay, Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio, Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino))
									{
										blnPeriodoAtividadeAux = false;
									}
								}

								if ((blnPeriodoAtividadeProcWebhookBraspag != blnPeriodoAtividadeAux) || (dtHrUltProcWebhookBraspagTransicaoPeriodoInatividade == DateTime.MinValue))
								{
									dtHrUltProcWebhookBraspagTransicaoPeriodoInatividade = DateTime.Now;
									blnPeriodoAtividadeProcWebhookBraspag = blnPeriodoAtividadeAux;
									if (blnPeriodoAtividadeProcWebhookBraspag)
									{
										strMsg = "O processamento dos dados recebidos pelo Webhook Braspag entrou no período de atividade (período de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									else
									{
										strMsg = "O processamento dos dados recebidos pelo Webhook Braspag entrou no período de inatividade (" + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino, "(nenhum)") + ")";
									}
									Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
								}

								if (blnPeriodoAtividadeProcWebhookBraspag)
								{
									#region [ Executa o processamento em intervalos regulares ]
									if (Global.Parametros.Geral.ExecutarProcessamentoWebhookBraspag)
									{
										// Executa nas seguintes condições:
										//	1) A rotina nunca foi executada.
										//	2) Já passou o intervalo de tempo de espera desde a última execução
										lngSegundosDecorridos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltProcWebhookBraspag);
										if (lngSegundosDecorridos >= Global.Parametros.Geral.TempoEntreProcessamentoWebhookBraspagEmSeg)
										{
											dtHrInicioProcessamento = DateTime.Now;
											if (Braspag.executaProcessamentoWebhook(out blnEmailAlertaEnviado, out strMsgInformativa, out strMsgErro))
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Sucesso na execução do processamento dos dados recebidos pelo Webhook Braspag (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgInformativa;
												if (blnEmailAlertaEnviado)
												{
													Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
													GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_WEBHOOK_BRASPAG, strMsg, out strMsgErro);
												}
												else
												{
													Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + strMsg);
												}
											}
											else
											{
												lngDuracaoProcessamentoEmSegundos = Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioProcessamento);
												strMsg = "Falha na execução do processamento dos dados recebidos pelo Webhook Braspag (duração: " + lngDuracaoProcessamentoEmSegundos.ToString() + " segundos): " + strMsgErro;
												Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
												GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PROCESSAMENTO_WEBHOOK_BRASPAG, strMsg, out strMsgErro);

												#region [ Envia email de alerta ]
												strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha na execução do processamento dos dados recebidos pelo Webhook Braspag [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
												strBody = strMsg;
												if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
												{
													strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
													Global.gravaLogAtividade(strMsg);
												}
												#endregion
											}

											dtHrUltProcWebhookBraspag = DateTime.Now;
											GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROCESSAMENTO_WEBHOOK_BRASPAG, dtHrUltProcWebhookBraspag);
										}
									}
									#endregion
								}
								else
								{
									if (Global.calculaTimeSpanMinutos(DateTime.Now - dtHrUltMsgInatividadeProcWebhookBraspag) >= 30)
									{
										dtHrUltMsgInatividadeProcWebhookBraspag = DateTime.Now;
										strMsg = "Processamento dos dados recebidos pelo Webhook Braspag está suspenso devido ao horário de inatividade: " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio, "(nenhum)") + " às " + Global.formataTimeSpanHorario(Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino, "(nenhum)");
										Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Information);
									}
								}
							}
							catch (Exception ex)
							{
								strMsg = ex.ToString();
								Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\nProcessamento dos dados recebidos pelo Webhook Braspag\r\n" + strMsg, EventLogEntryType.Error);
							}
							#endregion

							#region [ Processamento de clientes em atraso ]
							//try
							//{
							//if (dtHrUltVerificacaoProcClientesEmAtraso.AddMinutes(5) < DateTime.Now)
							//{
							//    #region [ Obtém a data/hora da última carga de arquivo de retorno de boletos ]
							//    // A carga do arquivo de retorno é acionada manualmente, portanto, pode ocorrer a qualquer hora do dia (na prática costuma ser realizada durante as manhãs)
							//    dtHrUltCargaArqRetornoBoleto = GeralDAO.getDataHoraUltCargaArqRetornoBoleto();
							//    #endregion

							//    if (dtHrUltProcClientesEmAtraso < dtHrUltCargaArqRetornoBoleto)
							//    {
							//        // TODO - Processa clientes em atraso


							//        // TODO - grava data/hora último processamento

							//        dtHrUltProcClientesEmAtraso = dtHrUltCargaArqRetornoBoleto;
							//        GeralDAO.setCampoDataTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.DT_HR_ULT_PROC_CLIENTES_EM_ATRASO, dtHrUltProcClientesEmAtraso);
							//    }

							//    dtHrUltVerificacaoProcClientesEmAtraso = DateTime.Now;
							//}
							//}
							//catch (Exception ex)
							//{
							//    strMsg = ex.ToString();
							//    Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
							//}
							#endregion
						}
						catch (Exception ex)
						{
							strMsg = ex.ToString();
							Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
						}

						ProcessaSleep(1000);
					} // while (true)
				}
				finally
				{
					#region [ Grava log no BD informando que serviço foi encerrado (tabela de log geral) ]
					strMsg = "Serviço do Windows encerrado '" + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + "'";
					GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_FINANCEIROSERVICE_ENCERRADO, strMsg, out strMsgErro);
					#endregion

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = strMsg;
					GeralDAO.gravaFinSvcLog(svcLog, out strMsgErroAux);
					#endregion
				}
				#endregion
			}
			catch (Exception ex)
			{
				#region [ Envia email de alerta ]
				if (BD.isConexaoOk())
				{
					strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Falha no serviço " + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de " + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + "\nException:\n" + ex.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				strMsg = ex.ToString();
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
				isThreadManutencaoEncerrada = true;
				finalizaExecucao();
			}
			finally
			{
				#region [ Envia email de alerta ]
				if (BD.isConexaoOk())
				{
					strSubject = Global.montaIdInstanciaServicoEmailSubject() + ": Serviço " + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + " encerrado às " + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
					strBody = "Mensagem de " + Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG + ": o serviço foi encerrado às " + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out strMsgErroAux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + strMsgErroAux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + "Thread de manutenção encerrada!", EventLogEntryType.Information);
				isThreadManutencaoEncerrada = true;
			}
		}
		#endregion
	}
}
