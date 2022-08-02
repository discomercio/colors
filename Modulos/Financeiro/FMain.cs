#region[ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Xml.Serialization;
using System.Net;
#endregion

namespace Financeiro
{
	public partial class FMain : Financeiro.FModelo
	{
		#region[ Atributos ]
		public static FMain fMain;
		private bool _InicializacaoOk;
		private String REGISTRY_PATH_FORM_OPTIONS;
		ToolStripMenuItem menuCobranca;
		ToolStripMenuItem menuModuloCobranca;
		ToolStripMenuItem menuFluxoCaixa;
		ToolStripMenuItem menuFluxoCaixaConsulta;
		ToolStripMenuItem menuFluxoCaixaEditaLote;
		ToolStripMenuItem menuFluxoCaixaDebito;
		ToolStripMenuItem menuFluxoCaixaDebitoLote;
		ToolStripMenuItem menuFluxoCaixaCredito;
		ToolStripMenuItem menuFluxoCaixaCreditoLote;
		ToolStripMenuItem menuFluxoCaixaRelatorioSinteticoCtaCorrente;
		ToolStripMenuItem menuFluxoCaixaRelatorioMovimentoSintetico;
		ToolStripMenuItem menuFluxoCaixaRelatorioMovimentoAnalitico;
		ToolStripMenuItem menuFluxoCaixaRelatorioMovimentoRateioSintetico;
		ToolStripMenuItem menuFluxoCaixaRelatorioMovimentoRateioAnalitico;
		ToolStripMenuItem menuBoleto;
		ToolStripMenuItem menuBoletoCadastra;
		ToolStripMenuItem menuBoletoCadastraAvulsoComPedido;
		ToolStripMenuItem menuBoletoCadastraAvulsoSemPedido;
		ToolStripMenuItem menuBoletoGeraArquivoRemessa;
		ToolStripMenuItem menuBoletoCarregaArquivoRetorno;
		ToolStripMenuItem menuBoletoRelatoriosArquivoRetorno;
		ToolStripMenuItem menuBoletoRelatorioArquivoRemessa;
		ToolStripMenuItem menuBoletoConsulta;
		ToolStripMenuItem menuBoletoOcorrencias;
		FFluxoCredito fFluxoCredito;
		FFluxoCreditoLote fFluxoCreditoLote;
		FFluxoDebito fFluxoDebito;
		FFluxoDebitoLote fFluxoDebitoLote;
		FFluxoConsulta fFluxoConsulta;
		FFluxoEditaLoteSeleciona fFluxoEditaLoteSeleciona;
		FFluxoRelatorioCtaCorrente fFluxoRelatorioCtaCorrente;
		FFluxoRelatorioMovimentoSintetico fFluxoRelatorioMovimentoSintetico;
		FFluxoRelatorioMovimentoSinteticoComparativo fFluxoRelatorioMovimentoSinteticoComparativo;
		FFluxoRelatorioMovimentoAnalitico fFluxoRelatorioMovimentoAnalitico;
		FFluxoRelatorioMovimentoRateioSintetico fFluxoRelatorioMovimentoRateioSintetico;
		FFluxoRelatorioMovimentoRateioAnalitico fFluxoRelatorioMovimentoRateioAnalitico;
		FBoletoCadastra fBoletoCadastra;
		FBoletoArqRemessa fBoletoArqRemessa;
		FBoletoArqRetorno fBoletoArqRetorno;
		FBoletoArqRetornoRelatorios fBoletoArqRetornoRelatorios;
		FBoletoArqRemessaRelatorio fBoletoArqRemessaRelatorio;
		FBoletoConsulta fBoletoConsulta;
		FBoletoOcorrencias fBoletoOcorrencias;
		FConfiguracao fConfiguracao;
		FBoletoAvulsoComPedidoSelPedido fBoletoAvulsoComPedidoSelPedido;
		FBoletoAvulsoComPedidoCadDetalhe fBoletoAvulsoComPedidoCadDetalhe;
		FCobrancaMain fCobrancaMain;
        FPlanilhasPagtoMarketplaceSeleciona fPlanilhaPagtoMarketplaceSeleciona;
        public List<String> listaNomeClienteAutoComplete = new List<String>();
		#endregion

		#region[ Métodos Privados ]

		#region [ Banco de Dados ]

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
				BDCep.abreConexao();
				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ carregaListaNomeClienteAutoComplete ]
		private void carregaListaNomeClienteAutoComplete()
		{
			Thread thrWorker;

			thrWorker = new Thread(new ThreadStart(this.carregaListaNomeClienteAutoCompleteThread));
			thrWorker.IsBackground = true;
			thrWorker.Priority = ThreadPriority.Normal;
			thrWorker.Start();
		}
		#endregion

		#region [ carregaListaNomeClienteAutoCompleteThread ]
		private void carregaListaNomeClienteAutoCompleteThread()
		{
			#region [ Declarações ]
			DateTime dtUltCache;
			String nomeArqCacheListaNomeClienteAutoComplete;
			String strSql;
			SqlConnection cnAux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			#region [ Carrega do arquivo de cache? ]
			nomeArqCacheListaNomeClienteAutoComplete = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + '\\' + Global.Cte.FIN.NOME_ARQ_CACHE_LISTA_NOME_CLIENTE_AUTO_COMPLETE;
			dtUltCache = Global.converteYyyyMmDdParaDateTime(Global.Usuario.Defaults.dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete);
			if (File.Exists(nomeArqCacheListaNomeClienteAutoComplete) && (dtUltCache == DateTime.Today))
			{
				XmlSerializer reader;
				StreamReader file;
				reader = new XmlSerializer(listaNomeClienteAutoComplete.GetType());
				file = new StreamReader(nomeArqCacheListaNomeClienteAutoComplete);
				listaNomeClienteAutoComplete = (List<String>)reader.Deserialize(file);
				file.Close();
				return;
			}
			#endregion

			// É necessário abrir uma conexão separada, pois se for acionado algum form
			// enquanto a thread ainda estiver em execução, pode ocorrer erro, principalmente
			// se for acionada alguma consulta que utilize DataReader
			cnAux = BD.abreNovaConexao();
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand(ref cnAux);
				daAdapter = BD.criaSqlDataAdapter();
				#endregion

				#region [ Monta o SQL da consulta ]
				strSql = "SELECT DISTINCT" +
							" nome" +
						" FROM t_CLIENTE" +
						" ORDER BY" +
							" nome";
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Carrega os dados no combo ]
				listaNomeClienteAutoComplete.Clear();
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					listaNomeClienteAutoComplete.Add(dtbResultado.Rows[i]["nome"].ToString());
				}
				#endregion

				#region [ Armazena os dados em arquivo local ]
				XmlSerializer xmlWriter = new XmlSerializer(listaNomeClienteAutoComplete.GetType());
				StreamWriter file = new StreamWriter(nomeArqCacheListaNomeClienteAutoComplete);
				xmlWriter.Serialize(file, listaNomeClienteAutoComplete);
				file.Close();
				#endregion

				#region [ Memoriza data do último cache gerado ]
				Global.Usuario.Defaults.dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete = Global.formataDataYyyyMmDdComSeparador(DateTime.Today);
				#endregion
			}
			finally
			{
				BD.fechaConexao(ref cnAux);
			}
		}
		#endregion

		#endregion

		#region [ inicializaConstrutoresEstaticosUnitsDAO ]
		private static void inicializaConstrutoresEstaticosUnitsDAO()
		{
			FinLogDAO.inicializaConstrutorEstatico();
			LancamentoFluxoCaixaDAO.inicializaConstrutorEstatico();
			BoletoDAO.inicializaConstrutorEstatico();
			PedidoHistPagtoDAO.inicializaConstrutorEstatico();
			PedidoDAO.inicializaConstrutorEstatico();
			UsuarioDAO.inicializaConstrutorEstatico();
			ComboDAO.inicializaConstrutorEstatico();
			SerasaDAO.inicializaConstrutorEstatico();
		}
		#endregion

		#region [ reinicializaObjetosEstaticosUnitsDAO ]
		private static void reinicializaObjetosEstaticosUnitsDAO()
		{
			try
			{
				FinLogDAO.inicializaObjetosEstaticos();
				LancamentoFluxoCaixaDAO.inicializaObjetosEstaticos();
				BoletoDAO.inicializaObjetosEstaticos();
				PedidoHistPagtoDAO.inicializaObjetosEstaticos();
				PedidoDAO.inicializaObjetosEstaticos();
				UsuarioDAO.inicializaObjetosEstaticos();
				ComboDAO.inicializaObjetosEstaticos();
				SerasaDAO.inicializaObjetosEstaticos();
				Global.gravaLogAtividade("Sucesso ao reinicializar os objetos estáticos das units de acesso ao Banco de Dados!!");
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao reinicializar os objetos estáticos das units de acesso ao Banco de Dados!!\n" + ex.Message);
			}
		}
		#endregion

		#region [ trataBotaoConfig ]
		private void trataBotaoConfig()
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

			fConfiguracao = new FConfiguracao();
			fConfiguracao.StartPosition = FormStartPosition.Manual;
			fConfiguracao.Left = this.Left + (this.Width - fConfiguracao.Width) / 2;
			fConfiguracao.Top = this.Top + (this.Height - fConfiguracao.Height) / 2;
			fConfiguracao.ShowDialog();
		}
		#endregion

		#region [ Módulo de Cobrança ]
		private void moduloCobrancaAciona()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_COBRANCA_ACESSO_AO_MODULO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fCobrancaMain = new FCobrancaMain();
			fCobrancaMain.Location = this.Location;
			fCobrancaMain.Show();
			if (!fCobrancaMain.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Fluxo de Caixa ]

		#region [ fluxoCaixaAbrePainelLancamentoCredito ]
		private void fluxoCaixaAbrePainelLancamentoCredito()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_CREDITO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoCredito = new FFluxoCredito();
			fFluxoCredito.ShowDialog();
		}
		#endregion

		#region [ fluxoCaixaCreditoLoteProximaQtdeLancamentos ]
		private int fluxoCaixaCreditoLoteProximaQtdeLancamentos(int qtdeSelecionadaAtual)
		{
			int[] listaOpcoesQtdeLancamentos = { 25, 50, 100, 150, 200 };  // Sempre em ordem crescente!!

			for (int i = 0; i < listaOpcoesQtdeLancamentos.Length; i++)
			{
				if (qtdeSelecionadaAtual < listaOpcoesQtdeLancamentos[i]) return listaOpcoesQtdeLancamentos[i];
			}

			// Está na última opção, volta para a 1ª da lista
			return listaOpcoesQtdeLancamentos[0];
		}
		#endregion

		#region [ fluxoCaixaAbrePainelLancamentoCreditoLote ]
		private void fluxoCaixaAbrePainelLancamentoCreditoLote()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_CREDITO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fFluxoCreditoLote = new FFluxoCreditoLote(this, Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos);
				fFluxoCreditoLote.ShowDialog();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ fluxoCaixaAbrePainelLancamentoDebito ]
		private void fluxoCaixaAbrePainelLancamentoDebito()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_DEBITO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoDebito = new FFluxoDebito();
			fFluxoDebito.ShowDialog();
		}
		#endregion

		#region [ fluxoCaixaDebitoLoteProximaQtdeLancamentos ]
		private int fluxoCaixaDebitoLoteProximaQtdeLancamentos(int qtdeSelecionadaAtual)
		{
			int[] listaOpcoesQtdeLancamentos = { 25, 50, 100, 150, 200 };  // Sempre em ordem crescente!!

			for (int i = 0; i < listaOpcoesQtdeLancamentos.Length; i++)
			{
				if (qtdeSelecionadaAtual < listaOpcoesQtdeLancamentos[i]) return listaOpcoesQtdeLancamentos[i];
			}

			// Está na última opção, volta para a 1ª da lista
			return listaOpcoesQtdeLancamentos[0];
		}
		#endregion

		#region [ fluxoCaixaAbrePainelLancamentoDebitoLote ]
		private void fluxoCaixaAbrePainelLancamentoDebitoLote()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_DEBITO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fFluxoDebitoLote = new FFluxoDebitoLote(this, Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos);
				fFluxoDebitoLote.ShowDialog();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ fluxoCaixaAbrePainelConsulta ]
		private void fluxoCaixaAbrePainelConsulta()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fFluxoConsulta = new FFluxoConsulta();
				fFluxoConsulta.Location = this.Location;
				fFluxoConsulta.Show();
				if (!fFluxoConsulta.ocorreuExceptionNaInicializacao) this.Visible = false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ fluxoCaixaAbrePainelEditaLote ]
		private void fluxoCaixaAbrePainelEditaLote()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fFluxoEditaLoteSeleciona = new FFluxoEditaLoteSeleciona();
				fFluxoEditaLoteSeleciona.Location = this.Location;
				fFluxoEditaLoteSeleciona.Show();
				if (!fFluxoEditaLoteSeleciona.ocorreuExceptionNaInicializacao) this.Visible = false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ fluxoCaixaAbrePainelRelatorioSinteticoCtaCorrente ]
		private void fluxoCaixaAbrePainelRelatorioSinteticoCtaCorrente()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoRelatorioCtaCorrente = new FFluxoRelatorioCtaCorrente();
			fFluxoRelatorioCtaCorrente.Location = this.Location;
			fFluxoRelatorioCtaCorrente.Show();
			if (!fFluxoRelatorioCtaCorrente.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ fluxoCaixaAbrePainelRelatorioMovimentoSintetico ]
		private void fluxoCaixaAbrePainelRelatorioMovimentoSintetico()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoRelatorioMovimentoSintetico = new FFluxoRelatorioMovimentoSintetico();
			fFluxoRelatorioMovimentoSintetico.Location = this.Location;
			fFluxoRelatorioMovimentoSintetico.Show();
			if (!fFluxoRelatorioMovimentoSintetico.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ fluxoCaixaAbrePainelRelatorioMovimentoSinteticoComparativo ]
		private void fluxoCaixaAbrePainelRelatorioMovimentoSinteticoComparativo()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoRelatorioMovimentoSinteticoComparativo = new FFluxoRelatorioMovimentoSinteticoComparativo();
			fFluxoRelatorioMovimentoSinteticoComparativo.Location = this.Location;
			fFluxoRelatorioMovimentoSinteticoComparativo.Show();
			if (!fFluxoRelatorioMovimentoSinteticoComparativo.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ fluxoCaixaAbrePainelRelatorioMovimentoAnalitico ]
		private void fluxoCaixaAbrePainelRelatorioMovimentoAnalitico()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoRelatorioMovimentoAnalitico = new FFluxoRelatorioMovimentoAnalitico();
			fFluxoRelatorioMovimentoAnalitico.Location = this.Location;
			fFluxoRelatorioMovimentoAnalitico.Show();
			if (!fFluxoRelatorioMovimentoAnalitico.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ fluxoCaixaAbrePainelRelatorioMovimentoRateioSintetico ]
		private void fluxoCaixaAbrePainelRelatorioMovimentoRateioSintetico()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoRelatorioMovimentoRateioSintetico = new FFluxoRelatorioMovimentoRateioSintetico();
			fFluxoRelatorioMovimentoRateioSintetico.Location = this.Location;
			fFluxoRelatorioMovimentoRateioSintetico.Show();
			if (!fFluxoRelatorioMovimentoRateioSintetico.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ fluxoCaixaAbrePainelRelatorioMovimentoRateioAnalitico ]
		private void fluxoCaixaAbrePainelRelatorioMovimentoRateioAnalitico()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fFluxoRelatorioMovimentoRateioAnalitico = new FFluxoRelatorioMovimentoRateioAnalitico();
			fFluxoRelatorioMovimentoRateioAnalitico.Location = this.Location;
			fFluxoRelatorioMovimentoRateioAnalitico.Show();
			if (!fFluxoRelatorioMovimentoRateioAnalitico.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#endregion

		#region [ Boleto ]

		#region [ Boleto: Cadastramento ]
		private void boletoCadastra()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoCadastra = new FBoletoCadastra();
			fBoletoCadastra.Location = this.Location;
			fBoletoCadastra.Show();
			if (!fBoletoCadastra.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Cadastramento Avulso (com pedido) ]
		private void boletoCadastraAvulsoComPedido()
		{
			#region [ Declarações ]
			DialogResult drResultado;
			List<String> listaPedidos;
			#endregion

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

			fBoletoAvulsoComPedidoSelPedido = new FBoletoAvulsoComPedidoSelPedido();
			fBoletoAvulsoComPedidoSelPedido.StartPosition = FormStartPosition.Manual;
			fBoletoAvulsoComPedidoSelPedido.Left = this.Left + (this.Width - fBoletoAvulsoComPedidoSelPedido.Width) / 2;
			fBoletoAvulsoComPedidoSelPedido.Top = this.Top + (this.Height - fBoletoAvulsoComPedidoSelPedido.Height) / 2;
			drResultado = fBoletoAvulsoComPedidoSelPedido.ShowDialog();

			if (drResultado != DialogResult.OK) return;

			listaPedidos = fBoletoAvulsoComPedidoSelPedido.listaPedidosSelecionados;

			fBoletoAvulsoComPedidoCadDetalhe = new FBoletoAvulsoComPedidoCadDetalhe(this, listaPedidos);
			fBoletoAvulsoComPedidoCadDetalhe.Location = this.Location;
			fBoletoAvulsoComPedidoCadDetalhe.Show();
			if (!fBoletoAvulsoComPedidoCadDetalhe.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Cadastramento Avulso (sem pedido) ]
		private void boletoCadastraAvulsoSemPedido()
		{
			// TODO - boletoCadastraAvulsoSemPedido()
		}
		#endregion

		#region [ Boleto: Gera Arquivo de Remessa ]
		private void boletoGeraArquivoRemessa()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoArqRemessa = new FBoletoArqRemessa();
			fBoletoArqRemessa.Location = this.Location;
			fBoletoArqRemessa.Show();
			if (!fBoletoArqRemessa.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Carrega Arquivo de Retorno ]
		private void boletoCarregaArquivoRetorno()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoArqRetorno = new FBoletoArqRetorno();
			fBoletoArqRetorno.Location = this.Location;
			fBoletoArqRetorno.Show();
			if (!fBoletoArqRetorno.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Relatórios do Arquivo de Retorno ]
		private void boletoRelatoriosArquivoRetorno()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoArqRetornoRelatorios = new FBoletoArqRetornoRelatorios();
			fBoletoArqRetornoRelatorios.Location = this.Location;
			fBoletoArqRetornoRelatorios.Show();
			if (!fBoletoArqRetornoRelatorios.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Relatório do Arquivo de Remessa ]
		private void boletoRelatorioArquivoRemessa()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoArqRemessaRelatorio = new FBoletoArqRemessaRelatorio();
			fBoletoArqRemessaRelatorio.Location = this.Location;
			fBoletoArqRemessaRelatorio.Show();
			if (!fBoletoArqRemessaRelatorio.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Consulta ]
		private void trataBotaoBoletoConsulta()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoConsulta = new FBoletoConsulta(this);
			fBoletoConsulta.Location = this.Location;
			fBoletoConsulta.Show();
			if (!fBoletoConsulta.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#region [ Boleto: Ocorrências ]
		private void trataBotaoBoletoOcorrencias()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}
			fBoletoOcorrencias = new FBoletoOcorrencias();
			fBoletoOcorrencias.Location = this.Location;
			fBoletoOcorrencias.Show();
			if (!fBoletoOcorrencias.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
        #endregion

        #endregion

        #region [ Planilha de Pagamentos Marketplace ]
        private void trataBotaoPlanilhaPagtosMarketplace()
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

            fPlanilhaPagtoMarketplaceSeleciona = new FPlanilhasPagtoMarketplaceSeleciona();
            fPlanilhaPagtoMarketplaceSeleciona.Location = new Point(this.Location.X + (this.Size.Width - fPlanilhaPagtoMarketplaceSeleciona.Size.Width) / 2, this.Location.Y + (this.Size.Height - fPlanilhaPagtoMarketplaceSeleciona.Size.Height) / 2);
            fPlanilhaPagtoMarketplaceSeleciona.Show();
            if (!fPlanilhaPagtoMarketplaceSeleciona.ocorreuExceptionNaInicializacao) this.Visible = false;
        } 
        #endregion

        #endregion

        #region [ Métodos Públicos ]

        #region [ reiniciaBancoDados ]
        public static bool reiniciaBancoDados()
		{
			#region [ Declarações ]
			String strMsgErroLog = "";
			FinLog finLog = new FinLog();
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

			try
			{
				if (BDCep.cnConexao != null)
				{
					if (BDCep.cnConexao.State != ConnectionState.Closed) BDCep.cnConexao.Close();
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
				BDCep.cnConexao=BDCep.abreNovaConexao();
				Global.gravaLogAtividade("Sucesso ao estabelecer nova conexão!!");
				reinicializaObjetosEstaticosUnitsDAO();
				Global.gravaLogAtividade("Sucesso ao reconectar com o Banco de Dados (processo concluído)!!");
				
				#region [ Grava log no BD ]
				finLog.usuario = Global.Usuario.usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.RECONEXAO_BD;
				finLog.descricao = "Sucesso ao reconectar com o Banco de Dados";
				FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
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

		#endregion

		#region[ Construtor ]
		public FMain()
		{
			InitializeComponent();

			fMain = this;

			if (!Directory.Exists(Global.PATH_BOLETO_ARQUIVO_REMESSA)) Directory.CreateDirectory(Global.PATH_BOLETO_ARQUIVO_REMESSA);

			REGISTRY_PATH_FORM_OPTIONS = Global.RegistryApp.REGISTRY_BASE_PATH + "\\" + this.Name;

			if (!Directory.Exists(Global.Cte.LogAtividade.PathLogAtividade)) Directory.CreateDirectory(Global.Cte.LogAtividade.PathLogAtividade);
			if (!Directory.Exists(Global.Cte.Imagens.PathImagens)) Directory.CreateDirectory(Global.Cte.Imagens.PathImagens);

			#region [ Monta menus ]

			#region [ Fluxo de Caixa ]
			// Menun principal do fluxo de caixa
			menuFluxoCaixa = new ToolStripMenuItem("&Fluxo de Caixa");
			menuFluxoCaixa.Name = "menuFluxoCaixa";
			// Lançamento de débito
			menuFluxoCaixaDebito = new ToolStripMenuItem("Lançamento de &Débito", null, menuFluxoCaixaDebito_Click);
			menuFluxoCaixaDebito.Name = "menuFluxoCaixaDebito";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaDebito);
			// Lançamento de débito em lote
			menuFluxoCaixaDebitoLote = new ToolStripMenuItem("Lançamento de Dé&bito em Lote", null, menuFluxoCaixaDebitoLote_Click);
			menuFluxoCaixaDebitoLote.Name = "menuFluxoCaixaDebitoLote";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaDebitoLote);
			// Lançamento de crédito
			menuFluxoCaixaCredito = new ToolStripMenuItem("Lançamento de &Crédito", null, menuFluxoCaixaCredito_Click);
			menuFluxoCaixaCredito.Name = "menuFluxoCaixaCredito";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaCredito);
			// Lançamento de crédito em lote
			menuFluxoCaixaCreditoLote = new ToolStripMenuItem("Lançamento de C&rédito em Lote", null, menuFluxoCaixaCreditoLote_Click);
			menuFluxoCaixaCreditoLote.Name = "menuFluxoCaixaCreditoLote";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaCreditoLote);
			// Consulta de Lançamentos
			menuFluxoCaixaConsulta = new ToolStripMenuItem("Con&sulta", null, menuFluxoCaixaConsulta_Click);
			menuFluxoCaixaConsulta.Name = "menuFluxoCaixaConsulta";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaConsulta);
			// Edição de Lançamentos em Lote
			menuFluxoCaixaEditaLote = new ToolStripMenuItem("&Edição em Lote", null, menuFluxoCaixaEditaLote_Click);
			menuFluxoCaixaEditaLote.Name = "menuFluxoCaixaEditaLote";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaEditaLote);
			// Relatório de Fluxo de Caixa Sintético
			menuFluxoCaixaRelatorioSinteticoCtaCorrente = new ToolStripMenuItem("Relatório Sintético de Flu&xo de Caixa", null, menuFluxoCaixaRelatorioSinteticoCtaCorrente_Click);
			menuFluxoCaixaRelatorioSinteticoCtaCorrente.Name = "menuFluxoCaixaRelatorioSinteticoCtaCorrente";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaRelatorioSinteticoCtaCorrente);
			// Relatório de Movimentos Sintético
			menuFluxoCaixaRelatorioMovimentoSintetico = new ToolStripMenuItem("Relatório Sintético de &Movimentos", null, menuFluxoCaixaRelatorioMovimentoSintetico_Click);
			menuFluxoCaixaRelatorioMovimentoSintetico.Name = "menuFluxoCaixaRelatorioMovimentoSintetico";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaRelatorioMovimentoSintetico);
			// Relatório de Movimentos Analítico
			menuFluxoCaixaRelatorioMovimentoAnalitico = new ToolStripMenuItem("Relatório A&nalítico de Movimentos", null, menuFluxoCaixaRelatorioMovimentoAnalitico_Click);
			menuFluxoCaixaRelatorioMovimentoAnalitico.Name = "menuFluxoCaixaRelatorioMovimentoAnalitico";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaRelatorioMovimentoAnalitico);
			// Relatório de Movimentos Sintético (Rateio)
			menuFluxoCaixaRelatorioMovimentoRateioSintetico = new ToolStripMenuItem("Relatório Sintético de &Movimentos (Rateio)", null, menuFluxoCaixaRelatorioMovimentoRateioSintetico_Click);
			menuFluxoCaixaRelatorioMovimentoRateioSintetico.Name = "menuFluxoCaixaRelatorioMovimentoRateioSintetico";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaRelatorioMovimentoRateioSintetico);
			// Relatório de Movimentos Analítico (Rateio)
			menuFluxoCaixaRelatorioMovimentoRateioAnalitico = new ToolStripMenuItem("Relatório A&nalítico de Movimentos (Rateio)", null, menuFluxoCaixaRelatorioMovimentoRateioAnalitico_Click);
			menuFluxoCaixaRelatorioMovimentoRateioAnalitico.Name = "menuFluxoCaixaRelatorioMovimentoRateioAnalitico";
			menuFluxoCaixa.DropDownItems.Add(menuFluxoCaixaRelatorioMovimentoRateioAnalitico);
			// Adiciona o menu do Fluxo de Caixa ao menu principal
			menuPrincipal.Items.Insert(1, menuFluxoCaixa);
			#endregion

			#region [ Boleto ]
			// Menun principal de Boleto
			menuBoleto = new ToolStripMenuItem("&Boleto");
			menuBoleto.Name = "menuBoleto";
			// Boleto: Cadastra
			menuBoletoCadastra = new ToolStripMenuItem("Cadastramento", null, menuBoletoCadastra_Click);
			menuBoletoCadastra.Name = "menuBoletoCadastra";
			menuBoleto.DropDownItems.Add(menuBoletoCadastra);
			// Boleto: Cadastra Avulso (com pedido)
			menuBoletoCadastraAvulsoComPedido = new ToolStripMenuItem("Cadastramento Avulso (com pedido)", null, menuBoletoCadastraAvulsoComPedido_Click);
			menuBoletoCadastraAvulsoComPedido.Name = "menuBoletoCadastraAvulsoComPedido";
			menuBoleto.DropDownItems.Add(menuBoletoCadastraAvulsoComPedido);
			// Boleto: Cadastra Avulso (sem pedido)
			menuBoletoCadastraAvulsoSemPedido = new ToolStripMenuItem("Cadastramento Avulso (sem pedido)", null, menuBoletoCadastraAvulsoSemPedido_Click);
			menuBoletoCadastraAvulsoSemPedido.Name = "menuBoletoCadastraAvulsoSemPedido";
			menuBoleto.DropDownItems.Add(menuBoletoCadastraAvulsoSemPedido);
			// Boleto: Gera Arquivo Remessa
			menuBoletoGeraArquivoRemessa = new ToolStripMenuItem("Gera Arquivo de Remessa", null, menuBoletoGeraArquivoRemessa_Click);
			menuBoletoGeraArquivoRemessa.Name = "menuBoletoGeraArquivoRemessa";
			menuBoleto.DropDownItems.Add(menuBoletoGeraArquivoRemessa);
			// Boleto: Relatório Arquivo Remessa
			menuBoletoRelatorioArquivoRemessa = new ToolStripMenuItem("Relatório do Arquivo de Remessa", null, menuBoletoRelatorioArquivoRemessa_Click);
			menuBoletoRelatorioArquivoRemessa.Name = "menuBoletoRelatorioArquivoRemessa";
			menuBoleto.DropDownItems.Add(menuBoletoRelatorioArquivoRemessa);
			// Boleto: Carrega Arquivo Retorno
			menuBoletoCarregaArquivoRetorno = new ToolStripMenuItem("Carrega Arquivo de Retorno", null, menuBoletoCarregaArquivoRetorno_Click);
			menuBoletoCarregaArquivoRetorno.Name = "menuBoletoCarregaArquivoRetorno";
			menuBoleto.DropDownItems.Add(menuBoletoCarregaArquivoRetorno);
			// Boleto: Relatórios Arquivo Retorno
			menuBoletoRelatoriosArquivoRetorno = new ToolStripMenuItem("Relatórios do Arquivo de Retorno", null, menuBoletoRelatoriosArquivoRetorno_Click);
			menuBoletoRelatoriosArquivoRetorno.Name = "menuBoletoRelatoriosArquivoRetorno";
			menuBoleto.DropDownItems.Add(menuBoletoRelatoriosArquivoRetorno);
			// Boleto: Consulta
			menuBoletoConsulta = new ToolStripMenuItem("Consulta", null, menuBoletoConsulta_Click);
			menuBoletoConsulta.Name = "menuBoletoConsulta";
			menuBoleto.DropDownItems.Add(menuBoletoConsulta);
			// Boleto: Ocorrências
			menuBoletoOcorrencias = new ToolStripMenuItem("Ocorrências", null, menuBoletoOcorrencias_Click);
			menuBoletoOcorrencias.Name = "menuBoletoOcorrencias";
			menuBoleto.DropDownItems.Add(menuBoletoOcorrencias);
			// Adiciona o menu de Boleto ao menu principal
			menuPrincipal.Items.Insert(2, menuBoleto);
			#endregion

			#region [ Módulo Cobrança ]
			menuCobranca = new ToolStripMenuItem("Co&brança");
			menuCobranca.Name = "menuCobranca";
			menuModuloCobranca = new ToolStripMenuItem("&Módulo de Cobrança", null, menuModuloCobranca_Click);
			menuModuloCobranca.Name = "menuModuloCobranca";
			menuCobranca.DropDownItems.Add(menuModuloCobranca);
			// Adiciona o menu de Cobrança ao menu principal
			menuPrincipal.Items.Insert(1, menuCobranca);
			#endregion

			#endregion
		}
		#endregion

		#region[ Eventos ]

		#region[ Form fMain ]

		#region[ fMain_Shown ]
		private void fMain_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strSenhaDescriptografada = "";
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strTop;
			String strLeft;
			int intTop;
			int intLeft;
			bool blnRestauraPosicaoAnterior;
			bool blnValidacaoUsuarioOk;
			String strMsg;
			String strUltimoUsuario;
			string sVersaoPermitida;
			string[] vListaVersaoPermitida;
			List<string> listaVersaoPermitida = new List<string>();
			Color? cor;
			DateTime dtHrServidor;
			UsuarioDAO usuarioDAO;
			VersaoModulo versaoModulo;
			FLogin fLogin = new FLogin();
			DialogResult drLogin;
			FinLog finLog = new FinLog();
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
#elif (NOTEBOOK_NX6325)
					this.Text += "  (Versão de DESENVOLVIMENTO no notebook NX6325)";
					if (!confirma("Versão exclusiva para o ambiente de desenvolvimento no notebook NX6325!!\nContinua assim mesmo?"))
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
#elif (HOME_PESSOAL)
					this.Text += "  (Versão EXCLUSIVA PARA LABORATÓRIO)";
#elif (PRODUCAO)
					// NOP
#else
					this.Text += "  (Versão DESCONHECIDA)";
					avisoErro("Versão DESCONHECIDA!!\nNão é possível continuar!!");
					Close();
					return;
#endif

					#region [ Registry: dados da sessão anterior ]
					Global.Usuario.Defaults.contaCorrente = (byte)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.contaCorrente, ""));
					Global.Usuario.Defaults.planoContasEmpresa = (byte)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.planoContasEmpresa, ""));
					Global.Usuario.Defaults.planoContasContaCredito = (int)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.planoContasContaCredito, ""));
					Global.Usuario.Defaults.planoContasContaDebito = (int)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.planoContasContaDebito, ""));
					Global.Usuario.Defaults.pathBoletoArquivoRetornoRelatorios = (String)regKey.GetValue(Global.RegistryApp.Chaves.pathBoletoArquivoRetornoRelatorios, "");
					Global.Usuario.Defaults.pathBoletoArquivoRemessaRelatorio = (String)regKey.GetValue(Global.RegistryApp.Chaves.pathBoletoArquivoRemessaRelatorio, "");
					Global.Usuario.Defaults.relatorioArqRetornoTipoSaida = (String)regKey.GetValue(Global.RegistryApp.Chaves.relatorioArqRetornoTipoSaida, "");
					Global.Usuario.Defaults.relatorioArqRemessaTipoSaida = (String)regKey.GetValue(Global.RegistryApp.Chaves.relatorioArqRemessaTipoSaida, "");
					Global.Usuario.Defaults.relatorioMovimentoChkIncluirAtrasados = (String)regKey.GetValue(Global.RegistryApp.Chaves.relatorioMovimentoChkIncluirAtrasados, "");
					Global.Usuario.Defaults.dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete = (String)regKey.GetValue(Global.RegistryApp.Chaves.dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete, "");
					Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos = (int)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.fluxoCreditoLoteQtdeLancamentos, ""));
					Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos = (int)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.fluxoDebitoLoteQtdeLancamentos, ""));
					Global.Usuario.Defaults.FBoletoArqRemessa.pathBoletoArquivoRemessa = (String)regKey.GetValue(Global.RegistryApp.Chaves.FBoletoArqRemessa.pathBoletoArquivoRemessa, "");
					Global.Usuario.Defaults.FBoletoArqRemessa.boletoCedente = (byte)Global.converteInteiro((String)regKey.GetValue(Global.RegistryApp.Chaves.FBoletoArqRemessa.boletoCedente, ""));
					Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno = (String)regKey.GetValue(Global.RegistryApp.Chaves.FBoletoArqRetorno.pathBoletoArquivoRetorno, "");

					strUltimoUsuario = (String)regKey.GetValue(Global.RegistryApp.Chaves.usuario, "");

					if (Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos <= 0) Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos = fluxoCaixaCreditoLoteProximaQtdeLancamentos(Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos);
					btnFluxoCaixaCreditoLote.Text += " (" + Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos.ToString() + ")";

					if (Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos <= 0) Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos = fluxoCaixaDebitoLoteProximaQtdeLancamentos(Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos);
					btnFluxoCaixaDebitoLote.Text += " (" + Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos.ToString() + ")";
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
							Global.Usuario.fin_smtp_enable_ssl = usuarioDAO.fin_smtp_enable_ssl;
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
								if ((!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_ACESSO_AO_MODULO)) &&
									(!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_COBRANCA_ACESSO_AO_MODULO)))
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

					#region [ Configura menu e botões conforme permissões de acesso ]
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_COBRANCA_ACESSO_AO_MODULO))
					{
						btnModuloCobranca.Enabled = false;
						menuModuloCobranca.Enabled = false;
					}
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
					{
						btnFluxoCaixaConsulta.Enabled = false;
						menuFluxoCaixaConsulta.Enabled = false;
						btnFluxoCaixaEditaLote.Enabled = false;
						menuFluxoCaixaEditaLote.Enabled = false;
						btnFluxoCaixaRelatorioSinteticoCtaCorrente.Enabled = false;
						menuFluxoCaixaRelatorioSinteticoCtaCorrente.Enabled = false;
						btnFluxoCaixaRelatorioMovimentoSintetico.Enabled = false;
						menuFluxoCaixaRelatorioMovimentoSintetico.Enabled = false;
						btnFluxoCaixaRelatorioMovimentoAnalitico.Enabled = false;
						menuFluxoCaixaRelatorioMovimentoAnalitico.Enabled = false;
						btnFluxoCaixaRelatorioMovimentoRateioSintetico.Enabled = false;
						menuFluxoCaixaRelatorioMovimentoRateioSintetico.Enabled = false;
						btnFluxoCaixaRelatorioMovimentoRateioAnalitico.Enabled = false;
						menuFluxoCaixaRelatorioMovimentoRateioAnalitico.Enabled = false;
					}
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_DEBITO))
					{
						btnFluxoCaixaDebito.Enabled = false;
						menuFluxoCaixaDebito.Enabled = false;
						btnFluxoCaixaDebitoLote.Enabled = false;
						menuFluxoCaixaDebitoLote.Enabled = false;
					}
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_CREDITO))
					{
						btnFluxoCaixaCredito.Enabled = false;
						menuFluxoCaixaCredito.Enabled = false;
						btnFluxoCaixaCreditoLote.Enabled = false;
						menuFluxoCaixaCreditoLote.Enabled = false;
					}
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
					{
						btnBoletoCadastra.Enabled = false;
						menuBoletoCadastra.Enabled = false;
						btnBoletoCadastraAvulsoComPedido.Enabled = false;
						menuBoletoCadastraAvulsoComPedido.Enabled = false;
						btnBoletoCadastraAvulsoSemPedido.Enabled = false;
						menuBoletoCadastraAvulsoSemPedido.Enabled = false;
						btnBoletoGeraArquivoRemessa.Enabled = false;
						menuBoletoGeraArquivoRemessa.Enabled = false;
						btnBoletoCarregaArquivoRetorno.Enabled = false;
						menuBoletoCarregaArquivoRetorno.Enabled = false;
						btnBoletoRelatoriosArquivoRetorno.Enabled = false;
						menuBoletoRelatoriosArquivoRetorno.Enabled = false;
						btnBoletoRelatorioArquivoRemessa.Enabled = false;
						menuBoletoRelatorioArquivoRemessa.Enabled = false;
						btnBoletoConsulta.Enabled = false;
						menuBoletoConsulta.Enabled = false;
						btnBoletoOcorrencias.Enabled = false;
						menuBoletoOcorrencias.Enabled = false;
					}
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
					versaoModulo = BD.getVersaoModulo("FIN", out strMsgErro);
					if (versaoModulo == null)
					{
						strMsgErro = "Falha ao tentar obter no banco de dados o número da versão em produção deste aplicativo!!\n" + strMsgErro;
						Global.gravaLogAtividade(strMsgErro);
						avisoErro(strMsgErro);
						Close();
						return;
					}

					sVersaoPermitida = versaoModulo.versao.Trim();
					sVersaoPermitida = sVersaoPermitida.Replace(';', '|');
					vListaVersaoPermitida = sVersaoPermitida.Split('|');
					foreach (string item in vListaVersaoPermitida)
					{
						if ((item ?? "").Trim().Length > 0)
						{
							listaVersaoPermitida.Add(item.Trim());
						}
					}

					if (!listaVersaoPermitida.Contains(Global.Cte.Aplicativo.VERSAO_NUMERO))
					{
						strMsgErro = "Versão inválida do aplicativo!!\n\nVersão deste programa: " + Global.Cte.Aplicativo.VERSAO_NUMERO + "\nVersão permitida: " + String.Join(", ", listaVersaoPermitida);
						Global.gravaLogAtividade(strMsgErro);
						avisoErro(strMsgErro);
						Close();
						return;
					}
					#endregion

					#region [ Carregando dados iniciais ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados iniciais");
					if (Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR))
					{
						carregaListaNomeClienteAutoComplete();
					}
					#endregion

					#region [ Carrega parâmetros ]
					Global.Parametro.FluxoCaixa_ConsiderarDataAtualizacaoAutomatica = ComumDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FIN_FluxoCaixa_ConsiderarDataAtualizacaoAutomatica_FlagHabilitacao);
					Global.Parametro.BoletoAvulso_PermitirDivergenciaValoresFormaPagtoVsPedido = ComumDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FIN_BoletoAvulso_PermitirDivergenciaValoresFormaPagtoVsPedido_FlagHabilitacao);
					#endregion

					#region [ Copia logotipo do Bradesco usado na geração da imagem do boleto ]
					if (!File.Exists(Global.Cte.Imagens.PathImagens + "\\" + Global.Cte.Imagens.ArqLogoBradesco))
					{
						WebClient wc = new WebClient();
						wc.DownloadFile("http://central85.com.br/images/" + Global.Cte.Imagens.ArqLogoBradesco, Global.Cte.Imagens.PathImagens + "\\" + Global.Cte.Imagens.ArqLogoBradesco);
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
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.LOGON;
					finLog.descricao = "Logon realizado na máquina=" +
										System.Environment.MachineName +
										"; OS=" + System.Environment.OSVersion.VersionString +
										"; OS Version=" + System.Environment.OSVersion.Version +
										"; OS SP=" + System.Environment.OSVersion.ServicePack +
										"; Processor Count=" + System.Environment.ProcessorCount.ToString() +
										"; Windows User Name=" + System.Environment.UserName +
										"; Versão=" + Global.Cte.Aplicativo.M_ID;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
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

		#region[ fMain_FormClosing ]
		private void fMain_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Declarações ]
			FinLog finLog;
			String strMsgErroLog = "";
			#endregion

			if (_InicializacaoOk)
			{
				#region [ Memoriza no registry ]
				RegistryKey regKey = Global.RegistryApp.criaRegistryKey(REGISTRY_PATH_FORM_OPTIONS);
				regKey.SetValue(Global.RegistryApp.Chaves.contaCorrente, Global.Usuario.Defaults.contaCorrente.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.planoContasEmpresa, Global.Usuario.Defaults.planoContasEmpresa.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.planoContasContaCredito, Global.Usuario.Defaults.planoContasContaCredito.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.planoContasContaDebito, Global.Usuario.Defaults.planoContasContaDebito.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.pathBoletoArquivoRetornoRelatorios, Global.Usuario.Defaults.pathBoletoArquivoRetornoRelatorios);
				regKey.SetValue(Global.RegistryApp.Chaves.pathBoletoArquivoRemessaRelatorio, Global.Usuario.Defaults.pathBoletoArquivoRemessaRelatorio);
				regKey.SetValue(Global.RegistryApp.Chaves.relatorioArqRetornoTipoSaida, Global.Usuario.Defaults.relatorioArqRetornoTipoSaida);
				regKey.SetValue(Global.RegistryApp.Chaves.relatorioArqRemessaTipoSaida, Global.Usuario.Defaults.relatorioArqRemessaTipoSaida);
				regKey.SetValue(Global.RegistryApp.Chaves.relatorioMovimentoChkIncluirAtrasados, Global.Usuario.Defaults.relatorioMovimentoChkIncluirAtrasados);
				regKey.SetValue(Global.RegistryApp.Chaves.dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete, Global.Usuario.Defaults.dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete);
				regKey.SetValue(Global.RegistryApp.Chaves.fluxoCreditoLoteQtdeLancamentos, Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.fluxoDebitoLoteQtdeLancamentos, Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.top, this.Top.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.left, this.Left.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.usuario, Global.Usuario.usuario);
				regKey.SetValue(Global.RegistryApp.Chaves.FBoletoArqRemessa.pathBoletoArquivoRemessa, Global.Usuario.Defaults.FBoletoArqRemessa.pathBoletoArquivoRemessa);
				regKey.SetValue(Global.RegistryApp.Chaves.FBoletoArqRemessa.boletoCedente, Global.Usuario.Defaults.FBoletoArqRemessa.boletoCedente.ToString());
				regKey.SetValue(Global.RegistryApp.Chaves.FBoletoArqRetorno.pathBoletoArquivoRetorno, Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno);
				#endregion

				#region [ Log em arquivo ]
				Global.gravaLogAtividade("Término do programa");
				Global.gravaLogAtividade(null);
				Global.gravaLogAtividade(null);
				#endregion

				#region [ Log de logoff realizado gravado no BD ]
				finLog = new FinLog();
				finLog.usuario = Global.Usuario.usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.LOGOFF;
				finLog.descricao = "Logoff após " + Global.formataDuracaoHMS(DateTime.Now - Global.dtHrInicioRefRelogioLocal);
				FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
				#endregion
			}
			BD.fechaConexao();
			BDCep.fechaConexao();
		}
		#endregion

		#endregion

		#region[ Botões / Menu ]

		#region [ btnModuloCobranca_Click ]
		private void btnModuloCobranca_Click(object sender, EventArgs e)
		{
			moduloCobrancaAciona();
		}
		#endregion

		#region [ menuModuloCobranca_Click ]
		private void menuModuloCobranca_Click(object sender, EventArgs e)
		{
			moduloCobrancaAciona();
		}
		#endregion

		#region [ Fluxo de Caixa: Crédito ]

		#region [ btnFluxoCaixaCredito_Click ]
		private void btnFluxoCaixaCredito_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelLancamentoCredito();
		}
		#endregion

		#region [ menuFluxoCaixaCredito_Click ]
		private void menuFluxoCaixaCredito_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelLancamentoCredito();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Crédito em Lote ]

		#region [ btnFluxoCaixaCreditoLote_MouseUp ]
		private void btnFluxoCaixaCreditoLote_MouseUp(object sender, MouseEventArgs e)
		{
			#region [ Declarações ]
			String[] v;
			#endregion

			if (
				(e.Button == MouseButtons.Right)
				||
				(e.Button == MouseButtons.Left) && (Control.ModifierKeys == Keys.Alt)
				||
				(e.Button == MouseButtons.Left) && (Control.ModifierKeys == Keys.Control)
				||
				(e.Button == MouseButtons.Left) && (Control.ModifierKeys == Keys.Shift)
				)
			{
				Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos = fluxoCaixaCreditoLoteProximaQtdeLancamentos(Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos);
			}

			v = btnFluxoCaixaCreditoLote.Text.Split('(');
			btnFluxoCaixaCreditoLote.Text = v[0] + "(" + Global.Usuario.Defaults.fluxoCreditoLoteQtdeLancamentos.ToString() + ")";
		}
		#endregion

		#region [ btnFluxoCaixaCreditoLote_Click ]
		private void btnFluxoCaixaCreditoLote_Click(object sender, EventArgs e)
		{
			if (Control.ModifierKeys != Keys.None) return;
			
			fluxoCaixaAbrePainelLancamentoCreditoLote();
		}
		#endregion

		#region [ menuFluxoCaixaCreditoLote_Click ]
		private void menuFluxoCaixaCreditoLote_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelLancamentoCreditoLote();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Débito ]

		#region [ btnFluxoCaixaDebito_Click ]
		private void btnFluxoCaixaDebito_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelLancamentoDebito();
		}
		#endregion

		#region [ menuFluxoCaixaDebito_Click ]
		private void menuFluxoCaixaDebito_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelLancamentoDebito();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Débito em Lote ]

		#region [ btnFluxoCaixaDebitoLote_MouseUp ]
		private void btnFluxoCaixaDebitoLote_MouseUp(object sender, MouseEventArgs e)
		{
			#region [ Declarações ]
			String[] v;
			#endregion

			if (
				(e.Button == MouseButtons.Right)
				||
				(e.Button == MouseButtons.Left) && (Control.ModifierKeys == Keys.Alt)
				||
				(e.Button == MouseButtons.Left) && (Control.ModifierKeys == Keys.Control)
				||
				(e.Button == MouseButtons.Left) && (Control.ModifierKeys == Keys.Shift)
				)
			{
				Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos = fluxoCaixaDebitoLoteProximaQtdeLancamentos(Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos);
			}

			v = btnFluxoCaixaDebitoLote.Text.Split('(');
			btnFluxoCaixaDebitoLote.Text = v[0] + "(" + Global.Usuario.Defaults.fluxoDebitoLoteQtdeLancamentos.ToString() + ")";
		}
		#endregion

		#region [ btnFluxoCaixaDebitoLote_Click ]
		private void btnFluxoCaixaDebitoLote_Click(object sender, EventArgs e)
		{
			if (Control.ModifierKeys != Keys.None) return;

			fluxoCaixaAbrePainelLancamentoDebitoLote();
		}
		#endregion

		#region [ menuFluxoCaixaDebitoLote_Click ]
		private void menuFluxoCaixaDebitoLote_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelLancamentoDebitoLote();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Consulta ]

		#region [ btnFluxoCaixaConsulta_Click ]
		private void btnFluxoCaixaConsulta_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelConsulta();
		}
		#endregion

		#region [ menuFluxoCaixaConsulta_Click ]
		private void menuFluxoCaixaConsulta_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelConsulta();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Edição em Lote ]

		#region [ btnFluxoCaixaEditaLote_Click ]
		private void btnFluxoCaixaEditaLote_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelEditaLote();
		}
		#endregion

		#region [ menuFluxoCaixaEditaLote_Click ]
		private void menuFluxoCaixaEditaLote_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelEditaLote();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Relatório de Fluxo de Caixa Sintético ]

		#region [ btnFluxoCaixaRelatorioSinteticoCtaCorrente_Click ]
		private void btnFluxoCaixaRelatorioSinteticoCtaCorrente_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioSinteticoCtaCorrente();
		}
		#endregion

		#region [ menuFluxoCaixaRelatorioSinteticoCtaCorrente_Click ]
		private void menuFluxoCaixaRelatorioSinteticoCtaCorrente_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioSinteticoCtaCorrente();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Relatório Sintético de Movimentos ]

		#region [ btnFluxoCaixaRelatorioMovimentoSintetico_Click ]
		private void btnFluxoCaixaRelatorioMovimentoSintetico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoSintetico();
		}
		#endregion

		#region [ menuFluxoCaixaRelatorioMovimentoSintetico_Click ]
		private void menuFluxoCaixaRelatorioMovimentoSintetico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoSintetico();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Relatório Comparativo Sintético de Movimentos (Excel) ]
		private void btnFluxoCaixaRelatorioMovimentoSinteticoComparativo_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoSinteticoComparativo();
		}
		#endregion

		#region [ Fluxo de Caixa: Relatório Analítico de Movimentos ]

		#region [ btnFluxoCaixaRelatorioMovimentoAnalitico_Click ]
		private void btnFluxoCaixaRelatorioMovimentoAnalitico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoAnalitico();
		}
		#endregion

		#region [ menuFluxoCaixaRelatorioMovimentoAnalitico_Click ]
		private void menuFluxoCaixaRelatorioMovimentoAnalitico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoAnalitico();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Relatório Sintético de Movimentos (Rateio) ]

		#region [ btnFluxoCaixaRelatorioMovimentoRateioSintetico_Click ]
		private void btnFluxoCaixaRelatorioMovimentoRateioSintetico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoRateioSintetico();
		}
		#endregion

		#region [ menuFluxoCaixaRelatorioMovimentoRateioSintetico_Click ]
		private void menuFluxoCaixaRelatorioMovimentoRateioSintetico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoRateioSintetico();
		}
		#endregion

		#endregion

		#region [ Fluxo de Caixa: Relatório Analítico de Movimentos (Rateio) ]

		#region [ btnFluxoCaixaRelatorioMovimentoRateioAnalitico_Click ]
		private void btnFluxoCaixaRelatorioMovimentoRateioAnalitico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoRateioAnalitico();
		}
		#endregion

		#region [ menuFluxoCaixaRelatorioMovimentoRateioAnalitico_Click ]
		private void menuFluxoCaixaRelatorioMovimentoRateioAnalitico_Click(object sender, EventArgs e)
		{
			fluxoCaixaAbrePainelRelatorioMovimentoRateioAnalitico();
		}
		#endregion

		#endregion

		#region [ Boleto: Cadastramento ]

		#region [ btnBoletoCadastra_Click ]
		private void btnBoletoCadastra_Click(object sender, EventArgs e)
		{
			boletoCadastra();
		}
		#endregion

		#region [ menuBoletoCadastra_Click ]
		private void menuBoletoCadastra_Click(object sender, EventArgs e)
		{
			boletoCadastra();
		}
		#endregion

		#endregion

		#region [ Boleto: Cadastramento Avulso (com pedido) ]

		#region [ btnBoletoCadastraAvulsoComPedido_Click ]
		private void btnBoletoCadastraAvulsoComPedido_Click(object sender, EventArgs e)
		{
			boletoCadastraAvulsoComPedido();
		}
		#endregion

		#region [ menuBoletoCadastraAvulsoComPedido_Click ]
		private void menuBoletoCadastraAvulsoComPedido_Click(object sender, EventArgs e)
		{
			boletoCadastraAvulsoComPedido();
		}
		#endregion

		#endregion

		#region [ Boleto: Cadastramento Avulso (sem pedido) ]

		#region [ btnBoletoCadastraAvulsoSemPedido_Click ]
		private void btnBoletoCadastraAvulsoSemPedido_Click(object sender, EventArgs e)
		{
			boletoCadastraAvulsoSemPedido();
		}
		#endregion

		#region [ menuBoletoCadastraAvulsoSemPedido_Click ]
		private void menuBoletoCadastraAvulsoSemPedido_Click(object sender, EventArgs e)
		{
			boletoCadastraAvulsoSemPedido();
		}
		#endregion

		#endregion

		#region [ Boleto: Gera Arquivo de Remessa ]

		#region [ btnBoletoGeraArquivoRemessa_Click ]
		private void btnBoletoGeraArquivoRemessa_Click(object sender, EventArgs e)
		{
			boletoGeraArquivoRemessa();
		}
		#endregion

		#region [ menuBoletoGeraArquivoRemessa_Click ]
		private void menuBoletoGeraArquivoRemessa_Click(object sender, EventArgs e)
		{
			boletoGeraArquivoRemessa();
		}
		#endregion

		#endregion

		#region [ btnBoletoRelatorioArquivoRemessa ]

		#region [ btnBoletoRelatorioArquivoRemessa_Click ]
		private void btnBoletoRelatorioArquivoRemessa_Click(object sender, EventArgs e)
		{
			boletoRelatorioArquivoRemessa();
		}
		#endregion

		#region [ menuBoletoRelatorioArquivoRemessa_Click ]
		private void menuBoletoRelatorioArquivoRemessa_Click(object sender, EventArgs e)
		{
			boletoRelatorioArquivoRemessa();
		}
		#endregion

		#endregion

		#region [ Boleto: Carrega Arquivo de Retorno ]

		#region [ btnBoletoCarregaArquivoRetorno_Click ]
		private void btnBoletoCarregaArquivoRetorno_Click(object sender, EventArgs e)
		{
			boletoCarregaArquivoRetorno();
		}
		#endregion

		#region [ menuBoletoCarregaArquivoRetorno_Click ]
		private void menuBoletoCarregaArquivoRetorno_Click(object sender, EventArgs e)
		{
			boletoCarregaArquivoRetorno();
		}
		#endregion

		#endregion

		#region [ Boleto: Relatórios do Arquivo de Retorno ]

		#region [ btnBoletoRelatoriosArquivoRetorno_Click ]
		private void btnBoletoRelatoriosArquivoRetorno_Click(object sender, EventArgs e)
		{
			boletoRelatoriosArquivoRetorno();
		}
		#endregion

		#region [ menuBoletoRelatoriosArquivoRetorno_Click ]
		private void menuBoletoRelatoriosArquivoRetorno_Click(object sender, EventArgs e)
		{
			boletoRelatoriosArquivoRetorno();
		}
		#endregion

		#endregion

		#region [ Boleto: Consulta ]

		#region [ btnBoletoConsulta_Click ]
		private void btnBoletoConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoBoletoConsulta();
		}
		#endregion

		#region [ menuBoletoConsulta_Click ]
		private void menuBoletoConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoBoletoConsulta();
		}
		#endregion

		#endregion

		#region [ Boleto: Ocorrências ]

		#region [ btnBoletoOcorrencias_Click ]
		private void btnBoletoOcorrencias_Click(object sender, EventArgs e)
		{
			trataBotaoBoletoOcorrencias();
		}
		#endregion

		#region [ menuBoletoOcorrencias_Click ]
		private void menuBoletoOcorrencias_Click(object sender, EventArgs e)
		{
			trataBotaoBoletoOcorrencias();
		}
        #endregion

        #endregion

        #region [ btnPlanilhaPagtosMktplace_Click ]
        private void btnPlanilhaPagtosMarketplace_Click(object sender, EventArgs e)
        {
            trataBotaoPlanilhaPagtosMarketplace();
        }
        #endregion

        #region [ Painel de Configuração ]

        #region [ btnConfig_Click ]
        private void btnConfig_Click(object sender, EventArgs e)
		{
			trataBotaoConfig();
		}
		#endregion

		#endregion

		#endregion

		#endregion
	}
}
