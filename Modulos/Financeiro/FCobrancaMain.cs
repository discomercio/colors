#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Financeiro
{
	public partial class FCobrancaMain : Financeiro.FModelo
	{
		#region [ Atributos ]

		#region [ Diversos ]
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

		FCobrancaAdministracao fCobrancaAdministracao;
		FBoletoConsulta fBoletoConsulta;
		FCobrancaFluxoConsulta fFluxoCaixaConsulta;
		#endregion

		#region [ Menu ]
		ToolStripMenuItem menuCobranca;
		ToolStripMenuItem menuAdministracaoCarteiraEmAtraso;
		ToolStripMenuItem menuBoletoConsulta;
		ToolStripMenuItem menuFluxoCaixaConsulta;
		#endregion

		#endregion

		#region [ Construtor ]
		public FCobrancaMain()
		{
			InitializeComponent();

			#region [ Menu ]
			// Menu principal de Cobrança
			menuCobranca = new ToolStripMenuItem("&Cobrança");
			menuCobranca.Name = "menuCobranca";
			// Administração da Carteira em Atraso
			menuAdministracaoCarteiraEmAtraso = new ToolStripMenuItem("&Administração da Carteira em Atraso", null, menuAdministracaoCarteiraEmAtraso_Click);
			menuAdministracaoCarteiraEmAtraso.Name = "menuAdministracaoCarteiraEmAtraso";
			menuCobranca.DropDownItems.Add(menuAdministracaoCarteiraEmAtraso);
			// Consulta de Boleto
			menuBoletoConsulta = new ToolStripMenuItem("&Consulta de Boletos", null, menuBoletoConsulta_Click);
			menuBoletoConsulta.Name = "menuBoletoConsulta";
			menuCobranca.DropDownItems.Add(menuBoletoConsulta);
			// Consulta do Fluxo de Caixa
			menuFluxoCaixaConsulta = new ToolStripMenuItem("Consulta do Flu&xo de Caixa", null, menuFluxoCaixaConsulta_Click);
			menuFluxoCaixaConsulta.Name = "menuFluxoCaixaConsulta";
			menuCobranca.DropDownItems.Add(menuFluxoCaixaConsulta);
			// Adiciona o menu Cobrança ao menu principal
			menuPrincipal.Items.Insert(1, menuCobranca);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ trataBotaoAdministracaoCarteiraEmAtraso ]
		private void trataBotaoAdministracaoCarteiraEmAtraso()
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

			if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_COBRANCA_ADMINISTRACAO_CARTEIRA_EM_ATRASO))
			{
				avisoErro("Nível de acesso insuficiente!!");
				return;
			}

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fCobrancaAdministracao = new FCobrancaAdministracao(this);
				fCobrancaAdministracao.Location = this.Location;
				fCobrancaAdministracao.Show();
				if (!fCobrancaAdministracao.ocorreuExceptionNaInicializacao) this.Visible = false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
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

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fBoletoConsulta = new FBoletoConsulta(this);
				fBoletoConsulta.Location = this.Location;
				fBoletoConsulta.Show();
				if (!fBoletoConsulta.ocorreuExceptionNaInicializacao) this.Visible = false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ Fluxo de Caixa: Consulta ]
		private void trataBotaoFluxoCaixaConsulta()
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

			info(ModoExibicaoMensagemRodape.EmExecucao, "carregando painel");
			try
			{
				fFluxoCaixaConsulta = new FCobrancaFluxoConsulta(this);
				fFluxoCaixaConsulta.Location = this.Location;
				fFluxoCaixaConsulta.Show();
				if (!fFluxoCaixaConsulta.ocorreuExceptionNaInicializacao) this.Visible = false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form: FCobrancaMain ]

		#region [ FCobrancaMain_Load ]
		private void FCobrancaMain_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
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

		#region [ FCobrancaMain_Shown ]
		private void FCobrancaMain_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Permissão de acesso ao módulo ]
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_COBRANCA_ADMINISTRACAO_CARTEIRA_EM_ATRASO))
					{
						btnAdministracaoCarteiraEmAtraso.Enabled = false;
						menuAdministracaoCarteiraEmAtraso.Enabled = false;
					}
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

		#region [ FCobrancaMain_FormClosing ]
		private void FCobrancaMain_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ Botões / Menu ]

		#region [ Administração da Carteira em Atraso ]

		#region [ btnAdministracaoCarteiraEmAtraso_Click ]
		private void btnAdministracaoCarteiraEmAtraso_Click(object sender, EventArgs e)
		{
			trataBotaoAdministracaoCarteiraEmAtraso();
		}
		#endregion

		#region [ menuAdministracaoCarteiraEmAtraso_Click ]
		private void menuAdministracaoCarteiraEmAtraso_Click(object sender, EventArgs e)
		{
			trataBotaoAdministracaoCarteiraEmAtraso();
		}
		#endregion

		#endregion

		#region [ Consulta de Boleto ]

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

		#region [ Consulta do Fluxo de Caixa ]

		#region [ btnFluxoCaixaConsulta_Click ]
		private void btnFluxoCaixaConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoFluxoCaixaConsulta();
		}
		#endregion

		#region [ menuFluxoCaixaConsulta_Click ]
		private void menuFluxoCaixaConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoFluxoCaixaConsulta();
		}
		#endregion

		#endregion

		#endregion

		#endregion
	}
}
