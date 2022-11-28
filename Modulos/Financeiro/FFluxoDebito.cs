#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Media;
#endregion

namespace Financeiro
{
	public partial class FFluxoDebito : Financeiro.FModelo
	{
		#region [ Atributos ]
		private bool _InicializacaoOk;
		ToolStripMenuItem menuLancamento;
		ToolStripMenuItem menuLancamentoGravar;
		ToolStripMenuItem menuLancamentoLimpar;
		#endregion

		#region [ Construtor ]
		public FFluxoDebito()
		{
			InitializeComponent();

			#region [ Menu Lançamento ]
			// Menu principal de Lançamento
			menuLancamento = new ToolStripMenuItem("&Lançamento");
			menuLancamento.Name = "menuLancamento";
			// Gravar
			menuLancamentoGravar = new ToolStripMenuItem("&Gravar", null, menuLancamentoGravar_Click);
			menuLancamentoGravar.Name = "menuLancamentoGravar";
			menuLancamento.DropDownItems.Add(menuLancamentoGravar);
			// Limpar
			menuLancamentoLimpar = new ToolStripMenuItem("Lim&par", null, menuLancamentoLimpar_Click);
			menuLancamentoLimpar.Name = "menuLancamentoLimpar";
			menuLancamento.DropDownItems.Add(menuLancamentoLimpar);
			// Adiciona o menu Lançamento ao menu principal
			menuPrincipal.Items.Insert(1, menuLancamento);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		void limpaCampos()
		{
			cbContaCorrente.SelectedIndex = -1;
			cbPlanoContasEmpresa.SelectedIndex = -1;
			cbPlanoContasConta.SelectedIndex = -1;
			txtDataCompetencia.Text = "";
			txtValor.Text = "";
			txtCnpjCpf.Text = "";
			txtNF.Text = "";
			txtDescricao.Text = "";
            txtComp2.Text = "";
		}
		#endregion

		#region [ comboContaCorrentePosicionaDefault ]
		private bool comboContaCorrentePosicionaDefault()
		{
			bool blnHaDefault = false;
			DsDataSource.DtbContaCorrenteComboRow rowContaCorrente;

			foreach (System.Data.DataRowView item in cbContaCorrente.Items)
			{
				rowContaCorrente = (DsDataSource.DtbContaCorrenteComboRow)item.Row;
				if (rowContaCorrente.id == Global.Usuario.Defaults.contaCorrente)
				{
					cbContaCorrente.SelectedIndex = cbContaCorrente.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
		}
		#endregion

		#region [ comboPlanoContasEmpresaPosicionaDefault ]
		private bool comboPlanoContasEmpresaPosicionaDefault()
		{
			bool blnHaDefault = false;
			DsDataSource.DtbPlanoContasEmpresaComboRow rowPlanoContasEmpresa;

			foreach (System.Data.DataRowView item in cbPlanoContasEmpresa.Items)
			{
				rowPlanoContasEmpresa = (DsDataSource.DtbPlanoContasEmpresaComboRow)item.Row;
				if (rowPlanoContasEmpresa.id == Global.Usuario.Defaults.planoContasEmpresa)
				{
					cbPlanoContasEmpresa.SelectedIndex = cbPlanoContasEmpresa.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
		}
		#endregion

		#region [ comboPlanoContasContaPosicionaDefault ]
		private bool comboPlanoContasContaPosicionaDefault()
		{
			bool blnHaDefault = false;
			DsDataSource.DtbPlanoContasContaComboRow rowPlanoContasConta;

			foreach (System.Data.DataRowView item in cbPlanoContasConta.Items)
			{
				rowPlanoContasConta = (DsDataSource.DtbPlanoContasContaComboRow)item.Row;
				if (rowPlanoContasConta.id == Global.Usuario.Defaults.planoContasContaDebito)
				{
					cbPlanoContasConta.SelectedIndex = cbPlanoContasConta.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
		}
		#endregion

		#region [ posicionaFocoPrimeiroCampoPreencher ]
		private void posicionaFocoPrimeiroCampoPreencher()
		{
			if (cbContaCorrente.SelectedIndex == -1)
			{
				cbContaCorrente.Focus();
				return;
			}
			if (cbPlanoContasEmpresa.SelectedIndex == -1)
			{
				cbPlanoContasEmpresa.Focus();
				return;
			}
			if (cbPlanoContasConta.SelectedIndex == -1)
			{
				cbPlanoContasConta.Focus();
				return;
			}
			if (txtDataCompetencia.Text.Trim().Length == 0)
			{
				txtDataCompetencia.Focus();
				return;
			}
			if (txtValor.Text.Trim().Length == 0)
			{
				txtValor.Focus();
				return;
			}
		}
		#endregion

		#region [ obtemDadosLancamentoCamposTela ]
		/// <summary>
		/// Carrega os dados dos campos na tela em um objeto da classe LancamentoFluxoCaixa
		/// </summary>
		/// <returns>
		/// Retorna um objeto LancamentoFluxoCaixa com os dados dos campos da tela
		/// </returns>
		private LancamentoFluxoCaixa obtemDadosLancamentoCamposTela()
		{
			LancamentoFluxoCaixa lancamento = new LancamentoFluxoCaixa();

			// O grupo de contas é obtido a partir da conta, ou seja, não é selecionado explicitamente pelo usuário
			// Lembrando que cada conta foi vinculada a um grupo de contas no momento do cadastramento
			System.Data.DataRowView dataRowView = (System.Data.DataRowView)cbPlanoContasConta.Items[cbPlanoContasConta.SelectedIndex];
			DsDataSource.DtbPlanoContasContaComboRow rowConta = (DsDataSource.DtbPlanoContasContaComboRow)dataRowView.Row;
			lancamento.id_plano_contas_grupo = (int)Global.converteInteiro(rowConta.id_plano_contas_grupo.ToString());
			lancamento.id_conta_corrente = (byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString());
			lancamento.id_plano_contas_empresa = (byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString());
			lancamento.id_plano_contas_conta = (int)Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString());
			lancamento.dt_competencia = Global.converteDdMmYyyyParaDateTime(txtDataCompetencia.Text);
            lancamento.dt_mes_competencia = Convert.ToDateTime(txtComp2.Text);
			lancamento.valor = Global.converteNumeroDecimal(txtValor.Text);
			lancamento.cnpj_cpf = Global.digitos(txtCnpjCpf.Text.Trim());
			lancamento.numero_NF = (int)Global.converteInteiro(Global.digitos(txtNF.Text.Trim()));
			lancamento.descricao = txtDescricao.Text.Trim();

			return lancamento;
		}
		#endregion

		#region [ consisteCampos ]
		/// <summary>
		/// Realiza a consistência dos campos na tela
		/// </summary>
		/// <returns>
		/// true: os campos estão devidamente preenchidos
		/// false: há campos não preenchidos corretamente
		/// </returns>
		private bool consisteCampos()
		{
			if (cbContaCorrente.SelectedIndex == -1)
			{
				avisoErro("Selecione uma conta corrente!!");
				cbContaCorrente.Focus();
				return false;
			}
			if (cbPlanoContasEmpresa.SelectedIndex == -1)
			{
				avisoErro("Selecione uma empresa!!");
				cbPlanoContasEmpresa.Focus();
				return false;
			}
			if (cbPlanoContasConta.SelectedIndex == -1)
			{
				avisoErro("Selecione um plano de conta!!");
				cbPlanoContasConta.Focus();
				return false;
			}
			if (txtDataCompetencia.Text.Trim().Length == 0)
			{
				avisoErro("Informe a data da competência!!");
				txtDataCompetencia.Focus();
				return false;
			}
            if (txtComp2.Text.Trim().Length == 0)
            {
                avisoErro("Informe o mês de competência!!");
                txtComp2.Focus();
                return false;
            }
            if (!Global.isDataOk(txtDataCompetencia.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCompetencia.Focus();
				return false;
			}
			if (txtValor.Text.Trim().Length == 0)
			{
				avisoErro("Informe o valor!!");
				txtValor.Focus();
				return false;
			}
			if (Global.converteNumeroDecimal(txtValor.Text) <= 0)
			{
				avisoErro("Valor inválido!!");
				txtValor.Focus();
				return false;
			}
			if (txtDescricao.Text.Trim().Length == 0)
			{
				avisoErro("Informe a descrição!!");
				txtDescricao.Focus();
				return false;
			}

			return true;
		}
		#endregion

		#region [ trataBotaoGravar ]
		void trataBotaoGravar()
		{
			#region [ Declarações ]
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			bool blnResultado;
			LancamentoFluxoCaixa lancamento;
			FinLog finLog = new FinLog();
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

			#region [ Consistência ]
			if (!consisteCampos()) return;
			#endregion

			#region [ Obtém valores ]
			lancamento = obtemDadosLancamentoCamposTela();
			lancamento.natureza = Global.Cte.FIN.Natureza.DEBITO;
			lancamento.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
			#endregion

			#region [ Grava no banco de dados ]
			blnResultado = LancamentoFluxoCaixaDAO.insere(	Global.Usuario.usuario,
															lancamento,
															ref strDescricaoLog,
															ref strMsgErro
															);
			#endregion

			#region [ Processamento pós tentativa de gravação no BD ]
			if (blnResultado)
			{
				#region [ Incrementa contador de lançamentos gravados ]
				Global.contadorLancamentoDebitoInserido++;
				lblContador.Text = Global.contadorLancamentoDebitoInserido.ToString().PadLeft(2, '0');
				#endregion

				#region [ Atualiza defaults do usuário ]
				Global.Usuario.Defaults.contaCorrente = (byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString());
				Global.Usuario.Defaults.planoContasEmpresa = (byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString());
				Global.Usuario.Defaults.planoContasContaDebito = (int)Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString());
				#endregion

				#region [ Grava log no BD ]
				finLog.usuario = Global.Usuario.usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_DEBITO_INSERE;
				finLog.natureza = lancamento.natureza;
				finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
				finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
				finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
				finLog.id_registro_origem = lancamento.id;
				finLog.id_conta_corrente = lancamento.id_conta_corrente;
				finLog.id_plano_contas_empresa = lancamento.id_plano_contas_empresa;
				finLog.id_plano_contas_grupo = lancamento.id_plano_contas_grupo;
				finLog.id_plano_contas_conta = lancamento.id_plano_contas_conta;
				finLog.id_cliente = lancamento.id_cliente;
				finLog.cnpj_cpf = lancamento.cnpj_cpf;
				finLog.descricao = strDescricaoLog;
				FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
				#endregion

				#region [ Prepara para cadastrar próximo lançamento ]
				limpaCampos();
				if (!comboContaCorrentePosicionaDefault()) cbContaCorrente.SelectedIndex = -1;
				if (!comboPlanoContasEmpresaPosicionaDefault()) cbPlanoContasEmpresa.SelectedIndex = -1;
				if (!comboPlanoContasContaPosicionaDefault()) cbPlanoContasConta.SelectedIndex = -1;
				posicionaFocoPrimeiroCampoPreencher();
				#endregion

				SystemSounds.Asterisk.Play();
			}
			else
			{
				avisoErro("Falha ao gravar o registro!!\n\n" + strMsgErro);
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoLimpar ]
		void trataBotaoLimpar()
		{
			limpaCampos();
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FFluxoDebito ]

		#region [ FFluxoDebito_Load ]
		private void FFluxoDebito_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				#region [ Limpa campos ]
				limpaCampos();
				lblContador.Text = Global.contadorLancamentoDebitoInserido.ToString().PadLeft(2, '0');
				#endregion

				#region [ Combo Conta Corrente ]
				cbContaCorrente.DataSource = ComboDAO.criaDtbContaCorrenteCombo(ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS);
				cbContaCorrente.ValueMember = "id";
				cbContaCorrente.DisplayMember = "contaComDescricao";
				if (Global.Usuario.Defaults.contaCorrente == 0)
					cbContaCorrente.SelectedIndex = -1;
				else
					if (!comboContaCorrentePosicionaDefault()) cbContaCorrente.SelectedIndex = -1;
				#endregion

				#region [ Combo Plano Contas Empresa ]
				cbPlanoContasEmpresa.DataSource = ComboDAO.criaDtbPlanoContasEmpresaCombo(ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS);
				cbPlanoContasEmpresa.ValueMember = "id";
				cbPlanoContasEmpresa.DisplayMember = "idComDescricao";
				if (Global.Usuario.Defaults.planoContasEmpresa == 0)
					cbPlanoContasEmpresa.SelectedIndex = -1;
				else
					if (!comboPlanoContasEmpresaPosicionaDefault()) cbPlanoContasEmpresa.SelectedIndex = -1;
				#endregion

				#region [ Combo Plano Contas Conta ]
				cbPlanoContasConta.DataSource = ComboDAO.criaDtbPlanoContasContaCombo(ComboDAO.eFiltraNatureza.SOMENTE_DEBITO, ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS, ComboDAO.eFiltraStSistema.SOMENTE_CONTAS_NORMAIS);
				cbPlanoContasConta.ValueMember = "id";
				cbPlanoContasConta.DisplayMember = "idComDescricao";
				if (Global.Usuario.Defaults.planoContasContaDebito == 0)
					cbPlanoContasConta.SelectedIndex = -1;
				else
					if (!comboPlanoContasContaPosicionaDefault()) cbPlanoContasConta.SelectedIndex = -1;
				#endregion

				#region [ Campo descrição ]
				txtDescricao.MaxLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
				#endregion

				blnSucesso = true;
			}
			catch (Exception ex)
			{
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

		#region [ FFluxoDebito_Shown ]
		private void FFluxoDebito_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					posicionaFocoPrimeiroCampoPreencher();
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
				// Se não inicializou corretamente, assegura-se de que o painel será fechado
				if (!_InicializacaoOk) Close();
			}
		}
		#endregion

		#endregion

		#region [ txtDataCompetencia ]

		#region [ txtDataCompetencia_Enter ]
		private void txtDataCompetencia_Enter(object sender, EventArgs e)
		{
			txtDataCompetencia.Select(0, txtDataCompetencia.Text.Length);
		}
		#endregion

		#region [ txtDataCompetencia_Leave ]
		private void txtDataCompetencia_Leave(object sender, EventArgs e)
		{
            string strDataAux;
			if (txtDataCompetencia.Text.Length == 0) return;
			txtDataCompetencia.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCompetencia.Text);
            if (!Global.isDataOk(txtDataCompetencia.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCompetencia.Focus();
				return;
			}
            strDataAux = Convert.ToDateTime(txtDataCompetencia.Text).ToString("MM/yyyy");
            if (txtComp2.Text.Trim().Length == 0) txtComp2.Text = strDataAux;
        }
		#endregion

		#region [ txtDataCompetencia_KeyPress ]
		private void txtDataCompetencia_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#region [ txtDataCompetencia_KeyDown ]
		private void txtDataCompetencia_KeyDown(object sender, KeyEventArgs e)
		{
			if ((e.Shift || e.Control) && ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Space)))
			{
				if (txtDataCompetencia.Text.Trim().Length == 0) txtDataCompetencia.Text = Global.formataDataDdMmYyyyComSeparador(DateTime.Now);
			}

			Global.trataTextBoxKeyDown(sender, e, txtComp2);
		}
        #endregion

        #endregion

        #region [ txtComp2 ]

        #region [ txtComp2_KeyDown ]

        private void txtComp2_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, txtValor);
        }

        #endregion

        #region [ txtComp2_KeyPress ]
        private void txtComp2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #region [ txtComp2_Leave ]

        private void txtComp2_Leave(object sender, EventArgs e)
        {
            if (txtComp2.Text.Length == 0) return;
            txtComp2.Text = Global.formataDataDigitadaParaMMYYYYComSeparador(txtComp2.Text);
            if (!Global.isDataMMYYYYOk(txtComp2.Text))
            {
                avisoErro("Formato inválido!!");
                txtComp2.Focus();
                return;
            }
        }

        #endregion

        #endregion

        #region [ txtValor ]

        #region [ txtValor_Enter ]
        private void txtValor_Enter(object sender, EventArgs e)
		{
			txtValor.Select(0, txtValor.Text.Length);
		}
		#endregion

		#region [ txtValor_Leave ]
		private void txtValor_Leave(object sender, EventArgs e)
		{
			txtValor.Text = Global.formataMoedaDigitada(txtValor.Text);
		}
		#endregion

		#region [ txtValor_KeyPress ]
		private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
		}
		#endregion

		#region [ txtValor_KeyDown ]
		private void txtValor_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtCnpjCpf);
		}
		#endregion

		#endregion

		#region [ txtDescricao ]

		#region [ txtDescricao_KeyPress ]
		private void txtDescricao_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#region [ txtDescricao_KeyDown ]
		private void txtDescricao_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnGravar);
		}
		#endregion

		#endregion

		#region [ txtCnpjCpf ]

		#region [ txtCnpjCpf_KeyPress ]
		private void txtCnpjCpf_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCnpjCpf(e.KeyChar);
		}
		#endregion

		#region [ txtCnpjCpf_KeyDown ]
		private void txtCnpjCpf_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtNF);
		}
		#endregion

		#region [ txtCnpjCpf_Enter ]
		private void txtCnpjCpf_Enter(object sender, EventArgs e)
		{
			txtCnpjCpf.Select(0, txtCnpjCpf.Text.Length);
		}
		#endregion

		#region [ txtCnpjCpf_Leave ]
		private void txtCnpjCpf_Leave(object sender, EventArgs e)
		{
			if (txtCnpjCpf.Text.Length == 0) return;
			txtCnpjCpf.Text = Global.formataCnpjCpf(txtCnpjCpf.Text);
			if (!Global.isCnpjCpfOk(txtCnpjCpf.Text))
			{
				avisoErro("CNPJ/CPF inválido!!");
				txtCnpjCpf.Focus();
				return;
			}
		}
		#endregion

		#endregion

		#region [ txtNF ]

		#region [ txtNF_Enter ]
		private void txtNF_Enter(object sender, EventArgs e)
		{
			txtNF.Select(0, txtNF.Text.Length);
		}
		#endregion

		#region [ txtNF_Leave ]
		private void txtNF_Leave(object sender, EventArgs e)
		{
			#region [ Declarações ]
			int numNF;
			#endregion

			numNF = (int)Global.converteInteiro(Global.digitos(txtNF.Text));
			txtNF.Text = (numNF == 0 ? "" : Global.formataInteiro(numNF));
		}
		#endregion

		#region [ txtNF_KeyDown ]
		private void txtNF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDescricao);
		}
		#endregion

		#region [ txtNF_KeyPress ]
		private void txtNF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ cbContaCorrente ]

		#region [ cbContaCorrente_KeyDown ]
		private void cbContaCorrente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasEmpresa);
		}
		#endregion

		#endregion

		#region [ cbPlanoContasEmpresa ]

		#region [ cbPlanoContasEmpresa_KeyDown ]
		private void cbPlanoContasEmpresa_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasConta);
		}
		#endregion

		#endregion

		#region [ cbPlanoContasConta ]

		#region [ cbPlanoContasConta_KeyDown ]
		private void cbPlanoContasConta_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtDataCompetencia);
		}
		#endregion

		#endregion

		#region [ Botões/Menu ]

		#region [ Gravar ]

		#region [ btnGravar_Click ]
		private void btnGravar_Click(object sender, EventArgs e)
		{
			trataBotaoGravar();
		}
		#endregion

		#region [ menuLancamentoGravar_Click ]
		private void menuLancamentoGravar_Click(object sender, EventArgs e)
		{
			trataBotaoGravar();
		}
		#endregion

		#endregion

		#region [ Limpar ]

		#region [ btnLimpar_Click ]
		private void btnLimpar_Click(object sender, EventArgs e)
		{
			trataBotaoLimpar();
		}
		#endregion

		#region [ menuLancamentoLimpar_Click ]
		private void menuLancamentoLimpar_Click(object sender, EventArgs e)
		{
			trataBotaoLimpar();
		}
		#endregion

		#endregion

		#endregion

		#endregion
	}
}
