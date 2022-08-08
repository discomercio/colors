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
	#region [ Delegate ]
	public delegate void FluxoEditaLancamentoAlteradoEventHandler();
	public delegate void FluxoEditaLancamentoExcluidoEventHandler();
	#endregion

	public partial class FFluxoEdita : Financeiro.FModelo
	{
		#region [ Eventos Customizados ]
		public event FluxoEditaLancamentoAlteradoEventHandler evtFluxoEditaLancamentoAlterado;
		public event FluxoEditaLancamentoExcluidoEventHandler evtFluxoEditaLancamentoExcluido;
		#endregion

		#region [ Atributos ]
		private bool _InicializacaoOk;
		private int _idLancamentoSelecionado;
		private bool _blnLancamentoEditadoFoiGravado = false;
		private bool _blnLancamentoFoiExcluido = false;
		private Form _formChamador = null;

		public int idLancamentoSelecionado
		{
			get { return _idLancamentoSelecionado; }
			set { _idLancamentoSelecionado = value; }
		}

		LancamentoFluxoCaixa lancamentoSelecionado;
		Cliente clienteSelecionado;

		ToolStripMenuItem menuLancamento;
		ToolStripMenuItem menuLancamentoAtualizar;
		ToolStripMenuItem menuLancamentoExcluir;
		#endregion

		#region [ Construtor ]
		public FFluxoEdita(Form formChamador)
		{
			InitializeComponent();

			_formChamador = formChamador;

			#region [ Menu Lançamento ]
			// Menu principal de Lançamento
			menuLancamento = new ToolStripMenuItem("&Lançamento");
			menuLancamento.Name = "menuLancamento";
			// Atualizar
			menuLancamentoAtualizar = new ToolStripMenuItem("&Atualizar", null, menuLancamentoAtualizar_Click);
			menuLancamentoAtualizar.Name = "menuLancamentoAtualizar";
			menuLancamento.DropDownItems.Add(menuLancamentoAtualizar);
			// Excluir
			menuLancamentoExcluir = new ToolStripMenuItem("E&xcluir", null, menuLancamentoExcluir_Click);
			menuLancamentoExcluir.Name = "menuLancamentoExcluir";
			menuLancamento.DropDownItems.Add(menuLancamentoExcluir);
			// Adiciona o menu Lançamento ao menu principal
			menuPrincipal.Items.Insert(1, menuLancamento);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ comboStSemEfeitoPosicionaDefault ]
		private bool comboStSemEfeitoPosicionaDefault()
		{
			bool blnHaDefault = false;

			foreach (Global.OpcaoFluxoCaixaStSemEfeito item in cbStSemEfeito.Items)
			{
				if (item.codigo == lancamentoSelecionado.st_sem_efeito)
				{
					cbStSemEfeito.SelectedIndex = cbStSemEfeito.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
		}
		#endregion

		#region [ comboStConfirmacaoPendentePosicionaDefault ]
		private bool comboStConfirmacaoPendentePosicionaDefault()
		{
			bool blnHaDefault = false;

			foreach (Global.OpcaoFluxoCaixaStConfirmacaoPendente item in cbStConfirmacaoPendente.Items)
			{
				if (item.codigo == lancamentoSelecionado.st_confirmacao_pendente)
				{
					cbStConfirmacaoPendente.SelectedIndex = cbStConfirmacaoPendente.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
		}
		#endregion
		
		#region [ comboCtrlPagtoStatusPosicionaDefault ]
		private bool comboCtrlPagtoStatusPosicionaDefault()
		{
			bool blnHaDefault = false;

			foreach (Global.OpcaoFluxoCaixaCtrlPagtoStatus item in cbCtrlPagtoStatus.Items)
			{
				if (item.codigo == lancamentoSelecionado.ctrl_pagto_status)
				{
					cbCtrlPagtoStatus.SelectedIndex = cbCtrlPagtoStatus.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
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
				if (rowContaCorrente.id == lancamentoSelecionado.id_conta_corrente)
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
				if (rowPlanoContasEmpresa.id == lancamentoSelecionado.id_plano_contas_empresa)
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
				if ((rowPlanoContasConta.id == lancamentoSelecionado.id_plano_contas_conta)
					 &&
					(rowPlanoContasConta.natureza == lancamentoSelecionado.natureza))
				{
					cbPlanoContasConta.SelectedIndex = cbPlanoContasConta.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
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
			LancamentoFluxoCaixa lancamentoEditado = new LancamentoFluxoCaixa();
			// O grupo de contas é obtido a partir da conta, ou seja, não é selecionado explicitamente pelo usuário
			// Lembrando que cada conta foi vinculada a um grupo de contas no momento do cadastramento
			System.Data.DataRowView dataRowView = (System.Data.DataRowView)cbPlanoContasConta.Items[cbPlanoContasConta.SelectedIndex];
			DsDataSource.DtbPlanoContasContaComboRow rowConta = (DsDataSource.DtbPlanoContasContaComboRow)dataRowView.Row;
			lancamentoEditado.id_plano_contas_grupo = (int)Global.converteInteiro(rowConta.id_plano_contas_grupo.ToString());
			// Neste painel de edição, são exibidas contas tanto de crédito quanto de débito
			// A alteração da natureza da operação é feita selecionando-se uma conta da natureza pretendida
			lancamentoEditado.natureza = rowConta.natureza;
			lancamentoEditado.id_conta_corrente = (byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString());
			lancamentoEditado.id_plano_contas_empresa = (byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString());
			lancamentoEditado.id_plano_contas_conta = (int)Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString());
			lancamentoEditado.dt_competencia = Global.converteDdMmYyyyParaDateTime(txtDataCompetencia.Text);
			lancamentoEditado.valor = Global.converteNumeroDecimal(txtValor.Text);
			lancamentoEditado.cnpj_cpf = Global.digitos(txtCnpjCpf.Text.Trim());
			lancamentoEditado.numero_NF = (int)Global.converteInteiro(Global.digitos(txtNF.Text.Trim()));
			lancamentoEditado.descricao = txtDescricao.Text.Trim();

            if (txtComp2.Enabled)
            {
                lancamentoEditado.dt_mes_competencia = Convert.ToDateTime(txtComp2.Text); 
            }

            lancamentoEditado.st_sem_efeito = (byte)Global.converteInteiro(cbStSemEfeito.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			if (lancamentoEditado.st_sem_efeito == Global.Cte.Etc.FLAG_NAO_SETADO)
			{
				avisoErro("O campo '" + lblTitStSemEfeito.Text + "' está com valor inválido: " + cbStSemEfeito.SelectedValue.ToString());
				return null;
			}
			
			lancamentoEditado.st_confirmacao_pendente = (byte)Global.converteInteiro(cbStConfirmacaoPendente.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			if (lancamentoEditado.st_confirmacao_pendente == Global.Cte.Etc.FLAG_NAO_SETADO)
			{
				avisoErro("O campo '" + lblTitStConfirmacaoPendente.Text + "' está com valor inválido: " + cbStConfirmacaoPendente.SelectedValue.ToString());
				return null;
			}

			lancamentoEditado.ctrl_pagto_status = (byte)Global.converteInteiro(cbCtrlPagtoStatus.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			if (lancamentoEditado.ctrl_pagto_status == Global.Cte.Etc.FLAG_NAO_SETADO)
			{
				avisoErro("O campo '" + lblTitCtrlPagtoStatus.Text + "' está com valor inválido: " + cbCtrlPagtoStatus.SelectedValue.ToString());
				return null;
			}

			return lancamentoEditado;
		}
		#endregion

		#region [ isLancamentoEditado ]
		/// <summary>
		/// Compara os dados dos dois objetos da classe LancamentoFluxoCaixa para verificar se o usuário fez alguma edição
		/// </summary>
		/// <param name="lancamentoOriginal">
		/// Objeto contendo os dados originais
		/// </param>
		/// <param name="lancamentoEditado">
		/// Objeto contendo os dados atuais, de acordo com o que está nos campos na tela
		/// </param>
		/// <returns>
		/// true: houve edição nos dados
		/// false: não houve nenhuma edição
		/// </returns>
		private bool isLancamentoEditado(LancamentoFluxoCaixa lancamentoOriginal, LancamentoFluxoCaixa lancamentoEditado)
		{
			if (lancamentoEditado.id_plano_contas_grupo != lancamentoSelecionado.id_plano_contas_grupo) return true;
			if (lancamentoEditado.natureza != lancamentoSelecionado.natureza) return true;
			if (lancamentoEditado.id_conta_corrente != lancamentoSelecionado.id_conta_corrente) return true;
			if (lancamentoEditado.id_plano_contas_empresa != lancamentoSelecionado.id_plano_contas_empresa) return true;
			if (lancamentoEditado.id_plano_contas_conta != lancamentoSelecionado.id_plano_contas_conta) return true;
            if (lancamentoEditado.dt_competencia != lancamentoSelecionado.dt_competencia) return true;
			if (lancamentoEditado.valor != lancamentoSelecionado.valor) return true;
			if (!Global.digitos(lancamentoEditado.cnpj_cpf).Equals(Global.digitos(lancamentoSelecionado.cnpj_cpf))) return true;
			if (lancamentoEditado.numero_NF != lancamentoSelecionado.numero_NF) return true;
			if (!lancamentoEditado.descricao.Equals(lancamentoSelecionado.descricao)) return true;
			if (lancamentoEditado.st_sem_efeito != lancamentoSelecionado.st_sem_efeito) return true;
			if (lancamentoEditado.st_confirmacao_pendente != lancamentoSelecionado.st_confirmacao_pendente) return true;
			if (lancamentoEditado.ctrl_pagto_status != lancamentoSelecionado.ctrl_pagto_status) return true;

            if (txtComp2.Enabled)
            {
                if (lancamentoEditado.dt_mes_competencia != lancamentoSelecionado.dt_mes_competencia) return true;
            }


            return false;
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

		#region [ consisteCtrlPagtoStatus ]
		/// <summary>
		/// Realiza a consistência do campo de status t_FIN_FLUXO_CAIXA.ctrl_pagto_status para 
		/// que sempre mantenha a coerência com relação ao campo t_FIN_FLUXO_CAIXA.st_sem_efeito
		/// </summary>
		/// <returns>
		/// true: consistência ok
		/// false: preenchimento está inconsistente
		/// </returns>
		private bool consisteCtrlPagtoStatus()
		{
			Global.Cte.FIN.eCtrlPagtoStatus enumCtrlPagtoStatus;
			bool blnErroConsistencia = false;
			String strMsgErroConsistencia = "";
			byte byteStSemEfeito;
			byte byteCtrlPagtoStatus;
			byteStSemEfeito = (byte)Global.converteInteiro(cbStSemEfeito.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			byteCtrlPagtoStatus = (byte)Global.converteInteiro(cbCtrlPagtoStatus.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);

			if (byteStSemEfeito == Global.Cte.Etc.FLAG_NAO_SETADO)
			{
				avisoErro("O campo " + lblTitStSemEfeito.Text + " está com valor inválido: " + cbStSemEfeito.SelectedValue.ToString());
				return false;
			}

			if (byteCtrlPagtoStatus == Global.Cte.Etc.FLAG_NAO_SETADO)
			{
				avisoErro("O campo " + lblTitCtrlPagtoStatus.Text + " está com valor inválido: " + cbCtrlPagtoStatus.SelectedValue.ToString());
				return false;
			}

			enumCtrlPagtoStatus = (Global.Cte.FIN.eCtrlPagtoStatus)byteCtrlPagtoStatus;
			switch (enumCtrlPagtoStatus)
			{
				case Global.Cte.FIN.eCtrlPagtoStatus.CONTROLE_MANUAL:
					break;
				case Global.Cte.FIN.eCtrlPagtoStatus.CADASTRADO_INICIAL:
					if (byteStSemEfeito == Global.Cte.FIN.StSemEfeito.FLAG_LIGADO)
					{
						blnErroConsistencia = true;
						strMsgErroConsistencia = "O lançamento NÃO pode ser cadastrado com\n\n" + lblTitCtrlPagtoStatus.Text + " = " + Global.retornaDescricaoFluxoCaixaCtrlPagtoStatus(enumCtrlPagtoStatus) + "\ne\n" + lblTitStSemEfeito.Text + " = " + Global.retornaDescricaoFluxoCaixaStSemEfeito(byteStSemEfeito);
					}
					break;
				case Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_BAIXADO:
					if (byteStSemEfeito == Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)
					{
						blnErroConsistencia = true;
						strMsgErroConsistencia = "O lançamento NÃO pode ser cadastrado com\n\n" + lblTitCtrlPagtoStatus.Text + " = " + Global.retornaDescricaoFluxoCaixaCtrlPagtoStatus(enumCtrlPagtoStatus) + "\ne\n" + lblTitStSemEfeito.Text + " = " + Global.retornaDescricaoFluxoCaixaStSemEfeito(byteStSemEfeito);
					}
					break;
				case Global.Cte.FIN.eCtrlPagtoStatus.CHEQUE_DEVOLVIDO:
					if (byteStSemEfeito == Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)
					{
						blnErroConsistencia = true;
						strMsgErroConsistencia = "O lançamento NÃO pode ser cadastrado com\n\n" + lblTitCtrlPagtoStatus.Text + " = " + Global.retornaDescricaoFluxoCaixaCtrlPagtoStatus(enumCtrlPagtoStatus) + "\ne\n" + lblTitStSemEfeito.Text + " = " + Global.retornaDescricaoFluxoCaixaStSemEfeito(byteStSemEfeito);
					}
					break;
				case Global.Cte.FIN.eCtrlPagtoStatus.VISA_CANCELADO:
					if (byteStSemEfeito == Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)
					{
						blnErroConsistencia = true;
						strMsgErroConsistencia = "O lançamento NÃO pode ser cadastrado com\n\n" + lblTitCtrlPagtoStatus.Text + " = " + Global.retornaDescricaoFluxoCaixaCtrlPagtoStatus(enumCtrlPagtoStatus) + "\ne\n" + lblTitStSemEfeito.Text + " = " + Global.retornaDescricaoFluxoCaixaStSemEfeito(byteStSemEfeito);
					}
					break;
				case Global.Cte.FIN.eCtrlPagtoStatus.PAGO:
					if (byteStSemEfeito == Global.Cte.FIN.StSemEfeito.FLAG_LIGADO)
					{
						blnErroConsistencia = true;
						strMsgErroConsistencia = "O lançamento NÃO pode ser cadastrado com\n\n" + lblTitCtrlPagtoStatus.Text + " = " + Global.retornaDescricaoFluxoCaixaCtrlPagtoStatus(enumCtrlPagtoStatus) + "\ne\n" + lblTitStSemEfeito.Text + " = " + Global.retornaDescricaoFluxoCaixaStSemEfeito(byteStSemEfeito);
					}
					break;
				default:
					break;
			}

			if (blnErroConsistencia)
			{
				avisoErro(strMsgErroConsistencia);
				return false;
			}

			return true;
		}
		#endregion

		#region [ trataBotaoExcluir ]
		private void trataBotaoExcluir()
		{
			#region [ Declarações ]
			String strAux;
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			bool blnResultado;
			FinLog finLog = new FinLog();
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
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

			#region[ Confirmação ]
			if (lancamentoSelecionado.dt_competencia < DateTime.Today)
			{
				#region [ Se é um lançamento de data passada, solicita a senha do login como confirmação ]
				strAux = "A data de competência deste lançamento é de uma data passada (" +
						 Global.formataDataDdMmYyyyComSeparador(lancamentoSelecionado.dt_competencia) + ")\n" +
						 "Digite a senha para confirmar a EXCLUSÃO!!";
				fAutorizacao = new FAutorizacao(strAux);
				drAutorizacao = fAutorizacao.ShowDialog();
				if (drAutorizacao != DialogResult.OK)
				{
					avisoErro("Operação não confirmada!!\nA exclusão não foi realizada!!");
					return;
				}
				if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
				{
					avisoErro("Senha inválida!!\nA exclusão não foi realizada!!");
					return;
				}
				#endregion

			}
			else
			{
				#region [ Confirmação simples ]
				if (!confirma("Confirma a EXCLUSÃO do lançamento?")) return;
				#endregion
			}
			#endregion

			#region [ Exclui do banco de dados ]
			blnResultado = LancamentoFluxoCaixaDAO.exclui(	Global.Usuario.usuario,
															lancamentoSelecionado.id,
															ref strDescricaoLog,
															ref strMsgErro
															);
			#endregion

			#region [ Processamento pós tentativa de exclusão do BD ]
			if (blnResultado)
			{
				#region [ Grava log no BD ]
				finLog.usuario = Global.Usuario.usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EXCLUI;
				finLog.natureza = lancamentoSelecionado.natureza;
				finLog.tipo_cadastro = lancamentoSelecionado.tipo_cadastro;
				finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
				finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
				finLog.id_registro_origem = lancamentoSelecionado.id;
				finLog.id_conta_corrente = lancamentoSelecionado.id_conta_corrente;
				finLog.id_plano_contas_empresa = lancamentoSelecionado.id_plano_contas_empresa;
				finLog.id_plano_contas_grupo = lancamentoSelecionado.id_plano_contas_grupo;
				finLog.id_plano_contas_conta = lancamentoSelecionado.id_plano_contas_conta;
				finLog.id_cliente = lancamentoSelecionado.id_cliente;
				finLog.cnpj_cpf = lancamentoSelecionado.cnpj_cpf;
				finLog.descricao = strDescricaoLog;
				finLog.st_sem_efeito = lancamentoSelecionado.st_sem_efeito;
				finLog.ctrl_pagto_status = lancamentoSelecionado.ctrl_pagto_status;
				FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
				#endregion

				_blnLancamentoFoiExcluido = true;
				aviso("Lançamento foi excluído!!");
				// Fecha o painel!!
				Close();
			}
			else
			{
				avisoErro("Falha ao gravar o registro!!\n\n" + strMsgErro);
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoAtualizar ]
		void trataBotaoAtualizar()
		{
			#region [ Declarações ]
			bool blnConfirmarAlteracaoDtPassadaComSenha = false;
			bool blnConfirmarAlteracao = false;
			String strAux;
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			bool blnResultado;
			bool blnHouveAlteracao = false;
			LancamentoFluxoCaixa lancamentoEditado;
			FinLog finLog = new FinLog();
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
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
			if (!consisteCtrlPagtoStatus()) return;
			#endregion

			#region [ Obtém valores ]
			lancamentoEditado = obtemDadosLancamentoCamposTela();
			if (lancamentoEditado == null) return;
			lancamentoEditado.id = lancamentoSelecionado.id;
			#endregion

			#region [ Se não houve alteração, não prossegue ]
			blnHouveAlteracao = isLancamentoEditado(lancamentoSelecionado, lancamentoEditado);
			if (!blnHouveAlteracao)
			{
				avisoErro("Não há alterações para gravar!!");
				return;
			}
			#endregion

			#region[ Confirmação ]
			if (lancamentoSelecionado.dt_competencia < DateTime.Today)
			{
				if (blnConfirmarAlteracaoDtPassadaComSenha)
				{
					#region [ Se é um lançamento de data passada, solicita a senha do login como confirmação ]
					strAux = "A data de competência deste lançamento é de uma data passada (" +
							 Global.formataDataDdMmYyyyComSeparador(lancamentoSelecionado.dt_competencia) + ")\n" +
							 "Digite a senha para confirmar a alteração!!";
					fAutorizacao = new FAutorizacao(strAux);
					drAutorizacao = fAutorizacao.ShowDialog();
					if (drAutorizacao != DialogResult.OK)
					{
						avisoErro("Operação não confirmada!!\nAs alterações não foram gravadas no banco de dados!!");
						return;
					}
					if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
					{
						avisoErro("Senha inválida!!\nAs alterações não foram gravadas no banco de dados!!");
						return;
					}
					#endregion
				}
				else
				{
					if (blnConfirmarAlteracao)
					{
						#region [ Confirmação simples ]
						strAux = "A data de competência deste lançamento é de uma data passada (" +
								 Global.formataDataDdMmYyyyComSeparador(lancamentoSelecionado.dt_competencia) + ")\n" +
								 "Confirma a gravação das alterações?";
						if (!confirma(strAux)) return;
						#endregion
					}
				}
			}
			else
			{
				if (blnConfirmarAlteracao)
				{
					#region [ Confirmação simples ]
					if (!confirma("Confirma a gravação das alterações?")) return;
					#endregion
				}
			}
			#endregion

			#region [ Grava no banco de dados ]
			blnResultado = LancamentoFluxoCaixaDAO.altera(	Global.Usuario.usuario,
															lancamentoEditado,
															ref strDescricaoLog,
															ref strMsgErro
															);
			#endregion

			#region [ Processamento pós tentativa de gravação no BD ]
			if (blnResultado)
			{
				#region [ Grava log no BD ]
				finLog.usuario = Global.Usuario.usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA;
				finLog.natureza = lancamentoEditado.natureza;
				finLog.tipo_cadastro = lancamentoSelecionado.tipo_cadastro;
				finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
				finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
				finLog.id_registro_origem = lancamentoSelecionado.id;
				finLog.id_conta_corrente = lancamentoEditado.id_conta_corrente;
				finLog.id_plano_contas_empresa = lancamentoEditado.id_plano_contas_empresa;
				finLog.id_plano_contas_grupo = lancamentoEditado.id_plano_contas_grupo;
				finLog.id_plano_contas_conta = lancamentoEditado.id_plano_contas_conta;
				finLog.id_cliente = lancamentoSelecionado.id_cliente;
				finLog.cnpj_cpf = lancamentoEditado.cnpj_cpf;
				finLog.descricao = strDescricaoLog;
				finLog.st_sem_efeito = lancamentoEditado.st_sem_efeito;
				finLog.ctrl_pagto_status = lancamentoEditado.ctrl_pagto_status;
				FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
				#endregion

				#region [ Atualiza os dados do objeto que armazena os dados originais do lançamento ]
				try
				{
					lancamentoSelecionado = LancamentoFluxoCaixaDAO.getLancamentoFluxoCaixa(_idLancamentoSelecionado);
				}
				catch (FinanceiroException ex)
				{
					avisoErro("Falha ao obter os dados do lançamento selecionado!!\n\n" + ex.Message);
					Close();
					return;
				}
				#endregion

				_blnLancamentoEditadoFoiGravado = true;
				SystemSounds.Asterisk.Play();
				// Fecha o painel!!
				Close();
			}
			else
			{
				avisoErro("Falha ao gravar o registro!!\n\n" + strMsgErro);
			}
			#endregion
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form: FFluxoEdita ]

		#region [ FFluxoEdita_Load ]
		private void FFluxoEdita_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			#endregion

			try
			{
				#region [ Obtém os dados do lançamento selecionado ]
				try
				{
					lancamentoSelecionado = LancamentoFluxoCaixaDAO.getLancamentoFluxoCaixa(_idLancamentoSelecionado);
					clienteSelecionado = ClienteDAO.getClienteCnpjCpf(lancamentoSelecionado.cnpj_cpf);

					_blnLancamentoEditadoFoiGravado = false;
					_blnLancamentoFoiExcluido = false;
				}
				catch (FinanceiroException ex)
				{
					avisoErro("Falha ao obter os dados do lançamento selecionado!!\n\n" + ex.Message);
					Close();
					return;
				}
				#endregion

				#region [ Combo StSemEfeito ]
				cbStSemEfeito.DataSource = Global.montaOpcaoFluxoCaixaStSemEfeito(Global.eOpcaoIncluirItemTodos.NAO_INCLUIR);
				cbStSemEfeito.DisplayMember = "descricao";
				cbStSemEfeito.ValueMember = "codigo";
				if (!comboStSemEfeitoPosicionaDefault())
				{
					cbStSemEfeito.SelectedIndex = -1;
					avisoErro("Falha ao tentar posicionar a lista de opções '" + lblTitStSemEfeito.Text + "' no código: " + lancamentoSelecionado.st_sem_efeito.ToString() + "!!");
				}
				#endregion

				#region [ Combo StConfirmacaoPendente ]
				cbStConfirmacaoPendente.DataSource = Global.montaOpcaoFluxoCaixaStConfirmacaoPendente(Global.eOpcaoIncluirItemTodos.NAO_INCLUIR);
				cbStConfirmacaoPendente.DisplayMember = "descricao";
				cbStConfirmacaoPendente.ValueMember = "codigo";
				if (!comboStConfirmacaoPendentePosicionaDefault())
				{
					cbStConfirmacaoPendente.SelectedIndex = -1;
					avisoErro("Falha ao tentar posicionar a lista de opções '" + lblTitStConfirmacaoPendente.Text + "' no código: " + lancamentoSelecionado.st_confirmacao_pendente.ToString() + "!!");
				}
				#endregion

				#region [ Combo CtrlPagtoStatus ]
				cbCtrlPagtoStatus.DataSource = Global.montaOpcaoFluxoCaixaCtrlPagtoStatus(Global.eOpcaoIncluirItemTodos.NAO_INCLUIR);
				cbCtrlPagtoStatus.DisplayMember = "descricao";
				cbCtrlPagtoStatus.ValueMember = "codigo";
				if (!comboCtrlPagtoStatusPosicionaDefault())
				{
					cbCtrlPagtoStatus.SelectedIndex = -1;
					avisoErro("Falha ao tentar posicionar a lista de opções '" + lblTitCtrlPagtoStatus.Text + "' no código: " + lancamentoSelecionado.ctrl_pagto_status.ToString() + "!!");
				}
				#endregion

				#region [ Combo Conta Corrente ]
				cbContaCorrente.DataSource = ComboDAO.criaDtbContaCorrenteCombo(ComboDAO.eFiltraStAtivo.TODOS);
				cbContaCorrente.ValueMember = "id";
				cbContaCorrente.DisplayMember = "contaComDescricao";
				if (!comboContaCorrentePosicionaDefault())
				{
					cbContaCorrente.SelectedIndex = -1;
					if (lancamentoSelecionado.id_conta_corrente != 0) avisoErro("Falha ao tentar posicionar a lista na conta corrente " + lancamentoSelecionado.id_conta_corrente.ToString() + "!!");
				}
				#endregion

				#region [ Combo Plano Contas Empresa ]
				cbPlanoContasEmpresa.DataSource = ComboDAO.criaDtbPlanoContasEmpresaCombo(ComboDAO.eFiltraStAtivo.TODOS);
				cbPlanoContasEmpresa.ValueMember = "id";
				cbPlanoContasEmpresa.DisplayMember = "idComDescricao";
				if (!comboPlanoContasEmpresaPosicionaDefault())
				{
					cbPlanoContasEmpresa.SelectedIndex = -1;
					if (lancamentoSelecionado.id_plano_contas_empresa != 0) avisoErro("Falha ao tentar posicionar a lista na empresa " + lancamentoSelecionado.id_plano_contas_empresa.ToString() + "!!");
				}
				#endregion

				#region [ Combo Plano Contas Conta ]
				cbPlanoContasConta.DataSource = ComboDAO.criaDtbPlanoContasContaCombo(ComboDAO.eFiltraNatureza.TODOS, ComboDAO.eFiltraStAtivo.TODOS, ComboDAO.eFiltraStSistema.TODOS);
				cbPlanoContasConta.ValueMember = "id";
				cbPlanoContasConta.DisplayMember = "idComDescricao";
				if (!comboPlanoContasContaPosicionaDefault())
				{
					cbPlanoContasConta.SelectedIndex = -1;
					if (lancamentoSelecionado.id_plano_contas_conta != 0) avisoErro("Falha ao tentar posicionar a lista na conta " + lancamentoSelecionado.id_plano_contas_conta.ToString() + "!!");
				}
				if (lancamentoSelecionado.tipo_cadastro == Global.Cte.FIN.TipoCadastro.SISTEMA) cbPlanoContasConta.Enabled = false;
				#endregion

				#region [ Campo descrição: tamanho máximo ]
				txtDescricao.MaxLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
				#endregion

				#region [ Demais campos ]
				txtDataCompetencia.Text = Global.formataDataDdMmYyyyComSeparador(lancamentoSelecionado.dt_competencia);
				txtValor.Text = Global.formataMoeda(lancamentoSelecionado.valor);
				txtCnpjCpf.Text = Global.formataCnpjCpf(lancamentoSelecionado.cnpj_cpf);
				txtNF.Text = (lancamentoSelecionado.numero_NF == 0 ? "" : Global.formataInteiro(lancamentoSelecionado.numero_NF));
				txtDescricao.Text = lancamentoSelecionado.descricao;
				lblCadastradoEm.Text = Global.formataDataDdMmYyyyHhMmComSeparador(lancamentoSelecionado.dt_hr_cadastro);
				lblCadastradoPor.Text = lancamentoSelecionado.usuario_cadastro;
				lblCadastradoModo.Text = Global.retornaDescricaoTipoCadastramento(lancamentoSelecionado.tipo_cadastro);
				if ((lancamentoSelecionado.editado_manual == Global.Cte.FIN.EditadoManual.SIM) || (Global.Parametro.FluxoCaixa_ConsiderarDataAtualizacaoAutomatica == 1))
				{
					lblAlteradoEm.Text = Global.formataDataDdMmYyyyHhMmComSeparador(lancamentoSelecionado.dt_hr_ult_atualizacao);
					lblAlteradoPor.Text = lancamentoSelecionado.usuario_ult_atualizacao;
				}
				else
				{
					lblAlteradoEm.Text = "";
					lblAlteradoPor.Text = "";
				}
				lblNatureza.Text = Global.retornaDescricaoFluxoCaixaNatureza(lancamentoSelecionado.natureza);
				lblNatureza.ForeColor = Global.retornaCorFluxoCaixaNatureza(lancamentoSelecionado.natureza);
				
				if (clienteSelecionado == null)
					lblNome.Text = "";
				else
					lblNome.Text = clienteSelecionado.nome;

                if (!lancamentoSelecionado.natureza.ToString().ToUpper().Equals("D"))
                {
                    lblTitComp2.Enabled = false;
                    lblTitComp2.Visible = false;

                    txtComp2.Enabled = false;
                    txtComp2.Visible = false;
                }
                else
                {
                    txtComp2.Text = lancamentoSelecionado.dt_mes_competencia.ToString("MM/yyyy");
                }
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

		#region [ FFluxoEdita_Shown ]
		private void FFluxoEdita_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					btnDummy.Focus();

					((FModelo)_formChamador).info(ModoExibicaoMensagemRodape.Normal);

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

		#region [ FFluxoEdita_FormClosing ]
		private void FFluxoEdita_FormClosing(object sender, FormClosingEventArgs e)
		{
			LancamentoFluxoCaixa lancamentoEditado;

			#region [ Trata situação em que lançamento foi excluído ]
			if (_blnLancamentoFoiExcluido)
			{
				// Aciona evento para refazer a pesquisa de lançamentos e atualizar os dados do grid
				if (evtFluxoEditaLancamentoExcluido != null) evtFluxoEditaLancamentoExcluido();
				return;
			}
			#endregion

			#region [ Verifica se houve alterações não salvas ]
			lancamentoEditado = obtemDadosLancamentoCamposTela();
			if (lancamentoEditado != null)
			{
				if (isLancamentoEditado(lancamentoSelecionado, lancamentoEditado))
				{
					if (!confirma("As alterações serão perdidas!!\nContinua assim mesmo?"))
					{
						e.Cancel = true;
						return;
					}
				}
			}
			#endregion

			#region [ Trata situação em que o lançamento foi alterado ]
			if (_blnLancamentoEditadoFoiGravado)
			{
				// Aciona evento para refazer a pesquisa de lançamentos e atualizar os dados do grid
				if (evtFluxoEditaLancamentoAlterado != null) evtFluxoEditaLancamentoAlterado();
			}
			#endregion
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
			if (txtDataCompetencia.Text.Length == 0) return;
			txtDataCompetencia.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCompetencia.Text);
			if (!Global.isDataOk(txtDataCompetencia.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCompetencia.Focus();
				return;
			}
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
			Global.trataTextBoxKeyDown(sender, e, btnAtualizar);
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

			if (txtNF.Text.Length == 0) return;

			numNF = (int)Global.converteInteiro(Global.digitos(txtNF.Text.Trim()));
			if (numNF < 0)
			{
				avisoErro("Número de NF inválido!!");
				txtNF.Focus();
				return;
			}

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

		#region [ cbStSemEfeito ]
		private void cbStSemEfeito_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbStConfirmacaoPendente);
		}
		#endregion

		#region [ cbStConfirmacaoPendente ]
		private void cbStConfirmacaoPendente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbCtrlPagtoStatus);
		}
		#endregion

		#region [ cbCtrlPagtoStatus ]
		private void cbCtrlPagtoStatus_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbContaCorrente);
		}
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

		#region [ Atualizar ]

		#region [ btnAtualizar_Click ]
		private void btnAtualizar_Click(object sender, EventArgs e)
		{
			trataBotaoAtualizar();
		}
		#endregion

		#region [ menuLancamentoAtualizar_Click ]
		private void menuLancamentoAtualizar_Click(object sender, EventArgs e)
		{
			trataBotaoAtualizar();
		}
		#endregion

		#endregion

		#region [ Excluir ]

		#region [ btnExcluir_Click ]
		private void btnExcluir_Click(object sender, EventArgs e)
		{
			trataBotaoExcluir();
		}
		#endregion

		#region [ menuLancamentoExcluir_Click ]
		private void menuLancamentoExcluir_Click(object sender, EventArgs e)
		{
			trataBotaoExcluir();
		}
		#endregion

		#endregion

		#endregion

		#endregion
	}
}
