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
	public partial class FFluxoCreditoLote : Financeiro.FModelo
	{
		#region [ Constantes ]
		const String COL_PLANO_CONTAS_CONTA = "colPlanoContasConta";
		const String COL_DATA_COMPETENCIA = "colDataCompetencia";
		const String COL_VALOR_LANCTO = "colValorLancto";
		const String COL_CNPJ_CPF = "colCnpjCpf";
		const String COL_NF = "colNF";
		const String COL_DESCRICAO = "colDescricao";
		const int LINHAS_GRID_LANCAMENTOS_DEFAULT = 80;
		#endregion

		#region [ Atributos ]
		private bool _InicializacaoOk;
		ToolStripMenuItem menuLancamento;
		ToolStripMenuItem menuLancamentoGravar;
		ToolStripMenuItem menuLancamentoLimpar;
		private Form _formChamador = null;
		private bool _blnComboPlanoContasPreencherAutomatico = false;
		private int LINHAS_GRID_LANCAMENTOS;
		#endregion

		#region [ Construtor ]
		public FFluxoCreditoLote(Form formChamador, int qtdeLancamentos)
		{
			InitializeComponent();

			_formChamador = formChamador;

			LINHAS_GRID_LANCAMENTOS = (qtdeLancamentos <= 0 ? LINHAS_GRID_LANCAMENTOS_DEFAULT : qtdeLancamentos);
			lblTitulo.Text += " (" + LINHAS_GRID_LANCAMENTOS.ToString() + ")";

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

		#region [ iniciaGrid ]
		private void iniciaGrid()
		{
			#region [ Declarações ]
			DsDataSource.DtbPlanoContasContaComboDataTable dtbPlanoContasConta;
			DsDataSource.DtbPlanoContasContaComboRow rowPlanoContasConta;
			#endregion

			#region [ Linhas do grid ]
			grdLote.Rows.Clear();
			grdLote.Rows.Add(LINHAS_GRID_LANCAMENTOS);
			#endregion

			#region [ Coluna com combo Plano Contas Conta ]
			dtbPlanoContasConta = ComboDAO.criaDtbPlanoContasContaCombo(ComboDAO.eFiltraNatureza.SOMENTE_CREDITO, ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS, ComboDAO.eFiltraStSistema.SOMENTE_CONTAS_NORMAIS);

			#region [ Inclui linha vazia ]
			rowPlanoContasConta = dtbPlanoContasConta.NewDtbPlanoContasContaComboRow();
			dtbPlanoContasConta.Rows.InsertAt(rowPlanoContasConta, 0);
			#endregion

			((DataGridViewComboBoxColumn)grdLote.Columns[COL_PLANO_CONTAS_CONTA]).ValueMember = "id";
			((DataGridViewComboBoxColumn)grdLote.Columns[COL_PLANO_CONTAS_CONTA]).DisplayMember = "idComDescricao";
			((DataGridViewComboBoxColumn)grdLote.Columns[COL_PLANO_CONTAS_CONTA]).DataSource = dtbPlanoContasConta;
			#endregion
		}
		#endregion

		#region [ limpaCampos ]
		void limpaCampos()
		{
			cbContaCorrente.SelectedIndex = -1;
			cbPlanoContasEmpresa.SelectedIndex = -1;
			cbPlanoContasConta.SelectedIndex = -1;
			txtDataCompetencia.Text = "";
			txtValor.Text = "";
			txtCnpjCpf.Text = "";
			txtDescricao.Text = "";
			iniciaGrid();
			lblQtdeLancamentos.Text = "00";
			lblValorTotal.Text = Global.formataMoeda(0m);
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
				if (rowPlanoContasConta.id == Global.Usuario.Defaults.planoContasContaCredito)
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

		#region [ obtemDadosLancamentoLinhaGrid ]
		/// <summary>
		/// Carrega os dados da linha especificada do grid em um objeto da classe LancamentoFluxoCaixa
		/// </summary>
		/// <returns>
		/// Retorna um objeto LancamentoFluxoCaixa com os dados da linha especificada do grid
		/// </returns>
		private LancamentoFluxoCaixa obtemDadosLancamentoLinhaGrid(int linhaGrid)
		{
			#region [ Declarações ]
			int idPlanoContasConta = 0;
			int idPlanoContasGrupo = 0;
			LancamentoFluxoCaixa lancamento = new LancamentoFluxoCaixa();
			DsDataSource.DtbPlanoContasContaComboDataTable dtbPlanoContasConta;
			DsDataSource.DtbPlanoContasContaComboRow rowPlanoContasConta;
			#endregion

			// O grupo de contas é obtido a partir da conta, ou seja, não é selecionado explicitamente pelo usuário
			// Lembrando que cada conta foi vinculada a um grupo de contas no momento do cadastramento
			if (grdLote.Rows[linhaGrid].Cells[COL_PLANO_CONTAS_CONTA].Value != null)
			{
				idPlanoContasConta = (int)Global.converteInteiro(grdLote.Rows[linhaGrid].Cells[COL_PLANO_CONTAS_CONTA].Value.ToString());
			}

			dtbPlanoContasConta = (DsDataSource.DtbPlanoContasContaComboDataTable)((DataGridViewComboBoxColumn)grdLote.Columns[COL_PLANO_CONTAS_CONTA]).DataSource;
			for (int i = 0; i < dtbPlanoContasConta.Rows.Count; i++)
			{
				rowPlanoContasConta = dtbPlanoContasConta[i];
				if (!rowPlanoContasConta.IsidNull())
				{
					if (rowPlanoContasConta.id == idPlanoContasConta)
					{
						idPlanoContasGrupo = (int)Global.converteInteiro(rowPlanoContasConta.id_plano_contas_grupo.ToString());
						break;
					}
				}
			}
			lancamento.id_plano_contas_grupo = idPlanoContasGrupo;
			lancamento.id_conta_corrente = (byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString());
			lancamento.id_plano_contas_empresa = (byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString());
			lancamento.id_plano_contas_conta = idPlanoContasConta;
			lancamento.dt_competencia = Global.converteDdMmYyyyParaDateTime(grdLote.Rows[linhaGrid].Cells[COL_DATA_COMPETENCIA].Value.ToString());
			lancamento.valor = Global.converteNumeroDecimal(grdLote.Rows[linhaGrid].Cells[COL_VALOR_LANCTO].Value.ToString());
			lancamento.cnpj_cpf = grdLote.Rows[linhaGrid].Cells[COL_CNPJ_CPF].Value != null ? Global.digitos(grdLote.Rows[linhaGrid].Cells[COL_CNPJ_CPF].Value.ToString()) : "";
			lancamento.numero_NF = grdLote.Rows[linhaGrid].Cells[COL_NF].Value != null ? (int)Global.converteInteiro(Global.digitos(grdLote.Rows[linhaGrid].Cells[COL_NF].Value.ToString())) : 0;
			lancamento.descricao = grdLote.Rows[linhaGrid].Cells[COL_DESCRICAO].Value.ToString().Trim();

			return lancamento;
		}
		#endregion

		#region [ isLinhaGridComAlgumDado ]
		private bool isLinhaGridComAlgumDado(int linhaGrid)
		{
			if (linhaGrid < 0) return false;
			if (linhaGrid >= grdLote.Rows.Count) return false;

			if (grdLote.Rows[linhaGrid].Cells[COL_PLANO_CONTAS_CONTA].Value != null)
			{
				if (grdLote.Rows[linhaGrid].Cells[COL_PLANO_CONTAS_CONTA].Value.ToString().Trim().Length > 0) return true;
			}

			if (grdLote.Rows[linhaGrid].Cells[COL_DATA_COMPETENCIA].Value != null)
			{
				if (grdLote.Rows[linhaGrid].Cells[COL_DATA_COMPETENCIA].Value.ToString().Trim().Length > 0) return true;
			}
			if (grdLote.Rows[linhaGrid].Cells[COL_VALOR_LANCTO].Value != null)
			{
				if (grdLote.Rows[linhaGrid].Cells[COL_VALOR_LANCTO].Value.ToString().Trim().Length > 0) return true;
			}
			if (grdLote.Rows[linhaGrid].Cells[COL_CNPJ_CPF].Value != null)
			{
				if (grdLote.Rows[linhaGrid].Cells[COL_CNPJ_CPF].Value.ToString().Trim().Length > 0) return true;
			}
			if (grdLote.Rows[linhaGrid].Cells[COL_NF].Value != null)
			{
				if (grdLote.Rows[linhaGrid].Cells[COL_NF].Value.ToString().Trim().Length > 0) return true;
			}
			if (grdLote.Rows[linhaGrid].Cells[COL_DESCRICAO].Value != null)
			{
				if (grdLote.Rows[linhaGrid].Cells[COL_DESCRICAO].Value.ToString().Trim().Length > 0) return true;
			}

			return false;
		}
		#endregion

		#region [ isLinhaGridPreenchidaOk ]
		private bool isLinhaGridPreenchidaOk(int linhaGrid)
		{
			if (linhaGrid < 0) return false;
			if (linhaGrid >= grdLote.Rows.Count) return false;

			if (grdLote.Rows[linhaGrid].Cells[COL_PLANO_CONTAS_CONTA].Value == null) return false;
			if (grdLote.Rows[linhaGrid].Cells[COL_PLANO_CONTAS_CONTA].Value.ToString().Trim().Length == 0) return false;

			if (grdLote.Rows[linhaGrid].Cells[COL_DATA_COMPETENCIA].Value == null) return false;
			if (grdLote.Rows[linhaGrid].Cells[COL_DATA_COMPETENCIA].Value.ToString().Trim().Length == 0) return false;

			if (grdLote.Rows[linhaGrid].Cells[COL_VALOR_LANCTO].Value == null) return false;
			if (grdLote.Rows[linhaGrid].Cells[COL_VALOR_LANCTO].Value.ToString().Trim().Length == 0) return false;

			if (grdLote.Rows[linhaGrid].Cells[COL_DESCRICAO].Value == null) return false;
			if (grdLote.Rows[linhaGrid].Cells[COL_DESCRICAO].Value.ToString().Trim().Length == 0) return false;

			return true;
		}
		#endregion

		#region [ isLinhaGridRepetida ]
		private bool isLinhaGridRepetida(int linhaGrid1, int linhaGrid2)
		{
			#region [ Declarações ]
			LancamentoFluxoCaixa lancamento1;
			LancamentoFluxoCaixa lancamento2;
			#endregion

			if (!isLinhaGridComAlgumDado(linhaGrid1)) return false;
			if (!isLinhaGridComAlgumDado(linhaGrid2)) return false;

			lancamento1 = obtemDadosLancamentoLinhaGrid(linhaGrid1);
			lancamento2 = obtemDadosLancamentoLinhaGrid(linhaGrid2);

			if (lancamento1.id_plano_contas_conta != lancamento2.id_plano_contas_conta) return false;
			if (lancamento1.dt_competencia != lancamento2.dt_competencia) return false;
			if (lancamento1.valor != lancamento2.valor) return false;
			if ((lancamento1.cnpj_cpf.ToString().Trim().Length > 0) && (lancamento2.cnpj_cpf.ToString().Trim().Length > 0))
			{
				if (!lancamento1.cnpj_cpf.ToString().Trim().Equals(lancamento2.cnpj_cpf.ToString().Trim())) return false;
			}
			if (lancamento1.numero_NF != lancamento2.numero_NF) return false;
			if (!lancamento1.descricao.ToString().Trim().ToUpper().Equals(lancamento2.descricao.ToString().Trim().ToUpper())) return false;

			return true;
		}
		#endregion

		#region [ isGridComAlgumDado ]
		private bool isGridComAlgumDado()
		{
			for (int i = 0; i < grdLote.Rows.Count; i++)
			{
				if (isLinhaGridComAlgumDado(i)) return true;
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
			#region [ Declarações ]
			int intNumLinha = 0;
			#endregion

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

			for (int intCounter = 0; intCounter < grdLote.Rows.Count; intCounter++)
			{
				intNumLinha++;

				if (isLinhaGridComAlgumDado(intCounter))
				{
					#region [ Plano de contas ]
					if (grdLote.Rows[intCounter].Cells[COL_PLANO_CONTAS_CONTA].Value == null)
					{
						avisoErro("Informe um plano de contas no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					if (grdLote.Rows[intCounter].Cells[COL_PLANO_CONTAS_CONTA].Value.ToString().Trim().Length == 0)
					{
						avisoErro("Selecione um plano de contas no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					#endregion

					#region [ Data de competência ]
					if (grdLote.Rows[intCounter].Cells[COL_DATA_COMPETENCIA].Value == null)
					{
						avisoErro("Informe a data de competência no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					if (!Global.isDataOk(grdLote.Rows[intCounter].Cells[COL_DATA_COMPETENCIA].Value.ToString().Trim()))
					{
						avisoErro("Data de competência inválida no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					#endregion

					#region [ Valor ]
					if (grdLote.Rows[intCounter].Cells[COL_VALOR_LANCTO].Value == null)
					{
						avisoErro("Informe o valor do lançamento na linha " + intNumLinha.ToString());
						return false;
					}
					if (Global.converteNumeroDecimal(grdLote.Rows[intCounter].Cells[COL_VALOR_LANCTO].Value.ToString().Trim()) <= 0)
					{
						avisoErro("Valor inválido no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					#endregion

					#region [ CNPJ/CPF ]
					if (grdLote.Rows[intCounter].Cells[COL_CNPJ_CPF].Value != null)
					{
						if (grdLote.Rows[intCounter].Cells[COL_CNPJ_CPF].Value.ToString().Trim().Length > 0)
						{
							if (!Global.isCnpjCpfOk(grdLote.Rows[intCounter].Cells[COL_CNPJ_CPF].Value.ToString().Trim()))
							{
								avisoErro("CNPJ/CPF inválido no lançamento da linha " + intNumLinha.ToString());
								return false;
							}
						}
					}
					#endregion

					#region [ NF ]
					if (grdLote.Rows[intCounter].Cells[COL_NF].Value != null)
					{
						if (grdLote.Rows[intCounter].Cells[COL_NF].Value.ToString().Trim().Length > 0)
						{
							if ((int)Global.converteInteiro(Global.digitos(grdLote.Rows[intCounter].Cells[COL_NF].Value.ToString().Trim())) < 0)
							{
								avisoErro("Número de NF inválido no lançamento da linha " + intNumLinha.ToString());
								return false;
							}
						}
					}
					#endregion

					#region [ Descrição ]
					if (grdLote.Rows[intCounter].Cells[COL_DESCRICAO].Value == null)
					{
						avisoErro("Preencha a descrição no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					if (grdLote.Rows[intCounter].Cells[COL_DESCRICAO].Value.ToString().Trim().Length == 0)
					{
						avisoErro("Informe a descrição no lançamento da linha " + intNumLinha.ToString());
						return false;
					}
					#endregion
				}
			}

			#region [ Há linhas repetidas? ]
			for (int intCounterExt = 0; intCounterExt < (grdLote.Rows.Count - 1); intCounterExt++)
			{
				for (int intCounterInt = (intCounterExt + 1); intCounterInt < grdLote.Rows.Count; intCounterInt++)
				{
					if (isLinhaGridRepetida(intCounterExt, intCounterInt))
					{
						avisoErro("Os lançamentos das linhas " + (intCounterExt + 1).ToString() + " e " + (intCounterInt + 1).ToString() + " são iguais!!");
						return false;
					}
				}
			}
			#endregion

			// Ok!
			return true;
		}
		#endregion

		#region [ trataBotaoGravar ]
		void trataBotaoGravar()
		{
			#region [ Declarações ]
			int contadorLancamentoCreditoLoteInseridoOriginal;
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			bool blnResultado = false;
			bool blnSucesso = false;
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

			contadorLancamentoCreditoLoteInseridoOriginal = Global.contadorLancamentoCreditoLoteInserido;

			info(ModoExibicaoMensagemRodape.EmExecucao, "gravando lançamentos");
			try
			{
				try
				{
					BD.iniciaTransacao();

					for (int intCounter = 0; intCounter < grdLote.Rows.Count; intCounter++)
					{
						#region [ Linha do grid tem dados? ]
						if (!isLinhaGridPreenchidaOk(intCounter))
						{
							if (isLinhaGridComAlgumDado(intCounter)) throw new Exception("Inconsistência encontrada na linha " + (intCounter + 1).ToString());
							continue;
						}
						#endregion

						#region [ Obtém valores ]
						lancamento = obtemDadosLancamentoLinhaGrid(intCounter);
						lancamento.natureza = Global.Cte.FIN.Natureza.CREDITO;
						lancamento.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
						#endregion

						#region [ Grava no banco de dados ]
						blnResultado = LancamentoFluxoCaixaDAO.insere(Global.Usuario.usuario,
																		lancamento,
																		ref strDescricaoLog,
																		ref strMsgErro
																		);
						#endregion

						#region [ Processamento pós tentativa de gravação no BD ]
						if (blnResultado)
						{
							#region [ Incrementa contador de lançamentos gravados ]
							Global.contadorLancamentoCreditoInserido++;
							lblContador.Text = Global.contadorLancamentoCreditoInserido.ToString().PadLeft(2, '0');
							#endregion

							#region [ Atualiza defaults do usuário ]
							if (cbContaCorrente.SelectedValue != null) Global.Usuario.Defaults.contaCorrente = (byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString());
							if (cbPlanoContasEmpresa.SelectedValue != null) Global.Usuario.Defaults.planoContasEmpresa = (byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString());
							#endregion

							#region [ Grava log no BD ]
							finLog.usuario = Global.Usuario.usuario;
							finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_CREDITO_LOTE_INSERE;
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
						}
						else
						{
							throw new Exception("Falha ao gravar o registro!!\n\n" + strMsgErro);
						}
						#endregion
					}

					// Gravou todos com sucesso!
					blnSucesso = true;
				}
				finally
				{
					#region [ Commit / Rollback ]
					if (blnSucesso)
					{
						#region [ Commit ]
						try
						{
							BD.commitTransacao();
						}
						catch (Exception ex)
						{
							blnSucesso = false;
							Global.gravaLogAtividade(ex.ToString());
							avisoErro(ex.ToString());
						}
						#endregion
					}
					else
					{
						#region [ Rollback ]
						try
						{
							Global.contadorLancamentoCreditoLoteInserido = contadorLancamentoCreditoLoteInseridoOriginal;
							lblContador.Text = Global.contadorLancamentoCreditoLoteInserido.ToString().PadLeft(2, '0');
							BD.rollbackTransacao();
						}
						catch (Exception ex)
						{
							Global.gravaLogAtividade(ex.ToString());
							avisoErro(ex.ToString());
						}
						#endregion
					}
					#endregion
				}
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				avisoErro(ex.ToString());
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}

			if (blnSucesso)
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "reiniciando painel");
				try
				{
					#region [ Prepara para cadastrar próximo lote de lançamentos ]
					limpaCampos();
					if (!comboContaCorrentePosicionaDefault()) cbContaCorrente.SelectedIndex = -1;
					if (!comboPlanoContasEmpresaPosicionaDefault()) cbPlanoContasEmpresa.SelectedIndex = -1;
					if (!comboPlanoContasContaPosicionaDefault()) cbPlanoContasConta.SelectedIndex = -1;
					posicionaFocoPrimeiroCampoPreencher();
					#endregion

					SystemSounds.Asterisk.Play();
				}
				finally
				{
					info(ModoExibicaoMensagemRodape.Normal);
				}
			}
		}
		#endregion

		#region [ trataBotaoLimpar ]
		void trataBotaoLimpar()
		{
			info(ModoExibicaoMensagemRodape.EmExecucao, "reiniciando painel");
			try
			{
				limpaCampos();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ atualizaValorTotalLancamentos ]
		private void atualizaValorTotalLancamentos()
		{
			#region [ Declarações ]
			int intQtdeLancamentos = 0;
			decimal vlAux;
			decimal vlTotal = 0;
			#endregion

			for (int i = 0; i < grdLote.Rows.Count; i++)
			{
				if (grdLote.Rows[i].Cells[COL_VALOR_LANCTO].Value != null)
				{
					vlAux = Global.converteNumeroDecimal(grdLote.Rows[i].Cells[COL_VALOR_LANCTO].Value.ToString().Trim());
					if (vlAux > 0)
					{
						vlTotal += vlAux;
						intQtdeLancamentos++;
					}
				}
			}

			lblQtdeLancamentos.Text = intQtdeLancamentos.ToString().PadLeft(2, '0');
			lblValorTotal.Text = Global.formataMoeda(vlTotal);
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FFluxoCredito ]

		#region [ FFluxoCredito_Load ]
		private void FFluxoCredito_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			#endregion

			try
			{
				#region [ Limpa campos ]
				limpaCampos();
				lblContador.Text = Global.contadorLancamentoCreditoInserido.ToString().PadLeft(2, '0');
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
				cbPlanoContasConta.DataSource = ComboDAO.criaDtbPlanoContasContaCombo(ComboDAO.eFiltraNatureza.SOMENTE_CREDITO, ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS, ComboDAO.eFiltraStSistema.SOMENTE_CONTAS_NORMAIS);
				cbPlanoContasConta.ValueMember = "id";
				cbPlanoContasConta.DisplayMember = "idComDescricao";
				if (Global.Usuario.Defaults.planoContasContaCredito == 0)
					cbPlanoContasConta.SelectedIndex = -1;
				else
					if (!comboPlanoContasContaPosicionaDefault()) cbPlanoContasConta.SelectedIndex = -1;
				#endregion

				#region [ Grid de lançamentos ]
				iniciaGrid();
				#endregion

				#region [ Campo descrição ]
				txtDescricao.MaxLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
				#endregion

				#region [ Tamanho da coluna descrição do grid ]
				((DataGridViewTextBoxColumn)grdLote.Columns[COL_DESCRICAO]).MaxInputLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
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

		#region [ FFluxoCredito_Shown ]
		private void FFluxoCredito_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					posicionaFocoPrimeiroCampoPreencher();
					#endregion

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

		#region [ FFluxoCreditoLote_FormClosing ]
		private void FFluxoCreditoLote_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (this.ActiveControl == grdLote)
			{
				if (grdLote.CurrentCell.IsInEditMode)
				{
					e.Cancel = true;
					return;
				}
			}

			if (isGridComAlgumDado())
			{
				if (!confirma("As alterações serão perdidas!\nContinua mesmo assim?"))
				{
					e.Cancel = true;
					return;
				}
			}
		}
		#endregion

		#region [ FFluxoCreditoLote_KeyDown ]
		private void FFluxoCreditoLote_KeyDown(object sender, KeyEventArgs e)
		{
			#region [ Declarações ]
			DataGridViewCell celula = null;
			string sCellNewValue;
			string sCellPreviousLineValue;
			DateTime dtDate;
			#endregion

			if (this.ActiveControl == grdLote)
			{
				#region [ DEL: limpa a célula selecionada ]
				if (e.KeyCode == Keys.Delete)
				{
					foreach (var item in grdLote.SelectedCells)
					{
						if (!((DataGridViewCell)item).ReadOnly)
						{
							((DataGridViewCell)item).Value = null;
						}
					}
					e.SuppressKeyPress = true;
				}
				#endregion

				#region [ ESC: ignora tecla p/ não fechar o form ]
				if (e.KeyCode == Keys.Escape)
				{
					e.SuppressKeyPress = true;
				}
				#endregion

				#region [ Teclas especiais p/ preencher a célula c/ o valor padrão ]
				if (Global.isTeclaEspecialCopiarValorPadrao(e))
				{
					celula = grdLote.CurrentCell;
					if (celula == null) return;

					if (grdLote.Columns[celula.ColumnIndex].Name.Equals(COL_PLANO_CONTAS_CONTA))
					{
						e.SuppressKeyPress = true;
						e.Handled = true;
						_blnComboPlanoContasPreencherAutomatico = true;
						if (celula.IsInEditMode) grdLote.CancelEdit();
						grdLote.BeginEdit(false);
					}
					else if (grdLote.Columns[celula.ColumnIndex].Name.Equals(COL_DATA_COMPETENCIA))
					{
						// Se a tecla ALT estiver pressionada e houver data preenchida na linha anterior do grid, preenche esta linha somando 1 mês à data da linha anterior
						e.SuppressKeyPress = true;
						e.Handled = true;
						sCellNewValue = txtDataCompetencia.Text;
						if ((e.Alt) && (celula.RowIndex >= 1))
						{
							sCellPreviousLineValue = (grdLote[celula.ColumnIndex, celula.RowIndex - 1].Value ?? "").ToString();
							if (sCellPreviousLineValue.Length > 0)
							{
								dtDate = Global.converteDdMmYyyyParaDateTime(sCellPreviousLineValue);
								dtDate = dtDate.AddMonths(1);
								sCellNewValue = Global.formataDataDdMmYyyyComSeparador(dtDate);
							}
						}
						celula.Value = sCellNewValue;
						grdLote.focusNextEditableCell();
					}
					else if (grdLote.Columns[celula.ColumnIndex].Name.Equals(COL_VALOR_LANCTO))
					{
						e.SuppressKeyPress = true;
						e.Handled = true;
						celula.Value = txtValor.Text;
						grdLote.focusNextEditableCell();
					}
					else if (grdLote.Columns[celula.ColumnIndex].Name.Equals(COL_CNPJ_CPF))
					{
						e.SuppressKeyPress = true;
						e.Handled = true;
						celula.Value = txtCnpjCpf.Text;
						grdLote.focusNextEditableCell();
					}
					else if (grdLote.Columns[celula.ColumnIndex].Name.Equals(COL_NF))
					{
						e.SuppressKeyPress = true;
						e.Handled = true;
						grdLote.focusNextEditableCell();
					}
					else if (grdLote.Columns[celula.ColumnIndex].Name.Equals(COL_DESCRICAO))
					{
						e.SuppressKeyPress = true;
						e.Handled = true;
						celula.Value = txtDescricao.Text;
						grdLote.focusNextEditableCell();
					}
				}
				#endregion
			}
		}
		#endregion

		#region [ FFluxoCreditoLote_KeyPress ]
		private void FFluxoCreditoLote_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (this.ActiveControl == null) return;
			if (this.ActiveControl is DataGridViewTextBoxEditingControl)
			{
				if (grdLote.CurrentCell == null) return;

				if (grdLote.Columns[grdLote.CurrentCell.ColumnIndex].Name.Equals(COL_DATA_COMPETENCIA))
				{
					e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
				}
				else if (grdLote.Columns[grdLote.CurrentCell.ColumnIndex].Name.Equals(COL_VALOR_LANCTO))
				{
					e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
				}
				else if (grdLote.Columns[grdLote.CurrentCell.ColumnIndex].Name.Equals(COL_CNPJ_CPF))
				{
					e.KeyChar = Global.filtraDigitacaoCnpjCpf(e.KeyChar);
				}
				else if (grdLote.Columns[grdLote.CurrentCell.ColumnIndex].Name.Equals(COL_NF))
				{
					e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
				}
				else if (grdLote.Columns[grdLote.CurrentCell.ColumnIndex].Name.Equals(COL_DESCRICAO))
				{
					e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
				}
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
			if ((e.Shift || e.Control) && ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Space)))
			{
				if (txtDataCompetencia.Text.Trim().Length == 0) txtDataCompetencia.Text = Global.formataDataDdMmYyyyComSeparador(DateTime.Now);
			}

			Global.trataTextBoxKeyDown(sender, e, txtValor);
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
			Global.trataTextBoxKeyDown(sender, e, grdLote);
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
			Global.trataTextBoxKeyDown(sender, e, txtDescricao);
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

		#region [ grdLote ]

		#region [ grdLote_EditingControlShowing ]
		private void grdLote_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
		{
			ComboBox cb;
			DsDataSource.DtbPlanoContasContaComboRow rowPlanoContasConta;
			System.Data.DataRowView item;
			int id_plano_contas_conta;

			cb = e.Control as ComboBox;
			if ((cb != null) && _blnComboPlanoContasPreencherAutomatico)
			{
				_blnComboPlanoContasPreencherAutomatico = false;
				if (cbPlanoContasConta.SelectedValue != null)
				{
					id_plano_contas_conta = (int)Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString());
					for (int i = 0; i < cb.Items.Count; i++)
					{
						item = (System.Data.DataRowView)cb.Items[i];
						rowPlanoContasConta = (DsDataSource.DtbPlanoContasContaComboRow)item.Row;
						if (!rowPlanoContasConta.IsidNull())
						{
							if (rowPlanoContasConta.id == id_plano_contas_conta)
							{
								cb.SelectedItem = cb.Items[i];
								grdLote.EndEdit();
								grdLote.focusNextEditableCell();
								break;
							}
						}
					}
				}
			}
		}
		#endregion

		#region [ grdLote_CellValidating ]
		private void grdLote_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
		{
			#region [ Há conteúdo na célula? ]
			if (e == null) return;
			if ((e.RowIndex < 0) || (e.ColumnIndex < 0)) return;
			if (e.FormattedValue == null) return;
			if (e.FormattedValue.ToString().Trim().Length == 0) return;
			#endregion

			#region [ Consistência dos dados da célula ]
			if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_DATA_COMPETENCIA))
			{
				if (!Global.isDataOk(e.FormattedValue.ToString().Trim()))
				{
					e.Cancel = true;
					avisoErro("Data inválida (" + e.FormattedValue.ToString().Trim() + ")!!");
					return;
				}
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_VALOR_LANCTO))
			{
				if (Global.formataMoedaDigitada(e.FormattedValue.ToString().Trim()).Length == 0)
				{
					e.Cancel = true;
					avisoErro("Valor inválido (" + e.FormattedValue.ToString().Trim() + ")!!");
					return;
				}
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_CNPJ_CPF))
			{
				if (!Global.isCnpjCpfOk(e.FormattedValue.ToString().Trim()))
				{
					e.Cancel = true;
					avisoErro("CNPJ/CPF inválido (" + e.FormattedValue.ToString().Trim() + ")!!");
					return;
				}
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_NF))
			{
				if ((int)Global.converteInteiro(Global.digitos(e.FormattedValue.ToString().Trim())) < 0)
				{
					e.Cancel = true;
					avisoErro("Número de NF inválido!!");
					return;
				}
			}
			#endregion
		}
		#endregion

		#region [ grdLote_CellValueChanged ]
		private void grdLote_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			#region [ Declarações ]
			int numNF;
			#endregion

			#region [ Há conteúdo na célula? ]
			if (e == null) return;
			if ((e.RowIndex < 0) || (e.ColumnIndex < 0)) return;
			if (grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
			{
				if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_VALOR_LANCTO)) atualizaValorTotalLancamentos();
				return;
			}
			#endregion

			#region [ Coluna Plano de Contas é combo-box, não faz formatação ]
			if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_PLANO_CONTAS_CONTA)) return;
			#endregion

			#region [ Tem conteúdo? ]
			grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim();
			if (grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim().Length == 0)
			{
				if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_VALOR_LANCTO)) atualizaValorTotalLancamentos();
				return;
			}
			#endregion

			#region [ Formata o conteúdo da célula ]
			if (grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null) return;

			if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_DATA_COMPETENCIA))
			{
				grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Global.formataDataDigitadaParaDDMMYYYYComSeparador(grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_VALOR_LANCTO))
			{
				grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Global.formataMoedaDigitada(grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
				atualizaValorTotalLancamentos();
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_CNPJ_CPF))
			{
				grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Global.formataCnpjCpf(grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_NF))
			{
				numNF = (int)Global.converteInteiro(Global.digitos(grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()));
				grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (numNF == 0 ? "" : Global.formataInteiro(numNF));
			}
			else if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_DESCRICAO))
			{
				grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Global.filtraTexto(grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
			}
			#endregion
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
