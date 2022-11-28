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
	public delegate void FluxoEditaLancamentoLoteAlteradoEventHandler();
	#endregion

	public partial class FFluxoEditaLote : Financeiro.FModelo
	{
		#region [ Eventos Customizados ]
		public event FluxoEditaLancamentoLoteAlteradoEventHandler evtFluxoEditaLancamentoLoteAlterado;
		#endregion

		#region [ Constantes ]
		const String COL_NATUREZA = "colNatureza";
		const String COL_ST_SEM_EFEITO = "colStSemEfeito";
		const String COL_CONFIRMACAO_PENDENTE = "colConfirmacaoPendente";
		const String COL_CONTA_CORRENTE = "colContaCorrente";
		const String COL_PLANO_CONTAS_CONTA = "colPlanoContasConta";
		const String COL_DATA_COMPETENCIA = "colDataCompetencia";
		const String COL_VALOR_LANCTO = "colValorLancto";
		const String COL_CNPJ_CPF = "colCnpjCpf";
		const String COL_DESCRICAO = "colDescricao";
		const String COL_ID_LANCTO = "colIdLancto";
        const String COL_COMP2 = "colComp2";
		#endregion

		#region [ Atributos ]
		private bool _InicializacaoOk;
		private bool _blnLancamentoEditadoFoiGravado = false;
		private List<int> _listaIdLancamentoSelecionado;

		private List<LancamentoFluxoCaixa> _listaLancamentoSelecionado = new List<LancamentoFluxoCaixa>();

		ToolStripMenuItem menuLancamento;
		ToolStripMenuItem menuLancamentoAtualizar;
		#endregion

		#region [ Construtor ]
		public FFluxoEditaLote(List<int> listaIdLancamentoSelecionado)
		{
			InitializeComponent();

			_listaIdLancamentoSelecionado = listaIdLancamentoSelecionado;

			#region [ Menu Lançamento ]
			// Menu principal de Lançamento
			menuLancamento = new ToolStripMenuItem("&Lançamento");
			menuLancamento.Name = "menuLancamento";
			// Atualizar
			menuLancamentoAtualizar = new ToolStripMenuItem("&Atualizar", null, menuLancamentoAtualizar_Click);
			menuLancamentoAtualizar.Name = "menuLancamentoAtualizar";
			menuLancamento.DropDownItems.Add(menuLancamentoAtualizar);
			// Adiciona o menu Lançamento ao menu principal
			menuPrincipal.Items.Insert(1, menuLancamento);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			// O 1ª item do combo é em branco com o valor da constante Global.Cte.Etc.FLAG_NAO_SETADO
			cbStSemEfeito.SelectedIndex = 0;
			cbStConfirmacaoPendente.SelectedIndex = 0;
			cbCtrlPagtoStatus.SelectedIndex = 0;
			txtDataCompetencia.Text = "";
            txtComp2.Text = "";
			lblQtdeLancamentos.Text = "";
			lblValorTotal.Text = "";
			cbContaCorrente.SelectedIndex = 0;
			cbPlanoContasEmpresa.SelectedIndex = 0;
			cbPlanoContasConta.SelectedIndex = 0;
		}
		#endregion

		#region [ isCampoEdicaoPreenchido ]
		/// <summary>
		/// Verifica se há algum campo de edição preenchido, lembrando que os campos vazios não 
		/// serão usados p/ efetuar as alterações e os campos preenchidos serão usados para todos 
		/// os lançamentos, da lista, ou seja, o mesmo valor será aplicado em todos os registros.
		/// </summary>
		/// <returns>
		/// true: há campo de edição preenchido
		/// false: não há nenhum campo de edição preenchido
		/// </returns>
		private bool isCampoEdicaoPreenchido()
		{
			#region [ Declarações ]
			bool blnHaAlteracao = false;
			#endregion

			if (cbStSemEfeito.SelectedIndex > -1)
			{
				if ((byte)Global.converteInteiro(cbStSemEfeito.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					blnHaAlteracao = true;
				}
			}

			if (cbStConfirmacaoPendente.SelectedIndex > -1)
			{
				if ((byte)Global.converteInteiro(cbStConfirmacaoPendente.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					blnHaAlteracao = true;
				}
			}

			if (cbCtrlPagtoStatus.SelectedIndex > -1)
			{
				if ((byte)Global.converteInteiro(cbCtrlPagtoStatus.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					blnHaAlteracao = true;
				}
			}

			if (txtDataCompetencia.Text.Trim().Length > 0)
			{
				if (Global.isDataOk(txtDataCompetencia.Text.Trim())) blnHaAlteracao = true;
			}

            if (txtComp2.Text.Trim().Length > 0)
            {
                if (Global.isDataMMYYYYOk(txtComp2.Text.Trim())) blnHaAlteracao = true;
            }

			if (cbContaCorrente.SelectedIndex > -1)
			{
				if ((byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					blnHaAlteracao = true;
				}
			}

			if (cbPlanoContasEmpresa.SelectedIndex > -1)
			{
				if ((byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					blnHaAlteracao = true;
				}
			}

			if (cbPlanoContasConta.SelectedIndex > -1)
			{
				if ((int)Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					blnHaAlteracao = true;
				}
			}

			return blnHaAlteracao;
		}
		#endregion

		#region [ houveEdicaoLancamento ]
		private bool houveEdicaoLancamento()
		{
			#region [ Declarações ]
			bool blnHaAlteracao = false;
			String strGridDescricao;
			#endregion

			#region [ Verifica se algum lançamento teve o campo 'Descrição' alterado ]
			for (int i = 0; i < _listaLancamentoSelecionado.Count; i++)
			{
				strGridDescricao = grdLote.Rows[i].Cells[COL_DESCRICAO].Value == null ? "" : grdLote.Rows[i].Cells[COL_DESCRICAO].Value.ToString().Trim();
				if (!_listaLancamentoSelecionado[i].descricao.Trim().Equals(strGridDescricao))
				{
					blnHaAlteracao = true;
					break;
				}
			}
			#endregion

			return blnHaAlteracao;
		}
		#endregion

		#region [ consisteCtrlPagtoStatus ]
		/// <summary>
		/// Para todos os lançamentos, realiza a consistência do campo de status 
		/// t_FIN_FLUXO_CAIXA.ctrl_pagto_status para que sempre mantenha a coerência 
		/// com relação ao campo t_FIN_FLUXO_CAIXA.st_sem_efeito
		/// </summary>
		/// <returns>
		/// true: consistência ok
		/// false: preenchimento está inconsistente
		/// </returns>
		private bool consisteCtrlPagtoStatus()
		{
			#region [ Declarações ]
			String strMsgErroConsistencia = "";
			byte byteStSemEfeito;
			byte byteCtrlPagtoStatus;
			#endregion

			byteCtrlPagtoStatus = (byte)Global.converteInteiro(cbCtrlPagtoStatus.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			byteStSemEfeito = (byte)Global.converteInteiro(cbStSemEfeito.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);

			if ((byteCtrlPagtoStatus == Global.Cte.Etc.FLAG_NAO_SETADO) && (byteStSemEfeito == Global.Cte.Etc.FLAG_NAO_SETADO))
			{
				return true;
			}
			else if ((byteCtrlPagtoStatus != Global.Cte.Etc.FLAG_NAO_SETADO) && (byteStSemEfeito != Global.Cte.Etc.FLAG_NAO_SETADO))
			{
				if (!consisteCtrlPagtoStatus((Global.Cte.FIN.eCtrlPagtoStatus)byteCtrlPagtoStatus, byteStSemEfeito, ref strMsgErroConsistencia))
				{
					avisoErro(strMsgErroConsistencia);
					return false;
				}
				return true;
			}
			else if ((byteCtrlPagtoStatus != Global.Cte.Etc.FLAG_NAO_SETADO) && (byteStSemEfeito == Global.Cte.Etc.FLAG_NAO_SETADO))
			{
				for (int i = 0; i < _listaLancamentoSelecionado.Count; i++)
				{
					if (!consisteCtrlPagtoStatus((Global.Cte.FIN.eCtrlPagtoStatus)byteCtrlPagtoStatus, _listaLancamentoSelecionado[i].st_sem_efeito, ref strMsgErroConsistencia))
					{
						strMsgErroConsistencia = "Erro no lançamento da linha " + (i + 1).ToString() + "\n" + strMsgErroConsistencia;
						avisoErro(strMsgErroConsistencia);
						return false;
					}
				}
				return true;
			}
			else if ((byteCtrlPagtoStatus == Global.Cte.Etc.FLAG_NAO_SETADO) && (byteStSemEfeito != Global.Cte.Etc.FLAG_NAO_SETADO))
			{
				for (int i = 0; i < _listaLancamentoSelecionado.Count; i++)
				{
					if (!consisteCtrlPagtoStatus((Global.Cte.FIN.eCtrlPagtoStatus)_listaLancamentoSelecionado[i].ctrl_pagto_status, byteStSemEfeito, ref strMsgErroConsistencia))
					{
						strMsgErroConsistencia = "Erro no lançamento da linha " + (i + 1).ToString() + "\n" + strMsgErroConsistencia;
						avisoErro(strMsgErroConsistencia);
						return false;
					}
				}
				return true;
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
		private bool consisteCtrlPagtoStatus(Global.Cte.FIN.eCtrlPagtoStatus enumCtrlPagtoStatus, byte byteStSemEfeito, ref String strMsgErroConsistencia)
		{
			#region [ Declarações ]
			bool blnErroConsistencia = false;
			#endregion

			strMsgErroConsistencia = "";

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

			if (blnErroConsistencia) return false;

			return true;
		}
		#endregion

		#region [ consisteLancamentos ]
		private bool consisteLancamentos()
		{
			#region [ Declarações ]
			String strDescricao;
			#endregion

			for (int i = 0; i < grdLote.Rows.Count; i++)
			{
				strDescricao = "";
				if (grdLote.Rows[i].Cells[COL_DESCRICAO].Value != null)
				{
					strDescricao = grdLote.Rows[i].Cells[COL_DESCRICAO].Value.ToString().Trim();
				}
				if (strDescricao.Length == 0)
				{
					avisoErro("O lançamento da linha " + (i + 1).ToString() + " está com a descrição em branco!");
					return false;
				}
			}

			return true;
		}
		#endregion

		#region [ trataBotaoAtualizar ]
		void trataBotaoAtualizar()
		{
			#region [ Declarações ]
			byte byteStSemEfeito;
			byte byteStConfirmacaoPendente;
			byte byteCtrlPagtoStatus;
			byte byteContaCorrente;
			byte bytePlanoContasEmpresa;
			DateTime dtCompetencia = DateTime.MinValue;
            DateTime dtComp2 = DateTime.MinValue;
            DateTime dtComp2Aux = DateTime.MinValue;
			String strIdLancto;
			int intIdLancto;
			int intPlanoContasGrupo;
			int intPlanoContasConta;
			bool blnResultado;
			bool blnSucesso = false;
			String strDescricaoLancto;
			String strMsgErro = "";
			String strDescricaoLog = "";
			String strMsgErroLog = "";
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
			if (!isCampoEdicaoPreenchido() && !houveEdicaoLancamento())
			{
				avisoErro("Não há alterações a serem feitas!!");
				return;
			}
			if (!consisteCtrlPagtoStatus()) return;
			if (!consisteLancamentos()) return;
			#endregion

			#region [ Confirmação simples ]
			if (!confirma("Confirma a gravação das alterações?")) return;
			#endregion

			byteStSemEfeito = (byte)Global.converteInteiro(cbStSemEfeito.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			byteStConfirmacaoPendente = (byte)Global.converteInteiro(cbStConfirmacaoPendente.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			byteCtrlPagtoStatus = (byte)Global.converteInteiro(cbCtrlPagtoStatus.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);

			if (Global.isDataOk(txtDataCompetencia.Text))
			{
				dtCompetencia = Global.converteDdMmYyyyParaDateTime(txtDataCompetencia.Text);
			}

            if (Global.isDataMMYYYYOk(txtComp2.Text))
            {
                dtComp2 = Convert.ToDateTime(txtComp2.Text);
            }

			if (cbContaCorrente.SelectedValue == null)
			{
				byteContaCorrente = Global.Cte.Etc.FLAG_NAO_SETADO;
			}
			else
			{
				byteContaCorrente = (byte)Global.converteInteiro(cbContaCorrente.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			}

			if (cbPlanoContasEmpresa.SelectedValue == null)
			{
				bytePlanoContasEmpresa = Global.Cte.Etc.FLAG_NAO_SETADO;
			}
			else
			{
				bytePlanoContasEmpresa = (byte)Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			}

			if (cbPlanoContasConta.SelectedValue == null)
			{
				intPlanoContasConta = Global.Cte.Etc.FLAG_NAO_SETADO;
			}
			else
			{
				intPlanoContasConta = (int)Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString(), Global.Cte.Etc.FLAG_NAO_SETADO);
			}

			if (intPlanoContasConta == Global.Cte.Etc.FLAG_NAO_SETADO)
			{
				intPlanoContasGrupo = Global.Cte.Etc.FLAG_NAO_SETADO;
			}
			else
			{
				// O grupo de contas é obtido a partir da conta, ou seja, não é selecionado explicitamente pelo usuário
				// Lembrando que cada conta foi vinculada a um grupo de contas no momento do cadastramento
				System.Data.DataRowView dataRowView = (System.Data.DataRowView)cbPlanoContasConta.Items[cbPlanoContasConta.SelectedIndex];
				DsDataSource.DtbPlanoContasContaComboRow rowConta = (DsDataSource.DtbPlanoContasContaComboRow)dataRowView.Row;
				intPlanoContasGrupo = (int)Global.converteInteiro(rowConta.id_plano_contas_grupo.ToString());
			}

			info(ModoExibicaoMensagemRodape.EmExecucao, "atualizando lançamentos");
			try
			{
				try
				{
					BD.iniciaTransacao();

					for (int ic = 0; ic < grdLote.Rows.Count; ic++)
					{
						#region [ Recupera campos do grid ]
						strIdLancto = grdLote.Rows[ic].Cells[COL_ID_LANCTO].Value.ToString();
						intIdLancto = (int)Global.converteInteiro(strIdLancto);
						if (intIdLancto < 0) throw new Exception("Falha ao tentar recuperar a identificação do registro do lançamento da linha " + (ic + 1).ToString() + "!!");
						strDescricaoLancto = grdLote.Rows[ic].Cells[COL_DESCRICAO].Value.ToString().Trim();
                        #endregion

                        #region [ Altera os dados do lançamento ]
                        if (grdLote.Rows[ic].Cells[COL_NATUREZA].Value.ToString().ToUpper().Equals("D"))
                            dtComp2Aux = dtComp2;
                        else
                            dtComp2Aux = DateTime.MinValue;

						blnResultado = LancamentoFluxoCaixaDAO.alteraPorEdicaoEmLote(Global.Usuario.usuario,
																					 intIdLancto,
																					 byteStSemEfeito,
																					 byteStConfirmacaoPendente,
																					 byteCtrlPagtoStatus,
																					 dtCompetencia,
                                                                                     dtComp2Aux,
																					 byteContaCorrente,
																					 bytePlanoContasEmpresa,
																					 intPlanoContasGrupo,
																					 intPlanoContasConta,
																					 strDescricaoLancto,
																					 ref strDescricaoLog,
																					 ref strMsgErro);
						#endregion

						if (blnResultado)
						{
							#region [ Grava log no BD ]
							finLog.usuario = Global.Usuario.usuario;
							finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_LOTE;
							finLog.natureza = _listaLancamentoSelecionado[ic].natureza;
							finLog.tipo_cadastro = _listaLancamentoSelecionado[ic].tipo_cadastro;
							finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
							finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
							finLog.id_registro_origem = _listaLancamentoSelecionado[ic].id;
							finLog.id_conta_corrente = (byteContaCorrente != Global.Cte.Etc.FLAG_NAO_SETADO ? byteContaCorrente : _listaLancamentoSelecionado[ic].id_conta_corrente);
							finLog.id_plano_contas_empresa = (bytePlanoContasEmpresa != Global.Cte.Etc.FLAG_NAO_SETADO ? bytePlanoContasEmpresa : _listaLancamentoSelecionado[ic].id_plano_contas_empresa);
							finLog.id_plano_contas_grupo = (intPlanoContasGrupo != Global.Cte.Etc.FLAG_NAO_SETADO ? intPlanoContasGrupo : _listaLancamentoSelecionado[ic].id_plano_contas_grupo);
							finLog.id_plano_contas_conta = (intPlanoContasConta != Global.Cte.Etc.FLAG_NAO_SETADO ? intPlanoContasConta : _listaLancamentoSelecionado[ic].id_plano_contas_conta);
							finLog.id_cliente = _listaLancamentoSelecionado[ic].id_cliente;
							finLog.cnpj_cpf = _listaLancamentoSelecionado[ic].cnpj_cpf;
							finLog.descricao = strDescricaoLog;
							finLog.st_sem_efeito = byteStSemEfeito;
							finLog.ctrl_pagto_status = byteCtrlPagtoStatus;
							FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
							#endregion
						}
						else
						{
							throw new Exception("Falha ao gravar as alterações!!\n\n" + strMsgErro);
						}
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

				#region [ Processamento pós tentativa de gravação no BD ]
				if (blnSucesso)
				{
					#region [ Atualiza os dados do objeto que armazena os dados originais do lançamento ]
					try
					{
						if (_listaLancamentoSelecionado.Count > 0) _listaLancamentoSelecionado.Clear();
						for (int i = 0; i < _listaIdLancamentoSelecionado.Count; i++)
						{
							_listaLancamentoSelecionado.Add(LancamentoFluxoCaixaDAO.getLancamentoFluxoCaixa(_listaIdLancamentoSelecionado[i]));
						}
					}
					catch (FinanceiroException ex)
					{
						avisoErro("Falha ao obter os dados dos lançamentos selecionados!!\n\n" + ex.Message);
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
					avisoErro("Falha ao gravar as alterações!!\n\n" + strMsgErro);
				}
				#endregion
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
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form: FFluxoEditaLote ]

		#region [ FFluxoEditaLote_Load ]
		private void FFluxoEditaLote_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			decimal decValorTotal = 0;
			#endregion

			try
			{
				#region [ Combo StSemEfeito ]
				cbStSemEfeito.DataSource = Global.montaOpcaoFluxoCaixaStSemEfeito(Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO);
				cbStSemEfeito.DisplayMember = "descricao";
				cbStSemEfeito.ValueMember = "codigo";
				// O 1ª item do combo é em branco com o valor da constante Global.Cte.Etc.FLAG_NAO_SETADO
				cbStSemEfeito.SelectedIndex = 0;
				#endregion

				#region [ Combo StConfirmacaoPendente ]
				cbStConfirmacaoPendente.DataSource = Global.montaOpcaoFluxoCaixaStConfirmacaoPendente(Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO);
				cbStConfirmacaoPendente.DisplayMember = "descricao";
				cbStConfirmacaoPendente.ValueMember = "codigo";
				// O 1ª item do combo é em branco com o valor da constante Global.Cte.Etc.FLAG_NAO_SETADO
				cbStConfirmacaoPendente.SelectedIndex = 0;
				#endregion

				#region [ Combo CtrlPagtoStatus ]
				cbCtrlPagtoStatus.DataSource = Global.montaOpcaoFluxoCaixaCtrlPagtoStatus(Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO);
				cbCtrlPagtoStatus.DisplayMember = "descricao";
				cbCtrlPagtoStatus.ValueMember = "codigo";
				// O 1ª item do combo é em branco com o valor da constante Global.Cte.Etc.FLAG_NAO_SETADO
				cbCtrlPagtoStatus.SelectedIndex = 0;
				#endregion

				#region [ Combo cbContaCorrente ]
				cbContaCorrente.DataSource = ComboDAO.criaDtbContaCorrenteCombo(ComboDAO.eFiltraStAtivo.TODOS, Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO);
				cbContaCorrente.ValueMember = "id";
				cbContaCorrente.DisplayMember = "contaComDescricao";
				cbContaCorrente.SelectedIndex = 0;
				#endregion

				#region [ Combo cbPlanoContasEmpresa ]
				cbPlanoContasEmpresa.DataSource = ComboDAO.criaDtbPlanoContasEmpresaCombo(ComboDAO.eFiltraStAtivo.TODOS, Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO);
				cbPlanoContasEmpresa.ValueMember = "id";
				cbPlanoContasEmpresa.DisplayMember = "idComDescricao";
				cbPlanoContasEmpresa.SelectedIndex = 0;
				#endregion

				#region [ Combo cbPlanoContasConta ]
				cbPlanoContasConta.DataSource = ComboDAO.criaDtbPlanoContasContaCombo(ComboDAO.eFiltraNatureza.TODOS, ComboDAO.eFiltraStAtivo.TODOS, ComboDAO.eFiltraStSistema.TODOS, Global.eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO);
				cbPlanoContasConta.ValueMember = "id";
				cbPlanoContasConta.DisplayMember = "idComDescricao";
				cbPlanoContasConta.SelectedIndex = 0;
				#endregion

				#region [ Demais campos ]
				limpaCampos();
				#endregion

				#region [ Tamanho da coluna descrição do grid ]
				((DataGridViewTextBoxColumn)grdLote.Columns[COL_DESCRICAO]).MaxInputLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
				#endregion

				#region [ Obtém os dados dos lançamentos selecionados ]
				try
				{
					if (_listaLancamentoSelecionado.Count > 0) _listaLancamentoSelecionado.Clear();
					for (int i = 0; i < _listaIdLancamentoSelecionado.Count; i++)
					{
						_listaLancamentoSelecionado.Add(LancamentoFluxoCaixaDAO.getLancamentoFluxoCaixa(_listaIdLancamentoSelecionado[i]));
					}

					_blnLancamentoEditadoFoiGravado = false;
				}
				catch (FinanceiroException ex)
				{
					avisoErro("Falha ao obter os dados dos lançamentos selecionados!!\n\n" + ex.Message);
					Close();
					return;
				}
				#endregion

				#region [ Exibe os dados dos lançamentos no grid ]
				if (grdLote.Rows.Count > 0) grdLote.Rows.Clear();
				grdLote.Rows.Add(_listaLancamentoSelecionado.Count);
				for (int i = 0; i < _listaLancamentoSelecionado.Count; i++)
				{
					grdLote.Rows[i].Cells[COL_ID_LANCTO].Value = _listaLancamentoSelecionado[i].id.ToString();
					grdLote.Rows[i].Cells[COL_NATUREZA].Value = _listaLancamentoSelecionado[i].natureza;
					grdLote.Rows[i].Cells[COL_NATUREZA].Style.ForeColor = (_listaLancamentoSelecionado[i].natureza == Global.Cte.FIN.Natureza.DEBITO ? Color.Red : Color.Green);
					grdLote.Rows[i].Cells[COL_ST_SEM_EFEITO].Value = Global.retornaDescricaoFluxoCaixaStSemEfeito(_listaLancamentoSelecionado[i].st_sem_efeito);
					grdLote.Rows[i].Cells[COL_ST_SEM_EFEITO].Style.ForeColor = (_listaLancamentoSelecionado[i].st_sem_efeito == Global.Cte.FIN.StSemEfeito.FLAG_LIGADO ? Color.Red : Color.Green);
					grdLote.Rows[i].Cells[COL_CONFIRMACAO_PENDENTE].Value = Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(_listaLancamentoSelecionado[i].st_confirmacao_pendente);
					grdLote.Rows[i].Cells[COL_CONFIRMACAO_PENDENTE].Style.ForeColor = (_listaLancamentoSelecionado[i].st_confirmacao_pendente == Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO ? Color.Red : Color.Green);
					grdLote.Rows[i].Cells[COL_CONTA_CORRENTE].Value = ComumDAO.getContaCorrenteNumeroConta(_listaLancamentoSelecionado[i].id_conta_corrente);
					grdLote.Rows[i].Cells[COL_PLANO_CONTAS_CONTA].Value = _listaLancamentoSelecionado[i].id_plano_contas_conta.ToString() + " - " + ComumDAO.getPlanoContasContaDescricao(_listaLancamentoSelecionado[i].id_plano_contas_conta);
					grdLote.Rows[i].Cells[COL_DATA_COMPETENCIA].Value = Global.formataDataDdMmYyyyComSeparador(_listaLancamentoSelecionado[i].dt_competencia);
                    grdLote.Rows[i].Cells[COL_COMP2].Value = Convert.ToDateTime(_listaLancamentoSelecionado[i].dt_mes_competencia) == DateTime.MinValue ? "" : Convert.ToDateTime(_listaLancamentoSelecionado[i].dt_mes_competencia).ToString("MM/yyyy");
					grdLote.Rows[i].Cells[COL_VALOR_LANCTO].Value = Global.formataMoeda(_listaLancamentoSelecionado[i].valor);
					grdLote.Rows[i].Cells[COL_CNPJ_CPF].Value = Global.formataCnpjCpf(_listaLancamentoSelecionado[i].cnpj_cpf);
					grdLote.Rows[i].Cells[COL_DESCRICAO].Value = _listaLancamentoSelecionado[i].descricao;
					decValorTotal += _listaLancamentoSelecionado[i].valor;
				}
				#endregion

				#region [ Exibe o grid sem nenhuma célula selecionada ]
				foreach (var item in grdLote.SelectedCells)
				{
					((DataGridViewCell)item).Selected = false;
				}
				#endregion

				#region [ Exibe totais ]
				lblQtdeLancamentos.Text = _listaLancamentoSelecionado.Count.ToString();
				lblValorTotal.Text = Global.formataMoeda(decValorTotal);
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

		#region [ FFluxoEditaLote_Shown ]
		private void FFluxoEditaLote_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					btnDummy.Focus();

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

		#region [ FFluxoEditaLote_FormClosing ]
		private void FFluxoEditaLote_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Verifica se houve edição em algum campo ]
			if (!_blnLancamentoEditadoFoiGravado)
			{
				if (isCampoEdicaoPreenchido() || houveEdicaoLancamento())
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
				if (evtFluxoEditaLancamentoLoteAlterado != null) evtFluxoEditaLancamentoLoteAlterado();
			}
			#endregion
		}
		#endregion

		#region [ FFluxoEditaLote_KeyPress ]
		private void FFluxoEditaLote_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (this.ActiveControl == null) return;
			if (this.ActiveControl is DataGridViewTextBoxEditingControl)
			{
				if (grdLote.CurrentCell == null) return;

				if (grdLote.Columns[grdLote.CurrentCell.ColumnIndex].Name.Equals(COL_DESCRICAO))
				{
					e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
				}
			}
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
			Global.trataComboBoxKeyDown(sender, e, txtDataCompetencia);
		}
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
            Global.trataComboBoxKeyDown(sender, e, txtComp2);
        }
        #endregion

        #endregion

        #region [ txtComp2 ]

        #region [ txtComp2_KeyDown ]

        private void txtComp2_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, cbContaCorrente);
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

		#region [ cbContaCorrente ]
		private void cbContaCorrente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasEmpresa);
		}
		#endregion

		#region [ cbPlanoContasEmpresa ]
		private void cbPlanoContasEmpresa_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasConta);
		}
		#endregion

		#region [ cbPlanoContasConta ]
		private void cbPlanoContasConta_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, grdLote);
			if (grdLote.SelectedCells.Count == 0) grdLote.focusNextEditableCell();
		}
		#endregion

		#region [ grdLote ]

		#region [ grdLote_CellValueChanged ]
		private void grdLote_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			#region [ Declarações ]
			String strDescricao;
			#endregion

			#region [ Há conteúdo na célula? ]
			if (e == null) return;
			if ((e.RowIndex < 0) || (e.ColumnIndex < 0)) return;
			#endregion

			#region [ Campo 'Descricao' ]
			if (grdLote.Columns[e.ColumnIndex].Name.Equals(COL_DESCRICAO))
			{
				#region [ Formata o conteúdo da célula ]
				if (grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
				{
					grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Global.filtraTexto(grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
				}
				#endregion

				#region [ Se a descrição foi editada, altera a cor da célula ]
				strDescricao = "";
				if (grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null) strDescricao = grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
				if (_listaLancamentoSelecionado[e.RowIndex].descricao.Trim().Equals(strDescricao.Trim()))
				{
					grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Empty;
				}
				else
				{
					grdLote.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
				}
				#endregion
			}
			#endregion
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

		#endregion

		#endregion
	}
}
