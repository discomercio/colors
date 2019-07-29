#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;
using System.Drawing.Drawing2D;
#endregion

namespace Financeiro
{
	public partial class FFluxoRelatorioCtaCorrente : Financeiro.FModelo
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

		DataTable _dtbConsulta;
		#endregion

		#region [ Memorização dos filtros ]
		private String _filtroDataCompetenciaInicial;
		private String _filtroDataCompetenciaFinal;
		private String _filtroDescricaoContaCorrente;
		private byte _filtroIdContaCorrente;
		private DateTime _filtroDtCompetenciaInicial;
		private DateTime _filtroDtCompetenciaFinal;
		#endregion

		#region [ Controle da impressão ]
		const String NOME_FONTE_DEFAULT = "Microsoft Sans Serif";
		private int _intConsultaImpressaoIdxLinha = 0;
		private int _intConsultaImpressaoNumPagina = 0;
		private String _strConsultaImpressaoDataEmissao;
		Impressao impressao;
		Font fonteTitulo;
		Font fonteListagem;
		Font fonteListagemNegrito;
		Font fonteDataEmissao;
		Font fonteFiltros;
		Font fonteNumPagina;
		Font fonteAtual;
		Brush brushPadrao;
		Pen penTracoTitulo;
		Pen penTracoPontilhado;
		float cxInicio;
		float cxFim;
		float cyInicio;
		float cyFim;
		float cyRodapeNumPagina;
		float larguraUtil;
		float alturaUtil;
		float ixDtCompetencia;
		float wxDtCompetencia;
		float ixValorCredito;
		float wxValorCredito;
		float ixValorDebito;
		float wxValorDebito;
		float ixSaldo;
		float wxSaldo;
		float ESPACAMENTO_COLUNAS;
		decimal vlSaldoInicial = 0;
		decimal vlSaldoAcumulado = 0;
		decimal vlCredito;
		decimal vlCreditoAcumulado = 0;
		decimal vlDebito;
		decimal vlDebitoAcumulado = 0;
		DateTime dtSaldoInicial = DateTime.MinValue;
		#endregion

		#endregion

		#region [ Menus ]
		ToolStripMenuItem menuLancamento;
		ToolStripMenuItem menuLancamentoLimpar;
		ToolStripMenuItem menuLancamentoImprimir;
		ToolStripMenuItem menuLancamentoPrintPreview;
		ToolStripMenuItem menuLancamentoPrinterDialog;
		#endregion

		#region [ Construtor ]
		public FFluxoRelatorioCtaCorrente()
		{
			InitializeComponent();

			#region [ Menu Lançamento ]
			// Menu principal de Lançamento
			menuLancamento = new ToolStripMenuItem("&Lançamento");
			menuLancamento.Name = "menuLancamento";
			// Limpar
			menuLancamentoLimpar = new ToolStripMenuItem("&Limpar", null, menuLancamentoLimpar_Click);
			menuLancamentoLimpar.Name = "menuLancamentoLimpar";
			menuLancamento.DropDownItems.Add(menuLancamentoLimpar);
			// Imprimir
			menuLancamentoImprimir = new ToolStripMenuItem("&Imprimir", null, menuLancamentoImprimir_Click);
			menuLancamentoImprimir.Name = "menuLancamentoImprimir";
			menuLancamento.DropDownItems.Add(menuLancamentoImprimir);
			// Visualizar Impressão
			menuLancamentoPrintPreview = new ToolStripMenuItem("&Visualizar Impressão", null, menuLancamentoPrintPreview_Click);
			menuLancamentoPrintPreview.Name = "menuLancamentoPrintPreview";
			menuLancamento.DropDownItems.Add(menuLancamentoPrintPreview);
			// Selecionar Impressora
			menuLancamentoPrinterDialog = new ToolStripMenuItem("&Selecionar Impressora", null, menuLancamentoPrinterDialog_Click);
			menuLancamentoPrinterDialog.Name = "menuLancamentoPrinterDialog";
			menuLancamento.DropDownItems.Add(menuLancamentoPrinterDialog);
			// Adiciona o menu Lançamento ao menu principal
			menuPrincipal.Items.Insert(1, menuLancamento);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtDataCompetenciaInicial.Text = "";
			txtDataCompetenciaFinal.Text = "";
            lbContaCorrente.ClearSelected();
			txtDataCompetenciaInicial.Focus();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Declarações ]
			const int MAX_PERIODO_EM_DIAS = 90;
			DateTime dtCompetenciaInicial = DateTime.MinValue;
			DateTime dtCompetenciaFinal = DateTime.MinValue;
			#endregion

			#region [ Período da Data de Competência ]
			if (txtDataCompetenciaInicial.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataCompetenciaInicial.Text))
				{
					avisoErro("Data inválida!!");
					txtDataCompetenciaInicial.Focus();
					return false;
				}
				else dtCompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaInicial.Text);
			}
			
			if (txtDataCompetenciaFinal.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataCompetenciaFinal.Text))
				{
					avisoErro("Data inválida!!");
					txtDataCompetenciaFinal.Focus();
					return false;
				}
				else dtCompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaFinal.Text);
			}

			if ((dtCompetenciaInicial > DateTime.MinValue) && (dtCompetenciaFinal > DateTime.MinValue))
			{
				if (dtCompetenciaInicial > dtCompetenciaFinal)
				{
					avisoErro("A data final do período é anterior à data inicial!!");
					txtDataCompetenciaFinal.Focus();
					return false;
				}
			}
			#endregion

			#region [ Alguma data foi informada? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if ((dtCompetenciaInicial == DateTime.MinValue) && (dtCompetenciaFinal == DateTime.MinValue))
				{
					avisoErro("É necessário informar pelo menos uma das datas para realizar a consulta!!");
					txtDataCompetenciaInicial.Focus();
					return false;
				}
			}
			#endregion

			#region [ Período de consulta é muito amplo? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if ((dtCompetenciaInicial > DateTime.MinValue) && (dtCompetenciaFinal > DateTime.MinValue))
				{
					if ((Global.calculaTimeSpanDias(dtCompetenciaFinal - dtCompetenciaInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
					{
						if (!confirma("O período de consulta excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
					}
				}
			}
			#endregion

			// Ok!
			return true;
		}
		#endregion

		#region [ calculaSaldoCtaCorrente ]
		private bool calculaSaldoCtaCorrente(DateTime dtReferencia, byte id_conta_corrente, ref decimal vlSaldoCalculado, ref String strMsgErro)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbCtaCorrente;
			SqlDataReader drFluxo;
			DateTime dtSaldoInicialCtaCorrente;
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return false;
				}
			}
			#endregion

			strMsgErro = "";
			vlSaldoCalculado = 0;

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				dtbCtaCorrente = new DataTable();
				#endregion

				if (id_conta_corrente > 0)
				{
					if (strWhere.Length > 0) strWhere += " AND";
					strWhere += " (id = " + id_conta_corrente.ToString() + ")";
				}

				if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

				dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

				strSql = "SELECT " +
							"*" +
						" FROM t_FIN_CONTA_CORRENTE" +
						strWhere;

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(dtbCtaCorrente);
				#endregion

				for (int i = 0; i < dtbCtaCorrente.Rows.Count; i++)
				{
					dtSaldoInicialCtaCorrente = (DateTime)dtbCtaCorrente.Rows[i]["dt_saldo_inicial"];
					if (dtReferencia < dtSaldoInicialCtaCorrente)
					{
						strMsgErro = "Não é possível calcular o saldo do dia " + Global.formataDataDdMmYyyyComSeparador(dtReferencia) + " porque é uma data anterior ao do saldo inicial (" + Global.formataDataDdMmYyyyComSeparador(dtSaldoInicialCtaCorrente) + ") da conta corrente: " + dtbCtaCorrente.Rows[i]["descricao"] + "!!";
						return false;
					}

					vlSaldoCalculado += (decimal)dtbCtaCorrente.Rows[i]["vl_saldo_inicial"];
					
					strSql = "SELECT" +
								" natureza," +
								" Coalesce(Sum(vl_total),0) AS vl_total" +
							" FROM " +
							"(" +
								"SELECT" +
									" natureza," +
									" Coalesce(Sum(valor),0) AS vl_total" +
								" FROM t_FIN_FLUXO_CAIXA" +
								" WHERE" +
									" (id_conta_corrente = " + dtbCtaCorrente.Rows[i]["id"].ToString() + ")" +
									" AND (st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")" +
									" AND (dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime((DateTime)dtbCtaCorrente.Rows[i]["dt_saldo_inicial"]) + ")" +
									" AND (dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferencia) + ")" +
									" AND" +
									" (" +
										"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
										" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO.ToString() + ")" +
									  ")" +
								" GROUP BY" +
									" natureza" +
								" UNION ALL " +
								"SELECT" +
										" natureza," +
										" Coalesce(Sum(valor),0) AS vl_total" +
									" FROM t_FIN_FLUXO_CAIXA" +
									" WHERE" +
										" (id_conta_corrente = " + dtbCtaCorrente.Rows[i]["id"].ToString() + ")" +
										" AND (st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")" +
										" AND (dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime((DateTime)dtbCtaCorrente.Rows[i]["dt_saldo_inicial"]) + ")" +
										" AND (dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferencia) + ")" +
										" AND" +
										" (" +
											"(dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
										  ")" +
									" GROUP BY" +
										" natureza" +
							") t" +
							" GROUP BY"+
								" natureza";
					cmCommand.CommandText = strSql;
					drFluxo = cmCommand.ExecuteReader();
					try
					{
						while (drFluxo.Read())
						{
							if (drFluxo["natureza"].ToString().Equals(Global.Cte.FIN.Natureza.CREDITO.ToString()))
							{
								vlSaldoCalculado += (decimal)drFluxo["vl_total"];
							}
							else if (drFluxo["natureza"].ToString().Equals(Global.Cte.FIN.Natureza.DEBITO.ToString()))
							{
								vlSaldoCalculado -= (decimal)drFluxo["vl_total"];
							}
						}
					}
					finally
					{
						drFluxo.Close();
					}
				}

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ montaClausulaWhere ]
		private String montaClausulaWhere()
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux = "";
			
			#region [ Data de competência ]
			_filtroDataCompetenciaInicial = txtDataCompetenciaInicial.Text;
			_filtroDataCompetenciaFinal = txtDataCompetenciaFinal.Text;
			_filtroDtCompetenciaInicial = DateTime.MinValue;
			_filtroDtCompetenciaFinal = DateTime.MinValue;

			if ((txtDataCompetenciaInicial.Text.Length > 0) && (txtDataCompetenciaFinal.Text.Length > 0))
			{
				_filtroDtCompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaInicial.Text);
				_filtroDtCompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaFinal.Text);
				// A data inicial é igual à data final?
				if (txtDataCompetenciaInicial.Text.Equals(txtDataCompetenciaFinal.Text))
				{
					strAux = " (dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((dt_competencia >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaInicial.Text) + ") AND (dt_competencia <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataCompetenciaInicial.Text.Length > 0) || (txtDataCompetenciaFinal.Text.Length > 0))
			{
				if (txtDataCompetenciaInicial.Text.Length > 0)
				{
					_filtroDtCompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaInicial.Text);
					_filtroDtCompetenciaFinal = _filtroDtCompetenciaInicial;
					strAux = " (dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaInicial.Text) + ")";
				}
				else if (txtDataCompetenciaFinal.Text.Length > 0)
				{
					_filtroDtCompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaFinal.Text);
					_filtroDtCompetenciaInicial = _filtroDtCompetenciaFinal;
					strAux = " (dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Conta Corrente ]
			_filtroIdContaCorrente = 0;
            _filtroDescricaoContaCorrente = "";
            strAux = "";
            if (lbContaCorrente.SelectedItems.Count > 0)
            {
                foreach (DataRowView item in lbContaCorrente.SelectedItems)
                {
                    if (!_filtroDescricaoContaCorrente.Equals("")) _filtroDescricaoContaCorrente += ", ";
                    _filtroDescricaoContaCorrente += item["contaComDescricao"];

                    if (strAux != "") strAux += " OR";
                    strAux += " (id_conta_corrente = " + item["id"].ToString() + ")";
                }
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append("(" + strAux + ")");
            }
            #endregion

            return sbWhere.ToString();
		}
		#endregion

		#region [ montaSqlConsulta ]
		private String montaSqlConsulta()
		{
			#region [ Declarações ]
			String strWhere;
			String strSql;
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			#region [ Monta cláusula Where ]
			strWhere = montaClausulaWhere();
			#endregion

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

			#region [ Monta Select ]
			// Datas posteriores à data de crédito do último arquivo de retorno: considerar todos os 
			//		lançamentos previstos válidos (st_sem_efeito=0)
			// Datas anteriores à data de crédito do último arquivo de retorno: considerar apenas os 
			//		lançamentos realizados e válidos (st_sem_efeito=0 e st_confirmacao_pendente=0)
			strSql = "SELECT" +
						" Coalesce(tCred.dt_competencia, tDeb.dt_competencia) AS dt_competencia," +
						" Coalesce(tCred.vl_credito,0) AS vl_credito," +
						" Coalesce(tDeb.vl_debito,0) AS vl_debito" +
					" FROM " +
							"(" +
							"SELECT" +
								" dt_competencia," +
								" Coalesce(Sum(valor),0) AS vl_credito" +
							" FROM" +
								" t_FIN_FLUXO_CAIXA" +
							" WHERE" +
								" (st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")" +
								" AND (natureza='" + Global.Cte.FIN.Natureza.CREDITO + "')" +
								" AND (" +
										"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
										" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO.ToString() + ")" +
									  ")" +
								(strWhere.Length > 0 ? " AND" : "") +
								strWhere +
							" GROUP BY" +
								" dt_competencia" +
							" UNION " +
							"SELECT" +
								" dt_competencia," +
								" Coalesce(Sum(valor),0) AS vl_credito" +
							" FROM" +
								" t_FIN_FLUXO_CAIXA" +
							" WHERE" +
								" (st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")" +
								" AND (natureza='" + Global.Cte.FIN.Natureza.CREDITO + "')" +
								" AND (" +
										"(dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
									  ")" +
								(strWhere.Length > 0 ? " AND" : "") +
								strWhere +
							" GROUP BY" +
								" dt_competencia" +
							") tCred" +
						" FULL OUTER JOIN " +
							"(" +
							"SELECT" +
								" dt_competencia," +
								" Coalesce(Sum(valor),0) AS vl_debito" +
							" FROM" +
								" t_FIN_FLUXO_CAIXA" +
							" WHERE" +
								" (st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")" +
								" AND (natureza='" + Global.Cte.FIN.Natureza.DEBITO + "')" +
								" AND (" +
										"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
										" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO.ToString() + ")" +
									  ")" +
								(strWhere.Length > 0 ? " AND" : "") +
								strWhere +
							" GROUP BY" +
								" dt_competencia" +
							" UNION " +
							"SELECT" +
								" dt_competencia," +
								" Coalesce(Sum(valor),0) AS vl_debito" +
							" FROM" +
								" t_FIN_FLUXO_CAIXA" +
							" WHERE" +
								" (st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")" +
								" AND (natureza='" + Global.Cte.FIN.Natureza.DEBITO + "')" +
								" AND (" +
										"(dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
									  ")" +
								(strWhere.Length > 0 ? " AND" : "") +
								strWhere +
							" GROUP BY" +
								" dt_competencia" +
							") tDeb" +
						" ON (tCred.dt_competencia=tDeb.dt_competencia)" +
					" ORDER BY" +
						" dt_competencia ASC";
			#endregion

			return strSql;
		}
		#endregion

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			#endregion

			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return false;
				}
			}
			#endregion

			try
			{
				#region [ Consistência dos parâmetros ]
				btnDummy.Focus();
				if (!consisteCampos()) return false;
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				_dtbConsulta = new DataTable();
				#endregion

				#region [ Monta o SQL da consulta ]
				strSql = montaSqlConsulta();
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(_dtbConsulta);
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return false;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ printPreview ]
		private void printPreview()
		{
			if (!executaPesquisa()) return;

			prnPreviewConsulta.WindowState = FormWindowState.Maximized;
			prnPreviewConsulta.MinimizeBox = true;
			prnPreviewConsulta.Text = Global.Cte.Aplicativo.M_ID + " - Visualização da Impressão";
			prnPreviewConsulta.PrintPreviewControl.Zoom = 1;
			prnPreviewConsulta.PrintPreviewControl.AutoZoom = true;
			prnPreviewConsulta.FormBorderStyle = FormBorderStyle.Sizable;
			prnPreviewConsulta.ShowDialog();
		}
		#endregion

		#region [ printerDialog ]
		private void printerDialog()
		{
			prnDialogConsulta.ShowDialog();
		}
		#endregion

		#region [ imprimeConsulta ]
		private void imprimeConsulta()
		{
			if (!executaPesquisa()) return;

			prnDocConsulta.Print();
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FFluxoRelatorioCtaCorrente ]

		#region [ FFluxoRelatorioCtaCorrente_Load ]
		private void FFluxoRelatorioCtaCorrente_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCampos();

				#region [ Combo Conta Corrente ]
				// Cria uma linha com a opção Todas
				DsDataSource.DtbContaCorrenteComboDataTable dtbContaCorrente = new DsDataSource.DtbContaCorrenteComboDataTable();
				DsDataSource.DtbContaCorrenteComboRow rowContaCorrente = dtbContaCorrente.NewDtbContaCorrenteComboRow();
				//rowContaCorrente.contaComDescricao = "Todas";
				//rowContaCorrente.id = 0;
				//dtbContaCorrente.AddDtbContaCorrenteComboRow(rowContaCorrente);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbContaCorrente.Merge(ComboDAO.criaDtbContaCorrenteCombo(ComboDAO.eFiltraStAtivo.TODOS));

                lbContaCorrente.DataSource = dtbContaCorrente;
                lbContaCorrente.ValueMember = "id";
                lbContaCorrente.DisplayMember = "contaComDescricao";
                lbContaCorrente.ClearSelected();
				#endregion

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

		#region [ FFluxoRelatorioCtaCorrente_Shown ]
		private void FFluxoRelatorioCtaCorrente_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					txtDataCompetenciaInicial.Focus();
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

		#region [ FFluxoRelatorioCtaCorrente_FormClosing ]
		private void FFluxoRelatorioCtaCorrente_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FFluxoRelatorioCtaCorrente_KeyDown ]
		private void FFluxoRelatorioCtaCorrente_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				e.SuppressKeyPress = true;
				printPreview();
				return;
			}
		}
		#endregion

		#endregion

		#region [ txtDataCompetenciaInicial ]

		#region [ txtDataCompetenciaInicial_Enter ]
		private void txtDataCompetenciaInicial_Enter(object sender, EventArgs e)
		{
			txtDataCompetenciaInicial.Select(0, txtDataCompetenciaInicial.Text.Length);
		}
		#endregion

		#region [ txtDataCompetenciaInicial_Leave ]
		private void txtDataCompetenciaInicial_Leave(object sender, EventArgs e)
		{
			if (txtDataCompetenciaInicial.Text.Length == 0) return;
			txtDataCompetenciaInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCompetenciaInicial.Text);
			if (!Global.isDataOk(txtDataCompetenciaInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCompetenciaInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataCompetenciaInicial_KeyDown ]
		private void txtDataCompetenciaInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDataCompetenciaFinal);
		}
		#endregion

		#region [ txtDataCompetenciaInicial_KeyPress ]
		private void txtDataCompetenciaInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDataCompetenciaFinal ]

		#region [ txtDataCompetenciaFinal_Enter ]
		private void txtDataCompetenciaFinal_Enter(object sender, EventArgs e)
		{
			txtDataCompetenciaFinal.Select(0, txtDataCompetenciaFinal.Text.Length);
		}
		#endregion

		#region [ txtDataCompetenciaFinal_Leave ]
		private void txtDataCompetenciaFinal_Leave(object sender, EventArgs e)
		{
			if (txtDataCompetenciaFinal.Text.Length == 0) return;
			txtDataCompetenciaFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCompetenciaFinal.Text);
			if (!Global.isDataOk(txtDataCompetenciaFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCompetenciaFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataCompetenciaFinal_KeyDown ]
		private void txtDataCompetenciaFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, lbContaCorrente);
		}
		#endregion

		#region [ txtDataCompetenciaFinal_KeyPress ]
		private void txtDataCompetenciaFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ cbContaCorrente ]

		#region [ cbContaCorrente_KeyDown ]
		private void cbContaCorrente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, btnDummy);
		}
		#endregion

		#endregion

		#region [ Limpar ]

		#region [ btnLimpar_Click ]
		private void btnLimpar_Click(object sender, EventArgs e)
		{
			limpaCampos();
		}
		#endregion

		#region [ menuLancamentoLimpar_Click ]
		private void menuLancamentoLimpar_Click(object sender, EventArgs e)
		{
			limpaCampos();
		}
		#endregion

		#endregion

		#region [ Imprimir ]

		#region [ btnImprimir_Click ]
		private void btnImprimir_Click(object sender, EventArgs e)
		{
			imprimeConsulta();
		}
		#endregion

		#region [ menuLancamentoImprimir_Click ]
		private void menuLancamentoImprimir_Click(object sender, EventArgs e)
		{
			imprimeConsulta();
		}
		#endregion

		#endregion

		#region [ Print Preview ]

		#region [ btnPrintPreview_Click ]
		private void btnPrintPreview_Click(object sender, EventArgs e)
		{
			printPreview();
		}
		#endregion

		#region [ menuLancamentoPrintPreview_Click ]
		private void menuLancamentoPrintPreview_Click(object sender, EventArgs e)
		{
			printPreview();
		}
		#endregion

		#endregion

		#region [ PrinterDialog ]

		#region [ btnPrinterDialog_Click ]
		private void btnPrinterDialog_Click(object sender, EventArgs e)
		{
			printerDialog();
		}
		#endregion

		#region [ menuLancamentoPrinterDialog_Click ]
		private void menuLancamentoPrinterDialog_Click(object sender, EventArgs e)
		{
			printerDialog();
		}
		#endregion

		#endregion

		#endregion

		#region [ Impressão ]

		#region [ prnDocConsulta_BeginPrint ]
		private void prnDocConsulta_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
		{
			String strMsgErro = "";
            decimal saldoAux = 0;
            byte idContaCorrente;

			if (_dtbConsulta == null)
			{
				e.Cancel = true;
				return;
			}

			_intConsultaImpressaoIdxLinha = 0;
			_intConsultaImpressaoNumPagina = 0;
			_strConsultaImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);

			impressao = new Impressao();

			#region [ Prepara elementos de impressão ]
			fonteTitulo = new Font(NOME_FONTE_DEFAULT, 18, FontStyle.Bold);
			fonteListagem = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Regular);
			fonteListagemNegrito = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Bold);
			fonteDataEmissao = new Font(NOME_FONTE_DEFAULT, 9f, FontStyle.Regular);
			fonteFiltros = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Italic);
			fonteNumPagina = new Font(NOME_FONTE_DEFAULT, 10f, FontStyle.Bold);
			brushPadrao = new SolidBrush(Color.Black);
			penTracoTitulo = new Pen(brushPadrao, .5f);
			penTracoPontilhado = Impressao.criaPenTracoPontilhado();
			#endregion

			#region [ Calcula o saldo inicial ]
			dtSaldoInicial = _filtroDtCompetenciaInicial.AddDays(-1);
            if (lbContaCorrente.SelectedItems.Count > 0)
            {
                foreach (DataRowView item in lbContaCorrente.SelectedItems)
                {
                    idContaCorrente = Convert.ToByte(item["id"]);
                    if (!calculaSaldoCtaCorrente(dtSaldoInicial, idContaCorrente, ref vlSaldoInicial, ref strMsgErro))
                    {
                        avisoErro(strMsgErro);
                        e.Cancel = true;
                        return;
                    }
                    saldoAux += vlSaldoInicial;
                }
                vlSaldoInicial = saldoAux;
            }
            else
            {
                if (!calculaSaldoCtaCorrente(dtSaldoInicial, _filtroIdContaCorrente, ref vlSaldoInicial, ref strMsgErro))
                {
                    avisoErro(strMsgErro);
                    e.Cancel = true;
                    return;
                }
            }
			#endregion

			#region [ Inicializações ]
			vlSaldoAcumulado = vlSaldoInicial;
			vlCreditoAcumulado = 0;
			vlDebitoAcumulado = 0;
			#endregion
		}
		#endregion

		#region [ prnDocConsulta_PrintPage ]
		private void prnDocConsulta_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			float cx;
			float cy;
			float hMax;
			String strTexto;
			int intLinhasImpressasNestaPagina = 0;
			#endregion

			#region [ Verifica se alguma consulta foi realizada ]
			if (_dtbConsulta == null)
			{
				e.Cancel = true;
				return;
			}
			#endregion

			#region [ Contador de página ]
			_intConsultaImpressaoNumPagina++;
			#endregion

			e.Graphics.PageUnit = GraphicsUnit.Millimeter;
			if (_intConsultaImpressaoNumPagina == 1)
			{
				#region [ Medidas do papel ]
				prnDocConsulta.DocumentName = "Relatório Sintético de Fluxo de Caixa";
				cxInicio = impressao.getLeftMarginInMm(e);
				larguraUtil = impressao.getWidthInMm(e);
				cxFim = cxInicio + larguraUtil;
				cyInicio = impressao.getTopMarginInMm(e);
				alturaUtil = impressao.getHeightInMm(e);
				cyFim = cyInicio + alturaUtil;
				cyRodapeNumPagina = cyFim - fonteNumPagina.GetHeight(e.Graphics) - 1;
				#endregion

				#region [ Layout das colunas ]
				wxDtCompetencia = 15f;
				wxValorCredito = 20f;
				wxValorDebito = 20f;
				wxSaldo = 20f;
				ESPACAMENTO_COLUNAS = (larguraUtil - wxDtCompetencia - wxValorCredito - wxValorDebito - wxSaldo) / 3;
				ixDtCompetencia = cxInicio;
				ixValorCredito = ixDtCompetencia + wxDtCompetencia + ESPACAMENTO_COLUNAS;
				ixValorDebito = ixValorCredito + wxValorCredito + ESPACAMENTO_COLUNAS;
				ixSaldo = cxInicio + larguraUtil - wxSaldo;
				#endregion
			}

			cx = cxInicio;
			cy = cyInicio;

			#region [ Título ]
			strTexto = "Relatório Sintético de Fluxo de Caixa";
			fonteAtual = fonteTitulo;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx - 1, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Data da emissão ]
			strTexto = "Emitido em: " + _strConsultaImpressaoDataEmissao;
			fonteAtual = fonteDataEmissao;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Filtros ]

			#region [ Configura fonte ]
			fonteAtual = fonteFiltros;
			#endregion

			#region [ Data de competência ]
			strTexto = "Competência: ";
			if ((_filtroDataCompetenciaInicial.Length > 0) && (_filtroDataCompetenciaFinal.Length > 0))
				strTexto += _filtroDataCompetenciaInicial + " a " + _filtroDataCompetenciaFinal;
			else if (_filtroDataCompetenciaInicial.Length > 0)
				strTexto += _filtroDataCompetenciaInicial;
			else if (_filtroDataCompetenciaFinal.Length > 0)
				strTexto += _filtroDataCompetenciaFinal;
			else strTexto += "N.I.";
			
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Conta Corrente ]
			strTexto = "Conta Corrente: ";
			if (_filtroDescricaoContaCorrente.Length > 0)
				strTexto += _filtroDescricaoContaCorrente;
			else
				strTexto += "Todas";

			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Saldo Inicial ]
			strTexto = "Saldo em " + Global.formataDataDdMmYyyyComSeparador(dtSaldoInicial) + ":   " + Global.formataMoeda(vlSaldoInicial);
			fonteAtual = fonteListagemNegrito;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Títulos ]
			cy += .5f;
			fonteAtual = fonteListagemNegrito;
			strTexto = "DATA";
			cx = ixDtCompetencia;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "CRÉDITO";
			cx = ixValorCredito + wxValorCredito - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "DÉBITO";
			cx = ixValorDebito + wxValorDebito - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "SALDO";
			cx = ixSaldo + wxSaldo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cy += fonteAtual.GetHeight(e.Graphics);
			cy += .5f;
			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
				   (_intConsultaImpressaoIdxLinha < _dtbConsulta.Rows.Count))
			{
				fonteAtual = fonteListagem;
				hMax = Math.Max(fonteListagem.GetHeight(e.Graphics), fonteListagemNegrito.GetHeight(e.Graphics));

				#region [ Data de competência ]
				cx = ixDtCompetencia;
				strTexto = Global.formataDataDdMmYyyyComSeparador((DateTime)_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["dt_competencia"]);
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				fonteAtual = fonteListagemNegrito;

				#region [ Valor Crédito ]
				vlCredito = (decimal)_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["vl_credito"];
				vlCreditoAcumulado += vlCredito;
				strTexto = Global.formataMoeda(vlCredito);
				cx = ixValorCredito + wxValorCredito - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Valor Débito ]
				vlDebito = (decimal)_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["vl_debito"];
				vlDebitoAcumulado += vlDebito;
				strTexto = Global.formataMoeda(vlDebito);
				cx = ixValorDebito + wxValorDebito - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Saldo ]
				vlSaldoAcumulado += vlCredito - vlDebito;
				strTexto = Global.formataMoeda(vlSaldoAcumulado);
				cx = ixSaldo + wxSaldo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				cy += hMax;

				intLinhasImpressasNestaPagina++;
				_intConsultaImpressaoIdxLinha++;

				#region [ Na última linha não imprime o tracejado ]
				if (_intConsultaImpressaoIdxLinha < _dtbConsulta.Rows.Count)
				{
					#region [ Traço pontilhado ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoPontilhado, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion
				}
				#endregion
			}

			#region [ Terminou a listagem? ]
			if (_intConsultaImpressaoIdxLinha < _dtbConsulta.Rows.Count)
			{
				e.HasMorePages = true;
			}
			else
			{
				e.HasMorePages = false;

				#region [ Há espaço suficiente? ]
				if ((cy + fonteListagemNegrito.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 10))
				{
					if (intLinhasImpressasNestaPagina > 0)
					{
						#region [ Traço ]
						cy += 1f;
						e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
						cy += 1f;
						#endregion
					}
					else cy += .5f;

					#region [ Imprime os totais ]
					fonteAtual = fonteListagemNegrito;
					cx = ixDtCompetencia;
					strTexto = "TOTAL";
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(vlCreditoAcumulado);
					cx = ixValorCredito + wxValorCredito - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(vlDebitoAcumulado);
					cx = ixValorDebito + wxValorDebito - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(vlSaldoAcumulado);
					cx = ixSaldo + wxSaldo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion
				}
				else e.HasMorePages = true;
				#endregion
			}
			#endregion

			#region [ Imprime nº página ]
			strTexto = _intConsultaImpressaoNumPagina.ToString();
			fonteAtual = fonteNumPagina;
			cy = cyRodapeNumPagina;
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion
		}
		#endregion

		#endregion
	}
}

