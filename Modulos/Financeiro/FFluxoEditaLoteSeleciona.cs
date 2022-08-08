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
	public partial class FFluxoEditaLoteSeleciona : Financeiro.FModelo
	{
		#region [ Constantes ]
		const String GRID_ST_SEM_EFEITO__CANCELADO = "Cancel";
		const String GRID_ST_SEM_EFEITO__VALIDO = "Válido";
		const String GRID_COL_CHECK_BOX = "colCheckBox";
		#endregion

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

		private bool _blnEventoCallBackEmProcessamento = false;
		FFluxoEditaLote fFluxoEditaLote;
		#endregion

		#region [ Menu ]
		ToolStripMenuItem menuLancamento;
		ToolStripMenuItem menuLancamentoPesquisar;
		ToolStripMenuItem menuLancamentoEditar;
		ToolStripMenuItem menuLancamentoLimpar;
		ToolStripMenuItem menuLancamentoImprimir;
		ToolStripMenuItem menuLancamentoPrintPreview;
		ToolStripMenuItem menuLancamentoPrinterDialog;
		#endregion

		#region [ Memorização dos filtros ]
		private String _filtroDataCompetenciaInicial;
		private String _filtroDataCompetenciaFinal;
        private String _filtroMesCompetenciaInicial;
        private String _filtroMesCompetenciaFinal;
		private String _filtroDataCadastroInicial;
		private String _filtroDataCadastroFinal;
		private String _filtroNatureza;
		private String _filtroStSemEfeito;
		private String _filtroAtrasados;
		private String _filtroCtrlPagtoStatus;
		private String _filtroValor;
		private String _filtroNomeCliente;
		private String _filtroCnpjCpf;
		private String _filtroNF;
		private String _filtroDescricao;
		private String _filtroContaCorrente;
		private String _filtroPlanoContasEmpresa;
		private String _filtroPlanoContasGrupo;
		private String _filtroPlanoContasConta;
		#endregion

		#region [ Controle da impressão ]
		private int _intConsultaImpressaoIdxLinhaGrid = 0;
		private int _intConsultaImpressaoNumPagina = 0;
		private String _strConsultaImpressaoDataEmissao;
		const String NOME_FONTE_DEFAULT = "Microsoft Sans Serif";
		const float ESPACAMENTO_COLUNAS = 4.0f;
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
        float ixComp2;
        float wxComp2;
		float ixContaCorrente;
		float wxContaCorrente;
		float ixPlanoContasConta;
		float wxPlanoContasConta;
		float ixDescricao;
		float wxDescricao;
		float ixValor;
		float wxValor;
		float ixNomeCnpjCpf;
		float wxNomeCnpjCpf;
		float ixObs;
		float wxObs;
		Impressao impressao;
		#endregion

		#endregion

		#region [ Construtor ]
		public FFluxoEditaLoteSeleciona()
		{
			InitializeComponent();

			#region [ Menu Lançamento ]
			// Menu principal de Lançamento
			menuLancamento = new ToolStripMenuItem("&Lançamento");
			menuLancamento.Name = "menuLancamento";
			// Pesquisar
			menuLancamentoPesquisar = new ToolStripMenuItem("&Pesquisar", null, menuLancamentoPesquisar_Click);
			menuLancamentoPesquisar.Name = "menuLancamentoPesquisar";
			menuLancamento.DropDownItems.Add(menuLancamentoPesquisar);
			// Limpar
			menuLancamentoLimpar = new ToolStripMenuItem("&Limpar", null, menuLancamentoLimpar_Click);
			menuLancamentoLimpar.Name = "menuLancamentoLimpar";
			menuLancamento.DropDownItems.Add(menuLancamentoLimpar);
			// Editar
			menuLancamentoEditar = new ToolStripMenuItem("&Editar", null, menuLancamentoEditar_Click);
			menuLancamentoEditar.Name = "menuLancamentoEditar";
			menuLancamento.DropDownItems.Add(menuLancamentoEditar);
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
            txtMesCompetenciaInicial.Text = "";
            txtMesCompetenciaFinal.Text = "";
			txtDataCadastroInicial.Text = "";
			txtDataCadastroFinal.Text = "";
			cbNatureza.SelectedIndex = -1;
			cbStSemEfeito.SelectedIndex = -1;
			cbAtrasados.SelectedIndex = -1;
			cbCtrlPagtoStatus.SelectedIndex = -1;
			txtValor.Text = "";
			txtNomeCliente.Text = "";
			txtCnpjCpf.Text = "";
			txtNF.Text = "";
			txtDescricao.Text = "";
			cbContaCorrente.SelectedIndex = -1;
			cbPlanoContasEmpresa.SelectedIndex = -1;
			cbPlanoContasGrupo.SelectedIndex = -1;
			cbPlanoContasConta.SelectedIndex = -1;
			lblTotalizacaoRegistros.Text = "";
			lblTotalizacaoValor.Text = "";
			gridDados.DataSource = null;
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
			DateTime dtCadastroInicial = DateTime.MinValue;
			DateTime dtCadastroFinal = DateTime.MinValue;
			DateTime dtAuxInicial = DateTime.MinValue;
			DateTime dtAuxFinal = DateTime.MinValue;
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

			#region [ Período da Data de Cadastramento ]
			if (txtDataCadastroInicial.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataCadastroInicial.Text))
				{
					avisoErro("Data inválida!!");
					txtDataCadastroInicial.Focus();
					return false;
				}
				else dtCadastroInicial = Global.converteDdMmYyyyParaDateTime(txtDataCadastroInicial.Text);
			}

			if (txtDataCadastroFinal.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataCadastroFinal.Text))
				{
					avisoErro("Data inválida!!");
					txtDataCadastroFinal.Focus();
					return false;
				}
				else dtCadastroFinal = Global.converteDdMmYyyyParaDateTime(txtDataCadastroFinal.Text);
			}

			if ((dtCadastroInicial > DateTime.MinValue) && (dtCadastroFinal > DateTime.MinValue))
			{
				if (dtCadastroInicial > dtCadastroFinal)
				{
					avisoErro("A data final do período é anterior à data inicial!!");
					txtDataCadastroFinal.Focus();
					return false;
				}
			}
			#endregion
			
			#region [ CNPJ/CPF ]
			if (txtCnpjCpf.Text.Trim().Length > 0)
			{
				if (!Global.isCnpjCpfOk(txtCnpjCpf.Text))
				{
					avisoErro("CNPJ/CPF inválido!!");
					txtCnpjCpf.Focus();
					return false;
				}
			}
			#endregion

			#region [ NF ]
			if (txtNF.Text.Trim().Length > 0)
			{
				if ((int)Global.converteInteiro(Global.digitos(txtNF.Text.Trim())) < 0)
				{
					avisoErro("Número de NF inválido!!");
					txtNF.Focus();
					return false;
				}
			}
			#endregion

			#region [ Alguma data foi informada? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if ((dtCompetenciaInicial == DateTime.MinValue) && (dtCompetenciaFinal == DateTime.MinValue) &&
					(dtCadastroInicial == DateTime.MinValue) && (dtCadastroFinal == DateTime.MinValue))
				{
					avisoErro("É necessário informar pelo menos uma das datas para realizar a consulta!!");
					txtDataCompetenciaInicial.Focus();
					return false;
				}
			}
			#endregion

			#region [ Período de consulta é muito amplo? ]
			if (!_blnEventoCallBackEmProcessamento)
			{
				if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
				{
					if ((dtCompetenciaInicial > DateTime.MinValue) && (dtCompetenciaFinal > DateTime.MinValue) &&
						(dtCadastroInicial > DateTime.MinValue) && (dtCadastroFinal > DateTime.MinValue))
					{
						if (dtCompetenciaInicial > dtCadastroInicial) dtAuxInicial = dtCompetenciaInicial; else dtAuxInicial = dtCadastroInicial;
						if (dtCompetenciaFinal < dtCadastroFinal) dtAuxFinal = dtCompetenciaFinal; else dtAuxFinal = dtCadastroFinal;
						if ((Global.calculaTimeSpanDias(dtAuxFinal - dtAuxInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
						{
							if (!confirma("O período de consulta excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
						}
					}
					else if ((dtCompetenciaInicial > DateTime.MinValue) && (dtCompetenciaFinal > DateTime.MinValue))
					{
						if ((Global.calculaTimeSpanDias(dtCompetenciaFinal - dtCompetenciaInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
						{
							if (!confirma("O período de consulta excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
						}
					}
					else if ((dtCadastroInicial > DateTime.MinValue) && (dtCadastroFinal > DateTime.MinValue))
					{
						if ((Global.calculaTimeSpanDias(dtCadastroFinal - dtCadastroInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
						{
							if (!confirma("O período de consulta excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
						}
					}
				}
			}
			#endregion

			// Ok!
			return true;
		}
		#endregion

		#region [ montaClausulaWhere ]
		private String montaClausulaWhere()
		{
			#region [ Declarações ]
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;
			int numNF;
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

			#region [ Data de competência ]
			if ((txtDataCompetenciaInicial.Text.Length > 0) && (txtDataCompetenciaFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtDataCompetenciaInicial.Text.Equals(txtDataCompetenciaFinal.Text))
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tFC.dt_competencia >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaInicial.Text) + ") AND (tFC.dt_competencia <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataCompetenciaInicial.Text.Length > 0) || (txtDataCompetenciaFinal.Text.Length > 0))
			{
				if (txtDataCompetenciaInicial.Text.Length > 0)
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaInicial.Text) + ")";
				}
				else if (txtDataCompetenciaFinal.Text.Length > 0)
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCompetenciaFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
            #endregion

            #region [ Mês de competência ]
            if ((txtMesCompetenciaInicial.Text.Length > 0) && (txtMesCompetenciaFinal.Text.Length > 0))
            {
                // O período inicial é igual ao período final?
                if (txtMesCompetenciaInicial.Text.Equals(txtMesCompetenciaFinal.Text))
                {
                    strAux = " (tFC.dt_mes_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime((Convert.ToDateTime(txtMesCompetenciaInicial.Text)).ToString("dd/MM/yyyy")) + ")";
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(strAux);
                }
                else
                {
                    strAux = " ((tFC.dt_mes_competencia >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime((Convert.ToDateTime(txtMesCompetenciaInicial.Text)).ToString("dd/MM/yyyy")) + ") AND (tFC.dt_mes_competencia <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime((Convert.ToDateTime(txtMesCompetenciaFinal.Text)).ToString("dd/MM/yyyy")) + "))";
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(strAux);
                }
            }
            else if ((txtMesCompetenciaInicial.Text.Length > 0) || (txtMesCompetenciaFinal.Text.Length > 0))
            {
                if (txtMesCompetenciaInicial.Text.Length > 0)
                {
                    strAux = " (tFC.dt_mes_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime((Convert.ToDateTime(txtMesCompetenciaInicial.Text)).ToString("dd/MM/yyyy")) + ")";
                }
                else if (txtMesCompetenciaFinal.Text.Length > 0)
                {
                    strAux = " (tFC.dt_mes_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime((Convert.ToDateTime(txtMesCompetenciaFinal.Text)).ToString("dd/MM/yyyy")) + ")";
                }
                else strAux = "";

                if (strAux.Length > 0)
                {
                    if (sbWhere.Length > 0) sbWhere.Append(" AND");
                    sbWhere.Append(strAux);
                }
            }
            #endregion

            #region [ Data de cadastramento ]
            if ((txtDataCadastroInicial.Text.Length > 0) && (txtDataCadastroFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtDataCadastroInicial.Text.Equals(txtDataCadastroFinal.Text))
				{
					strAux = " (tFC.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCadastroInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tFC.dt_cadastro >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCadastroInicial.Text) + ") AND (tFC.dt_cadastro <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCadastroFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataCadastroInicial.Text.Length > 0) || (txtDataCadastroFinal.Text.Length > 0))
			{
				if (txtDataCadastroInicial.Text.Length > 0)
				{
					strAux = " (tFC.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCadastroInicial.Text) + ")";
				}
				else if (txtDataCadastroFinal.Text.Length > 0)
				{
					strAux = " (tFC.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCadastroFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Natureza ]
			if ((cbNatureza.SelectedIndex > -1) && (cbNatureza.SelectedValue.ToString().Trim().Length > 0))
			{
				strAux = " (tFC.natureza = '" + (char)cbNatureza.SelectedValue + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Sem Efeito ]
			if (cbStSemEfeito.SelectedIndex > -1)
			{
				if (((byte)cbStSemEfeito.SelectedValue) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					strAux = " (tFC.st_sem_efeito = " + (byte)cbStSemEfeito.SelectedValue + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Atrasados ]
			if (cbAtrasados.SelectedIndex > -1)
			{
				strAux = " (" +
							"(tFC.st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO.ToString() + ")" +
							" AND (tFC.st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO.ToString() + ")" +
							" AND (tFC.dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
							" AND (tFC.ctrl_pagto_status <> " + ((byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_PAGO_CHEQUE_VINCULADO).ToString() + ")" +
						  ")";
				
				if (((byte)cbAtrasados.SelectedValue) == Global.Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.APENAS_ATRASADOS)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else if (((byte)cbAtrasados.SelectedValue) == Global.Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.IGNORAR_ATRASADOS)
				{
					strAux = " (NOT " + strAux + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ CtrlPagtoStatus ]
			if (cbCtrlPagtoStatus.SelectedIndex > -1)
			{
				if (((byte)cbCtrlPagtoStatus.SelectedValue) != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					strAux = " (tFC.ctrl_pagto_status = " + (byte)cbCtrlPagtoStatus.SelectedValue + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region[ Valor ]
			if (txtValor.Text.Trim().Length > 0)
			{
				if (Global.converteNumeroDecimal(txtValor.Text) > 0)
				{
					strAux = Global.sqlFormataDecimal(Global.converteNumeroDecimal(txtValor.Text));
					strAux = " (tFC.valor = " + strAux + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Nome do cliente ]
			if (txtNomeCliente.Text.Trim().Length > 0)
			{
				strAux = " (tC.nome LIKE '" + txtNomeCliente.Text + BD.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ CNPJ/CPF ]
			if (Global.digitos(txtCnpjCpf.Text).Length > 0)
			{
				strAux = " (tFC.cnpj_cpf = '" + Global.digitos(txtCnpjCpf.Text) + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ NF ]
			if (txtNF.Text.Trim().Length > 0)
			{
				numNF = (int)Global.converteInteiro(Global.digitos(txtNF.Text.Trim()));
				if (numNF > 0)
				{
					strAux = " (tFC.numero_NF = " + numNF.ToString() + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Descrição ]
			if (txtDescricao.Text.Trim().Length > 0)
			{
				strAux = " (tFC.descricao LIKE '" + BD.CARACTER_CURINGA_TODOS + txtDescricao.Text + BD.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Conta Corrente ]
			if ((cbContaCorrente.SelectedIndex > -1) && ((Global.converteInteiro(cbContaCorrente.SelectedValue.ToString())) > 0))
			{
				strAux = " (tFC.id_conta_corrente = " + cbContaCorrente.SelectedValue.ToString() + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Plano de Contas Empresa ]
			if ((cbPlanoContasEmpresa.SelectedIndex > -1) && ((Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString())) > 0))
			{
				strAux = " (tFC.id_plano_contas_empresa = " + cbPlanoContasEmpresa.SelectedValue.ToString() + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Plano Contas Grupo ]
			if ((cbPlanoContasGrupo.SelectedIndex > -1) && ((Global.converteInteiro(cbPlanoContasGrupo.SelectedValue.ToString())) > 0))
			{
				strAux = " (tFC.id_plano_contas_grupo = " + cbPlanoContasGrupo.SelectedValue.ToString() + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Plano Contas Conta ]
			if ((cbPlanoContasConta.SelectedIndex > -1) && ((Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString())) > 0))
			{
				// Obtém o código da natureza desta conta
				System.Data.DataRowView dataRowView = (System.Data.DataRowView)cbPlanoContasConta.Items[cbPlanoContasConta.SelectedIndex];
				DsDataSource.DtbPlanoContasContaComboRow rowConta = (DsDataSource.DtbPlanoContasContaComboRow)dataRowView.Row;
				// Monta SQL
				strAux = " ((tFC.id_plano_contas_conta = " + cbPlanoContasConta.SelectedValue.ToString() + ")" +
						 " AND (tFC.natureza = '" + rowConta.natureza + "'))";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
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

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

			#region [ Monta cláusula Where ]
			strWhere = montaClausulaWhere();
			if (strWhere.Length > 0) strWhere = " WHERE " + strWhere;
			#endregion

			#region [ Monta Select ]
			strSql = "SELECT " +
						" tFC.id," +
						" tFC.id_conta_corrente," +
						" tFC.id_plano_contas_empresa," +
						" tFC.id_plano_contas_grupo," +
						" tFC.id_plano_contas_conta," +
						" tFC.natureza," +
						" tFC.st_sem_efeito," +
						" tFC.st_confirmacao_pendente," +
						" tFC.dt_competencia," +
                        " tFC.dt_mes_competencia," +
                        " tFC.valor," +
						" Lower(tFC.descricao) AS descricao," +
						" tFC.ctrl_pagto_id_parcela," +
						" tFC.ctrl_pagto_modulo," +
						" tFC.ctrl_pagto_status," +
						" tFC.id_cliente," +
						" Coalesce(tFC.cnpj_cpf,'') AS cnpj_cpf," +
						" tFC.tipo_cadastro," +
						" tFC.editado_manual," +
						" tFC.dt_cadastro," +
						" tFC.dt_hr_cadastro," +
						" tFC.usuario_cadastro," +
						" tFC.dt_ult_atualizacao," +
						" tFC.dt_hr_ult_atualizacao," +
						" tFC.usuario_ult_atualizacao," +
						" Coalesce(tCC.conta,'') AS descricao_conta_corrente," +
						" Coalesce(tPCE.descricao,'') AS descricao_plano_contas_empresa," +
						" Coalesce(tPCG.descricao,'') AS descricao_plano_contas_grupo," +
						Global.sqlMontaPadLeftCampoNumerico("tPCC.id", '0', Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA) + " + ' - ' + Coalesce(tPCC.descricao,'') AS descricao_plano_contas_conta," +
						" Coalesce(tC.nome,'') AS nome_cliente," +
						" tFC.st_boleto_pago_cheque," +
						" tFC.dt_ocorrencia_banco_boleto_pago_cheque," +
						Global.sqlMontaCaseWhenParaFluxoCaixaAguardandoLiquidacao("tFC", "st_aguardando_liquidacao") + "," +
						Global.sqlMontaCaseWhenParaFluxoCaixaEmAtraso(dtReferenciaLimitePagamentoEmAtraso, "tFC", "st_em_atraso") + "," +
						Global.sqlMontaCaseWhenParaFluxoCaixaCalculaDiasEmAtraso(dtReferenciaLimitePagamentoEmAtraso, "tFC", "qtde_dias_em_atraso") +
					" FROM t_FIN_FLUXO_CAIXA tFC" +
						" LEFT JOIN t_FIN_CONTA_CORRENTE tCC" +
							" ON (tFC.id_conta_corrente=tCC.id)" +
						" LEFT JOIN t_FIN_PLANO_CONTAS_EMPRESA tPCE" +
							" ON (tFC.id_plano_contas_empresa=tPCE.id)" +
						" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
							" ON (tFC.id_plano_contas_grupo=tPCG.id)" +
						" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
							" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
						" LEFT JOIN t_CLIENTE tC" +
							" ON (tFC.cnpj_cpf=tC.cnpj_cpf)" +
					strWhere +
					" ORDER BY" +
						" st_em_atraso," +
						" st_aguardando_liquidacao," +
						" qtde_dias_em_atraso," +
						" tFC.dt_competencia," +
						" tFC.natureza," +
						" tFC.valor," +
						" tFC.descricao," +
						" tFC.id";
			#endregion

			return strSql;
		}
		#endregion

		#region [ memorizaFiltrosParaImpressao ]
		/// <summary>
		/// Memoriza os parâmetros usados na última pesquisa para serem usados na impressão.
		/// </summary>
		private void memorizaFiltrosParaImpressao()
		{
			_filtroDataCompetenciaInicial = txtDataCompetenciaInicial.Text;
			_filtroDataCompetenciaFinal = txtDataCompetenciaFinal.Text;
            _filtroMesCompetenciaInicial = txtMesCompetenciaInicial.Text;
            _filtroMesCompetenciaFinal = txtMesCompetenciaFinal.Text;
			_filtroDataCadastroInicial = txtDataCadastroInicial.Text;
			_filtroDataCadastroFinal = txtDataCadastroFinal.Text;
			_filtroNatureza = cbNatureza.Text;
			_filtroStSemEfeito = cbStSemEfeito.Text;
			_filtroAtrasados = cbAtrasados.Text;
			_filtroCtrlPagtoStatus = cbCtrlPagtoStatus.Text;
			_filtroValor = txtValor.Text;
			_filtroNomeCliente = txtNomeCliente.Text;
			_filtroCnpjCpf = txtCnpjCpf.Text;
			_filtroNF = txtNF.Text;
			_filtroDescricao = txtDescricao.Text;
			_filtroContaCorrente = cbContaCorrente.Text;
			_filtroPlanoContasEmpresa = cbPlanoContasEmpresa.Text;
			_filtroPlanoContasGrupo = cbPlanoContasGrupo.Text;
			_filtroPlanoContasConta = cbPlanoContasConta.Text;
		}
		#endregion

		#region [ executaPesquisa ]
		private void executaPesquisa()
		{
			#region [ Declarações ]
			Decimal decTotalizacaoValor = 0;
			int intQtdeRegistros = 0;
			String strSql;
			String strNomeCnpjCpf;
			DateTime dtOcorrenciaBanco;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DsDataSource.DtbFinFluxoCaixaGridDataTable dtbConsulta = new DsDataSource.DtbFinFluxoCaixaGridDataTable();
			DsDataSource.DtbFinFluxoCaixaGridRow rowConsulta;
			#endregion

			try
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

				#region [ Consistência dos parâmetros ]
				btnDummy.Focus();
				if (!consisteCampos()) return;
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados (processamento no servidor e transferência de dados)");

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				#endregion

				#region [ Monta o SQL da consulta ]
				strSql = montaSqlConsulta();
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Prepara alguns campos que necessitam de formatação ]
				for (int i = 0; i < dtbConsulta.Count; i++)
				{
					rowConsulta = (DsDataSource.DtbFinFluxoCaixaGridRow)dtbConsulta.Rows[i];
					rowConsulta.descricao_natureza = Global.retornaDescricaoFluxoCaixaNatureza(rowConsulta.natureza);
					if (rowConsulta.st_sem_efeito == Global.Cte.FIN.StSemEfeito.FLAG_LIGADO)
						rowConsulta.descricao_st_sem_efeito = GRID_ST_SEM_EFEITO__CANCELADO;
					else
						rowConsulta.descricao_st_sem_efeito = GRID_ST_SEM_EFEITO__VALIDO;
					
					rowConsulta.dt_competencia_formatada = Global.formataDataDdMmYyyyComSeparador(rowConsulta.dt_competencia);
					rowConsulta.valor_formatado = Global.formataMoeda(rowConsulta.valor);
					rowConsulta.cnpj_cpf_formatado = Global.formataCnpjCpf(rowConsulta.cnpj_cpf);
					rowConsulta.dt_cadastro_formatada = Global.formataDataDdMmYyyyComSeparador(rowConsulta.dt_cadastro);
					rowConsulta.dt_hr_cadastro_formatada = Global.formataDataDdMmYyyyComSeparador(rowConsulta.dt_hr_cadastro);
					rowConsulta.dt_ult_atualizacao_formatada = Global.formataDataDdMmYyyyComSeparador(rowConsulta.dt_ult_atualizacao);
					rowConsulta.dt_hr_ult_atualizacao_formatada = Global.formataDataDdMmYyyyComSeparador(rowConsulta.dt_hr_ult_atualizacao);
					
					strNomeCnpjCpf = BD.readToString(rowConsulta.nome_cliente);
					if (strNomeCnpjCpf.Length > 0)
						strNomeCnpjCpf += " (" + rowConsulta.cnpj_cpf_formatado + ")";
					else
						strNomeCnpjCpf = rowConsulta.cnpj_cpf_formatado;

					rowConsulta.nome_cnpj_cpf = strNomeCnpjCpf;

					#region [ Boleto pago? ]
					if ((rowConsulta.ctrl_pagto_modulo == Global.Cte.FIN.CtrlPagtoModulo.BOLETO) &&
						(rowConsulta.ctrl_pagto_status == (byte)Global.Cte.FIN.eCtrlPagtoStatus.PAGO))
					{
						if (rowConsulta.IsobservacoesNull()) rowConsulta.observacoes = "";
						if (rowConsulta.observacoes.Length > 0) rowConsulta.observacoes += "\n";
						rowConsulta.observacoes += "Boleto pago";
					}
					#endregion

					#region [ Está em atraso? ]
					if (rowConsulta.st_em_atraso == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						if (rowConsulta.IsobservacoesNull()) rowConsulta.observacoes = "";
						if (rowConsulta.observacoes.Length > 0) rowConsulta.observacoes += "\n";
						rowConsulta.observacoes += "Atrasado há " + rowConsulta.qtde_dias_em_atraso.ToString();
						if (rowConsulta.qtde_dias_em_atraso == 1)
							rowConsulta.observacoes += " dia";
						else
							rowConsulta.observacoes += " dias";
					}
					#endregion

					#region [ Aguardando liquidação ]
					if (rowConsulta.st_aguardando_liquidacao == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						if (rowConsulta.Isdt_ocorrencia_banco_boleto_pago_chequeNull())
							dtOcorrenciaBanco = DateTime.MinValue;
						else
							dtOcorrenciaBanco = rowConsulta.dt_ocorrencia_banco_boleto_pago_cheque;

						if (rowConsulta.IsobservacoesNull()) rowConsulta.observacoes = "";
						if (rowConsulta.observacoes.Length > 0) rowConsulta.observacoes += "\n";
						rowConsulta.observacoes += "Aguardando liquidação (cheque em " + Global.formataDataDdMmYyyyComSeparador(dtOcorrenciaBanco) + ")";
					}
					#endregion

					decTotalizacaoValor += rowConsulta.valor;
					intQtdeRegistros++;
				}
				#endregion

				#region [ Exibição dos dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição (processamento local)");
					
					gridDados.SuspendLayout();

					#region [ Carrega os dados no Grid ]
					gridDados.DataSource = dtbConsulta;
					#endregion

					#region [ Formata alguns campos no Grid ]
					for (int i = 0; i < gridDados.Rows.Count; i++)
					{
                        #region [ Comp2 ]
                        gridDados.Rows[i].Cells["dt_mes_competencia"].Value = gridDados.Rows[i].Cells["dt_mes_competencia"].Value != DBNull.Value ? Convert.ToDateTime(gridDados.Rows[i].Cells["dt_mes_competencia"].Value).ToString("MM/yyyy") : "";
                        #endregion

                        #region [ Natureza ]
                        if (gridDados.Rows[i].Cells["natureza"].Value.ToString().Equals(Global.Cte.FIN.Natureza.DEBITO.ToString()))
						{
							gridDados.Rows[i].Cells["natureza"].Style.ForeColor = Color.Red;
						}
						else if (gridDados.Rows[i].Cells["natureza"].Value.ToString().Equals(Global.Cte.FIN.Natureza.CREDITO.ToString()))
						{
							gridDados.Rows[i].Cells["natureza"].Style.ForeColor = Color.Green;
						}
						#endregion

						#region [ StSemEfeito ]
						if (gridDados.Rows[i].Cells["descricao_st_sem_efeito"].Value.ToString().Equals(GRID_ST_SEM_EFEITO__CANCELADO))
						{
							gridDados.Rows[i].Cells["descricao_st_sem_efeito"].Style.ForeColor = Color.Red;
						}
						else if (gridDados.Rows[i].Cells["descricao_st_sem_efeito"].Value.ToString().Equals(GRID_ST_SEM_EFEITO__VALIDO))
						{
							gridDados.Rows[i].Cells["descricao_st_sem_efeito"].Style.ForeColor = Color.Green;
						}
						#endregion

						#region [ Em atraso? ]
						if (gridDados.Rows[i].Cells["st_em_atraso"].Value.ToString().Equals(Global.Cte.FIN.StCampoFlag.FLAG_LIGADO.ToString()))
						{
							gridDados.Rows[i].Cells["obs"].Style.ForeColor = Color.Red;
						}
						#endregion

						#region [ Aguardando liquidação? ]
						if (gridDados.Rows[i].Cells["st_aguardando_liquidacao"].Value.ToString().Equals(Global.Cte.FIN.StCampoFlag.FLAG_LIGADO.ToString()))
						{
							gridDados.Rows[i].Cells["obs"].Style.ForeColor = Color.Red;
						}
						#endregion
					}
					#endregion

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < gridDados.Rows.Count; i++)
					{
						if (gridDados.Rows[i].Selected) gridDados.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					gridDados.ResumeLayout();
				}
				#endregion

				#region [ Exibe totalização ]
				lblTotalizacaoValor.Text = Global.formataMoeda(decTotalizacaoValor);
				lblTotalizacaoRegistros.Text = intQtdeRegistros.ToString();
				#endregion

				memorizaFiltrosParaImpressao();

				gridDados.Focus();

				// Feedback da conclusão da pesquisa
				if (!_blnEventoCallBackEmProcessamento) SystemSounds.Exclamation.Play();
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
			}
		}
		#endregion

		#region [ editaLancamentosSelecionadosEmLote ]
		private void editaLancamentosSelecionadosEmLote()
		{
			#region [ Declarações ]
			List<int> listaIdLancamentoSelecionado = new List<int>();
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

			info(ModoExibicaoMensagemRodape.EmExecucao, "preparando edição do lançamento");
			try
			{
				for (int i = 0; i < gridDados.Rows.Count; i++)
				{
					if (gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value != null)
					{
						if ((bool)gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value)
						{
							listaIdLancamentoSelecionado.Add((int)Global.converteInteiro(gridDados.Rows[i].Cells["id"].Value.ToString()));
						}
					}
					
				}

				if (listaIdLancamentoSelecionado.Count == 0)
				{
					avisoErro("Nenhum lançamento foi selecionado!!");
					return;
				}

				fFluxoEditaLote = new FFluxoEditaLote(listaIdLancamentoSelecionado);
				fFluxoEditaLote.evtFluxoEditaLancamentoLoteAlterado += new FluxoEditaLancamentoLoteAlteradoEventHandler(TrataEventoFluxoEditaLancamentoLoteAlterado);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}

			fFluxoEditaLote.ShowDialog();
		}
		#endregion

		#region [ printPreview ]
		private void printPreview()
		{
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
			prnDocConsulta.Print();
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FFluxoConsulta ]

		#region [ FFluxoConsulta_Load ]
		private void FFluxoConsulta_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCampos();

				#region [ Combo Conta Corrente ]
				// Cria uma linha com a opção Todas
				DsDataSource.DtbContaCorrenteComboDataTable dtbContaCorrente = new DsDataSource.DtbContaCorrenteComboDataTable();
				DsDataSource.DtbContaCorrenteComboRow rowContaCorrente = dtbContaCorrente.NewDtbContaCorrenteComboRow();
				rowContaCorrente.contaComDescricao = "Todas";
				rowContaCorrente.id = 0;
				dtbContaCorrente.AddDtbContaCorrenteComboRow(rowContaCorrente);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbContaCorrente.Merge(ComboDAO.criaDtbContaCorrenteCombo(ComboDAO.eFiltraStAtivo.TODOS));
				cbContaCorrente.DataSource = dtbContaCorrente;
				cbContaCorrente.ValueMember = "id";
				cbContaCorrente.DisplayMember = "contaComDescricao";
				cbContaCorrente.SelectedIndex = -1;
				#endregion

				#region [ Combo Plano Contas Empresa ]
				// Cria uma linha com a opção Todas
				DsDataSource.DtbPlanoContasEmpresaComboDataTable dtbPlanoContasEmpresa = new DsDataSource.DtbPlanoContasEmpresaComboDataTable();
				DsDataSource.DtbPlanoContasEmpresaComboRow rowPlanoContasEmpresa = dtbPlanoContasEmpresa.NewDtbPlanoContasEmpresaComboRow();
				rowPlanoContasEmpresa.id = 0;
				rowPlanoContasEmpresa.idComDescricao = "Todas";
				dtbPlanoContasEmpresa.AddDtbPlanoContasEmpresaComboRow(rowPlanoContasEmpresa);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbPlanoContasEmpresa.Merge(ComboDAO.criaDtbPlanoContasEmpresaCombo(ComboDAO.eFiltraStAtivo.TODOS));
				cbPlanoContasEmpresa.DataSource = dtbPlanoContasEmpresa;
				cbPlanoContasEmpresa.ValueMember = "id";
				cbPlanoContasEmpresa.DisplayMember = "idComDescricao";
				cbPlanoContasEmpresa.SelectedIndex = -1;
				#endregion

				#region [ Combo Plano Contas Grupo ]
				// Cria uma linha com a opção Todas
				DsDataSource.DtbPlanoContasGrupoComboDataTable dtbPlanoContasGrupo = new DsDataSource.DtbPlanoContasGrupoComboDataTable();
				DsDataSource.DtbPlanoContasGrupoComboRow rowPlanoContasGrupo = dtbPlanoContasGrupo.NewDtbPlanoContasGrupoComboRow();
				rowPlanoContasGrupo.id = 0;
				rowPlanoContasGrupo.idComDescricao = "Todos";
				dtbPlanoContasGrupo.AddDtbPlanoContasGrupoComboRow(rowPlanoContasGrupo);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbPlanoContasGrupo.Merge(ComboDAO.criaDtbPlanoContasGrupoCombo(ComboDAO.eFiltraStAtivo.TODOS));
				cbPlanoContasGrupo.DataSource = dtbPlanoContasGrupo;
				cbPlanoContasGrupo.ValueMember = "id";
				cbPlanoContasGrupo.DisplayMember = "idComDescricao";
				cbPlanoContasGrupo.SelectedIndex = -1;
				#endregion

				#region [ Combo Plano Contas Conta ]
				// Cria uma linha com a opção Todas
				DsDataSource.DtbPlanoContasContaComboDataTable dtbPlanoContasConta = new DsDataSource.DtbPlanoContasContaComboDataTable();
				DsDataSource.DtbPlanoContasContaComboRow rowPlanoContasConta = dtbPlanoContasConta.NewDtbPlanoContasContaComboRow();
				rowPlanoContasConta.id = 0;
				rowPlanoContasConta.idComDescricao = "Todas";
				dtbPlanoContasConta.AddDtbPlanoContasContaComboRow(rowPlanoContasConta);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbPlanoContasConta.Merge(ComboDAO.criaDtbPlanoContasContaCombo(ComboDAO.eFiltraNatureza.TODOS, ComboDAO.eFiltraStAtivo.TODOS, ComboDAO.eFiltraStSistema.TODOS));
				cbPlanoContasConta.DataSource = dtbPlanoContasConta;
				cbPlanoContasConta.ValueMember = "id";
				cbPlanoContasConta.DisplayMember = "idComDescricao";
				cbPlanoContasConta.SelectedIndex = -1;
				#endregion

				#region [ Combo Natureza ]
				cbNatureza.DataSource = Global.montaOpcaoFluxoCaixaNatureza(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbNatureza.DisplayMember = "descricao";
				cbNatureza.ValueMember = "codigo";
				cbNatureza.SelectedIndex = -1;
				#endregion

				#region [ Combo StSemEfeito ]
				cbStSemEfeito.DataSource = Global.montaOpcaoFluxoCaixaStSemEfeito(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbStSemEfeito.DisplayMember = "descricao";
				cbStSemEfeito.ValueMember = "codigo";
				cbStSemEfeito.SelectedIndex = -1;
				for (int i = 0; i < cbStSemEfeito.Items.Count; i++)
				{
					if (((Global.OpcaoFluxoCaixaStSemEfeito)cbStSemEfeito.Items[i]).codigo == Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)
					{
						cbStSemEfeito.SelectedIndex = i;
						break;
					}
				}
				#endregion

				#region [ Combo Atrasados ]
				cbAtrasados.DataSource = Global.montaOpcaoFluxoCaixaPesquisaLancamentoAtrasado(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbAtrasados.DisplayMember = "descricao";
				cbAtrasados.ValueMember = "codigo";
				cbAtrasados.SelectedIndex = -1;
				#endregion

				#region [ Combo CtrlPagtoStatus ]
				cbCtrlPagtoStatus.DataSource = Global.montaOpcaoFluxoCaixaCtrlPagtoStatus(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbCtrlPagtoStatus.DisplayMember = "descricao";
				cbCtrlPagtoStatus.ValueMember = "codigo";
				cbCtrlPagtoStatus.SelectedIndex = -1;
				#endregion

				#region [ Campo descrição ]
				txtDescricao.MaxLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
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

		#region [ FFluxoConsulta_Shown ]
		private void FFluxoConsulta_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Permissão de acesso ao módulo ]
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_EDITAR_LANCTO))
					{
						btnEditar.Enabled = false;
						menuLancamentoEditar.Enabled = false;
					}
					#endregion

					#region [ Prepara lista de auto complete do campo nome do cliente ]
					txtNomeCliente.AutoCompleteCustomSource.AddRange(FMain.fMain.listaNomeClienteAutoComplete.ToArray());
					#endregion

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

		#region [ FFluxoConsulta_FormClosing ]
		private void FFluxoConsulta_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FFluxoConsulta_KeyDown ]
		private void FFluxoConsulta_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				e.SuppressKeyPress = true;
				executaPesquisa();
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
            Global.trataTextBoxKeyDown(sender, e, txtMesCompetenciaInicial);
        }
        #endregion

        #region [ txtDataCompetenciaFinal_KeyPress ]
        private void txtDataCompetenciaFinal_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #endregion

        #region [ txtMesCompetenciaInicial ]

        #region [ txtMesCompetenciaInicial_KeyDown ]

        private void txtMesCompetenciaInicial_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, txtMesCompetenciaFinal);
        }

        #endregion

        #region [ txtMesCompetenciaInicial_KeyPress ]
        private void txtMesCompetenciaInicial_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #region [ txtMesCompetenciaInicial_Leave ]

        private void txtMesCompetenciaInicial_Leave(object sender, EventArgs e)
        {
            if (txtMesCompetenciaInicial.Text.Length == 0) return;
            txtMesCompetenciaInicial.Text = Global.formataDataDigitadaParaMMYYYYComSeparador(txtMesCompetenciaInicial.Text);
            if (!Global.isDataMMYYYYOk(txtMesCompetenciaInicial.Text))
            {
                avisoErro("Formato inválido!!");
                txtMesCompetenciaInicial.Focus();
                return;
            }
        }

        #endregion

        #endregion

        #region [ txtMesCompetenciaFinal ]

        #region [ txtMesCompetenciaFinal_KeyDown ]

        private void txtMesCompetenciaFinal_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, txtDataCadastroInicial);
        }

        #endregion

        #region [ txtMesCompetenciaFinal_KeyPress ]
        private void txtMesCompetenciaFinal_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #region [ txtMesCompetenciaFinal_Leave ]

        private void txtMesCompetenciaFinal_Leave(object sender, EventArgs e)
        {
            if (txtMesCompetenciaFinal.Text.Length == 0) return;
            txtMesCompetenciaFinal.Text = Global.formataDataDigitadaParaMMYYYYComSeparador(txtMesCompetenciaFinal.Text);
            if (!Global.isDataMMYYYYOk(txtMesCompetenciaFinal.Text))
            {
                avisoErro("Formato inválido!!");
                txtMesCompetenciaFinal.Focus();
                return;
            }
        }

        #endregion

        #endregion

        #region [ txtDataCadastroInicial ]

        #region [ txtDataCadastroInicial_Enter ]
        private void txtDataCadastroInicial_Enter(object sender, EventArgs e)
        {
            txtDataCadastroInicial.Select(0, txtDataCadastroInicial.Text.Length);
        }
        #endregion

        #region [ txtDataCadastroInicial_Leave ]
        private void txtDataCadastroInicial_Leave(object sender, EventArgs e)
        {
            if (txtDataCadastroInicial.Text.Length == 0) return;
            txtDataCadastroInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCadastroInicial.Text);
            if (!Global.isDataOk(txtDataCadastroInicial.Text))
            {
                avisoErro("Data inválida!!");
                txtDataCadastroInicial.Focus();
                return;
            }
        }
        #endregion

        #region [ txtDataCadastroInicial_KeyDown ]
        private void txtDataCadastroInicial_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, txtDataCadastroFinal);
        }
        #endregion

        #region [ txtDataCadastroInicial_KeyPress ]
        private void txtDataCadastroInicial_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #endregion

        #region [ txtDataCadastroFinal ]

        #region [ txtDataCadastroFinal_Enter ]
        private void txtDataCadastroFinal_Enter(object sender, EventArgs e)
        {
            txtDataCadastroFinal.Select(0, txtDataCadastroFinal.Text.Length);
        }
        #endregion

        #region [ txtDataCadastroFinal_Leave ]
        private void txtDataCadastroFinal_Leave(object sender, EventArgs e)
        {
            if (txtDataCadastroFinal.Text.Length == 0) return;
            txtDataCadastroFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCadastroFinal.Text);
            if (!Global.isDataOk(txtDataCadastroFinal.Text))
            {
                avisoErro("Data inválida!!");
                txtDataCadastroFinal.Focus();
                return;
            }
        }
        #endregion

        #region [ txtDataCadastroFinal_KeyDown ]
        private void txtDataCadastroFinal_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, cbStSemEfeito);
        }
        #endregion

        #region [ txtDataCadastroFinal_KeyPress ]
        private void txtDataCadastroFinal_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #endregion

        #region [ cbNatureza ]

        #region [ cbNatureza_KeyDown ]
        private void cbNatureza_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataComboBoxKeyDown(sender, e, cbAtrasados);
        }
        #endregion

        #endregion

        #region [ cbStSemEfeito ]

        #region [ cbStSemEfeito_KeyDown ]
        private void cbStSemEfeito_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataComboBoxKeyDown(sender, e, cbNatureza);
        }
        #endregion

        #endregion

        #region [ cbAtrasados ]

        #region [ cbAtrasados_KeyDown ]
        private void cbAtrasados_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataComboBoxKeyDown(sender, e, cbCtrlPagtoStatus);
        }
        #endregion

        #endregion

        #region [ cbCtrlPagtoStatus ]

        #region [ cbCtrlPagtoStatus_KeyDown ]
        private void cbCtrlPagtoStatus_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtValor);
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

		#region [ txtValor_KeyDown ]
		private void txtValor_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDescricao);
		}
		#endregion

		#region [ txtValor_KeyPress ]
		private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDescricao ]

		#region [ txtDescricao_Enter ]
		private void txtDescricao_Enter(object sender, EventArgs e)
		{
			txtDescricao.Select(0, txtDescricao.Text.Length);
		}
		#endregion

		#region [ txtDescricao_Leave ]
		private void txtDescricao_Leave(object sender, EventArgs e)
		{
			txtDescricao.Text = txtDescricao.Text.Trim();
		}
		#endregion

		#region [ txtDescricao_KeyDown ]
		private void txtDescricao_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtNomeCliente);
		}
		#endregion

		#region [ txtDescricao_KeyPress ]
		private void txtDescricao_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtNomeCliente ]

		#region [ txtNomeCliente_Enter ]
		private void txtNomeCliente_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNomeCliente_Leave ]
		private void txtNomeCliente_Leave(object sender, EventArgs e)
		{
			txtNomeCliente.Text = txtNomeCliente.Text.Trim();
		}
		#endregion

		#region [ txtNomeCliente_KeyDown ]
		private void txtNomeCliente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtCnpjCpf);
		}
		#endregion

		#region [ txtNomeCliente_KeyPress ]
		private void txtNomeCliente_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtCnpjCpf ]

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

		#region [ txtCnpjCpf_KeyDown ]
		private void txtCnpjCpf_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtNF);
		}
		#endregion

		#region [ txtCnpjCpf_KeyPress ]
		private void txtCnpjCpf_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCnpjCpf(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, cbContaCorrente);
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
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasGrupo);
		}
		#endregion

		#endregion

		#region [ cbPlanoContasGrupo ]

		#region [ cbPlanoContasGrupo_KeyDown ]
		private void cbPlanoContasGrupo_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasConta);
		}
		#endregion

		#endregion

		#region [ cbPlanoContasConta ]

		#region [ cbPlanoContasConta_KeyDown ]
		private void cbPlanoContasConta_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, btnPesquisar);
		}
		#endregion

		#endregion

		#region [ gridDados ]

		#region [ gridDados_CellContentClick ]
		private void gridDados_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e == null) return;
			if (e.ColumnIndex == 0)
			{
				DataGridViewCheckBoxCell chkBox = (DataGridViewCheckBoxCell)this.gridDados[e.ColumnIndex, e.RowIndex];
				if (chkBox.EditingCellFormattedValue.ToString().ToUpper().Equals("TRUE"))
				{
					this.gridDados.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSkyBlue;
				}
				else
				{
					this.gridDados.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Empty;
				}
			}
		}
		#endregion

		#region [ gridDados_DoubleClick ]
		private void gridDados_DoubleClick(object sender, EventArgs e)
		{
			editaLancamentosSelecionadosEmLote();
		}
		#endregion

		#endregion

		#region [ Botões / Menu ]

		#region [ Pesquisar ]

		#region [ btnPesquisar_Click ]
		private void btnPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#region [ menuLancamentoPesquisar_Click ]
		private void menuLancamentoPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#endregion

		#region [ Editar ]

		#region [ btnEditar_Click ]
		private void btnEditar_Click(object sender, EventArgs e)
		{
			editaLancamentosSelecionadosEmLote();
		}
		#endregion

		#region [ menuLancamentoEditar_Click ]
		private void menuLancamentoEditar_Click(object sender, EventArgs e)
		{
			editaLancamentosSelecionadosEmLote();
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

		#region [ Marcar todos os lançamentos ]

		#region [ btnMarcarTodos_Click ]
		private void btnMarcarTodos_Click(object sender, EventArgs e)
		{
			gridDados.SuspendLayout();
			try
			{
				for (int i = 0; i < gridDados.Rows.Count; i++)
				{
					if (gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value != null)
					{
						if (!(bool)gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value)
						{
							gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value = true;
							gridDados.Rows[i].DefaultCellStyle.BackColor = Color.LightSkyBlue;
						}
					}
					else
					{
						gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value = true;
						gridDados.Rows[i].DefaultCellStyle.BackColor = Color.LightSkyBlue;
					}
				}
			}
			finally
			{
				gridDados.ResumeLayout();
				gridDados.Update();
			}
		}
		#endregion

		#endregion

		#region [ Desmarcar todos os lançamentos ]

		#region [ btnDesmarcarTodos_Click ]
		private void btnDesmarcarTodos_Click(object sender, EventArgs e)
		{
			gridDados.SuspendLayout();
			try
			{
				for (int i = 0; i < gridDados.Rows.Count; i++)
				{
					if (gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value != null)
					{
						if ((bool)gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value)
						{
							gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value = false;
							gridDados.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
						}
					}
					else
					{
						gridDados.Rows[i].Cells[GRID_COL_CHECK_BOX].Value = false;
						gridDados.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
					}
				}
			}
			finally
			{
				gridDados.ResumeLayout();
				gridDados.Update();
			}
		}
		#endregion

		#endregion

		#endregion

		#region [ Eventos acionados pelo painel FFluxoEditaLote ]

		#region [ TrataEventoFluxoEditaLancamentoLoteAlterado ]
		public void TrataEventoFluxoEditaLancamentoLoteAlterado()
		{
			int intLancamentoSelecionado = 0;

			_blnEventoCallBackEmProcessamento = true;
			try
			{
				#region [ Memoriza o item atualmente selecionado ]
				foreach (DataGridViewRow item in gridDados.SelectedRows)
				{
					intLancamentoSelecionado = (int)Global.converteInteiro(item.Cells["id"].Value.ToString());
				}
				#endregion

				#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
				executaPesquisa();
				#endregion

				#region [ Restaura o item que estava anteriormente selecionado ]
				if (intLancamentoSelecionado > 0)
				{
					for (int i = 0; i < gridDados.Rows.Count; i++)
					{
						if (intLancamentoSelecionado == Global.converteInteiro(gridDados.Rows[i].Cells["id"].Value.ToString()))
						{
							gridDados.Rows[i].Selected = true;
							break;
						}
					}
				}
				#endregion
			}
			finally
			{
				_blnEventoCallBackEmProcessamento = false;
			}
		}
		#endregion

		#endregion

		#region [ Impressão ]

		#region [ prnDocConsulta_BeginPrint ]
		private void prnDocConsulta_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
		{
			if (gridDados.DataSource == null)
			{
				e.Cancel = true;
				return;
			}

			prnDocConsulta.DefaultPageSettings.Landscape = true;
			impressao = new Impressao(prnDocConsulta.DefaultPageSettings.Landscape);

			#region [ Prepara elementos de impressão ]
			fonteTitulo = new Font(NOME_FONTE_DEFAULT, 18, FontStyle.Bold);
			fonteListagem = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Regular);
			fonteListagemNegrito = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Bold);
			fonteDataEmissao = new Font(NOME_FONTE_DEFAULT, 9f, FontStyle.Regular);
			fonteFiltros = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Italic);
			fonteNumPagina = new Font(NOME_FONTE_DEFAULT, 10f, FontStyle.Bold);
			brushPadrao = new SolidBrush(Color.Black);
			penTracoTitulo = new Pen(brushPadrao, .5f);
			penTracoPontilhado = Impressao.criaPenTracoPontilhado();
			#endregion

			#region [ Inicialização ]
			_intConsultaImpressaoIdxLinhaGrid = 0;
			_intConsultaImpressaoNumPagina = 0;
			_strConsultaImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
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
			RectangleF r;
			String strTexto;
			#endregion

			#region [ Verifica se alguma consulta foi realizada ]
			if (gridDados.DataSource == null)
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
				prnDocConsulta.DocumentName = "Resultado da Consulta de Lançamentos";
				cxInicio = impressao.getLeftMarginInMm(e);
				larguraUtil = impressao.getWidthInMm(e);
				cxFim = cxInicio + larguraUtil;
				cyInicio = impressao.getTopMarginInMm(e);
				alturaUtil = impressao.getHeightInMm(e);
				cyFim = cyInicio + alturaUtil;
				cyRodapeNumPagina = cyFim - fonteNumPagina.GetHeight(e.Graphics) - 1;
				#endregion

				#region [ Layout das colunas ]
				ixDtCompetencia = cxInicio;
				wxDtCompetencia = 15f;
                ixComp2 = ixDtCompetencia + wxDtCompetencia + ESPACAMENTO_COLUNAS;
                wxComp2 = 15f;
                ixContaCorrente = ixComp2 + wxComp2 + ESPACAMENTO_COLUNAS;
                wxContaCorrente = 20f;
				ixPlanoContasConta = ixContaCorrente + wxContaCorrente + ESPACAMENTO_COLUNAS;
				wxPlanoContasConta = 50f;
				ixDescricao = ixPlanoContasConta + wxPlanoContasConta + ESPACAMENTO_COLUNAS;
				wxObs = 50f;
				ixObs = cxInicio + larguraUtil - wxObs;
				wxNomeCnpjCpf = 50f;
				ixNomeCnpjCpf = ixObs - wxNomeCnpjCpf - ESPACAMENTO_COLUNAS;
				wxValor = 20f;
				ixValor = ixNomeCnpjCpf - wxValor - ESPACAMENTO_COLUNAS;
				wxDescricao = ixValor - ixDescricao - ESPACAMENTO_COLUNAS;
				#endregion
			}

			cx = cxInicio;
			cy = cyInicio;

			#region [ Título ]
			strTexto = "Consulta de Lançamentos";
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

            #region [ Comp2 ]
            strTexto = "Comp2: ";
            if ((_filtroMesCompetenciaInicial.Length > 0) && (_filtroMesCompetenciaFinal.Length > 0))
                strTexto += _filtroMesCompetenciaInicial + " a " + _filtroMesCompetenciaFinal;
            else if (_filtroMesCompetenciaInicial.Length > 0)
                strTexto += _filtroMesCompetenciaInicial;
            else if (_filtroMesCompetenciaFinal.Length > 0)
                strTexto += _filtroMesCompetenciaFinal;
            else strTexto += "N.I.";

            cx = cxInicio + larguraUtil * .26f;
            e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
            #endregion

            #region [ Data de cadastro ]
            strTexto = "Cadastramento: ";
            if ((_filtroDataCadastroInicial.Length > 0) && (_filtroDataCadastroFinal.Length > 0))
                strTexto += _filtroDataCadastroInicial + " a " + _filtroDataCadastroFinal;
            else if (_filtroDataCadastroInicial.Length > 0)
                strTexto += _filtroDataCadastroInicial;
            else if (_filtroDataCadastroFinal.Length > 0)
                strTexto += _filtroDataCadastroFinal;
            else strTexto += "N.I.";

            cx = cxInicio + larguraUtil * .48f;
            e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
            #endregion

            #region [ Natureza ]
            strTexto = "Natureza: ";
            if (_filtroNatureza.Length > 0)
                strTexto += _filtroNatureza;
            else
                strTexto += "Todas";

            cx = cxInicio + larguraUtil * .78f;
            e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
            #endregion

            #region [ Nova linha ]
            cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ StSemEfeito ]
			strTexto = "Efeito: ";
			if (_filtroStSemEfeito.Length > 0)
				strTexto += _filtroStSemEfeito;
			else
				strTexto += "Todos";

			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ CtrlPagtoStatus ]
			strTexto = "Status: ";
			if (_filtroCtrlPagtoStatus.Length > 0)
				strTexto += _filtroCtrlPagtoStatus;
			else
				strTexto += "Todos";

			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Atrasados ]
			strTexto = "Atrasados: ";
			if (_filtroAtrasados.Length > 0)
				strTexto += _filtroAtrasados;
			else
				strTexto += "Todos";

			cx = cxInicio + larguraUtil * .75f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Valor ]
			strTexto = "Valor: ";
			if (_filtroValor.Length > 0)
				strTexto += _filtroValor;
			else
				strTexto += "N.I.";

			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Descrição ]
			strTexto = "Descrição: ";
			if (_filtroDescricao.Length > 0)
				strTexto += _filtroDescricao;
			else
				strTexto += "N.I.";
			
			cx = cxInicio + larguraUtil * .5f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Conta Corrente ]
			strTexto = "Conta Corrente: ";
			if (_filtroContaCorrente.Length > 0)
				strTexto += _filtroContaCorrente;
			else
				strTexto += "Todas";

			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Plano Contas Empresa ]
			strTexto = "Empresa: ";
			if (_filtroPlanoContasEmpresa.Length > 0)
				strTexto += _filtroPlanoContasEmpresa;
			else
				strTexto += "Todas";

			cx = cxInicio + larguraUtil * .5f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Plano Contas Grupo ]
			strTexto = "Grupo: ";
			if (_filtroPlanoContasGrupo.Length > 0)
				strTexto += _filtroPlanoContasGrupo;
			else
				strTexto += "Todos";

			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Plano Contas Conta ]
			strTexto = "Conta: ";
			if (_filtroPlanoContasConta.Length > 0)
				strTexto += _filtroPlanoContasConta;
			else
				strTexto += "Todos";

			cx = cxInicio + larguraUtil * .5f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Nome do Cliente ]
			strTexto = "Nome Cliente: ";
			if (_filtroNomeCliente.Length > 0)
				strTexto += _filtroNomeCliente;
			else
				strTexto += "N.I.";

			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ CNPJ/CPF ]
			strTexto = "CNPJ/CPF: ";
			if (_filtroCnpjCpf.Length > 0)
				strTexto += _filtroCnpjCpf;
			else
				strTexto += "N.I.";

			cx = cxInicio + larguraUtil * .5f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ NF ]
			strTexto = "NF: ";
			if (_filtroNF.Length > 0)
				strTexto += _filtroNF;
			else
				strTexto += "N.I.";

			cx = cxInicio + larguraUtil * .75f;
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
			strTexto = "Data";
			cx = ixDtCompetencia;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

            strTexto = "Comp2";
            cx = ixComp2;
            e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

            strTexto = "Cta Corrente";
			cx = ixContaCorrente;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Plano de Contas";
			cx = ixPlanoContasConta;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Descrição";
			cx = ixDescricao;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Valor";
			cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Nome/CNPJ/CPF";
			cx = ixNomeCnpjCpf;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Obs";
			cx = ixObs;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cy += fonteAtual.GetHeight(e.Graphics);
			cy += .5f;
			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
				   (_intConsultaImpressaoIdxLinhaGrid < gridDados.Rows.Count))
			{
				fonteAtual = fonteListagem;
				hMax = Math.Max(fonteListagem.GetHeight(e.Graphics), fonteListagemNegrito.GetHeight(e.Graphics));

				#region [ Data de competência ]
				cx = ixDtCompetencia;
				strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["dt_competencia_formatada"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
                #endregion

                cx = ixComp2;
                strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["dt_mes_competencia"].Value.ToString();
                e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

                cx = ixContaCorrente;
				strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["descricao_conta_corrente"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				cx = ixPlanoContasConta;
				r = new RectangleF(ixPlanoContasConta, cy, wxPlanoContasConta, 20);
				strTexto = "(" + gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["natureza"].Value.ToString() + ") " +
							gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["descricao_plano_contas_conta"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxPlanoContasConta).Height);

				cx = ixDescricao;
				r = new RectangleF(ixDescricao, cy, wxDescricao, 20);
				strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["descricao"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxDescricao).Height);

				cx = ixNomeCnpjCpf;
				r = new RectangleF(ixNomeCnpjCpf, cy, wxNomeCnpjCpf, 20);
				strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["nome_cnpj_cpf"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxNomeCnpjCpf).Height);

				cx = ixObs;
				r = new RectangleF(ixObs, cy, wxObs, 20);
				strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["obs"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxObs).Height);

				fonteAtual = fonteListagemNegrito;
				strTexto = gridDados.Rows[_intConsultaImpressaoIdxLinhaGrid].Cells["valor_formatado"].Value.ToString();
				cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				cy += hMax;
				cy += .5f;
				e.Graphics.DrawLine(penTracoPontilhado, cxInicio, cy, cxFim, cy);
				cy += .5f;

				_intConsultaImpressaoIdxLinhaGrid++;
			}

			#region [ Imprime nº página ]
			strTexto = _intConsultaImpressaoNumPagina.ToString();
			fonteAtual = fonteNumPagina;
			cy = cyRodapeNumPagina;
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			if (_intConsultaImpressaoIdxLinhaGrid < gridDados.Rows.Count)
				e.HasMorePages = true;
			else
				e.HasMorePages = false;
		}
		#endregion

		#endregion
	}
}
