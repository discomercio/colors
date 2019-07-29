#region [ using ]
using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing.Drawing2D;
#endregion

namespace Financeiro
{
	public partial class FFluxoRelatorioMovimentoAnalitico : Financeiro.FModelo
	{
		#region [ Enum ]
		enum eOpcaoFiltroPeriodoCompetencia
		{
			APLICAR_FILTRO = 1,
			IGNORAR_FILTRO = 2
		}
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

		DataTable _dtbConsulta;
		#endregion

		#region [ Menus ]
		ToolStripMenuItem menuLancamento;
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
		private String _filtroValor;
		private String _filtroCnpjCpf;
		private String _filtroDescricao;
		private String _filtroContaCorrente;
		private String _filtroPlanoContasEmpresa;
		private String _filtroPlanoContasGrupo;
		private String _filtroPlanoContasConta;
		private String _filtroPlanoContasGrupoInicial;
		private String _filtroPlanoContasGrupoFinal;
		private String _filtroChkIncluirAtrasados;
		private String _filtroChkCPF;
		private String _filtroChkCNPJ;
		#endregion

		#region [ Controle da Impressão ]
		const String NOME_FONTE_DEFAULT = "Microsoft Sans Serif";
		const float ESPACAMENTO_COLUNAS = 5.0f;
		private int _intConsultaImpressaoIdxLinha = 0;
		private int _intConsultaImpressaoNumPagina = 0;
		private String _strConsultaImpressaoDataEmissao;
		private String _strPlanoContasGrupoAnteriorId;
		private String _strPlanoContasGrupoAnteriorDescricao;
		private String _strPlanoContasContaAnteriorId;
		private String _strPlanoContasContaAnteriorDescricao;
		private bool _blnImprimeTitulos;
		private bool _blnQuebrarGrupo;
		private bool _blnImprimirTotalGrupo;
		private bool _blnQuebrarPlanoConta;
		private bool _blnImprimirTotalPlanoConta;
		private bool _blnTotalUltimoGrupoJaFoiImpresso;
		private bool _blnTotalUltimoPlanoContaJaFoiImpresso;
		private int _intLinhasImpressasTotal;
		Font fonteTitulo;
		Font fonteListagem;
		Font fonteListagemNegrito;
		Font fonteTituloGrupoContas;
		Font fonteDataEmissao;
		Font fonteFiltros;
		Font fonteNumPagina;
		Font fonteAtual;
		Brush brushPadrao;
		Pen penTracoTitulo;
		Pen penTracoPontilhado;
		float alturaLinhaListagem;
		float alturaLinhaListagemNegrito;
		float cxInicio;
		float cxFim;
		float cyInicio;
		float cyFim;
		float cyRodapeNumPagina;
		float larguraUtil;
		float alturaUtil;
		float ixNatureza;
		float wxNatureza;
		float ixDtCompetencia;
		float wxDtCompetencia;
        float ixComp2;
        float wxComp2;
        float ixDescricao;
		float wxDescricao;
		float ixValor;
		float wxValor;
		Impressao impressao;
		decimal vlValor;
		decimal vlTotalAcumulado;
		decimal vlSubTotalPlanoContasGrupo;
		decimal vlSubTotalPlanoContasConta;
		#endregion

		#endregion

		#region [ Construtor ]
		public FFluxoRelatorioMovimentoAnalitico()
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
            txtMesCompetenciaInicial.Text = "";
            txtMesCompetenciaFinal.Text = "";
            txtDataCadastroInicial.Text = "";
			txtDataCadastroFinal.Text = "";
			cbNatureza.SelectedIndex = -1;
			txtValor.Text = "";
			txtCnpjCpf.Text = "";
			txtDescricao.Text = "";
			cbContaCorrente.SelectedIndex = -1;
			cbPlanoContasEmpresa.SelectedIndex = -1;
			cbPlanoContasGrupo.SelectedIndex = -1;
			cbPlanoContasConta.SelectedIndex = -1;
			cbPlanoContasGrupoInicial.SelectedIndex = -1;
			cbPlanoContasGrupoFinal.SelectedIndex = -1;
			txtDataCompetenciaInicial.Focus();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Declarações ]
			const int MAX_PERIODO_EM_DIAS = 90;
			int intPlanoContasGrupo;
			int intPlanoContasGrupoInicial;
			int intPlanoContasGrupoFinal;
			DateTime dtCompetenciaInicial = DateTime.MinValue;
			DateTime dtCompetenciaFinal = DateTime.MinValue;
            DateTime dtMesCompetenciaInicial = DateTime.MinValue;
            DateTime dtMesCompetenciaFinal = DateTime.MinValue;
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

            #region [ Período do Mês de Competência ]
            if (txtMesCompetenciaInicial.Text.Trim().Length > 0)
            {
                if (!Global.isDataMMYYYYOk(txtMesCompetenciaInicial.Text))
                {
                    avisoErro("Data inválida!!");
                    txtMesCompetenciaInicial.Focus();
                    return false;
                }
                else dtMesCompetenciaInicial = Global.converteDdMmYyyyParaDateTime(Convert.ToDateTime(txtMesCompetenciaInicial.Text).ToString("dd/MM/yyyy"));
            }

            if (txtMesCompetenciaFinal.Text.Trim().Length > 0)
            {
                if (!Global.isDataMMYYYYOk(txtMesCompetenciaFinal.Text))
                {
                    avisoErro("Data inválida!!");
                    txtMesCompetenciaFinal.Focus();
                    return false;
                }
                else dtMesCompetenciaFinal = Global.converteDdMmYyyyParaDateTime(Convert.ToDateTime(txtMesCompetenciaFinal.Text).ToString("dd/MM/yyyy"));
            }

            if ((dtMesCompetenciaInicial > DateTime.MinValue) && (dtMesCompetenciaFinal > DateTime.MinValue))
            {
                if (dtMesCompetenciaInicial > dtMesCompetenciaFinal)
                {
                    avisoErro("A data final do período de competência é anterior à data inicial!!");
                    txtMesCompetenciaFinal.Focus();
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
			#endregion

			#region [ Obtém os valores dos combo-boxes de grupos de contas ]
			intPlanoContasGrupo = cbPlanoContasGrupo.SelectedValue == null ? 0 : (int)Global.converteInteiro(cbPlanoContasGrupo.SelectedValue.ToString());
			intPlanoContasGrupoInicial = cbPlanoContasGrupoInicial.SelectedValue == null ? 0 : (int)Global.converteInteiro(cbPlanoContasGrupoInicial.SelectedValue.ToString());
			intPlanoContasGrupoFinal = cbPlanoContasGrupoFinal.SelectedValue == null ? 0 : (int)Global.converteInteiro(cbPlanoContasGrupoFinal.SelectedValue.ToString());
			#endregion

			#region [ Faixa de Plano de Contas Grupo ]
			if ((intPlanoContasGrupoInicial > 0) && (intPlanoContasGrupoFinal > 0))
			{
				if (intPlanoContasGrupoInicial > intPlanoContasGrupoFinal)
				{
					avisoErro("Grupo de contas inicial é maior que o grupo de contas final!!");
					return false;
				}
			}
			#endregion

			#region [ Informou um grupo de contas específico e também uma faixa? ]
			if (intPlanoContasGrupo > 0)
			{
				if ((intPlanoContasGrupoInicial > 0) || (intPlanoContasGrupoFinal > 0))
				{
					avisoErro("Foi selecionado um grupo de contas específico, mas também foi definida uma faixa de grupos de contas para a consulta!!");
					return false;
				}
			}
			#endregion

			// Ok!
			return true;
		}
		#endregion

		#region [ montaClausulaWhere ]
		private String montaClausulaWhere(eOpcaoFiltroPeriodoCompetencia opcaoAplicarFiltroPeriodoCompetencia)
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;

			if (opcaoAplicarFiltroPeriodoCompetencia == eOpcaoFiltroPeriodoCompetencia.APLICAR_FILTRO)
			{
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
			}

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

			#region [ CNPJ/CPF ]
			if (Global.digitos(txtCnpjCpf.Text).Length > 0)
			{
				strAux = " (tFC.cnpj_cpf = '" + Global.digitos(txtCnpjCpf.Text) + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
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

			#region [ Faixa de Plano Contas Grupo ]

			#region [ Grupo Inicial ]
			if ((cbPlanoContasGrupoInicial.SelectedIndex > -1) && ((Global.converteInteiro(cbPlanoContasGrupoInicial.SelectedValue.ToString())) > 0))
			{
				strAux = " (tFC.id_plano_contas_grupo >= " + cbPlanoContasGrupoInicial.SelectedValue.ToString() + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Grupo Final ]
			if ((cbPlanoContasGrupoFinal.SelectedIndex > -1) && ((Global.converteInteiro(cbPlanoContasGrupoFinal.SelectedValue.ToString())) > 0))
			{
				strAux = " (tFC.id_plano_contas_grupo <= " + cbPlanoContasGrupoFinal.SelectedValue.ToString() + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#endregion

			#region [ Restrição Fixa ]
			strAux = " (tFC.st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO + ")";
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
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
			String strWhereAtrasados;
			String strWhereSomenteCpfCnpj = "";
			String strSqlAtrasados = "";
			DateTime dtCompetenciaInicial = DateTime.MinValue;
			DateTime dtCompetenciaFinal = DateTime.MinValue;
			DateTime dtReferenciaFinal = DateTime.MinValue;
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

			#region [ Somente lançamentos c/ CPF e/ou CNPJ cadastrados ]
			// Lembrando que os lançamentos gerados automaticamente devido aos boletos sempre possuem o CPF/CNPJ preenchido
			if (chkCPF.Checked)
			{
				if (strWhereSomenteCpfCnpj.Length > 0) strWhereSomenteCpfCnpj += " OR";
				strWhereSomenteCpfCnpj += " (tamanho_cnpj_cpf = " + Global.Cte.Etc.TAMANHO_CPF.ToString() + ")";
			}

			if (chkCNPJ.Checked)
			{
				if (strWhereSomenteCpfCnpj.Length > 0) strWhereSomenteCpfCnpj += " OR";
				strWhereSomenteCpfCnpj += " (tamanho_cnpj_cpf = " + Global.Cte.Etc.TAMANHO_CNPJ.ToString() + ")";
			}

			if (strWhereSomenteCpfCnpj.Length > 0) strWhereSomenteCpfCnpj = " (" + strWhereSomenteCpfCnpj + ")";
			#endregion

			#region [ Inclui atrasados? ]
			// O total em lançamentos atrasados que atendam aos critérios (exceto critérios de período de datas)
			// é projetado para o futuro, pois trata-se de um montante que a empresa espera receber futuramente.
			// Qualquer data futura que seja consultada com a opção de incluir os atrasados irá contabilizar
			// o valor dos lançamentos atrasados.
			// Ao consultar um período que contenha apenas datas passadas, contabiliza-se apenas o valor
			// realizado (confirmado), desprezando-se os atrasados, mesmo que a opção esteja assinalada.
			// A data que define se o lançamento está atrasado ou não é a data de crédito do último arquivo
			// de retorno processado.
			if (chkIncluirAtrasados.Checked)
			{
				if (txtDataCompetenciaInicial.Text.Length > 0) dtCompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaInicial.Text);
				if (txtDataCompetenciaFinal.Text.Length > 0) dtCompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtDataCompetenciaFinal.Text);
				dtReferenciaFinal = dtCompetenciaFinal;
				if (dtCompetenciaInicial > dtReferenciaFinal) dtReferenciaFinal = dtCompetenciaInicial;
				
				// Se a consulta envolve um período que está inteiramente antes da data limite dos
				// pagamentos em atraso, então o relatório exibe apenas o total de pagamentos realizados
				// (confirmados).
				// Mas se o período de consulta envolve um intervalo que é posterior à data limite dos
				// pagamentos em atraso, então o relatório vai exibir também os pagamentos previstos.
				// Neste caso, os pagamentos em atraso são computados, já que os pagamentos em atraso
				// tornam-se uma previsão de fluxo de caixa, que será realizado em algum momento no futuro.
				// IMPORTANTE: o id do registro do lançamento foi incluído porque quando há 2 ou mais
				// lançamentos cujos dados são todos iguais, apenas 1 é retornado.
				if ((dtReferenciaFinal == DateTime.MinValue) || (dtReferenciaFinal > dtReferenciaLimitePagamentoEmAtraso))
				{
					strWhereAtrasados = montaClausulaWhere(eOpcaoFiltroPeriodoCompetencia.IGNORAR_FILTRO);

					strSqlAtrasados =
							"SELECT " +
								" tFC.id AS id_lancamento," +
								" tFC.id_plano_contas_grupo," +
								" tPCG.descricao AS descricao_id_plano_contas_grupo," +
								" tFC.id_plano_contas_conta," +
								" tPCC.descricao AS descricao_id_plano_contas_conta," +
								" tFC.dt_competencia," +
								" tFC.natureza," +
								" tFC.descricao," +
								" tFC.valor," +
								" 'S' AS flag_em_atraso" +
							" FROM t_FIN_FLUXO_CAIXA tFC" +
								" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
									" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
								" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
									" ON (tFC.id_plano_contas_grupo=tPCG.id)" +
							" WHERE" +
								" (" +
									"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
									" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO.ToString() + ")" +
								  ")" +
								(strWhereAtrasados.Length > 0 ? " AND" : "") + strWhereAtrasados +
								(strWhereSomenteCpfCnpj.Length > 0 ? " AND" : "") + strWhereSomenteCpfCnpj;
				}
			}
			#endregion

			#region [ Monta cláusula Where ]
			strWhere = montaClausulaWhere(eOpcaoFiltroPeriodoCompetencia.APLICAR_FILTRO);
			#endregion

			#region [ Monta Select ]
			// Datas posteriores à data de crédito do último arquivo de retorno: considerar todos os 
			//		lançamentos previstos válidos (st_sem_efeito=0)
			// Datas anteriores à data de crédito do último arquivo de retorno: considerar apenas os 
			//		lançamentos realizados e válidos (st_sem_efeito=0 e st_confirmacao_pendente=0)
			// IMPORTANTE: o id do registro do lançamento foi incluído porque quando há 2 ou mais
			// lançamentos cujos dados são todos iguais, apenas 1 é retornado.
			strSql = "SELECT " +
						" id_lancamento," +
						" id_plano_contas_grupo," +
						" descricao_id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" descricao_id_plano_contas_conta," +
						" dt_competencia," +
                        " dt_mes_competencia," +
						" natureza," +
						" descricao," +
						" valor," +
						" flag_em_atraso" +
					" FROM " +
					"(" +
						"SELECT " +
							" tFC.id AS id_lancamento," +
							" tFC.id_plano_contas_grupo," +
							" tPCG.descricao AS descricao_id_plano_contas_grupo," +
							" tFC.id_plano_contas_conta," +
							" tPCC.descricao AS descricao_id_plano_contas_conta," +
							" tFC.dt_competencia," +
                            " tFC.dt_mes_competencia," +
							" tFC.natureza," +
							" tFC.descricao," +
							" tFC.valor," +
							" 'N' AS flag_em_atraso" +
						" FROM t_FIN_FLUXO_CAIXA tFC" +
							" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
								" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
							" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
								" ON (tFC.id_plano_contas_grupo=tPCG.id)" +
						" WHERE" +
							" (" +
								"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
								" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO.ToString() + ")" +
							  ")" +
							(strWhere.Length > 0 ? " AND" : "") + strWhere +
							(strWhereSomenteCpfCnpj.Length > 0 ? " AND" : "") + strWhereSomenteCpfCnpj +
						" UNION " +
						"SELECT " +
							" tFC.id AS id_lancamento," +
							" tFC.id_plano_contas_grupo," +
							" tPCG.descricao AS descricao_id_plano_contas_grupo," +
							" tFC.id_plano_contas_conta," +
							" tPCC.descricao AS descricao_id_plano_contas_conta," +
							" tFC.dt_competencia," +
                            " tFC.dt_mes_competencia," +
							" tFC.natureza," +
							" tFC.descricao," +
							" tFC.valor," +
							" 'N' AS flag_em_atraso" +
						" FROM t_FIN_FLUXO_CAIXA tFC" +
							" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
								" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
							" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
								" ON (tFC.id_plano_contas_grupo=tPCG.id)" +
						" WHERE" +
							" (" +
								"(dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
							  ")" +
							(strWhere.Length > 0 ? " AND" : "") + strWhere +
							(strWhereSomenteCpfCnpj.Length > 0 ? " AND" : "") + strWhereSomenteCpfCnpj +
						(strSqlAtrasados.Length > 0 ? " UNION " + strSqlAtrasados : "") +
					") t" +
					" ORDER BY" +
						" id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" dt_competencia," +
						" natureza," +
						" valor," +
						" descricao";
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
			_filtroValor = txtValor.Text;
			_filtroCnpjCpf = txtCnpjCpf.Text;
			_filtroDescricao = txtDescricao.Text;
			_filtroContaCorrente = cbContaCorrente.Text;
			_filtroPlanoContasEmpresa = cbPlanoContasEmpresa.Text;
			_filtroPlanoContasGrupo = cbPlanoContasGrupo.Text;
			_filtroPlanoContasConta = cbPlanoContasConta.Text;
			_filtroPlanoContasGrupoInicial = cbPlanoContasGrupoInicial.Text;
			_filtroPlanoContasGrupoFinal = cbPlanoContasGrupoFinal.Text;
			_filtroChkIncluirAtrasados = (chkIncluirAtrasados.Checked ? "Sim" : "Não");
			_filtroChkCPF = (chkCPF.Checked ? "Sim" : "Não");
			_filtroChkCNPJ = (chkCNPJ.Checked ? "Sim" : "Não");
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

				memorizaFiltrosParaImpressao();

				Global.Usuario.Defaults.relatorioMovimentoChkIncluirAtrasados = chkIncluirAtrasados.Checked.ToString();

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

		#region [ Form FFluxoRelatorio ]

		#region [ FFluxoRelatorio_Load ]
		private void FFluxoRelatorio_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			#endregion

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

				#region [ Combo Plano Contas Grupo - Inicial ]
				// Cria uma linha com a opção em branco
				DsDataSource.DtbPlanoContasGrupoComboDataTable dtbPlanoContasGrupoInicial = new DsDataSource.DtbPlanoContasGrupoComboDataTable();
				DsDataSource.DtbPlanoContasGrupoComboRow rowPlanoContasGrupoInicial = dtbPlanoContasGrupoInicial.NewDtbPlanoContasGrupoComboRow();
				rowPlanoContasGrupoInicial.id = 0;
				rowPlanoContasGrupoInicial.idComDescricao = "";
				dtbPlanoContasGrupoInicial.AddDtbPlanoContasGrupoComboRow(rowPlanoContasGrupoInicial);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbPlanoContasGrupoInicial.Merge(ComboDAO.criaDtbPlanoContasGrupoCombo(ComboDAO.eFiltraStAtivo.TODOS));
				cbPlanoContasGrupoInicial.DataSource = dtbPlanoContasGrupoInicial;
				cbPlanoContasGrupoInicial.ValueMember = "id";
				cbPlanoContasGrupoInicial.DisplayMember = "idComDescricao";
				cbPlanoContasGrupoInicial.SelectedIndex = -1;
				#endregion

				#region [ Combo Plano Contas Grupo - Final ]
				// Cria uma linha com a opção em branco
				DsDataSource.DtbPlanoContasGrupoComboDataTable dtbPlanoContasGrupoFinal = new DsDataSource.DtbPlanoContasGrupoComboDataTable();
				DsDataSource.DtbPlanoContasGrupoComboRow rowPlanoContasGrupoFinal = dtbPlanoContasGrupoFinal.NewDtbPlanoContasGrupoComboRow();
				rowPlanoContasGrupoFinal.id = 0;
				rowPlanoContasGrupoFinal.idComDescricao = "";
				dtbPlanoContasGrupoFinal.AddDtbPlanoContasGrupoComboRow(rowPlanoContasGrupoFinal);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbPlanoContasGrupoFinal.Merge(ComboDAO.criaDtbPlanoContasGrupoCombo(ComboDAO.eFiltraStAtivo.TODOS));
				cbPlanoContasGrupoFinal.DataSource = dtbPlanoContasGrupoFinal;
				cbPlanoContasGrupoFinal.ValueMember = "id";
				cbPlanoContasGrupoFinal.DisplayMember = "idComDescricao";
				cbPlanoContasGrupoFinal.SelectedIndex = -1;
				#endregion

				#region [ Checkbox: Incluir Atrasados ]
				if (Global.Usuario.Defaults.relatorioMovimentoChkIncluirAtrasados.Trim().Length > 0)
				{
					if (Global.Usuario.Defaults.relatorioMovimentoChkIncluirAtrasados.ToUpper().Equals("TRUE")) chkIncluirAtrasados.Checked = true;
				}
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

		#region [ FFluxoRelatorio_Shown ]
		private void FFluxoRelatorio_Shown(object sender, EventArgs e)
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

		#region [ FFluxoRelatorio_FormClosing ]
		private void FFluxoRelatorio_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FFluxoRelatorioMovimento_KeyDown ]
		private void FFluxoRelatorioMovimento_KeyDown(object sender, KeyEventArgs e)
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
			Global.trataTextBoxKeyDown(sender, e, cbNatureza);
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
			Global.trataTextBoxKeyDown(sender, e, txtCnpjCpf);
		}
		#endregion

		#region [ txtValor_KeyPress ]
		private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, txtDescricao);
		}
		#endregion

		#region [ txtCnpjCpf_KeyPress ]
		private void txtCnpjCpf_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCnpjCpf(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, cbContaCorrente);
		}
		#endregion

		#region [ txtDescricao_KeyPress ]
		private void txtDescricao_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
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
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasGrupoInicial);
		}
		#endregion

		#endregion

		#region [ cbPlanoContasGrupoInicial ]

		#region [ cbPlanoContasGrupoInicial_KeyDown ]
		private void cbPlanoContasGrupoInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbPlanoContasGrupoFinal);
		}
		#endregion

		#endregion

		#region [ cbPlanoContasGrupoFinal ]

		#region [ cbPlanoContasGrupoFinal_KeyDown ]
		private void cbPlanoContasGrupoFinal_KeyDown(object sender, KeyEventArgs e)
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
			if (_dtbConsulta == null)
			{
				e.Cancel = true;
				return;
			}

			if (_dtbConsulta.Rows.Count == 0)
			{
				e.Cancel = true;
				avisoErro("Não há dados!!");
				return;
			}

			impressao = new Impressao();

			#region [ Prepara elementos de impressão ]
			fonteTitulo = new Font(NOME_FONTE_DEFAULT, 18, FontStyle.Bold);
			fonteListagem = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Regular);
			fonteListagemNegrito = new Font(NOME_FONTE_DEFAULT, 8f, FontStyle.Bold);
			fonteTituloGrupoContas = new Font(NOME_FONTE_DEFAULT, 10f, FontStyle.Bold);
			fonteDataEmissao = new Font(NOME_FONTE_DEFAULT, 9f, FontStyle.Regular);
			fonteFiltros = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Italic);
			fonteNumPagina = new Font(NOME_FONTE_DEFAULT, 10f, FontStyle.Bold);
			brushPadrao = new SolidBrush(Color.Black);
			penTracoTitulo = new Pen(brushPadrao, .5f);
			penTracoPontilhado = Impressao.criaPenTracoPontilhado();
			#endregion

			_intConsultaImpressaoIdxLinha = 0;
			_intConsultaImpressaoNumPagina = 0;
			_strConsultaImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
			vlTotalAcumulado = 0;
			vlSubTotalPlanoContasGrupo = 0;
			_strPlanoContasGrupoAnteriorId = "XXXXXXXXXXXXXXXXXX";
			_strPlanoContasGrupoAnteriorDescricao = "";
			vlSubTotalPlanoContasConta = 0;
			_strPlanoContasContaAnteriorId = "XXXXXXXXXXXXXXXXXX";
			_strPlanoContasContaAnteriorDescricao = "";
			_blnImprimeTitulos = false;
			_blnQuebrarGrupo = false;
			_blnImprimirTotalGrupo = false;
			_blnQuebrarPlanoConta = false;
			_blnImprimirTotalPlanoConta = false;
			_blnTotalUltimoGrupoJaFoiImpresso = false;
			_blnTotalUltimoPlanoContaJaFoiImpresso = false;
			_intLinhasImpressasTotal = 0;
		}
		#endregion

		#region [ prnDocConsulta_PrintPage ]
		private void prnDocConsulta_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			float cx;
			float cy;
			float cy_i;
			float cy_f;
			float hMax;
			RectangleF r;
			String strTexto;
			String strAux;
			int intLinhasImpressasNestaPagina = 0;
			bool blnImprimiuTituloGrupo;
			bool blnUltLinhaDestePlanoConta;
			float hAux;
			#endregion

			#region [ Verifica se alguma consulta foi realizada ]
			if (_dtbConsulta == null)
			{
				e.Cancel = true;
				return;
			}

			if (_dtbConsulta.Rows.Count == 0)
			{
				e.Cancel = true;
				avisoErro("Não há dados!!");
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
				prnDocConsulta.DocumentName = "Relatório Analítico de Movimentos";
				cxInicio = impressao.getLeftMarginInMm(e);
				larguraUtil = impressao.getWidthInMm(e);
				cxFim = cxInicio + larguraUtil;
				cyInicio = impressao.getTopMarginInMm(e);
				alturaUtil = impressao.getHeightInMm(e);
				cyFim = cyInicio + alturaUtil;
				cyRodapeNumPagina = cyFim - fonteNumPagina.GetHeight(e.Graphics) - 1;
				alturaLinhaListagem = fonteListagem.GetHeight(e.Graphics);
				alturaLinhaListagemNegrito = fonteListagemNegrito.GetHeight(e.Graphics);
				#endregion

				#region [ Layout das colunas ]
				ixNatureza = cxInicio;
				wxNatureza = 14f;
				ixDtCompetencia = ixNatureza + wxNatureza + ESPACAMENTO_COLUNAS;
				wxDtCompetencia = 20f;
                ixComp2 = ixDtCompetencia + wxDtCompetencia + ESPACAMENTO_COLUNAS;
                wxComp2 = 15f;
                ixDescricao = ixComp2 + wxComp2 + ESPACAMENTO_COLUNAS;
				wxValor = 20f;
				ixValor = cxInicio + larguraUtil - wxValor;
				wxDescricao = ixValor - ixDescricao - ESPACAMENTO_COLUNAS;
				#endregion
			}

			cx = cxInicio;
			cy = cyInicio;

			#region [ Título ]
			strTexto = "Relatório Analítico de Movimentos";
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

			#region [ Linha ]
			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;
			#endregion

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

			#region [ Incluir atrasados ]
			strTexto = "Incluir Atrasados: ";
			if (_filtroChkIncluirAtrasados.Length > 0)
				strTexto += _filtroChkIncluirAtrasados;
			else
				strTexto += "N.I.";

			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Checkbox CPF ]
			strTexto = "CPF: ";
			if (_filtroChkCPF.Length > 0)
				strTexto += _filtroChkCPF;
			else
				strTexto += "N.I.";

			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Checkbox CNPJ ]
			strTexto = "CNPJ: ";
			if (_filtroChkCNPJ.Length > 0)
				strTexto += _filtroChkCNPJ;
			else
				strTexto += "N.I.";

			cx = cxInicio + larguraUtil * .66f;
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

			#region [ CNPJ/CPF ]
			strTexto = "CNPJ/CPF: ";
			if (_filtroCnpjCpf.Length > 0)
				strTexto += _filtroCnpjCpf;
			else
				strTexto += "N.I.";

			cx = cxInicio + larguraUtil * .25f;
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
			{
				strTexto += _filtroPlanoContasGrupo;
			}
			else
			{
				if ((_filtroPlanoContasGrupoInicial.Length > 0) || (_filtroPlanoContasGrupoFinal.Length > 0))
					strTexto += "N.I.";
				else
					strTexto += "Todos";
			}

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

			#region [ Plano Contas Grupo (Inicial) ]
			strTexto = "Grupo (inicial): ";
			if (_filtroPlanoContasGrupoInicial.Length > 0)
				strTexto += _filtroPlanoContasGrupoInicial;
			else
				strTexto += "N.I.";

			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Plano Contas Grupo (Final) ]
			strTexto = "Grupo (final): ";
			if (_filtroPlanoContasGrupoFinal.Length > 0)
				strTexto += _filtroPlanoContasGrupoFinal;
			else
				strTexto += "N.I.";

			cx = cxInicio + larguraUtil * .5f;
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

			_blnImprimeTitulos = true;

			while (_intConsultaImpressaoIdxLinha < _dtbConsulta.Rows.Count)
			{
				#region [ Mudou o grupo? ]
				// Lembre-se: pode ter impresso o total do grupo de contas anterior e ter pulado de página
				// por falta de espaço p/ o cabeçalho do próximo grupo de contas. Sem o 'if', o total seria
				// impresso novamente.
				if (!_blnQuebrarGrupo)
				{
					if (!_strPlanoContasGrupoAnteriorId.Equals(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["id_plano_contas_grupo"].ToString()))
					{
						_blnQuebrarGrupo = true;
						_blnImprimirTotalGrupo = true;
						_blnQuebrarPlanoConta = true;
						_blnImprimirTotalPlanoConta = true;
					}
				}
				#endregion

				#region [ Mudou o plano de conta? ]
				// Lembre-se: pode ter impresso o total do plano de conta anterior e ter pulado de página
				// por falta de espaço p/ o cabeçalho do próximo plano de conta. Sem o 'if', o total seria
				// impresso novamente.
				if (!_blnQuebrarPlanoConta)
				{
					if (!_strPlanoContasContaAnteriorId.Equals(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["id_plano_contas_conta"].ToString()))
					{
						_blnQuebrarPlanoConta = true;
						_blnImprimirTotalPlanoConta = true;
					}
				}
				#endregion

				#region [ Imprime total do plano de contas anterior? ]
				if (_blnImprimirTotalPlanoConta)
				{
					if (_intLinhasImpressasTotal > 0)
					{
						#region [ Espaçamento ]
						cy += 1f;
						#endregion

						#region [ Há espaço? ]
						if ((cy + alturaLinhaListagemNegrito + 2) > cyRodapeNumPagina) break;
						#endregion

						#region [ Imprime o total do plano de contas anterior ]
						fonteAtual = fonteListagemNegrito;
						strTexto = "Total do Plano de Contas  (" + _strPlanoContasContaAnteriorId.PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA, '0') + " - " + _strPlanoContasContaAnteriorDescricao + ")";
						cx = ixValor - ESPACAMENTO_COLUNAS - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

						strTexto = Global.formataMoeda(vlSubTotalPlanoContasConta);
						cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						cy += fonteAtual.GetHeight(e.Graphics);
						#endregion

						#region [ Espaçamento para o próximo plano de contas ]
						if (!_blnImprimirTotalGrupo) cy += 2f;
						#endregion
					}
					vlSubTotalPlanoContasConta = 0;
					_blnImprimirTotalPlanoConta = false;
				}
				#endregion

				#region [ Imprime total do grupo anterior? ]
				if (_blnImprimirTotalGrupo)
				{
					if (_intLinhasImpressasTotal > 0)
					{
						#region [ Espaçamento ]
						cy += 2f;
						#endregion

						#region [ Há espaço? ]
						if ((cy + alturaLinhaListagemNegrito + 2) > cyRodapeNumPagina) break;
						#endregion

						#region [ Imprime o total do grupo ]
						e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy, cxFim, cy);
						fonteAtual = fonteListagemNegrito;
						cy_i = cy;
						cy += 1f;
						cy_f = cy + fonteAtual.GetHeight(e.Graphics) + 1f;
						e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy_i, ixDescricao, cy_f);
						e.Graphics.DrawLine(penTracoTitulo, cxFim, cy_i, cxFim, cy_f);
						e.Graphics.FillRectangle(new LinearGradientBrush(new RectangleF(ixDescricao, cy_i, cxFim - ixDescricao, cy_f - cy_i), Color.White, Color.WhiteSmoke, LinearGradientMode.Vertical), ixDescricao, cy_i, cxFim - ixDescricao, cy_f - cy_i);
						strTexto = "Total do Grupo de Contas  (" + _strPlanoContasGrupoAnteriorId.PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO, '0') + " - " + _strPlanoContasGrupoAnteriorDescricao + ")";
						cx = ixValor - ESPACAMENTO_COLUNAS - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

						strTexto = Global.formataMoeda(vlSubTotalPlanoContasGrupo);
						cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
						e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
						cy += fonteAtual.GetHeight(e.Graphics);
						cy += 1f;
						e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy, cxFim, cy);
						#endregion
					}
					vlSubTotalPlanoContasGrupo = 0;
					_blnImprimirTotalGrupo = false;
				}
				#endregion

				#region [ Quebra por grupo? ]
				blnImprimiuTituloGrupo = false;
				if (_blnQuebrarGrupo)
				{
					#region [ Espaçamento ]
					if (intLinhasImpressasNestaPagina == 0)
						cy += 3f;
					else
						cy += 10f;
					#endregion

					#region [ Há espaço? ]
					if ((cy + 35) > cyRodapeNumPagina) break;
					#endregion

					#region [ Traço ]
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion

					#region [ Imprime nome do grupo ]
					fonteAtual = fonteTituloGrupoContas;
					strTexto = _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["id_plano_contas_grupo"].ToString().PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO, '0') + " - " + _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["descricao_id_plano_contas_grupo"].ToString();
					if (!_blnQuebrarGrupo) strTexto += "  (continuação)";
					cx = cxInicio + 3f;
					cy_i = cy - .5f;
					cy_f = cy + fonteAtual.GetHeight(e.Graphics) + .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy_i, cxInicio, cy_f);
					e.Graphics.DrawLine(penTracoTitulo, cxFim, cy_i, cxFim, cy_f);
					e.Graphics.FillRectangle(new LinearGradientBrush(new RectangleF(cxInicio, cy_i, larguraUtil, cy_f - cy_i), Color.WhiteSmoke, Color.LightGray, LinearGradientMode.Vertical), cxInicio, cy_i, larguraUtil, cy_f - cy_i);
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					cy += fonteAtual.GetHeight(e.Graphics);
					#endregion

					#region [ Traço ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion

					_strPlanoContasGrupoAnteriorId = _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["id_plano_contas_grupo"].ToString();
					_strPlanoContasGrupoAnteriorDescricao = _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["descricao_id_plano_contas_grupo"].ToString();
					_blnQuebrarGrupo = false;
					blnImprimiuTituloGrupo = true;
				}
				#endregion
				
				#region [ Imprime títulos/Quebra por plano de conta? ]
				if (_blnImprimeTitulos || _blnQuebrarPlanoConta)
				{
					#region [ Espaçamento ]
					cy += 2f;
					#endregion

					#region [ Há espaço? ]
					if (!blnImprimiuTituloGrupo)
					{
						if ((cy + 25) > cyRodapeNumPagina) break;
					}
					#endregion

					#region [ Traço ]
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion

					#region [ Imprime nome do plano de contas ]
					fonteAtual = fonteListagemNegrito;
					strTexto = _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["id_plano_contas_conta"].ToString().PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA, '0') + " - " + _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["descricao_id_plano_contas_conta"].ToString();
					if (!_blnQuebrarPlanoConta) strTexto += "  (continuação)";
					cx = cxInicio + 3f;
					cy_i = cy - .5f;
					cy_f = cy + fonteAtual.GetHeight(e.Graphics) + .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy_i, cxInicio, cy_f);
					e.Graphics.DrawLine(penTracoTitulo, cxFim, cy_i, cxFim, cy_f);
					e.Graphics.FillRectangle(new LinearGradientBrush(new RectangleF(cxInicio, cy_i, larguraUtil, cy_f - cy_i), Color.White, Color.WhiteSmoke, LinearGradientMode.Vertical), cxInicio, cy_i, larguraUtil, cy_f - cy_i);
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					cy += fonteAtual.GetHeight(e.Graphics);
					#endregion

					#region [ Traço ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion

					#region [ Títulos ]
					cy += .5f;
					fonteAtual = fonteListagemNegrito;

					strTexto = "Natureza";
					cx = ixNatureza;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = "Competência";
					cx = ixDtCompetencia;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

                    strTexto = "Comp2";
                    cx = ixComp2;
                    e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

                    strTexto = "Descrição";
					cx = ixDescricao;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = "Valor";
					cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					cy += fonteAtual.GetHeight(e.Graphics);
					cy += .5f;
					#endregion

					#region [ Linha ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion

					if (_blnQuebrarPlanoConta)
					{
						_strPlanoContasContaAnteriorId = _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["id_plano_contas_conta"].ToString();
						_strPlanoContasContaAnteriorDescricao = _dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["descricao_id_plano_contas_conta"].ToString();
						_blnQuebrarPlanoConta = false;
					}
				}
				#endregion

				_blnImprimeTitulos = false;

				#region [ Há espaço para mais 1 linha da listagem? ]
				blnUltLinhaDestePlanoConta = false;
				if ((cy + (6 * alturaLinhaListagemNegrito)) > cyRodapeNumPagina)
				{
					if (_intConsultaImpressaoIdxLinha < (_dtbConsulta.Rows.Count - 1))
					{
						if ((!_strPlanoContasGrupoAnteriorId.Equals(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha + 1]["id_plano_contas_grupo"].ToString()))
							||
							(!_strPlanoContasContaAnteriorId.Equals(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha + 1]["id_plano_contas_conta"].ToString())))
						{
							blnUltLinhaDestePlanoConta = true;
						}
					}
				}
				hAux = alturaLinhaListagem + 2;
				if ((cy + (blnUltLinhaDestePlanoConta ? ((2 * hAux) + 2) : hAux)) > cyRodapeNumPagina) break;
				#endregion

				fonteAtual = fonteListagem;
				hMax = Math.Max(alturaLinhaListagem, alturaLinhaListagemNegrito);

				#region [ Natureza ]
				cx = ixNatureza;
				strTexto = Global.retornaDescricaoFluxoCaixaNatureza(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["natureza"].ToString().ToCharArray()[0]);
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Data Competência ]
				cx = ixDtCompetencia;
				strTexto = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["dt_competencia"]));
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
                #endregion

                #region [ Comp2 ]
                cx = ixComp2;
                strTexto = Convert.ToDateTime(BD.readToDateTime(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["dt_mes_competencia"])) == DateTime.MinValue ? "" : Convert.ToDateTime(BD.readToDateTime(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["dt_mes_competencia"])).ToString("MM/yyyy");
                e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy); 
                #endregion

                #region [ Descrição ]
                cx = ixDescricao;
				r = new RectangleF(ixDescricao, cy, wxDescricao, 20);
				strTexto = BD.readToString(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["descricao"]);
				strAux = BD.readToString(_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["flag_em_atraso"]);
				if (strAux.Equals("S")) strTexto += "  (*** em atraso ***)";
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxDescricao).Height);
				#endregion

				fonteAtual = fonteListagemNegrito;

				#region [ Valor ]
				vlValor = (decimal)_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["valor"];
				if (_dtbConsulta.Rows[_intConsultaImpressaoIdxLinha]["natureza"].ToString().Equals(Global.Cte.FIN.Natureza.DEBITO.ToString()))
				{
					vlValor *= -1;
				}
				vlTotalAcumulado += vlValor;
				vlSubTotalPlanoContasGrupo += vlValor;
				vlSubTotalPlanoContasConta += vlValor;
				strTexto = Global.formataMoeda(vlValor);
				cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				cy += hMax;

				_intLinhasImpressasTotal++;
				intLinhasImpressasNestaPagina++;
				_intConsultaImpressaoIdxLinha++;

				#region [ Traço pontilhado ]
				cy += .5f;
				e.Graphics.DrawLine(penTracoPontilhado, cxInicio, cy, cxFim, cy);
				cy += .5f;
				#endregion
			}  // while (listagem)

			#region [ Imprime nº página ]
			strTexto = _intConsultaImpressaoNumPagina.ToString();
			fonteAtual = fonteNumPagina;
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cyRodapeNumPagina);
			#endregion

			#region [ Terminou a listagem? ]
			if (_intConsultaImpressaoIdxLinha < _dtbConsulta.Rows.Count)
			{
				e.HasMorePages = true;
			}
			else
			{
				e.HasMorePages = false;

				#region [ Imprime o total do último plano de contas? ]
				if (!_blnTotalUltimoPlanoContaJaFoiImpresso)
				{
					#region [ Espaçamento ]
					cy += 1f;
					#endregion

					#region [ Há espaço suficiente? ]
					if ((cy + alturaLinhaListagemNegrito + 2) > cyRodapeNumPagina)
					{
						e.HasMorePages = true;
						return;
					}
					#endregion

					fonteAtual = fonteListagemNegrito;
					strTexto = "Total do Plano de Contas  (" + _strPlanoContasContaAnteriorId.PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA, '0') + " - " + _strPlanoContasContaAnteriorDescricao + ")";
					cx = ixValor - ESPACAMENTO_COLUNAS - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(vlSubTotalPlanoContasConta);
					cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					cy += fonteAtual.GetHeight(e.Graphics);

					_blnTotalUltimoPlanoContaJaFoiImpresso = true;
				}
				#endregion

				#region [ Imprime o total do último grupo? ]
				if (!_blnTotalUltimoGrupoJaFoiImpresso)
				{
					#region [ Espaçamento ]
					cy += 2f;
					#endregion

					#region [ Há espaço suficiente? ]
					if ((cy + alturaLinhaListagemNegrito + 2) > cyRodapeNumPagina)
					{
						e.HasMorePages = true;
						return;
					}
					#endregion

					e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy, cxFim, cy);
					fonteAtual = fonteListagemNegrito;
					cy_i = cy;
					cy += 1f;
					cy_f = cy + fonteAtual.GetHeight(e.Graphics) + 1f;
					e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy_i, ixDescricao, cy_f);
					e.Graphics.DrawLine(penTracoTitulo, cxFim, cy_i, cxFim, cy_f);
					e.Graphics.FillRectangle(new LinearGradientBrush(new RectangleF(ixDescricao, cy_i, cxFim - ixDescricao, cy_f - cy_i), Color.White, Color.WhiteSmoke, LinearGradientMode.Vertical), ixDescricao, cy_i, cxFim - ixDescricao, cy_f - cy_i);
					strTexto = "Total do Grupo de Contas  (" + _strPlanoContasGrupoAnteriorId.PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO, '0') + " - " + _strPlanoContasGrupoAnteriorDescricao + ")";
					cx = ixValor - ESPACAMENTO_COLUNAS - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(vlSubTotalPlanoContasGrupo);
					cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					cy += fonteAtual.GetHeight(e.Graphics);
					cy += 1f;
					e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy, cxFim, cy);

					_blnTotalUltimoGrupoJaFoiImpresso = true;
				}
				#endregion

				#region [ Imprime o total geral ]

				#region [ Espaçamento ]
				if (intLinhasImpressasNestaPagina > 0)
				{
					cy += 6f;
				}
				else cy += 3f;
				#endregion

				#region [ Há espaço suficiente? ]
				if ((cy + alturaLinhaListagemNegrito + 4) > cyRodapeNumPagina)
				{
					e.HasMorePages = true;
					return;
				}
				#endregion

				#region [ Imprime o total geral ]
				e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy, cxFim, cy);
				fonteAtual = fonteListagemNegrito;
				cy_i = cy;
				cy += 1f;
				cy_f = cy + fonteAtual.GetHeight(e.Graphics) + 1f;
				e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy_i, ixDescricao, cy_f);
				e.Graphics.DrawLine(penTracoTitulo, cxFim, cy_i, cxFim, cy_f);
				e.Graphics.FillRectangle(new LinearGradientBrush(new RectangleF(ixDescricao, cy_i, cxFim - ixDescricao, cy_f - cy_i), Color.WhiteSmoke, Color.LightGray, LinearGradientMode.Vertical), ixDescricao, cy_i, cxFim - ixDescricao, cy_f - cy_i);
				strTexto = "TOTAL GERAL";
				cx = ixValor - ESPACAMENTO_COLUNAS - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

				strTexto = Global.formataMoeda(vlTotalAcumulado);
				cx = ixValor + wxValor - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				cy += fonteAtual.GetHeight(e.Graphics);
				cy += 1f;
				e.Graphics.DrawLine(penTracoTitulo, ixDescricao, cy, cxFim, cy);
				#endregion

				#endregion
			}
			#endregion
		}
		#endregion

		#endregion
	}
}
