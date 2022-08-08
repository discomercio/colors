#region [ using ]
using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing.Drawing2D;
using System.Media;
using System.Reflection;
using System.Collections.Generic;
using System.Threading;
#endregion

namespace Financeiro
{
	public partial class FFluxoRelatorioMovimentoSinteticoComparativo : Financeiro.FModelo
	{
		enum eOpcaoFiltroPeriodoCompetencia
		{
			APLICAR_FILTRO = 1,
			IGNORAR_FILTRO = 2
		}

		enum eOpcaoFiltroTipoSaida
		{
			NENHUM = 0,
			COMPARATIVO_ENTRE_PERIODOS = 1,
			COMPARATIVO_MENSAL = 2
		}

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
		List<RelSinteticoComparativoMensal> _saidaMensal = new List<RelSinteticoComparativoMensal>();
		#endregion

		#region [ Menus ]
		ToolStripMenuItem menuRelatorio;
		ToolStripMenuItem menuRelatorioExecutarConsulta;
		ToolStripMenuItem menuRelatorioLimpar;
		#endregion

		#region [ Memorização dos filtros ]
		private eOpcaoFiltroTipoSaida _filtroTipoSaida;
		private String _filtroPeriodo1DataCompetenciaInicial;
		private String _filtroPeriodo1DataCompetenciaFinal;
		private String _filtroPeriodo2DataCompetenciaInicial;
		private String _filtroPeriodo2DataCompetenciaFinal;
		private int _filtroPeriodoInicialMesCompetencia;
		private int _filtroPeriodoInicialAnoCompetencia;
		private int _filtroPeriodoFinalMesCompetencia;
		private int _filtroPeriodoFinalAnoCompetencia;
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
		DateTime dtMesCompetenciaInicial;
		DateTime dtMesCompetenciaFinal;
		DateTime dtAux;
		DateTime dtIteracao;
		#endregion

		#region [ Controle da Impressão ]
		private String _strConsultaImpressaoDataEmissao;
		private String _strPlanoContasGrupoAnterior;
		private bool _blnImprimirTotalGrupo;
		private bool _blnQuebrarGrupo;
		decimal vlValor;
		decimal vlTotalAcumuladoPeriodo1;
		decimal vlTotalAcumuladoPeriodo2;
		decimal vlSubTotalPlanoContasGrupoPeriodo1;
		decimal vlSubTotalPlanoContasGrupoPeriodo2;
		#endregion

		#endregion

		#region [ Construtor ]
		public FFluxoRelatorioMovimentoSinteticoComparativo()
		{
			InitializeComponent();

			#region [ Menu Relatorio ]
			// Menu principal de Relatorio
			menuRelatorio = new ToolStripMenuItem("&Relatório");
			menuRelatorio.Name = "menuRelatorio";
			// Executar consulta
			menuRelatorioExecutarConsulta = new ToolStripMenuItem("&Executar Consulta", null, menuRelatorioExecutarConsulta_Click);
			menuRelatorioExecutarConsulta.Name = "menuRelatorioExecutarConsulta";
			menuRelatorio.DropDownItems.Add(menuRelatorioExecutarConsulta);
			// Limpar
			menuRelatorioLimpar = new ToolStripMenuItem("&Limpar", null, menuRelatorioLimpar_Click);
			menuRelatorioLimpar.Name = "menuRelatorioLimpar";
			menuRelatorio.DropDownItems.Add(menuRelatorioLimpar);
			// Adiciona o menu Relatorio ao menu principal
			menuPrincipal.Items.Insert(1, menuRelatorio);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtPeriodo1DataCompetenciaInicial.Text = "";
			txtPeriodo1DataCompetenciaFinal.Text = "";
			txtPeriodo2DataCompetenciaInicial.Text = "";
			txtPeriodo2DataCompetenciaFinal.Text = "";
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
			txtPeriodo1DataCompetenciaInicial.Focus();
			_saidaMensal.Clear();
		}
		#endregion

		#region [ preencheCamposDefault ]
		private void preencheCamposDefault()
		{
			#region [ cbNatureza ]
			cbNatureza.SelectedIndex = -1;
			foreach (Global.OpcaoFluxoCaixaNatureza item in cbNatureza.Items)
			{
				if (item.codigo.Equals(Global.Cte.FIN.Natureza.DEBITO))
				{
					cbNatureza.SelectedIndex = cbNatureza.Items.IndexOf(item);
					break;
				}
			}
			#endregion

			#region [ Combo mês competência (inicial) ]
			cbCompetenciaMesInicial.SelectedIndex = -1;
			foreach (Global.OpcaoMes item in cbCompetenciaMesInicial.Items)
			{
				if (item.numero == DateTime.Today.Month)
				{
					cbCompetenciaMesInicial.SelectedIndex = cbCompetenciaMesInicial.Items.IndexOf(item);
					break;
				}
			}
			#endregion

			#region [ Combo ano competência (inicial) ]
			cbCompetenciaAnoInicial.SelectedIndex = -1;
			foreach (Global.OpcaoAno item in cbCompetenciaAnoInicial.Items)
			{
				if (item.numero == (DateTime.Today.Year - 1))
				{
					cbCompetenciaAnoInicial.SelectedIndex = cbCompetenciaAnoInicial.Items.IndexOf(item);
					break;
				}
			}
			#endregion

			#region [ Combo mês competência (final) ]
			cbCompetenciaMesFinal.SelectedIndex = -1;
			foreach (Global.OpcaoMes item in cbCompetenciaMesFinal.Items)
			{
				if (item.numero == DateTime.Today.Month)
				{
					cbCompetenciaMesFinal.SelectedIndex = cbCompetenciaMesFinal.Items.IndexOf(item);
					break;
				}
			}
			#endregion

			#region [ Combo ano competência (final) ]
			cbCompetenciaAnoFinal.SelectedIndex = -1;
			foreach (Global.OpcaoAno item in cbCompetenciaAnoFinal.Items)
			{
				if (item.numero == (DateTime.Today.Year))
				{
					cbCompetenciaAnoFinal.SelectedIndex = cbCompetenciaAnoFinal.Items.IndexOf(item);
					break;
				}
			}
			#endregion

			rbCompetenciaComparativoEntrePeriodos.Checked = false;
			rbCompetenciaComparativoMensal.Checked = false;
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
			int iMesAux;
			int iAnoAux;
			DateTime dtPeriodo1CompetenciaInicial = DateTime.MinValue;
			DateTime dtPeriodo1CompetenciaFinal = DateTime.MinValue;
			DateTime dtPeriodo2CompetenciaInicial = DateTime.MinValue;
			DateTime dtPeriodo2CompetenciaFinal = DateTime.MinValue;
			DateTime dtCompetenciaMesInicial = DateTime.MinValue;
			DateTime dtCompetenciaMesFinal = DateTime.MinValue;
			DateTime dtMesCompetenciaInicial = DateTime.MinValue;
			DateTime dtMesCompetenciaFinal = DateTime.MinValue;
			DateTime dtCadastroInicial = DateTime.MinValue;
			DateTime dtCadastroFinal = DateTime.MinValue;
			DateTime dtAuxInicial = DateTime.MinValue;
			DateTime dtAuxFinal = DateTime.MinValue;
			#endregion

			if ((!rbCompetenciaComparativoEntrePeriodos.Checked) && (!rbCompetenciaComparativoMensal.Checked))
			{
				avisoErro("Nenhuma opção de tipo de saída do relatório foi selecionada!");
				return false;
			}

			#region [ Tipo de saída: comparativo entre períodos ]

			#region [ Período Inicial da Data de Competência ]
			if (rbCompetenciaComparativoEntrePeriodos.Checked)
			{
				if (txtPeriodo1DataCompetenciaInicial.Text.Trim().Length > 0)
				{
					if (!Global.isDataOk(txtPeriodo1DataCompetenciaInicial.Text))
					{
						avisoErro("Data inválida!!");
						txtPeriodo1DataCompetenciaInicial.Focus();
						return false;
					}
					else dtPeriodo1CompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtPeriodo1DataCompetenciaInicial.Text);
				}

				if (txtPeriodo1DataCompetenciaFinal.Text.Trim().Length > 0)
				{
					if (!Global.isDataOk(txtPeriodo1DataCompetenciaFinal.Text))
					{
						avisoErro("Data inválida!!");
						txtPeriodo1DataCompetenciaFinal.Focus();
						return false;
					}
					else dtPeriodo1CompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtPeriodo1DataCompetenciaFinal.Text);
				}

				if ((dtPeriodo1CompetenciaInicial > DateTime.MinValue) && (dtPeriodo1CompetenciaFinal > DateTime.MinValue))
				{
					if (dtPeriodo1CompetenciaInicial > dtPeriodo1CompetenciaFinal)
					{
						avisoErro("A data final do período é anterior à data inicial!!");
						txtPeriodo1DataCompetenciaFinal.Focus();
						return false;
					}
				}
			}
			#endregion

			#region [ Período Final da Data de Competência ]
			if (rbCompetenciaComparativoEntrePeriodos.Checked)
			{
				if (txtPeriodo2DataCompetenciaInicial.Text.Trim().Length > 0)
				{
					if (!Global.isDataOk(txtPeriodo2DataCompetenciaInicial.Text))
					{
						avisoErro("Data inválida!!");
						txtPeriodo2DataCompetenciaInicial.Focus();
						return false;
					}
					else dtPeriodo2CompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtPeriodo2DataCompetenciaInicial.Text);
				}

				if (txtPeriodo2DataCompetenciaFinal.Text.Trim().Length > 0)
				{
					if (!Global.isDataOk(txtPeriodo2DataCompetenciaFinal.Text))
					{
						avisoErro("Data inválida!!");
						txtPeriodo2DataCompetenciaFinal.Focus();
						return false;
					}
					else dtPeriodo2CompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtPeriodo2DataCompetenciaFinal.Text);
				}

				if ((dtPeriodo2CompetenciaInicial > DateTime.MinValue) && (dtPeriodo2CompetenciaFinal > DateTime.MinValue))
				{
					if (dtPeriodo2CompetenciaInicial > dtPeriodo2CompetenciaFinal)
					{
						avisoErro("A data final do período é anterior à data inicial!!");
						txtPeriodo2DataCompetenciaFinal.Focus();
						return false;
					}
				}
			}
			#endregion

			#region [ Período inicial e final estão preenchidos corretamente? ]
			if (rbCompetenciaComparativoEntrePeriodos.Checked)
			{
				if (dtPeriodo1CompetenciaInicial == DateTime.MinValue)
				{
					avisoErro("A data inicial do período 1 não foi informada!");
					txtPeriodo1DataCompetenciaInicial.Focus();
					return false;
				}

				if (dtPeriodo1CompetenciaFinal == DateTime.MinValue)
				{
					avisoErro("A data final do período 1 não foi informada!");
					txtPeriodo1DataCompetenciaFinal.Focus();
					return false;
				}

				if (dtPeriodo2CompetenciaInicial == DateTime.MinValue)
				{
					avisoErro("A data inicial do período 2 não foi informada!");
					txtPeriodo2DataCompetenciaInicial.Focus();
					return false;
				}

				if (dtPeriodo2CompetenciaFinal == DateTime.MinValue)
				{
					avisoErro("A data final do período 2 não foi informada!");
					txtPeriodo2DataCompetenciaFinal.Focus();
					return false;
				}
			}
			#endregion

			#endregion

			#region [ Tipo de saída: comparativo mês a mês ]
			if (rbCompetenciaComparativoMensal.Checked)
			{
				if (cbCompetenciaMesInicial.SelectedIndex == -1)
				{
					avisoErro("Não foi selecionado o mês inicial do período de consulta!");
					cbCompetenciaMesInicial.Focus();
					return false;
				}

				if (cbCompetenciaAnoInicial.SelectedIndex == -1)
				{
					avisoErro("Não foi selecionado o ano inicial do período de consulta!");
					cbCompetenciaAnoInicial.Focus();
					return false;
				}

				if (cbCompetenciaMesFinal.SelectedIndex == -1)
				{
					avisoErro("Não foi selecionado o mês final do período de consulta!");
					cbCompetenciaMesFinal.Focus();
					return false;
				}

				if (cbCompetenciaAnoFinal.SelectedIndex == -1)
				{
					avisoErro("Não foi selecionado o ano final do período de consulta!");
					cbCompetenciaAnoFinal.Focus();
					return false;
				}

				iMesAux = ((Global.OpcaoMes)cbCompetenciaMesInicial.Items[cbCompetenciaMesInicial.SelectedIndex]).numero;
				iAnoAux = ((Global.OpcaoAno)cbCompetenciaAnoInicial.Items[cbCompetenciaAnoInicial.SelectedIndex]).numero;
				dtCompetenciaMesInicial = new DateTime(iAnoAux, iMesAux, 1);

				iMesAux = ((Global.OpcaoMes)cbCompetenciaMesFinal.Items[cbCompetenciaMesFinal.SelectedIndex]).numero;
				iAnoAux = ((Global.OpcaoAno)cbCompetenciaAnoFinal.Items[cbCompetenciaAnoFinal.SelectedIndex]).numero;
				dtCompetenciaMesFinal = new DateTime(iAnoAux, iMesAux, 1);

				if (dtCompetenciaMesInicial > dtCompetenciaMesFinal)
				{
					avisoErro("O mês/ano de competência inicial é posterior ao mês/ano de competência final!");
					cbCompetenciaMesInicial.Focus();
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

			#region [ Período de consulta é muito amplo? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if (rbCompetenciaComparativoEntrePeriodos.Checked)
				{
					if ((dtPeriodo1CompetenciaInicial > DateTime.MinValue) && (dtPeriodo1CompetenciaFinal > DateTime.MinValue))
					{
						if ((Global.calculaTimeSpanDias(dtPeriodo1CompetenciaFinal - dtPeriodo1CompetenciaInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
						{
							if (!confirma("O período 1 excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
						}
					}

					if ((dtPeriodo2CompetenciaInicial > DateTime.MinValue) && (dtPeriodo2CompetenciaFinal > DateTime.MinValue))
					{
						if ((Global.calculaTimeSpanDias(dtPeriodo2CompetenciaFinal - dtPeriodo2CompetenciaInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
						{
							if (!confirma("O período 2 excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
						}
					}
				}

				if ((dtCadastroInicial > DateTime.MinValue) && (dtCadastroFinal > DateTime.MinValue))
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

		#region [ montaClausulaWhereBase ]
		private String montaClausulaWhereBase()
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;

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

		#region [ montaClausulaWherePeriodo1 ]
		private String montaClausulaWherePeriodo1()
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;

			#region [ Data de competência ]
			if ((txtPeriodo1DataCompetenciaInicial.Text.Length > 0) && (txtPeriodo1DataCompetenciaFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtPeriodo1DataCompetenciaInicial.Text.Equals(txtPeriodo1DataCompetenciaFinal.Text))
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo1DataCompetenciaInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tFC.dt_competencia >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo1DataCompetenciaInicial.Text) + ") AND (tFC.dt_competencia <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo1DataCompetenciaFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtPeriodo1DataCompetenciaInicial.Text.Length > 0) || (txtPeriodo1DataCompetenciaFinal.Text.Length > 0))
			{
				if (txtPeriodo1DataCompetenciaInicial.Text.Length > 0)
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo1DataCompetenciaInicial.Text) + ")";
				}
				else if (txtPeriodo1DataCompetenciaFinal.Text.Length > 0)
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo1DataCompetenciaFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			return sbWhere.ToString();
		}
		#endregion

		#region [ montaClausulaWherePeriodo2 ]
		private String montaClausulaWherePeriodo2()
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;

			#region [ Data de competência ]
			if ((txtPeriodo2DataCompetenciaInicial.Text.Length > 0) && (txtPeriodo2DataCompetenciaFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtPeriodo2DataCompetenciaInicial.Text.Equals(txtPeriodo2DataCompetenciaFinal.Text))
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo2DataCompetenciaInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tFC.dt_competencia >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo2DataCompetenciaInicial.Text) + ") AND (tFC.dt_competencia <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo2DataCompetenciaFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtPeriodo2DataCompetenciaInicial.Text.Length > 0) || (txtPeriodo2DataCompetenciaFinal.Text.Length > 0))
			{
				if (txtPeriodo2DataCompetenciaInicial.Text.Length > 0)
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo2DataCompetenciaInicial.Text) + ")";
				}
				else if (txtPeriodo2DataCompetenciaFinal.Text.Length > 0)
				{
					strAux = " (tFC.dt_competencia = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtPeriodo2DataCompetenciaFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			return sbWhere.ToString();
		}
		#endregion

		#region [ montaSqlConsulta ]
		private String montaSqlConsulta()
		{
			#region [ Declarações ]
			String strWherePeriodo1;
			String strWherePeriodo2;
			String strWhereBase;
			String strSql = "";
			String strWhereSomenteCpfCnpj = "";
			String strSqlAtrasadosPeriodo1 = "";
			String strSqlAtrasadosPeriodo2 = "";
			String strAtrasadosSelectBase;
			String strAtrasadosFromBase;
			String strAtrasadosWhereBase;
			String strAtrasadosGroupByBase;
			String strConfirmadosSelectBase;
			String strConfirmadosFromBase;
			String strConfirmadosWhereBase;
			String strConfirmadosGroupByBase;
			String strPrevistosSelectBase;
			String strPrevistosFromBase;
			String strPrevistosWhereBase;
			String strPrevistosGroupByBase;
			StringBuilder sbSql;
			DateTime dtPeriodo1CompetenciaInicial = DateTime.MinValue;
			DateTime dtPeriodo1CompetenciaFinal = DateTime.MinValue;
			DateTime dtPeriodo2CompetenciaInicial = DateTime.MinValue;
			DateTime dtPeriodo2CompetenciaFinal = DateTime.MinValue;
			DateTime dtReferenciaFinalPeriodo1 = DateTime.MinValue;
			DateTime dtReferenciaFinalPeriodo2 = DateTime.MinValue;
			DateTime dtReferenciaFinalMensal;
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

			strWhereBase = montaClausulaWhereBase();

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

			#region [ Monta SQL base usado repetidamente ]

			#region [ Lançamentos em atraso ]
			strAtrasadosSelectBase =
					" tFC.id_plano_contas_grupo," +
					" tPCG.descricao AS descricao_id_plano_contas_grupo," +
					" tFC.id_plano_contas_conta," +
					" tPCC.descricao," +
					" tFC.natureza,";

			strAtrasadosFromBase =
					" t_FIN_FLUXO_CAIXA tFC" +
					" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
						" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
					" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
						" ON (tFC.id_plano_contas_grupo=tPCG.id)";

			strAtrasadosWhereBase =
					" (" +
						"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
						" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO.ToString() + ")" +
					")" +
					(strWhereBase.Length > 0 ? " AND" : "") + strWhereBase +
					(strWhereSomenteCpfCnpj.Length > 0 ? " AND" : "") + strWhereSomenteCpfCnpj;

			strAtrasadosGroupByBase =
					" tFC.id_plano_contas_grupo," +
					" tPCG.descricao," +
					" tFC.id_plano_contas_conta," +
					" tPCC.descricao," +
					" tFC.natureza";
			#endregion

			#region [ Lançamentos confirmados ]
			strConfirmadosSelectBase =
					" tFC.id_plano_contas_grupo," +
					" tPCG.descricao AS descricao_id_plano_contas_grupo," +
					" tFC.id_plano_contas_conta," +
					" tPCC.descricao," +
					" tFC.natureza,";

			strConfirmadosFromBase =
					" t_FIN_FLUXO_CAIXA tFC" +
					" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
						" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
					" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
						" ON (tFC.id_plano_contas_grupo=tPCG.id)";

			strConfirmadosWhereBase =
					" (" +
					"(dt_competencia <= " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
					" AND (st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO.ToString() + ")" +
					")" +
					(strWhereSomenteCpfCnpj.Length > 0 ? " AND" : "") + strWhereSomenteCpfCnpj;

			strConfirmadosGroupByBase =
					" tFC.id_plano_contas_grupo," +
					" tPCG.descricao," +
					" tFC.id_plano_contas_conta," +
					" tPCC.descricao," +
					" tFC.natureza";
			#endregion

			#region [ Lançamentos previstos ]
			strPrevistosSelectBase =
					" tFC.id_plano_contas_grupo," +
					" tPCG.descricao AS descricao_id_plano_contas_grupo," +
					" tFC.id_plano_contas_conta," +
					" tPCC.descricao," +
					" tFC.natureza,";

			strPrevistosFromBase =
					" t_FIN_FLUXO_CAIXA tFC" +
					" LEFT JOIN t_FIN_PLANO_CONTAS_CONTA tPCC" +
						" ON (tFC.id_plano_contas_conta=tPCC.id) AND (tFC.natureza=tPCC.natureza)" +
					" LEFT JOIN t_FIN_PLANO_CONTAS_GRUPO tPCG" +
						" ON (tFC.id_plano_contas_grupo=tPCG.id)";

			strPrevistosWhereBase =
					" (" +
						"(dt_competencia > " + Global.sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")" +
					")" +
					(strWhereSomenteCpfCnpj.Length > 0 ? " AND" : "") + strWhereSomenteCpfCnpj;

			strPrevistosGroupByBase =
					" tFC.id_plano_contas_grupo," +
					" tPCG.descricao," +
					" tFC.id_plano_contas_conta," +
					" tPCC.descricao," +
					" tFC.natureza";
			#endregion

			#endregion

			#region [ Tipo de saída: comparativo entre períodos ]
			if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
			{
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
					#region [ Período 1 ]
					if (txtPeriodo1DataCompetenciaInicial.Text.Length > 0) dtPeriodo1CompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtPeriodo1DataCompetenciaInicial.Text);
					if (txtPeriodo1DataCompetenciaFinal.Text.Length > 0) dtPeriodo1CompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtPeriodo1DataCompetenciaFinal.Text);
					dtReferenciaFinalPeriodo1 = dtPeriodo1CompetenciaFinal;
					if (dtPeriodo1CompetenciaInicial > dtReferenciaFinalPeriodo1) dtReferenciaFinalPeriodo1 = dtPeriodo1CompetenciaInicial;

					// Se a consulta envolve um período que está inteiramente antes da data limite dos
					// pagamentos em atraso, então o relatório exibe apenas o total de pagamentos realizados
					// (confirmados).
					// Mas se o período de consulta envolve um intervalo que é posterior à data limite dos
					// pagamentos em atraso, então o relatório vai exibir também os pagamentos previstos.
					// Neste caso, os pagamentos em atraso são computados, já que os pagamentos em atraso
					// tornam-se uma previsão de fluxo de caixa, que será realizado em algum momento no futuro.
					if ((dtReferenciaFinalPeriodo1 == DateTime.MinValue) || (dtReferenciaFinalPeriodo1 > dtReferenciaLimitePagamentoEmAtraso))
					{
						strSqlAtrasadosPeriodo1 =
								"SELECT" +
									strAtrasadosSelectBase +
									" Coalesce(Sum(tFC.valor),0) AS vl_total_periodo1," +
									" Sum(0) AS vl_total_periodo2" +
								" FROM" +
									strAtrasadosFromBase +
								" WHERE" +
									strAtrasadosWhereBase +
								" GROUP BY" +
									strAtrasadosGroupByBase;
					}
					#endregion

					#region [ Período 2 ]
					if (txtPeriodo2DataCompetenciaInicial.Text.Length > 0) dtPeriodo2CompetenciaInicial = Global.converteDdMmYyyyParaDateTime(txtPeriodo2DataCompetenciaInicial.Text);
					if (txtPeriodo2DataCompetenciaFinal.Text.Length > 0) dtPeriodo2CompetenciaFinal = Global.converteDdMmYyyyParaDateTime(txtPeriodo2DataCompetenciaFinal.Text);
					dtReferenciaFinalPeriodo2 = dtPeriodo2CompetenciaFinal;
					if (dtPeriodo2CompetenciaInicial > dtReferenciaFinalPeriodo2) dtReferenciaFinalPeriodo2 = dtPeriodo2CompetenciaInicial;

					// Se a consulta envolve um período que está inteiramente antes da data limite dos
					// pagamentos em atraso, então o relatório exibe apenas o total de pagamentos realizados
					// (confirmados).
					// Mas se o período de consulta envolve um intervalo que é posterior à data limite dos
					// pagamentos em atraso, então o relatório vai exibir também os pagamentos previstos.
					// Neste caso, os pagamentos em atraso são computados, já que os pagamentos em atraso
					// tornam-se uma previsão de fluxo de caixa, que será realizado em algum momento no futuro.
					if ((dtReferenciaFinalPeriodo2 == DateTime.MinValue) || (dtReferenciaFinalPeriodo2 > dtReferenciaLimitePagamentoEmAtraso))
					{
						strSqlAtrasadosPeriodo2 =
								"SELECT " +
									strAtrasadosSelectBase +
									" Sum(0) AS vl_total_periodo1," +
									" Coalesce(Sum(tFC.valor),0) AS vl_total_periodo2" +
								" FROM" +
									strAtrasadosFromBase +
								" WHERE" +
									strAtrasadosWhereBase +
								" GROUP BY" +
									strAtrasadosGroupByBase;
					}
					#endregion
				}
				#endregion

				#region [ Monta cláusula Where ]
				strWherePeriodo1 = montaClausulaWherePeriodo1();
				if ((strWherePeriodo1.Length > 0) && (strWhereBase.Length > 0)) strWherePeriodo1 += " AND";
				strWherePeriodo1 += strWhereBase;

				strWherePeriodo2 = montaClausulaWherePeriodo2();
				if ((strWherePeriodo2.Length > 0) && (strWhereBase.Length > 0)) strWherePeriodo2 += " AND";
				strWherePeriodo2 += strWhereBase;
				#endregion

				#region [ Monta Select ]
				// Datas posteriores à data de crédito do último arquivo de retorno: considerar todos os 
				//		lançamentos previstos válidos (st_sem_efeito=0)
				// Datas anteriores à data de crédito do último arquivo de retorno: considerar apenas os 
				//		lançamentos realizados e válidos (st_sem_efeito=0 e st_confirmacao_pendente=0)
				strSql = "SELECT" +
						   " id_plano_contas_grupo," +
						   " descricao_id_plano_contas_grupo," +
						   " id_plano_contas_conta," +
						   " descricao," +
						   " natureza," +
						   " Coalesce(Sum(vl_total_periodo1),0) AS vl_total_periodo1," +
						   " Coalesce(Sum(vl_total_periodo2),0) AS vl_total_periodo2" +
					   " FROM " +
					   "(" +
						   "SELECT " +
							   strConfirmadosSelectBase +
							   " Coalesce(Sum(tFC.valor),0) AS vl_total_periodo1," +
							   " Sum(0) AS vl_total_periodo2" +
						   " FROM" +
							   strConfirmadosFromBase +
						   " WHERE" +
							   strConfirmadosWhereBase +
							   (strWherePeriodo1.Length > 0 ? " AND" : "") + strWherePeriodo1 +
						   " GROUP BY" +
							   strConfirmadosGroupByBase +
						   " UNION ALL " +
						   "SELECT " +
							   strPrevistosSelectBase +
							   " Coalesce(Sum(tFC.valor),0) AS vl_total_periodo1," +
							   " Sum(0) AS vl_total_periodo2" +
						   " FROM" +
							   strPrevistosFromBase +
						   " WHERE" +
							   strPrevistosWhereBase +
							   (strWherePeriodo1.Length > 0 ? " AND" : "") + strWherePeriodo1 +
						   " GROUP BY" +
							   strPrevistosGroupByBase +
						   (strSqlAtrasadosPeriodo1.Length > 0 ? " UNION ALL " + strSqlAtrasadosPeriodo1 : "") +
						   " UNION ALL " +
						   "SELECT " +
							   strConfirmadosSelectBase +
							   " Sum(0) AS vl_total_periodo1," +
							   " Coalesce(Sum(tFC.valor),0) AS vl_total_periodo2" +
						   " FROM" +
							   strConfirmadosFromBase +
						   " WHERE" +
							   strConfirmadosWhereBase +
							   (strWherePeriodo2.Length > 0 ? " AND" : "") + strWherePeriodo2 +
						   " GROUP BY" +
							   strConfirmadosGroupByBase +
						   " UNION ALL " +
						   "SELECT " +
							   strPrevistosSelectBase +
							   " Sum(0) AS vl_total_periodo1," +
							   " Coalesce(Sum(tFC.valor),0) AS vl_total_periodo2" +
						   " FROM" +
							   strPrevistosFromBase +
						   " WHERE" +
							   strPrevistosWhereBase +
							   (strWherePeriodo2.Length > 0 ? " AND" : "") + strWherePeriodo2 +
						   " GROUP BY" +
							   strPrevistosGroupByBase +
						   (strSqlAtrasadosPeriodo2.Length > 0 ? " UNION ALL " + strSqlAtrasadosPeriodo2 : "") +
					   ") t" +
					   " GROUP BY" +
						   " id_plano_contas_grupo," +
						   " descricao_id_plano_contas_grupo," +
						   " id_plano_contas_conta," +
						   " descricao," +
						   " natureza" +
					   " ORDER BY" +
						   " id_plano_contas_grupo," +
						   " descricao_id_plano_contas_grupo," +
						   " id_plano_contas_conta," +
						   " descricao," +
						   " natureza";
				#endregion
			}
			#endregion

			#region [ Tipo de saída: comparativo mês a mês ]
			if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
			{
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
					for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
					{
						// Se a consulta envolve um período que está inteiramente antes da data limite dos
						// pagamentos em atraso, então o relatório exibe apenas o total de pagamentos realizados
						// (confirmados).
						// Mas se o período de consulta envolve um intervalo que é posterior à data limite dos
						// pagamentos em atraso, então o relatório vai exibir também os pagamentos previstos.
						// Neste caso, os pagamentos em atraso são computados, já que os pagamentos em atraso
						// tornam-se uma previsão de fluxo de caixa, que será realizado em algum momento no futuro.
						dtReferenciaFinalMensal = _saidaMensal[iMensal].dtMesAno.AddMonths(1).AddDays(-1);
						sbSql = new StringBuilder("");
						if (dtReferenciaFinalMensal > dtReferenciaLimitePagamentoEmAtraso)
						{
							sbSql.Append("SELECT" + strAtrasadosSelectBase);
							for (int i = 0; i < _saidaMensal.Count; i++)
							{
								if (i == iMensal)
								{
									sbSql.Append(" Coalesce(Sum(tFC.valor),0) AS " + _saidaMensal[i].keyName);
								}
								else
								{
									sbSql.Append(" Sum(0) AS " + _saidaMensal[i].keyName);
								}

								if (i < (_saidaMensal.Count - 1)) sbSql.Append(",");
							}
							sbSql.Append(" FROM" + strAtrasadosFromBase);
							sbSql.Append(" WHERE" + strAtrasadosWhereBase);
							sbSql.Append(" GROUP BY" + strAtrasadosGroupByBase);
							_saidaMensal[iMensal].sqlAtrasados = sbSql.ToString();
						}
					}
				}
				#endregion

				#region [ Monta cláusula Where ]
				for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
				{
					_saidaMensal[iMensal].sqlWherePeriodo = " ((tFC.dt_competencia >= " + Global.sqlMontaDateTimeParaSqlDateTime(_saidaMensal[iMensal].dtMesAno) + ") AND (tFC.dt_competencia < " + Global.sqlMontaDateTimeParaSqlDateTime(_saidaMensal[iMensal].dtMesAno.AddMonths(1)) + "))" +
															" AND " + strWhereBase;
				}
				#endregion

				#region [ Monta Select ]
				// Datas posteriores à data de crédito do último arquivo de retorno: considerar todos os 
				//		lançamentos previstos válidos (st_sem_efeito=0)
				// Datas anteriores à data de crédito do último arquivo de retorno: considerar apenas os 
				//		lançamentos realizados e válidos (st_sem_efeito=0 e st_confirmacao_pendente=0)
				sbSql = new StringBuilder("");
				sbSql.Append("SELECT" +
								" id_plano_contas_grupo," +
								" descricao_id_plano_contas_grupo," +
								" id_plano_contas_conta," +
								" descricao," +
								" natureza,");
				for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
				{
					sbSql.Append(" Coalesce(Sum(" + _saidaMensal[iMensal].keyName + "),0) AS " + _saidaMensal[iMensal].keyName);
					if (iMensal < (_saidaMensal.Count - 1)) sbSql.Append(",");
				}
				sbSql.Append(" FROM ");
				sbSql.Append("(");
				for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
				{
					if (iMensal > 0) sbSql.Append(" UNION ALL ");

					sbSql.Append("SELECT " + strConfirmadosSelectBase);
					for (int i = 0; i < _saidaMensal.Count; i++)
					{
						if (i == iMensal)
						{
							sbSql.Append(" Coalesce(Sum(tFC.valor),0) AS " + _saidaMensal[i].keyName);
						}
						else
						{
							sbSql.Append(" Sum(0) AS " + _saidaMensal[i].keyName);
						}
						if (i < (_saidaMensal.Count - 1)) sbSql.Append(",");
					}
					sbSql.Append(" FROM" + strConfirmadosFromBase);
					sbSql.Append(" WHERE" + strConfirmadosWhereBase);
					if (_saidaMensal[iMensal].sqlWherePeriodo.Length > 0) sbSql.Append(" AND" + _saidaMensal[iMensal].sqlWherePeriodo);
					sbSql.Append(" GROUP BY" + strConfirmadosGroupByBase);
					sbSql.Append(" UNION ALL ");
					sbSql.Append("SELECT " + strPrevistosSelectBase);
					for (int i = 0; i < _saidaMensal.Count; i++)
					{
						if (i == iMensal)
						{
							sbSql.Append(" Coalesce(Sum(tFC.valor),0) AS " + _saidaMensal[i].keyName);
						}
						else
						{
							sbSql.Append(" Sum(0) AS " + _saidaMensal[i].keyName);
						}
						if (i < (_saidaMensal.Count - 1)) sbSql.Append(",");
					}
					sbSql.Append(" FROM" + strPrevistosFromBase);
					sbSql.Append(" WHERE" + strPrevistosWhereBase);
					if (_saidaMensal[iMensal].sqlWherePeriodo.Length > 0) sbSql.Append(" AND" + _saidaMensal[iMensal].sqlWherePeriodo);
					sbSql.Append(" GROUP BY" + strPrevistosGroupByBase);
					if (_saidaMensal[iMensal].sqlAtrasados.Length > 0)
					{
						sbSql.Append(" UNION ALL " + _saidaMensal[iMensal].sqlAtrasados);
					}
				}
				sbSql.Append(") t");
				sbSql.Append(" GROUP BY" +
								" id_plano_contas_grupo," +
								" descricao_id_plano_contas_grupo," +
								" id_plano_contas_conta," +
								" descricao," +
								" natureza");
				sbSql.Append(" ORDER BY" +
								" id_plano_contas_grupo," +
								" descricao_id_plano_contas_grupo," +
								" id_plano_contas_conta," +
								" descricao," +
								" natureza");
				#endregion

				strSql = sbSql.ToString();
			}
			#endregion

			Global.gravaLogAtividade(strSql);

			return strSql;
		}
		#endregion

		#region [ memorizaFiltrosParaMontagemRelatorio ]
		/// <summary>
		/// Memoriza os parâmetros usados na última pesquisa para serem usados na montagem do relatório.
		/// </summary>
		private void memorizaFiltrosParaMontagemRelatorio()
		{
			if (rbCompetenciaComparativoEntrePeriodos.Checked)
			{
				_filtroTipoSaida = eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS;
			}
			else if (rbCompetenciaComparativoMensal.Checked)
			{
				_filtroTipoSaida = eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL;
			}
			else
			{
				_filtroTipoSaida = eOpcaoFiltroTipoSaida.NENHUM;
			}

			_filtroPeriodo1DataCompetenciaInicial = "";
			_filtroPeriodo1DataCompetenciaFinal = "";
			_filtroPeriodo2DataCompetenciaInicial = "";
			_filtroPeriodo2DataCompetenciaFinal = "";
			_filtroPeriodoInicialMesCompetencia = 0;
			_filtroPeriodoInicialAnoCompetencia = 0;
			_filtroPeriodoFinalMesCompetencia = 0;
			_filtroPeriodoFinalAnoCompetencia = 0;

			if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
			{
				_filtroPeriodo1DataCompetenciaInicial = txtPeriodo1DataCompetenciaInicial.Text;
				_filtroPeriodo1DataCompetenciaFinal = txtPeriodo1DataCompetenciaFinal.Text;
				_filtroPeriodo2DataCompetenciaInicial = txtPeriodo2DataCompetenciaInicial.Text;
				_filtroPeriodo2DataCompetenciaFinal = txtPeriodo2DataCompetenciaFinal.Text;
			}

			if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
			{
				if (cbCompetenciaMesInicial.SelectedIndex > -1) _filtroPeriodoInicialMesCompetencia = ((Global.OpcaoMes)cbCompetenciaMesInicial.Items[cbCompetenciaMesInicial.SelectedIndex]).numero;
				if (cbCompetenciaAnoInicial.SelectedIndex > -1) _filtroPeriodoInicialAnoCompetencia = ((Global.OpcaoAno)cbCompetenciaAnoInicial.Items[cbCompetenciaAnoInicial.SelectedIndex]).numero;
				if (cbCompetenciaMesFinal.SelectedIndex > -1) _filtroPeriodoFinalMesCompetencia = ((Global.OpcaoMes)cbCompetenciaMesFinal.Items[cbCompetenciaMesFinal.SelectedIndex]).numero;
				if (cbCompetenciaAnoFinal.SelectedIndex > -1) _filtroPeriodoFinalAnoCompetencia = ((Global.OpcaoAno)cbCompetenciaAnoFinal.Items[cbCompetenciaAnoFinal.SelectedIndex]).numero;

				if ((_filtroPeriodoInicialMesCompetencia > 0) && (_filtroPeriodoInicialAnoCompetencia > 0) && (_filtroPeriodoFinalMesCompetencia > 0) && (_filtroPeriodoFinalAnoCompetencia > 0))
				{
					_saidaMensal.Clear();
					dtMesCompetenciaInicial = new DateTime(_filtroPeriodoInicialAnoCompetencia, _filtroPeriodoInicialMesCompetencia, 1);
					dtMesCompetenciaFinal = new DateTime(_filtroPeriodoFinalAnoCompetencia, _filtroPeriodoFinalMesCompetencia, 1);
					if (dtMesCompetenciaInicial > dtMesCompetenciaFinal)
					{
						dtAux = dtMesCompetenciaFinal;
						dtMesCompetenciaFinal = dtMesCompetenciaInicial;
						dtMesCompetenciaInicial = dtAux;
					}

					dtIteracao = dtMesCompetenciaInicial;
					while (dtIteracao <= dtMesCompetenciaFinal)
					{
						_saidaMensal.Add(new RelSinteticoComparativoMensal(dtIteracao.Month, dtIteracao.Year));
						dtIteracao = dtIteracao.AddMonths(1);
					}
				}
			}

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

				memorizaFiltrosParaMontagemRelatorio();

				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				_dtbConsulta = new DataTable();
				#endregion

				#region [ Monta o SQL da consulta ]
				strSql = montaSqlConsulta();
				if ((strSql ?? "").Length == 0)
				{
					avisoErro("Falha ao montar a consulta!!");
					return false;
				}
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

		#region [ geraPlanilhaExcel ]
		private bool geraPlanilhaExcel()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FFluxoRelatorioMovimentoSinteticoComparativo.geraPlanilhaExcel()";
			const int MAX_LINHAS_EXCEL = 65536;
			const String FN_LISTAGEM = "Arial";
			const int FS_LISTAGEM = 8;
			const int FS_CABECALHO = 8;
			bool _blnImprimeTitulos;
			bool blnFlag;
			bool blnExcelSuportaUseSystemSeparators = false;
			bool blnExcelSuportaDecimalDataType = false;
			int iNumLinha = 1;
			int iPrimeiraLinhaDadosGrupo = 0;
			int iUltimaLinhaDadosGrupo = 0;
			int iXlDadosMinCol;
			int iXlDadosMaxCol;
			int iXlMargemEsq;
			int iXlNatureza;
			int iXlPlanoContas;
			int iXlVlPeriodo1;
			int iXlVlPeriodo2;
			int iXlAux;
			string strMsg;
			string strTexto;
			string strAux;
			string strExcelDecimalSeparator = "";
			string strExcelThousandsSeparator = "";
			StringBuilder sbTexto;
			object oXL = null;
			object oWBs = null;
			object oWB = null;
			object oWS = null;
			object oWindow = null;
			object oWindows = null;
			object oPageSetup = null;
			object oStyles = null;
			object oStyle = null;
			object oFont = null;
			object oBorders = null;
			object oBorder = null;
			object oCells = null;
			object oCell = null;
			object oColumns = null;
			object oColumn = null;
			object oRows = null;
			object oRow = null;
			object oRange = null;
			object oRangeInterior = null;
			object oApplication = null;
			List<int> listaNumLinhaTotalGrupo = new List<int>();
			#endregion

			#region [ Observações Importantes ]
			// Observações:
			// ============
			// 1) Todas as referências ao Excel devem ser devidamentes desalocadas, senão o processo do Excel não é encerrado ao final.
			//    Se uma variável for ser reutilizada para acessar outro objeto, como é o caso de 'range' por exemplo, antes de atribuir as novas referências, as anteriores devem ser desalocadas.
			//    Os comandos para desalocar as referências foram encapsuladas na rotina ExcelAutomation.NAR(), seguindo orientações do artigo https://support.microsoft.com/en-us/kb/317109
			// 2) O comando p/ maximizar a janela do Excel deve ser evitado porque senão o processo do Excel não é encerrado ao final, mesmo executando os comandos p/ desalocar as referências:
			//    ExcelAutomation.SetProperty(oXL, "WindowState", ExcelAutomation.XlWindowState.xlMaximized);
			// 3) Após realizar alterações nesta rotina, deve-se verificar se o processo do Excel está sendo encerrado ao final ou se está ficando pendente.
			#endregion

			if (_dtbConsulta == null)
			{
				aviso("O resultado da pesquisa não possui dados!");
				return false;
			}

			if (_dtbConsulta.Rows.Count == 0)
			{
				aviso("Não há dados no resultado da pesquisa!");
				return false;
			}

			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "gerando planilha Excel");

				_strConsultaImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
				vlTotalAcumuladoPeriodo1 = 0;
				vlTotalAcumuladoPeriodo2 = 0;
				vlSubTotalPlanoContasGrupoPeriodo1 = 0;
				vlSubTotalPlanoContasGrupoPeriodo2 = 0;
				_strPlanoContasGrupoAnterior = "XXXXXXXXXXXXXXXXXX";
				_blnQuebrarGrupo = false;
				_blnImprimirTotalGrupo = false;

				try // finally
				{
					#region [ Cria instância do Excel ]
					try
					{
						oXL = ExcelAutomation.CriaInstanciaExcel();
					}
					catch (Exception ex)
					{
						strMsg = "Falha ao acionar o Excel!!\nVerifique se o Excel está instalado!!\n\n" + ex.ToString();
						avisoErro(strMsg);
						return false;
					}
					#endregion

					#region [ Inicializa planilha ]
					ExcelAutomation.SetProperty(oXL, "Visible", true);
					ExcelAutomation.SetProperty(oXL, "DisplayAlerts", false);
					ExcelAutomation.SetProperty(oXL, "SheetsInNewWorkbook", 1);
					ExcelAutomation.NAR(oWBs);
					oWBs = ExcelAutomation.GetProperty(oXL, ExcelAutomation.PropertyType.Workbooks);
					oWB = ExcelAutomation.InvokeMethod(oWBs, "Add", Missing.Value);
					ExcelAutomation.NAR(oWindows);
					oWindows = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.Windows);
					ExcelAutomation.NAR(oWindow);
					oWindow = ExcelAutomation.GetProperty(oWindows, "Item", 1);
					ExcelAutomation.SetProperty(oWindow, "DisplayGridlines", false);
					ExcelAutomation.SetProperty(oWindow, "DisplayHeadings", true);
					ExcelAutomation.SetProperty(oWindow, "WindowState", ExcelAutomation.XlWindowState.xlMaximized);
					ExcelAutomation.NAR(oWS);
					oWS = ExcelAutomation.GetProperty(oWB, ExcelAutomation.PropertyType.ActiveSheet);
					try
					{
						ExcelAutomation.NAR(oPageSetup);
						oPageSetup = ExcelAutomation.GetProperty(oWS, "PageSetup");
						ExcelAutomation.SetProperty(oPageSetup, "PaperSize", ExcelAutomation.XlPaperSize.xlPaperA4);
						ExcelAutomation.SetProperty(oPageSetup, "Orientation", ExcelAutomation.XlPageOrientation.xlLandscape);
						ExcelAutomation.SetProperty(oPageSetup, "LeftMargin", 2);
						ExcelAutomation.SetProperty(oPageSetup, "RightMargin", 2);
						ExcelAutomation.SetProperty(oPageSetup, "TopMargin", 15);
						ExcelAutomation.SetProperty(oPageSetup, "BottomMargin", 15);
						ExcelAutomation.SetProperty(oPageSetup, "HeaderMargin", 5);
						ExcelAutomation.SetProperty(oPageSetup, "FooterMargin", 5);
						ExcelAutomation.SetProperty(oPageSetup, "CenterHorizontally", true);
						ExcelAutomation.SetProperty(oPageSetup, "CenterVertically", false);
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Exception\n" + ex.ToString());
					}
					ExcelAutomation.NAR(oStyles);
					oStyles = ExcelAutomation.GetProperty(oWB, "Styles");
					ExcelAutomation.NAR(oStyle);
					oStyle = ExcelAutomation.GetProperty(oStyles, "Item", "Normal");
					ExcelAutomation.SetProperty(oStyle, "IncludeNumber", true);
					ExcelAutomation.SetProperty(oStyle, "IncludeFont", true);
					ExcelAutomation.SetProperty(oStyle, "IncludeAlignment", true);
					ExcelAutomation.SetProperty(oStyle, "IncludeBorder", true);
					ExcelAutomation.SetProperty(oStyle, "IncludePatterns", true);
					ExcelAutomation.SetProperty(oStyle, "IncludeProtection", true);
					ExcelAutomation.SetProperty(oStyle, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
					ExcelAutomation.SetProperty(oStyle, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignTop);
					ExcelAutomation.SetProperty(oStyle, "WrapText", false);
					ExcelAutomation.SetProperty(oStyle, "IndentLevel", 0);
					ExcelAutomation.SetProperty(oStyle, "ShrinkToFit", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oStyle, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Bold", false);
					ExcelAutomation.SetProperty(oFont, "Italic", false);
					ExcelAutomation.SetProperty(oFont, "Underline", ExcelAutomation.XlUnderlineStyle.xlUnderlineStyleNone);
					ExcelAutomation.SetProperty(oFont, "Strikethrough", false);
					ExcelAutomation.SetProperty(oFont, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
					ExcelAutomation.NAR(oCells);
					oCells = ExcelAutomation.GetProperty(oWS, "Cells");
					ExcelAutomation.SetProperty(oCells, "Style", "Normal");
					ExcelAutomation.SetProperty(oCells, "NumberFormat", "@");
					ExcelAutomation.SetProperty(oWS, "DisplayPageBreaks", false);
					ExcelAutomation.SetProperty(oWS, "Name", "Comparativo");
					ExcelAutomation.SetProperty(oXL, "DisplayAlerts", true);
					ExcelAutomation.SetProperty(oXL, "UserControl", true);
					#endregion

					#region [ Verifica se o Excel suporta o tipo 'decimal' ]
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", 1, 1);
					try
					{
						ExcelAutomation.SetProperty(oCell, "Value", (decimal)0.5);
						blnExcelSuportaDecimalDataType = true;
					}
					catch (Exception)
					{
						blnExcelSuportaDecimalDataType = false;
					}
					finally
					{
						ExcelAutomation.SetProperty(oCell, "Value", null);
					}
					#endregion

					#region [ Índices que definem a posição das colunas ]
					iXlMargemEsq = 1;
					iXlNatureza = iXlMargemEsq + 1;
					iXlPlanoContas = iXlNatureza + 2;
					iXlVlPeriodo1 = iXlPlanoContas + 2;
					iXlVlPeriodo2 = iXlVlPeriodo1 + 2;
					#endregion

					#region [ Colunas que definem os limites da planilha ]
					iXlDadosMinCol = iXlMargemEsq + 1;
					if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
						iXlDadosMaxCol = iXlVlPeriodo1 + 2;
					else
						iXlDadosMaxCol = iXlVlPeriodo1 + (_saidaMensal.Count - 1) * 2;
					#endregion

					#region [ Configura largura das colunas ]
					ExcelAutomation.NAR(oColumns);
					oColumns = ExcelAutomation.GetProperty(oWS, "Columns");
					// Margem
					ExcelAutomation.NAR(oColumn);
					oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlMargemEsq, Missing.Value);
					ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
					// Natureza
					ExcelAutomation.NAR(oColumn);
					oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlNatureza, Missing.Value);
					ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 9);
					ExcelAutomation.NAR(oColumn);
					oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlNatureza + 1, Missing.Value);
					ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
					// Plano de contas
					ExcelAutomation.NAR(oColumn);
					oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlPlanoContas, Missing.Value);
					ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 40);
					ExcelAutomation.NAR(oColumn);
					oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlPlanoContas + 1, Missing.Value);
					ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
					if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
					{
						// Valor do período 1
						ExcelAutomation.NAR(oColumn);
						oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlPeriodo1, Missing.Value);
						ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 15);
						ExcelAutomation.NAR(oColumn);
						oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlPeriodo1 + 1, Missing.Value);
						ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
						// Valor do período 2
						ExcelAutomation.NAR(oColumn);
						oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlPeriodo2, Missing.Value);
						ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 15);
						ExcelAutomation.NAR(oColumn);
						oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlPeriodo2 + 1, Missing.Value);
						ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
					}
					else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
					{
						for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
						{
							iXlAux = iXlVlPeriodo1 + iMensal * 2;
							ExcelAutomation.NAR(oColumn);
							oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlAux, Missing.Value);
							ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 15);
							ExcelAutomation.NAR(oColumn);
							oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlAux + 1, Missing.Value);
							ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
						}
					}
					#endregion

					#region [ Linha usada como margem superior ]
					ExcelAutomation.NAR(oRows);
					oRows = ExcelAutomation.GetProperty(oWS, "Rows");
					ExcelAutomation.NAR(oRow);
					oRow = ExcelAutomation.GetProperty(oRows, "Item", iNumLinha, Missing.Value);
					ExcelAutomation.SetProperty(oRow, "RowHeight", 5);
					iNumLinha++;
					#endregion

					#region [ Cabeçalho do relatório ]

					ExcelAutomation.NAR(oCells);
					oCells = ExcelAutomation.GetProperty(oWS, "Cells");

					#region [ Título do relatório ]
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", 14);
					ExcelAutomation.SetProperty(oFont, "Bold", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", "Relatório Sintético Comparativo de Movimentos");
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Tipo de saída do relatório ]
					if (rbCompetenciaComparativoEntrePeriodos.Checked)
					{
						strTexto = "Tipo de saída: Comparativo entre períodos";
					}
					else
					{
						strTexto = "Tipo de saída: Comparativo mês a mês";
					}
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Data/hora da emissão ]
					strTexto = "Emissão: " + Global.formataDataDdMmYyyyHhMmComSeparador(DateTime.Now);
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Tipo de saída: comparativo entre períodos ]
					if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
					{
						#region [ Filtro: Período 1 ]
						strTexto = "Período 1: ";
						if ((_filtroPeriodo1DataCompetenciaInicial.Length > 0) && (_filtroPeriodo1DataCompetenciaFinal.Length > 0))
							strTexto += _filtroPeriodo1DataCompetenciaInicial + " a " + _filtroPeriodo1DataCompetenciaFinal;
						else if (_filtroPeriodo1DataCompetenciaInicial.Length > 0)
							strTexto += _filtroPeriodo1DataCompetenciaInicial;
						else if (_filtroPeriodo1DataCompetenciaFinal.Length > 0)
							strTexto += _filtroPeriodo1DataCompetenciaFinal;
						else strTexto += "N.I.";
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
						ExcelAutomation.SetProperty(oCell, "WrapText", false);
						ExcelAutomation.NAR(oFont);
						oFont = ExcelAutomation.GetProperty(oCell, "Font");
						ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Italic", true);
						ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
						ExcelAutomation.SetProperty(oCell, "Value", strTexto);
						#endregion

						#region [ Filtro: Período 2 ]
						strTexto = "Período 2: ";
						if ((_filtroPeriodo2DataCompetenciaInicial.Length > 0) && (_filtroPeriodo2DataCompetenciaFinal.Length > 0))
							strTexto += _filtroPeriodo2DataCompetenciaInicial + " a " + _filtroPeriodo2DataCompetenciaFinal;
						else if (_filtroPeriodo2DataCompetenciaInicial.Length > 0)
							strTexto += _filtroPeriodo2DataCompetenciaInicial;
						else if (_filtroPeriodo2DataCompetenciaFinal.Length > 0)
							strTexto += _filtroPeriodo2DataCompetenciaFinal;
						else strTexto += "N.I.";
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
						ExcelAutomation.SetProperty(oCell, "WrapText", false);
						ExcelAutomation.NAR(oFont);
						oFont = ExcelAutomation.GetProperty(oCell, "Font");
						ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Italic", true);
						ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
						ExcelAutomation.SetProperty(oCell, "Value", strTexto);
						#endregion
					}
					#endregion

					#region [ Tipo de saída: comparativo mensal ]
					if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
					{
						#region [ Filtro: Mês e ano inicial ]
						strTexto = "Mês inicial: " + Global.retornaDescricaoMesAbreviado(_filtroPeriodoInicialMesCompetencia) + "/" + _filtroPeriodoInicialAnoCompetencia;
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
						ExcelAutomation.SetProperty(oCell, "WrapText", false);
						ExcelAutomation.NAR(oFont);
						oFont = ExcelAutomation.GetProperty(oCell, "Font");
						ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Italic", true);
						ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
						ExcelAutomation.SetProperty(oCell, "Value", strTexto);
						#endregion

						#region [ Filtro: Mês e ano final ]
						strTexto = "Mês final: " + Global.retornaDescricaoMesAbreviado(_filtroPeriodoFinalMesCompetencia) + "/" + _filtroPeriodoFinalAnoCompetencia;
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
						ExcelAutomation.SetProperty(oCell, "WrapText", false);
						ExcelAutomation.NAR(oFont);
						oFont = ExcelAutomation.GetProperty(oCell, "Font");
						ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
						ExcelAutomation.SetProperty(oFont, "Italic", true);
						ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
						ExcelAutomation.SetProperty(oCell, "Value", strTexto);
						#endregion
					}
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Data de cadastro ]
					strTexto = "Cadastramento: ";
					if ((_filtroDataCadastroInicial.Length > 0) && (_filtroDataCadastroFinal.Length > 0))
						strTexto += _filtroDataCadastroInicial + " a " + _filtroDataCadastroFinal;
					else if (_filtroDataCadastroInicial.Length > 0)
						strTexto += _filtroDataCadastroInicial;
					else if (_filtroDataCadastroFinal.Length > 0)
						strTexto += _filtroDataCadastroFinal;
					else strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: Mês de competência ]
					strTexto = "Comp2: ";
					if ((_filtroMesCompetenciaInicial.Length > 0) && (_filtroMesCompetenciaFinal.Length > 0))
						strTexto += _filtroMesCompetenciaInicial + " a " + _filtroMesCompetenciaFinal;
					else if (_filtroMesCompetenciaInicial.Length > 0)
						strTexto += _filtroMesCompetenciaInicial;
					else if (_filtroMesCompetenciaFinal.Length > 0)
						strTexto += _filtroMesCompetenciaFinal;
					else strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Natureza ]
					strTexto = "Natureza: ";
					if (_filtroNatureza.Length > 0)
						strTexto += _filtroNatureza;
					else
						strTexto += "Todas";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: Incluir Atrasados ]
					strTexto = "Incluir Atrasados: ";
					if (_filtroChkIncluirAtrasados.Length > 0)
						strTexto += _filtroChkIncluirAtrasados;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Checkbox CPF ]
					strTexto = "CPF: ";
					if (_filtroChkCPF.Length > 0)
						strTexto += _filtroChkCPF;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: Checkbox CNPJ ]
					strTexto = "CNPJ: ";
					if (_filtroChkCNPJ.Length > 0)
						strTexto += _filtroChkCNPJ;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Valor ]
					strTexto = "Valor: ";
					if (_filtroValor.Length > 0)
						strTexto += _filtroValor;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: CNPJ/CPF ]
					strTexto = "CNPJ/CPF: ";
					if (_filtroCnpjCpf.Length > 0)
						strTexto += _filtroCnpjCpf;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Descrição ]
					strTexto = "Descrição: ";
					if (_filtroDescricao.Length > 0)
						strTexto += _filtroDescricao;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Conta Corrente ]
					strTexto = "Conta Corrente: ";
					if (_filtroContaCorrente.Length > 0)
						strTexto += _filtroContaCorrente;
					else
						strTexto += "Todas";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: Plano Contas Empresa ]
					strTexto = "Empresa: ";
					if (_filtroPlanoContasEmpresa.Length > 0)
						strTexto += _filtroPlanoContasEmpresa;
					else
						strTexto += "Todas";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Plano Contas Grupo ]
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
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: Plano Contas Conta ]
					strTexto = "Conta: ";
					if (_filtroPlanoContasConta.Length > 0)
						strTexto += _filtroPlanoContasConta;
					else
						strTexto += "Todos";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#region [ Filtro: Plano Contas Grupo (Inicial) ]
					strTexto = "Grupo (inicial): ";
					if (_filtroPlanoContasGrupoInicial.Length > 0)
						strTexto += _filtroPlanoContasGrupoInicial;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Filtro: Plano Contas Grupo (Final) ]
					strTexto = "Grupo (final): ";
					if (_filtroPlanoContasGrupoFinal.Length > 0)
						strTexto += _filtroPlanoContasGrupoFinal;
					else
						strTexto += "N.I.";
					ExcelAutomation.NAR(oCell);
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
					ExcelAutomation.SetProperty(oCell, "WrapText", false);
					ExcelAutomation.NAR(oFont);
					oFont = ExcelAutomation.GetProperty(oCell, "Font");
					ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
					ExcelAutomation.SetProperty(oFont, "Italic", true);
					ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
					ExcelAutomation.SetProperty(oCell, "Value", strTexto);
					#endregion

					#region [ Nova Linha ]
					iNumLinha++;
					#endregion

					#endregion

					#region [ Obtém separador decimal usado pelo Excel ]
					ExcelAutomation.NAR(oApplication);
					oApplication = ExcelAutomation.GetProperty(oXL, "Application");
					try
					{
						blnFlag = (bool)ExcelAutomation.GetProperty(oApplication, "UseSystemSeparators");
						if (blnFlag)
						{
							System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
							strExcelDecimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
							strExcelThousandsSeparator = ci.NumberFormat.NumberGroupSeparator;
						}
						else
						{
							strExcelDecimalSeparator = (string)ExcelAutomation.GetProperty(oApplication, "DecimalSeparator");
							strExcelThousandsSeparator = (string)ExcelAutomation.GetProperty(oApplication, "ThousandsSeparator");
						}

						blnExcelSuportaUseSystemSeparators = true;
					}
					catch (Exception)
					{
						blnExcelSuportaUseSystemSeparators = false;
					}

					if (!blnExcelSuportaUseSystemSeparators || (strExcelDecimalSeparator.Length == 0) || (strExcelThousandsSeparator.Length == 0))
					{
						System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
						strExcelDecimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
						strExcelThousandsSeparator = ci.NumberFormat.NumberGroupSeparator;
					}
					#endregion

					#region [ Formatação/alinhamento das colunas ]

					if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
					{
						#region [ Valor do Período 1 ]
						strAux = Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + MAX_LINHAS_EXCEL.ToString();
						ExcelAutomation.NAR(oRange);
						oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
						ExcelAutomation.SetProperty(oRange, "NumberFormat", "#,##0.00");
						ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
						#endregion

						#region [ Valor do Período 2 ]
						strAux = Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + MAX_LINHAS_EXCEL.ToString();
						ExcelAutomation.NAR(oRange);
						oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
						ExcelAutomation.SetProperty(oRange, "NumberFormat", "#,##0.00");
						ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
						#endregion
					}
					else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
					{
						for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
						{
							iXlAux = iXlVlPeriodo1 + iMensal * 2;
							strAux = Global.excel_converte_numeracao_digito_para_letra(iXlAux) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlAux) + MAX_LINHAS_EXCEL.ToString();
							ExcelAutomation.NAR(oRange);
							oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
							ExcelAutomation.SetProperty(oRange, "NumberFormat", "#,##0.00");
							ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
						}
					}
					#endregion

					#region [ Laço para listagem ]
					_blnImprimeTitulos = true;
					for (int iRow = 0; iRow < _dtbConsulta.Rows.Count; iRow++)
					{
						#region [ Mudou o grupo? ]
						if (!_blnQuebrarGrupo)
						{
							if (!_strPlanoContasGrupoAnterior.Equals(_dtbConsulta.Rows[iRow]["id_plano_contas_grupo"].ToString()))
							{
								_blnQuebrarGrupo = true;
								_blnImprimirTotalGrupo = true;
							}
						}
						#endregion

						#region [ Imprime total do grupo anterior? ]
						if (_blnImprimirTotalGrupo)
						{
							if (iRow > 0)
							{
								#region [ Imprime o total do grupo ]
								iNumLinha++;

								listaNumLinhaTotalGrupo.Add(iNumLinha);

								#region [ Texto 'Total Grupo' ]
								ExcelAutomation.NAR(oCell);
								oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPlanoContas);
								ExcelAutomation.NAR(oFont);
								oFont = ExcelAutomation.GetProperty(oCell, "Font");
								ExcelAutomation.SetProperty(oFont, "Bold", true);
								ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
								ExcelAutomation.SetProperty(oCell, "Value", "TOTAL GRUPO");
								#endregion

								if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
								{
									#region [ Valor do Período 1 ]
									ExcelAutomation.NAR(oCell);
									oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
									ExcelAutomation.NAR(oFont);
									oFont = ExcelAutomation.GetProperty(oCell, "Font");
									ExcelAutomation.SetProperty(oFont, "Bold", true);
									strTexto = "=SUM(" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + iPrimeiraLinhaDadosGrupo.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + iUltimaLinhaDadosGrupo.ToString() + ")";
									ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
									#endregion

									#region [ Valor do Período 2 ]
									ExcelAutomation.NAR(oCell);
									oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo2);
									ExcelAutomation.NAR(oFont);
									oFont = ExcelAutomation.GetProperty(oCell, "Font");
									ExcelAutomation.SetProperty(oFont, "Bold", true);
									strTexto = "=SUM(" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + iPrimeiraLinhaDadosGrupo.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + iUltimaLinhaDadosGrupo.ToString() + ")";
									ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
									#endregion
								}
								else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
								{
									for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
									{
										iXlAux = iXlVlPeriodo1 + iMensal * 2;
										ExcelAutomation.NAR(oCell);
										oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlAux);
										ExcelAutomation.NAR(oFont);
										oFont = ExcelAutomation.GetProperty(oCell, "Font");
										ExcelAutomation.SetProperty(oFont, "Bold", true);
										strTexto = "=SUM(" + Global.excel_converte_numeracao_digito_para_letra(iXlAux) + iPrimeiraLinhaDadosGrupo.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlAux) + iUltimaLinhaDadosGrupo.ToString() + ")";
										ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
									}
								}
								#endregion
							}

							iPrimeiraLinhaDadosGrupo = 0;
							iUltimaLinhaDadosGrupo = 0;
							vlSubTotalPlanoContasGrupoPeriodo1 = 0;
							vlSubTotalPlanoContasGrupoPeriodo2 = 0;
							if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
							{
								for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
								{
									_saidaMensal[iMensal].vlSubTotalPlanoContasGrupo = 0;
								}
							}
							_blnImprimirTotalGrupo = false;
						}
						#endregion

						#region [ Imprime títulos/Quebra por grupo? ]
						if (_blnImprimeTitulos || _blnQuebrarGrupo)
						{
							if (iRow == 0)
								iNumLinha++;
							else
								iNumLinha = iNumLinha + 3;

							#region [ Imprime nome do grupo ]

							#region [ Bordas ]

							#region [ Prepara objetos p/ tratamento das bordas ]
							strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinCol) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxCol) + iNumLinha.ToString();
							ExcelAutomation.NAR(oRange);
							oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
							ExcelAutomation.NAR(oBorders);
							oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
							#endregion

							#region [ Cor de fundo ]
							ExcelAutomation.NAR(oRangeInterior);
							oRangeInterior = ExcelAutomation.GetProperty(oRange, ExcelAutomation.PropertyType.Interior);
							ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
							ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
							ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.Color, ExcelAutomation.XlColor.Gray1);
							ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.TintAndShade, 0d);
							ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
							#endregion

							#region [ Borda Superior ]
							ExcelAutomation.NAR(oBorder);
							oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeTop);
							ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
							ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
							ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
							#endregion

							#region [ Borda Inferior ]
							ExcelAutomation.NAR(oBorder);
							oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
							ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
							ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
							ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
							#endregion

							#region [ Borda Esquerda ]
							ExcelAutomation.NAR(oBorder);
							oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeLeft);
							ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
							ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
							ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
							#endregion

							#region [ Borda Direita ]
							ExcelAutomation.NAR(oBorder);
							oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeRight);
							ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
							ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
							ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
							#endregion

							#endregion

							#region [ Nome do grupo de contas ]
							strTexto = " " + _dtbConsulta.Rows[iRow]["id_plano_contas_grupo"].ToString().PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_GRUPO, '0') + " - " + _dtbConsulta.Rows[iRow]["descricao_id_plano_contas_grupo"].ToString();
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
							ExcelAutomation.SetProperty(oCell, "Value", strTexto);
							#endregion

							#endregion

							#region [ Bordas dos títulos das colunas ]
							iNumLinha++;
							strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinCol) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxCol) + iNumLinha.ToString();
							ExcelAutomation.NAR(oRange);
							oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
							ExcelAutomation.NAR(oBorders);
							oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
							ExcelAutomation.NAR(oBorder);
							oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
							ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
							ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
							ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
							#endregion

							#region [ Título das colunas ]

							#region [ Natureza ]
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
							ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
							ExcelAutomation.SetProperty(oCell, "Value", "Natureza");
							#endregion

							#region [ Plano de Contas ]
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPlanoContas);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
							ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
							ExcelAutomation.SetProperty(oCell, "Value", "Plano de Contas");
							#endregion

							if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
							{
								#region [ Valor do Período 1 ]
								ExcelAutomation.NAR(oCell);
								oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
								ExcelAutomation.SetProperty(oCell, "NumberFormat", "@");
								ExcelAutomation.NAR(oFont);
								oFont = ExcelAutomation.GetProperty(oCell, "Font");
								ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
								ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
								ExcelAutomation.SetProperty(oFont, "Bold", true);
								ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
								ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
								ExcelAutomation.SetProperty(oCell, "Value", "VL Período 1");
								#endregion

								#region [ Valor do Período 2 ]
								ExcelAutomation.NAR(oCell);
								oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo2);
								ExcelAutomation.SetProperty(oCell, "NumberFormat", "@");
								ExcelAutomation.NAR(oFont);
								oFont = ExcelAutomation.GetProperty(oCell, "Font");
								ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
								ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
								ExcelAutomation.SetProperty(oFont, "Bold", true);
								ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
								ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
								ExcelAutomation.SetProperty(oCell, "Value", "VL Período 2");
								#endregion
							}
							else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
							{
								for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
								{
									iXlAux = iXlVlPeriodo1 + iMensal * 2;
									ExcelAutomation.NAR(oCell);
									oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlAux);
									ExcelAutomation.SetProperty(oCell, "NumberFormat", "@");
									ExcelAutomation.NAR(oFont);
									oFont = ExcelAutomation.GetProperty(oCell, "Font");
									ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
									ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
									ExcelAutomation.SetProperty(oFont, "Bold", true);
									ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
									ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
									strTexto = Global.retornaDescricaoMesAbreviado(_saidaMensal[iMensal].mes) + "/" + _saidaMensal[iMensal].ano.ToString();
									ExcelAutomation.SetProperty(oCell, "Value", strTexto);
								}
							}
							#endregion

							if (_blnQuebrarGrupo)
							{
								_strPlanoContasGrupoAnterior = _dtbConsulta.Rows[iRow]["id_plano_contas_grupo"].ToString();
								_blnQuebrarGrupo = false;
							}

							_blnImprimeTitulos = false;
						}
						#endregion

						iNumLinha++;

						#region [ Memoriza intervalo de linhas p/ a fórmula ]
						if (iPrimeiraLinhaDadosGrupo <= 0)
						{
							iPrimeiraLinhaDadosGrupo = iNumLinha;
						}
						iUltimaLinhaDadosGrupo = iNumLinha;
						#endregion

						#region [ Natureza ]
						strTexto = Global.retornaDescricaoFluxoCaixaNatureza(_dtbConsulta.Rows[iRow]["natureza"].ToString().ToCharArray()[0]);
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNatureza);
						ExcelAutomation.SetProperty(oCell, "Value", strTexto);
						#endregion

						#region [ Plano de Contas ]
						strTexto = _dtbConsulta.Rows[iRow]["id_plano_contas_conta"].ToString().PadLeft(Global.Cte.FIN.TamanhoCampo.PLANO_CONTAS_CONTA, '0') + " - " + _dtbConsulta.Rows[iRow]["descricao"].ToString();
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPlanoContas);
						ExcelAutomation.SetProperty(oCell, "Value", strTexto);
						#endregion

						if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
						{
							#region [ Valor Período 1 ]
							vlValor = (decimal)_dtbConsulta.Rows[iRow]["vl_total_periodo1"];
							vlTotalAcumuladoPeriodo1 += vlValor;
							vlSubTotalPlanoContasGrupoPeriodo1 += vlValor;
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
							if (blnExcelSuportaDecimalDataType)
							{
								ExcelAutomation.SetProperty(oCell, "Value", vlValor);
							}
							else
							{
								ExcelAutomation.SetProperty(oCell, "Value", (double)vlValor);
							}
							#endregion

							#region [ Valor Período 2 ]
							vlValor = (decimal)_dtbConsulta.Rows[iRow]["vl_total_periodo2"];
							vlTotalAcumuladoPeriodo2 += vlValor;
							vlSubTotalPlanoContasGrupoPeriodo2 += vlValor;
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo2);
							if (blnExcelSuportaDecimalDataType)
							{
								ExcelAutomation.SetProperty(oCell, "Value", vlValor);
							}
							else
							{
								ExcelAutomation.SetProperty(oCell, "Value", (double)vlValor);
							}
							#endregion
						}
						else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
						{
							for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
							{
								iXlAux = iXlVlPeriodo1 + iMensal * 2;
								vlValor = (decimal)_dtbConsulta.Rows[iRow][_saidaMensal[iMensal].keyName];
								_saidaMensal[iMensal].vlTotalAcumulado += vlValor;
								_saidaMensal[iMensal].vlSubTotalPlanoContasGrupo += vlValor;
								ExcelAutomation.NAR(oCell);
								oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlAux);
								if (blnExcelSuportaDecimalDataType)
								{
									ExcelAutomation.SetProperty(oCell, "Value", vlValor);
								}
								else
								{
									ExcelAutomation.SetProperty(oCell, "Value", (double)vlValor);
								}
							}
						}
					} // Laço
					#endregion

					#region [ Total do último grupo ]
					if (_dtbConsulta.Rows.Count > 0)
					{
						iNumLinha++;

						listaNumLinhaTotalGrupo.Add(iNumLinha);

						#region [ Texto 'Total Grupo' ]
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPlanoContas);
						ExcelAutomation.NAR(oFont);
						oFont = ExcelAutomation.GetProperty(oCell, "Font");
						ExcelAutomation.SetProperty(oFont, "Bold", true);
						ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
						ExcelAutomation.SetProperty(oCell, "Value", "TOTAL GRUPO");
						#endregion

						if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
						{
							#region [ Valor do Período 1 ]
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							strTexto = "=SUM(" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + iPrimeiraLinhaDadosGrupo.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + iUltimaLinhaDadosGrupo.ToString() + ")";
							ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
							#endregion

							#region [ Valor do Período 2 ]
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo2);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							strTexto = "=SUM(" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + iPrimeiraLinhaDadosGrupo.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + iUltimaLinhaDadosGrupo.ToString() + ")";
							ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
							#endregion
						}
						else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
						{
							for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
							{
								iXlAux = iXlVlPeriodo1 + iMensal * 2;
								ExcelAutomation.NAR(oCell);
								oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlAux);
								ExcelAutomation.NAR(oFont);
								oFont = ExcelAutomation.GetProperty(oCell, "Font");
								ExcelAutomation.SetProperty(oFont, "Bold", true);
								strTexto = "=SUM(" + Global.excel_converte_numeracao_digito_para_letra(iXlAux) + iPrimeiraLinhaDadosGrupo.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlAux) + iUltimaLinhaDadosGrupo.ToString() + ")";
								ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
							}
						}
					}
					#endregion

					#region [ Total geral ]
					if (_dtbConsulta.Rows.Count > 0)
					{
						iNumLinha++;
						iNumLinha++;

						#region [ Bordas ]

						#region [ Prepara objetos p/ tratamento das bordas ]
						strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinCol) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxCol) + iNumLinha.ToString();
						ExcelAutomation.NAR(oRange);
						oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
						ExcelAutomation.NAR(oBorders);
						oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
						#endregion

						#region [ Cor de fundo ]
						ExcelAutomation.NAR(oRangeInterior);
						oRangeInterior = ExcelAutomation.GetProperty(oRange, ExcelAutomation.PropertyType.Interior);
						ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.Pattern, ExcelAutomation.XlPattern.xlSolid);
						ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.PatternColorIndex, ExcelAutomation.XlPatternColorIndex.xlAutomatic);
						ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.Color, ExcelAutomation.XlColor.Gray1);
						ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.TintAndShade, 0d);
						ExcelAutomation.SetProperty(oRangeInterior, ExcelAutomation.PropertyType.PatternTintAndShade, 0d);
						#endregion

						#region [ Borda Superior ]
						ExcelAutomation.NAR(oBorder);
						oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeTop);
						ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
						ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
						ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
						#endregion

						#region [ Borda Inferior ]
						ExcelAutomation.NAR(oBorder);
						oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
						ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
						ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
						ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
						#endregion

						#region [ Borda Esquerda ]
						ExcelAutomation.NAR(oBorder);
						oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeLeft);
						ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
						ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
						ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
						#endregion

						#region [ Borda Direita ]
						ExcelAutomation.NAR(oBorder);
						oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeRight);
						ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
						ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
						ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
						#endregion

						#endregion

						#region [ Texto 'Total Geral' ]
						ExcelAutomation.NAR(oCell);
						oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPlanoContas);
						ExcelAutomation.NAR(oFont);
						oFont = ExcelAutomation.GetProperty(oCell, "Font");
						ExcelAutomation.SetProperty(oFont, "Bold", true);
						ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
						ExcelAutomation.SetProperty(oCell, "Value", "TOTAL GERAL");
						#endregion

						if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_ENTRE_PERIODOS)
						{
							#region [ Total Geral do Período 1 ]
							sbTexto = new StringBuilder("");
							for (int i = 0; i < listaNumLinhaTotalGrupo.Count; i++)
							{
								// No Excel, o separador entre as células passadas no parâmetro da fórmula é o ponto e vírgula.
								// Entretanto, o uso do ponto e vírgula na string enviada via automação causa uma exception.
								// O funcionamento esperado foi obtido usando a vírgula. Na planilha gerada, a fórmula é exibida corretamente (com ponto e vírgula ao invés da vírgula).
								if (sbTexto.Length > 0) sbTexto.Append(",");
								sbTexto.Append(Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo1) + listaNumLinhaTotalGrupo[i].ToString());
							}
							strTexto = "=SUM(" + sbTexto.ToString() + ")";
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo1);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
							#endregion

							#region [ Total Geral do Período 2 ]
							sbTexto = new StringBuilder("");
							for (int i = 0; i < listaNumLinhaTotalGrupo.Count; i++)
							{
								if (sbTexto.Length > 0) sbTexto.Append(",");
								sbTexto.Append(Global.excel_converte_numeracao_digito_para_letra(iXlVlPeriodo2) + listaNumLinhaTotalGrupo[i].ToString());
							}
							strTexto = "=SUM(" + sbTexto.ToString() + ")";
							ExcelAutomation.NAR(oCell);
							oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlPeriodo2);
							ExcelAutomation.NAR(oFont);
							oFont = ExcelAutomation.GetProperty(oCell, "Font");
							ExcelAutomation.SetProperty(oFont, "Bold", true);
							ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
							#endregion
						}
						else if (_filtroTipoSaida == eOpcaoFiltroTipoSaida.COMPARATIVO_MENSAL)
						{
							for (int iMensal = 0; iMensal < _saidaMensal.Count; iMensal++)
							{
								iXlAux = iXlVlPeriodo1 + iMensal * 2;
								sbTexto = new StringBuilder("");
								for (int i = 0; i < listaNumLinhaTotalGrupo.Count; i++)
								{
									// No Excel, o separador entre as células passadas no parâmetro da fórmula é o ponto e vírgula.
									// Entretanto, o uso do ponto e vírgula na string enviada via automação causa uma exception.
									// O funcionamento esperado foi obtido usando a vírgula. Na planilha gerada, a fórmula é exibida corretamente (com ponto e vírgula ao invés da vírgula).
									if (sbTexto.Length > 0) sbTexto.Append(",");
									sbTexto.Append(Global.excel_converte_numeracao_digito_para_letra(iXlAux) + listaNumLinhaTotalGrupo[i].ToString());
								}
								strTexto = "=SUM(" + sbTexto.ToString() + ")";
								ExcelAutomation.NAR(oCell);
								oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlAux);
								ExcelAutomation.NAR(oFont);
								oFont = ExcelAutomation.GetProperty(oCell, "Font");
								ExcelAutomation.SetProperty(oFont, "Bold", true);
								ExcelAutomation.SetProperty(oCell, "Formula", strTexto);
							}
						}
					}
					#endregion
				}
				finally
				{
					ExcelAutomation.NAR(oFont);
					ExcelAutomation.NAR(oCell);
					ExcelAutomation.NAR(oCells);
					ExcelAutomation.NAR(oBorder);
					ExcelAutomation.NAR(oBorders);
					ExcelAutomation.NAR(oRangeInterior);
					ExcelAutomation.NAR(oRange);
					ExcelAutomation.NAR(oRow);
					ExcelAutomation.NAR(oRows);
					ExcelAutomation.NAR(oColumn);
					ExcelAutomation.NAR(oColumns);
					ExcelAutomation.NAR(oStyle);
					ExcelAutomation.NAR(oStyles);
					ExcelAutomation.NAR(oPageSetup);
					ExcelAutomation.NAR(oWindow);
					ExcelAutomation.NAR(oWindows);
					ExcelAutomation.NAR(oWS);
					ExcelAutomation.NAR(oWB);
					ExcelAutomation.NAR(oWBs);
					ExcelAutomation.NAR(oApplication);
					ExcelAutomation.NAR(oXL);
					Thread.Sleep(1000);
				}

				// Feedback da conclusão da rotina
				SystemSounds.Exclamation.Play();

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

		#region [ executaConsulta ]
		private void executaConsulta()
		{
			if (!executaPesquisa()) return;

			geraPlanilhaExcel();
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
			int iAno;
			int qtdeAnosPeriodo;
			int qtdeAnosPeriodoFuturo;
			int qtdeAnosPeriodoTotal;
			#endregion

			try
			{
				limpaCampos();

				qtdeAnosPeriodo = ComumDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FIN_REL_SINT_COMP_MOVTO_COMP_MES_A_MES_PERIODO_EM_ANOS);
				if (qtdeAnosPeriodo <= 0) qtdeAnosPeriodo = 5;

				qtdeAnosPeriodoFuturo = ComumDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FIN_REL_SINT_COMP_MOVTO_COMP_MES_A_MES_FUTURO_PERIODO_EM_ANOS);
				if (qtdeAnosPeriodoFuturo < 0) qtdeAnosPeriodoFuturo = 0;

				qtdeAnosPeriodoTotal = qtdeAnosPeriodo + qtdeAnosPeriodoFuturo;

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
				cbNatureza.DataSource = Global.montaOpcaoFluxoCaixaNatureza(Global.eOpcaoIncluirItemTodos.NAO_INCLUIR);
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

				#region [ Combo mês competência (inicial) ]
				cbCompetenciaMesInicial.DataSource = Global.montaOpcaoMes(Global.eOpcaoIncluirItemTodos.NAO_INCLUIR);
				cbCompetenciaMesInicial.DisplayMember = "nome";
				cbCompetenciaMesInicial.ValueMember = "numero";
				cbCompetenciaMesInicial.SelectedIndex = -1;
				#endregion

				#region [ Combo ano competência (inicial) ]
				cbCompetenciaAnoInicial.Items.Clear();
				for (int i = 0; i < qtdeAnosPeriodoTotal; i++)
				{
					iAno = (DateTime.Today.Year + qtdeAnosPeriodoFuturo) - i;
					cbCompetenciaAnoInicial.Items.Add(new Global.OpcaoAno(iAno, iAno.ToString()));
				}
				cbCompetenciaAnoInicial.DisplayMember = "descricao";
				cbCompetenciaAnoInicial.ValueMember = "numero";
				cbCompetenciaAnoInicial.SelectedIndex = -1;
				#endregion

				#region [ Combo mês competência (final) ]
				cbCompetenciaMesFinal.DataSource = Global.montaOpcaoMes(Global.eOpcaoIncluirItemTodos.NAO_INCLUIR);
				cbCompetenciaMesFinal.DisplayMember = "nome";
				cbCompetenciaMesFinal.ValueMember = "numero";
				cbCompetenciaMesFinal.SelectedIndex = -1;
				#endregion

				#region [ Combo ano competência (final) ]
				cbCompetenciaAnoFinal.Items.Clear();
				for (int i = 0; i < qtdeAnosPeriodoTotal; i++)
				{
					iAno = (DateTime.Today.Year + qtdeAnosPeriodoFuturo) - i;
					cbCompetenciaAnoFinal.Items.Add(new Global.OpcaoAno(iAno, iAno.ToString()));
				}
				cbCompetenciaAnoFinal.DisplayMember = "descricao";
				cbCompetenciaAnoFinal.ValueMember = "numero";
				cbCompetenciaAnoFinal.SelectedIndex = -1;
				#endregion

				#region [ Campo descrição ]
				txtDescricao.MaxLength = Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO;
				#endregion

				preencheCamposDefault();

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
					txtPeriodo1DataCompetenciaInicial.Focus();
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
				executaConsulta();
				return;
			}
		}
		#endregion

		#endregion

		#region [ txtPeriodo1DataCompetenciaInicial ]

		#region [ txtPeriodo1DataCompetenciaInicial_Enter ]
		private void txtPeriodo1DataCompetenciaInicial_Enter(object sender, EventArgs e)
		{
			txtPeriodo1DataCompetenciaInicial.Select(0, txtPeriodo1DataCompetenciaInicial.Text.Length);
		}
		#endregion

		#region [ txtPeriodo1DataCompetenciaInicial_Leave ]
		private void txtPeriodo1DataCompetenciaInicial_Leave(object sender, EventArgs e)
		{
			if (txtPeriodo1DataCompetenciaInicial.Text.Length == 0) return;
			txtPeriodo1DataCompetenciaInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtPeriodo1DataCompetenciaInicial.Text);
			if (!Global.isDataOk(txtPeriodo1DataCompetenciaInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtPeriodo1DataCompetenciaInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtPeriodo1DataCompetenciaInicial_KeyDown ]
		private void txtPeriodo1DataCompetenciaInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtPeriodo1DataCompetenciaFinal);
		}
		#endregion

		#region [ txtPeriodo1DataCompetenciaInicial_KeyPress ]
		private void txtPeriodo1DataCompetenciaInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
			rbCompetenciaComparativoEntrePeriodos.Checked = true;
		}
		#endregion

		#endregion

		#region [ txtPeriodo1DataCompetenciaFinal ]

		#region [ txtPeriodo1DataCompetenciaFinal_Enter ]
		private void txtPeriodo1DataCompetenciaFinal_Enter(object sender, EventArgs e)
		{
			txtPeriodo1DataCompetenciaFinal.Select(0, txtPeriodo1DataCompetenciaFinal.Text.Length);
		}
		#endregion

		#region [ txtPeriodo1DataCompetenciaFinal_Leave ]
		private void txtPeriodo1DataCompetenciaFinal_Leave(object sender, EventArgs e)
		{
			if (txtPeriodo1DataCompetenciaFinal.Text.Length == 0) return;
			txtPeriodo1DataCompetenciaFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtPeriodo1DataCompetenciaFinal.Text);
			if (!Global.isDataOk(txtPeriodo1DataCompetenciaFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtPeriodo1DataCompetenciaFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtPeriodo1DataCompetenciaFinal_KeyDown ]
		private void txtPeriodo1DataCompetenciaFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtPeriodo2DataCompetenciaInicial);
		}
		#endregion

		#region [ txtPeriodo1DataCompetenciaFinal_KeyPress ]
		private void txtPeriodo1DataCompetenciaFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
			rbCompetenciaComparativoEntrePeriodos.Checked = true;
		}
		#endregion

		#endregion

		#region [ txtPeriodo2DataCompetenciaInicial ]

		#region [ txtPeriodo2DataCompetenciaInicial_Enter ]
		private void txtPeriodo2DataCompetenciaInicial_Enter(object sender, EventArgs e)
		{
			txtPeriodo2DataCompetenciaInicial.Select(0, txtPeriodo2DataCompetenciaInicial.Text.Length);
		}
		#endregion

		#region [ txtPeriodo2DataCompetenciaInicial_Leave ]
		private void txtPeriodo2DataCompetenciaInicial_Leave(object sender, EventArgs e)
		{
			if (txtPeriodo2DataCompetenciaInicial.Text.Length == 0) return;
			txtPeriodo2DataCompetenciaInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtPeriodo2DataCompetenciaInicial.Text);
			if (!Global.isDataOk(txtPeriodo2DataCompetenciaInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtPeriodo2DataCompetenciaInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtPeriodo2DataCompetenciaInicial_KeyDown ]
		private void txtPeriodo2DataCompetenciaInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtPeriodo2DataCompetenciaFinal);
		}
		#endregion

		#region [ txtPeriodo2DataCompetenciaInicial_KeyPress ]
		private void txtPeriodo2DataCompetenciaInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
			rbCompetenciaComparativoEntrePeriodos.Checked = true;
		}
		#endregion

		#endregion

		#region [ txtPeriodo2DataCompetenciaFinal ]

		#region [ txtPeriodo2DataCompetenciaFinal_Enter ]
		private void txtPeriodo2DataCompetenciaFinal_Enter(object sender, EventArgs e)
		{
			txtPeriodo2DataCompetenciaFinal.Select(0, txtPeriodo2DataCompetenciaFinal.Text.Length);
		}
		#endregion

		#region [ txtPeriodo2DataCompetenciaFinal_Leave ]
		private void txtPeriodo2DataCompetenciaFinal_Leave(object sender, EventArgs e)
		{
			if (txtPeriodo2DataCompetenciaFinal.Text.Length == 0) return;
			txtPeriodo2DataCompetenciaFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtPeriodo2DataCompetenciaFinal.Text);
			if (!Global.isDataOk(txtPeriodo2DataCompetenciaFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtPeriodo2DataCompetenciaFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtPeriodo2DataCompetenciaFinal_KeyDown ]
		private void txtPeriodo2DataCompetenciaFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtMesCompetenciaInicial);
		}
		#endregion

		#region [ txtPeriodo2DataCompetenciaFinal_KeyPress ]
		private void txtPeriodo2DataCompetenciaFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
			rbCompetenciaComparativoEntrePeriodos.Checked = true;
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

		#region [ Executar Consulta ]

		#region [ menuRelatorioExecutarConsulta_Click ]
		private void menuRelatorioExecutarConsulta_Click(object sender, EventArgs e)
		{
			executaConsulta();
		}
		#endregion

		#region [ btnConsultar_Click ]
		private void btnConsultar_Click(object sender, EventArgs e)
		{
			executaConsulta();
		}
		#endregion

		#endregion

		#region [ Limpar ]

		#region [ btnLimpar_Click ]
		private void btnLimpar_Click(object sender, EventArgs e)
		{
			limpaCampos();
			preencheCamposDefault();
		}
		#endregion

		#region [ menuRelatorioLimpar_Click ]
		private void menuRelatorioLimpar_Click(object sender, EventArgs e)
		{
			limpaCampos();
			preencheCamposDefault();
		}
		#endregion

		#endregion

		#region [ cbCompetenciaMesInicial ]

		#region [ cbCompetenciaMesInicial_SelectedIndexChanged ]
		private void cbCompetenciaMesInicial_SelectedIndexChanged(object sender, EventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#region [ cbCompetenciaMesInicial_MouseClick ]
		private void cbCompetenciaMesInicial_MouseClick(object sender, MouseEventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#endregion

		#region [ cbCompetenciaAnoInicial ]

		#region [ cbCompetenciaAnoInicial_SelectedIndexChanged ]
		private void cbCompetenciaAnoInicial_SelectedIndexChanged(object sender, EventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#region [ cbCompetenciaAnoInicial_MouseClick ]
		private void cbCompetenciaAnoInicial_MouseClick(object sender, MouseEventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#endregion

		#region [ cbCompetenciaMesFinal ]

		#region [ cbCompetenciaMesFinal_SelectedIndexChanged ]
		private void cbCompetenciaMesFinal_SelectedIndexChanged(object sender, EventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#region [ cbCompetenciaMesFinal_MouseClick ]
		private void cbCompetenciaMesFinal_MouseClick(object sender, MouseEventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#endregion

		#region [ cbCompetenciaAnoFinal ]

		#region [ cbCompetenciaAnoFinal_SelectedIndexChanged ]
		private void cbCompetenciaAnoFinal_SelectedIndexChanged(object sender, EventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#region [ cbCompetenciaAnoFinal_MouseClick ]
		private void cbCompetenciaAnoFinal_MouseClick(object sender, MouseEventArgs e)
		{
			rbCompetenciaComparativoMensal.Checked = true;
		}
		#endregion

		#endregion

		#endregion
	}

	#region [ Classe: RelSinteticoComparativoMensal ]
	class RelSinteticoComparativoMensal
	{
		public int mes { get; set; } = 0;
		public int ano { get; set; } = 0;
		public DateTime dtMesAno { get; set; } = DateTime.MinValue;
		public string sqlAtrasados { get; set; } = "";
		public string sqlMovtoMes { get; set; } = "";
		public string sqlWherePeriodo { get; set; } = "";
		public string keyName { get; set; } = "";
		public decimal vlSubTotalPlanoContasGrupo { get; set; } = 0;
		public decimal vlTotalAcumulado { get; set; } = 0;
		public RelSinteticoComparativoMensal(int Mes, int Ano)
		{
			mes = Mes;
			ano = Ano;
			dtMesAno = new DateTime(ano, mes, 1);
			keyName = Global.retornaDescricaoMesAbreviado(mes) + "_" + ano.ToString();
		}
	}
	#endregion
}
