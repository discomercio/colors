#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;
using System.Reflection;
#endregion

namespace Financeiro
{
	public partial class FCobrancaAdministracao : Financeiro.FModelo
	{
		#region [ Constantes ]
		const String GRID_COL_CHECK_BOX = "colCheckBox";
		const String GRID_COL_ID_CLIENTE = "colGridIdCliente";
		const String GRID_COL_NOME_CNPJ_CPF = "colGridNomeCnpjCpf";
		const String GRID_COL_QTDE_PARCELAS_EM_ATRASO = "colGridQtdeParcelasEmAtraso";
		const String GRID_COL_MAX_DIAS_EM_ATRASO = "colGridMaxDiasEmAtraso";
		const String GRID_COL_DESCRICAO_PARCELAS = "colGridDescricaoParcelas";
		const String GRID_COL_VALOR_TOTAL_EM_ATRASO = "colGridValorTotalEmAtraso";
		const String GRID_COL_NUM_PARCELA_MAIOR_ATRASO = "colGridNumParcelaMaiorAtraso";
		const String GRID_COL_VENDEDOR = "colGridVendedor";
		const String GRID_COL_INDICADOR = "colGridIndicador";
		const String GRID_COL_UF = "colGridUF";
		#endregion

		#region [ Atributos ]
		private Form _formChamador = null;
		CobrancaAdminListaClienteEmAtraso _listaClientesEmAtraso = new CobrancaAdminListaClienteEmAtraso();

		DataTable _dtbBaseClientesEmAtraso = new DataTable("dtbBaseClientesEmAtraso");
		DataTable _dtbTodasParcelasEmAtraso = new DataTable("dtbTodasParcelasEmAtraso");
		DataTable _dtbDadosPedidoTodasParcelasEmAtraso = new DataTable("dtbDadosPedidoTodasParcelasEmAtraso");
		DataView _dvTodasParcelasEmAtraso;
		DataView _dvDadosPedidoTodasParcelasEmAtraso;

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
		#endregion

		#region [ Memorização dos filtros ]
		private String _filtroSituacaoAdmCobranca;
		private String _filtroQtdeDiasAtrasoInicial;
		private String _filtroQtdeDiasAtrasoFinal;
		private String _filtroContaCorrente;
		private String _filtroPlanoContasEmpresa;
		private String _filtroPlanoContasGrupo;
		private String _filtroPlanoContasConta;
		private String _filtroEquipeVendas;
		private String _filtroVendedor;
		private String _filtroIndicador;
		private String _filtroGarantia;
		private String _filtroNomeCliente;
		private String _filtroCnpjCpf;
		private String _filtroSomentePrimeiraParcelaEmAtraso;
		#endregion

		#region [ Controle da impressão ]
		private int _intImpressaoIdxLinha = 0;
		private int _intImpressaoNumPagina = 0;
		private String _strImpressaoDataEmissao;
		private int _intQtdeTotalRegistros;
		private decimal _vlTotalRegistros;
		Impressao impressao;
		const String NOME_FONTE_DEFAULT = "Courier New";
		Font fonteTitulo;
		Font fonteListagem;
		Font fonteListagemNegrito;
		Font fonteDataEmissao;
		Font fonteNumPagina;
		Font fonteFiltros;
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

		#region [ Colunas Listagem ]
		float ixNome;
		float wxNome;
		float ixTelefone;
		float wxTelefone;
		float ixPedido;
		float wxPedido;
		float ixQtdeParcelasEmAtraso;
		float wxQtdeParcelasEmAtraso;
		float ixMaiorAtraso;
		float wxMaiorAtraso;
		float ixValorTotalEmAtraso;
		float wxValorTotalEmAtraso;
		float ixDescricaoParcelasEmAtraso;
		float wxDescricaoParcelasEmAtraso;
		float ESPACAMENTO_COLUNAS;
		#endregion

		#endregion

		#region [ Construtor ]
		public FCobrancaAdministracao(Form formChamador)
		{
			InitializeComponent();

			_formChamador = formChamador;
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			cbSituacao.SelectedIndex = -1;
			txtQtdeDiasAtrasoInicial.Text = "";
			txtQtdeDiasAtrasoFinal.Text = "";
			cbContaCorrente.SelectedIndex = -1;
			cbPlanoContasEmpresa.SelectedIndex = -1;
			cbPlanoContasGrupo.SelectedIndex = -1;
			cbPlanoContasConta.SelectedIndex = -1;
			cbEquipeVendas.SelectedIndex = -1;
			cbVendedor.SelectedIndex = -1;
			cbIndicador.SelectedIndex = -1;
			cbGarantia.SelectedIndex = -1;
			txtNomeCliente.Text = "";
			txtCnpjCpf.Text = "";
			lblTotalizacaoRegistros.Text = "";
			lblTotalizacaoValor.Text = "";
			gridDados.DataSource = null;
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ CNPJ/CPF ]
			if (txtCnpjCpf.Text.Trim().Length > 0)
			{
				if (!Global.isCnpjCpfOk(txtCnpjCpf.Text))
				{
					avisoErro("CNPJ/CPF inválido!!");
					return false;
				}
			}
			#endregion

			#region [ Atrasado Entre ]
			if ((txtQtdeDiasAtrasoInicial.Text.Length > 0) && (txtQtdeDiasAtrasoFinal.Text.Length > 0))
			{
				if (Global.converteInteiro(txtQtdeDiasAtrasoInicial.Text) > Global.converteInteiro(txtQtdeDiasAtrasoFinal.Text))
				{
					avisoErro("Filtro '" + lblTitAtrasadoEntre.Text + "' está preenchido incorretamente:\nO menor valor deve ser informado no 1º campo e o maior valor no 2º campo!!");
					return false;
				}
			}
			#endregion

			return true;
		}
		#endregion

		#region [ printerDialog ]
		private void printerDialog()
		{
			prnDialogConsulta.ShowDialog();
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

		#region [ imprimeConsulta ]
		private void imprimeConsulta()
		{
			prnDocConsulta.Print();
		}
		#endregion

		#region [ memorizaFiltrosParaImpressao ]
		/// <summary>
		/// Memoriza os parâmetros usados na última pesquisa para serem usados na impressão.
		/// </summary>
		private void memorizaFiltrosParaImpressao()
		{
			_filtroSituacaoAdmCobranca = cbSituacao.Text;
			_filtroQtdeDiasAtrasoInicial = (txtQtdeDiasAtrasoInicial.Text.Trim().Length == 0 ? "" : Global.formataInteiro((int)Global.converteInteiro(Global.digitos(txtQtdeDiasAtrasoInicial.Text))));
			_filtroQtdeDiasAtrasoFinal = (txtQtdeDiasAtrasoFinal.Text.Trim().Length == 0 ? "" : Global.formataInteiro((int)Global.converteInteiro(Global.digitos(txtQtdeDiasAtrasoFinal.Text))));
			_filtroContaCorrente = cbContaCorrente.Text;
			_filtroPlanoContasEmpresa = cbPlanoContasEmpresa.Text;
			_filtroPlanoContasGrupo = cbPlanoContasGrupo.Text;
			_filtroPlanoContasConta = cbPlanoContasConta.Text;
			_filtroEquipeVendas = cbEquipeVendas.Text;
			_filtroVendedor = cbVendedor.Text;
			_filtroIndicador = cbIndicador.Text;
			_filtroGarantia = cbGarantia.Text;
			_filtroNomeCliente = txtNomeCliente.Text;
			_filtroCnpjCpf = txtCnpjCpf.Text;
			if (ckb_somente_primeira_parcela_em_atraso.Checked)
				_filtroSomentePrimeiraParcelaEmAtraso = "Sim";
			else
				_filtroSomentePrimeiraParcelaEmAtraso = "Não";
		}
		#endregion

		#region [ montaClausulaWhereBaseClientesEmAtraso ]
		private String montaClausulaWhereBaseClientesEmAtraso()
		{
			#region [ Declarações ]
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;
			#endregion

			#region [ Critérios comuns ]
			strAux = Global.sqlMontaRestricoesClausulaWhereParaBoletosEmAtraso("tFC");
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
			#endregion

			#region [ Nome do cliente ]
			if (txtNomeCliente.Text.Trim().Length > 0)
			{
				strAux = " (tC.nome LIKE '" + txtNomeCliente.Text.Trim() + BD.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Cnpj/Cpf do cliente ]
			if (txtCnpjCpf.Text.Trim().Length > 0)
			{
				strAux = " (tC.cnpj_cpf = '" + Global.digitos(txtCnpjCpf.Text.Trim()) + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			return sbWhere.ToString();
		}
		#endregion

		#region [ montaSqlConsultaBaseClientesEmAtraso ]
		private String montaSqlConsultaBaseClientesEmAtraso()
		{
			#region [ Declarações ]
			String strSelect;
			String strFrom;
			String strWhereInterno;
			String strWhereExterno = "";
			String strSql;
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();
			
			#region [ Monta cláusula 'Where' do 'Select' interno ]
			strWhereInterno = montaClausulaWhereBaseClientesEmAtraso();
			if (strWhereInterno.Length > 0) strWhereInterno = " WHERE " + strWhereInterno;
			#endregion

			#region [ Monta cláusula 'From' do 'Select' interno ]
			strFrom = " FROM t_FIN_FLUXO_CAIXA tFC" +
							" LEFT JOIN t_CLIENTE tC" +
								" ON (tFC.id_cliente=tC.id)";
			#endregion

			#region [ Monta campos do 'Select' interno ]
			strSelect = "SELECT" +
							" tFC.id_cliente," +
							" Coalesce(tC.nome,'') AS nome_cliente," +
							" Coalesce(tC.cnpj_cpf,'') AS cnpj_cpf," +
							" Coalesce(tC.endereco,'') AS endereco," +
							" Coalesce(tC.endereco_numero,'') AS endereco_numero," +
							" Coalesce(tC.endereco_complemento,'') AS endereco_complemento," +
							" Coalesce(tC.bairro,'') AS bairro," +
							" Coalesce(tC.cidade,'') AS cidade," +
							" Coalesce(tC.uf,'') AS uf," +
							" Coalesce(tC.cep,'') AS cep," +
							" Coalesce(tC.ddd_res,'') AS ddd_res," +
							" Coalesce(tC.tel_res,'') AS tel_res," +
							" Coalesce(tC.ddd_com,'') AS ddd_com," +
							" Coalesce(tC.tel_com,'') AS tel_com," +
							" Coalesce(tC.ramal_com,'') AS ramal_com," +
							" Coalesce(tC.contato,'') AS contato," +
							" Min(tFC.dt_competencia) AS dt_competencia";
			#endregion

			#region [ Atrasado entre ]
			if (txtQtdeDiasAtrasoInicial.Text.Length > 0)
			{
				if (strWhereExterno.Length > 0) strWhereExterno += " AND";
				strWhereExterno += " (qtde_dias_em_atraso >= " + Global.digitos(txtQtdeDiasAtrasoInicial.Text) + ")";
			}

			if (txtQtdeDiasAtrasoFinal.Text.Length > 0)
			{
				if (strWhereExterno.Length > 0) strWhereExterno += " AND";
				strWhereExterno += " (qtde_dias_em_atraso <= " + Global.digitos(txtQtdeDiasAtrasoFinal.Text) + ")";
			}
			#endregion

			#region [ Monta o 'Select' interno ]
			strSql = strSelect +
					 strFrom +
					 strWhereInterno +
					 " GROUP BY" +
						" tFC.id_cliente," +
						" tC.nome," +
						" tC.cnpj_cpf," +
						" tC.endereco," +
						" tC.endereco_numero," +
						" tC.endereco_complemento," +
						" tC.bairro," +
						" tC.cidade," +
						" tC.uf," +
						" tC.cep," +
						" tC.ddd_res," +
						" tC.tel_res," +
						" tC.ddd_com," +
						" tC.tel_com," +
						" tC.ramal_com," +
						" tC.contato";
			#endregion

			#region[ Monta o 'Select' externo ]
			if (strWhereExterno.Length > 0) strWhereExterno = " WHERE" + strWhereExterno;

			strSql = "SELECT " +
						"*" +
					 " FROM " +
						"(" +
							"SELECT" +
								" id_cliente," +
								" nome_cliente," +
								" cnpj_cpf," +
								" endereco," +
								" endereco_numero," +
								" endereco_complemento," +
								" bairro," +
								" cidade," +
								" uf," +
								" cep," +
								" ddd_res," +
								" tel_res," +
								" ddd_com," +
								" tel_com," +
								" ramal_com," +
								" contato," +
								Global.sqlMontaExpressaoCalculaDiasEmAtraso(dtReferenciaLimitePagamentoEmAtraso, "", "qtde_dias_em_atraso") +
							 " FROM " +
								"(" +
									strSql +
								") t1" +
						") t2" +
					 strWhereExterno +
					 " ORDER BY" +
						" qtde_dias_em_atraso DESC," +
						" nome_cliente";
			#endregion

			return strSql;
		}
		#endregion

		#region [ montaClausulaWhereTodasParcelasEmAtraso ]
		private String montaClausulaWhereTodasParcelasEmAtraso()
		{
			#region [ Declarações ]
			String strWhere = "";
			String strAux;
			#endregion

			#region [ Critérios comuns ]
			strAux = Global.sqlMontaRestricoesClausulaWhereParaBoletosEmAtraso("tFC");
			if (strWhere.Length > 0) strWhere += " AND";
			strWhere += strAux;
			#endregion

			return strWhere;
		}
		#endregion

		#region [ montaSqlConsultaTodasParcelasEmAtraso ]
		private String montaSqlConsultaTodasParcelasEmAtraso(List<String> listaClientes)
		{
			#region [ Declarações ]
			String strSelect;
			String strFrom;
			String strWhere = "";
			String strSql;
			String strAux;
			StringBuilder sbWhere = new StringBuilder("");
			StringBuilder sbWhereCliente = new StringBuilder("");
			DateTime dtReferenciaLimitePagamentoEmAtraso;
			#endregion

			dtReferenciaLimitePagamentoEmAtraso = Global.obtemDataReferenciaLimitePagamentoEmAtraso();

			#region [ Critério fixo: parcelas em atraso ]
			strAux = Global.sqlMontaRestricoesClausulaWhereParaBoletosEmAtraso("tFC");
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
			#endregion

			#region [ Lista de clientes ]
			for (int i = 0; i < listaClientes.Count; i++)
			{
				if (sbWhereCliente.Length > 0) sbWhereCliente.Append(",");
				sbWhereCliente.Append("'" + listaClientes[i] + "'");
			}

			if (sbWhereCliente.Length > 0)
			{
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tFC.id_cliente IN (" + sbWhereCliente.ToString() + "))");
			}
			#endregion

			#region [ Monta cláusula 'Where' ]
			if (sbWhere.Length > 0) strWhere = " WHERE " + sbWhere.ToString();
			#endregion

			#region [ Monta cláusula 'From' ]
			strFrom = " FROM t_FIN_FLUXO_CAIXA tFC" +
						  " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFC.ctrl_pagto_id_parcela=tFBI.id) AND (tFC.ctrl_pagto_modulo=" + Global.Cte.FIN.CtrlPagtoModulo.BOLETO + ")";
			#endregion

			#region [ Monta cláusula 'Select' ]
			strSelect = "SELECT" +
							" tFC.id_cliente," +
							" tFC.id," +
							" tFC.id_conta_corrente," +
							" tFC.id_plano_contas_empresa," +
							" tFC.id_plano_contas_grupo," +
							" tFC.id_plano_contas_conta," +
							" tFC.dt_competencia," +
							" tFC.valor," +
							" tFC.descricao," +
							" tFC.ctrl_pagto_id_parcela," +
							" tFC.ctrl_pagto_modulo," +
							Global.sqlMontaExpressaoCalculaDiasEmAtraso(dtReferenciaLimitePagamentoEmAtraso, "tFC", "qtde_dias_em_atraso") + "," +
							" tFBI.id AS tFBI_id," +
							" tFBI.id_boleto AS tFBI_id_boleto," +
							" tFBI.num_parcela AS tFBI_num_parcela," +
							" tFBI.status AS tFBI_status";
			#endregion

			#region [ Monta SQL completo ]
			strSql = strSelect +
					 strFrom +
					 strWhere +
					 " ORDER BY" +
						" tFC.id_cliente," +
						" tFC.id";
			#endregion

			return strSql;
		}
		#endregion

		#region [ montaSqlConsultaDadosPedidoTodasParcelasEmAtraso ]
		private String montaSqlConsultaDadosPedidoTodasParcelasEmAtraso(List<String> listaClientes)
		{
			#region [ Declarações ]
			String strSelect;
			String strFrom;
			String strWhere = "";
			String strSql;
			String strAux;
			StringBuilder sbWhere = new StringBuilder("");
			StringBuilder sbWhereCliente = new StringBuilder("");
			#endregion

			#region [ Critérios comuns ]
			strAux = Global.sqlMontaRestricoesClausulaWhereParaBoletosEmAtraso("tFC");
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
			#endregion

			#region [ Lista de clientes ]
			for (int i = 0; i < listaClientes.Count; i++)
			{
				if (sbWhereCliente.Length > 0) sbWhereCliente.Append(",");
				sbWhereCliente.Append("'" + listaClientes[i] + "'");
			}

			if (sbWhereCliente.Length > 0)
			{
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(" (tP.id_cliente IN (" + sbWhereCliente.ToString() + "))");
			}
			#endregion

			#region [ Monta cláusula 'Where' ]
			if (sbWhere.Length > 0) strWhere = " WHERE " + sbWhere.ToString();
			#endregion

			#region [ Monta cláusula 'From' ]
			strFrom = " FROM t_FIN_FLUXO_CAIXA tFC" +
						" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tBIR" +
							" ON (tFC.ctrl_pagto_id_parcela=tBIR.id_boleto_item)" +
						" INNER JOIN t_PEDIDO tP" +
							" ON (tBIR.pedido=tP.pedido)" +
						" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (tP.pedido_base=t_PEDIDO__BASE.pedido)" +
						" LEFT JOIN t_EQUIPE_VENDAS_X_USUARIO tEVU" +
							" ON (tP.vendedor=tEVU.usuario)" +
						" LEFT JOIN t_EQUIPE_VENDAS tEV" +
							" ON (tEVU.id_equipe_vendas=tEV.id)" +
						" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR tOI" +
							" ON (tP.indicador=tOI.apelido)";
			#endregion

			#region [ Monta cláusula 'Select' ]
			strSelect = "SELECT" +
							" tP.id_cliente," +
							" tFC.id," +
							" tBIR.id_boleto_item," +
							" tBIR.pedido," +
							" tP.vendedor," +
							" t_PEDIDO__BASE.analise_credito," +
							" t_PEDIDO__BASE.analise_credito_data," +
							" Coalesce(tEV.id, 0) AS id_equipe_vendas," +
							" Coalesce(tEV.apelido, '') AS equipe_vendas," +
							" tP.indicador," +
							" tOI.email AS indicador_email," +
							" tP.GarantiaIndicadorStatus," +
							" (" +
								"SELECT" +
									" Coalesce(Sum(qtde*(preco_NF-preco_venda)),0)" +
								" FROM t_PEDIDO_ITEM tPI INNER JOIN t_PEDIDO tP2 ON (tPI.pedido=tP2.pedido)" +
								" WHERE" +
									" (tP2.st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
									" AND (tP2.pedido=tP.pedido)" +
							") AS vl_RA," +
							" (" +
								"SELECT" +
									" Coalesce(Sum(qtde*(preco_NF-preco_venda)),0)" +
								" FROM t_PEDIDO_ITEM_DEVOLVIDO tPID INNER JOIN t_PEDIDO tP3 ON (tPID.pedido=tP3.pedido)" +
								" WHERE" +
									" (tP3.st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
									" AND (tP3.pedido=tP.pedido)" +
							") AS vl_RA_devolucao";
			#endregion

			#region [ Monta SQL completo ]
			strSql = strSelect +
					 strFrom +
					 strWhere +
					 " ORDER BY" +
						" tP.id_cliente," +
						" tFC.id";
			#endregion

			return strSql;
		}
		#endregion

		#region [ montaDescricaoParcelas ]
		private String montaDescricaoParcelas(List<CobrancaAdminParcelaEmAtraso> listaParcelas)
		{
			#region [ Declarações ]
			String strAux;
			StringBuilder sbAux = new StringBuilder();
			CobrancaAdminParcelaEmAtraso parcelaEmAtraso;
			#endregion

			for (int i = 0; i < listaParcelas.Count; i++)
			{
				parcelaEmAtraso = listaParcelas[i];
				if (sbAux.Length > 0) sbAux.Append("\n");
				strAux = Global.formataMoeda(parcelaEmAtraso.valor).PadLeft(12, ' ');
				sbAux.Append(Global.formataDataDdMmYyyyComSeparador(parcelaEmAtraso.dt_competencia) + ' ' + strAux);
			}
			return sbAux.ToString();
		}
		#endregion

		#region [ montaDescricaoDadosPedido ]
		private String montaDescricaoDadosPedido(List<CobrancaAdminDadosPedidoParcelaEmAtraso> listaDadosPedido)
		{
			#region [ Declarações ]
			StringBuilder sbAux = new StringBuilder();
			CobrancaAdminDadosPedidoParcelaEmAtraso dadosPedido;
			#endregion

			for (int i = 0; i < listaDadosPedido.Count; i++)
			{
				dadosPedido = listaDadosPedido[i];
				if (sbAux.Length > 0) sbAux.Append("\n");
				sbAux.Append(dadosPedido.pedido);
				sbAux.Append(" (Vend: " + dadosPedido.vendedor.ToUpper());
				if (dadosPedido.indicador.Trim().Length > 0)
				{
					sbAux.Append("; Ind: " + dadosPedido.indicador.ToUpper());
				}
				if (dadosPedido.garantiaIndicadorStatus != Global.Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO_DEFINIDO)
				{
					sbAux.Append("; Gar: ");
					if (dadosPedido.garantiaIndicadorStatus == Global.Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM)
					{
						sbAux.Append("sim");
					}
					else if (dadosPedido.garantiaIndicadorStatus == Global.Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO)
					{
						sbAux.Append("não");
					}
				}
				if (dadosPedido.equipe_vendas.Trim().Length > 0)
				{
					sbAux.Append("; Eq: " + dadosPedido.equipe_vendas.ToUpper());
				}
				sbAux.Append(")");
			}
			return sbAux.ToString();
		}
		#endregion

		#region [ montaDescricaoDadosPedidoComValorRA ]
		private String montaDescricaoDadosPedidoComValorRA(List<CobrancaAdminDadosPedidoParcelaEmAtraso> listaDadosPedido)
		{
			#region [ Declarações ]
			StringBuilder sbAux = new StringBuilder();
			CobrancaAdminDadosPedidoParcelaEmAtraso dadosPedido;
			#endregion

			for (int i = 0; i < listaDadosPedido.Count; i++)
			{
				dadosPedido = listaDadosPedido[i];
				if (sbAux.Length > 0) sbAux.Append("\n");
				sbAux.Append(dadosPedido.pedido);
				sbAux.Append(" (Vend: " + dadosPedido.vendedor.ToUpper());
				if (dadosPedido.indicador.Trim().Length > 0)
				{
					sbAux.Append("; Ind: " + dadosPedido.indicador.ToUpper());
				}
				
				if (dadosPedido.garantiaIndicadorStatus != Global.Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO_DEFINIDO)
				{
					sbAux.Append("; Gar: ");
					if (dadosPedido.garantiaIndicadorStatus == Global.Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM)
					{
						sbAux.Append("sim");
					}
					else if (dadosPedido.garantiaIndicadorStatus == Global.Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO)
					{
						sbAux.Append("não");
					}
				}
				
				if (dadosPedido.equipe_vendas.Trim().Length > 0)
				{
					sbAux.Append("; Eq: " + dadosPedido.equipe_vendas.ToUpper());
				}
				
				if (dadosPedido.vl_RA > 0)
				{
					sbAux.Append("; RA bruto: " + Global.formataMoeda(dadosPedido.vl_RA));
				}

				sbAux.Append(")");
			}
			return sbAux.ToString();
		}
		#endregion

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			bool blnAchou;
			bool blnTemSerieBoletoComSomentePrimeiraParcelaEmAtraso;
			Decimal decTotalizacaoValor = 0;
			Decimal vl_RA;
			int n;
			int intQtdeRegistros = 0;
			CobrancaAdminRegistroClienteEmAtraso clienteEmAtraso;
			String s;
			String strSql;
			CobrancaAdminParcelaEmAtraso parcelaEmAtraso;
			CobrancaAdminDadosPedidoParcelaEmAtraso dadosPedido;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataRow rowCliente;
			List<String> listaIdBoleto = new List<String>();
			List<String> listaClienteIdEmAtraso = new List<String>();
			#endregion

			try
			{
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

				#region [ Consistência dos parâmetros ]
				btnDummy.Focus();
				if (!consisteCampos()) return false;
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				#endregion

				#region [ Prepara data adapter ]
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.SelectCommand = cmCommand;
				#endregion

				#region [ Executa a consulta no BD ]

				#region [ Relação dos clientes em atraso ]
				_dtbBaseClientesEmAtraso.Reset();
				strSql = montaSqlConsultaBaseClientesEmAtraso();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(_dtbBaseClientesEmAtraso);
				#endregion

				#region [ Obtém lista c/ ID's dos clientes em atraso ]
				for (int i = 0; i < _dtbBaseClientesEmAtraso.Rows.Count; i++)
				{
					listaClienteIdEmAtraso.Add(BD.readToString(_dtbBaseClientesEmAtraso.Rows[i]["id_cliente"]));
				}
				#endregion

				#region [ Dados de todas as parcelas em atraso ]
				_dtbTodasParcelasEmAtraso.Reset();
				_dvTodasParcelasEmAtraso = new DataView();
				strSql = montaSqlConsultaTodasParcelasEmAtraso(listaClienteIdEmAtraso);
				cmCommand.CommandText = strSql;
				daAdapter.Fill(_dtbTodasParcelasEmAtraso);
				_dvTodasParcelasEmAtraso.Table = _dtbTodasParcelasEmAtraso;
				_dvTodasParcelasEmAtraso.Sort = "id_cliente";
				#endregion

				#region [ Dados dos pedidos de todas as parcelas em atraso ]
				_dtbDadosPedidoTodasParcelasEmAtraso.Reset();
				_dvDadosPedidoTodasParcelasEmAtraso = new DataView();
				strSql = montaSqlConsultaDadosPedidoTodasParcelasEmAtraso(listaClienteIdEmAtraso);
				cmCommand.CommandText = strSql;
				daAdapter.Fill(_dtbDadosPedidoTodasParcelasEmAtraso);
				_dvDadosPedidoTodasParcelasEmAtraso.Table = _dtbDadosPedidoTodasParcelasEmAtraso;
				_dvDadosPedidoTodasParcelasEmAtraso.Sort = "id_cliente";
				#endregion

				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "processando dados");
				
				#region [ Limpa dados anteriores ]
				for (int i = (_listaClientesEmAtraso.contaClientes() - 1); i >= 0; i--)
				{
					clienteEmAtraso = _listaClientesEmAtraso.getClienteEmAtraso(i);
					clienteEmAtraso.listaParcelas.Clear();
				}
				_listaClientesEmAtraso.Clear();
				#endregion

				#region [ Carrega dados em lista na memória ]
				for (int i = 0; i < _dtbBaseClientesEmAtraso.Rows.Count; i++)
				{
					#region [ DoEvents ]
					Application.DoEvents();
					#endregion

					rowCliente = _dtbBaseClientesEmAtraso.Rows[i];

					#region [ Verifica filtros antes de adicionar este cliente à lista ]

					#region [ Conta Corrente ]
					if ((cbContaCorrente.SelectedIndex > -1) && (Global.converteInteiro(cbContaCorrente.SelectedValue.ToString()) > 0))
					{
						_dvTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
															 " AND (id_conta_corrente = " + cbContaCorrente.SelectedValue.ToString() + ")";
						if (_dvTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Plano de Contas Empresa ]
					if ((cbPlanoContasEmpresa.SelectedIndex > -1) && (Global.converteInteiro(cbPlanoContasEmpresa.SelectedValue.ToString()) > 0))
					{
						_dvTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
															 " AND (id_plano_contas_empresa = " + cbPlanoContasEmpresa.SelectedValue.ToString() + ")";
						if (_dvTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Plano Contas Grupo ]
					if ((cbPlanoContasGrupo.SelectedIndex > -1) && (Global.converteInteiro(cbPlanoContasGrupo.SelectedValue.ToString()) > 0))
					{
						_dvTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
															 " AND (id_plano_contas_grupo = " + cbPlanoContasGrupo.SelectedValue.ToString() + ")";
						if (_dvTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Plano Contas Conta ]
					if ((cbPlanoContasConta.SelectedIndex > -1) && (Global.converteInteiro(cbPlanoContasConta.SelectedValue.ToString()) > 0))
					{
						_dvTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
															 " AND (id_plano_contas_conta = " + cbPlanoContasConta.SelectedValue.ToString() + ")";
						if (_dvTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Equipe de Vendas ]
					if ((cbEquipeVendas.SelectedIndex > -1) && (Global.converteInteiro(cbEquipeVendas.SelectedValue.ToString()) > 0))
					{
						_dvDadosPedidoTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
																		" AND (id_equipe_vendas = " + cbEquipeVendas.SelectedValue.ToString() + ")";
						if (_dvDadosPedidoTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Vendedor ]
					if ((cbVendedor.SelectedIndex > -1) && (cbVendedor.SelectedValue.ToString().Trim().Length > 0))
					{
						_dvDadosPedidoTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
																		" AND (vendedor = '" + cbVendedor.SelectedValue.ToString().Trim() + "')";
						if (_dvDadosPedidoTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Indicador ]
					if ((cbIndicador.SelectedIndex > -1) && (cbIndicador.SelectedValue.ToString().Trim().Length > 0))
					{
						_dvDadosPedidoTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
																		" AND (indicador = '" + cbIndicador.SelectedValue.ToString().Trim() + "')";
						if (_dvDadosPedidoTodasParcelasEmAtraso.Count == 0) continue;
					}
					#endregion

					#region [ Garantia ]
					if ((cbGarantia.SelectedIndex > -1) && (Global.converteInteiro(cbGarantia.SelectedValue.ToString()) != Global.Cte.Etc.FLAG_NAO_SETADO))
					{
						_dvDadosPedidoTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "')" +
																		" AND (GarantiaIndicadorStatus = " + cbGarantia.SelectedValue.ToString() + ")";
						if (_dvDadosPedidoTodasParcelasEmAtraso.Count == 0) continue;
					}					
					#endregion

					#region [ Somente 1ª parcela ]
					if (ckb_somente_primeira_parcela_em_atraso.Checked)
					{
						_dvTodasParcelasEmAtraso.RowFilter = "id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "'";

						#region [ Dos boletos em atraso, obtém a lista de Id de cada série de boletos ]
						listaIdBoleto.Clear();
						for (int j = 0; j < _dvTodasParcelasEmAtraso.Count; j++)
						{
							blnAchou = false;
							for (int k = 0; k < listaIdBoleto.Count; k++)
							{
								if (BD.readToByte(_dvTodasParcelasEmAtraso[j].Row["ctrl_pagto_modulo"]) == Global.Cte.FIN.CtrlPagtoModulo.BOLETO)
								{
									if (listaIdBoleto[k].Equals(BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["tFBI_id_boleto"]).ToString()))
									{
										blnAchou = true;
										break;
									}
								}
							}
							if (!blnAchou)
							{
								listaIdBoleto.Add(BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["tFBI_id_boleto"]).ToString());
							}
						}
						#endregion

						#region [ Cliente possui alguma série de boletos em que somente a 1ª parcela esteja em atraso? ]
						blnTemSerieBoletoComSomentePrimeiraParcelaEmAtraso = false;

						for (int j = 0; j < listaIdBoleto.Count; j++)
						{
							_dvTodasParcelasEmAtraso.RowFilter = "(id_cliente = '" + BD.readToString(rowCliente["id_cliente"]) + "') AND (tFBI_id_boleto = " + listaIdBoleto[j] + ")";
							if (_dvTodasParcelasEmAtraso.Count == 1)
							{
								if (BD.readToByte(_dvTodasParcelasEmAtraso[0].Row["ctrl_pagto_modulo"]) == Global.Cte.FIN.CtrlPagtoModulo.BOLETO)
								{
									if ((BD.readToInt16(_dvTodasParcelasEmAtraso[0].Row["tFBI_status"]) != Global.Cte.FIN.CodBoletoItemStatus.BOLETO_BAIXADO) &&
										(BD.readToByte(_dvTodasParcelasEmAtraso[0].Row["tFBI_num_parcela"]) == 1))
									{
										blnTemSerieBoletoComSomentePrimeiraParcelaEmAtraso = true;
										break;
									}
								}
							}
						}

						if (!blnTemSerieBoletoComSomentePrimeiraParcelaEmAtraso) continue;
						#endregion
					}
					#endregion

					#endregion

					#region [ Cliente satisfaz aos critérios: adiciona na lista ]
					if (_listaClientesEmAtraso.isIdClienteExiste(BD.readToString(rowCliente["id_cliente"])))
					{
						throw new Exception("Cliente já consta na lista de clientes em atraso!\nNome: " + BD.readToString(rowCliente["nome_cliente"]) + "\nCNPJ/CPF: " + Global.formataCnpjCpf(BD.readToString(rowCliente["cnpj_cpf"])) + "\nID: " + BD.readToString(rowCliente["id_cliente"]));
					}
					clienteEmAtraso = _listaClientesEmAtraso.adicionaClienteEmAtraso(
														BD.readToString(rowCliente["id_cliente"]),
														BD.readToString(rowCliente["cnpj_cpf"]),
														BD.readToString(rowCliente["nome_cliente"]),
														BD.readToInt(rowCliente["qtde_dias_em_atraso"]),
														BD.readToString(rowCliente["ddd_res"]),
														BD.readToString(rowCliente["tel_res"]),
														BD.readToString(rowCliente["ddd_com"]),
														BD.readToString(rowCliente["tel_com"]),
														BD.readToString(rowCliente["ramal_com"]),
														BD.readToString(rowCliente["contato"]),
														BD.readToString(rowCliente["uf"]));
					#endregion

					#region [ Adiciona na lista todas as parcelas em atraso deste cliente ]
					_dvTodasParcelasEmAtraso.RowFilter = "id_cliente = '" + clienteEmAtraso.id_cliente + "'";
					for (int j = 0; j < _dvTodasParcelasEmAtraso.Count; j++)
					{
						parcelaEmAtraso = new CobrancaAdminParcelaEmAtraso(
													BD.readToString(_dvTodasParcelasEmAtraso[j].Row["id_cliente"]),
													BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["id"]),
													BD.readToByte(_dvTodasParcelasEmAtraso[j].Row["id_conta_corrente"]),
													BD.readToByte(_dvTodasParcelasEmAtraso[j].Row["id_plano_contas_empresa"]),
													BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["id_plano_contas_grupo"]),
													BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["id_plano_contas_conta"]),
													BD.readToDateTime(_dvTodasParcelasEmAtraso[j].Row["dt_competencia"]),
													BD.readToDecimal(_dvTodasParcelasEmAtraso[j].Row["valor"]),
													BD.readToString(_dvTodasParcelasEmAtraso[j].Row["descricao"]),
													BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["ctrl_pagto_id_parcela"]),
													BD.readToByte(_dvTodasParcelasEmAtraso[j].Row["ctrl_pagto_modulo"]),
													BD.readToInt(_dvTodasParcelasEmAtraso[j].Row["qtde_dias_em_atraso"]),
													BD.readToInt16(_dvTodasParcelasEmAtraso[j].Row["tFBI_status"]),
													BD.readToByte(_dvTodasParcelasEmAtraso[j].Row["tFBI_num_parcela"]));
						clienteEmAtraso.adicionaParcelaEmAtraso(parcelaEmAtraso);
					}
					#endregion

					#region [ Adiciona na lista todos os pedidos referentes às parcelas em atraso deste cliente ]
					_dvDadosPedidoTodasParcelasEmAtraso.RowFilter = "id_cliente = '" + clienteEmAtraso.id_cliente + "'";
					for (int j = 0; j < _dvDadosPedidoTodasParcelasEmAtraso.Count; j++)
					{
						if (!clienteEmAtraso.isPedidoExiste(BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["pedido"])))
						{
							vl_RA = BD.readToDecimal(_dvDadosPedidoTodasParcelasEmAtraso[j]["vl_RA"]) - BD.readToDecimal(_dvDadosPedidoTodasParcelasEmAtraso[j]["vl_RA_devolucao"]);
							dadosPedido = new CobrancaAdminDadosPedidoParcelaEmAtraso(
														BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["id_cliente"]),
														BD.readToInt(_dvDadosPedidoTodasParcelasEmAtraso[j]["id"]),
														BD.readToInt(_dvDadosPedidoTodasParcelasEmAtraso[j]["id_boleto_item"]),
														BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["pedido"]),
														BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["vendedor"]),
														BD.readToInt(_dvDadosPedidoTodasParcelasEmAtraso[j]["id_equipe_vendas"]),
														BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["equipe_vendas"]),
														BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["indicador"]),
														BD.readToString(_dvDadosPedidoTodasParcelasEmAtraso[j]["indicador_email"]),
														BD.readToByte(_dvDadosPedidoTodasParcelasEmAtraso[j]["GarantiaIndicadorStatus"]),
														vl_RA,
														BD.readToInt(_dvDadosPedidoTodasParcelasEmAtraso[j]["analise_credito"]),
														BD.readToDateTime(_dvDadosPedidoTodasParcelasEmAtraso[j]["analise_credito_data"]));
							clienteEmAtraso.adicionaPedido(dadosPedido);
						}
					}
					#endregion
				}
				#endregion

				#region [ Carrega dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					gridDados.SuspendLayout();

					gridDados.Rows.Clear();
					if (_listaClientesEmAtraso.contaClientes() > 0) gridDados.Rows.Add(_listaClientesEmAtraso.contaClientes());

					for (int i = 0; i < _listaClientesEmAtraso.contaClientes(); i++)
					{
						clienteEmAtraso = _listaClientesEmAtraso.getClienteEmAtraso(i);
						gridDados.Rows[i].Cells[GRID_COL_ID_CLIENTE].Value = clienteEmAtraso.id_cliente;
						gridDados.Rows[i].Cells[GRID_COL_NOME_CNPJ_CPF].Value = Texto.iniciaisEmMaiusculas(clienteEmAtraso.nome_cliente) + " (" + Global.formataCnpjCpf(clienteEmAtraso.cnpj_cpf) + ")";
						gridDados.Rows[i].Cells[GRID_COL_QTDE_PARCELAS_EM_ATRASO].Value = clienteEmAtraso.qtde_parcelas_em_atraso.ToString().PadLeft(2, '0');
						gridDados.Rows[i].Cells[GRID_COL_MAX_DIAS_EM_ATRASO].Value = Global.formataInteiro(clienteEmAtraso.max_qtde_dias_em_atraso).PadLeft(2, '0');
						gridDados.Rows[i].Cells[GRID_COL_VALOR_TOTAL_EM_ATRASO].Value = Global.formataMoeda(clienteEmAtraso.vl_total_em_atraso);
						gridDados.Rows[i].Cells[GRID_COL_DESCRICAO_PARCELAS].Value = montaDescricaoParcelas(clienteEmAtraso.listaParcelas);
						n = clienteEmAtraso.getNumeroParcelaMaiorAtraso();
						s = (n == 0 ? "" : n.ToString() + "º");
						gridDados.Rows[i].Cells[GRID_COL_NUM_PARCELA_MAIOR_ATRASO].Value = s;
						gridDados.Rows[i].Cells[GRID_COL_VENDEDOR].Value = Texto.iniciaisEmMaiusculas(clienteEmAtraso.getVendedor());
						gridDados.Rows[i].Cells[GRID_COL_INDICADOR].Value = Texto.iniciaisEmMaiusculas(clienteEmAtraso.getIndicador());
						gridDados.Rows[i].Cells[GRID_COL_UF].Value = clienteEmAtraso.uf;
						decTotalizacaoValor += clienteEmAtraso.vl_total_em_atraso;
						intQtdeRegistros++;
					}

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

				gridDados.Focus();

				memorizaFiltrosParaImpressao();

				// Feedback da conclusão da pesquisa
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

		#region [ geraPlanilhaExcel ]
		private bool geraPlanilhaExcel()
		{
			#region [ Declarações ]
			const int MAX_LINHAS_EXCEL = 65536;
			const String FN_LISTAGEM = "Arial";
			const int FS_LISTAGEM = 8;
			const int FS_CABECALHO = 8;
			bool blnExcelSuportaUseSystemSeparators = false;
			bool blnExcelSuportaDecimalDataType = false;
			bool blnFlag;
			String strMsg;
			String strAux;
			String strTelefone;
			String strExcelDecimalSeparator = "";
			String strExcelThousandsSeparator = "";
			String strTexto;
			int intQtdeRegistros = 0;
			int intPrimeiraLinhaDados = 0;
			int intUltimaLinhaDados = 0;
			int iNumLinha = 1;
			int iOffSetArray = 2;
			int iXlDadosMinIndex;
			int iXlDadosMaxIndex;
			int iXlMargemEsq;
			int iXlCliente;
			int iXlTelefone;
			int iXlPedido;
			int iXlParcelasEmAtraso;
			int iXlMaiorAtraso;
			int iXlValorRA;
			int iXlVlTotalEmAtraso;
			int iXlDescricaoParcelasEmAtraso;
			int iXlNumParcelaMaiorAtraso;
			int iXlDtCreditoOk;
			int iXlVendedor;
			int iXlParceiro;
			int iXlParceiroEmail;
			int iXlUF;
			decimal vl_total_RA;
			DateTime dtCreditoOk;
			object oXL;
			object oWBs;
			object oWB;
			object oWS;
			object oWindow;
			object oWindows;
			object oPageSetup;
			object oStyles;
			object oStyle;
			object oFont;
			object oBorders;
			object oBorder;
			object oCells;
			object oCell;
			object oColumns;
			object oColumn;
			object oRows;
			object oRow;
			object oRange;
			object oApplication;
			String[] vDados;
			CobrancaAdminRegistroClienteEmAtraso clienteEmAtraso;
			#endregion

			try
			{
				#region [ Consistência ]
				if (_listaClientesEmAtraso == null)
				{
					return false;
				}

				if (_listaClientesEmAtraso.contaClientes() == 0)
				{
					avisoErro("Não há dados!!");
					return false;
				}
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "gerando planilha Excel");

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
				oWBs = ExcelAutomation.GetProperty(oXL, "Workbooks");
				oWB = ExcelAutomation.InvokeMethod(oWBs, "Add", Missing.Value);
				oWindows = ExcelAutomation.GetProperty(oWB, "Windows");
				oWindow = ExcelAutomation.GetProperty(oWindows, "Item", 1);
				ExcelAutomation.SetProperty(oWindow, "DisplayGridlines", false);
				ExcelAutomation.SetProperty(oWindow, "DisplayHeadings", true);
				ExcelAutomation.SetProperty(oWindow, "WindowState", ExcelAutomation.XlWindowState.xlMaximized);
				oWS = ExcelAutomation.GetProperty(oWB, "ActiveSheet");
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
				oStyles = ExcelAutomation.GetProperty(oWB, "Styles");
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
				oFont = ExcelAutomation.GetProperty(oStyle, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Bold", false);
				ExcelAutomation.SetProperty(oFont, "Italic", false);
				ExcelAutomation.SetProperty(oFont, "Underline", ExcelAutomation.XlUnderlineStyle.xlUnderlineStyleNone);
				ExcelAutomation.SetProperty(oFont, "Strikethrough", false);
				ExcelAutomation.SetProperty(oFont, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				oCells = ExcelAutomation.GetProperty(oWS, "Cells");
				ExcelAutomation.SetProperty(oCells, "Style", "Normal");
				ExcelAutomation.SetProperty(oCells, "NumberFormat", "@");
				ExcelAutomation.SetProperty(oWS, "DisplayPageBreaks", false);
				ExcelAutomation.SetProperty(oWS, "Name", "Carteira_em_Atraso");
				ExcelAutomation.SetProperty(oXL, "DisplayAlerts", true);
				ExcelAutomation.SetProperty(oXL, "UserControl", true);
				#endregion

				#region [ Verifica se o Excel suporta o tipo 'decimal' ]
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
				iXlCliente = iXlMargemEsq + 1;
				iXlTelefone = iXlCliente + 2;
				iXlPedido = iXlTelefone + 2;
				iXlParcelasEmAtraso = iXlPedido + 2;
				iXlMaiorAtraso = iXlParcelasEmAtraso + 2;
				iXlValorRA = iXlMaiorAtraso + 2;
				iXlVlTotalEmAtraso = iXlValorRA + 2;
				iXlDescricaoParcelasEmAtraso = iXlVlTotalEmAtraso + 2;
				iXlNumParcelaMaiorAtraso = iXlDescricaoParcelasEmAtraso + 2;
				iXlDtCreditoOk = iXlNumParcelaMaiorAtraso + 2;
				iXlVendedor = iXlDtCreditoOk + 2;
				iXlParceiro = iXlVendedor + 2;
				iXlParceiroEmail = iXlParceiro + 2;
				iXlUF = iXlParceiroEmail + 2;
				#endregion

				#region [ Array usado p/ transferir dados p/ o Excel ]
				iXlDadosMinIndex = iXlMargemEsq + 1;
				iXlDadosMaxIndex = iXlUF;
				vDados = new string[(iXlDadosMaxIndex - iXlDadosMinIndex + 1)];
				#endregion

				#region [ Configura largura das colunas ]
				oColumns = ExcelAutomation.GetProperty(oWS, "Columns");
				// Margem
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlMargemEsq, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Cliente
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlCliente, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 23);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlCliente + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Telefone
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlTelefone, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 19);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlTelefone + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Pedido
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlPedido, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 26);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlPedido + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.2);
				// Parcelas em atraso
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlParcelasEmAtraso, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 8);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlParcelasEmAtraso + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.2);
				// Maior atraso
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlMaiorAtraso, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 7);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlMaiorAtraso + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.2);
				// Valor RA
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlValorRA, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 9);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlValorRA + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Valor total em atraso
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlTotalEmAtraso, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 11);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVlTotalEmAtraso + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Descrição das parcelas em atraso
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDescricaoParcelasEmAtraso, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 28);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDescricaoParcelasEmAtraso + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.2);
				// Nº parcela mais antiga em atraso
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlNumParcelaMaiorAtraso, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 7);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlNumParcelaMaiorAtraso + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.2);
				// Data Crédito Ok
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtCreditoOk, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 9);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlDtCreditoOk + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Vendedor
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVendedor, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 11);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlVendedor + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Parceiro
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlParceiro, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 11);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlParceiro + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// Email do Parceiro
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlParceiroEmail, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 20);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlParceiroEmail + 1, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 0.5);
				// UF
				oColumn = ExcelAutomation.GetProperty(oColumns, "Item", iXlUF, Missing.Value);
				ExcelAutomation.SetProperty(oColumn, "ColumnWidth", 3);
				ExcelAutomation.SetProperty(oColumn, "WrapText", true);
				#endregion

				#region [ Linha usada como margem superior ]
				oRows = ExcelAutomation.GetProperty(oWS, "Rows");
				oRow = ExcelAutomation.GetProperty(oRows, "Item", iNumLinha, Missing.Value);
				ExcelAutomation.SetProperty(oRow, "RowHeight", 5);
				iNumLinha++;
				#endregion

				#region [ Cabeçalho do relatório ]

				oCells = ExcelAutomation.GetProperty(oWS, "Cells");

				#region [ Título do relatório ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", 14);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "CARTEIRA EM ATRASO");
				#endregion

				#region [ Filtro: Somente 1ª parcela em atraso ]
				strTexto = "Somente 1ª parcela em atraso: " + _filtroSomentePrimeiraParcelaEmAtraso;
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPedido);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Data/hora da emissão ]
				strTexto = "Emissão: " + Global.formataDataDdMmYyyyHhMmComSeparador(DateTime.Now);
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricaoParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
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

				#region [ Filtro: Situação ]
				strTexto = "Situação: ";
				strTexto += (_filtroSituacaoAdmCobranca.Length > 0) ? _filtroSituacaoAdmCobranca : "Todas";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Atrasado entre ]
				strTexto = "Atrasado entre: ";
				strTexto += (_filtroQtdeDiasAtrasoInicial.Length > 0) ? _filtroQtdeDiasAtrasoInicial : "N.I.";
				strTexto += " e ";
				strTexto += (_filtroQtdeDiasAtrasoFinal.Length > 0) ? _filtroQtdeDiasAtrasoFinal : "N.I.";
				strTexto += " dias";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPedido);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Conta Corrente ]
				strTexto = "Conta Corrente: ";
				strTexto += (_filtroContaCorrente.Length > 0) ? _filtroContaCorrente : "Todas";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricaoParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
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

				#region [ Filtro: Plano Contas Empresa ]
				strTexto = "Empresa: ";
				strTexto += (_filtroPlanoContasEmpresa.Length > 0) ? _filtroPlanoContasEmpresa : "Todas";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Plano Contas Grupo ]
				strTexto = "Grupo: ";
				strTexto += (_filtroPlanoContasGrupo.Length > 0) ? _filtroPlanoContasGrupo : "Todos";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPedido);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Plano Contas Conta ]
				strTexto = "Conta: ";
				strTexto += (_filtroPlanoContasConta.Length > 0) ? _filtroPlanoContasConta : "Todos";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricaoParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
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

				#region [ Filtro: Equipe de Vendas ]
				strTexto = "Equipe de Vendas: ";
				strTexto += (_filtroEquipeVendas.Length > 0) ? _filtroEquipeVendas : "Todas";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Vendedor ]
				strTexto = "Vendedor: ";
				strTexto += (_filtroVendedor.Length > 0) ? _filtroVendedor : "Todos";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPedido);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Indicador ]
				strTexto = "Indicador: ";
				strTexto += (_filtroIndicador.Length > 0) ? _filtroIndicador : "Todos";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricaoParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
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

				#region [ Filtro: Garantia ]
				strTexto = "Garantia: ";
				strTexto += (_filtroGarantia.Length > 0) ? _filtroGarantia : "N.I.";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: Nome Cliente ]
				strTexto = "Nome Cliente: ";
				strTexto += (_filtroNomeCliente.Length > 0) ? _filtroNomeCliente : "N.I.";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPedido);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#region [ Filtro: CNPJ/CPF ]
				strTexto = "CNPJ/CPF: ";
				strTexto += (_filtroCnpjCpf.Length > 0) ? Global.formataCnpjCpf(_filtroCnpjCpf) : "N.I.";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricaoParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Italic", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", strTexto);
				#endregion

				#endregion

				#region [ Bordas dos títulos das colunas ]
				iNumLinha++;
				oRow = ExcelAutomation.GetProperty(oRows, "Item", iNumLinha, Missing.Value);
				ExcelAutomation.SetProperty(oRow, "RowHeight", 4);
				iNumLinha++;
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
				oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeTop);
				ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
				ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
				ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
				ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
				ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
				ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				#endregion

				#region [ Título das colunas ]

				#region [ Cliente ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Cliente");
				#endregion

				#region [ Telefone ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlTelefone);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Telefone");
				#endregion

				#region [ Pedido ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlPedido);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Pedido");
				#endregion

				#region [ Parcelas em Atraso ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Parc em Atraso");
				#endregion

				#region [ Maior Atraso ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlMaiorAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Maior Atraso");
				#endregion

				#region [ Valor RA ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlValorRA);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "RA Bruto");
				#endregion

				#region [ Valor Total em Atraso ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlTotalEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Valor Total em Atraso");
				#endregion

				#region [ Descrição das Parcelas em Atraso ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDescricaoParcelasEmAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Descrição das\r\nParcelas em Atraso");
				#endregion

				#region [ Nº Parcela Maior Atraso ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNumParcelaMaiorAtraso);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Nº Parc Maior Atraso");
				#endregion

				#region [ Data Crédito Ok ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtCreditoOk);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Data Créd. Ok");
				#endregion

				#region [ Vendedor ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVendedor);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Vendedor");
				#endregion

				#region [ Parceiro ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlParceiro);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Parceiro");
				#endregion

				#region [ Email do Parceiro ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlParceiroEmail);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "Email");
				#endregion

				#region [ UF ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlUF);
				ExcelAutomation.SetProperty(oCell, "WrapText", true);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", FN_LISTAGEM);
				ExcelAutomation.SetProperty(oFont, "Size", FS_CABECALHO);
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				ExcelAutomation.SetProperty(oCell, "VerticalAlignment", ExcelAutomation.XlVAlign.xlVAlignBottom);
				ExcelAutomation.SetProperty(oCell, "Value", "UF");
				#endregion

				#endregion

				#region [ Obtém separador decimal usado pelo Excel ]
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

				#region [ Coluna: Parcelas em Atraso ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlParcelasEmAtraso) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlParcelasEmAtraso) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormat", "#00");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Coluna: Maior Atraso ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlMaiorAtraso) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlMaiorAtraso) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormat", "#" + strExcelThousandsSeparator + "#00");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Valor RA ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlValorRA) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlValorRA) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormat", "#" + strExcelThousandsSeparator + "##0" + strExcelDecimalSeparator + "00");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
				#endregion

				#region [ Coluna: Valor Total em Atraso ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlVlTotalEmAtraso) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlTotalEmAtraso) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormat", "#" + strExcelThousandsSeparator + "##0" + strExcelDecimalSeparator + "00");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignRight);
				#endregion

				#region [ Coluna: Descrição das Parcelas em Atraso ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDescricaoParcelasEmAtraso) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDescricaoParcelasEmAtraso) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignLeft);
				oFont = ExcelAutomation.GetProperty(oRange, "Font");
				ExcelAutomation.SetProperty(oFont, "Name", "Courier New");
				#endregion

				#region [ Coluna: Nº Parcela Maior Atraso ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlNumParcelaMaiorAtraso) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlNumParcelaMaiorAtraso) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormat", "##0º");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Data Crédito Ok ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDtCreditoOk) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDtCreditoOk) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "NumberFormat", "dd/mm/aaaa");
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#region [ Coluna: UF ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlUF) + (iNumLinha + 1).ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlUF) + MAX_LINHAS_EXCEL.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				ExcelAutomation.SetProperty(oRange, "HorizontalAlignment", ExcelAutomation.XlHAlign.xlHAlignCenter);
				#endregion

				#endregion

				#region [ Laço para listagem ]
				for (int i = 0; i < _listaClientesEmAtraso.contaClientes(); i++)
				{
					intQtdeRegistros++;
					iNumLinha++;
					if (intPrimeiraLinhaDados == 0) intPrimeiraLinhaDados = iNumLinha;
					intUltimaLinhaDados = iNumLinha;

					clienteEmAtraso = _listaClientesEmAtraso.getClienteEmAtraso(i);

					#region [ Transfere dados para o Excel (campos texto) ]

					#region [ Cliente ]
					vDados[iXlCliente - iOffSetArray] = Texto.iniciaisEmMaiusculas(clienteEmAtraso.nome_cliente) + " (" + Global.formataCnpjCpf(clienteEmAtraso.cnpj_cpf) + ")";
					#endregion

					#region [ Telefone ]
					strTelefone = Global.formataTelefone(clienteEmAtraso.ddd_res, clienteEmAtraso.tel_res);
					if (strTelefone.Length > 0) strTelefone = "R: " + strTelefone;
					strAux = Global.formataTelefone(clienteEmAtraso.ddd_com, clienteEmAtraso.tel_com, clienteEmAtraso.ramal_com);
					if (strAux.Length > 0) strAux = "C: " + strAux;
					if ((strAux.Length > 0) && (clienteEmAtraso.contato.Length > 0)) strAux += "\ncontato: " + Texto.iniciaisEmMaiusculas(clienteEmAtraso.contato);
					if ((strTelefone.Length > 0) && (strAux.Length > 0)) strTelefone += "\n";
					strTelefone += strAux;
					vDados[iXlTelefone - iOffSetArray] = strTelefone;
					#endregion

					#region [ Pedido ]
					vDados[iXlPedido - iOffSetArray] = montaDescricaoDadosPedidoComValorRA(clienteEmAtraso.listaDadosPedido);
					#endregion

					#region [ Descrição das Parcelas em Atraso ]
					vDados[iXlDescricaoParcelasEmAtraso - iOffSetArray] = montaDescricaoParcelas(clienteEmAtraso.listaParcelas);
					#endregion

					#region [ Vendedor ]
					vDados[iXlVendedor - iOffSetArray] = Texto.iniciaisEmMaiusculas(clienteEmAtraso.getVendedor());
					#endregion

					#region [ Parceiro ]
					vDados[iXlParceiro - iOffSetArray] = Texto.iniciaisEmMaiusculas(clienteEmAtraso.getIndicador());
					#endregion

					#region [ Email do Parceiro ]
					vDados[iXlParceiroEmail - iOffSetArray] = clienteEmAtraso.getIndicadorEmail();
					#endregion

					#region [ UF ]
					vDados[iXlUF - iOffSetArray] = clienteEmAtraso.uf;
					#endregion

					#region [ Transfere dados do vetor p/ o Excel ]
					strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
					oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
					ExcelAutomation.SetProperty(oRange, "Value2", vDados);
					#endregion

					#endregion

					#region [ Transfere dados para o Excel (campos datetime) ]

					#region [ Data Crédito Ok ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlDtCreditoOk);
					dtCreditoOk = clienteEmAtraso.getDataCreditoOk();
					if (dtCreditoOk != DateTime.MinValue) ExcelAutomation.SetProperty(oCell, "Value", dtCreditoOk.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador));
					#endregion

					#endregion

					#region [ Transfere dados para o Excel (campos numéricos) ]

					#region [ Parcelas em Atraso ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlParcelasEmAtraso);
					ExcelAutomation.SetProperty(oCell, "Value", clienteEmAtraso.qtde_parcelas_em_atraso);
					#endregion

					#region [ Maior Atraso ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlMaiorAtraso);
					ExcelAutomation.SetProperty(oCell, "Value", clienteEmAtraso.max_qtde_dias_em_atraso);
					#endregion

					#region [ Valor RA ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlValorRA);
					vl_total_RA = 0;
					for (int j = 0; j < clienteEmAtraso.listaDadosPedido.Count; j++)
					{
						vl_total_RA += clienteEmAtraso.listaDadosPedido[j].vl_RA;
					}
					if (blnExcelSuportaDecimalDataType)
					{
						ExcelAutomation.SetProperty(oCell, "Value", vl_total_RA);
					}
					else
					{
						ExcelAutomation.SetProperty(oCell, "Value", (double)vl_total_RA);
					}
					#endregion

					#region [ Valor Total em Atraso ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlTotalEmAtraso);
					if (blnExcelSuportaDecimalDataType)
					{
						ExcelAutomation.SetProperty(oCell, "Value", clienteEmAtraso.vl_total_em_atraso);
					}
					else
					{
						ExcelAutomation.SetProperty(oCell, "Value", (double)clienteEmAtraso.vl_total_em_atraso);
					}
					#endregion

					#region [ Nº da Parcela Maior Atraso ]
					oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlNumParcelaMaiorAtraso);
					ExcelAutomation.SetProperty(oCell, "Value", clienteEmAtraso.getNumeroParcelaMaiorAtraso());
					#endregion

					#endregion

					#region [ Borda inferior da linha ]
					oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
					oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeBottom);
					ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlDot);
					ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlHairline);
					ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
					#endregion
				}
				#endregion

				#region [ Linha com os totais ]

				#region [ Nova Linha ]
				iNumLinha++;
				#endregion

				#region [ Borda ]
				strAux = Global.excel_converte_numeracao_digito_para_letra(iXlDadosMinIndex) + iNumLinha.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlDadosMaxIndex) + iNumLinha.ToString();
				oRange = ExcelAutomation.GetProperty(oWS, "Range", strAux);
				oBorders = ExcelAutomation.GetProperty(oRange, "Borders");
				oBorder = ExcelAutomation.GetProperty(oBorders, "Item", ExcelAutomation.XlBordersIndex.xlEdgeTop);
				ExcelAutomation.SetProperty(oBorder, "LineStyle", ExcelAutomation.XlLineStyle.xlContinuous);
				ExcelAutomation.SetProperty(oBorder, "Weight", ExcelAutomation.XlBorderWeight.xlMedium);
				ExcelAutomation.SetProperty(oBorder, "ColorIndex", ExcelAutomation.XlColorIndex.xlColorIndexAutomatic);
				#endregion

				#region [ Total de registros ]
				strAux = "TOTAL: " + intQtdeRegistros.ToString() + " registro(s)";
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlCliente);
				ExcelAutomation.SetProperty(oCell, "WrapText", false);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				ExcelAutomation.SetProperty(oCell, "Value", strAux);
				#endregion

				#region [ Soma do valor total do RA ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlValorRA);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				strAux = "=SOMA(" + Global.excel_converte_numeracao_digito_para_letra(iXlValorRA) + intPrimeiraLinhaDados.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlValorRA) + intUltimaLinhaDados.ToString() + ")";
				ExcelAutomation.SetProperty(oCell, "Formula", strAux);
				#endregion

				#region [ Soma do valor total em atraso ]
				oCell = ExcelAutomation.GetProperty(oCells, "Item", iNumLinha, iXlVlTotalEmAtraso);
				oFont = ExcelAutomation.GetProperty(oCell, "Font");
				ExcelAutomation.SetProperty(oFont, "Bold", true);
				strAux = "=SOMA(" + Global.excel_converte_numeracao_digito_para_letra(iXlVlTotalEmAtraso) + intPrimeiraLinhaDados.ToString() + ":" + Global.excel_converte_numeracao_digito_para_letra(iXlVlTotalEmAtraso) + intUltimaLinhaDados.ToString() + ")";
				ExcelAutomation.SetProperty(oCell, "Formula", strAux);
				#endregion

				#endregion

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

		#endregion

		#region [ Eventos ]

		#region [ Form: FCobrancaAdministracao ]

		#region [ FCobrancaAdministracao_Load ]
		private void FCobrancaAdministracao_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			#endregion

			try
			{
				limpaCampos();

				#region [ Combo Situação ]
				cbSituacao.DataSource = Global.montaOpcaoCobrancaAdmSituacao(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbSituacao.DisplayMember = "descricao";
				cbSituacao.ValueMember = "codigo";
				cbSituacao.SelectedIndex = -1;
				#endregion

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

				#region [ Equipe de Vendas ]
				// Cria uma linha com a opção Todas
				DsDataSource.DtbEquipeVendasComboDataTable dtbEquipeVendas = new DsDataSource.DtbEquipeVendasComboDataTable();
				DsDataSource.DtbEquipeVendasComboRow rowEquipeVendas = dtbEquipeVendas.NewDtbEquipeVendasComboRow();
				rowEquipeVendas.id = 0;
				rowEquipeVendas.apelidoComDescricao = "Todas";
				dtbEquipeVendas.AddDtbEquipeVendasComboRow(rowEquipeVendas);
				// Obtém os dados do BD e faz um merge com a opção Todas
				dtbEquipeVendas.Merge(ComboDAO.criaDtbEquipeVendasCombo());
				cbEquipeVendas.DataSource = dtbEquipeVendas;
				cbEquipeVendas.ValueMember = "id";
				cbEquipeVendas.DisplayMember = "apelidoComDescricao";
				cbEquipeVendas.SelectedIndex = -1;
				#endregion

				#region [ Vendedor ]
				// Cria uma linha com a opção Todos
				DsDataSource.DtbVendedorComboDataTable dtbVendedor = new DsDataSource.DtbVendedorComboDataTable();
				DsDataSource.DtbVendedorComboRow rowVendedor = dtbVendedor.NewDtbVendedorComboRow();
				rowVendedor.usuario = "";
				rowVendedor.usuarioComNome = "Todos";
				dtbVendedor.AddDtbVendedorComboRow(rowVendedor);
				// Obtém os dados do BD e faz um merge com a opção Todos
				dtbVendedor.Merge(ComboDAO.criaDtbVendedor());
				cbVendedor.DataSource = dtbVendedor;
				cbVendedor.ValueMember = "usuario";
				cbVendedor.DisplayMember = "usuarioComNome";
				cbVendedor.SelectedIndex = -1;
				#endregion

				#region [ Indicador ]
				// Cria uma linha com a opção Todos
				DsDataSource.DtbIndicadorComboDataTable dtbIndicador = new DsDataSource.DtbIndicadorComboDataTable();
				DsDataSource.DtbIndicadorComboRow rowIndicador = dtbIndicador.NewDtbIndicadorComboRow();
				rowIndicador.apelido = "";
				rowIndicador.apelidoComRazaoSocialNome = "Todos";
				dtbIndicador.AddDtbIndicadorComboRow(rowIndicador);
				// Obtém os dados do BD e faz um merge com a opção Todos
				dtbIndicador.Merge(ComboDAO.criaDtbIndicador());
				cbIndicador.DataSource = dtbIndicador;
				cbIndicador.ValueMember = "apelido";
				cbIndicador.DisplayMember = "apelidoComRazaoSocialNome";
				cbIndicador.SelectedIndex = -1;
				#endregion

				#region [ Combo Garantia ]
				cbGarantia.DataSource = Global.montaOpcaoPedidoComGarantiaIndicador(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbGarantia.DisplayMember = "descricao";
				cbGarantia.ValueMember = "codigo";
				cbGarantia.SelectedIndex = -1;
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

		#region [ FCobrancaAdministracao_Shown ]
		private void FCobrancaAdministracao_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Prepara lista de auto complete do campo nome do cliente ]
					txtNomeCliente.AutoCompleteCustomSource.AddRange(FMain.fMain.listaNomeClienteAutoComplete.ToArray());
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

		#region [ FCobrancaAdministracao_FormClosing ]
		private void FCobrancaAdministracao_FormClosing(object sender, FormClosingEventArgs e)
		{
			try
			{
				// Campo nome do cliente exibe uma lista de sugestões
				if (ActiveControl == txtNomeCliente)
				{
					btnDummy.Focus();
					txtNomeCliente.Focus();
					Global.textBoxPosicionaCursorNoFinal(txtNomeCliente);
					e.Cancel = true;
				}
			}
			finally
			{
				#region [ Torna visível o form chamador? ]
				if (!e.Cancel)
				{
					if (_formChamador != null)
					{
						_formChamador.Location = this.Location;
						_formChamador.Visible = true;
						this.Visible = false;
					}
				}
				#endregion
			}
		}
		#endregion

		#region [ FCobrancaAdministracao_KeyDown ]
		private void FCobrancaAdministracao_KeyDown(object sender, KeyEventArgs e)
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

		#region [ cbSituacao ]

		#region [ cbSituacao_KeyDown ]
		private void cbSituacao_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtQtdeDiasAtrasoInicial);
		}
		#endregion

		#endregion

		#region [ txtQtdeDiasAtrasoInicial ]

		#region [ txtQtdeDiasAtrasoInicial_Enter ]
		private void txtQtdeDiasAtrasoInicial_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtQtdeDiasAtrasoInicial_Leave ]
		private void txtQtdeDiasAtrasoInicial_Leave(object sender, EventArgs e)
		{
			txtQtdeDiasAtrasoInicial.Text = txtQtdeDiasAtrasoInicial.Text.Trim();
		}
		#endregion

		#region [ txtQtdeDiasAtrasoInicial_KeyDown ]
		private void txtQtdeDiasAtrasoInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtQtdeDiasAtrasoFinal);
		}
		#endregion

		#region [ txtQtdeDiasAtrasoInicial_KeyPress ]
		private void txtQtdeDiasAtrasoInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtQtdeDiasAtrasoFinal ]

		#region [ txtQtdeDiasAtrasoFinal_Enter ]
		private void txtQtdeDiasAtrasoFinal_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtQtdeDiasAtrasoFinal_Leave ]
		private void txtQtdeDiasAtrasoFinal_Leave(object sender, EventArgs e)
		{
			txtQtdeDiasAtrasoFinal.Text = txtQtdeDiasAtrasoFinal.Text.Trim();
		}
		#endregion

		#region [ txtQtdeDiasAtrasoFinal_KeyDown ]
		private void txtQtdeDiasAtrasoFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, cbContaCorrente);
		}
		#endregion

		#region [ txtQtdeDiasAtrasoFinal_KeyPress ]
		private void txtQtdeDiasAtrasoFinal_KeyPress(object sender, KeyPressEventArgs e)
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
			Global.trataComboBoxKeyDown(sender, e, cbEquipeVendas);
		}
		#endregion

		#endregion

		#region [ cbEquipeVendas ]

		#region [ cbEquipeVendas_KeyDown ]
		private void cbEquipeVendas_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbVendedor);
		}
		#endregion

		#endregion

		#region [ cbVendedor ]

		#region [ cbVendedor_KeyDown ]
		private void cbVendedor_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbIndicador);
		}
		#endregion

		#endregion

		#region [ cbIndicador ]

		#region [ cbIndicador_KeyDown ]
		private void cbIndicador_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbGarantia);
		}
		#endregion

		#endregion

		#region [ cbGarantia ]

		#region [ cbGarantia_KeyDown ]
		private void cbGarantia_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtNomeCliente);
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
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtCnpjCpf_Leave ]
		private void txtCnpjCpf_Leave(object sender, EventArgs e)
		{
			txtCnpjCpf.Text = Global.formataCnpjCpf(txtCnpjCpf.Text.Trim());
			if (txtCnpjCpf.Text.Trim().Length > 0)
			{
				if (!Global.isCnpjCpfOk(txtCnpjCpf.Text.Trim()))
				{
					avisoErro("CNPJ/CPF inválido!");
					txtCnpjCpf.Focus();
					return;
				}
			}
		}
		#endregion

		#region [ txtCnpjCpf_KeyDown ]
		private void txtCnpjCpf_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, gridDados);
		}
		#endregion

		#region [ txtCnpjCpf_KeyPress ]
		private void txtCnpjCpf_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCnpjCpf(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ gridDados ]

		#region [ gridDados_SortCompare ]
		private void gridDados_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			if (e.Column.Name.Equals(GRID_COL_VALOR_TOTAL_EM_ATRASO))
			{
				e.SortResult = System.Decimal.Compare(Global.converteNumeroDecimal(e.CellValue1.ToString()), Global.converteNumeroDecimal(e.CellValue2.ToString()));
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_COL_QTDE_PARCELAS_EM_ATRASO))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_COL_MAX_DIAS_EM_ATRASO))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_COL_NUM_PARCELA_MAIOR_ATRASO))
			{
				e.SortResult = (int)Global.converteInteiro(Global.digitos(e.CellValue1.ToString())) - (int)Global.converteInteiro(Global.digitos(e.CellValue2.ToString()));
				e.Handled = true;
			}
		}
		#endregion

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

		#endregion

		#region [ Botões / Menu ]

		#region [ Marcar todas as linhas ]

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

		#region [ Desmarcar todas as linhas ]

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

		#region [ Pesquisar ]

		#region [ btnPesquisar_Click ]
		private void btnPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#endregion

		#region [ Planilha Excel ]

		#region [ btnExcel_Click ]
		private void btnExcel_Click(object sender, EventArgs e)
		{
			geraPlanilhaExcel();
		}
		#endregion

		#endregion

		#region [ btnPrinterDialog ]
		private void btnPrinterDialog_Click(object sender, EventArgs e)
		{
			printerDialog();
		}
		#endregion

		#region [ btnPrintPreview ]
		private void btnPrintPreview_Click(object sender, EventArgs e)
		{
			printPreview();
		}
		#endregion

		#region [ btnImprimir ]
		private void btnImprimir_Click(object sender, EventArgs e)
		{
			imprimeConsulta();
		}
		#endregion

		#endregion

		#endregion

		#region [ Impressão ]

		#region [ prnDocConsulta_BeginPrint ]
		private void prnDocConsulta_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
		{
			if (_listaClientesEmAtraso == null)
			{
				e.Cancel = true;
				return;
			}

			if (_listaClientesEmAtraso.contaClientes() == 0)
			{
				e.Cancel = true;
				avisoErro("Não há dados!!");
				return;
			}

			prnDocConsulta.DefaultPageSettings.Landscape = true;

			impressao = new Impressao(prnDocConsulta.DefaultPageSettings.Landscape);

			#region [ Prepara elementos de impressão ]
			fonteTitulo = new Font(NOME_FONTE_DEFAULT, 12f, FontStyle.Bold);
			fonteListagem = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteListagemNegrito = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteDataEmissao = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteFiltros = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Italic | FontStyle.Bold);
			fonteNumPagina = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			brushPadrao = new SolidBrush(Color.Black);
			penTracoTitulo = new Pen(brushPadrao, .5f);
			penTracoPontilhado = Impressao.criaPenTracoPontilhado();
			#endregion

			_intImpressaoIdxLinha = 0;
			_intImpressaoNumPagina = 0;
			_strImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
			_intQtdeTotalRegistros = 0;
			_vlTotalRegistros = 0m;
		}
		#endregion

		#region [ prnDocConsulta_PrintPage ]
		private void prnDocConsulta_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			const float hPrevista = 40f;
			float cx;
			float cy;
			float hMax;
			float hNome;
			float hDescricaoParcelasEmAtraso;
			float hTelefone;
			float hPedido;
			decimal vlTotalEmAtrasoAux;
			RectangleF r;
			String strAux;
			String strTexto;
			String strNome;
			String strDescricaoParcelasEmAtraso;
			String strTelefone;
			String strPedido;
			int intLinhasImpressasNestaPagina = 0;
			CobrancaAdminRegistroClienteEmAtraso clienteEmAtraso;
			#endregion

			#region [ Consistência ]
			if ((_listaClientesEmAtraso == null) || (_listaClientesEmAtraso.contaClientes() == 0))
			{
				e.Cancel = true;
				return;
			}
			#endregion

			#region [ Contador de página ]
			_intImpressaoNumPagina++;
			#endregion

			e.Graphics.PageUnit = GraphicsUnit.Millimeter;
			if (_intImpressaoNumPagina == 1)
			{
				#region [ Medidas do papel ]
				prnDocConsulta.DocumentName = "Carteira em Atraso";
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

				#region [ Layout das colunas da listagem ]
				ESPACAMENTO_COLUNAS = 2f;
				wxTelefone = 35f;
				wxPedido = 55f;
				wxQtdeParcelasEmAtraso = 14f;
				wxMaiorAtraso = 12f;
				wxValorTotalEmAtraso = 18f;
				wxDescricaoParcelasEmAtraso = 38f;
				wxNome = larguraUtil
						 - ESPACAMENTO_COLUNAS - wxTelefone
						 - ESPACAMENTO_COLUNAS - wxPedido
						 - ESPACAMENTO_COLUNAS - wxQtdeParcelasEmAtraso
						 - ESPACAMENTO_COLUNAS - wxMaiorAtraso
						 - ESPACAMENTO_COLUNAS - wxValorTotalEmAtraso
						 - ESPACAMENTO_COLUNAS - wxDescricaoParcelasEmAtraso;
				ixNome = cxInicio;
				ixTelefone = ixNome + wxNome + ESPACAMENTO_COLUNAS;
				ixPedido = ixTelefone + wxTelefone + ESPACAMENTO_COLUNAS;
				ixQtdeParcelasEmAtraso = ixPedido + wxPedido + 2 * ESPACAMENTO_COLUNAS;  // Dá um espaçamento maior p/ melhorar visual
				ixMaiorAtraso = ixQtdeParcelasEmAtraso + wxQtdeParcelasEmAtraso + ESPACAMENTO_COLUNAS;
				ixValorTotalEmAtraso = ixMaiorAtraso + wxMaiorAtraso + ESPACAMENTO_COLUNAS;
				ixDescricaoParcelasEmAtraso = ixValorTotalEmAtraso + wxValorTotalEmAtraso + ESPACAMENTO_COLUNAS;
				#endregion
			}

			cy = cyInicio;

			#region [ Título ]
			strTexto = "CARTEIRA EM ATRASO";
			fonteAtual = fonteTitulo;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Somente 1ª parcela em atraso? ]
			strTexto = "Somente 1ª parcela em atraso: " + _filtroSomentePrimeiraParcelaEmAtraso;
			fonteAtual = fonteFiltros;
			cy -= fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Data da emissão ]
			strTexto = "Emissão: " + _strImpressaoDataEmissao;
			fonteAtual = fonteListagemNegrito;
			cy -= fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion
			
			#region [ Informações no cabeçalho ]

			cy += .5f;

			#region [ Filtros ]

			#region [ Configura fonte ]
			fonteAtual = fonteFiltros;
			#endregion

			#region [ Situação ]
			strTexto = "Situação: ";
			strTexto += (_filtroSituacaoAdmCobranca.Length > 0) ? _filtroSituacaoAdmCobranca : "Todas";
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Atrasado entre ]
			strTexto = "Atrasado entre: ";
			strTexto += (_filtroQtdeDiasAtrasoInicial.Length > 0) ? _filtroQtdeDiasAtrasoInicial : "N.I.";
			strTexto += " e ";
			strTexto += (_filtroQtdeDiasAtrasoFinal.Length > 0) ? _filtroQtdeDiasAtrasoFinal : "N.I.";
			strTexto += " dias";
			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Conta Corrente ]
			strTexto = "Conta Corrente: ";
			strTexto += (_filtroContaCorrente.Length > 0) ? _filtroContaCorrente : "Todas";
			cx = cxInicio + larguraUtil * .66f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Plano Contas Empresa ]
			strTexto = "Empresa: ";
			strTexto += (_filtroPlanoContasEmpresa.Length > 0) ? _filtroPlanoContasEmpresa : "Todas";
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Plano Contas Grupo ]
			strTexto = "Grupo: ";
			strTexto += (_filtroPlanoContasGrupo.Length > 0) ? _filtroPlanoContasGrupo : "Todos";
			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Plano Contas Conta ]
			strTexto = "Conta: ";
			strTexto += (_filtroPlanoContasConta.Length > 0) ? _filtroPlanoContasConta : "Todos";
			cx = cxInicio + larguraUtil * .66f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion
			
			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Equipe de Vendas ]
			strTexto = "Equipe de Vendas: ";
			strTexto += (_filtroEquipeVendas.Length > 0) ? _filtroEquipeVendas : "Todas";
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Vendedor ]
			strTexto = "Vendedor: ";
			strTexto += (_filtroVendedor.Length > 0) ? _filtroVendedor : "Todos";
			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Indicador ]
			strTexto = "Indicador: ";
			strTexto += (_filtroIndicador.Length > 0) ? _filtroIndicador : "Todos";
			cx = cxInicio + larguraUtil * .66f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			#region [ Garantia ]
			strTexto = "Garantia: ";
			strTexto += (_filtroGarantia.Length > 0) ? _filtroGarantia : "N.I.";
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nome Cliente ]
			strTexto = "Nome Cliente: ";
			strTexto += (_filtroNomeCliente.Length > 0) ? _filtroNomeCliente : "N.I.";
			cx = cxInicio + larguraUtil * .33f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ CNPJ/CPF ]
			strTexto = "CNPJ/CPF: ";
			strTexto += (_filtroCnpjCpf.Length > 0) ? Global.formataCnpjCpf(_filtroCnpjCpf) : "N.I.";
			cx = cxInicio + larguraUtil * .66f;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Nova linha ]
			cy += fonteAtual.GetHeight(e.Graphics);
			cx = cxInicio;
			#endregion

			cy += .5f;
			#endregion

			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Títulos da listagem ]
			cy += .5f;
			fonteAtual = fonteListagemNegrito;
			strTexto = "CLIENTE";
			cx = ixNome;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			strTexto = "TELEFONE";
			cx = ixTelefone;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			strTexto = "PEDIDO";
			cx = ixPedido;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			strTexto = "PARC EM";
			cx = ixQtdeParcelasEmAtraso;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			strTexto = "ATRASO";
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			strTexto = "MAIOR";
			cx = ixMaiorAtraso;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			strTexto = "ATRASO";
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			strTexto = "VALOR TOTAL";
			cx = ixValorTotalEmAtraso + wxValorTotalEmAtraso - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			strTexto = "EM ATRASO";
			cx = ixValorTotalEmAtraso + wxValorTotalEmAtraso - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			strTexto = "DESCRIÇÃO DAS PARCELAS";
			cx = ixDescricaoParcelasEmAtraso;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			strTexto = "EM ATRASO";
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy + alturaLinhaListagemNegrito);

			cy += 2 * alturaLinhaListagemNegrito;
			cy += .5f;
			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Laço para listagem ]
			fonteAtual = fonteListagem;
			while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
				   (_intImpressaoIdxLinha < _listaClientesEmAtraso.contaClientes()))
			{
				hMax = alturaLinhaListagem;

				clienteEmAtraso = _listaClientesEmAtraso.getClienteEmAtraso(_intImpressaoIdxLinha);

				#region [ Há espaço suficiente p/ os campos maiores? ]
				strNome = clienteEmAtraso.nome_cliente.ToUpper() + " (" + Global.formataCnpjCpf(clienteEmAtraso.cnpj_cpf) + ")";
				hNome = e.Graphics.MeasureString(strNome, fonteAtual, (int)wxNome).Height;

				strDescricaoParcelasEmAtraso = montaDescricaoParcelas(clienteEmAtraso.listaParcelas);
				hDescricaoParcelasEmAtraso = e.Graphics.MeasureString(strDescricaoParcelasEmAtraso, fonteAtual, (int)wxDescricaoParcelasEmAtraso).Height;

				strTelefone = Global.formataTelefone(clienteEmAtraso.ddd_res, clienteEmAtraso.tel_res);
				if (strTelefone.Length > 0) strTelefone = "R: " + strTelefone;
				strAux = Global.formataTelefone(clienteEmAtraso.ddd_com, clienteEmAtraso.tel_com, clienteEmAtraso.ramal_com);
				if (strAux.Length > 0) strAux = "C: " + strAux;
				if ((strAux.Length > 0) && (clienteEmAtraso.contato.Length > 0)) strAux += "\ncontato: " + clienteEmAtraso.contato;
				if ((strTelefone.Length > 0) && (strAux.Length > 0)) strTelefone += "\n";
				strTelefone += strAux;
				hTelefone = e.Graphics.MeasureString(strTelefone, fonteAtual, (int)wxTelefone).Height;

				strPedido = montaDescricaoDadosPedido(clienteEmAtraso.listaDadosPedido);
				hPedido = e.Graphics.MeasureString(strPedido, fonteAtual, (int)wxPedido).Height;

				hMax = Math.Max(hMax, hNome);
				hMax = Math.Max(hMax, hDescricaoParcelasEmAtraso);
				hMax = Math.Max(hMax, hTelefone);
				hMax = Math.Max(hMax, hPedido);

				if ((cy + hMax) > (cyRodapeNumPagina - 5)) break;
				#endregion

				#region [ Nome ]
				r = new RectangleF(ixNome, cy, wxNome, hPrevista);
				e.Graphics.DrawString(strNome, fonteAtual, brushPadrao, r);
				#endregion

				#region [ Telefone ]
				r = new RectangleF(ixTelefone, cy, wxTelefone, hPrevista);
				e.Graphics.DrawString(strTelefone, fonteAtual, brushPadrao, r);
				#endregion

				#region [ Pedido ]
				r = new RectangleF(ixPedido, cy, wxPedido, hPrevista);
				e.Graphics.DrawString(strPedido, fonteAtual, brushPadrao, r);
				#endregion

				#region [ Parcelas em Atraso ]
				strTexto = clienteEmAtraso.qtde_parcelas_em_atraso.ToString().PadLeft(2, '0');
				cx = ixQtdeParcelasEmAtraso;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Maior Atraso ]
				strTexto = Global.formataInteiro(clienteEmAtraso.max_qtde_dias_em_atraso).PadLeft(2, '0').PadLeft(5, ' ');
				cx = ixMaiorAtraso;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Valor Total em Atraso ]
				strTexto = Global.formataMoeda(clienteEmAtraso.vl_total_em_atraso);
				vlTotalEmAtrasoAux = Global.converteNumeroDecimal(strTexto);
				cx = ixValorTotalEmAtraso + wxValorTotalEmAtraso - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion


				#region [ Descrição das Parcelas em Atraso ]
				cx = ixDescricaoParcelasEmAtraso;
				r = new RectangleF(ixDescricaoParcelasEmAtraso, cy, wxDescricaoParcelasEmAtraso, hPrevista);
				strTexto = montaDescricaoParcelas(clienteEmAtraso.listaParcelas);
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				#endregion

				cy += hMax;

				_intQtdeTotalRegistros++;
				_vlTotalRegistros += vlTotalEmAtrasoAux;

				intLinhasImpressasNestaPagina++;
				_intImpressaoIdxLinha++;

				#region [ Na última linha não imprime o tracejado ]
				if (_intImpressaoIdxLinha < _listaClientesEmAtraso.contaClientes())
				{
					#region [ Traço pontilhado ]
					cy += .5f;
					e.Graphics.DrawLine(penTracoPontilhado, cxInicio, cy, cxFim, cy);
					cy += .5f;
					#endregion
				}
				#endregion
			}
			#endregion

			#region [ Tem mais páginas para imprimir? ]
			if (_intImpressaoIdxLinha < _listaClientesEmAtraso.contaClientes())
			{
				e.HasMorePages = true;
			}
			else
			{
				e.HasMorePages = false;

				#region [ Há espaço suficiente? ]
				if ((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 10))
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
					cx = cxInicio;
					strTexto = "TOTAL";
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataInteiro(_intQtdeTotalRegistros) + " registro(s)";
					cx = ixTelefone;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(_vlTotalRegistros);
					cx = ixValorTotalEmAtraso + wxValorTotalEmAtraso - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
					#endregion
				}
				else e.HasMorePages = true;
				#endregion
			}
			#endregion

			#region [ Imprime nº página ]
			strTexto = "Página: " + _intImpressaoNumPagina.ToString().PadLeft(2, ' ');
			fonteAtual = fonteNumPagina;
			cy = cyRodapeNumPagina;
			cx = cxInicio + larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion
		}
		#endregion

		#endregion
	}

	#region [ Classe CobrancaAdminListaClienteEmAtraso ]
	class CobrancaAdminListaClienteEmAtraso
	{
		#region [ Atributos ]
		private StringBuilder _relacaoIdClientesNaLista;
		private List<CobrancaAdminRegistroClienteEmAtraso> _listaClienteEmAtraso;
		#endregion

		#region [ Construtor ]
		public CobrancaAdminListaClienteEmAtraso()
		{
			_relacaoIdClientesNaLista = new StringBuilder();
			_relacaoIdClientesNaLista.Append('|');
			_listaClienteEmAtraso = new List<CobrancaAdminRegistroClienteEmAtraso>();
		}
		#endregion

		#region [ isIdClienteExiste ]
		public bool isIdClienteExiste(String idCliente)
		{
			if (idCliente == null) return false;
			if (idCliente.Trim().Length == 0) return false;
			if (_relacaoIdClientesNaLista.ToString().IndexOf('|' + idCliente + '|') != -1) return true;
			return false;
		}
		#endregion

		#region [ adicionaClienteEmAtraso ]
		public CobrancaAdminRegistroClienteEmAtraso adicionaClienteEmAtraso(
										String id_cliente, 
										String cnpj_cpf, 
										String nome_cliente, 
										int max_qtde_dias_em_atraso,
										String ddd_res,
										String tel_res,
										String ddd_com,
										String tel_com,
										String ramal_com,
										String contato,
										String uf)
		{
			#region [ Declarações ]
			CobrancaAdminRegistroClienteEmAtraso cliente;
			#endregion

			#region [ Cliente já existe na lista? ]
			if (isIdClienteExiste(id_cliente))
			{
				throw new Exception("Cliente " + nome_cliente + " (id=" + id_cliente + ", CNPJ/CPF=" + cnpj_cpf + ") já consta na lista!!");
			}
			#endregion

			cliente = new CobrancaAdminRegistroClienteEmAtraso(
										id_cliente,
										cnpj_cpf,
										nome_cliente,
										max_qtde_dias_em_atraso,
										ddd_res,
										tel_res,
										ddd_com,
										tel_com,
										ramal_com,
										contato,
										uf);
			_listaClienteEmAtraso.Add(cliente);
			_relacaoIdClientesNaLista.Append(id_cliente + '|');
			return cliente;
		}
		#endregion

		#region [ getClienteEmAtraso ]
		public CobrancaAdminRegistroClienteEmAtraso getClienteEmAtraso(String id_cliente)
		{
			for (int i = (_listaClienteEmAtraso.Count - 1); i >= 0; i--)
			{
				if (_listaClienteEmAtraso[i].id_cliente.Equals(id_cliente))
				{
					return _listaClienteEmAtraso[i];
				}
			}
			return null;
		}

		public CobrancaAdminRegistroClienteEmAtraso getClienteEmAtraso(int indice)
		{
			if ((indice < 0) || (indice >= _listaClienteEmAtraso.Count)) return null;
			return _listaClienteEmAtraso[indice];
		}
		#endregion

		#region [ contaClientes ]
		public int contaClientes()
		{
			return _listaClienteEmAtraso.Count;
		}
		#endregion

		#region [ Clear ]
		public void Clear()
		{
			_listaClienteEmAtraso.Clear();
			_relacaoIdClientesNaLista.Remove(0, _relacaoIdClientesNaLista.Length);
			_relacaoIdClientesNaLista.Append('|');
		}
		#endregion
	}
	#endregion

	#region [ Classe CobrancaAdminRegistroClienteEmAtraso ]
	class CobrancaAdminRegistroClienteEmAtraso
	{
		#region[ Atributos ]
		private StringBuilder _relacaoPedidoNaLista;

		private List<CobrancaAdminParcelaEmAtraso> _listaParcelas;
		public List<CobrancaAdminParcelaEmAtraso> listaParcelas
		{
			get { return _listaParcelas; }
		}

		private List<CobrancaAdminDadosPedidoParcelaEmAtraso> _listaDadosPedido;
		public List<CobrancaAdminDadosPedidoParcelaEmAtraso> listaDadosPedido
		{
			get { return _listaDadosPedido; }
		}

		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
		}

		private String _cnpj_cpf;
		public String cnpj_cpf
		{
			get { return _cnpj_cpf; }
		}

		private String _nome_cliente;
		public String nome_cliente
		{
			get { return _nome_cliente; }
		}

		private decimal _vl_total_em_atraso;
		public decimal vl_total_em_atraso
		{
			get { return _vl_total_em_atraso; }
		}

		private int _max_qtde_dias_em_atraso;
		public int max_qtde_dias_em_atraso
		{
			get { return _max_qtde_dias_em_atraso; }
		}

		private int _qtde_parcelas_em_atraso;
		public int qtde_parcelas_em_atraso
		{
			get { return _qtde_parcelas_em_atraso; }
		}

		private String _ddd_res;
		public String ddd_res
		{
			get { return _ddd_res; }
		}

		private String _tel_res;
		public String tel_res
		{
			get { return _tel_res; }
		}

		private String _ddd_com;
		public String ddd_com
		{
			get { return _ddd_com; }
		}

		private String _tel_com;
		public String tel_com
		{
			get { return _tel_com; }
		}

		private String _ramal_com;
		public String ramal_com
		{
			get { return _ramal_com; }
		}

		private String _contato;
		public String contato
		{
			get { return _contato; }
		}

		private String _uf;
		public String uf
		{
			get { return _uf; }
		}
		#endregion

		#region [ Construtor ]
		public CobrancaAdminRegistroClienteEmAtraso(
									String id_cliente,
									String cnpj_cpf,
									String nome_cliente,
									int max_qtde_dias_em_atraso,
									String ddd_res,
									String tel_res,
									String ddd_com,
									String tel_com,
									String ramal_com,
									String contato,
									String uf)
		{
			_relacaoPedidoNaLista = new StringBuilder();
			_relacaoPedidoNaLista.Append('|');
			_id_cliente = id_cliente;
			_cnpj_cpf = cnpj_cpf;
			_nome_cliente = nome_cliente;
			_max_qtde_dias_em_atraso = max_qtde_dias_em_atraso;
			_ddd_res = ddd_res;
			_tel_res = tel_res;
			_ddd_com = ddd_com;
			_tel_com = tel_com;
			_ramal_com = ramal_com;
			_contato = contato;
			_uf = uf;
			_listaParcelas = new List<CobrancaAdminParcelaEmAtraso>();
			_listaDadosPedido = new List<CobrancaAdminDadosPedidoParcelaEmAtraso>();
		}
		#endregion

		#region [ adicionaParcelaEmAtraso ]
		public void adicionaParcelaEmAtraso(CobrancaAdminParcelaEmAtraso parcela)
		{
			#region [ Adiciona os itens nas listas ]
			_listaParcelas.Add(parcela);
			#endregion

			#region [ Processa dados ]
			_qtde_parcelas_em_atraso++;
			_vl_total_em_atraso += parcela.valor;
			#endregion
		}
		#endregion

		#region [ isPedidoExiste ]
		public bool isPedidoExiste(String pedido)
		{
			if (pedido == null) return false;
			if (pedido.Trim().Length == 0) return false;
			if (_relacaoPedidoNaLista.ToString().IndexOf('|' + pedido + '|') != -1) return true;
			return false;
		}
		#endregion

		#region [ adicionaPedido ]
		public void adicionaPedido(CobrancaAdminDadosPedidoParcelaEmAtraso dadosPedido)
		{
			_listaDadosPedido.Add(dadosPedido);
			_relacaoPedidoNaLista.Append(dadosPedido.pedido + '|');
		}
		#endregion

		#region [ getIndicador ]
		public String getIndicador()
		{
			#region [ Declarações ]
			String strAux;
			String strResposta = "";
			String strRelacaoIndicadores = "";
			#endregion

			for (int i = 0; i < _listaDadosPedido.Count; i++)
			{
				if (_listaDadosPedido[i].indicador.Length > 0)
				{
					strAux = "|" + _listaDadosPedido[i].indicador + "|";
					if (strRelacaoIndicadores.IndexOf(strAux) == -1)
					{
						strRelacaoIndicadores += strAux;
						if (strResposta.Length > 0) strResposta += ", ";
						strResposta += _listaDadosPedido[i].indicador;
					}
				}
			}
			return strResposta;
		}
		#endregion

		#region [ getIndicadorEmail ]
		public String getIndicadorEmail()
		{
			#region [ Declarações ]
			String strAux;
			String strResposta = "";
			String strRelacaoEmails = "";
			#endregion

			for (int i = 0; i < _listaDadosPedido.Count; i++)
			{
				if (_listaDadosPedido[i].indicador_email.Length > 0)
				{
					strAux = "|" + _listaDadosPedido[i].indicador_email + "|";
					if (strRelacaoEmails.IndexOf(strAux) == -1)
					{
						strRelacaoEmails += strAux;
						if (strResposta.Length > 0) strResposta += "; ";
						strResposta += _listaDadosPedido[i].indicador_email;
					}
				}
			}
			strResposta = strResposta.ToLower();
			strResposta = strResposta.Replace(',', ';');
			strResposta = strResposta.Replace('/', ';');
			strResposta = strResposta.Replace(" ", "");
			strResposta = strResposta.Replace(";", "; ");
			return strResposta;
		}
		#endregion

		#region [ getVendedor ]
		public String getVendedor()
		{
			#region [ Declarações ]
			String strAux;
			String strResposta = "";
			String strRelacaoVendedores = "";
			#endregion

			for (int i = 0; i < _listaDadosPedido.Count; i++)
			{
				if (_listaDadosPedido[i].vendedor.Length > 0)
				{
					strAux = "|" + _listaDadosPedido[i].vendedor + "|";
					if (strRelacaoVendedores.IndexOf(strAux) == -1)
					{
						strRelacaoVendedores += strAux;
						if (strResposta.Length > 0) strResposta += ", ";
						strResposta += _listaDadosPedido[i].vendedor;
					}
				}
			}
			return strResposta;
		}
		#endregion

		#region [ getNumeroParcelaMaiorAtraso ]
		public int getNumeroParcelaMaiorAtraso()
		{
			#region [ Declarações ]
			int numeroParcelaMaiorAtraso = 0;
			int qtdeDiasMaiorAtraso = 0;
			#endregion

			for (int i = 0; i < _listaParcelas.Count; i++)
			{
				if (_listaParcelas[i].ctrl_pagto_modulo != Global.Cte.FIN.CtrlPagtoModulo.BOLETO) continue;
				if (_listaParcelas[i].qtde_dias_em_atraso > qtdeDiasMaiorAtraso)
				{
					qtdeDiasMaiorAtraso = _listaParcelas[i].qtde_dias_em_atraso;
					numeroParcelaMaiorAtraso = _listaParcelas[i].tFBI_num_parcela;
				}
			}
			return numeroParcelaMaiorAtraso;
		}
		#endregion

		#region [ getDataCreditoOk ]
		public DateTime getDataCreditoOk()
		{
			#region [ Declarações ]
			DateTime dtCreditoOk = DateTime.MinValue;
			#endregion

			for (int i = 0; i < _listaDadosPedido.Count; i++)
			{
				// Localiza e retorna a data de crédito ok mais antiga
				if (_listaDadosPedido[i].analise_credito == Global.Cte.FIN.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
				{
					if (_listaDadosPedido[i].analise_credito_data != DateTime.MinValue)
					{
						if ((dtCreditoOk == DateTime.MinValue) || (_listaDadosPedido[i].analise_credito_data < dtCreditoOk)) dtCreditoOk = _listaDadosPedido[i].analise_credito_data;
					}
				}
			}

			return dtCreditoOk;
		}
		#endregion
	}
	#endregion

	#region [ CobrancaAdminParcelaEmAtraso ]
	class CobrancaAdminParcelaEmAtraso
	{
		#region [ Atributos ]
		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
		}

		private int _id;
		public int id
		{
			get { return _id; }
		}

		private byte _id_conta_corrente;
		public byte id_conta_corrente
		{
			get { return _id_conta_corrente; }
		}

		private byte _id_plano_contas_empresa;
		public byte id_plano_contas_empresa
		{
			get { return _id_plano_contas_empresa; }
		}

		private int _id_plano_contas_grupo;
		public int id_plano_contas_grupo
		{
			get { return _id_plano_contas_grupo; }
		}

		private int _id_plano_contas_conta;
		public int id_plano_contas_conta
		{
			get { return _id_plano_contas_conta; }
		}

		private DateTime _dt_competencia;
		public DateTime dt_competencia
		{
			get { return _dt_competencia; }
		}

		private decimal _valor;
		public decimal valor
		{
			get { return _valor; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
		}

		private int _ctrl_pagto_id_parcela;
		public int ctrl_pagto_id_parcela
		{
			get { return _ctrl_pagto_id_parcela; }
		}

		private byte _ctrl_pagto_modulo;
		public byte ctrl_pagto_modulo
		{
			get { return _ctrl_pagto_modulo; }
		}

		private int _qtde_dias_em_atraso;
		public int qtde_dias_em_atraso
		{
			get { return _qtde_dias_em_atraso; }
		}

		private Int16 _tFBI_status;
		public Int16 tFBI_status
		{
			get { return _tFBI_status; }
		}

		private int _tFBI_num_parcela;
		public int tFBI_num_parcela
		{
			get { return _tFBI_num_parcela; }
		}
		#endregion

		#region [ Construtor ]
		public CobrancaAdminParcelaEmAtraso(String id_cliente,
											int id,
											byte id_conta_corrente,
											byte id_plano_contas_empresa,
											int id_plano_contas_grupo,
											int id_plano_contas_conta,
											DateTime dt_competencia,
											decimal valor,
											String descricao,
											int ctrl_pagto_id_parcela,
											byte ctrl_pagto_modulo,
											int qtde_dias_em_atraso,
											Int16 tFBI_status,
											int tFBI_num_parcela
											)
		{
			_id_cliente = id_cliente;
			_id = id;
			_id_conta_corrente = id_conta_corrente;
			_id_plano_contas_empresa = id_plano_contas_empresa;
			_id_plano_contas_grupo = id_plano_contas_grupo;
			_id_plano_contas_conta = id_plano_contas_conta;
			_dt_competencia = dt_competencia;
			_valor = valor;
			_descricao = descricao;
			_ctrl_pagto_id_parcela = ctrl_pagto_id_parcela;
			_ctrl_pagto_modulo = ctrl_pagto_modulo;
			_qtde_dias_em_atraso = qtde_dias_em_atraso;
			_tFBI_status = tFBI_status;
			_tFBI_num_parcela = tFBI_num_parcela;
		}
		#endregion
	}
	#endregion

	#region [ CobrancaAdminDadosPedidoParcelaEmAtraso ]
	class CobrancaAdminDadosPedidoParcelaEmAtraso
	{
		#region [ Atributos ]
		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
		}

		private int _id;
		public int id
		{
			get { return _id; }
		}

		private int _id_boleto_item;
		public int id_boleto_item
		{
			get { return _id_boleto_item; }
		}

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
		}

		private String _vendedor;
		public String vendedor
		{
			get { return _vendedor; }
		}

		private int _id_equipe_vendas;
		public int id_equipe_vendas
		{
			get { return _id_equipe_vendas; }
		}

		private String _equipe_vendas;
		public String equipe_vendas
		{
			get { return _equipe_vendas; }
		}

		private String _indicador;
		public String indicador
		{
			get { return _indicador; }
		}

		private String _indicador_email;
		public String indicador_email
		{
			get { return _indicador_email; }
		}

		private byte _garantiaIndicadorStatus;
		public byte garantiaIndicadorStatus
		{
			get { return _garantiaIndicadorStatus; }
		}

		private Decimal _vl_RA;
		public Decimal vl_RA
		{
			get { return _vl_RA; }
		}

		private int _analise_credito;
		public int analise_credito
		{
			get { return _analise_credito; }
		}

		private DateTime _analise_credito_data;
		public DateTime analise_credito_data
		{
			get { return _analise_credito_data; }
		}
		#endregion

		#region [ Construtor ]
		public CobrancaAdminDadosPedidoParcelaEmAtraso(String id_cliente,
														int id,
														int id_boleto_item,
														String pedido,
														String vendedor,
														int id_equipe_vendas,
														String equipe_vendas,
														String indicador,
														String indicador_email,
														byte garantiaIndicadorStatus,
														Decimal vl_RA,
														int analise_credito,
														DateTime analise_credito_data)
		{
			_id_cliente = id_cliente;
			_id = id;
			_id_boleto_item = id_boleto_item;
			_pedido = pedido;
			_vendedor = vendedor;
			_id_equipe_vendas = id_equipe_vendas;
			_equipe_vendas = equipe_vendas;
			_indicador = indicador;
			_indicador_email = indicador_email;
			_garantiaIndicadorStatus = garantiaIndicadorStatus;
			_vl_RA = vl_RA;
			_analise_credito = analise_credito;
			_analise_credito_data = analise_credito_data;
		}
		#endregion
	}
	#endregion
}
