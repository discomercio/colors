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
using System.Net.Mail;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using mshtml;
#endregion

namespace Financeiro
{
	public partial class FBoletoConsulta : Financeiro.FModelo
	{
		#region [ interface IHTMLElementRender ]
		// Replacement for mshtml imported interface, Tlbimp.exe generates wrong signatures
		[ComImport, InterfaceType((short)1), Guid("3050F669-98B5-11CF-BB82-00AA00BDCE0B")]
		private interface IHTMLElementRender
		{
			void DrawToDC(IntPtr hdc);
			void SetDocumentPrinter(string bstrPrinterName, IntPtr hdc);
		}
		#endregion

		#region [ interface IViewObject ]
		[ComVisible(true), ComImport()]
		[GuidAttribute("0000010d-0000-0000-C000-000000000046")]
		[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]

		private interface IViewObject
		{
			[return: MarshalAs(UnmanagedType.I4)]
			[PreserveSig]
			int Draw(
				//tagDVASPECT
				[MarshalAs(UnmanagedType.U4)] UInt32 dwDrawAspect,
				int lindex,
				IntPtr pvAspect,
				[In] IntPtr ptd,
				//[MarshalAs(UnmanagedType.Struct)] ref DVTARGETDEVICE ptd,
				IntPtr hdcTargetDev,
				IntPtr hdcDraw,
				[MarshalAs(UnmanagedType.Struct)] ref tagRECT lprcBounds,
				[MarshalAs(UnmanagedType.Struct)] ref tagRECT lprcWBounds,
				IntPtr pfnContinue,
				[MarshalAs(UnmanagedType.U4)] UInt32 dwContinue);
		}
		#endregion

		#region [ Constantes ]
		const String GRID_COL_CHECK_BOX = "colCheck";
		#endregion

		#region [ Atributos ]
		private Form _formChamador = null;

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

		FBoletoHtml fBoletoHtml;
		FConfiguracao fConfiguracao;
		FEmailParametros fEmailParametros;
		#endregion

		#region [ Menu ]
		ToolStripMenuItem menuBoleto;
		ToolStripMenuItem menuBoletoPesquisar;
		ToolStripMenuItem menuBoletoDetalhe;
		ToolStripMenuItem menuBoletoLimpar;
		#endregion

		#region [ Controle da impressão ]
		private int _intImpressaoIdxLinhaGrid = 0;
		private int _intImpressaoNumPagina = 0;
		private String _strImpressaoDataEmissao;
		private int _intQtdeTotalRegistros;
		private decimal _vlTotalRegistros;
		Impressao impressao;
		const String NOME_FONTE_DEFAULT = "Courier New";
		Font fonteTitulo;
		Font fonteListagem;
		Font fonteDataEmissao;
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

		#region [ Colunas Listagem ]
		float ixCliente;
		float wxCliente;
		float ixCnpjCpf;
		float wxCnpjCpf;
		float ixLoja;
		float wxLoja;
		float ixPedido;
		float wxPedido;
		float ixNumeroDocumento;
		float wxNumeroDocumento;
		float ixParcela;
		float wxParcela;
		float ixSituacao;
		float wxSituacao;
		float ixDtVencto;
		float wxDtVencto;
		float ixVlTitulo;
		float wxVlTitulo;
		float ESPACAMENTO_COLUNAS;
		#endregion

		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoConsulta(Form formChamador)
		{
			InitializeComponent();

			_formChamador = formChamador;

			#region [ Menu Boleto ]
			// Menu principal de Boleto
			menuBoleto = new ToolStripMenuItem("&Boleto");
			menuBoleto.Name = "menuBoleto";
			// Pesquisar
			menuBoletoPesquisar = new ToolStripMenuItem("&Pesquisar", null, menuBoletoPesquisar_Click);
			menuBoletoPesquisar.Name = "menuBoletoPesquisar";
			menuBoleto.DropDownItems.Add(menuBoletoPesquisar);
			// Limpar
			menuBoletoLimpar = new ToolStripMenuItem("&Limpar", null, menuBoletoLimpar_Click);
			menuBoletoLimpar.Name = "menuBoletoLimpar";
			menuBoleto.DropDownItems.Add(menuBoletoLimpar);
			// Detalhe
			menuBoletoDetalhe = new ToolStripMenuItem("&Detalhes do Boleto", null, menuBoletoDetalhe_Click);
			menuBoletoDetalhe.Name = "menuBoletoDetalhe";
			menuBoleto.DropDownItems.Add(menuBoletoDetalhe);
			// Adiciona o menu Boleto ao menu principal
			menuPrincipal.Items.Insert(1, menuBoleto);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtDataCargaRetornoInicial.Text = "";
			txtDataCargaRetornoFinal.Text = "";
			txtDataVenctoInicial.Text = "";
			txtDataVenctoFinal.Text = "";
			txtNumNF.Text = "";
			cbOcorrencia.SelectedIndex = -1;
			cbBoletoCedente.SelectedIndex = -1;
			txtNumPedido.Text = "";
			txtValor.Text = "";
			txtNomeCliente.Text = "";
			txtCnpjCpf.Text = "";
			lblTotalizacaoRegistros.Text = "";
			lblTotalizacaoValor.Text = "";
			gridDados.DataSource = null;
			txtDataCargaRetornoInicial.Focus();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Declarações ]
			const int MAX_PERIODO_EM_DIAS = 90;
			DateTime dtVenctoInicial = DateTime.MinValue;
			DateTime dtVenctoFinal = DateTime.MinValue;
			DateTime dtCargaRetornoInicial = DateTime.MinValue;
			DateTime dtCargaRetornoFinal = DateTime.MinValue;
			DateTime dtAuxInicial = DateTime.MinValue;
			DateTime dtAuxFinal = DateTime.MinValue;
			#endregion

			#region [ Período do Vencimento da Parcela ]
			if (txtDataVenctoInicial.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataVenctoInicial.Text))
				{
					avisoErro("Data inválida!!");
					txtDataVenctoInicial.Focus();
					return false;
				}
				else dtVenctoInicial = Global.converteDdMmYyyyParaDateTime(txtDataVenctoInicial.Text);
			}
			
			if (txtDataVenctoFinal.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataVenctoFinal.Text))
				{
					avisoErro("Data inválida!!");
					txtDataVenctoFinal.Focus();
					return false;
				}
				else dtVenctoFinal = Global.converteDdMmYyyyParaDateTime(txtDataVenctoFinal.Text);
			}

			if ((dtVenctoInicial > DateTime.MinValue) && (dtVenctoFinal > DateTime.MinValue))
			{
				if (dtVenctoInicial > dtVenctoFinal)
				{
					avisoErro("A data final do período é anterior à data inicial!!");
					txtDataVenctoFinal.Focus();
					return false;
				}
			}
			#endregion

			#region [ Período da carga do arquivo de retorno ]
			if (txtDataCargaRetornoInicial.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataCargaRetornoInicial.Text))
				{
					avisoErro("Data inválida!!");
					txtDataCargaRetornoInicial.Focus();
					return false;
				}
				else dtCargaRetornoInicial = Global.converteDdMmYyyyParaDateTime(txtDataCargaRetornoInicial.Text);
			}

			if (txtDataCargaRetornoFinal.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataCargaRetornoFinal.Text))
				{
					avisoErro("Data inválida!!");
					txtDataCargaRetornoFinal.Focus();
					return false;
				}
				else dtCargaRetornoFinal = Global.converteDdMmYyyyParaDateTime(txtDataCargaRetornoFinal.Text);
			}

			if ((dtCargaRetornoInicial > DateTime.MinValue) && (dtCargaRetornoFinal > DateTime.MinValue))
			{
				if (dtCargaRetornoInicial > dtCargaRetornoFinal)
				{
					avisoErro("A data final do período é anterior à data inicial!!");
					txtDataCargaRetornoFinal.Focus();
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
				if ((dtVenctoInicial == DateTime.MinValue) && (dtVenctoFinal == DateTime.MinValue) &&
					(dtCargaRetornoInicial == DateTime.MinValue) && (dtCargaRetornoFinal == DateTime.MinValue))
				{
					avisoErro("É necessário informar pelo menos uma das datas para realizar a consulta!!");
					txtDataCargaRetornoInicial.Focus();
					return false;
				}
			}
			#endregion

			#region [ Período de consulta é muito amplo? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if ((dtVenctoInicial > DateTime.MinValue) && (dtVenctoFinal > DateTime.MinValue) &&
					(dtCargaRetornoInicial > DateTime.MinValue) && (dtCargaRetornoFinal > DateTime.MinValue))
				{
					if (dtVenctoInicial > dtCargaRetornoInicial) dtAuxInicial = dtVenctoInicial; else dtAuxInicial = dtCargaRetornoInicial;
					if (dtVenctoFinal < dtCargaRetornoFinal) dtAuxFinal = dtVenctoFinal; else dtAuxFinal = dtCargaRetornoFinal;
					if ((Global.calculaTimeSpanDias(dtAuxFinal - dtAuxInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
					{
						if (!confirma("O período de consulta excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
					}
				}
				else if ((dtVenctoInicial > DateTime.MinValue) && (dtVenctoFinal > DateTime.MinValue))
				{
					if ((Global.calculaTimeSpanDias(dtVenctoFinal - dtVenctoInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
					{
						if (!confirma("O período de consulta excede " + MAX_PERIODO_EM_DIAS.ToString() + " dias!!\nContinua mesmo assim?")) return false;
					}
				}
				else if ((dtCargaRetornoInicial > DateTime.MinValue) && (dtCargaRetornoFinal > DateTime.MinValue))
				{
					if ((Global.calculaTimeSpanDias(dtCargaRetornoFinal - dtCargaRetornoInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
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

		#region [ montaClausulaWhere ]
		private String montaClausulaWhere()
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;

			#region [ Restrição fixa ]
			strAux = " (" +
						" (tB.status <> " + Global.Cte.FIN.CodBoletoStatus.CANCELADO_MANUAL.ToString() + ")" +
						" AND (tBI.status <> " + Global.Cte.FIN.CodBoletoItemStatus.CANCELADO_MANUAL.ToString() + ")" +
						" AND (tBI.status <> " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")" +
						" AND (tBI.status <> " + Global.Cte.FIN.CodBoletoItemStatus.ENVIADO_REMESSA_BANCO.ToString() + ")" +
					 ")";
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
			#endregion

			#region [ Data da carga do arquivo de retorno ]
			if ((txtDataCargaRetornoInicial.Text.Length > 0) && (txtDataCargaRetornoFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtDataCargaRetornoInicial.Text.Equals(txtDataCargaRetornoFinal.Text))
				{
					strAux = " (tBI.ult_data_carga_arq_retorno = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCargaRetornoInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tBI.ult_data_carga_arq_retorno >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCargaRetornoInicial.Text) + ") AND (tBI.ult_data_carga_arq_retorno <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCargaRetornoFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataCargaRetornoInicial.Text.Length > 0) || (txtDataCargaRetornoFinal.Text.Length > 0))
			{
				if (txtDataCargaRetornoInicial.Text.Length > 0)
				{
					strAux = " (tBI.ult_data_carga_arq_retorno = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCargaRetornoInicial.Text) + ")";
				}
				else if (txtDataCargaRetornoFinal.Text.Length > 0)
				{
					strAux = " (tBI.ult_data_carga_arq_retorno = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataCargaRetornoFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Data do vencimento da parcela ]
			if ((txtDataVenctoInicial.Text.Length > 0) && (txtDataVenctoFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtDataVenctoInicial.Text.Equals(txtDataVenctoFinal.Text))
				{
					strAux = " (tBI.dt_vencto = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataVenctoInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tBI.dt_vencto >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataVenctoInicial.Text) + ") AND (tBI.dt_vencto <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataVenctoFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataVenctoInicial.Text.Length > 0) || (txtDataVenctoFinal.Text.Length > 0))
			{
				if (txtDataVenctoInicial.Text.Length > 0)
				{
					strAux = " (tBI.dt_vencto = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataVenctoInicial.Text) + ")";
				}
				else if (txtDataVenctoFinal.Text.Length > 0)
				{
					strAux = " (tBI.dt_vencto = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataVenctoFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Nº NF ]
			if (txtNumNF.Text.Length > 0)
			{
				strAux = " (tB.numero_NF = " + txtNumNF.Text + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Situação ]
			if (cbOcorrencia.SelectedIndex > -1)
			{
				if ((cbOcorrencia.SelectedValue.ToString().Length > 0) &&
					(!cbOcorrencia.SelectedValue.ToString().Equals(Global.Cte.Etc.FLAG_NAO_SETADO.ToString())))
				{
					strAux = " (tBI.ult_identificacao_ocorrencia = '" + cbOcorrencia.SelectedValue.ToString() + "')";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Cedente ]
			if (cbBoletoCedente.SelectedIndex > -1)
			{
				if (cbBoletoCedente.SelectedValue.ToString().Trim().Length > 0)
				{
					strAux = " (tB.id_boleto_cedente = " + cbBoletoCedente.SelectedValue.ToString() + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Nº pedido ]
			if (txtNumPedido.Text.Length > 0)
			{
				#region [ Exibe de um pedido específico ]
				strAux = " (tBI.id IN" +
								"(" +
									"SELECT DISTINCT" +
										" id_boleto_item" +
									" FROM t_FIN_BOLETO_ITEM_RATEIO" +
									" WHERE" +
										" pedido = '" + txtNumPedido.Text.Trim() + "'" +
								")" +
							")";
				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				#endregion
			}
			#endregion

			#region[ Valor ]
			if (txtValor.Text.Trim().Length > 0)
			{
				if (Global.converteNumeroDecimal(txtValor.Text) > 0)
				{
					strAux = Global.sqlFormataDecimal(Global.converteNumeroDecimal(txtValor.Text));
					strAux = " (tBI.valor = " + strAux + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Nome do cliente ]
			if (txtNomeCliente.Text.Trim().Length > 0)
			{
				strAux = " (tB.nome_sacado LIKE '" + BD.CARACTER_CURINGA_TODOS + txtNomeCliente.Text + BD.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}

			#region [ CNPJ/CPF ]
			if (Global.digitos(txtCnpjCpf.Text).Length > 0)
			{
				strAux = " (tB.num_inscricao_sacado = '" + Global.digitos(txtCnpjCpf.Text) + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#endregion

			return sbWhere.ToString();
		}
		#endregion

		#region [ montaSqlConsulta ]
		private String montaSqlConsulta()
		{
			String strWhere;
			String strSql;

			#region [ Monta cláusula Where ]
			strWhere = montaClausulaWhere();
			if (strWhere.Length > 0) strWhere = " WHERE " + strWhere;
			#endregion

			#region [ Monta Select ]
			strSql = "SELECT " +
						" tBI.id AS id_boleto_item," +
						" tB.nome_sacado," +
						" tB.num_inscricao_sacado," +
						" tBI.numero_documento," +
						" tBI.num_parcela," +
						" tB.qtde_parcelas," +
						" tBI.ult_identificacao_ocorrencia," +
						" tBI.ult_motivos_rejeicoes," +
						" tBI.ult_motivo_ocorrencia_19," +
						" tBI.dt_vencto," +
						BD.strSchema + ".ConcatenaPedidosTabelaFinBoletoItemRateio(tBI.id, ', ') AS pedido," +
						" valor" +
					" FROM t_FIN_BOLETO tB" +
						" INNER JOIN t_FIN_BOLETO_ITEM tBI" +
							" ON tB.id=tBI.id_boleto" +
					strWhere +
					" ORDER BY" +
						" tB.nome_sacado," +
						" tBI.id_boleto," +
						" tBI.dt_vencto";
			#endregion

			return strSql;
		}
		#endregion

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			Decimal decTotalizacaoValor = 0;
			int intQtdeRegistros = 0;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
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

				#region [ Monta o SQL da consulta ]
				strSql = montaSqlConsulta();
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Carrega dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					gridDados.SuspendLayout();

					gridDados.Rows.Clear();
					if (dtbConsulta.Rows.Count > 0) gridDados.Rows.Add(dtbConsulta.Rows.Count);

					for (int i = 0; i < dtbConsulta.Rows.Count; i++)
					{
						rowConsulta = dtbConsulta.Rows[i];
						gridDados.Rows[i].Cells["situacao"].Value = Global.montaDescricaoOcorrenciaBoleto(BD.readToString(rowConsulta["ult_identificacao_ocorrencia"]), BD.readToString(rowConsulta["ult_motivos_rejeicoes"]), BD.readToString(rowConsulta["ult_motivo_ocorrencia_19"]));
						gridDados.Rows[i].Cells["id_boleto_item"].Value = BD.readToInt(rowConsulta["id_boleto_item"]).ToString();
						gridDados.Rows[i].Cells["cliente"].Value = BD.readToString(rowConsulta["nome_sacado"]);
						gridDados.Rows[i].Cells["cnpj_cpf_formatado"].Value = Global.formataCnpjCpf(BD.readToString(rowConsulta["num_inscricao_sacado"]));
						gridDados.Rows[i].Cells["pedido"].Value = BD.readToString(rowConsulta["pedido"]);
						gridDados.Rows[i].Cells["num_documento"].Value = BD.readToString(rowConsulta["numero_documento"]);
						gridDados.Rows[i].Cells["num_parcela"].Value = BD.readToByte(rowConsulta["num_parcela"]).ToString() + " / " + BD.readToByte(rowConsulta["qtde_parcelas"]).ToString();
						gridDados.Rows[i].Cells["dt_vencto_formatada"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_vencto"]));
						gridDados.Rows[i].Cells["valor_formatado"].Value = Global.formataMoeda(BD.readToDecimal(rowConsulta["valor"]));

						decTotalizacaoValor += BD.readToDecimal(rowConsulta["valor"]);
						intQtdeRegistros++;
					}
                    #region [ Exibe o grid sem nenhuma linha pré-selecionada ]
                    gridDados.ClearSelection();
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

		#region [ consultaDetalheRegistroSelecionado ]
		private void consultaDetalheRegistroSelecionado()
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vReciboSacadoInstrucoes = new String[6];
			String[] vFichaCompensacaoInstrucoes = new String[6];
			int idxReciboSacadoInstrucoes = 0;
			int idxFichaCompensacaoInstrucoes = 0;
			BoletoCedente boletoCedente;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
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

			#region [ Consistências ]
			if (gridDados.SelectedRows.Count == 0)
			{
				avisoErro("Nenhum boleto foi selecionado!!");
				return;
			}

			if (gridDados.SelectedRows.Count > 1)
			{
				avisoErro("Não é permitida a seleção de múltiplos boletos!!");
				return;
			}
			#endregion

			#region [ Obtém Id do registro ]
			foreach (DataGridViewRow item in gridDados.SelectedRows)
			{
				idBoletoItem = (int)Global.converteInteiro(item.Cells["id_boleto_item"].Value.ToString());
			}
			if (idBoletoItem == 0)
			{
				avisoErro("Não foi possível obter a identificação do registro do boleto!!");
				return;
			}
			#endregion

			#region [ Obtém dados do boleto selecionado ]
			rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
			rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
			#endregion

			#region [ Boleto teve entrada confirmada pela banco? ]
			if (rowBoletoItem.Isdt_entrada_confirmadaNull())
			{
				avisoErro("Boleto não teve a entrada confirmada pelo banco!!");
				return;
			}
			#endregion

			#region [ Obtém os dados do cedente ]
			boletoCedente = BoletoCedenteDAO.getBoletoCedente(rowBoletoPrincipal.id_boleto_cedente);
			if (boletoCedente == null)
			{
				avisoErro("Falha ao obter os dados do cedente!!");
				return;
			}
			#endregion

			#region [ Mensagens do recibo do sacado ]
			for (int i = 0; i < vReciboSacadoInstrucoes.Length; i++)
			{
				vReciboSacadoInstrucoes[i] = "";
			}
			if (rowBoletoItem.valor_por_dia_atraso > 0)
			{
				vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = "APÓS O VENCIMENTO MORA DIA... " + Global.formataMoeda(rowBoletoItem.valor_por_dia_atraso);
				idxReciboSacadoInstrucoes++;
			}
			if (rowBoletoPrincipal.qtde_dias_protesto > 0)
			{
				vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = "APÓS " + rowBoletoPrincipal.qtde_dias_protesto.ToString() + " DIAS DO VENCIMENTO, PROTESTAR O TÍTULO.";
				idxReciboSacadoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_1.Trim().Length > 0)
			{
				vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_1;
				idxReciboSacadoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_2.Trim().Length > 0)
			{
				vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_2;
				idxReciboSacadoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_3.Trim().Length > 0)
			{
				vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_3;
				idxReciboSacadoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_4.Trim().Length > 0)
			{
				vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_4;
				idxReciboSacadoInstrucoes++;
			}
			#endregion

			#region [ Mensagens da ficha de compensação ]
			for (int i = 0; i < vFichaCompensacaoInstrucoes.Length; i++)
			{
				vFichaCompensacaoInstrucoes[i] = "";
			}
			if (rowBoletoItem.valor_por_dia_atraso > 0)
			{
				vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = "APÓS O VENCIMENTO MORA DIA... " + Global.formataMoeda(rowBoletoItem.valor_por_dia_atraso);
				idxFichaCompensacaoInstrucoes++;
			}
			if (rowBoletoPrincipal.qtde_dias_protesto > 0)
			{
				vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = "APÓS " + rowBoletoPrincipal.qtde_dias_protesto.ToString() + " DIAS DO VENCIMENTO, PROTESTAR O TÍTULO.";
				idxFichaCompensacaoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_1.Trim().Length > 0)
			{
				vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_1;
				idxFichaCompensacaoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_2.Trim().Length > 0)
			{
				vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_2;
				idxFichaCompensacaoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_3.Trim().Length > 0)
			{
				vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_3;
				idxFichaCompensacaoInstrucoes++;
			}
			if (rowBoletoPrincipal.mensagem_4.Trim().Length > 0)
			{
				vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_4;
				idxFichaCompensacaoInstrucoes++;
			}
			#endregion

			#region [ Exibe form para visualizar o boleto ]
			fBoletoHtml = new FBoletoHtml(	this,
											rowBoletoPrincipal.email_sacado,
											Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto),
											Global.formataMoeda(rowBoletoItem.valor),
											Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_entrada_confirmada),
											Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_entrada_confirmada),
											rowBoletoItem.numero_documento,
											boletoCedente.nome_empresa,
											boletoCedente.carteira,
											boletoCedente.agencia + '-' + boletoCedente.digito_agencia + '/' + boletoCedente.conta + '-' + boletoCedente.digito_conta,
											boletoCedente.carteira + '/' + rowBoletoItem.nosso_numero + '-' + rowBoletoItem.digito_nosso_numero,
											rowBoletoPrincipal.nome_sacado,
											rowBoletoPrincipal.num_inscricao_sacado,
											rowBoletoPrincipal.endereco_sacado,
											Global.formataCep(rowBoletoPrincipal.cep_sacado) + " - " + rowBoletoPrincipal.cidade_sacado + " - " + rowBoletoPrincipal.uf_sacado,
											rowBoletoItem.linha_digitavel,
											rowBoletoItem.codigo_barras,
											vReciboSacadoInstrucoes[0],
											vReciboSacadoInstrucoes[1],
											vReciboSacadoInstrucoes[2],
											vReciboSacadoInstrucoes[3],
											vReciboSacadoInstrucoes[4],
											vReciboSacadoInstrucoes[5],
											vFichaCompensacaoInstrucoes[0],
											vFichaCompensacaoInstrucoes[1],
											vFichaCompensacaoInstrucoes[2],
											vFichaCompensacaoInstrucoes[3],
											vFichaCompensacaoInstrucoes[4],
											vFichaCompensacaoInstrucoes[5]
											);
			fBoletoHtml.StartPosition = FormStartPosition.Manual;
			fBoletoHtml.Left = this.Left + (this.Width - fBoletoHtml.Width) / 2;
			fBoletoHtml.Top = this.Top + (this.Height - fBoletoHtml.Height) / 2;
			fBoletoHtml.ShowDialog();
			#endregion
		}
		#endregion

		#region [ configuraParametrosEmail ]
		private bool configuraParametrosEmail()
		{
			DialogResult drResultado;
			fConfiguracao = new FConfiguracao();
			fConfiguracao.StartPosition = FormStartPosition.Manual;
			fConfiguracao.Left = this.Left + (this.Width - fConfiguracao.Width) / 2;
			fConfiguracao.Top = this.Top + (this.Height - fConfiguracao.Height) / 2;
			drResultado = fConfiguracao.ShowDialog();

			if (drResultado == DialogResult.OK)
				return true;
			else
				return false;
		}
		#endregion

		#region [ trataBotaoBoletoEmail ]
		private void trataBotaoBoletoEmail()
		{
			#region [ Declarações ]
			int intQtdeAssinalados = 0;
			int idBoletoItem;
			int intScrollRectangleHeight;
			DialogResult drResultado;
			List<BoletoDadosEmailBatch> listaBoletoBatch = new List<BoletoDadosEmailBatch>();
			BoletoDadosEmailBatch boletoDadosEmailBatch;
			BoletoCedente boletoCedente;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			String[] vReciboSacadoInstrucoes = new String[6];
			String[] vFichaCompensacaoInstrucoes = new String[6];
			int idxReciboSacadoInstrucoes;
			int idxFichaCompensacaoInstrucoes;
			String assuntoSelecionado = "Cobrança Bradesco, boleto bancário";
			String destinatarioParaSelecionado = "";
			String destinatarioCopiaSelecionado = "";
			String strCorpoEmail = "";
			SmtpClient smtpCliente;
			MailMessage mailMensagem;
			MailAddress mailAddressFrom;
			MailAddress mailAddressTo;
			MailAddress mailAddressCc;
			MailAddress mailAddressBcc;
			String strDestinatarioPara;
			String strDestinatarioCopia;
			String[] v;
			Attachment attachment;
			MemoryStream msStream;
			List<StreamWriter> listaSwWriter = new List<StreamWriter>();
			DateTime dtHrInicioEspera;
			WebBrowser wb;
			HtmlElement c_codigo_barras_loaded;
			String strOuterHtml;
			bool blnCodigoBarrasLoaded;
			int intTentativas;
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

			#region [ Consistências ]
			for (int i = 0; i < gridDados.Rows.Count; i++)
			{
				if (gridDados.Rows[i].Cells["colCheck"].Value != null)
				{
					if ((bool)gridDados.Rows[i].Cells["colCheck"].Value) intQtdeAssinalados++;
				}
			}
			if (intQtdeAssinalados == 0)
			{
				avisoErro("Nenhum boleto foi selecionado!!");
				return;
			}
			#endregion

			#region [ Consistência dos parâmetros de envio de e-mails ]
			if (Global.Usuario.fin_servidor_smtp_endereco.Trim().Length == 0)
			{
				if (!confirma("É necessário configurar os parâmetros para envio de e-mails!!\nDeseja configurar agora?")) return;
				if (!configuraParametrosEmail())
				{
					avisoErro("Os parâmetros para envio de e-mails não foram configurados corretamente!!");
					return;
				}
			}
			#endregion

			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "obtendo dados dos boletos");

				#region [ Obtém dados dos registros selecionados ]
				for (int i = 0; i < gridDados.Rows.Count; i++)
				{
					if (gridDados.Rows[i].Cells["colCheck"].Value != null)
					{
						if ((bool)gridDados.Rows[i].Cells["colCheck"].Value)
						{
							#region [ Obtém Id do boleto ]
							idBoletoItem = (int)Global.converteInteiro(gridDados.Rows[i].Cells["id_boleto_item"].Value.ToString());
							if (idBoletoItem == 0)
							{
								avisoErro("Não foi possível obter a identificação do registro do boleto!!");
								return;
							}
							#endregion

							boletoDadosEmailBatch = new BoletoDadosEmailBatch();
							boletoDadosEmailBatch.idBoletoItem = idBoletoItem;

							#region [ Obtém dados do boleto selecionado ]
							rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(boletoDadosEmailBatch.idBoletoItem);
							rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(boletoDadosEmailBatch.idBoletoItem);
							#endregion

							#region [ Obtém os dados do cedente ]
							boletoCedente = BoletoCedenteDAO.getBoletoCedente(rowBoletoPrincipal.id_boleto_cedente);
							if (boletoCedente == null)
							{
								avisoErro("Falha ao obter os dados do cedente!!");
								return;
							}
							#endregion

							#region [ Boleto teve entrada confirmada pela banco? ]
							if (rowBoletoItem.Isdt_entrada_confirmadaNull())
							{
								avisoErro("Boleto com nº documento '" + rowBoletoItem.numero_documento.ToString() + "' não teve a entrada confirmada pelo banco!!");
								return;
							}
							#endregion

							#region [ Mensagens do recibo do sacado ]
							idxReciboSacadoInstrucoes = 0;
							idxFichaCompensacaoInstrucoes = 0;

							for (int j = 0; j < vReciboSacadoInstrucoes.Length; j++)
							{
								vReciboSacadoInstrucoes[j] = "";
							}
							if (rowBoletoItem.valor_por_dia_atraso > 0)
							{
								vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = "APÓS O VENCIMENTO MORA DIA... " + Global.formataMoeda(rowBoletoItem.valor_por_dia_atraso);
								idxReciboSacadoInstrucoes++;
							}
							if (rowBoletoPrincipal.qtde_dias_protesto > 0)
							{
								vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = "APÓS " + rowBoletoPrincipal.qtde_dias_protesto.ToString() + " DIAS DO VENCIMENTO, PROTESTAR O TÍTULO.";
								idxReciboSacadoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_1.Trim().Length > 0)
							{
								vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_1;
								idxReciboSacadoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_2.Trim().Length > 0)
							{
								vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_2;
								idxReciboSacadoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_3.Trim().Length > 0)
							{
								vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_3;
								idxReciboSacadoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_4.Trim().Length > 0)
							{
								vReciboSacadoInstrucoes[idxReciboSacadoInstrucoes] = rowBoletoPrincipal.mensagem_4;
								idxReciboSacadoInstrucoes++;
							}
							#endregion

							#region [ Mensagens da ficha de compensação ]
							for (int j = 0; j < vFichaCompensacaoInstrucoes.Length; j++)
							{
								vFichaCompensacaoInstrucoes[j] = "";
							}
							if (rowBoletoItem.valor_por_dia_atraso > 0)
							{
								vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = "APÓS O VENCIMENTO MORA DIA... " + Global.formataMoeda(rowBoletoItem.valor_por_dia_atraso);
								idxFichaCompensacaoInstrucoes++;
							}
							if (rowBoletoPrincipal.qtde_dias_protesto > 0)
							{
								vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = "APÓS " + rowBoletoPrincipal.qtde_dias_protesto.ToString() + " DIAS DO VENCIMENTO, PROTESTAR O TÍTULO.";
								idxFichaCompensacaoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_1.Trim().Length > 0)
							{
								vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_1;
								idxFichaCompensacaoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_2.Trim().Length > 0)
							{
								vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_2;
								idxFichaCompensacaoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_3.Trim().Length > 0)
							{
								vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_3;
								idxFichaCompensacaoInstrucoes++;
							}
							if (rowBoletoPrincipal.mensagem_4.Trim().Length > 0)
							{
								vFichaCompensacaoInstrucoes[idxFichaCompensacaoInstrucoes] = rowBoletoPrincipal.mensagem_4;
								idxFichaCompensacaoInstrucoes++;
							}
							#endregion

							#region [ Adiciona boleto na lista de envio ]
							boletoDadosEmailBatch.emailCliente = rowBoletoPrincipal.email_sacado;
							boletoDadosEmailBatch.dataVencimento = Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto);
							boletoDadosEmailBatch.valorDocumento = Global.formataMoeda(rowBoletoItem.valor);
							boletoDadosEmailBatch.dataProcessamento = Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_entrada_confirmada);
							boletoDadosEmailBatch.dataDocumento = Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_entrada_confirmada);
							boletoDadosEmailBatch.numeroDocumento = rowBoletoItem.numero_documento;
							boletoDadosEmailBatch.nomeCedente = boletoCedente.nome_empresa;
							boletoDadosEmailBatch.carteira = boletoCedente.carteira;
							boletoDadosEmailBatch.agenciaECodigoCedente = boletoCedente.agencia + '-' + boletoCedente.digito_agencia + '/' + boletoCedente.conta + '-' + boletoCedente.digito_conta;
							boletoDadosEmailBatch.nossoNumero = boletoCedente.carteira + '/' + rowBoletoItem.nosso_numero + '-' + rowBoletoItem.digito_nosso_numero;
							boletoDadosEmailBatch.nomeSacado = rowBoletoPrincipal.nome_sacado;
							boletoDadosEmailBatch.numInscricaoSacado = rowBoletoPrincipal.num_inscricao_sacado;
							boletoDadosEmailBatch.enderecoSacado = rowBoletoPrincipal.endereco_sacado;
							boletoDadosEmailBatch.cepCidadeUfSacado = Global.formataCep(rowBoletoPrincipal.cep_sacado) + " - " + rowBoletoPrincipal.cidade_sacado + " - " + rowBoletoPrincipal.uf_sacado;
							boletoDadosEmailBatch.linhaDigitavel = rowBoletoItem.linha_digitavel;
							boletoDadosEmailBatch.codigoBarras = rowBoletoItem.codigo_barras;
							boletoDadosEmailBatch.reciboSacadoInstrucoesLinha1 = vReciboSacadoInstrucoes[0];
							boletoDadosEmailBatch.reciboSacadoInstrucoesLinha2 = vReciboSacadoInstrucoes[1];
							boletoDadosEmailBatch.reciboSacadoInstrucoesLinha3 = vReciboSacadoInstrucoes[2];
							boletoDadosEmailBatch.reciboSacadoInstrucoesLinha4 = vReciboSacadoInstrucoes[3];
							boletoDadosEmailBatch.reciboSacadoInstrucoesLinha5 = vReciboSacadoInstrucoes[4];
							boletoDadosEmailBatch.reciboSacadoInstrucoesLinha6 = vReciboSacadoInstrucoes[5];
							boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha1 = vFichaCompensacaoInstrucoes[0];
							boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha2 = vFichaCompensacaoInstrucoes[1];
							boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha3 = vFichaCompensacaoInstrucoes[2];
							boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha4 = vFichaCompensacaoInstrucoes[3];
							boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha5 = vFichaCompensacaoInstrucoes[4];
							boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha6 = vFichaCompensacaoInstrucoes[5];
							boletoDadosEmailBatch.boletoHtml = new BoletoHtml(
												boletoDadosEmailBatch.dataVencimento,
												boletoDadosEmailBatch.valorDocumento,
												boletoDadosEmailBatch.dataProcessamento,
												boletoDadosEmailBatch.dataDocumento,
												boletoDadosEmailBatch.numeroDocumento,
												boletoDadosEmailBatch.nomeCedente,
												boletoDadosEmailBatch.carteira,
												boletoDadosEmailBatch.agenciaECodigoCedente,
												boletoDadosEmailBatch.nossoNumero,
												boletoDadosEmailBatch.nomeSacado,
                                                boletoDadosEmailBatch.numInscricaoSacado,
                                                boletoDadosEmailBatch.enderecoSacado,
												boletoDadosEmailBatch.cepCidadeUfSacado,
												boletoDadosEmailBatch.linhaDigitavel,
												boletoDadosEmailBatch.codigoBarras,
												boletoDadosEmailBatch.reciboSacadoInstrucoesLinha1,
												boletoDadosEmailBatch.reciboSacadoInstrucoesLinha2,
												boletoDadosEmailBatch.reciboSacadoInstrucoesLinha3,
												boletoDadosEmailBatch.reciboSacadoInstrucoesLinha4,
												boletoDadosEmailBatch.reciboSacadoInstrucoesLinha5,
												boletoDadosEmailBatch.reciboSacadoInstrucoesLinha6,
												boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha1,
												boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha2,
												boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha3,
												boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha4,
												boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha5,
												boletoDadosEmailBatch.fichaCompensacaoInstrucoesLinha6
												);

							listaBoletoBatch.Add(boletoDadosEmailBatch);
							#endregion
						}
					}
				}
				#endregion

				#region [ Consistências ]
				if (listaBoletoBatch.Count == 0)
				{
					avisoErro("Nenhum boleto selecionado foi identificado corretamente!!");
					return;
				}

				for (int i = 0; i < listaBoletoBatch.Count - 1; i++)
				{
					if (listaBoletoBatch[i].numInscricaoSacado.Trim().Length > 0)
					{
						if (!listaBoletoBatch[i].numInscricaoSacado.Trim().Equals(listaBoletoBatch[i + 1].numInscricaoSacado.Trim()))
						{
							avisoErro("Os boletos selecionados são de clientes diferentes!!\n" + listaBoletoBatch[i].nomeSacado.Trim() + " (" + Global.formataCnpjCpf(listaBoletoBatch[i].numInscricaoSacado) + ")\n" + listaBoletoBatch[i + 1].nomeSacado + " (" + Global.formataCnpjCpf(listaBoletoBatch[i + 1].numInscricaoSacado) + ")\n\nNão é possível continuar!!");
							return;
						}
					}
				}
				#endregion

				#region [ Obtém endereços p/ enviar o e-mail ]
				info(ModoExibicaoMensagemRodape.Normal);
				destinatarioParaSelecionado = listaBoletoBatch[0].emailCliente;
				fEmailParametros = new FEmailParametros(Global.Usuario.fin_email_remetente,
														Global.Usuario.fin_display_name_remetente,
														assuntoSelecionado,
														destinatarioParaSelecionado,
														destinatarioCopiaSelecionado);
				fEmailParametros.StartPosition = FormStartPosition.Manual;
				fEmailParametros.Left = this.Left + (this.Width - fEmailParametros.Width) / 2;
				fEmailParametros.Top = this.Top + (this.Height - fEmailParametros.Height) / 2;
				drResultado = fEmailParametros.ShowDialog();

				if (drResultado != DialogResult.OK)
				{
					avisoErro("Envio do e-mail foi cancelado!!");
					return;
				}

				assuntoSelecionado = fEmailParametros.assuntoEmail;
				destinatarioParaSelecionado = fEmailParametros.destinatarioPara;
				destinatarioCopiaSelecionado = fEmailParametros.destinatarioCopia;
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "gerando e-mail");

				#region [ Prepara o e-mail ]
				mailMensagem = new MailMessage();
				smtpCliente = new SmtpClient();

				mailAddressFrom = new MailAddress(Global.Usuario.fin_email_remetente, Global.Usuario.fin_display_name_remetente);
				mailMensagem.From = mailAddressFrom;

				#region[ Preenche destinatário ]
				strDestinatarioPara = destinatarioParaSelecionado.Trim();
				if (strDestinatarioPara.Length > 0)
				{
					strDestinatarioPara = strDestinatarioPara.Replace("\n", " ");
					strDestinatarioPara = strDestinatarioPara.Replace("\r", " ");
					strDestinatarioPara = strDestinatarioPara.Replace(",", " ");
					strDestinatarioPara = strDestinatarioPara.Replace(";", " ");
					v = strDestinatarioPara.Split(' ');
					for (int i = 0; i < v.Length; i++)
					{
						if (v[i].Trim().Length > 0)
						{
							mailAddressTo = new MailAddress(v[i].Trim());
							mailMensagem.To.Add(mailAddressTo);
						}
					}
				}
				#endregion

				#region [ Se houver Cópia Para, preenche campos ]
				strDestinatarioCopia = destinatarioCopiaSelecionado.Trim();
				if (strDestinatarioCopia.Length > 0)
				{
					strDestinatarioCopia = strDestinatarioCopia.Replace("\n", " ");
					strDestinatarioCopia = strDestinatarioCopia.Replace("\r", " ");
					strDestinatarioCopia = strDestinatarioCopia.Replace(",", " ");
					strDestinatarioCopia = strDestinatarioCopia.Replace(";", " ");
					v = strDestinatarioCopia.Split(' ');
					for (int i = 0; i < v.Length; i++)
					{
						if (v[i].Trim().Length > 0)
						{
							mailAddressCc = new MailAddress(v[i].Trim());
							mailMensagem.CC.Add(mailAddressCc);
						}
					}
				}
				#endregion

				#region [ Bcc para o próprio remetente: fica como um comprovante ]
				mailAddressBcc = new MailAddress(Global.Usuario.fin_email_remetente);
				mailMensagem.Bcc.Add(mailAddressBcc);
				#endregion

				#endregion

				#region [ Laço para gerar todos os boletos e anexá-los ao e-mail ]
				for (int i = 0; i < listaBoletoBatch.Count; i++)
				{
					if (i > 0) strCorpoEmail += "\n";
					strCorpoEmail += "\n" +
									 "N. do título...: " + listaBoletoBatch[i].numeroDocumento +
									 "\n" +
									 "Vencimento.....: " + listaBoletoBatch[i].dataVencimento +
									 "\n" +
									 "Valor do título: " + listaBoletoBatch[i].valorDocumento +
									 "\n" +
									 "Linha digitável: " + listaBoletoBatch[i].linhaDigitavel;
				}

				strCorpoEmail = "Data do envio: " + Global.formataDataDdMmYyyyComSeparador(DateTime.Now) +
								"\n\n" +
								"Prezado sacado," +
								"\n\n" +
								listaBoletoBatch[0].nomeSacado + " - " + Global.formataCnpjCpf(listaBoletoBatch[0].numInscricaoSacado) +
								"\n" +
								listaBoletoBatch[0].enderecoSacado +
								"\n" +
								listaBoletoBatch[0].cepCidadeUfSacado +
								"\n\n" +
								"Ref: Boleto(s) de pagamento transmitido via e-mail" +
								"\n\n" +
								"Dados do(s) título(s)" +
								"\n" +
								"=====================" +
								strCorpoEmail +
								"\n\n\n" +
								"Instruções para impressão do boleto bancário:" +
								"\n" +
								"=============================================" +
								"\n\n" +
								"1) Dê um duplo \"click\" sobre o ícone do arquivo em anexo (Arquivo de imagem JPG)" +
								"\n" +
								"2) Imprima o boleto conforme instruções abaixo:" +
								"\n" +
								"   a) Utilize impressora jato de tinta ou laser;" +
								"\n" +
								"   b) Utilize papel de tamanho A4 (210x297mm);" +
								"\n" +
								"   c) Configure a impressora para modo Normal de Impressão;" +
								"\n\n\n" +
								"Atenciosamente," +
								"\n\n" +
								listaBoletoBatch[0].nomeCedente +
								"\n";

				mailMensagem.Subject = assuntoSelecionado;
				mailMensagem.Priority = MailPriority.High;
				mailMensagem.BodyEncoding = Encoding.GetEncoding("Windows-1252");
				mailMensagem.Body = strCorpoEmail;

				#region [ Cria componente WebBrowser para renderizar o html ]
				wb = new WebBrowser();
				wb.ScrollBarsEnabled = false;
				wb.AllowNavigation = true;
				wb.Width = 780;
				wb.Height = 1022;
				#endregion

				#region [ Laço para gerar o anexo para cada um dos boletos ]
				for (int i = 0; i < listaBoletoBatch.Count; i++)
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "gerando e-mail: criando imagem do boleto " + (i + 1).ToString() + "/" + listaBoletoBatch.Count.ToString());

					#region [ Laço de tentativas para renderizar o html ]
					intTentativas = 0;
					do
					{
						intTentativas++;

						#region [ Se após várias tentativas o problema persiste, tenta recriar o componente WebBrowser ]
						if (intTentativas > 10)
						{
							if (wb != null) wb.Dispose();
							wb = new WebBrowser();
							wb.ScrollBarsEnabled = false;
							wb.AllowNavigation = true;
							wb.Width = 780;
							wb.Height = 1022;
						}
						#endregion

						#region [ Carrega o html no WebBrowser ]
						if (wb.Document != null) wb.Document.OpenNew(true);
						wb.DocumentText = listaBoletoBatch[i].boletoHtml.textoBoletoHtml;
						Application.DoEvents();
						#endregion

						#region [ Aguarda o WebBrowser processar o html ]
						dtHrInicioEspera = DateTime.Now;
						while (wb.ReadyState != WebBrowserReadyState.Complete)
						{
							Application.DoEvents();
							Thread.Sleep(200);
							if (Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioEspera) > 120)
							{
								throw new FinanceiroException("Falha ao gerar a imagem do boleto: timeout na renderização do código html!!");
							}
						}
						Application.DoEvents();
						wb.Update();
						Application.DoEvents();
						#endregion

						#region [ O html foi renderizado? ]

						#region [ wb.Document.Body.ScrollRectangle.Height ]
						intScrollRectangleHeight = 0;
						if (wb.Document != null)
						{
							if (wb.Document.Body != null)
							{
								if (wb.Document.Body.ScrollRectangle != null) intScrollRectangleHeight = wb.Document.Body.ScrollRectangle.Height;
							}
						}
						#endregion

						#region [ c_codigo_barras_loaded ]
						blnCodigoBarrasLoaded = false;
						c_codigo_barras_loaded = wb.Document.GetElementById("c_codigo_barras_loaded");
						if (c_codigo_barras_loaded != null)
						{
							strOuterHtml = c_codigo_barras_loaded.OuterHtml;
							if (strOuterHtml != null) blnCodigoBarrasLoaded = strOuterHtml.ToUpper().Contains("VALUE=S");
						}
						#endregion

						if ((intScrollRectangleHeight < 600)
								||
							(!blnCodigoBarrasLoaded))
						{
							if (intTentativas > 100) throw new FinanceiroException("Falha ao gerar a imagem do boleto: falha na renderização do código html!!");
							Application.DoEvents();
						}
						#endregion

					} while ((intScrollRectangleHeight < 600)
								||
							 (!blnCodigoBarrasLoaded));
					#endregion

					#region [ Ajusta o tamanho do componente WebBrowser em função do conteúdo do html renderizado ]
					wb.Width = wb.Document.Body.ScrollRectangle.Width;
					wb.Height = wb.Document.Body.ScrollRectangle.Height;
					#endregion

					#region [ Captura a imagem do html renderizado pelo componente WebBrowser ]
					// Get the view object of the browser
					IViewObject VObject = wb.Document.DomDocument as IViewObject;
					// Construct a bitmap as big as the required image.
					Bitmap bmp = new Bitmap(wb.Document.Body.ClientRectangle.Width, wb.Document.Body.ClientRectangle.Height);
					// The size of the portion of the web page to be captured.
					mshtml.tagRECT SourceRect = new tagRECT();
					SourceRect.left = 0;
					SourceRect.top = 0;
					SourceRect.right = wb.Right;
					SourceRect.bottom = wb.Bottom;

					// The size to render the target image. This can be used to shrink the image to a thumbnail.
					mshtml.tagRECT TargetRect = new tagRECT();
					TargetRect.left = 0;
					TargetRect.top = 0;
					TargetRect.right = wb.Right;
					TargetRect.bottom = wb.Bottom;

					// Draw the web page into the bitmap.
					using (Graphics gr = Graphics.FromImage(bmp))
					{
						IntPtr hdc = gr.GetHdc();
						int hr =
							VObject.Draw((int)DVASPECT.DVASPECT_CONTENT,
								(int)-1, IntPtr.Zero, IntPtr.Zero,
								IntPtr.Zero, hdc, ref TargetRect, ref SourceRect,
								IntPtr.Zero, (uint)0);
						gr.ReleaseHdc();
					}
					#endregion

					#region [ Anexa no email a imagem capturada no formato jpeg ]
					msStream = new MemoryStream();
					bmp.Save(msStream, System.Drawing.Imaging.ImageFormat.Jpeg);
					msStream.Position = 0;
					attachment = new Attachment(msStream, "BOLETO_" + (i + 1).ToString().PadLeft(2, '0') + ".JPG", System.Net.Mime.MediaTypeNames.Image.Jpeg);
					mailMensagem.Attachments.Add(attachment);
					#endregion

					bmp.Dispose();
				}
				#endregion

				#endregion

				#region [ Transmite o e-mail ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "enviando e-mail");

				smtpCliente.Host = Global.Usuario.fin_servidor_smtp_endereco;
				if (Global.Usuario.fin_servidor_smtp_porta > 0) smtpCliente.Port = Global.Usuario.fin_servidor_smtp_porta;
				smtpCliente.Credentials = new System.Net.NetworkCredential(Global.Usuario.fin_usuario_smtp, Global.Usuario.fin_senha_smtp);

				info(ModoExibicaoMensagemRodape.EmExecucao, "transmitindo o e-mail");
				smtpCliente.Send(mailMensagem);
				SystemSounds.Exclamation.Play();
				#endregion
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				return;
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

		#region [ Form FBoletoConsulta ]

		#region [ FBoletoConsulta_Load ]
		private void FBoletoConsulta_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCampos();

				#region [ Combo Ocorrências ]
				cbOcorrencia.DataSource = Global.montaOpcaoBoletoIdentificacaoOcorrencia(Global.eOpcaoIncluirItemTodos.INCLUIR);
				cbOcorrencia.DisplayMember = "descricao";
				cbOcorrencia.ValueMember = "codigo";
				cbOcorrencia.SelectedIndex = -1;
				#endregion

				#region [ Combo Cedente ]
				cbBoletoCedente.ValueMember = "id";
				cbBoletoCedente.DisplayMember = "descricao_formatada";
				cbBoletoCedente.DataSource = ComboDAO.criaDtbBoletoCedenteCombo(ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS);
				cbBoletoCedente.SelectedIndex = -1;
				// Se houver apenas 1 opção, então seleciona
				if ((cbBoletoCedente.Items.Count == 1) && (cbBoletoCedente.SelectedIndex == -1)) cbBoletoCedente.SelectedIndex = 0;
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

		#region [ FBoletoConsulta_Shown ]
		private void FBoletoConsulta_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Permissão de acesso ao módulo ]
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
					{
						btnDetalhe.Enabled = false;
						menuBoletoDetalhe.Enabled = false;
					}
					#endregion

					#region [ Prepara lista de auto complete do campo nome do cliente ]
					txtNomeCliente.AutoCompleteCustomSource.AddRange(FMain.fMain.listaNomeClienteAutoComplete.ToArray());
					#endregion

					#region [ Posiciona foco ]
					txtDataCargaRetornoInicial.Focus();
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

		#region [ FBoletoConsulta_FormClosing ]
		private void FBoletoConsulta_FormClosing(object sender, FormClosingEventArgs e)
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

		#region [ FBoletoConsulta_KeyDown ]
		private void FBoletoConsulta_KeyDown(object sender, KeyEventArgs e)
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

		#region [ txtDataCargaRetornoInicial ]

		#region [ txtDataCargaRetornoInicial_Enter ]
		private void txtDataCargaRetornoInicial_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataCargaRetornoInicial_Leave ]
		private void txtDataCargaRetornoInicial_Leave(object sender, EventArgs e)
		{
			if (txtDataCargaRetornoInicial.Text.Length == 0) return;
			txtDataCargaRetornoInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCargaRetornoInicial.Text);
			if (!Global.isDataOk(txtDataCargaRetornoInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCargaRetornoInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataCargaRetornoInicial_KeyDown ]
		private void txtDataCargaRetornoInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDataCargaRetornoFinal);
		}
		#endregion

		#region [ txtDataCargaRetornoInicial_KeyPress ]
		private void txtDataCargaRetornoInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDataCargaRetornoFinal ]

		#region [ txtDataCargaRetornoFinal_Enter ]
		private void txtDataCargaRetornoFinal_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataCargaRetornoFinal_Leave ]
		private void txtDataCargaRetornoFinal_Leave(object sender, EventArgs e)
		{
			if (txtDataCargaRetornoFinal.Text.Length == 0) return;
			txtDataCargaRetornoFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataCargaRetornoFinal.Text);
			if (!Global.isDataOk(txtDataCargaRetornoFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtDataCargaRetornoFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataCargaRetornoFinal_KeyDown ]
		private void txtDataCargaRetornoFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDataVenctoInicial);
		}
		#endregion

		#region [ txtDataCargaRetornoFinal_KeyPress ]
		private void txtDataCargaRetornoFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDataVenctoInicial ]

		#region [ txtDataVenctoInicial_Enter ]
		private void txtDataVenctoInicial_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataVenctoInicial_Leave ]
		private void txtDataVenctoInicial_Leave(object sender, EventArgs e)
		{
			if (txtDataVenctoInicial.Text.Length == 0) return;
			txtDataVenctoInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataVenctoInicial.Text);
			if (!Global.isDataOk(txtDataVenctoInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtDataVenctoInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataVenctoInicial_KeyDown ]
		private void txtDataVenctoInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDataVenctoFinal);
		}
		#endregion

		#region [ txtDataVenctoInicial_KeyPress ]
		private void txtDataVenctoInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDataVenctoFinal ]

		#region [ txtDataVenctoFinal_Enter ]
		private void txtDataVenctoFinal_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataVenctoFinal_Leave ]
		private void txtDataVenctoFinal_Leave(object sender, EventArgs e)
		{
			if (txtDataVenctoFinal.Text.Length == 0) return;
			txtDataVenctoFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataVenctoFinal.Text);
			if (!Global.isDataOk(txtDataVenctoFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtDataVenctoFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataVenctoFinal_KeyDown ]
		private void txtDataVenctoFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtNumNF);
		}
		#endregion

		#region [ txtDataVenctoFinal_KeyPress ]
		private void txtDataVenctoFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtNumNF ]

		#region [ txtNumNF_Enter ]
		private void txtNumNF_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumNF_Leave ]
		private void txtNumNF_Leave(object sender, EventArgs e)
		{
			txtNumNF.Text = txtNumNF.Text.Trim();
		}
		#endregion

		#region [ txtNumNF_KeyDown ]
		private void txtNumNF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, cbOcorrencia);
		}
		#endregion

		#region [ txtNumNF_KeyPress ]
		private void txtNumNF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ cbOcorrencia ]

		#region [ cbOcorrencia_KeyDown ]
		private void cbOcorrencia_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtNumPedido);
		}
		#endregion

		#endregion

		#region [ cbBoletoCedente ]

		#region [ cbBoletoCedente_KeyDown ]
		private void cbBoletoCedente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtValor);
		}
		#endregion

		#endregion

		#region [ txtNumPedido ]

		#region [ txtNumPedido_Enter ]
		private void txtNumPedido_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumPedido_Leave ]
		private void txtNumPedido_Leave(object sender, EventArgs e)
		{
			String strNumPedido;
			if (txtNumPedido.Text.Length == 0) return;
			strNumPedido = Global.normalizaNumeroPedido(txtNumPedido.Text);
			if (strNumPedido.Length == 0)
			{
				avisoErro("Nº pedido em formato inválido!!");
				txtNumPedido.Focus();
				return;
			}
			txtNumPedido.Text = strNumPedido;
		}
		#endregion

		#region [ txtNumPedido_KeyDown ]
		private void txtNumPedido_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, cbBoletoCedente);
		}
		#endregion

		#region [ txtNumPedido_KeyPress ]
		private void txtNumPedido_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroPedido(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtValor ]

		#region [ txtValor_Enter ]
		private void txtValor_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
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
			Global.trataTextBoxKeyDown(sender, e, txtNomeCliente);
		}
		#endregion

		#region [ txtValor_KeyPress ]
		private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, btnPesquisar);
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

		#region [ gridDados_KeyDown ]
		private void gridDados_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				consultaDetalheRegistroSelecionado();
				return;
			}
		}
		#endregion

		#region [ gridDados_DoubleClick ]
		private void gridDados_DoubleClick(object sender, EventArgs e)
		{
			consultaDetalheRegistroSelecionado();
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

		#region [ menuBoletoPesquisar_Click ]
		private void menuBoletoPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#endregion

		#region [ Detalhe ]

		#region [ btnDetalhe_Click ]
		private void btnDetalhe_Click(object sender, EventArgs e)
		{
			consultaDetalheRegistroSelecionado();
		}
		#endregion

		#region [ menuBoletoDetalhe_Click ]
		private void menuBoletoDetalhe_Click(object sender, EventArgs e)
		{
			consultaDetalheRegistroSelecionado();
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

		#region [ menuBoletoLimpar_Click ]
		private void menuBoletoLimpar_Click(object sender, EventArgs e)
		{
			limpaCampos();
		}
		#endregion

		#endregion

		#region [ Enviar boleto por e-mail ]

		#region [ btnBoletoEmail_Click ]
		private void btnBoletoEmail_Click(object sender, EventArgs e)
		{
			trataBotaoBoletoEmail();
		}
		#endregion

		#endregion

		#region [ btnPrinterDialog ]

		#region [ btnPrinterDialog_Click ]
		private void btnPrinterDialog_Click(object sender, EventArgs e)
		{
			printerDialog();
		}
		#endregion

		#endregion

		#region [ btnPrintPreview ]

		#region [ btnPrintPreview_Click ]
		private void btnPrintPreview_Click(object sender, EventArgs e)
		{
			printPreview();
		}
		#endregion

		#endregion

		#region [ btnImprimir ]

		#region [ btnImprimir_Click ]
		private void btnImprimir_Click(object sender, EventArgs e)
		{
			imprimeConsulta();
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

		#endregion

		#endregion

		#region [ Impressão ]

		#region [ prnDocConsulta_QueryPageSettings ]
		private void prnDocConsulta_QueryPageSettings(object sender, System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			executaQueryPageSettingsListagem(ref sender, ref e);
		}
		#endregion

		#region [ prnDocConsulta_BeginPrint ]
		private void prnDocConsulta_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
		{
			executaBeginPrintListagem(ref sender, ref e);
		}
		#endregion

		#region [ prnDocConsulta_PrintPage ]
		private void prnDocConsulta_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			executaPrintPageListagem(ref sender, ref e);
		}
		#endregion

		#region [ executaQueryPageSettingsListagem ]
		private void executaQueryPageSettingsListagem(ref object sender, ref System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			e.PageSettings.Landscape = true;
		}
		#endregion

		#region [ executaBeginPrintListagem ]
		private void executaBeginPrintListagem(ref object sender, ref System.Drawing.Printing.PrintEventArgs e)
		{
			#region [ Consistência ]
			if (gridDados.Rows.Count == 0)
			{
				e.Cancel = true;
				return;
			}
			#endregion

			_intImpressaoIdxLinhaGrid = 0;
			_intImpressaoNumPagina = 0;
			_strImpressaoDataEmissao = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now);
			_intQtdeTotalRegistros = 0;
			_vlTotalRegistros = 0m;

			prnDocConsulta.DefaultPageSettings.Landscape = true;

			impressao = new Impressao(prnDocConsulta.DefaultPageSettings.Landscape);

			#region [ Prepara elementos de impressão ]
			fonteTitulo = new Font(NOME_FONTE_DEFAULT, 10f, FontStyle.Bold);
			fonteListagem = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteDataEmissao = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			fonteNumPagina = new Font(NOME_FONTE_DEFAULT, 7f, FontStyle.Bold);
			brushPadrao = new SolidBrush(Color.Black);
			penTracoTitulo = new Pen(brushPadrao, .5f);
			penTracoPontilhado = Impressao.criaPenTracoPontilhado();
			#endregion
		}
		#endregion

		#region [ executaPrintPageListagem ]
		private void executaPrintPageListagem(ref object sender, ref System.Drawing.Printing.PrintPageEventArgs e)
		{
			#region [ Declarações ]
			float cx;
			float cy;
			float hMax;
			RectangleF r;
			int idBoletoItem;
			String[] v;
			String strAux;
			String strTexto;
			String strIdBoletoItem;
			String strPedido;
			String strLoja;
			int intLinhasImpressasNestaPagina = 0;
			decimal vlTitulo;
			List<String> listaPedidoLoja;
			BoletoCedente boletoCedente;
			Global.OpcaoBoletoIdentificacaoOcorrencia opcaoOcorrencia;
			#endregion

			#region [ Consistência ]
			if (gridDados.Rows.Count == 0)
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
				prnDocConsulta.DocumentName = "Consulta de Boletos";
				cxInicio = impressao.getLeftMarginInMm(e);
				larguraUtil = impressao.getWidthInMm(e);
				cxFim = cxInicio + larguraUtil;
				cyInicio = impressao.getTopMarginInMm(e);
				alturaUtil = impressao.getHeightInMm(e);
				cyFim = cyInicio + alturaUtil;
				cyRodapeNumPagina = cyFim - fonteNumPagina.GetHeight(e.Graphics) - 1;
				#endregion

				#region [ Layout das colunas da listagem ]
				ESPACAMENTO_COLUNAS = 2f;
				wxCliente = 40f;
				wxCnpjCpf = 30f;
				wxLoja = 10f;
				wxPedido = 15f;
				wxNumeroDocumento = 19f;
				wxParcela = 12f;
				wxDtVencto = 16f;
				wxVlTitulo = 23f;
				wxSituacao = larguraUtil
							 - wxCliente				// A 1ª coluna não tem espaçamento
							 - ESPACAMENTO_COLUNAS		// Espaçamento da própria coluna "Situação"
							 - ESPACAMENTO_COLUNAS - wxCnpjCpf
							 - ESPACAMENTO_COLUNAS - wxLoja
							 - ESPACAMENTO_COLUNAS - wxPedido
							 - ESPACAMENTO_COLUNAS - wxNumeroDocumento
							 - ESPACAMENTO_COLUNAS - wxParcela
							 - ESPACAMENTO_COLUNAS - wxDtVencto
							 - ESPACAMENTO_COLUNAS - wxVlTitulo;

				ixCliente = cxInicio;
				ixCnpjCpf = ixCliente + wxCliente + ESPACAMENTO_COLUNAS;
				ixLoja = ixCnpjCpf + wxCnpjCpf + ESPACAMENTO_COLUNAS;
				ixPedido = ixLoja + wxLoja + ESPACAMENTO_COLUNAS;
				ixNumeroDocumento = ixPedido + wxPedido + ESPACAMENTO_COLUNAS;
				ixParcela = ixNumeroDocumento + wxNumeroDocumento + ESPACAMENTO_COLUNAS;
				ixSituacao = ixParcela + wxParcela + ESPACAMENTO_COLUNAS;
				ixDtVencto = ixSituacao + wxSituacao + ESPACAMENTO_COLUNAS;
				ixVlTitulo = ixDtVencto + wxDtVencto + ESPACAMENTO_COLUNAS;
				#endregion
			}

			cy = cyInicio;

			#region [ Título ]
			strTexto = "CONSULTA DE BOLETOS";
			fonteAtual = fonteTitulo;
			cx = cxInicio + (larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			cy += 1f;
			#endregion

			#region [ Informações no cabeçalho ]

			#region [ Data da emissão ]
			strTexto = "Emissão: " + _strImpressaoDataEmissao;
			fonteAtual = fonteListagem;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Filtros: Data carga do retorno / Vencimento / Nº NF ]
			strTexto = "Carga do Retorno: " +
					   ((txtDataCargaRetornoInicial.Text.Trim().Length > 0) ? txtDataCargaRetornoInicial.Text : "N.I.") +
					   " a " +
					   ((txtDataCargaRetornoFinal.Text.Trim().Length > 0) ? txtDataCargaRetornoFinal.Text : "N.I.");
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Vencimento: " +
					   ((txtDataVenctoInicial.Text.Trim().Length > 0) ? txtDataVenctoInicial.Text : "N.I.") +
					   " a " +
					   ((txtDataVenctoFinal.Text.Trim().Length > 0) ? txtDataVenctoFinal.Text : "N.I.");
			cx = cxInicio + (larguraUtil / 3);
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Nº NF: " +
					   ((txtNumNF.Text.Trim().Length > 0) ? txtNumNF.Text : "N.I.");
			cx = cxInicio + (2 * (larguraUtil / 3));
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Nº Pedido / Valor / CNPJ/CPF ]
			strTexto = "Nº Pedido: " +
					   ((txtNumPedido.Text.Trim().Length > 0) ? txtNumPedido.Text : "N.I.");
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil / 3);
			strTexto = "Valor (" + Global.Cte.Etc.SIMBOLO_MONETARIO + "): " +
					   ((txtValor.Text.Trim().Length > 0) ? txtValor.Text : "N.I.");
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "CNPJ/CPF: " +
					   ((txtCnpjCpf.Text.Trim().Length > 0) ? txtCnpjCpf.Text : "N.I.");
			cx = cxInicio + (2 * (larguraUtil / 3));
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion
			
			#region [ Ocorrência / Cedente ]

			#region [ Ocorrência ]
			if (cbOcorrencia.SelectedIndex == -1)
			{
				strTexto = "N.I.";
			}
			else
			{
				opcaoOcorrencia = (Global.OpcaoBoletoIdentificacaoOcorrencia)cbOcorrencia.SelectedItem;
				strTexto = opcaoOcorrencia.descricao;
			}
			strTexto = "Ocorrência: " + strTexto;
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			#region [ Cedente ]
			if (cbBoletoCedente.SelectedIndex == -1)
			{
				strTexto = "N.I.";
			}
			else
			{
				boletoCedente = BoletoCedenteDAO.getBoletoCedente((int)Global.converteInteiro(cbBoletoCedente.SelectedValue.ToString()));
				strTexto = boletoCedente.apelido.ToUpper();
			}
			strTexto = "Cedente: " + strTexto;
			cx = cxInicio + (2 * (larguraUtil / 3));
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Nome do Cliente ]
			strTexto = "Nome Cliente: " +
					   ((txtNomeCliente.Text.Trim().Length > 0) ? txtNomeCliente.Text : "N.I.");
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Títulos da listagem ]
			cy += .5f;
			fonteAtual = fonteListagem;
			strTexto = "CLIENTE";
			cx = ixCliente;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "CNPJ/CPF";
			cx = ixCnpjCpf;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "LOJA";
			cx = ixLoja;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "PEDIDO";
			cx = ixPedido;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Nº DOC";
			cx = ixNumeroDocumento;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "PARC";
			cx = ixParcela;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "SITUAÇÃO";
			cx = ixSituacao;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "DT VENCTO";
			cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "VALOR";
			cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cy += fonteAtual.GetHeight(e.Graphics);
			cy += .5f;
			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Laço para listagem ]
			while (((cy + fonteListagem.GetHeight(e.Graphics)) < (cyRodapeNumPagina - 5)) &&
				   (_intImpressaoIdxLinhaGrid < gridDados.Rows.Count))
			{
				#region [ Consulta BD para obter pedido+loja ]
				strPedido = "";
				strLoja = "";
				strIdBoletoItem = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["id_boleto_item"].Value.ToString().Trim();
				if (strIdBoletoItem.Length > 0)
				{
					idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
					listaPedidoLoja = BoletoDAO.obtemBoletoInformacaoPedidoLoja(idBoletoItem);
					for (int i = 0; i < listaPedidoLoja.Count; i++)
					{
						strAux = listaPedidoLoja[i];
						if (strAux == null) continue;
						if (strAux.Length == 0) continue;
						v = strAux.Split('=');
						if (strPedido.Length > 0) strPedido += ", ";
						strPedido += v[0];
						if (strLoja.Length > 0) strLoja += ", ";
						strLoja += v[1];
					}
				}
				#endregion

				hMax = fonteListagem.GetHeight(e.Graphics);

				#region [ Cliente ]
				cx = ixCliente;
				r = new RectangleF(ixCliente, cy, wxCliente, 20);
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["cliente"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxCliente).Height);
				#endregion

				#region [ CNPJ/CPF ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["cnpj_cpf_formatado"].Value.ToString();
				cx = ixCnpjCpf;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Loja ]
				cx = ixLoja;
				r = new RectangleF(ixLoja, cy, wxLoja, 20);
				strTexto = strLoja;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxLoja).Height);
				#endregion

				#region [ Pedido ]
				cx = ixPedido;
				r = new RectangleF(ixPedido, cy, wxPedido, 20);
				strTexto = strPedido;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxPedido).Height);
				#endregion

				#region [ Nº Documento ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["num_documento"].Value.ToString();
				cx = ixNumeroDocumento;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Parcela ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["num_parcela"].Value.ToString();
				strTexto = strTexto.Replace(" ", "");
				cx = ixParcela;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Situação ]
				cx = ixSituacao;
				r = new RectangleF(ixSituacao, cy, wxSituacao, 20);
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["situacao"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxSituacao).Height);
				#endregion

				#region [ Data do vencimento ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["dt_vencto_formatada"].Value.ToString();
				cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Valor do título ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["valor_formatado"].Value.ToString();
				vlTitulo = Global.converteNumeroDecimal(strTexto);
				cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				cy += hMax;

				_intQtdeTotalRegistros++;
				_vlTotalRegistros += vlTitulo;

				intLinhasImpressasNestaPagina++;
				_intImpressaoIdxLinhaGrid++;

				#region [ Na última linha não imprime o tracejado ]
				if (_intImpressaoIdxLinhaGrid < gridDados.Rows.Count)
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
			if (_intImpressaoIdxLinhaGrid < gridDados.Rows.Count)
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
					fonteAtual = fonteListagem;
					cx = cxInicio;
					strTexto = "TOTAL";
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataInteiro(_intQtdeTotalRegistros) + " registro(s)";
					cx = ixCnpjCpf;
					e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

					strTexto = Global.formataMoeda(_vlTotalRegistros);
					cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
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

		#endregion
	}

	#region [ BoletoDadosEmailBatch ]
	class BoletoDadosEmailBatch
	{
		public int idBoletoItem;
		public String emailCliente;
		public String dataVencimento;
		public String valorDocumento;
		public String dataProcessamento;
		public String dataDocumento;
		public String numeroDocumento;
		public String nomeCedente;
		public String carteira;
		public String agenciaECodigoCedente;
		public String nossoNumero;
		public String nomeSacado;
		public String numInscricaoSacado;
		public String enderecoSacado;
		public String cepCidadeUfSacado;
		public String linhaDigitavel;
		public String codigoBarras;
		public String reciboSacadoInstrucoesLinha1;
		public String reciboSacadoInstrucoesLinha2;
		public String reciboSacadoInstrucoesLinha3;
		public String reciboSacadoInstrucoesLinha4;
		public String reciboSacadoInstrucoesLinha5;
		public String reciboSacadoInstrucoesLinha6;
		public String fichaCompensacaoInstrucoesLinha1;
		public String fichaCompensacaoInstrucoesLinha2;
		public String fichaCompensacaoInstrucoesLinha3;
		public String fichaCompensacaoInstrucoesLinha4;
		public String fichaCompensacaoInstrucoesLinha5;
		public String fichaCompensacaoInstrucoesLinha6;
		public BoletoHtml boletoHtml;
	}
	#endregion
}
