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
	public partial class FBoletoOcorrencias : Financeiro.FModelo
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

		private bool _atualizacaoAutomaticaPesquisaEmAndamento = false;
		FBoletoTrataOcorrenciaCepInvalido fBoletoTrataOcorrenciaCepInvalido;
		FBoletoTrataOcorrenciaValaComum fBoletoTrataOcorrenciaValaComum;
		#endregion

		#region [ Menu ]
		ToolStripMenuItem menuOcorrencia;
		ToolStripMenuItem menuOcorrenciaPesquisar;
		ToolStripMenuItem menuOcorrenciaTratar;
		ToolStripMenuItem menuOcorrenciaLimpar;
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
		float ixDataOcorrencia;
		float wxDataOcorrencia;
		float ixCliente;
		float wxCliente;
		float ixNumeroDocumento;
		float wxNumeroDocumento;
		float ixDtVencto;
		float wxDtVencto;
		float ixVlTitulo;
		float wxVlTitulo;
		float ixLoja;
		float wxLoja;
		float ixPedido;
		float wxPedido;
		float ixOcorrencia;
		float wxOcorrencia;
		float ixOcorrenciaObs;
		float wxOcorrenciaObs;
		float ESPACAMENTO_COLUNAS;
		#endregion

		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoOcorrencias()
		{
			InitializeComponent();

			#region [ Menu Boleto ]
			// Menu principal de Ocorrências
			menuOcorrencia = new ToolStripMenuItem("&Ocorrências");
			menuOcorrencia.Name = "menuOcorrencia";
			// Pesquisar
			menuOcorrenciaPesquisar = new ToolStripMenuItem("&Pesquisar", null, menuOcorrenciaPesquisar_Click);
			menuOcorrenciaPesquisar.Name = "menuOcorrenciaPesquisar";
			menuOcorrencia.DropDownItems.Add(menuOcorrenciaPesquisar);
			// Limpar
			menuOcorrenciaLimpar = new ToolStripMenuItem("&Limpar", null, menuOcorrenciaLimpar_Click);
			menuOcorrenciaLimpar.Name = "menuOcorrenciaLimpar";
			menuOcorrencia.DropDownItems.Add(menuOcorrenciaLimpar);
			// Tratar
			menuOcorrenciaTratar = new ToolStripMenuItem("&Tratar Ocorrência", null, menuOcorrenciaTratar_Click);
			menuOcorrenciaTratar.Name = "menuOcorrenciaTratar";
			menuOcorrencia.DropDownItems.Add(menuOcorrenciaTratar);
			// Adiciona o menu Boleto ao menu principal
			menuPrincipal.Items.Insert(1, menuOcorrencia);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtDataInicial.Text = "";
			txtDataFinal.Text = "";
			txtNumDocumento.Text = "";
			txtValor.Text = "";
			txtNossoNumero.Text = "";
			lblTotalizacaoRegistros.Text = "";
			gridDados.Rows.Clear();
			cbOcorrencia.SelectedIndex = -1;
			cbBoletoCedente.SelectedIndex = -1;
			txtDataInicial.Focus();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Declarações ]
			const int MAX_PERIODO_EM_DIAS = 90;
			DateTime dtInicial = DateTime.MinValue;
			DateTime dtFinal = DateTime.MinValue;
			#endregion

			#region [ Período de consulta ]
			if (txtDataInicial.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataInicial.Text))
				{
					avisoErro("Data inválida!!");
					txtDataInicial.Focus();
					return false;
				}
				else dtInicial = Global.converteDdMmYyyyParaDateTime(txtDataInicial.Text);
			}
			
			if (txtDataFinal.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataFinal.Text))
				{
					avisoErro("Data inválida!!");
					txtDataFinal.Focus();
					return false;
				}
				else dtFinal = Global.converteDdMmYyyyParaDateTime(txtDataFinal.Text);
			}

			if ((dtInicial > DateTime.MinValue) && (dtFinal > DateTime.MinValue))
			{
				if (dtInicial > dtFinal)
				{
					avisoErro("A data final do período é anterior à data inicial!!");
					txtDataFinal.Focus();
					return false;
				}
			}
			#endregion

			#region [ Alguma data foi informada? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if ((dtInicial == DateTime.MinValue) && (dtFinal == DateTime.MinValue))
				{
					avisoErro("É necessário informar pelo menos uma das datas para realizar a consulta!!");
					txtDataInicial.Focus();
					return false;
				}
			}
			#endregion

			#region [ Período de consulta é muito amplo? ]
			if (Global.Cte.FIN.FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA)
			{
				if ((dtInicial > DateTime.MinValue) && (dtFinal > DateTime.MinValue))
				{
					if ((Global.calculaTimeSpanDias(dtFinal - dtInicial) > MAX_PERIODO_EM_DIAS) && (MAX_PERIODO_EM_DIAS > 0))
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
			strAux = " (tBO.st_ocorrencia_tratada = " + Global.Cte.FIN.CodBoletoOcorrenciaStOcorrenciaTratada.NAO_TRATADA.ToString() + ")";
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
			#endregion

			#region [ Período de consulta ]
			if ((txtDataInicial.Text.Length > 0) && (txtDataFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtDataInicial.Text.Equals(txtDataFinal.Text))
				{
					strAux = " (tBO.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tBO.dt_cadastro >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataInicial.Text) + ") AND (tBO.dt_cadastro <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataInicial.Text.Length > 0) || (txtDataFinal.Text.Length > 0))
			{
				if (txtDataInicial.Text.Length > 0)
				{
					strAux = " (tBO.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataInicial.Text) + ")";
				}
				else if (txtDataFinal.Text.Length > 0)
				{
					strAux = " (tBO.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Somente com divergência de valor ]
			if (ckb_somente_divergencia_valor.Checked)
			{
				strAux = " (tBO.st_divergencia_valor <> 0)";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Valor ]
			if (Global.converteNumeroDecimal(txtValor.Text) > 0)
			{
				strAux = " (tBO.vl_titulo = " + Global.sqlFormataDecimal(Global.converteNumeroDecimal(txtValor.Text)) + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Nº Documento ]
			if (txtNumDocumento.Text.Length > 0)
			{
				strAux = " (tBO.numero_documento = '" + txtNumDocumento.Text + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Ocorrência ]
			if (cbOcorrencia.SelectedIndex > -1)
			{
				if ((cbOcorrencia.SelectedValue.ToString().Length > 0) &&
					(!cbOcorrencia.SelectedValue.ToString().Equals(Global.Cte.Etc.FLAG_NAO_SETADO.ToString())))
				{
					strAux = " (tBO.identificacao_ocorrencia = '" + cbOcorrencia.SelectedValue.ToString() + "')";
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
					strAux = " (tBO.id_boleto_cedente = " + cbBoletoCedente.SelectedValue.ToString() + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Nosso Número ]
			if (Global.digitos(txtNossoNumero.Text).Length > 0)
			{
				strAux = " (tBO.nosso_numero = '" + Global.digitos(txtNossoNumero.Text) + "')";
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
			#endregion

			#region [ Monta cláusula Where ]
			strWhere = montaClausulaWhere();
			if (strWhere.Length > 0) strWhere = " WHERE " + strWhere;
			#endregion

			#region [ Monta Select ]
			strSql = "SELECT " +
						" tBO.id AS id_boleto_ocorrencia," +
						" tBO.id_boleto_item," +
						" tBO.id_boleto," +
						" tBO.id_boleto_cedente," +
						" tBO.dt_cadastro," +
						" tBO.numero_documento," +
						" tBO.nosso_numero," +
						" tBO.digito_nosso_numero," +
						" tBO.dt_vencto," +
						" tBO.vl_titulo," +
						" tBO.identificacao_ocorrencia," +
						" tBO.motivos_rejeicoes," +
						" tBO.motivo_ocorrencia_19," +
						" tBO.obs_ocorrencia," +
						" tBO.registro_arq_retorno," +
						" tB.nome_sacado," +
						" tB.num_inscricao_sacado" +
					" FROM t_FIN_BOLETO_OCORRENCIA tBO" +
						" LEFT JOIN t_FIN_BOLETO tB" +
							" ON (tBO.id_boleto=tB.id)" +
					strWhere +
					" ORDER BY" +
						" tBO.dt_cadastro," +
						" tBO.identificacao_ocorrencia," +
						" tBO.id";
			#endregion

			return strSql;
		}
		#endregion

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			int intQtdeRegistros = 0;
			String strSql;
			String strCnpjCpf;
			String strCliente;
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
				if (!_atualizacaoAutomaticaPesquisaEmAndamento)
				{
					if (!consisteCampos()) return false;
				}
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

				#region [ Exibição dos dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					gridDados.SuspendLayout();

					#region [ Carrega os dados no grid ]
					gridDados.Rows.Clear();
					if (dtbConsulta.Rows.Count > 0) gridDados.Rows.Add(dtbConsulta.Rows.Count);
					for (int i = 0; i < dtbConsulta.Rows.Count; i++)
					{
						rowConsulta = dtbConsulta.Rows[i];
						gridDados.Rows[i].Cells["id_boleto_ocorrencia"].Value = BD.readToInt(rowConsulta["id_boleto_ocorrencia"]);
						gridDados.Rows[i].Cells["id_boleto_item"].Value = BD.readToInt(rowConsulta["id_boleto_item"]);
						gridDados.Rows[i].Cells["id_boleto"].Value = BD.readToInt(rowConsulta["id_boleto"]);
						gridDados.Rows[i].Cells["id_boleto_cedente"].Value = BD.readToInt(rowConsulta["id_boleto_cedente"]);
						gridDados.Rows[i].Cells["dt_cadastro"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_cadastro"]));
						gridDados.Rows[i].Cells["numero_documento"].Value = BD.readToString(rowConsulta["numero_documento"]);
						gridDados.Rows[i].Cells["dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["dt_vencto"]));
						gridDados.Rows[i].Cells["vl_titulo"].Value = Global.formataMoeda(BD.readToDecimal(rowConsulta["vl_titulo"]));
						gridDados.Rows[i].Cells["identificacao_ocorrencia"].Value = BD.readToString(rowConsulta["identificacao_ocorrencia"]);
						gridDados.Rows[i].Cells["ocorrencia"].Value = Global.montaDescricaoOcorrenciaBoleto(BD.readToString(rowConsulta["identificacao_ocorrencia"]), BD.readToString(rowConsulta["motivos_rejeicoes"]), BD.readToString(rowConsulta["motivo_ocorrencia_19"]));
						gridDados.Rows[i].Cells["obs"].Value = BD.readToString(rowConsulta["obs_ocorrencia"]);

						strCnpjCpf = Global.formataCnpjCpf(BD.readToString(rowConsulta["num_inscricao_sacado"]));
						if (strCnpjCpf.Length > 0) strCnpjCpf = " (" + strCnpjCpf + ")";
						strCliente = BD.readToString(rowConsulta["nome_sacado"]) + strCnpjCpf;
						gridDados.Rows[i].Cells["cliente"].Value = strCliente;

						gridDados.Rows[i].Cells["registro_arq_retorno"].Value = BD.readToString(rowConsulta["registro_arq_retorno"]);
						intQtdeRegistros++;
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

		#region [ trataOcorrenciaSelecionada ]
		private void trataOcorrenciaSelecionada()
		{
			#region [ Declarações ]
			String strIdentificacaoOcorrencia = "";
			#endregion

			#region [ Consistência ]
			if (gridDados.SelectedRows.Count == 0)
			{
				avisoErro("Nenhum registro foi selecionado!!");
				return;
			}

			if (gridDados.SelectedRows.Count > 1)
			{
				avisoErro("Não é permitida a seleção de múltiplos registros!!");
				return;
			}
			#endregion

			#region [ Obtém a ocorrência ]
			foreach (DataGridViewRow item in gridDados.SelectedRows)
			{
				strIdentificacaoOcorrencia = item.Cells["identificacao_ocorrencia"].Value.ToString();
			}
			#endregion

			#region [ Verifica qual o tipo da ocorrência e fornece um tratamento adequado ]
			if (strIdentificacaoOcorrencia.Equals("24"))
			{
				trataCorrecaoCep();
			}
			else
			{
				trataMarcarComoJaTratada();
			}
			#endregion
		}
		#endregion

		#region [ trataCorrecaoCep ]
		private void trataCorrecaoCep()
		{
			#region [ Declarações ]
			DialogResult drResultado;
			int intIdBoletoOcorrencia = 0;
			int intIdBoleto = 0;
			int intIdBoletoItem = 0;
			int intIdBoletoCedente = 0;
			String strIdentificacaoOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			String enderecoCorrigido;
			String bairroCorrigido;
			String cepCorrigido;
			String cidadeCorrigido;
			String ufCorrigido;
			String strMsgErro = "";
			String strDescricaoLog;
			String strMsgErroLog = "";
			bool blnSucesso = false;
			FinLog finLog = new FinLog();
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

				#region [ Consistência ]
				if (gridDados.SelectedRows.Count == 0)
				{
					avisoErro("Nenhum registro foi selecionado!!");
					return;
				}

				if (gridDados.SelectedRows.Count > 1)
				{
					avisoErro("Não é permitida a seleção de múltiplos registros!!");
					return;
				}
				#endregion

				#region [ Obtém Id do registro ]
				foreach (DataGridViewRow item in gridDados.SelectedRows)
				{
					strIdentificacaoOcorrencia = item.Cells["identificacao_ocorrencia"].Value.ToString();
					intIdBoletoOcorrencia = (int)Global.converteInteiro(item.Cells["id_boleto_ocorrencia"].Value.ToString());
					intIdBoleto = (int)Global.converteInteiro(item.Cells["id_boleto"].Value.ToString());
					intIdBoletoItem = (int)Global.converteInteiro(item.Cells["id_boleto_item"].Value.ToString());
					intIdBoletoCedente = (int)Global.converteInteiro(item.Cells["id_boleto_cedente"].Value.ToString());
				}
				#endregion

				#region [ É ocorrência de CEP irregular? ]
				if (!strIdentificacaoOcorrencia.Equals("24"))
				{
					trataMarcarComoJaTratada();
					return;
				}
				#endregion

				#region [ Id do registro é válida? ]
				if (intIdBoletoOcorrencia == 0)
				{
					avisoErro("Falha ao obter o nº identificação do registro da ocorrência!!");
					return;
				}

				if (intIdBoleto == 0)
				{
					avisoErro("Falha ao obter o nº identificação do registro do boleto associado a esta ocorrência!!");
					return;
				}
				#endregion

				#region [ Recupera dados do endereço ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoleto(intIdBoleto);
				if (rowBoletoPrincipal == null)
				{
					avisoErro("Falha ao recuperar os dados do boleto!!");
					return;
				}
				#endregion

				#region [ Exibe painel p/ reeditar o endereço ]
				fBoletoTrataOcorrenciaCepInvalido = new FBoletoTrataOcorrenciaCepInvalido(
																this,
																intIdBoletoCedente,
																rowBoletoPrincipal.nome_sacado,
																rowBoletoPrincipal.num_inscricao_sacado,
																rowBoletoPrincipal.endereco_sacado,
																rowBoletoPrincipal.bairro_sacado,
																rowBoletoPrincipal.cep_sacado,
																rowBoletoPrincipal.cidade_sacado,
																rowBoletoPrincipal.uf_sacado);
				fBoletoTrataOcorrenciaCepInvalido.StartPosition = FormStartPosition.Manual;
				fBoletoTrataOcorrenciaCepInvalido.Left = this.Left + (this.Width - fBoletoTrataOcorrenciaCepInvalido.Width) / 2;
				fBoletoTrataOcorrenciaCepInvalido.Top = this.Top + (this.Height - fBoletoTrataOcorrenciaCepInvalido.Height) / 2;
				drResultado = fBoletoTrataOcorrenciaCepInvalido.ShowDialog();
				if (drResultado != DialogResult.OK) return;
				#endregion

				#region [ Altera o endereço e reseta o status p/ poder ser enviado novamente no arquivo remessa ]
				enderecoCorrigido = fBoletoTrataOcorrenciaCepInvalido.enderecoCorrigido;
				bairroCorrigido = fBoletoTrataOcorrenciaCepInvalido.bairroCorrigido;
				cepCorrigido = fBoletoTrataOcorrenciaCepInvalido.cepCorrigido;
				cidadeCorrigido = fBoletoTrataOcorrenciaCepInvalido.cidadeCorrigido;
				ufCorrigido = fBoletoTrataOcorrenciaCepInvalido.ufCorrigido;

				BD.iniciaTransacao();
				try
				{
					if (!BoletoDAO.corrigeBoletoOcorrencia24CepIrregular(
													Global.Usuario.usuario,
													intIdBoleto,
													enderecoCorrigido,
													bairroCorrigido,
													cepCorrigido,
													cidadeCorrigido,
													ufCorrigido,
													ref strMsgErro))
					{
						throw new Exception("Falha ao atualizar os dados do endereço no banco de dados!!\n\n" + strMsgErro);
					}

					if (!BoletoDAO.marcaBoletoOcorrenciasComoJaTratadasByIdBoleto(
													Global.Usuario.usuario,
													intIdBoleto,
													"Endereço corrigido para: " + Global.formataEndereco(enderecoCorrigido, "", "", bairroCorrigido, cidadeCorrigido, ufCorrigido, cepCorrigido),
													ref strMsgErro))
					{
						throw new Exception("Falha ao marcar a ocorrência como já tratada!!\n\n" + strMsgErro);
					}

					#region [ Grava o log no BD ]
					strDescricaoLog = "Tratamento em 'Boleto - Ocorrências' com correção de endereço: ocorrência 24 (CEP irregular)" +
									  " \n" + "t_FIN_BOLETO.id=" + intIdBoleto.ToString() +
									  " \n" + "t_FIN_BOLETO_ITEM.id=" + intIdBoletoItem.ToString() +
									  " \n" + "t_FIN_BOLETO_OCORRENCIA.id=" + intIdBoletoOcorrencia.ToString();
					if (!rowBoletoPrincipal.endereco_sacado.Equals(enderecoCorrigido)) strDescricaoLog += " \n" + "Endereço: " + rowBoletoPrincipal.endereco_sacado + " -> " + enderecoCorrigido;
					if (!rowBoletoPrincipal.bairro_sacado.Equals(bairroCorrigido)) strDescricaoLog += " \n" + "Bairro: " + rowBoletoPrincipal.bairro_sacado + " -> " + bairroCorrigido;
					if (!rowBoletoPrincipal.cep_sacado.Equals(cepCorrigido)) strDescricaoLog += " \n" + "CEP: " + rowBoletoPrincipal.cep_sacado + " -> " + cepCorrigido;
					if (!rowBoletoPrincipal.cidade_sacado.Equals(cidadeCorrigido)) strDescricaoLog += " \n" + "Cidade: " + rowBoletoPrincipal.cidade_sacado + " -> " + cidadeCorrigido;
					if (!rowBoletoPrincipal.uf_sacado.Equals(ufCorrigido)) strDescricaoLog += " \n" + "UF: " + rowBoletoPrincipal.uf_sacado + " -> " + ufCorrigido;
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_OCORRENCIAS_TRATA_CEP_IRREGULAR;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO_ITEM;
					finLog.id_registro_origem = intIdBoletoItem;
					finLog.id_boleto_cedente = intIdBoletoCedente;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					blnSucesso = true;
				}
				finally
				{
					if (blnSucesso)
					{
						BD.commitTransacao();
					}
					else
					{
						BD.rollbackTransacao();
					}
				}
				#endregion

				#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
				_atualizacaoAutomaticaPesquisaEmAndamento = true;
				try
				{
					executaPesquisa();
				}
				finally
				{
					_atualizacaoAutomaticaPesquisaEmAndamento = false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				avisoErro(ex.ToString());
			}
		}
		#endregion

		#region [ trataMarcarComoJaTratada ]
		private void trataMarcarComoJaTratada()
		{
			#region [ Declarações ]
			int intIdBoletoOcorrencia = 0;
			int intIdBoleto = 0;
			int intIdBoletoItem = 0;
			int intIdBoletoCedente = 0;
			String strIdentificacaoOcorrencia = "";
			String strNomeCliente = "";
			String strCnpjCpf = "";
			String strRegistroArqRetorno = "";
			String strComentarioOcorrenciaTratada;
			String strMsgErro = "";
			String strDescricaoLog;
			String strMsgErroLog = "";
			DialogResult drResultado;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			bool blnSucesso = false;
			FinLog finLog = new FinLog();
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

				#region [ Consistência ]
				if (gridDados.SelectedRows.Count == 0)
				{
					avisoErro("Nenhum registro foi selecionado!!");
					return;
				}

				if (gridDados.SelectedRows.Count > 1)
				{
					avisoErro("Não é permitida a seleção de múltiplos registros!!");
					return;
				}
				#endregion

				#region [ Obtém Id do registro ]
				foreach (DataGridViewRow item in gridDados.SelectedRows)
				{
					strIdentificacaoOcorrencia = item.Cells["identificacao_ocorrencia"].Value.ToString();
					intIdBoletoOcorrencia = (int)Global.converteInteiro(item.Cells["id_boleto_ocorrencia"].Value.ToString());
					intIdBoleto = (int)Global.converteInteiro(item.Cells["id_boleto"].Value.ToString());
					intIdBoletoItem = (int)Global.converteInteiro(item.Cells["id_boleto_item"].Value.ToString());
					intIdBoletoCedente = (int)Global.converteInteiro(item.Cells["id_boleto_cedente"].Value.ToString());
					strRegistroArqRetorno = item.Cells["registro_arq_retorno"].Value.ToString();
				}
				#endregion

				#region [ Id do registro é válida? ]
				if (intIdBoletoOcorrencia == 0)
				{
					avisoErro("Falha ao obter o nº identificação do registro da ocorrência!!");
					return;
				}
				#endregion

				#region [ Se há id do boleto associado, obtém dados do cliente ]
				if (intIdBoleto > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoleto(intIdBoleto);
					if (rowBoletoPrincipal == null)
					{
						avisoErro("Falha ao recuperar os dados do boleto!!");
						return;
					}
					strNomeCliente = rowBoletoPrincipal.nome_sacado;
					strCnpjCpf = rowBoletoPrincipal.num_inscricao_sacado;
				}
				#endregion

				#region [ Exibe form p/ exibir detalhes e editar comentário sobre o tratamento dado ]
				fBoletoTrataOcorrenciaValaComum = new FBoletoTrataOcorrenciaValaComum(intIdBoletoCedente, strNomeCliente, strCnpjCpf, strRegistroArqRetorno);
				fBoletoTrataOcorrenciaValaComum.StartPosition = FormStartPosition.Manual;
				fBoletoTrataOcorrenciaValaComum.Left = this.Left + (this.Width - fBoletoTrataOcorrenciaValaComum.Width) / 2;
				fBoletoTrataOcorrenciaValaComum.Top = this.Top + (this.Height - fBoletoTrataOcorrenciaValaComum.Height) / 2;
				drResultado = fBoletoTrataOcorrenciaValaComum.ShowDialog();
				if (drResultado != DialogResult.OK) return;
				#endregion

				#region [ Marca a ocorrência como já tratada ]
				strComentarioOcorrenciaTratada = fBoletoTrataOcorrenciaValaComum.comentarioOcorrenciaTratada;

				BD.iniciaTransacao();
				try
				{
					if (!BoletoDAO.marcaBoletoOcorrenciaComoJaTratada(
											Global.Usuario.usuario,
											intIdBoletoOcorrencia,
											strComentarioOcorrenciaTratada,
											ref strMsgErro))
					{
						throw new Exception("Falha ao marcar a ocorrência como já tratada!!\n\n" + strMsgErro);
					}

					#region [ Grava o log no BD ]
					strDescricaoLog = "Tratamento em 'Boleto - Ocorrências' como vala comum: ocorrência " + strIdentificacaoOcorrencia + " (" + Global.decodificaIdentificacaoOcorrencia(strIdentificacaoOcorrencia) + ")" +
									  " \n" + "t_FIN_BOLETO.id=" + intIdBoleto.ToString() +
									  " \n" + "t_FIN_BOLETO_ITEM.id=" + intIdBoletoItem.ToString() +
									  " \n" + "t_FIN_BOLETO_OCORRENCIA.id=" + intIdBoletoOcorrencia.ToString() +
									  " \n" + "Comentários: " + strComentarioOcorrenciaTratada;
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_OCORRENCIAS_TRATA_VALA_COMUM;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					if (intIdBoletoItem == 0)
					{
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO_OCORRENCIA;
						finLog.id_registro_origem = intIdBoletoOcorrencia;
					}
					else
					{
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO_ITEM;
						finLog.id_registro_origem = intIdBoletoItem;
					}
					finLog.id_boleto_cedente = intIdBoletoCedente;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					blnSucesso = true;
				}
				finally
				{
					if (blnSucesso)
					{
						BD.commitTransacao();
					}
					else
					{
						BD.rollbackTransacao();
					}
				}
				#endregion

				#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
				_atualizacaoAutomaticaPesquisaEmAndamento = true;
				try
				{
					executaPesquisa();
				}
				finally
				{
					_atualizacaoAutomaticaPesquisaEmAndamento = false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				avisoErro(ex.ToString());
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

		#region [ Form FBoletoOcorrencias ]

		#region [ FBoletoOcorrencias_Load ]
		private void FBoletoOcorrencias_Load(object sender, EventArgs e)
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

		#region [ FBoletoOcorrencias_Shown ]
		private void FBoletoOcorrencias_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Permissão de acesso ao módulo ]
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
					{
						btnOcorrenciaTratar.Enabled = false;
						menuOcorrenciaTratar.Enabled = false;
					}
					#endregion

					#region [ Posiciona foco ]
					txtDataInicial.Focus();
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

		#region [ FBoletoOcorrencias_FormClosing ]
		private void FBoletoOcorrencias_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FBoletoOcorrencias_KeyDown ]
		private void FBoletoOcorrencias_KeyDown(object sender, KeyEventArgs e)
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

		#region [ txtDataInicial ]

		#region [ txtDataInicial_Enter ]
		private void txtDataInicial_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataInicial_Leave ]
		private void txtDataInicial_Leave(object sender, EventArgs e)
		{
			if (txtDataInicial.Text.Length == 0) return;
			txtDataInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataInicial.Text);
			if (!Global.isDataOk(txtDataInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtDataInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataInicial_KeyDown ]
		private void txtDataInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDataFinal);
		}
		#endregion

		#region [ txtDataInicial_KeyPress ]
		private void txtDataInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDataFinal ]

		#region [ txtDataFinal_Enter ]
		private void txtDataFinal_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataFinal_Leave ]
		private void txtDataFinal_Leave(object sender, EventArgs e)
		{
			if (txtDataFinal.Text.Length == 0) return;
			txtDataFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataFinal.Text);
			if (!Global.isDataOk(txtDataFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtDataFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataFinal_KeyDown ]
		private void txtDataFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtValor);
		}
		#endregion

		#region [ txtDataFinal_KeyPress ]
		private void txtDataFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, cbBoletoCedente);
		}
		#endregion

		#region [ txtValor_KeyPress ]
		private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtNumDocumento ]

		#region [ txtNumDocumento_Enter ]
		private void txtNumDocumento_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumDocumento_Leave ]
		private void txtNumDocumento_Leave(object sender, EventArgs e)
		{
			txtNumDocumento.Text = txtNumDocumento.Text.Trim();
		}
		#endregion

		#region [ txtNumDocumento_KeyDown ]
		private void txtNumDocumento_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, cbOcorrencia);
		}
		#endregion

		#region [ txtNumDocumento_KeyPress ]
		private void txtNumDocumento_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtNossoNumero ]

		#region [ txtNossoNumero_Enter ]
		private void txtNossoNumero_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNossoNumero_Leave ]
		private void txtNossoNumero_Leave(object sender, EventArgs e)
		{
			txtNossoNumero.Text = txtNossoNumero.Text.Trim();
		}
		#endregion

		#region [ txtNossoNumero_KeyDown ]
		private void txtNossoNumero_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnPesquisar);
		}
		#endregion

		#region [ txtNossoNumero_KeyPress ]
		private void txtNossoNumero_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ cbBoletoCedente ]

		#region [ cbBoletoCedente_KeyDown ]
		private void cbBoletoCedente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtNumDocumento);
		}
		#endregion

		#endregion

		#region [ cbOcorrencia ]

		#region [ cbOcorrencia_KeyDown ]
		private void cbOcorrencia_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtNossoNumero);
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
				trataOcorrenciaSelecionada();
				return;
			}
		}
		#endregion

		#region [ gridDados_DoubleClick ]
		private void gridDados_DoubleClick(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
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

		#region [ menuOcorrenciaPesquisar_Click ]
		private void menuOcorrenciaPesquisar_Click(object sender, EventArgs e)
		{
			executaPesquisa();
		}
		#endregion

		#endregion

		#region [ Tratar Ocorrência ]

		#region [ btnOcorrenciaTratar_Click ]
		private void btnOcorrenciaTratar_Click(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion

		#region [ menuOcorrenciaTratar_Click ]
		private void menuOcorrenciaTratar_Click(object sender, EventArgs e)
		{
			trataOcorrenciaSelecionada();
		}
		#endregion

		#endregion

		#region [ Marcar como já tratada ]

		#region [ btnMarcarComoJaTratada_Click ]
		private void btnMarcarComoJaTratada_Click(object sender, EventArgs e)
		{
			trataMarcarComoJaTratada();
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

		#region [ menuOcorrenciaLimpar_Click ]
		private void menuOcorrenciaLimpar_Click(object sender, EventArgs e)
		{
			limpaCampos();
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
				prnDocConsulta.DocumentName = "Boleto: Ocorrências";
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
				wxDataOcorrencia = 16f;
				wxCliente = 40f;
				wxNumeroDocumento = 19f;
				wxDtVencto = 16f;
				wxVlTitulo = 23f;
				wxLoja = 10f;
				wxPedido = 15f;
				wxOcorrencia = 55f;
				wxOcorrenciaObs = larguraUtil
								  - wxDataOcorrencia            // A 1ª coluna não tem espaçamento
								  - ESPACAMENTO_COLUNAS     // Espaçamento da própria coluna "Obs"
								  - ESPACAMENTO_COLUNAS - wxCliente
								  - ESPACAMENTO_COLUNAS - wxNumeroDocumento
								  - ESPACAMENTO_COLUNAS - wxDtVencto
								  - ESPACAMENTO_COLUNAS - wxVlTitulo
								  - ESPACAMENTO_COLUNAS - wxLoja
								  - ESPACAMENTO_COLUNAS - wxPedido
								  - ESPACAMENTO_COLUNAS - wxOcorrencia;

				ixDataOcorrencia = cxInicio;
				ixCliente = ixDataOcorrencia + wxDataOcorrencia + ESPACAMENTO_COLUNAS;
				ixNumeroDocumento = ixCliente + wxCliente + ESPACAMENTO_COLUNAS;
				ixDtVencto = ixNumeroDocumento + wxNumeroDocumento + ESPACAMENTO_COLUNAS;
				ixVlTitulo = ixDtVencto + wxDtVencto + ESPACAMENTO_COLUNAS;
				ixLoja = ixVlTitulo + wxVlTitulo + ESPACAMENTO_COLUNAS;
				ixPedido = ixLoja + wxLoja + ESPACAMENTO_COLUNAS;
				ixOcorrencia = ixPedido + wxPedido + ESPACAMENTO_COLUNAS;
				ixOcorrenciaObs = ixOcorrencia + wxOcorrencia + ESPACAMENTO_COLUNAS;
				#endregion
			}

			cy = cyInicio;

			#region [ Título ]
			strTexto = "BOLETO: OCORRÊNCIAS";
			fonteAtual = fonteTitulo;
			cx = cxInicio + (larguraUtil - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			cy += 1f;
			#endregion

			#region [ Informações no cabeçalho ]

			#region [ Data da emissão / Cedente ]

			#region [ Data da emissão ]
			strTexto = "Emissão: " + _strImpressaoDataEmissao;
			fonteAtual = fonteListagem;
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
			cx = cxInicio + (larguraUtil * 0.33f);
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			#endregion

			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Filtros: Data / Valor ]
			strTexto = "Data: " +
					   ((txtDataInicial.Text.Trim().Length > 0) ? txtDataInicial.Text : "N.I.") +
					   " a " +
					   ((txtDataFinal.Text.Trim().Length > 0) ? txtDataFinal.Text : "N.I.");
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil * 0.33f);
			strTexto = "Valor (" + Global.Cte.Etc.SIMBOLO_MONETARIO + "): " +
					   ((txtValor.Text.Trim().Length > 0) ? txtValor.Text : "N.I.");
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#region [ Nº Documento / Nosso Número / Somente com divergência de valor]
			strTexto = "Nº Documento: " +
					   ((txtNumDocumento.Text.Trim().Length > 0) ? txtNumDocumento.Text : "N.I.");
			cx = cxInicio;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil * 0.33f);
			strTexto = "Nosso Número: " +
					   ((txtNossoNumero.Text.Trim().Length > 0) ? txtNossoNumero.Text : "N.I.");
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			cx = cxInicio + (larguraUtil * 0.66f);
			strTexto = "Somente com divergência de valor: " + (ckb_somente_divergencia_valor.Checked ? "Sim" : "Não");
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

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
			cy += fonteAtual.GetHeight(e.Graphics);
			#endregion

			#endregion

			cy += .5f;
			e.Graphics.DrawLine(penTracoTitulo, cxInicio, cy, cxFim, cy);
			cy += .5f;

			#region [ Títulos da listagem ]
			cy += .5f;
			fonteAtual = fonteListagem;
			strTexto = "DATA";
			cx = ixDataOcorrencia;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "CLIENTE";
			cx = ixCliente;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "Nº DOC";
			cx = ixNumeroDocumento;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "DT VENCTO";
			cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "VL TÍTULO";
			cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
			
			strTexto = "LOJA";
			cx = ixLoja;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "PEDIDO";
			cx = ixPedido;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "OCORRÊNCIA";
			cx = ixOcorrencia;
			e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);

			strTexto = "OBS";
			cx = ixOcorrenciaObs;
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
				#region [ Há espaço suficiente p/ os textos multi-linhas? ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["cliente"].Value.ToString();
				if ((cy + e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxCliente).Height) > (cyRodapeNumPagina - 5)) break;

				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["ocorrencia"].Value.ToString();
				if ((cy + e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxOcorrencia).Height) > (cyRodapeNumPagina - 5)) break;

				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["obs"].Value.ToString();
				if ((cy + e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxOcorrenciaObs).Height) > (cyRodapeNumPagina - 5)) break;
				#endregion

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

				#region [ Data da ocorrência ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["dt_cadastro"].Value.ToString();
				cx = ixDataOcorrencia;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Cliente ]
				cx = ixCliente;
				r = new RectangleF(ixCliente, cy, wxCliente, 20);
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["cliente"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxCliente).Height);
				#endregion

				#region [ Nº Documento ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["numero_documento"].Value.ToString();
				cx = ixNumeroDocumento;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Data do vencimento ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["dt_vencto"].Value.ToString();
				cx = ixDtVencto + (wxDtVencto - e.Graphics.MeasureString(strTexto, fonteAtual).Width) / 2;
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, cx, cy);
				#endregion

				#region [ Valor do título ]
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["vl_titulo"].Value.ToString();
				vlTitulo = Global.converteNumeroDecimal(strTexto);
				cx = ixVlTitulo + wxVlTitulo - e.Graphics.MeasureString(strTexto, fonteAtual).Width;
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

				#region [ Ocorrência ]
				cx = ixOcorrencia;
				r = new RectangleF(ixOcorrencia, cy, wxOcorrencia, 20);
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["ocorrencia"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxOcorrencia).Height);
				#endregion

				#region [ Obs ]
				cx = ixOcorrenciaObs;
				r = new RectangleF(ixOcorrenciaObs, cy, wxOcorrenciaObs, 30);
				strTexto = gridDados.Rows[_intImpressaoIdxLinhaGrid].Cells["obs"].Value.ToString();
				e.Graphics.DrawString(strTexto, fonteAtual, brushPadrao, r);
				hMax = Math.Max(hMax, e.Graphics.MeasureString(strTexto, fonteAtual, (int)wxOcorrenciaObs).Height);
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
					cx = ixCliente;
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
}
