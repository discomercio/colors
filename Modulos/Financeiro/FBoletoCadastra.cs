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
	public partial class FBoletoCadastra : Financeiro.FModelo
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

		FBoletoCadastraDetalhe fBoletoCadastraDetalhe;

		int _flagPedidoUsarMemorizacaoCompletaEnderecos;
		#endregion

		#region [ Menu ]
		ToolStripMenuItem menuBoleto;
		ToolStripMenuItem menuBoletoPesquisar;
		ToolStripMenuItem menuBoletoDetalhe;
		ToolStripMenuItem menuBoletoLimpar;
		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoCadastra()
		{
			InitializeComponent();

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
			txtDataEmissaoInicial.Text = "";
			txtDataEmissaoFinal.Text = "";
			txtNumPedido.Text = "";
			txtNumNF.Text = "";
			txtNumParcelas.Text = "";
			txtValor.Text = "";
			txtNomeCliente.Text = "";
			txtCnpjCpf.Text = "";
			txtNumLoja.Text = "";
			cbBoletoCedente.SelectedIndex = -1;
			lblTotalizacaoRegistros.Text = "";
			lblTotalizacaoValor.Text = "";
			gridDados.DataSource = null;
			txtDataEmissaoInicial.Focus();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Declarações ]
			DateTime dtEmissaoInicial = DateTime.MinValue;
			DateTime dtEmissaoFinal = DateTime.MinValue;
			#endregion

			#region [ Período da Data de Emissão da NF ]
			if (txtDataEmissaoInicial.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataEmissaoInicial.Text))
				{
					avisoErro("Data inválida!!");
					txtDataEmissaoInicial.Focus();
					return false;
				}
				else dtEmissaoInicial = Global.converteDdMmYyyyParaDateTime(txtDataEmissaoInicial.Text);
			}
			
			if (txtDataEmissaoFinal.Text.Trim().Length > 0)
			{
				if (!Global.isDataOk(txtDataEmissaoFinal.Text))
				{
					avisoErro("Data inválida!!");
					txtDataEmissaoFinal.Focus();
					return false;
				}
				else dtEmissaoFinal = Global.converteDdMmYyyyParaDateTime(txtDataEmissaoFinal.Text);
			}

			if ((dtEmissaoInicial > DateTime.MinValue) && (dtEmissaoFinal > DateTime.MinValue))
			{
				if (dtEmissaoInicial > dtEmissaoFinal)
				{
					avisoErro("A data final do período é anterior à data inicial!!");
					txtDataEmissaoFinal.Focus();
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

			// Ok!
			return true;
		}
		#endregion

		#region [ montaClausulaWhereInterno ]
		private String montaClausulaWhereInterno()
		{
			#region [ Declarações ]
			StringBuilder sbWhere = new StringBuilder("");
			StringBuilder sbWhereNFeEmitente = new StringBuilder("");
			String strAux;
			List<int> listaNFeEmitente;
			#endregion

			#region [ Restrição fixa ]
			strAux = " (tFNPP.status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")";
			if (sbWhere.Length > 0) sbWhere.Append(" AND");
			sbWhere.Append(strAux);
			#endregion

			#region [ Data de emissão ]
			if ((txtDataEmissaoInicial.Text.Length > 0) && (txtDataEmissaoFinal.Text.Length > 0))
			{
				// A data inicial é igual à data final?
				if (txtDataEmissaoInicial.Text.Equals(txtDataEmissaoFinal.Text))
				{
					strAux = " (tFNPP.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataEmissaoInicial.Text) + ")";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " ((tFNPP.dt_cadastro >= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataEmissaoInicial.Text) + ") AND (tFNPP.dt_cadastro <= " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataEmissaoFinal.Text) + "))";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			else if ((txtDataEmissaoInicial.Text.Length > 0) || (txtDataEmissaoFinal.Text.Length > 0))
			{
				if (txtDataEmissaoInicial.Text.Length > 0)
				{
					strAux = " (tFNPP.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataEmissaoInicial.Text) + ")";
				}
				else if (txtDataEmissaoFinal.Text.Length > 0)
				{
					strAux = " (tFNPP.dt_cadastro = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(txtDataEmissaoFinal.Text) + ")";
				}
				else strAux = "";

				if (strAux.Length > 0)
				{
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			#region [ Nº pedido ]
			if (txtNumPedido.Text.Length > 0)
			{
				#region [ Exibe de um pedido específico ]
				strAux = " (tFNPP.id IN" +
								"(" +
									"SELECT" +
										" tFNPPI.id_nf_parcela_pagto" +
									" FROM t_FIN_NF_PARCELA_PAGTO_ITEM tFNPPI" +
										" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tFNPPIR" +
											" ON tFNPPI.id=tFNPPIR.id_nf_parcela_pagto_item" +
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
			else
			{
				#region [ Se não foi informado o nº de um pedido específico, exibe apenas dos pedidos entregues ]
				strAux = " (tFNPP.id IN" +
								"(" +
									"SELECT" +
										" tFNPPI.id_nf_parcela_pagto" +
									" FROM t_FIN_NF_PARCELA_PAGTO tFNPP" +
										" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM tFNPPI" +
											" ON tFNPP.id=tFNPPI.id_nf_parcela_pagto" +
										" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tFNPPIR" +
											" ON tFNPPI.id=tFNPPIR.id_nf_parcela_pagto_item" +
										" INNER JOIN t_PEDIDO tP" +
											" ON tFNPPIR.pedido=tP.pedido" +
									" WHERE" +
										" (tFNPP.status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")" +
										" AND (tP.st_entrega = '" + Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE + "')" +
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

			#region [ Nº NF ]
			if (txtNumNF.Text.Length > 0)
			{
				strAux = " (tFNPP.numero_NF = " + txtNumNF.Text + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Nº Parcelas ]
			if (txtNumParcelas.Text.Length > 0)
			{
				strAux = " (tFNPP.qtde_parcelas_boleto = " + txtNumParcelas.Text + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Nome do cliente ]
			if (txtNomeCliente.Text.Trim().Length > 0)
			{
				strAux = " (tC.nome LIKE '" + BD.CARACTER_CURINGA_TODOS + txtNomeCliente.Text + BD.CARACTER_CURINGA_TODOS + "'" + Global.Cte.Etc.SQL_COLLATE_CASE_ACCENT + ")";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ CNPJ/CPF ]
			if (Global.digitos(txtCnpjCpf.Text).Length > 0)
			{
				strAux = " (tC.cnpj_cpf = '" + Global.digitos(txtCnpjCpf.Text) + "')";
				if (sbWhere.Length > 0) sbWhere.Append(" AND");
				sbWhere.Append(strAux);
			}
			#endregion

			#region [ Nº Loja ]
			if (txtNumLoja.Text.Length > 0)
			{
				strAux = " (tFNPP.id IN" +
								"(" +
									"SELECT" +
										" tFNPPI.id_nf_parcela_pagto" +
									" FROM t_FIN_NF_PARCELA_PAGTO tFNPP" +
										" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM tFNPPI" +
											" ON tFNPP.id=tFNPPI.id_nf_parcela_pagto" +
										" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tFNPPIR" +
											" ON tFNPPI.id=tFNPPIR.id_nf_parcela_pagto_item" +
										" INNER JOIN t_PEDIDO tP" +
											" ON tFNPPIR.pedido=tP.pedido" +
									" WHERE" +
										" (tFNPP.status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")" +
										" AND (CONVERT(smallint, tP.loja) = " + txtNumLoja.Text.Trim() + ")" +
								")" +
							")";
				if (strAux.Length > 0)
				{
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
					#region [ Não é o cedente padrão! ]
					listaNFeEmitente = BoletoCedenteDAO.getNFeEmitentesBoletoCedente((int)Global.converteInteiro(cbBoletoCedente.SelectedValue.ToString()));
					foreach (int id_nfe_emitente in listaNFeEmitente)
					{
						if (sbWhereNFeEmitente.Length > 0) sbWhereNFeEmitente.Append(", ");
						sbWhereNFeEmitente.Append(id_nfe_emitente.ToString());
					}
					if (sbWhereNFeEmitente.Length > 0)
					{
						strAux = " (tFNPP.id IN " +
										"(" +
											"SELECT" +
												" tFNPPI.id_nf_parcela_pagto" +
											" FROM t_FIN_NF_PARCELA_PAGTO tFNPP" +
												" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM tFNPPI" +
													" ON tFNPP.id=tFNPPI.id_nf_parcela_pagto" +
												" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tFNPPIR" +
													" ON tFNPPI.id=tFNPPIR.id_nf_parcela_pagto_item" +
												" INNER JOIN t_PEDIDO tP" +
													" ON tFNPPIR.pedido=tP.pedido" +
											" WHERE" +
												" (tFNPP.status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")" +
												" AND (tP.id_nfe_emitente IN (" + sbWhereNFeEmitente.ToString() + "))" +
										")" +
									")";
						if (strAux.Length > 0)
						{
							if (sbWhere.Length > 0) sbWhere.Append(" AND");
							sbWhere.Append(strAux);
						}
					}
					#endregion
				}
			}
			#endregion

			return sbWhere.ToString();
		}
		#endregion

		#region [ montaClausulaWhereExterno ]
		private String montaClausulaWhereExterno()
		{
			StringBuilder sbWhere = new StringBuilder("");
			String strAux;

			#region[ Valor ]
			if (txtValor.Text.Trim().Length > 0)
			{
				if (Global.converteNumeroDecimal(txtValor.Text) > 0)
				{
					strAux = Global.sqlFormataDecimal(Global.converteNumeroDecimal(txtValor.Text));
					strAux = " (vl_total = " + strAux + ")";
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
			String strWhereInterno;
			String strWhereExterno;
			String strSql;
			String strPedidoEnderecoNome;

			#region [ Monta cláusula Where ]
			strWhereInterno = montaClausulaWhereInterno();
			if (strWhereInterno.Length > 0) strWhereInterno = " WHERE " + strWhereInterno;

			strWhereExterno = montaClausulaWhereExterno();
			if (strWhereExterno.Length > 0) strWhereExterno = " WHERE " + strWhereExterno;
			#endregion

			if (_flagPedidoUsarMemorizacaoCompletaEnderecos == 0)
			{
				strPedidoEnderecoNome = "NULL";
			}
			else
			{
				strPedidoEnderecoNome = "SELECT TOP 1 endereco_nome FROM t_PEDIDO tP INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tRateio ON (tP.pedido = tRateio.pedido) WHERE (tRateio.id_nf_parcela_pagto = tFNPP.id) AND (tP.st_memorizacao_completa_enderecos <> 0) ORDER BY tP.data_hora";
			}

			#region [ Monta Select ]
			strSql = "SELECT" +
						" id," +
						" dt_cadastro," +
						" numero_NF," +
						" qtde_parcelas_boleto," +
						" status," +
						" id_cliente," +
						" Coalesce(pedido_endereco_nome, nome) AS nome," +
						" cnpj_cpf," +
						" pedido," +
						" '' AS st_pagto," +
						" '' AS st_pagto_descricao," +
						" vl_total" +
					" FROM " +
						"(" +
							"SELECT " +
								" tFNPP.id," +
								" tFNPP.dt_cadastro," +
								" tFNPP.numero_NF," +
								" tFNPP.qtde_parcelas_boleto," +
								" tFNPP.status," +
								" tFNPP.id_cliente," +
								" tC.nome," +
								" (" + strPedidoEnderecoNome + ") AS pedido_endereco_nome," +
								" tC.cnpj_cpf, " +
								BD.strSchema + ".ConcatenaPedidosTabelaFinNfParcelaPagtoItemRateio(tFNPP.id, ', ') AS pedido," +
								" (" +
									"SELECT" +
										" Coalesce(Sum(valor),0) AS vl_total" +
									" FROM t_FIN_NF_PARCELA_PAGTO_ITEM" +
									" WHERE" +
										" (id_nf_parcela_pagto=tFNPP.id)" +
										" AND (forma_pagto = " + Global.Cte.FIN.FormaPagto.ID_FORMA_PAGTO_BOLETO.ToString() + ")" +
								") AS vl_total" +
							" FROM t_FIN_NF_PARCELA_PAGTO tFNPP" +
								" INNER JOIN t_CLIENTE tC" +
									" ON tFNPP.id_cliente=tC.id" +
							strWhereInterno +
						") t" +
					strWhereExterno +
					" ORDER BY" +
						" dt_cadastro," +
						" id";
			#endregion

			return strSql;
		}
		#endregion

		#region [ executaPesquisa ]
		private void executaPesquisa()
		{
			#region [ Declarações ]
			Decimal decTotalizacaoValor = 0;
			int intQtdeRegistros = 0;
			String strSql;
			string st_pagto;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DsDataSource.DtbNfParcelaPagtoGridDataTable dtbConsulta = new DsDataSource.DtbNfParcelaPagtoGridDataTable();
			DsDataSource.DtbNfParcelaPagtoGridRow rowConsulta;
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

				#region [ Prepara alguns campos que necessitam de formatação ]
				for (int i = 0; i < dtbConsulta.Count; i++)
				{
					rowConsulta = (DsDataSource.DtbNfParcelaPagtoGridRow)dtbConsulta.Rows[i];
					rowConsulta.valor_formatado = Global.formataMoeda(rowConsulta.vl_total);
					rowConsulta.cnpj_cpf_formatado = Global.formataCnpjCpf(rowConsulta.cnpj_cpf);
					rowConsulta.dt_cadastro_formatada = Global.formataDataDdMmYyyyComSeparador(rowConsulta.dt_cadastro);

					try
					{
						st_pagto = PedidoDAO.getPedidoStPagto(rowConsulta["pedido"].ToString());
					}
					catch (Exception)
					{
						st_pagto = "";
					}

					if ((st_pagto ?? "").Trim().Length > 0)
					{
						rowConsulta["st_pagto"] = st_pagto;
						rowConsulta["st_pagto_descricao"] = Global.stPagtoPedidoDescricao(st_pagto);
					}

					decTotalizacaoValor += rowConsulta.vl_total;
					intQtdeRegistros++;
				}
				#endregion

				#region [ Exibição dos dados no grid ]
				try
				{
					gridDados.SuspendLayout();

					#region [ Carrega os dados no Grid ]
					gridDados.DataSource = dtbConsulta;
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

				gridDados.Focus();

				// Feedback da conclusão da pesquisa
				SystemSounds.Exclamation.Play();
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

		#region [ consultaDetalheRegistroSelecionado ]
		private void consultaDetalheRegistroSelecionado()
		{
			#region [ Declarações ]
			int intRegistroSelecionado = 0;
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

			foreach (DataGridViewRow item in gridDados.SelectedRows)
			{
				intRegistroSelecionado = (int)Global.converteInteiro(item.Cells["id"].Value.ToString());
			}

			fBoletoCadastraDetalhe = new FBoletoCadastraDetalhe(this);
			fBoletoCadastraDetalhe.evtRegistroGravado += new RegistroGravadoEventHandler(TrataEventoDadosBoletoCadastrado);
			fBoletoCadastraDetalhe.evtRegistroAnulado += new RegistroAnuladoEventHandler(TrataEventoDadosBoletoCancelado);
			fBoletoCadastraDetalhe.idBoletoPreCadastradoSelecionado = intRegistroSelecionado;
			fBoletoCadastraDetalhe.Location = this.Location;
			fBoletoCadastraDetalhe.Show();
			if (!fBoletoCadastraDetalhe.ocorreuExceptionNaInicializacao) this.Visible = false;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FBoletoCadastra ]

		#region [ FBoletoCadastra_Load ]
		private void FBoletoCadastra_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCampos();

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

		#region [ FBoletoCadastra_Shown ]
		private void FBoletoCadastra_Shown(object sender, EventArgs e)
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

					#region [ Combo Cedente ]
					cbBoletoCedente.ValueMember = "id";
					cbBoletoCedente.DisplayMember = "descricao_formatada";
					cbBoletoCedente.DataSource = ComboDAO.criaDtbBoletoCedenteCombo(ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS);
					cbBoletoCedente.SelectedIndex = -1;
					// Se houver apenas 1 opção, então seleciona
					if ((cbBoletoCedente.Items.Count == 1) && (cbBoletoCedente.SelectedIndex == -1)) cbBoletoCedente.SelectedIndex = 0;
					#endregion

					#region [ Parâmetros do BD ]
					_flagPedidoUsarMemorizacaoCompletaEnderecos = ComumDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS, 0);
					#endregion

					#region [ Posiciona foco ]
					txtDataEmissaoInicial.Focus();
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

		#region [ FBoletoCadastra_FormClosing ]
		private void FBoletoCadastra_FormClosing(object sender, FormClosingEventArgs e)
		{
			// Campo nome do cliente exibe uma lista de sugestões
			if (ActiveControl == txtNomeCliente)
			{
				btnDummy.Focus();
				txtNomeCliente.Focus();
				Global.textBoxPosicionaCursorNoFinal(txtNomeCliente);
				e.Cancel = true;
				return;
			}

			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#region [ FBoletoCadastra_KeyDown ]
		private void FBoletoCadastra_KeyDown(object sender, KeyEventArgs e)
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

		#region [ txtDataEmissaoInicial ]

		#region [ txtDataEmissaoInicial_Enter ]
		private void txtDataEmissaoInicial_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataEmissaoInicial_Leave ]
		private void txtDataEmissaoInicial_Leave(object sender, EventArgs e)
		{
			if (txtDataEmissaoInicial.Text.Length == 0) return;
			txtDataEmissaoInicial.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataEmissaoInicial.Text);
			if (!Global.isDataOk(txtDataEmissaoInicial.Text))
			{
				avisoErro("Data inválida!!");
				txtDataEmissaoInicial.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataEmissaoInicial_KeyDown ]
		private void txtDataEmissaoInicial_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDataEmissaoFinal);
		}
		#endregion

		#region [ txtDataEmissaoInicial_KeyPress ]
		private void txtDataEmissaoInicial_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDataEmissaoFinal ]

		#region [ txtDataEmissaoFinal_Enter ]
		private void txtDataEmissaoFinal_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDataEmissaoFinal_Leave ]
		private void txtDataEmissaoFinal_Leave(object sender, EventArgs e)
		{
			if (txtDataEmissaoFinal.Text.Length == 0) return;
			txtDataEmissaoFinal.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtDataEmissaoFinal.Text);
			if (!Global.isDataOk(txtDataEmissaoFinal.Text))
			{
				avisoErro("Data inválida!!");
				txtDataEmissaoFinal.Focus();
				return;
			}
		}
		#endregion

		#region [ txtDataEmissaoFinal_KeyDown ]
		private void txtDataEmissaoFinal_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtNumPedido);
		}
		#endregion

		#region [ txtDataEmissaoFinal_KeyPress ]
		private void txtDataEmissaoFinal_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, txtNumParcelas);
		}
		#endregion

		#region [ txtNumPedido_KeyPress ]
		private void txtNumPedido_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroPedido(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtNumParcelas ]

		#region [ txtNumParcelas_Enter ]
		private void txtNumParcelas_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumParcelas_Leave ]
		private void txtNumParcelas_Leave(object sender, EventArgs e)
		{
			txtNumParcelas.Text = txtNumParcelas.Text.Trim();
		}
		#endregion

		#region [ txtNumParcelas_KeyDown ]
		private void txtNumParcelas_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtNumNF);
		}
		#endregion

		#region [ txtNumParcelas_KeyPress ]
		private void txtNumParcelas_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ cbBoletoCedente ]

		#region [ cbBoletoCedente_KeyDown ]
		private void cbBoletoCedente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtNumLoja);
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
			Global.trataTextBoxKeyDown(sender, e, txtNumLoja);
		}
		#endregion

		#region [ txtNumNF_KeyPress ]
		private void txtNumNF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ Nº Loja ]

		#region [ txtNumLoja_Enter ]
		private void txtNumLoja_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumLoja_Leave ]
		private void txtNumLoja_Leave(object sender, EventArgs e)
		{
			txtNumLoja.Text = txtNumLoja.Text.Trim();
		}
		#endregion

		#region [ txtNumLoja_KeyDown ]
		private void txtNumLoja_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtValor);
		}
		#endregion

		#region [ txtNumLoja_KeyPress ]
		private void txtNumLoja_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
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

		#endregion

		#endregion

		#region [ Eventos acionados pelo painel FBoletoCadastraDetalhe ]

		#region [ TrataEventoDadosBoletoCadastrado ]
		public void TrataEventoDadosBoletoCadastrado()
		{
			#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
			executaPesquisa();
			#endregion
		}
		#endregion

		#region [ TrataEventoDadosBoletoCancelado ]
		public void TrataEventoDadosBoletoCancelado()
		{
			#region [ Refaz a pesquisa p/ atualizar os dados no grid ]
			executaPesquisa();
			#endregion
		}
		#endregion

		#endregion
	}
}
