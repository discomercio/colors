#region [ using ]
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Media;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
#endregion

namespace PrnDANFE
{
	#region [ FDANFEImprime ]
	public partial class FDANFEImprime : PrnDANFE.FModelo
	{
		#region [ enum ]
		private enum eFiltroPreenchimentoObrigatorio
		{
			OBRIGATORIO = 1,
			OPCIONAL = 2
		}
		#endregion

		#region [ Constantes ]
		const String GRID_PESQ_COL_CHECKBOX = "colGrdPesqCheckBox";
		const String GRID_PESQ_COL_PEDIDO = "colGrdPesqPedido";
		const String GRID_PESQ_COL_CIDADE = "colGrdPesqCidade";
		const String GRID_PESQ_COL_UF = "colGrdPesqUF";
		const String GRID_PESQ_COL_DATA_ENTREGA = "colGrdPesqDataEntrega";
		const String GRID_PESQ_COL_TRANSPORTADORA = "colGrdPesqTransportadora";
		const String GRID_PESQ_COL_NFE = "colGrdPesqNFE";
		const String GRID_PESQ_COL_SERIE = "colGrdPesqSerie";
		#endregion

		#region [ Atributos ]
		private string[] _ArquivosGravados;
		private bool _listagemMarcada = false;
		private bool _emProcessamento = false;
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

		private DataTable _dtbConsulta = new DataTable();
		#endregion

		#region [ Impressão ]
		const char CODIGO_SOH = (char)0x01;
		const char CODIGO_STX = (char)0x02;
		const char CODIGO_CR = (char)0x0D;
		#endregion

		#region [ Construtor ]
		public FDANFEImprime()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ haPdfSelecionado ]
		private bool haPdfSelecionado()
		{
			for (int i = 0; i < grdPesquisa.Rows.Count; i++)
			{
				if (grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CHECKBOX].Value == null) continue;
				if ((bool)grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CHECKBOX].Value) return true;
			}

			return false;
		}
		#endregion

		#region [ baixaPDFsSelecionados ]
		private bool baixaPDFsSelecionados()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "baixaPDFsSelecionados()";
			bool existePDFSelecionado = false;
			bool conseguiuBaixarPDF = false;
			string mensagemRetorno = "";
			string[] arqsSelecionados;
			int intIndexItem = 0;
			int i;
			#endregion

			try
			{
				arqsSelecionados = new string[grdPesquisa.Rows.Count];
				for (i = 0; i < grdPesquisa.Rows.Count; i++)
				{
					if (grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CHECKBOX].Value == null) continue;

					if ((bool)grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CHECKBOX].Value)
					{
						conseguiuBaixarPDF = PDFBaixar.executa_download_pdf_danfe_parametro_emitente(Global.Usuario.emit_id,
																grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_NFE].Value.ToString(),
																grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_SERIE].Value.ToString(),
																ref mensagemRetorno);
						if (conseguiuBaixarPDF)
						{
							existePDFSelecionado = true;
							arqsSelecionados.SetValue(grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_PEDIDO].Value.ToString(), intIndexItem);
							intIndexItem = intIndexItem + 1;
							if (mensagemRetorno != "")
							{
								aviso(mensagemRetorno);
								mensagemRetorno = "";
							}
						}
						else
						{
							grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CHECKBOX].Value = false;
							if (mensagemRetorno != "")
							{
								aviso(mensagemRetorno);
								mensagemRetorno = "";
							}
						}
					}
				}

				_ArquivosGravados = new string[intIndexItem];
				for (i = 0; i < intIndexItem; i++)
				{
					_ArquivosGravados[i] = arqsSelecionados[i];
				}

				return existePDFSelecionado;
			}
			catch (Exception ex)
			{
				aviso(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaPDFsBaixados ]
		private void atualizaPDFsBaixados()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "atualizaPDFsBaixados()";
			int intRetorno;
			bool blnSucesso = false;
			SqlCommand cmCommand;
			String strSql;
			#endregion

			try
			{
				if (_ArquivosGravados == null) return;

				try
				{
					BD.iniciaTransacao();

					cmCommand = BD.criaSqlCommand();
					for (int i = 0; i < _ArquivosGravados.Length; i++)
					{
						if (_ArquivosGravados[i] == null) continue;

						strSql = "UPDATE t_PEDIDO SET" +
									" danfe_impressa_status = " + Global.Cte.Etc.COD_DANFE_IMPRESSA_STATUS__OK + ", " +
									" danfe_impressa_data = getdate(), " +
									" danfe_impressa_data_hora = getdate(), " +
									" danfe_impressa_usuario = '" + Global.Usuario.usuario + "', " +
									" danfe_a_imprimir_status = " + Global.Cte.Etc.COD_DANFE_A_IMPRIMIR_STATUS__IMPRESSA + ", " +
									" danfe_a_imprimir_data_hora = getdate(), " +
									" danfe_a_imprimir_usuario = '" + Global.Usuario.usuario + "' " +
								" WHERE" +
									" (pedido = '" + _ArquivosGravados[i] + "')";
						cmCommand.CommandText = strSql;
						intRetorno = BD.executaNonQuery(ref cmCommand);
					}

					_ArquivosGravados = null;
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
			}
			catch (Exception ex)
			{
				aviso(NOME_DESTA_ROTINA + " - " + ex.ToString());
			}
		}
		#endregion

		#region [ carregaTransportadoras ]
		private void carregaTransportadoras()
		{
			#region [ Declarações ]
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbTransportadora = new DataTable();
			int i;
			#endregion

			cmCommand = BD.criaSqlCommand();
			daAdapter = BD.criaSqlDataAdapter();

			cmCommand.CommandText = "SELECT id, nome FROM t_TRANSPORTADORA ORDER BY id";
			daAdapter.SelectCommand = cmCommand;
			daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daAdapter.Fill(dtbTransportadora);

			cbTransportadora.Items.Clear();
			cbTransportadora.Items.Add("");
			for (i = 0; i < dtbTransportadora.Rows.Count; i++)
			{
				cbTransportadora.Items.Add(BD.readToString(dtbTransportadora.Rows[i]["id"]) + " - " + BD.readToString(dtbTransportadora.Rows[i]["nome"]));
			}
		}
		#endregion

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaPesquisa()";
			int flagPedidoUsarMemorizacaoCompletaEnderecos;
			String strSql;
			String strFrom;
			String strWhere = "";
			String strIdTransp;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			try
			{
				#region [ Limpa campos dos dados de resposta ]
				limpaCamposDados();
				#endregion

				#region [ Monta restrições da cláusula 'Where' ]
				strWhere = " WHERE (" +
					"(t_PEDIDO.st_entrega = '" + Global.Cte.StEntregaPedido.ST_ENTREGA_SEPARAR + "')" +
					" OR " +
					"(t_PEDIDO.st_entrega = '" + Global.Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR + "')" +
				")" +
			  " AND (" +
					"(t_PEDIDO.danfe_impressa_status = " + Global.Cte.Etc.COD_DANFE_IMPRESSA_STATUS__INICIAL + ")" +
					" OR " +
					"(t_PEDIDO.danfe_impressa_status = " + Global.Cte.Etc.COD_DANFE_IMPRESSA_STATUS__NAO_DEFINIDO + ")" +
				")" +
			  " AND (t_PEDIDO__BASE.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK + ")" +
			  " AND (t_PEDIDO.st_etg_imediata = " + Global.Cte.T_PEDIDO__ST_ETG_IMEDIATA.OPCAO_SIM + ") ";

				//CONDIÇÕES DO ROMANEIO
				strWhere = strWhere +
				" AND (t_PEDIDO.a_entregar_data_marcada IS NOT NULL) " +
				" AND (t_PEDIDO.transportadora_id IS NOT NULL) " +
				" AND ((" +
					"SELECT" +
						" TOP 1 NFe_numero_NF" +
					" FROM t_NFe_EMISSAO tNE" +
					" WHERE" +
						" (tNE.pedido=t_PEDIDO.pedido)" +
						" AND (tipo_NF = '1')" +
						" AND (st_anulado = 0)" +
						" AND (codigo_retorno_NFe_T1 = 1)" +
					" ORDER BY" +
						" id DESC" +
				") IS NOT NULL) ";

				//Obter apenas registros marcados para impressão
				strWhere = strWhere +
				" AND (t_PEDIDO.danfe_a_imprimir_status = " + Global.Cte.Etc.COD_DANFE_A_IMPRIMIR_STATUS__MARCADA + ") ";

				//Obter apenas pedidos referentes ao Emitente selecionado
				strWhere = strWhere +
				" AND (t_PEDIDO.id_nfe_emitente = " + Global.Usuario.emit_id + ") ";

				//Critérios definidos em tela
				if (dtpDataEntrega.Checked)
				{
					if (dtpDataEntrega.Value != null)
					{
						strWhere += " AND (t_PEDIDO.a_entregar_data_marcada = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(dtpDataEntrega.Value.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador)) + ") ";
					}
				}
				if (cbTransportadora.Text.Trim() != "")
				{
					//obtendo id da transportadora (à esquerda do hífen)
					strIdTransp = cbTransportadora.Text.Substring(0, cbTransportadora.Text.IndexOf("-")).Trim();
					strWhere += " AND (t_PEDIDO.transportadora_id = '" + strIdTransp + "') ";
				}
				if (txtNFe.Text.Trim() != "")
				{
					strWhere += " AND ((" +
									"SELECT" +
										" TOP 1 NFe_numero_NF" +
									" FROM t_NFe_EMISSAO tNE" +
									" WHERE" +
										" (tNE.pedido=t_PEDIDO.pedido)" +
										" AND (tipo_NF = '1')" +
										" AND (st_anulado = 0)" +
										" AND (codigo_retorno_NFe_T1 = 1)" +
									" ORDER BY" +
										" id DESC) = '" + txtNFe.Text + "') ";
				}
				if (txtPedido.Text.Trim() != "")
				{
					strWhere += " AND (t_PEDIDO.pedido = '" + Global.normalizaNumeroPedido(txtPedido.Text) + "') ";
				}

				strFrom = " FROM t_PEDIDO" +
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" +
				" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id) ";
				#endregion

				this.Cursor = Cursors.WaitCursor;
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				#endregion

				#region [ Inicialização ]
				flagPedidoUsarMemorizacaoCompletaEnderecos = ParametroDAO.getCampoInteiroTabelaParametro(Global.Cte.ID_T_PARAMETRO.ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS, 0);
				#endregion

				#region [ Monta o SQL ]
				strSql = "SELECT" +
				" t_PEDIDO.pedido," +
				" t_PEDIDO.data," +
				" t_PEDIDO.a_entregar_data_marcada," +
				" t_PEDIDO.transportadora_id," +
				" t_PEDIDO.st_end_entrega," +
				" t_PEDIDO.EndEtg_endereco," +
				" t_PEDIDO.EndEtg_endereco_numero," +
				" t_PEDIDO.EndEtg_endereco_complemento," +
				" t_PEDIDO.EndEtg_bairro," +
				" t_PEDIDO.EndEtg_cidade," +
				" t_PEDIDO.EndEtg_uf," +
				" t_PEDIDO.EndEtg_cep,";

				if (flagPedidoUsarMemorizacaoCompletaEnderecos == 0)
				{
					strSql +=
					" t_CLIENTE.endereco," +
					" t_CLIENTE.endereco_numero," +
					" t_CLIENTE.endereco_complemento," +
					" t_CLIENTE.bairro," +
					" t_CLIENTE.cidade," +
					" t_CLIENTE.uf," +
					" t_CLIENTE.cep,";
				}
				else
				{
					strSql +=
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.endereco ELSE t_PEDIDO.endereco_logradouro END) AS endereco," +
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.endereco_numero ELSE t_PEDIDO.endereco_numero END) AS endereco_numero," +
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.endereco_complemento ELSE t_PEDIDO.endereco_complemento END) AS endereco_complemento," +
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.bairro ELSE t_PEDIDO.endereco_bairro END) AS bairro," +
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.cidade ELSE t_PEDIDO.endereco_cidade END) AS cidade," +
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.uf ELSE t_PEDIDO.endereco_uf END) AS uf," +
					" (CASE t_PEDIDO.st_memorizacao_completa_enderecos WHEN 0 THEN t_CLIENTE.cep ELSE t_PEDIDO.endereco_cep END) AS cep,";
				}

				strSql +=
				" (" +
					"SELECT" +
						" Count(*)" +
					" FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA tPNES" +
					" WHERE" +
						" (tPNES.pedido=t_PEDIDO.pedido)" +
						" AND (" +
							"(nfe_emitida_status=0)" +
							" OR " +
							"(nfe_emitida_status=1)" +
							")" +
				") AS qtde_solicitacao_emissao_nfe," +
				" (" +
					"SELECT" +
						" TOP 1 NFe_numero_NF" +
					" FROM t_NFe_EMISSAO tNE" +
					" WHERE" +
						" (tNE.pedido=t_PEDIDO.pedido)" +
						" AND (tipo_NF = '1')" +
						" AND (st_anulado = 0)" +
						" AND (codigo_retorno_NFe_T1 = 1)" +
					" ORDER BY" +
						" id DESC" +
				") AS numeroNFe," +
				" (" +
					"SELECT" +
						" TOP 1 NFe_serie_NF" +
					" FROM t_NFe_EMISSAO tNE" +
					" WHERE" +
						" (tNE.pedido=t_PEDIDO.pedido)" +
						" AND (tipo_NF = '1')" +
						" AND (st_anulado = 0)" +
						" AND (codigo_retorno_NFe_T1 = 1)" +
					" ORDER BY" +
						" id DESC" +
				") AS serieNFe " +
				strFrom +
				strWhere +
				" ORDER BY t_PEDIDO.transportadora_id, t_PEDIDO.data, t_PEDIDO.pedido";
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
					grdPesquisa.SuspendLayout();

					grdPesquisa.Rows.Clear();
					if (dtbConsulta.Rows.Count > 0) grdPesquisa.Rows.Add(dtbConsulta.Rows.Count);

					for (int i = 0; i < dtbConsulta.Rows.Count; i++)
					{
						rowConsulta = dtbConsulta.Rows[i];
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_PEDIDO].Value = BD.readToString(rowConsulta["pedido"]);
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_PEDIDO].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
						if (BD.readToString(rowConsulta["st_end_entrega"]) != "0")
						{
							grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CIDADE].Value = BD.readToString(rowConsulta["EndEtg_cidade"]).ToUpper();
							grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_UF].Value = BD.readToString(rowConsulta["EndEtg_uf"]).ToUpper();
						}
						else
						{
							grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CIDADE].Value = BD.readToString(rowConsulta["cidade"]).ToUpper();
							grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_UF].Value = BD.readToString(rowConsulta["uf"]).ToUpper();
						}
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CIDADE].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_UF].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_DATA_ENTREGA].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["a_entregar_data_marcada"]));
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_DATA_ENTREGA].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_TRANSPORTADORA].Value = BD.readToString(rowConsulta["transportadora_id"]);
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_TRANSPORTADORA].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_NFE].Value = BD.readToString(rowConsulta["numeroNFe"]);
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_NFE].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_SERIE].Value = BD.readToString(rowConsulta["serieNFe"]);
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_SERIE].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					}

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdPesquisa.Rows.Count; i++)
					{
						if (grdPesquisa.Rows[i].Selected) grdPesquisa.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					grdPesquisa.ResumeLayout();
				}
				#endregion

				#region [Totais]
				lblTotalRegistros.Text = Global.formataInteiro(dtbConsulta.Rows.Count);
				#endregion

				this.Cursor = Cursors.Default;

				grdPesquisa.Focus();

				// Feedback da conclusão da pesquisa
				SystemSounds.Exclamation.Play();

				return true;
			}
			catch (Exception ex)
			{
				this.Cursor = Cursors.Default;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
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

		#region [ limpaCamposPesquisa ]
		private void limpaCamposPesquisa()
		{
			grdPesquisa.Rows.Clear();
			dtpDataEntrega.Value = DateTime.Now.Date;
			dtpDataEntrega.Checked = false;
			cbTransportadora.SelectedIndex = -1;
			txtNFe.Text = "";
			txtPedido.Text = "";
		}
		#endregion

		#region [ limpaCamposDados ]
		private void limpaCamposDados()
		{
			lblTotalRegistros.Text = "";
		}
		#endregion

		#region [ limpaCamposFiltro ]
		private void limpaCamposFiltro()
		{
			dtpDataEntrega.Value = DateTime.Now.Date;
			dtpDataEntrega.Checked = false;
			cbTransportadora.TabIndex = 0;
			txtNFe.Text = "";
			txtPedido.Text = "";
		}
		#endregion

		#region [ trataBotaoPrinterDialog ]
		private void trataBotaoPrinterDialog()
		{
			printDialog.ShowDialog();
		}
		#endregion

		#region [ trataBotaoMarcarTodos ]
		private void TrataBotaoMarcarTodos()
		{
			_listagemMarcada = !_listagemMarcada;
			for (int i = 0; i < grdPesquisa.Rows.Count; i++)
			{
				grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_CHECKBOX].Value = _listagemMarcada;
			}
		}
		#endregion

		#region [ trataBotaoPesquisar ]
		private bool trataBotaoPesquisar()
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

			if (!executaPesquisa()) return false;

			return true;
		}
		#endregion

		#region [ trataBotaoLimparFiltro ]
		private bool trataBotaoLimparFiltro()
		{
			limpaCamposPesquisa();
			return true;
		}
		#endregion

		#region [ TrataBotaoPastaPDFAgrup ]
		private void TrataBotaoPastaPDFAgrup()
		{
			String strPasta = Application.StartupPath + "\\" + Global.Usuario.strPastaEmitente + "\\PDF_AGRUPADO";
			if (Directory.Exists(strPasta))
			{
				Process.Start(strPasta);
			}
			else
			{
				aviso("Não há arquivos para exibir");
			}
		}
		#endregion

		#region [ TrataBotaoPastaPDFInd ]
		private void TrataBotaoPastaPDFInd()
		{
			String strPasta = Application.StartupPath + "\\" + Global.Usuario.strPastaEmitente + "\\PDF_INDIVIDUAL";
			if (Directory.Exists(strPasta))
			{
				Process.Start(strPasta);
			}
			else
			{
				aviso("Não há arquivos para exibir");
			}
		}
		#endregion

		#region [ trataBotaoGravarUm ]
		private void trataBotaoGravarUm()
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "trataBotaoGravarUm()";
			String strMsgErro;
			String strNomeArquivo = "";
			Log log = new Log();
			#endregion

			try
			{
				if (!haPdfSelecionado())
				{
					aviso("Não há nenhum PDF selecionado!!");
					return;
				}

				if (!confirma("Confirma a gravação dos PDF's selecionados em um ÚNICO arquivo?"))
				{
					return;
				}

				info(ModoExibicaoMensagemRodape.EmExecucao, "obtendo PDFs dos DANFES selecionados");

				if (!PDF.limpaPastaManipulaPDFs())
				{
					aviso("Problemas na preparação para gerar os PDF's: falha ao tentar limpar o diretório de trabalho!");
					return;
				}

				if (!baixaPDFsSelecionados())
				{
					aviso("Os PDFs não foram selecionados");
					return;
				}

				//gerando arquivos agrupados e individuais, para que os dois fiquem disponíveis
				if ((!PDF.concatenaPDFs(ref strNomeArquivo)) || (!PDF.copiaPastaManipulaPDFs()))
				{
					aviso("Problema na geração de um dos arquivos PDFs");
					return;
				}
				else
				{
					if (confirma("Deseja visualizar o arquivo gravado?"))
					{
						PDF.abrePDF(strNomeArquivo);
					}

					atualizaPDFsBaixados();
					trataBotaoPesquisar();
				}

				// Feedback da conclusão da consulta
				SystemSounds.Exclamation.Play();

			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				avisoErro(strMsgErro);
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [trataBotaoGravarVarios ]
		private void trataBotaoGravarVarios()
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "trataBotaoGravarVarios()";
			string YYYYMMDD_HHMMSS = DateTime.Now.ToString(Global.Cte.DataHora.FmtYYYYMMDD) + "_" + DateTime.Now.ToString(Global.Cte.DataHora.FmtHHMMSS);
			String strMsgErro;
			String strNomeArquivo = "";
			Log log = new Log();
			#endregion

			try
			{
				if (!haPdfSelecionado())
				{
					aviso("Não há nenhum PDF selecionado!!");
					return;
				}

				if (!confirma("Confirma a gravação dos PDFs selecionados em ARQUIVOS SEPARADOS?"))
				{
					return;
				}

				info(ModoExibicaoMensagemRodape.EmExecucao, "obtendo PDFs dos DANFES selecionados");

				if (!PDF.limpaPastaManipulaPDFs())
				{
					aviso("Problemas na preparação para gerar os PDF's: falha ao tentar limpar o diretório de trabalho!");
					return;
				}

				if (!baixaPDFsSelecionados())
				{
					aviso("Os PDFs não foram selecionados");
					return;
				}

				//gerando arquivos agrupados e individuais, para que os dois fiquem disponíveis
				if ((!PDF.concatenaPDFs(ref strNomeArquivo)) || (!PDF.copiaPastaManipulaPDFs()))
				{
					aviso("Problema na geração de um dos arquivos PDFs");
					return;
				}

				if (confirma("Deseja abrir a pasta de PDFs para visualizar os arquivos gravados?"))
				{
					Process.Start(Global.Cte.Etc.PathPDFIndividual);
				}

				atualizaPDFsBaixados();
				trataBotaoPesquisar();

				// Feedback da conclusão da consulta
				SystemSounds.Exclamation.Play();
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				avisoErro(strMsgErro);
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoImprimir ]
		private void trataBotaoImprimir()
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "trataBotaoImprimir()";
			bool imprimirInvisivel = false;
			bool imprimirEFechar = true;
			String strMsgErro;
			String strNomeArquivo = "";
			Log log = new Log();
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

			try
			{
				if (!haPdfSelecionado())
				{
					aviso("Não há nenhum PDF selecionado!!");
					return;
				}

				if (!confirma("Confirma a impressão dos PDFs selecionados?"))
				{
					return;
				}

				info(ModoExibicaoMensagemRodape.EmExecucao, "obtendo PDFs dos DANFES selecionados");

				if (!PDF.limpaPastaManipulaPDFs())
				{
					aviso("Problemas na preparação para gerar os PDF's: falha ao tentar limpar o diretório de trabalho!");
					return;
				}

				if (!baixaPDFsSelecionados())
				{
					aviso("Problema ao obter PDFs selecionados");
					return;
				}

				//gerando arquivos agrupados e individuais, para que os dois fiquem disponíveis
				if ((!PDF.concatenaPDFs(ref strNomeArquivo)) || (!PDF.copiaPastaManipulaPDFs()))
				{
					aviso("Problema na geração de um dos arquivos PDFs");
					return;
				}
				else
				{
					PDF.imprimePDF(Global.Cte.Etc.ArqMergePDF, imprimirInvisivel, imprimirEFechar);
					atualizaPDFsBaixados();
					trataBotaoPesquisar();
				}

				// Feedback da conclusão da consulta
				SystemSounds.Exclamation.Play();
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				avisoErro(strMsgErro);
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FDANFEImprime ]

		#region [ FDANFEImprime_Load ]
		private void FDANFEImprime_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCamposPesquisa();
				limpaCamposDados();
				limpaCamposFiltro();

				carregaTransportadoras();

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

		#region [ FDANFEImprime_Shown ]
		private void FDANFEImprime_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Ajusta layout do header do grid (resultado da pesquisa) ]
					grdPesquisa.Columns[GRID_PESQ_COL_CHECKBOX].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdPesquisa.Columns[GRID_PESQ_COL_PEDIDO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdPesquisa.Columns[GRID_PESQ_COL_PEDIDO].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_CIDADE].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdPesquisa.Columns[GRID_PESQ_COL_CIDADE].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_UF].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdPesquisa.Columns[GRID_PESQ_COL_UF].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_DATA_ENTREGA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdPesquisa.Columns[GRID_PESQ_COL_DATA_ENTREGA].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_TRANSPORTADORA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdPesquisa.Columns[GRID_PESQ_COL_TRANSPORTADORA].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_NFE].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdPesquisa.Columns[GRID_PESQ_COL_NFE].ReadOnly = true;
					#endregion

					#region [ Informa Emitente ]
					lblEmit.Text = Global.Usuario.emit;
					#endregion

					#region [ Posiciona foco ]
					btnDummy.Focus();
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

		#region [ FDANFEImprime_FormClosing ]
		private void FDANFEImprime_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (_emProcessamento)
			{
				SystemSounds.Exclamation.Play();
				e.Cancel = true;
				return;
			}

			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnPrinterDialog ]

		#region [ btnPrinterDialog_Click ]
		private void btnPrinterDialog_Click(object sender, EventArgs e)
		{
			trataBotaoPrinterDialog();
		}
		#endregion

		#endregion

		#region [ btnPesquisar ]

		#region [ btnPesquisar_Click ]
		private void btnPesquisar_Click(object sender, EventArgs e)
		{
			trataBotaoPesquisar();
		}
		#endregion

		#endregion

		#region [ btnLimparFiltro ]

		#region [ btnLimparFiltro_Click ]
		private void btnLimparFiltro_Click(object sender, EventArgs e)
		{
			trataBotaoLimparFiltro();
		}
		#endregion

		#endregion

		#region [ btnGravarUm ]

		#region [ btnGravarUm_Click ]
		private void btnGravarUm_Click(object sender, EventArgs e)
		{
			trataBotaoGravarUm();
		}
		#endregion

		#endregion

		#region [ btnGravarVarios ]
		private void btnGravarVarios_Click(object sender, EventArgs e)
		{
			trataBotaoGravarVarios();
		}
		#endregion

		#region [ btnImprimir ]

		#region [ btnImprimir_Click ]
		private void btnImprimir_Click(object sender, EventArgs e)
		{
			trataBotaoImprimir();
		}
		#endregion

		#endregion

		#region [ grdPesquisa ]

		#region [ grdPesquisa_CellContentClick ]
		private void grdPesquisa_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{

			if (e == null) return;
			if (e.ColumnIndex == 0)
			{
				DataGridViewCheckBoxCell chkBox = (DataGridViewCheckBoxCell)this.grdPesquisa[e.ColumnIndex, e.RowIndex];
				if (chkBox.EditingCellFormattedValue.ToString().ToUpper().Equals("TRUE"))
				{
					this.grdPesquisa.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSkyBlue;
				}
				else
				{
					this.grdPesquisa.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Empty;
				}

			}
		}
		#endregion

		#endregion

		#region [ txtNFe ]

		#region [ txtNFe_Enter ]
		private void txtNFe_Enter(object sender, EventArgs e)
		{
			txtNFe.Select(0, txtNFe.Text.Length);
		}
		#endregion

		#region [ txtNFe_Leave ]
		private void txtNFe_Leave(object sender, EventArgs e)
		{
			txtNFe.Text = Global.digitos(txtNFe.Text);
		}
		#endregion

		#region [ txtNFe_KeyPress ]
		private void txtNFe_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtPedido ]

		#region [ txtPedido_Enter ]
		private void txtPedido_Enter(object sender, EventArgs e)
		{
			txtPedido.Select(0, txtPedido.Text.Length);
		}
		#endregion

		#region [ txtPedido_Leave ]
		private void txtPedido_Leave(object sender, EventArgs e)
		{
			#region [ Declarações ]
			string strPedido;
			#endregion

			strPedido = Global.normalizaNumeroPedido(txtPedido.Text);
			if (strPedido.Length > 0)
			{
				txtPedido.Text = strPedido;
			}
			else if (txtPedido.Text.Length > 0)
			{
				avisoErro("Número de pedido inválido!!");
				txtPedido.Focus();
			}
		}
		#endregion

		#region [ txtPedido_KeyPress ]
		private void txtPedido_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroPedido(e.KeyChar);
			if (e.KeyChar != '\0') e.KeyChar = (char)e.KeyChar.ToString().ToUpper()[0];
		}
		#endregion

		#region [ txtPedido_KeyDown ]
		private void txtPedido_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnPesquisar);
		}
		#endregion

		#region [ btnMarcarTodos_Click ]
		private void btnMarcarTodos_Click(object sender, EventArgs e)
		{
			TrataBotaoMarcarTodos();
		}
		#endregion

		#region [ btnPastaPDFAgrup_Click ]
		private void btnPastaPDFAgrup_Click(object sender, EventArgs e)
		{
			TrataBotaoPastaPDFAgrup();
		}
		#endregion

		#region [ btnPastaPDFInd_Click ]
		private void btnPastaPDFInd_Click(object sender, EventArgs e)
		{
			TrataBotaoPastaPDFInd();
		}
		#endregion

		#endregion

		#endregion
	}
	#endregion
}
