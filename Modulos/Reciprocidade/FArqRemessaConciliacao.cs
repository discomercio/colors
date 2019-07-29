#region [ using ]
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FArqRemessaConciliacao : FModelo
	{
		#region [ Atributos ]
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

		private DataSet _dsConsulta = null;
		private ArqRemessa _arquivoRemessa = new ArqRemessa();
		#endregion

		#region [ Constantes ]
		private const int ST_GERACAO_EM_ANDAMENTO = 0;
		private const int ST_GERACAO_SUCESSO = 1;
		private const int ST_GERACAO_FALHA = 2;
		private const int ST_ENVIADO_SERASA_SUCESSO = 1;
		#endregion

		#region [ Construtor ]
		public FArqRemessaConciliacao()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]
		#region [ pathTituloArquivoRemessaValorDefault ]
		private String pathTituloArquivoRemessaValorDefault()
		{
			#region [ Declarações ]
			String strResp;
			#endregion

			strResp = Global.PATH_DEFAULT_TITULO_ARQUIVO_REMESSA;
			if (Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa))
				{
					strResp = Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa;
				}
			}
			return strResp;
		}
		#endregion

		#region [ ajustaPosicaoLblTotalGridBoletos ]
		private void ajustaPosicaoLblTotalGridBoletos()
		{
			lblTotalGridBoletos.Left = grdBoletos.Left + grdBoletos.Width - lblTotalGridBoletos.Width - 3;
			if (Global.isVScrollBarVisible(grdBoletos)) lblTotalGridBoletos.Left -= Global.getVScrollBarWidth(grdBoletos);
		}
		#endregion

		#region [ limpaCamposResposta ]
		private void limpaCamposResposta()
		{
			lblTotalGridBoletos.Text = "";
			lblTotalRegistros.Text = "";
			grdBoletos.Rows.Clear();
			ajustaPosicaoLblTotalGridBoletos();
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtDiretorio.Text = "";
			limpaCamposResposta();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Diretório ]
			if (txtDiretorio.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o diretório em que o arquivo de remessa será gerado!!");
				return false;
			}
			if (!Directory.Exists(txtDiretorio.Text))
			{
				avisoErro("O diretório selecionado para gerar o arquivo de remessa não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoSelecionaDiretorio ]
		private void trataBotaoSelecionaDiretorio()
		{
			DialogResult dr;
			folderBrowserDialog.SelectedPath = txtDiretorio.Text;
			dr = folderBrowserDialog.ShowDialog();
			if (dr != DialogResult.OK) return;
			txtDiretorio.Text = folderBrowserDialog.SelectedPath;
			Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa = folderBrowserDialog.SelectedPath;
		}
		#endregion

		#region [ trataBotaoExecutaConsulta ]
		private bool trataBotaoExecutaConsulta()
		{
			#region [ Declarações ]
			int intIndiceLinha = 0;
			decimal soma = 0;
			#endregion

			#region [ Consistência ]
			if (!consisteCampos()) return false;
			#endregion

			#region [ Verifica se existe item no combobox ]
			if (cmbTitulos.Items.Count == 0)
			{
				aviso("Não há títulos conciliados para remessa");
				return false;
			}
			#endregion

			#region [ Verifica se o usuário selecionou algum item do combobox ]
			if (cmbTitulos.SelectedItem == null)
			{
				avisoErro("Selecione a data do arquivo que contém os títulos a gerar!!");
				return false;
			}
			#endregion

			ComboItemHelper itemSelecionado = (ComboItemHelper)cmbTitulos.SelectedItem;

			info(ModoExibicaoMensagemRodape.EmExecucao, "executando consulta");
			try
			{
				#region [ Limpa o grid ]
				limpaCamposResposta();
				#endregion

				#region [ Obtém dados para gerar o arquivo de remessa ]
				_dsConsulta = ConciliacaoTituloDAO.selecionaBoletosParaRemessa(itemSelecionado.id);
				#endregion

				if (_dsConsulta.Tables["DtbBoleto"].Rows.Count == 0)
				{
					aviso("Não há títulos conciliados para remessa!!");
					return false;
				}

				#region [ Prepara dados p/ exibição no grid ]
				if (_dsConsulta.Tables["DtbBoleto"].Rows.Count > 0) grdBoletos.Rows.Add(_dsConsulta.Tables["DtbBoleto"].Rows.Count);
				foreach (DataRow rowBoleto in _dsConsulta.Tables["DtbBoleto"].Rows)
				{
					String numTitulo = rowBoleto["num_titulo_estendido"].ToString().Trim();

					grdBoletos.Rows[intIndiceLinha].Cells["cnpj"].Value = Global.formataCnpjCpf(BD.readToString(rowBoleto["cnpj_cliente"]));
					grdBoletos.Rows[intIndiceLinha].Cells["num_titulo"].Value = numTitulo.Substring(0, numTitulo.Length - 1) + "-" + numTitulo.Substring(numTitulo.Length - 1);
					grdBoletos.Rows[intIndiceLinha].Cells["data_emissao"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_data_emissao"]);
					
					if (rowBoleto["dt_data_vencto_editado"] != DBNull.Value)
					{
						grdBoletos.Rows[intIndiceLinha].Cells["data_vencimento"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_data_vencto_editado"]);
					}
					else
					{
						grdBoletos.Rows[intIndiceLinha].Cells["data_vencimento"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_data_vencto_original"]);
					}

					if ((decimal)rowBoleto["vl_valor_titulo_editado"] != 0)
					{
						grdBoletos.Rows[intIndiceLinha].Cells["valor"].Value = Global.formataMoeda((decimal)rowBoleto["vl_valor_titulo_editado"]);
					}
					else
					{
						grdBoletos.Rows[intIndiceLinha].Cells["valor"].Value = Global.formataMoeda((decimal)rowBoleto["vl_valor_titulo_original"]);
					}

					object dtPagto = rowBoleto["dt_data_pagto_editado"];
					if (dtPagto != DBNull.Value)
					{
						grdBoletos.Rows[intIndiceLinha].Cells["data_pagamento"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_data_pagto_editado"]);
					}
					else
					{
						grdBoletos.Rows[intIndiceLinha].Cells["data_pagamento"].Value = "";
					}

					intIndiceLinha++;

					if (BD.readToDecimal(rowBoleto["vl_valor_titulo_editado"]) != 0)
					{
						soma += BD.readToDecimal(rowBoleto["vl_valor_titulo_editado"]);
					}
					else
					{
						soma += BD.readToDecimal(rowBoleto["vl_valor_titulo_original"]);
					}
				}
				#endregion

				#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
				for (int i = 0; i < grdBoletos.Rows.Count; i++)
				{
					if (grdBoletos.Rows[i].Selected) grdBoletos.Rows[i].Selected = false;
				}
				#endregion

				ajustaPosicaoLblTotalGridBoletos();
				lblTotalGridBoletos.Text = Global.formataMoeda(soma);
				lblTotalRegistros.Text = Global.formataInteiro(intIndiceLinha);
				return true;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoGravaArqRemessa ]
		private void trataBotaoGravaArqRemessa()
		{
			#region [ Declarações ]
			bool blnSucesso;
			bool blnGerouNsu;
			int qtdeRegProcessado;
			int percProgressoAtual;
			int percProgressoAnterior;
			String strMsgErro = "";
			String strMsgProgresso;
			String strNomeBasicoArqRemessa;
			String strNomeCompletoArqRemessa;
			String strPathCompleto = "";
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			StreamWriter sw;
			DateTime dtInicioProcessamento;
			DateTime dtFimProcessamento;
			int totalRegTitulos = 0;
			#endregion

			#region [ Consistência ]
			if (_dsConsulta == null)
			{
				avisoErro("Nenhuma consulta foi realizada!!");
				return;
			}

			if (_dsConsulta.Tables["DtbBoleto"].Rows.Count == 0)
			{
				avisoErro("Não há títulos para gerar!!");
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma a geração do arquivo de remessa?")) return;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "gerando arquivo de remessa");
			try
			{
				dtInicioProcessamento = DateTime.Now;
				_arquivoRemessa = new ArqRemessa();

				#region [ Prepara nome do arquivo de remessa ]
				strNomeBasicoArqRemessa = "RemessaSerasaConciliacao_" +
										  Global.digitos(Global.formataDataYyyyMmDdComSeparador(DateTime.Now)) +
										  ".txt";
				#endregion

				#region [ Obtém path completo ]
				strPathCompleto = txtDiretorio.Text;

				if (!Directory.Exists(strPathCompleto))
				{
					Directory.CreateDirectory(strPathCompleto);
					if (!Directory.Exists(strPathCompleto))
					{
						avisoErro("Falha ao tentar criar o diretório:\n" + strPathCompleto);
						return;
					}
				}
				#endregion

				#region [ Nome completo do arquivo de remessa ]
				strNomeCompletoArqRemessa = Global.barraInvertidaAdd(strPathCompleto) + strNomeBasicoArqRemessa;
				#endregion

				#region [ Verifica se já existe arquivo c/ o mesmo nome ]
				if (File.Exists(strNomeCompletoArqRemessa))
				{
					avisoErro("Já existe um arquivo no diretório especificado com este nome!!\n" + strNomeCompletoArqRemessa);
					return;
				}
				#endregion

				sw = new StreamWriter(strNomeCompletoArqRemessa, true, encode);

				try
				{
					#region [ Monta Header ]
					//Para conciliação, o "período inicial" é definido pela constante CONCILIA.                                        
					String strFim = BD.readToString(_dsConsulta.Tables["DtbBoleto"].Rows[0].GetParentRow("dtbArqConciliacaoInput_dtbBoleto")["s_data_final_periodo"]);
					DateTime fim = Global.converteYyyyMmDdSemSeparadorParaDateTime(strFim);

					ArqRemessa.LinhaHeader header = new ArqRemessa.LinhaHeader(fim);
					_arquivoRemessa.linhaHeader = header;
					sw.WriteLine(header.ToString());
					#endregion

					#region [ Monta os registros do arquivo de remessa ]
					#region [Títulos]
					qtdeRegProcessado = 0;
					percProgressoAnterior = 0;
					foreach (DataRow rowBoleto in _dsConsulta.Tables["DtbBoleto"].Rows)
					{
						qtdeRegProcessado++;
						percProgressoAtual = 100 * qtdeRegProcessado / _dsConsulta.Tables["DtbBoleto"].Rows.Count;
						if (percProgressoAtual != percProgressoAnterior)
						{
							percProgressoAnterior = percProgressoAtual;
							strMsgProgresso = "Processando registro: " + qtdeRegProcessado.ToString() + "   (" + percProgressoAtual.ToString() + "%)";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							Application.DoEvents();
						}

						int concTituloId = (int)rowBoleto["id"];
						String sacadoPJ = rowBoleto["cnpj_cliente"].ToString();
						String numTitulo = rowBoleto["num_titulo_estendido"].ToString().Trim();
						DateTime dataEmissao = (DateTime)rowBoleto["dt_data_emissao"];
						
						Decimal valorTitulo;
						String strValorTitulo;
						if ((Decimal)rowBoleto["vl_valor_titulo_editado"] != 0)
						{
							valorTitulo = (Decimal)rowBoleto["vl_valor_titulo_editado"];
							strValorTitulo = BD.readToString(rowBoleto["s_valor_titulo_editado"]);
						}
						else
						{
							valorTitulo = (Decimal)rowBoleto["vl_valor_titulo_original"];
							strValorTitulo = BD.readToString(rowBoleto["s_valor_titulo_original"]);
						}

						DateTime dataVencimento;
						if (rowBoleto["dt_data_vencto_editado"] != DBNull.Value)
						{
							dataVencimento = BD.readToDateTime(rowBoleto["dt_data_vencto_editado"]);
						}
						else
						{
							dataVencimento = BD.readToDateTime(rowBoleto["dt_data_vencto_original"]);
						}

						DateTime dataPagamento;
						if (rowBoleto["dt_data_pagto_editado"] != DBNull.Value)
						{
							dataPagamento = BD.readToDateTime(rowBoleto["dt_data_pagto_editado"]);
						}
						else
						{
							dataPagamento = DateTime.MinValue;
						}

						ArqRemessa.DetalheTitulo dt = new ArqRemessa.DetalheTitulo(concTituloId, 0, sacadoPJ, numTitulo, dataEmissao, valorTitulo, dataVencimento, dataPagamento);

						if (strValorTitulo.Equals("9999999999999"))
						{
							dt.isTituloExcluido = true;
						}

						_arquivoRemessa.addDetalheTitulo(dt);
						sw.WriteLine(dt.ToString());
						totalRegTitulos++;
					}
					#endregion
					#endregion

					#region [ Monta Trailler ]
					ArqRemessa.LinhaTrailler t = new ArqRemessa.LinhaTrailler(0, totalRegTitulos);
					_arquivoRemessa.linhaTrailler = t;
					sw.Write(t.ToString());
					#endregion
				}
				finally
				{
					sw.Flush();
					sw.Close();
				}

				dtFimProcessamento = DateTime.Now;
				int id_serasa_arq_conciliacao_output = 0;

				try
				{
					BD.iniciaTransacao();
					//obtem o NSU e passa por ref na funcao abaixo
					blnGerouNsu = false;
					blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_ARQ_CONCILIACAO_OUTPUT, ref id_serasa_arq_conciliacao_output, ref strMsgErro);
					if (!blnGerouNsu)
					{
						throw new Exception("Falha ao tentar gerar o NSU para o registro de histórico de arquivos de conciliação!!\n" + strMsgErro);
					}

					//insere um registro da tabela t_SERASA_ARQ_CONCILIACAO_OUTPUT
					if (!ArqConciliacaoOutputDAO.insere(id_serasa_arq_conciliacao_output,
														dtInicioProcessamento,
														dtInicioProcessamento,
														Global.Usuario.usuario,
														ArqRemessa.LinhaHeader.CNPJ_EMPRESA_CONVENIADA,
														Global.formataDataYyyyMmDdSemSeparador(_arquivoRemessa.linhaHeader.dataFim),
														_arquivoRemessa.linhaHeader.dataFim,
														ArqRemessa.LinhaHeader.PERIODICIDADE_REMESSA,
														ArqRemessa.LinhaHeader.RESERVADO_SERASA,
														ArqRemessa.LinhaHeader.ID_GRUPO_RELATO_SEGMENTO,
														ArqRemessa.LinhaHeader.ID_VERSAO_LAYOUT,
														ArqRemessa.LinhaHeader.NUM_VERSAO_LAYOUT,
														_arquivoRemessa.linhaTrailler.qtdeRegTitulo,
														dtFimProcessamento.Subtract(dtInicioProcessamento).Seconds,
														strNomeBasicoArqRemessa,
														txtDiretorio.Text,
														ST_GERACAO_EM_ANDAMENTO,
														null))
					{
						throw new Exception("Falha ao tentar inserir um registro na tabela t_SERASA_ARQ_CONCILIACAO_OUTPUT");
					}

					BD.commitTransacao();
					blnSucesso = true;
				}
				catch (Exception e)
				{
					Global.gravaLogAtividade(e.ToString());
					strMsgErro = e.ToString();
					blnSucesso = false;
				}

				if (!blnSucesso)
				{
					BD.rollbackTransacao();

					#region [ Se o arquivo de remessa foi gravado, renomeia para indicar que houve uma falha ]
					if (File.Exists(strNomeCompletoArqRemessa)) File.Move(strNomeCompletoArqRemessa, strNomeCompletoArqRemessa + ".ERR");
					#endregion

					info(ModoExibicaoMensagemRodape.Normal);
					avisoErro(strMsgErro);
				}
				else
				{
					try
					{
						BD.iniciaTransacao();
						qtdeRegProcessado = 0;
						percProgressoAnterior = 0;
						foreach (ArqRemessa.DetalheTitulo detTitulo in _arquivoRemessa.detTituloList)
						{
							qtdeRegProcessado++;
							percProgressoAtual = 100 * qtdeRegProcessado / _arquivoRemessa.detTituloList.Count;
							if (percProgressoAtual != percProgressoAnterior)
							{
								percProgressoAnterior = percProgressoAtual;
								strMsgProgresso = "Atualizando banco de dados: " + percProgressoAtual.ToString() + "%";
								info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
								Application.DoEvents();
							}

							//atualiza id_serasa_arq_conciliacao_output e st_enviado_serasa da tabela t_SERASA_CONCILIACAO_TITULO
							if (!ConciliacaoTituloDAO.atualizaArqConciliacaoOutputEStEnviadoSerasa(id_serasa_arq_conciliacao_output,
																								   ST_ENVIADO_SERASA_SUCESSO,
																								   detTitulo.tituloMovimentoId))
							{
								throw new Exception("Falha ao atualizar um registro para o título do arquivo de conciliação!!");
							}
						}

						if (!ArqConciliacaoOutputDAO.atualizaStatusGeracao(ST_GERACAO_SUCESSO, null, id_serasa_arq_conciliacao_output)) //1 = Gerado com sucesso
						{
							throw new Exception("Falha ao atualizar o status da geração do arquivo de conciliação!!");
						}

						blnSucesso = true;
					}
					catch (Exception e)
					{
						Global.gravaLogAtividade(e.ToString());
						strMsgErro = e.ToString();
						blnSucesso = false;
					}

					if (blnSucesso)
					{
						BD.commitTransacao();
						Global.gravaLogAtividade("Arquivo de remessa ID " + id_serasa_arq_conciliacao_output + " gerado com sucesso!");
						info(ModoExibicaoMensagemRodape.Normal);
						aviso("Arquivo de remessa gerado com sucesso!!\n\n" + strNomeCompletoArqRemessa);
						Close();
					}
					else
					{
						BD.rollbackTransacao();

						if (File.Exists(strNomeCompletoArqRemessa)) File.Move(strNomeCompletoArqRemessa, strNomeCompletoArqRemessa + ".ERR");
						info(ModoExibicaoMensagemRodape.Normal);
						avisoErro("Não foi possível gerar o arquivo de remessa!");

						//tenta atualizar o status de geração do arquivo da tabela t_SERASA_ARQ_REMESSA_NORMAL
						try
						{
							BD.iniciaTransacao();
							if (!ArqConciliacaoOutputDAO.atualizaStatusGeracao(ST_GERACAO_FALHA, "Falha na geração do arquivo", id_serasa_arq_conciliacao_output))
							{
								throw new Exception("Falha na tentativa de atualizar o status da geração do arquivo de remessa!");
							}
							BD.commitTransacao();
						}
						catch (Exception e)
						{
							Global.gravaLogAtividade(e.ToString());
						}
					}
				}
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ carregaComboDatas ]
		private void carregaComboDatas()
		{
			const int ID_COLUNA = 0;
			const int DATA_HORA_COLUNA = 1;

			try
			{
				DataTable dtConsulta = ArqConciliacaoInputDAO.selecionaDatasParaCombobox();

				foreach (DataRow row in dtConsulta.Rows)
				{
					DateTime dataHora = BD.readToDateTime(row[DATA_HORA_COLUNA]);
					if (dataHora != DateTime.MinValue)
					{
						int id = BD.readToInt(row[ID_COLUNA]);
						ComboItemHelper item = new ComboItemHelper(id, dataHora);
						cmbTitulos.Items.Add(item);
					}
				}
			}
			catch (Exception e)
			{
				Global.gravaLogAtividade(e.ToString());
				avisoErro(e.ToString());
			}
		}
		#endregion
		#endregion

		#region [ Eventos ]
		#region [ FArqRemessaConciliacao ]
		#region [ FArqRemessaConciliacao_Load ]
		private void FArqRemessaConciliacao_Load(object sender, EventArgs e)
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

		#region [ FArqRemessaConciliacao_Shown ]
		private void FArqRemessaConciliacao_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

					#region [ Preenchimento dos campos ]
					txtDiretorio.Text = pathTituloArquivoRemessaValorDefault();
					carregaComboDatas();
					if (cmbTitulos.Items.Count == 1)
					{
						cmbTitulos.SelectedIndex = 0;
						trataBotaoExecutaConsulta();
					}
					#endregion

					#region [ Ajusta o label com o valor total ]
					ajustaPosicaoLblTotalGridBoletos();
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

		#region [ FArqRemessaConciliacao_KeyDown ]
		private void FArqRemessaConciliacao_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				trataBotaoExecutaConsulta();
				return;
			}
		}
		#endregion

		#region [ FArqRemessaConciliacao_FormClosing ]
		private void FArqRemessaConciliacao_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain._fMain.Location = this.Location;
			FMain._fMain.Visible = true;
			this.Visible = false;
		}
		#endregion
		#endregion

		#region [ btnSelecionaDiretorio ]
		#region [ btnSelecionaDiretorio_Click ]
		private void btnSelecionaDiretorio_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaDiretorio();
		}
		#endregion
		#endregion

		#region [ btnExecutaConsulta ]
		#region [ btnExecutaConsulta_Click ]
		private void btnExecutaConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoExecutaConsulta();
		}
		#endregion
		#endregion

		#region [ btnGravaArqRemessa ]
		#region [ btnGravaArqRemessa_Click ]
		private void btnGravaArqRemessa_Click(object sender, EventArgs e)
		{
			trataBotaoGravaArqRemessa();
		}
		#endregion
		#endregion

		#region [ txtDiretorio ]
		#region [ txtDiretorio_Enter ]
		private void txtDiretorio_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDiretorio_DoubleClick ]
		private void txtDiretorio_DoubleClick(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion
		#endregion
		#endregion
	}
}
