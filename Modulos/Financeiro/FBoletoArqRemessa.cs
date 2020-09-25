#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	public partial class FBoletoArqRemessa : Financeiro.FModelo
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

		private BoletoCedente _boletoCedenteSelecionado = null;
		private DataSet _dsConsulta = null;
		#endregion

		#region [ Construtor ]
		public FBoletoArqRemessa()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ comboBoletoCedentePosicionaDefault ]
		private bool comboBoletoCedentePosicionaDefault()
		{
			bool blnHaDefault = false;
			DsDataSource.DtbBoletoCedenteComboRow rowBoletoCedente;

			foreach (System.Data.DataRowView item in cbBoletoCedente.Items)
			{
				rowBoletoCedente = (DsDataSource.DtbBoletoCedenteComboRow)item.Row;
				if (rowBoletoCedente.id == Global.Usuario.Defaults.FBoletoArqRemessa.boletoCedente)
				{
					cbBoletoCedente.SelectedIndex = cbBoletoCedente.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}
			return blnHaDefault;
		}
		#endregion

		#region [ montaPathBoletoArquivoRemessaBoletoCedente ]
		private bool montaPathBoletoArquivoRemessaBoletoCedente(BoletoCedente boletoCedenteSelecionado, String strPathSelecionado, ref String strPathBase, ref String strPathCompleto, ref String strMsgErro)
		{
			#region [ Declarações ]
			const String NOME_BASE_DIRETORIO_CEDENTE = "Cedente_";
			String strNomeDiretorioCedente;
			String strPath;
			String[] vDir;
			#endregion

			#region [ Inicializa variáveis de retorno ]
			strPathBase = "";
			strPathCompleto = "";
			strMsgErro = "";
			#endregion

			#region [ Consistências ]
			if (boletoCedenteSelecionado.id <= 0)
			{
				strMsgErro = "Não há informações sobre o cedente!!";
				return false;
			}
			if (boletoCedenteSelecionado.apelido == null)
			{
				strMsgErro = "Cedente não está com o campo 'apelido' cadastrado!!";
				return false;
			}
			if (boletoCedenteSelecionado.apelido.Length == 0)
			{
				strMsgErro = "Cedente está com o campo 'apelido' vazio!!";
				return false;
			}
			if (strPathSelecionado == null)
			{
				strMsgErro = "Não foi informado o diretório para o arquivo de remessa!!";
				return false;
			}
			if (strPathSelecionado.Length == 0)
			{
				strMsgErro = "Diretório inválido para o arquivo de remessa!!";
				return false;
			}
			#endregion

			strNomeDiretorioCedente = NOME_BASE_DIRETORIO_CEDENTE + boletoCedenteSelecionado.id.ToString() + "_" + boletoCedenteSelecionado.apelido.ToUpper();
			strPath = Global.barraInvertidaAdd(strPathSelecionado);

			#region [ Obtém o path base (sem o diretório do cedente) ]
			vDir = Global.barraInvertidaDel(strPath).Split('\\');
			if (vDir.Length > 0)
			{
				// Se o trecho final do diretório contém o a pasta de um dos cedentes, retira essa pasta do path
				if (vDir[vDir.Length - 1].ToUpper().Contains(NOME_BASE_DIRETORIO_CEDENTE.ToUpper()))
				{
					strPath = String.Join("\\", vDir, 0, vDir.Length - 1);
					strPath = Global.barraInvertidaAdd(strPath);
				}
			}

			// Retorna o path base (sem o diretório do cedente)
			strPathBase = strPath;
			#endregion

			#region [ Obtém o path completo (com o diretório do cedente) ]
			if (!strPath.ToUpper().EndsWith(Global.barraInvertidaAdd(strNomeDiretorioCedente.ToUpper())))
			{
				strPath += Global.barraInvertidaAdd(strNomeDiretorioCedente);
			}

			// Retorna o path completo (com o diretório do cedente)
			strPathCompleto = strPath;
			#endregion

			return true;
		}
		#endregion

		#region [ pathBoletoArquivoRemessaValorDefault ]
		private String pathBoletoArquivoRemessaValorDefault()
		{
			#region [ Declarações ]
			String strResp;
			#endregion

			strResp = Global.PATH_BOLETO_ARQUIVO_REMESSA;
			if (Global.Usuario.Defaults.FBoletoArqRemessa.pathBoletoArquivoRemessa.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FBoletoArqRemessa.pathBoletoArquivoRemessa))
				{
					strResp = Global.Usuario.Defaults.FBoletoArqRemessa.pathBoletoArquivoRemessa;
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
			lblTotalSerieBoletos.Text = "";
			lblTotalParcelas.Text = "";
			lblTotalGridBoletos.Text = "";
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
			#region [ Cedente ]
			if (cbBoletoCedente.SelectedIndex == -1)
			{
				avisoErro("É necessário informar a conta do cedente!!");
				cbBoletoCedente.Focus();
				return false;
			}
			#endregion

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
		}
		#endregion
		
		#region [ trataBotaoExecutaConsulta ]
		private bool trataBotaoExecutaConsulta()
		{
			#region [ Declarações ]
			short id_boleto_cedente;
			int intIndiceLinha = 0;
			int intTotalSerieBoletos = 0;
			int intTotalParcelas = 0;
			decimal vlSubTotal;
			decimal vlTotalGeral = 0m;
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

			#region [ Consistência ]
			if (!consisteCampos()) return false;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "executando consulta");
			try
			{
				#region [ Limpa o grid ]
				limpaCamposResposta();
				#endregion

				#region [ Obtém dados do cedente ]
				id_boleto_cedente = (short)Global.converteInteiro(cbBoletoCedente.SelectedValue.ToString());
				_boletoCedenteSelecionado = BoletoCedenteDAO.getBoletoCedente(id_boleto_cedente);
				#endregion

				#region [ Obtém dados para gerar o arquivo de remessa ]
				_dsConsulta = BoletoDAO.selecionaBoletosParaArqRemessa(id_boleto_cedente);
				#endregion

				#region [ Prepara dados p/ exibição no grid ]
				if (_dsConsulta.Tables["DtbFinBoleto"].Rows.Count > 0) grdBoletos.Rows.Add(_dsConsulta.Tables["DtbFinBoleto"].Rows.Count);
				foreach (DsDataSource.DtbFinBoletoRow rowBoleto in _dsConsulta.Tables["DtbFinBoleto"].Rows)
				{
					intTotalSerieBoletos++;
					grdBoletos.Rows[intIndiceLinha].Cells["id_boleto"].Value = rowBoleto.id.ToString();
					grdBoletos.Rows[intIndiceLinha].Cells["cliente"].Value = rowBoleto.nome_sacado + " (" + Global.formataCnpjCpf(rowBoleto.num_inscricao_sacado) + ")";

					vlSubTotal = 0m;
					grdBoletos.Rows[intIndiceLinha].Cells["num_documento"].Value = "";
					grdBoletos.Rows[intIndiceLinha].Cells["parcelas"].Value = "";
					foreach (DsDataSource.DtbFinBoletoItemRow rowBoletoItem in rowBoleto.GetChildRows("DtbFinBoleto_DtbFinBoletoItem"))
					{
						intTotalParcelas++;
						if (grdBoletos.Rows[intIndiceLinha].Cells["num_documento"].Value.ToString().Length > 0)
							grdBoletos.Rows[intIndiceLinha].Cells["num_documento"].Value += "\n";

						grdBoletos.Rows[intIndiceLinha].Cells["num_documento"].Value += rowBoletoItem.numero_documento;

						if (grdBoletos.Rows[intIndiceLinha].Cells["parcelas"].Value.ToString().Length > 0)
							grdBoletos.Rows[intIndiceLinha].Cells["parcelas"].Value += "\n";

						grdBoletos.Rows[intIndiceLinha].Cells["parcelas"].Value +=
								Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto) +
								Global.formataMoeda(rowBoletoItem.valor).PadLeft(18, ' ');

						vlSubTotal += rowBoletoItem.valor;
						vlTotalGeral += rowBoletoItem.valor;
					}

					grdBoletos.Rows[intIndiceLinha].Cells["parcelas"].Value += "\n" + "".PadLeft(10 + 18, '=') + "\n" + " ".PadLeft(10, ' ') + Global.formataMoeda(vlSubTotal).PadLeft(18, ' ');

					intIndiceLinha++;
				}
				#endregion

				#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
				for (int i = 0; i < grdBoletos.Rows.Count; i++)
				{
					if (grdBoletos.Rows[i].Selected) grdBoletos.Rows[i].Selected = false;
				}
				#endregion

				ajustaPosicaoLblTotalGridBoletos();
				lblTotalSerieBoletos.Text = Global.formataInteiro(intTotalSerieBoletos);
				lblTotalParcelas.Text = Global.formataInteiro(intTotalParcelas);
				lblTotalGridBoletos.Text = Global.formataMoeda(vlTotalGeral);

				return true;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoCancela ]
		private void trataBotaoCancela()
		{
			#region [ Declarações ]
			bool blnSucesso;
			int intLinhaGridSelecionado = -1;
			int id_boleto_selecionado;
			decimal vlTotal = 0m;
			String strAux;
			String strDescricaoLog = "";
			String strDescricaoLogAux = "";
			String strMsgErro = "";
			String strMsgErroLog = "";
			DsDataSource.DtbFinBoletoRow rowBoleto = null;
			DataRow[] vRowsSelect;
			FinLog finLog = new FinLog();
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
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

			#region [ Obtém índice no grid do boleto selecionado ]
			for (int i = 0; i < grdBoletos.Rows.Count; i++)
			{
				if (grdBoletos.Rows[i].Selected)
				{
					intLinhaGridSelecionado = i;
					break;
				}
			}

			if (intLinhaGridSelecionado < 0)
			{
				avisoErro("Nenhum boleto foi selecionado!!");
				return;
			}
			#endregion

			#region [ Confirmação ]
			strAux = "Confirma o cancelamento do boleto selecionado?\n" + grdBoletos.Rows[intLinhaGridSelecionado].Cells["cliente"].Value + "\nATENÇÃO: esta operação cancela definitivamente o boleto!! Para reeditar os dados, acione a operação \"Desfazer boleto\"!!" + "\nDigite a senha para confirmar a operação!!";
			fAutorizacao = new FAutorizacao(strAux);
			drAutorizacao = fAutorizacao.ShowDialog();
			if (drAutorizacao != DialogResult.OK)
			{
				avisoErro("Operação não confirmada!!");
				return;
			}
			if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
			{
				avisoErro("Senha inválida!!\nA operação não foi realizada!!");
				return;
			}
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "executando cancelamento do boleto");
			try
			{
				#region [ Obtém id do boleto selecionado ]
				id_boleto_selecionado = (int)Global.converteInteiro(grdBoletos.Rows[intLinhaGridSelecionado].Cells["id_boleto"].Value.ToString());
				if (id_boleto_selecionado == 0)
				{
					avisoErro("Falha ao obter o id do registro a ser cancelado!!");
					return;
				}
				#endregion

				#region [ Obtém registro principal do boleto ]
				vRowsSelect = _dsConsulta.Tables["DtbFinBoleto"].Select("id=" + id_boleto_selecionado.ToString());
				if (vRowsSelect.Length != 1)
				{
					throw new Exception("Falha ao obter o registro referente ao boleto que será cancelado!!");
				}
				rowBoleto = (DsDataSource.DtbFinBoletoRow)vRowsSelect[0];
				#endregion

				#region [ Cancela o boleto e suas parcelas ]
				blnSucesso = false;
				try
				{
					#region [ Inicia a transação ]
					BD.iniciaTransacao();
					#endregion

					#region [ Cancela o registro do boleto ]
					if (!BoletoDAO.marcaBoletoCanceladoManual(Global.Usuario.usuario,
															  id_boleto_selecionado,
															  ref strMsgErro))
					{
						throw new Exception("Falha ao marcar o registro id=" + id_boleto_selecionado.ToString() + " do boleto como cancelado!!\n" + strMsgErro);
					}
					strDescricaoLog = "Boleto id=" + id_boleto_selecionado.ToString();
					#endregion

					#region [ Cancela os registros da parcelas ]
					foreach (DsDataSource.DtbFinBoletoItemRow rowBoletoItem in rowBoleto.GetChildRows("DtbFinBoleto_DtbFinBoletoItem"))
					{
						vlTotal += rowBoletoItem.valor;
						if (!BoletoDAO.marcaBoletoItemCanceladoManual(Global.Usuario.usuario,
																	 rowBoletoItem.id,
																	 ref strMsgErro))
						{
							throw new Exception("Falha ao marcar o registro id=" + rowBoletoItem.id.ToString() + " da parcela do boleto como cancelado!!\n" + strMsgErro);
						}
						if (strDescricaoLogAux.Length > 0) strDescricaoLogAux += ", ";
						strDescricaoLogAux += "id=" + rowBoletoItem.id.ToString() + " (" + Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto) + " " + Global.formataMoeda(rowBoletoItem.valor) + ")";
					}
					strDescricaoLog += "; Parcelas: " + strDescricaoLogAux + "; Valor total cancelado: " + Global.formataMoeda(vlTotal);
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					strMsgErro = ex.ToString();
					blnSucesso = false;
				}
				#endregion

				#region [ Finaliza a transação ]
				if (blnSucesso)
				{
					BD.commitTransacao();

					#region [ Grava o log no BD ]
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_CANCELA_MANUAL;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.id_registro_origem = id_boleto_selecionado;
					finLog.id_cliente = rowBoleto.id_cliente;
					finLog.cnpj_cpf = rowBoleto.num_inscricao_sacado;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion
				}
				else
				{
					BD.rollbackTransacao();
					strMsgErro = "Falha ao marcar o registro do boleto como cancelado!!\n\n" + strMsgErro;
					avisoErro(strMsgErro);
				}
				#endregion

				#region [ Atualiza o grid ]
				trataBotaoExecutaConsulta();
				#endregion
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
			int id_boleto_arq_remessa = 0;
			int intNumSequencialRegistro = 0;
			int intNumSequencialRemessa = 0;
			int intIndiceArquivoRemessaNoDia = 0;
			int intTotalSerieBoletos = 0;
			int intTotalParcelas = 0;
			decimal vlTotal = 0m;
			bool blnSucesso;
			bool blnGerouNsu;
			String strDescricaoLog = "";
			String strMsgErro = "";
			String strMsgErroAux = "";
			String strMsgErroLog = "";
			String strNomeBasicoArqRemessa;
			String strNomeCompletoArqRemessa;
			String strPathBase = "";
			String strPathCompleto = "";
            B237HeaderArqRemessa b237LinhaHeader;
            B237TraillerArqRemessa b237LinhaTrailler;
            B237RegTipo1ArqRemessa b237LinhaTipo1;
            B237RegTipo2ArqRemessa b237LinhaTipo2;
            B422HeaderArqRemessa b422LinhaHeader;
            B422TraillerArqRemessa b422LinhaTrailler;
            B422RegTipo1ArqRemessa b422LinhaTipo1;
            B422RegTipo2ArqRemessa b422LinhaTipo2;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			StreamWriter sw;
			FinLog finLog = new FinLog();
			BoletoArqRemessa boletoArqRemessa = new BoletoArqRemessa();
			DateTime dtInicioProcessamento;
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

			#region [ Consistência ]
			if (_dsConsulta == null)
			{
				avisoErro("Nenhuma consulta foi realizada!!");
				return;
			}

			if (_dsConsulta.Tables["DtbFinBoleto"].Rows.Count == 0)
			{
				avisoErro("Não há boletos para gerar!!");
				return;
			}

			if (_boletoCedenteSelecionado == null)
			{
				avisoErro("Os dados do cedente não foram obtidos corretamente do banco de dados!!");
				return;
			}

			if (((short)Global.converteInteiro(cbBoletoCedente.SelectedValue.ToString())) != _boletoCedenteSelecionado.id)
			{
				avisoErro("A última consulta executada não foi do cedente que está selecionado atualmente!!\nExecute novamente a consulta antes de gerar o arquivo de remessa!!");
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

				#region [ Gera índice para compor nome do arquivo ]
				if (!BoletoCedenteDAO.geraIndiceArqRemessaNoDia((short)_boletoCedenteSelecionado.id, ref intIndiceArquivoRemessaNoDia, ref strMsgErro))
				{
					avisoErro("Falha ao tentar gerar o número sequencial diário para compor o nome do arquivo de remessa!!\n\n" + strMsgErro);
					return;
				}
				#endregion

				#region [ Prepara nome do arquivo de remessa ]
				strNomeBasicoArqRemessa = "CB" +
										  Texto.leftStr(Global.digitos(Global.formataDataDdMmYyComSeparador(DateTime.Now)), 4) +
										  intIndiceArquivoRemessaNoDia.ToString().PadLeft(2, '0') +
										  ".REM";
				#endregion

				#region [ Obtém path completo ]
				if (!montaPathBoletoArquivoRemessaBoletoCedente(_boletoCedenteSelecionado, txtDiretorio.Text, ref strPathBase, ref strPathCompleto, ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n\n" + strMsgErro;
					strMsgErro = "Falha ao tentar montar o path completo para o arquivo de remessa do cedente '" + _boletoCedenteSelecionado.apelido + "'!!" + strMsgErro;
					avisoErro(strMsgErro);
					return;
				}

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

				#region [ Gera próximo número sequencial de remessa ]
				if (!BoletoCedenteDAO.geraNumSequencialRemessa((short)_boletoCedenteSelecionado.id, ref intNumSequencialRemessa, ref strMsgErro))
				{
					avisoErro("Falha ao gerar o número sequencial de remessa!!\n\n" + strMsgErro);
					return;
				}
				#endregion

				#region [ Dados para o histórico de arquivos de remessa ]
				blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO_ARQ_REMESSA, ref id_boleto_arq_remessa, ref strMsgErro);
				if (!blnGerouNsu)
				{
					avisoErro("Falha ao tentar gerar o NSU para o registro de histórico de arquivos de remessa!!\n" + strMsgErro);
					return;
				}
				boletoArqRemessa.id = id_boleto_arq_remessa;
				boletoArqRemessa.nsu_arq_remessa = intNumSequencialRemessa;
				boletoArqRemessa.nome_arq_remessa = strNomeBasicoArqRemessa;
				boletoArqRemessa.caminho_arq_remessa = Global.barraInvertidaDel(strPathCompleto);
				boletoArqRemessa.id_boleto_cedente = (short)_boletoCedenteSelecionado.id;
				boletoArqRemessa.codigo_empresa = _boletoCedenteSelecionado.codigo_empresa;
				boletoArqRemessa.nome_empresa = _boletoCedenteSelecionado.nome_empresa;
				boletoArqRemessa.num_banco = _boletoCedenteSelecionado.num_banco;
				boletoArqRemessa.nome_banco = _boletoCedenteSelecionado.nome_banco.ToUpper();
				boletoArqRemessa.agencia = _boletoCedenteSelecionado.agencia;
				boletoArqRemessa.digito_agencia = _boletoCedenteSelecionado.digito_agencia;
				boletoArqRemessa.conta = _boletoCedenteSelecionado.conta;
				boletoArqRemessa.digito_conta = _boletoCedenteSelecionado.digito_conta;
				boletoArqRemessa.carteira = _boletoCedenteSelecionado.carteira;
				#endregion

				sw = new StreamWriter(strNomeCompletoArqRemessa, true, encode);
                try
                {
                    #region [ Monta Header ]
                    intNumSequencialRegistro++;

                    if (_boletoCedenteSelecionado.num_banco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
                    {
                        #region [ Safra ]
                        b422LinhaHeader = new B422HeaderArqRemessa();
                        b422LinhaHeader.codigoEmpresa.valor = _boletoCedenteSelecionado.codigo_empresa;
                        b422LinhaHeader.nomeEmpresa.valor = _boletoCedenteSelecionado.nome_empresa;
                        b422LinhaHeader.numeroBanco.valor = _boletoCedenteSelecionado.num_banco;
                        b422LinhaHeader.nomeBanco.valor = "SAFRA";
                        b422LinhaHeader.dataGravacaoArquivo.valor = Global.digitos(Global.formataDataDdMmYyComSeparador(DateTime.Now));
                        b422LinhaHeader.numSequencialArquivo.valor = intNumSequencialRemessa.ToString();
                        b422LinhaHeader.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                        sw.WriteLine(Global.filtraAcentuacao(b422LinhaHeader.ToString()));
                        #endregion
                    }
                    else
                    {
                        #region [ Bradesco ]
                        b237LinhaHeader = new B237HeaderArqRemessa();
                        b237LinhaHeader.codigoEmpresa.valor = _boletoCedenteSelecionado.codigo_empresa;
                        b237LinhaHeader.nomeEmpresa.valor = _boletoCedenteSelecionado.nome_empresa;
                        b237LinhaHeader.numeroBanco.valor = _boletoCedenteSelecionado.num_banco;
                        b237LinhaHeader.nomeBanco.valor = _boletoCedenteSelecionado.nome_banco.ToUpper();
                        b237LinhaHeader.dataGravacaoArquivo.valor = Global.digitos(Global.formataDataDdMmYyComSeparador(DateTime.Now));
                        b237LinhaHeader.numSequencialRemessa.valor = intNumSequencialRemessa.ToString();
                        b237LinhaHeader.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                        sw.WriteLine(Global.filtraAcentuacao(b237LinhaHeader.ToString()));
                        #endregion
                    }
                    #endregion

                    #region [ Monta os registros do arquivo de remessa ]
                    foreach (DsDataSource.DtbFinBoletoRow rowBoleto in _dsConsulta.Tables["DtbFinBoleto"].Rows)
                    {
                        intTotalSerieBoletos++;
                        foreach (DsDataSource.DtbFinBoletoItemRow rowBoletoItem in rowBoleto.GetChildRows("DtbFinBoleto_DtbFinBoletoItem"))
                        {
                            intTotalParcelas++;
                            vlTotal += rowBoletoItem.valor;

                            #region [ Registro do tipo 1 ]
                            intNumSequencialRegistro++;

                            if (_boletoCedenteSelecionado.num_banco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
                            {
                                #region [ Safra ]
                                b422LinhaTipo1 = new B422RegTipo1ArqRemessa();
                                b422LinhaTipo1.tipoInscricaoEmpresa.valor = Global.Cte.FIN.BoletoSafra.TipoInscricaoEmpresa.CNPJ;
                                b422LinhaTipo1.numInscricao.valor = Global.digitos(_boletoCedenteSelecionado.cnpj);
                                b422LinhaTipo1.codEmpresa.valor = _boletoCedenteSelecionado.codigo_empresa;
                                b422LinhaTipo1.numControleParticipante.valor = rowBoletoItem.num_controle_participante;
                                // Código IOF Operações de Seguro: 0=Isento; 1=2%; 2=4%
                                b422LinhaTipo1.codIOF.valor = "0";
                                if ((rowBoletoItem.st_instrucao_protesto == 1) && (rowBoleto.primeira_instrucao.Equals("06")))
                                {
                                    // Como o sistema iniciou com o Boleto Bradesco, os campos do BD contém: primeira_instrucao = '06' (protestar); segunda_instrucao = nº dias p/ protesto
                                    // No Safra, os campos para informar o protesto são: instrucao2 = '10' (protestar); instrucao3 = nº dias p/ protesto
                                    b422LinhaTipo1.instrucao2.valor = "10";
                                    b422LinhaTipo1.instrucao3.valor = rowBoleto.segunda_instrucao;
                                }
                                else
                                {
                                    b422LinhaTipo1.instrucao2.valor = "00";
                                    b422LinhaTipo1.instrucao3.valor = "00";
                                }

                                b422LinhaTipo1.codCarteira.valor = Global.Cte.FIN.BoletoSafra.TipoCarteira.COBRANCA_SIMPLES;
                                b422LinhaTipo1.numDocumento.valor = rowBoletoItem.numero_documento.ToUpper();
                                b422LinhaTipo1.dataVenctoTitulo.valor = Global.digitos(Global.formataDataDdMmYyComSeparador(rowBoletoItem.dt_vencto));
                                b422LinhaTipo1.valorTitulo.valor = Global.digitos(Global.formataMoeda(rowBoletoItem.valor));
                                b422LinhaTipo1.bancoEncarregadoCobranca.valor = _boletoCedenteSelecionado.num_banco;
                                b422LinhaTipo1.agenciaDepositaria.valor = _boletoCedenteSelecionado.agencia;
                                b422LinhaTipo1.dataEmissaoTitulo.valor = Global.digitos(Global.formataDataDdMmYyComSeparador(DateTime.Now));

                                // Códigos para 1ª instrução de cobrança:
                                //    01 - NÃO RECEBER PRINCIPAL, SEM JUROS DE MORA
                                //    02 - DEVOLVER, SE NÃO PAGO, ATÉ 15 DIAS APÓS O VENCIMENTO
                                //    03 - DEVOLVER, SE NÃO PAGO, ATÉ 30 DIAS APÓS O VENCIMENTO
                                //    07 - NÃO PROTESTAR
                                //    08 - NÃO COBRAR JUROS DE MORA
                                //    16 - MULTA (*)
                                //      (*) Para tratamento de multa, formatar no campo “abatimento” (pos. 206 a 218), as seguintes informações:
                                //      Posição 206 a 211 a data a partir da qual a multa deve ser cobrada (ddmmaa)
                                //      Posição 212 a 215 o percentual referente à multa no formato 99v99.
                                //      Posição 216 a 218 zeros
                                b422LinhaTipo1.instrucao1.valor = "";
                                b422LinhaTipo1.valorPorDiaAtraso.valor = Global.digitos(Global.formataMoeda(rowBoletoItem.valor_por_dia_atraso));
                                b422LinhaTipo1.dataLimiteConcessaoDesconto.valor = "000000";
                                b422LinhaTipo1.valorDesconto.valor = "0";
                                b422LinhaTipo1.identificacaoTipoInscricaoSacado.valor = rowBoleto.tipo_sacado;
                                b422LinhaTipo1.numInscricaoSacado.valor = rowBoleto.num_inscricao_sacado;
                                b422LinhaTipo1.nomeSacado.valor = rowBoleto.nome_sacado.ToUpper();
                                b422LinhaTipo1.enderecoCompleto.valor = rowBoleto.endereco_sacado.ToUpper();
                                b422LinhaTipo1.enderecoBairro.valor = Texto.leftStr(rowBoleto.bairro_sacado, 10);
                                b422LinhaTipo1.cep.valor = Texto.leftStr(Global.digitos(rowBoleto.cep_sacado), 5);
                                if (Global.digitos(rowBoleto.cep_sacado).Length == 8)
                                    b422LinhaTipo1.sufixoCep.valor = Texto.rightStr(Global.digitos(rowBoleto.cep_sacado), 3);
                                else
                                    b422LinhaTipo1.sufixoCep.valor = "000";
                                b422LinhaTipo1.enderecoCidade.valor = Texto.leftStr(rowBoleto.cidade_sacado, 15);
                                b422LinhaTipo1.enderecoUF.valor = rowBoleto.uf_sacado;
                                b422LinhaTipo1.nomeSacadorAvalista.valor = rowBoletoItem.primeira_mensagem.ToUpper();
                                b422LinhaTipo1.bancoEmitenteBoleto.valor = _boletoCedenteSelecionado.num_banco;
                                b422LinhaTipo1.numSequencialArquivo.valor = intNumSequencialRemessa.ToString();
                                b422LinhaTipo1.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                                sw.WriteLine(Global.filtraAcentuacao(b422LinhaTipo1.ToString()));
                                #endregion
                            }
                            else
                            {
                                #region [ Bradesco ]
                                b237LinhaTipo1 = new B237RegTipo1ArqRemessa();
                                b237LinhaTipo1.identifCedenteCarteira.valor = rowBoleto.carteira;
                                b237LinhaTipo1.identifCedenteAgencia.valor = rowBoleto.agencia;
                                b237LinhaTipo1.identifCedenteCtaCorrente.valor = rowBoleto.conta;
                                b237LinhaTipo1.identifCedenteDigitoCtaCorrente.valor = rowBoleto.digito_conta;
                                b237LinhaTipo1.numControleParticipante.valor = rowBoletoItem.num_controle_participante;
                                if (rowBoleto.perc_multa > 0)
                                {
                                    b237LinhaTipo1.campoMulta.valor = "2";
                                    b237LinhaTipo1.percentualMulta.valor = Global.digitos(Global.formataPercentualCom2Decimais(rowBoleto.perc_multa));
                                }
                                if (rowBoletoItem.bonificacao_por_dia > 0)
                                {
                                    b237LinhaTipo1.descontoBonificacaoPorDia.valor = Global.digitos(Global.formataMoeda(rowBoletoItem.bonificacao_por_dia));
                                }
                                b237LinhaTipo1.numDocumento.valor = rowBoletoItem.numero_documento.ToUpper();
                                b237LinhaTipo1.dataVenctoTitulo.valor = Global.digitos(Global.formataDataDdMmYyComSeparador(rowBoletoItem.dt_vencto));
                                b237LinhaTipo1.valorTitulo.valor = Global.digitos(Global.formataMoeda(rowBoletoItem.valor));
                                b237LinhaTipo1.dataEmissaoTitulo.valor = Global.digitos(Global.formataDataDdMmYyComSeparador(DateTime.Now));
                                if (rowBoletoItem.st_instrucao_protesto == 1)
                                {
                                    b237LinhaTipo1.primeiraInstrucao.valor = rowBoleto.primeira_instrucao;
                                    b237LinhaTipo1.segundaInstrucao.valor = rowBoleto.segunda_instrucao;
                                }
                                else
                                {
                                    b237LinhaTipo1.primeiraInstrucao.valor = "00";
                                    b237LinhaTipo1.segundaInstrucao.valor = "00";
                                }
                                b237LinhaTipo1.valorPorDiaAtraso.valor = Global.digitos(Global.formataMoeda(rowBoletoItem.valor_por_dia_atraso));
                                b237LinhaTipo1.dataLimiteConcessaoDesconto.valor = "000000";
                                b237LinhaTipo1.valorDesconto.valor = "0";
                                b237LinhaTipo1.identificacaoTipoInscricaoSacado.valor = rowBoleto.tipo_sacado;
                                b237LinhaTipo1.numInscricaoSacado.valor = rowBoleto.num_inscricao_sacado;
                                b237LinhaTipo1.nomeSacado.valor = rowBoleto.nome_sacado.ToUpper();
                                b237LinhaTipo1.enderecoCompleto.valor = rowBoleto.endereco_sacado.ToUpper();
                                b237LinhaTipo1.primeiraMensagem.valor = rowBoletoItem.primeira_mensagem.ToUpper();
                                b237LinhaTipo1.cep.valor = Texto.leftStr(Global.digitos(rowBoleto.cep_sacado), 5);
                                if (Global.digitos(rowBoleto.cep_sacado).Length == 8)
                                    b237LinhaTipo1.sufixoCep.valor = Texto.rightStr(Global.digitos(rowBoleto.cep_sacado), 3);
                                else
                                    b237LinhaTipo1.sufixoCep.valor = "000";
                                b237LinhaTipo1.sacadorAvalistaOuSegundaMensagem.valor = rowBoleto.segunda_mensagem.ToUpper();
                                b237LinhaTipo1.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                                sw.WriteLine(Global.filtraAcentuacao(b237LinhaTipo1.ToString()));
                                #endregion
                            }
                            #endregion

                            #region [ Registro do tipo 2 ]
                            intNumSequencialRegistro++;

                            if (_boletoCedenteSelecionado.num_banco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
                            {
                                #region [ Safra ]
                                // TODO
                                #endregion
                            }
                            else
                            {
                                #region [ Bradesco ]
                                b237LinhaTipo2 = new B237RegTipo2ArqRemessa();
                                b237LinhaTipo2.mensagem_1.valor = rowBoleto.mensagem_1;
                                b237LinhaTipo2.mensagem_2.valor = rowBoleto.mensagem_2;
                                b237LinhaTipo2.mensagem_3.valor = rowBoleto.mensagem_3;
                                b237LinhaTipo2.mensagem_4.valor = rowBoleto.mensagem_4;
                                b237LinhaTipo2.carteira.valor = rowBoleto.carteira;
                                b237LinhaTipo2.agencia.valor = rowBoleto.agencia;
                                b237LinhaTipo2.contaCorrente.valor = rowBoleto.conta;
                                b237LinhaTipo2.digitoContaCorrente.valor = rowBoleto.digito_conta;
                                b237LinhaTipo2.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                                sw.WriteLine(Global.filtraAcentuacao(b237LinhaTipo2.ToString()));
                                #endregion
                            }
                            #endregion
                        }
                    }
                    #endregion

                    #region [ Monta Trailler ]
                    intNumSequencialRegistro++;

                    if (_boletoCedenteSelecionado.num_banco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
                    {
                        #region [ Safra ]
                        b422LinhaTrailler = new B422TraillerArqRemessa();
                        b422LinhaTrailler.qtdeTitulos.valor = intTotalParcelas.ToString();
                        b422LinhaTrailler.valorTotalTitulos.valor = Global.digitos(Global.formataMoeda(vlTotal));
                        b422LinhaTrailler.numSequencialArquivo.valor = intNumSequencialRemessa.ToString();
                        b422LinhaTrailler.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                        sw.WriteLine(Global.filtraAcentuacao(b422LinhaTrailler.ToString()));
                        #endregion
                    }
                    else
                    {
                        #region [ Bradesco ]
                        b237LinhaTrailler = new B237TraillerArqRemessa();
                        b237LinhaTrailler.numSequencialRegistro.valor = intNumSequencialRegistro.ToString();
                        sw.WriteLine(Global.filtraAcentuacao(b237LinhaTrailler.ToString()));
                        #endregion
                    }
                    #endregion
                }
                finally
                {
                    sw.Flush();
                    sw.Close();
                }

				#region [ Dados para o histórico de arquivos de remessa ]
				boletoArqRemessa.qtde_serie_boletos = intTotalSerieBoletos;
				boletoArqRemessa.qtde_registros = intTotalParcelas;
				boletoArqRemessa.vl_total = vlTotal;
				#endregion

				#region [ Assinala os boletos como gravados no arquivo de remessa ]
				blnSucesso = false;
				try
				{
					BD.iniciaTransacao();

					foreach (DsDataSource.DtbFinBoletoRow rowBoleto in _dsConsulta.Tables["DtbFinBoleto"].Rows)
					{
						if (!BoletoDAO.marcaBoletoEnviadoRemessaBanco(Global.Usuario.usuario,
																 rowBoleto.id,
																 id_boleto_arq_remessa,
																 ref strMsgErro))
						{
							throw new Exception("Falha ao marcar o registro id=" + rowBoleto.id.ToString() + " do boleto como já gravado no arquivo de remessa!!\n" + strMsgErro);
						}

						foreach (DsDataSource.DtbFinBoletoItemRow rowBoletoItem in rowBoleto.GetChildRows("DtbFinBoleto_DtbFinBoletoItem"))
						{
							if (!BoletoDAO.marcaBoletoItemEnviadoRemessaBanco(Global.Usuario.usuario,
																		 rowBoletoItem.id,
																		 ref strMsgErro))
							{
								throw new Exception("Falha ao marcar o registro id=" + rowBoletoItem.id.ToString() + " da parcela do boleto como já gravado no arquivo de remessa!!\n" + strMsgErro);
							}
						}
					}

					#region [ Grava o registro em t_FIN_BOLETO_ARQ_REMESSA ]
					boletoArqRemessa.st_geracao = Global.Cte.FIN.CodBoletoArqRemessaStGeracao.SUCESSO;
					boletoArqRemessa.duracao_proc_em_seg = Global.calculaTimeSpanSegundos(DateTime.Now - dtInicioProcessamento);
					if (!BoletoDAO.boletoArqRemessaInsere(Global.Usuario.usuario, boletoArqRemessa, ref strMsgErroAux))
					{
						throw new Exception("Falha ao gravar o histórico de arquivos de remessas no banco de dados!!");
					}
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					strMsgErro = ex.ToString();
					blnSucesso = false;
				}
				#endregion

				if (blnSucesso)
				{
					BD.commitTransacao();

					#region [ Grava o log no BD ]
					strDescricaoLog = "Arquivo de remessa gerado: " + strNomeCompletoArqRemessa + ", nº sequencial de remessa=" + intNumSequencialRemessa.ToString() + ", contendo " + Global.formataInteiro(intTotalSerieBoletos) + " séries de boletos com um total de " + Global.formataInteiro(intTotalParcelas) + " parcelas que totalizam o valor de " + Global.formataMoeda(vlTotal);
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_GERA_ARQ_REMESSA;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					Global.Usuario.Defaults.FBoletoArqRemessa.pathBoletoArquivoRemessa = Global.barraInvertidaDel(strPathBase);

					System.Diagnostics.Process.Start(txtDiretorio.Text);
					info(ModoExibicaoMensagemRodape.Normal);
					aviso("Arquivo de remessa gerado com sucesso!!\n\n" + strNomeCompletoArqRemessa);
					Close();
				}
				else
				{
					BD.rollbackTransacao();

					#region [ Se o arquivo de remessa foi gravado, renomeia para indicar que houve uma falha ]
					if (File.Exists(strNomeCompletoArqRemessa)) File.Move(strNomeCompletoArqRemessa, strNomeCompletoArqRemessa + ".ERR");
					#endregion
					
					info(ModoExibicaoMensagemRodape.Normal);
					strMsgErro = "Falha ao marcar os registros dos boletos como já gravados no arquivo de remessa!!\n\n" + strMsgErro;
					avisoErro(strMsgErro);
				}
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoDesfazerBoleto ]
		private void trataBotaoDesfazerBoleto()
		{
			#region [ Declarações ]
			int intLinhaGridSelecionado = -1;
			int id_boleto_selecionado;
			bool blnSucesso;
			String strMsgErro = "";
			String strDescricaoLog = "";
			String strMsgErroLog = "";
			DsDataSource.DtbFinBoletoRow rowBoleto = null;
			DataRow[] vRowsSelect;
			FinLog finLog = new FinLog();
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

			#region [ Obtém índice no grid do boleto selecionado ]
			for (int i = 0; i < grdBoletos.Rows.Count; i++)
			{
				if (grdBoletos.Rows[i].Selected)
				{
					intLinhaGridSelecionado = i;
					break;
				}
			}

			if (intLinhaGridSelecionado < 0)
			{
				avisoErro("Nenhum boleto foi selecionado!!");
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma que o boleto selecionado deve ser desfeito, retornando ao estágio inicial?\n\n" + grdBoletos.Rows[intLinhaGridSelecionado].Cells["cliente"].Value)) return;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "desfazendo boleto");
			try
			{
				#region [ Obtém id do boleto selecionado ]
				id_boleto_selecionado = (int)Global.converteInteiro(grdBoletos.Rows[intLinhaGridSelecionado].Cells["id_boleto"].Value.ToString());
				if (id_boleto_selecionado == 0)
				{
					avisoErro("Falha ao obter o id do boleto a ser desfeito!!");
					return;
				}
				#endregion

				#region [ Obtém registro principal do boleto ]
				vRowsSelect = _dsConsulta.Tables["DtbFinBoleto"].Select("id=" + id_boleto_selecionado.ToString());
				if (vRowsSelect.Length != 1)
				{
					throw new Exception("Falha ao obter o registro referente ao boleto que será desfeito!!");
				}
				rowBoleto = (DsDataSource.DtbFinBoletoRow)vRowsSelect[0];
				#endregion

				#region [ Exclui o boleto e reverte o status do registro gerado na impressão da NF ]
				blnSucesso = false;
				try
				{
					#region [ Inicia a transação ]
					BD.iniciaTransacao();
					#endregion

					#region [ Exclui os dados do boleto ]
					if (!BoletoDAO.excluiBoletoEmStatusInicial(
												Global.Usuario.usuario,
												id_boleto_selecionado,
												ref strMsgErro))
					{
						throw new Exception("Falha ao excluir o boleto com registro id=" + id_boleto_selecionado.ToString() + "!!\n" + strMsgErro);
					}
					strDescricaoLog = "Boleto id=" + id_boleto_selecionado.ToString() + "; Cliente=" + Global.formataCnpjCpf(rowBoleto.num_inscricao_sacado) + " - " + rowBoleto.nome_sacado;
					#endregion

					#region [ Restaura status inicial dos dados gerados na emissão da NF (se não for um boleto avulso) ]
					if (rowBoleto.id_nf_parcela_pagto > 0)
					{
						if (!BoletoPreCadastradoDAO.restauraStatusInicial(
														Global.Usuario.usuario,
														rowBoleto.id_nf_parcela_pagto,
														ref strMsgErro))
						{
							throw new Exception("Falha ao restaurar status inicial dos dados gerados na emissão da NF (registro id=" + rowBoleto.id_nf_parcela_pagto.ToString() + ")!!\n" + strMsgErro);
						}
						strDescricaoLog += "; restaurado status inicial de t_FIN_NF_PARCELA_PAGTO.id=" + rowBoleto.id_nf_parcela_pagto.ToString();
					}
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					strMsgErro = ex.ToString();
					blnSucesso = false;
				}
				#endregion

				#region [ Finaliza a transação ]
				if (blnSucesso)
				{
					BD.commitTransacao();

					#region [ Grava o log no BD ]
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_DESFEITO;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.id_registro_origem = id_boleto_selecionado;
					finLog.id_cliente = rowBoleto.id_cliente;
					finLog.cnpj_cpf = rowBoleto.num_inscricao_sacado;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion
				}
				else
				{
					BD.rollbackTransacao();
					strMsgErro = "Falha ao desfazer o boleto!!\n\n" + strMsgErro;
					avisoErro(strMsgErro);
				}
				#endregion

				#region [ Atualiza o grid ]
				trataBotaoExecutaConsulta();
				#endregion
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoArqRemessa ]

		#region [ FBoletoArqRemessa_Load ]
		private void FBoletoArqRemessa_Load(object sender, EventArgs e)
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

		#region [ FBoletoArqRemessa_Shown ]
		private void FBoletoArqRemessa_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

					#region [ Preenchimento dos campos ]

					#region [ Combo Cedente ]
					cbBoletoCedente.ValueMember = "id";
					cbBoletoCedente.DisplayMember = "descricao_formatada";
					cbBoletoCedente.DataSource = ComboDAO.criaDtbBoletoCedenteCombo(ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS);
					if (Global.Usuario.Defaults.FBoletoArqRemessa.boletoCedente == 0)
						cbBoletoCedente.SelectedIndex = -1;
					else
						if (!comboBoletoCedentePosicionaDefault()) cbBoletoCedente.SelectedIndex = -1;
					// Se houver apenas 1 opção, então seleciona
					if ((cbBoletoCedente.Items.Count == 1) && (cbBoletoCedente.SelectedIndex == -1)) cbBoletoCedente.SelectedIndex = 0;
					#endregion

					txtDiretorio.Text = pathBoletoArquivoRemessaValorDefault();
					#endregion

					#region [ Faz a consulta automaticamente? ]
					if (cbBoletoCedente.SelectedIndex != -1) trataBotaoExecutaConsulta();
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

		#region [ FBoletoArqRemessa_KeyDown ]
		private void FBoletoArqRemessa_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				trataBotaoExecutaConsulta();
				return;
			}
		}
		#endregion

		#region [ FBoletoArqRemessa_FormClosing ]
		private void FBoletoArqRemessa_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
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

		#region [ btnCancela ]

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			trataBotaoCancela();
		}
		#endregion

		#endregion

		#region [ btnDesfazerBoleto ]

		#region [ btnDesfazerBoleto_Click ]
		private void btnDesfazerBoleto_Click(object sender, EventArgs e)
		{
			trataBotaoDesfazerBoleto();
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

		#region [ cbBoletoCedente ]

		#region [ cbBoletoCedente_SelectionChangeCommitted ]
		private void cbBoletoCedente_SelectionChangeCommitted(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strPathBase = "";
			String strPathCompleto = "";
			String strMsgErro = "";
			#endregion

			if (!_InicializacaoOk) return;
			trataBotaoExecutaConsulta();

			#region [ Sincroniza campo contendo o nome do diretório ]
			if (montaPathBoletoArquivoRemessaBoletoCedente(_boletoCedenteSelecionado, txtDiretorio.Text, ref strPathBase, ref strPathCompleto, ref strMsgErro))
			{
				if (Directory.Exists(strPathCompleto))
				{
					txtDiretorio.Text = Global.barraInvertidaDel(strPathCompleto);
				}
				else
				{
					txtDiretorio.Text = Global.barraInvertidaDel(strPathBase);
				}
			}
			#endregion
		}
		#endregion

		#endregion

		#endregion
	}
}
