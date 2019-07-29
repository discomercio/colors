#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Media;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.IO;
#endregion

namespace EtqWms
{
	#region [ FEtiquetaImprime ]
	public partial class FEtiquetaImprime : EtqWms.FModelo
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
		const String GRID_PESQ_COL_DATA_HORA = "colGrdPesqDataHora";
		const String GRID_PESQ_COL_NSU = "colGrdPesqNsu";
		const String GRID_PESQ_COL_QTDE_PEDIDOS = "colGrdPesqQtdePedidos";
		const String GRID_PESQ_COL_QTDE_VOLUMES = "colGrdPesqQtdeVolumes";
		const String GRID_PESQ_COL_USUARIO = "colGrdPesqUsuario";
		const String GRID_DADOS_COL_NSU_N1 = "colGrdDadosNsuN1";
		const String GRID_DADOS_COL_NSU_N2 = "colGrdDadosNsuN2";
		const String GRID_DADOS_COL_NSU_N3 = "colGrdDadosNsuN3";
		const String GRID_DADOS_COL_CHECKBOX = "colGrdDadosCheckBox";
		const String GRID_DADOS_COL_SEQUENCIA = "colGrdDadosSequencia";
		const String GRID_DADOS_COL_PEDIDO = "colGrdDadosPedido";
		const String GRID_DADOS_COL_NUMERO_NF = "colGrdDadosNumeroNF";
		const String GRID_DADOS_COL_TRANSPORTADORA = "colGrdDadosTransportadora";
		const String GRID_DADOS_COL_CLIENTE = "colGrdDadosCliente";
		const String GRID_DADOS_COL_ZONA = "colGrdDadosZona";
		const String GRID_DADOS_COL_QTDE = "colGrdDadosQtde";
		const String GRID_DADOS_COL_QTDE_VOLUMES = "colGrdDadosQtdeVolumes";
		const String GRID_DADOS_COL_PRODUTO = "colGrdDadosProduto";
		#endregion

		#region [ Atributos ]
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
		private int _idNsuUltimoSelecionado = 0;
		private DateTime _dtHrNsuUltimoSelecionado = DateTime.MinValue;
		private List<EtiquetaDados> _listaEtqCompleta = new List<EtiquetaDados>();
		private List<EtiquetaDados> _listaEtqParcial = new List<EtiquetaDados>();
		private List<EtiquetaCtrlNumeracaoVolumes> _listaCtrlNumeracaoVolumes = new List<EtiquetaCtrlNumeracaoVolumes>();
		#endregion

		#region [ Impressão ]
		const char CODIGO_SOH = (char)0x01;
		const char CODIGO_STX = (char)0x02;
		const char CODIGO_CR = (char)0x0D;
		#endregion

		#region [ Construtor ]
		public FEtiquetaImprime()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ calculaTotalVolumes ]
		private bool calculaTotalVolumes(List<EtiquetaDados> lista, out int totalVolumes, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "calculaTotalVolumes()";
			int intTotalizador = 0;
			#endregion

			totalVolumes = 0;
			strMsgErro = "";
			try
			{
				for (int i = 0; i < lista.Count; i++)
				{
					intTotalizador += lista[i].qtde * lista[i].qtde_volumes;
				}

				totalVolumes = intTotalizador;
				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ executaConsulta ]
		private bool executaConsulta()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaConsulta()";
			int idNsuSelecionado = 0;
			int intQtdePedidos = 0;
			int intQtdeTotalVolumes = 0;
			int intSequencia = 0;
			bool blnAchou;
			bool blnInconsistenciaEmits = false;
			String strMsg;
			String strSql;
			String strPedido;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataRow rowConsulta;
			StringBuilder sbPedidos = new StringBuilder("");
			EtiquetaDados etiqueta;
			EtiquetaCtrlNumeracaoVolumes etiquetaCtrlNumeracaoVolumes = null;
			CtrlNumeracaoVolume ctrlNumeracaoVolume;
			#endregion

			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				#region [ Há linha selecionada? ]
				for (int i = 0; i < grdPesquisa.Rows.Count; i++)
				{
					if (grdPesquisa.Rows[i].Selected)
					{
						idNsuSelecionado = (int)Global.converteInteiro(grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_NSU].Value.ToString());
						break;
					}
				}

				if (idNsuSelecionado == 0)
				{
					avisoErro("Nenhum relatório foi selecionado!!");
					return false;
				}
				#endregion

				#region [ É a mesma consulta que a anterior? ]
				if ((idNsuSelecionado == _idNsuUltimoSelecionado) && (grdDados.Rows.Count > 0))
				{
					// Evita cliques acidentais, mas após algum tempo permite recarregar os dados
					if (Global.calculaTimeSpanSegundos(DateTime.Now - _dtHrNsuUltimoSelecionado) <= 10)
					{
						return true;
					}
				}
				#endregion

				#region [ Exibe o NSU selecionado ]
				lblNsuDadosRelatorio.Text = Global.formataInteiro(idNsuSelecionado);
				#endregion

				#region [ Limpa listas que armazenam os dados p/ impressão das etiquetas ]
				_listaEtqCompleta.Clear();
				_listaEtqParcial.Clear();
				_listaCtrlNumeracaoVolumes.Clear();
				#endregion

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				#endregion

				#region [ Monta o SQL ]
				// IMPORTANTE: as etiquetas devem estar ordenadas por Zona e Produto.
				//		A) Zona: para poder separar as etiquetas facilmente p/ cada "separador" no depósito.
				//		B) Produto: para o "separador" poder separar e colar as etiquetas em todas as unidades
				//				 de uma só vez.

				strSql = "SELECT" +
							" tN1.id AS id_N1," +
							" tN2.id AS id_N2," +
							" tN3.id AS id_N3," +
							" tN2.pedido," +
							" tN2.obs_2," +
							" tN2.obs_3," +
							" tN2.loja," +
							" tN2.transportadora_id," +
							" tN2.id_cliente," +
							" tCli.cnpj_cpf AS cnpj_cpf_cliente," +
							" tCli.nome_iniciais_em_maiusculas AS nome_cliente," +
							" tFab.nome AS nome_fabricante," +
							" tFab.razao_social AS razao_social_fabricante," +
							" tN3.fabricante," +
							" tN3.produto," +
							" tN3.qtde," +
							" tN3.qtde_volumes," +
							" tProd.descricao," +
							" tProd.descricao_html," +
							" tN3.zona_id," +
							" tN3.zona_codigo," +
							" tN2.numeroNFe," +
							" tN2.qtde_volumes_pedido," +
							" tN2.destino_tipo_endereco," +
							" tN2.destino_endereco," +
							" tN2.destino_endereco_numero," +
							" tN2.destino_endereco_complemento," +
							" tN2.destino_bairro," +
							" tN2.destino_cidade," +
							" tN2.destino_uf," +
							" tN2.destino_cep," +
							" tPed.st_entrega," +
							" tCli.uf AS origem_uf," +
							" tPed.id_nfe_emitente" +
						" FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO tN1" +
							" INNER JOIN t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO tN2 ON (tN1.id = tN2.id_wms_etq_n1)" +
							" INNER JOIN t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO tN3 ON (tN2.id = tN3.id_wms_etq_n2)" +
							" INNER JOIN t_CLIENTE tCli ON (tN2.id_cliente = tCli.id)" +
							" INNER JOIN t_FABRICANTE tFab ON (tN3.fabricante = tFab.fabricante)" +
							" INNER JOIN t_PRODUTO tProd ON ((tN3.fabricante = tProd.fabricante) AND (tN3.produto = tProd.produto))" +
							" INNER JOIN t_PEDIDO tPed ON (tN2.pedido = tPed.pedido)" +
						" WHERE" +
							" (tN1.id = " + idNsuSelecionado.ToString() + ")" +
						" ORDER BY" +
							" tN3.zona_codigo," +
							" tN3.fabricante," +
							" tN3.produto," +
							" tN1.id," +
							" tN2.id," +
							" tN3.id";
				#endregion

				#region [ Executa a consulta no BD ]
				_dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(_dtbConsulta);
				#endregion

				#region [ Consiste Emitentes dos Pedidos ]
				grdDados.SuspendLayout();
				grdDados.Rows.Clear();
				strMsg = "";
				for (int iEmit = 0; (iEmit < _dtbConsulta.Rows.Count) && (!blnInconsistenciaEmits); iEmit++)
				{
					rowConsulta = _dtbConsulta.Rows[iEmit];
					if ((BD.readToString(rowConsulta["id_nfe_emitente"]) != Global.Usuario.emit_id))
					{
						strMsg = "Inconsistência no Emitente do pedido " + BD.readToString(rowConsulta["pedido"]) + "!!";
						blnInconsistenciaEmits = true;
					}
				}
				if (blnInconsistenciaEmits)
				{
					avisoErro(strMsg);
					return false;
				}
				#endregion

				#region [ Carrega dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					//grdDados.SuspendLayout();

					//grdDados.Rows.Clear();
					if (_dtbConsulta.Rows.Count > 0) grdDados.Rows.Add(_dtbConsulta.Rows.Count);

					for (int i = 0; i < _dtbConsulta.Rows.Count; i++)
					{
						intSequencia++;
						rowConsulta = _dtbConsulta.Rows[i];
						strPedido = "|" + BD.readToString(rowConsulta["pedido"]) + "|";
						if (!sbPedidos.ToString().Contains(strPedido))
						{
							sbPedidos.Append(strPedido);
							intQtdePedidos++;
						}
						etiqueta = new EtiquetaDados();
						etiqueta.id_N1 = BD.readToInt(rowConsulta["id_N1"]);
						etiqueta.id_N2 = BD.readToInt(rowConsulta["id_N2"]);
						etiqueta.id_N3 = BD.readToInt(rowConsulta["id_N3"]);
						etiqueta.pedido = BD.readToString(rowConsulta["pedido"]);
						etiqueta.obs_2 = BD.readToString(rowConsulta["obs_2"]);
						etiqueta.obs_3 = BD.readToString(rowConsulta["obs_3"]);
						etiqueta.loja = BD.readToString(rowConsulta["loja"]);
						etiqueta.transportadora_id = BD.readToString(rowConsulta["transportadora_id"]);
						etiqueta.id_cliente = BD.readToString(rowConsulta["id_cliente"]);
						etiqueta.cnpj_cpf_cliente = BD.readToString(rowConsulta["cnpj_cpf_cliente"]);
						etiqueta.nome_cliente = BD.readToString(rowConsulta["nome_cliente"]);
						etiqueta.nome_fabricante = BD.readToString(rowConsulta["nome_fabricante"]);
						etiqueta.razao_social_fabricante = BD.readToString(rowConsulta["razao_social_fabricante"]);
						etiqueta.fabricante = BD.readToString(rowConsulta["fabricante"]);
						etiqueta.produto = BD.readToString(rowConsulta["produto"]);
						etiqueta.qtde = BD.readToInt(rowConsulta["qtde"]);
						etiqueta.qtde_volumes = BD.readToInt(rowConsulta["qtde_volumes"]);
						etiqueta.descricao = BD.readToString(rowConsulta["descricao"]);
						etiqueta.descricao_html = BD.readToString(rowConsulta["descricao_html"]);
						etiqueta.zona_id = BD.readToInt(rowConsulta["zona_id"]);
						etiqueta.zona_codigo = BD.readToString(rowConsulta["zona_codigo"]);
						etiqueta.numeroNFe = BD.readToString(rowConsulta["numeroNFe"]);
						etiqueta.qtde_volumes_pedido = BD.readToInt(rowConsulta["qtde_volumes_pedido"]);
						etiqueta.st_entrega = BD.readToString(rowConsulta["st_entrega"]);
						etiqueta.destino_tipo_endereco = BD.readToString(rowConsulta["destino_tipo_endereco"]);
						etiqueta.destino_endereco = BD.readToString(rowConsulta["destino_endereco"]);
						etiqueta.destino_endereco_numero = BD.readToString(rowConsulta["destino_endereco_numero"]);
						etiqueta.destino_endereco_complemento = BD.readToString(rowConsulta["destino_endereco_complemento"]);
						etiqueta.destino_bairro = BD.readToString(rowConsulta["destino_bairro"]);
						etiqueta.destino_cidade = BD.readToString(rowConsulta["destino_cidade"]);
						etiqueta.destino_uf = BD.readToString(rowConsulta["destino_uf"]);
						etiqueta.destino_cep = BD.readToString(rowConsulta["destino_cep"]);
						etiqueta.origem_uf = BD.readToString(rowConsulta["origem_uf"]);
						etiqueta.sequencia = intSequencia;

						#region [ Numeração dos volumes ]
						blnAchou = false;
						for (int j = 0; j < _listaCtrlNumeracaoVolumes.Count; j++)
						{
							if (_listaCtrlNumeracaoVolumes[j].pedido.Equals(etiqueta.pedido))
							{
								blnAchou = true;
								etiquetaCtrlNumeracaoVolumes = _listaCtrlNumeracaoVolumes[j];
								break;
							}
						}

						if (!blnAchou)
						{
							etiquetaCtrlNumeracaoVolumes = new EtiquetaCtrlNumeracaoVolumes();
							etiquetaCtrlNumeracaoVolumes.pedido = etiqueta.pedido;
							etiquetaCtrlNumeracaoVolumes.qtde_volumes_pedido = etiqueta.qtde_volumes_pedido;
							_listaCtrlNumeracaoVolumes.Add(etiquetaCtrlNumeracaoVolumes);
						}

						for (int j = 0; j < (etiqueta.qtde); j++)
						{
							for (int k = 0; k < etiqueta.qtde_volumes; k++)
							{
								etiquetaCtrlNumeracaoVolumes.contador_numeracao_volume_pedido++;
								if (etiquetaCtrlNumeracaoVolumes.contador_numeracao_volume_pedido > etiquetaCtrlNumeracaoVolumes.qtde_volumes_pedido)
								{
									throw new Exception("Falha ao atribuir numeração para os volumes do pedido " + etiquetaCtrlNumeracaoVolumes.pedido + ": a numeração (" + etiquetaCtrlNumeracaoVolumes.contador_numeracao_volume_pedido.ToString() + ") excedeu o total de volumes (" + etiquetaCtrlNumeracaoVolumes.qtde_volumes_pedido.ToString() + ")!!");
								}
								ctrlNumeracaoVolume = new CtrlNumeracaoVolume();
								ctrlNumeracaoVolume.numeracaoVolume = etiquetaCtrlNumeracaoVolumes.contador_numeracao_volume_pedido;
								if (etiqueta.qtde_volumes > 1)
								{
									if (k == 0)
									{
										ctrlNumeracaoVolume.identificacaoVolumeProdutoComposto = "I";
									}
									else if (k == 1)
									{
										ctrlNumeracaoVolume.identificacaoVolumeProdutoComposto = "O";
									}
								}
								etiqueta.ctrlNumeracaoVolume.Add(ctrlNumeracaoVolume);
							}
						}
						#endregion

						_listaEtqCompleta.Add(etiqueta);
						_listaEtqParcial.Add(etiqueta);

						preencheLinhaGrid(i, etiqueta);
						intQtdeTotalVolumes += etiqueta.qtde * etiqueta.qtde_volumes;
					}

					#region [ Confere a numeração dos volumes ]
					strMsg = "";
					for (int i = 0; i < _listaCtrlNumeracaoVolumes.Count; i++)
					{
						if (_listaCtrlNumeracaoVolumes[i].contador_numeracao_volume_pedido != _listaCtrlNumeracaoVolumes[i].qtde_volumes_pedido)
						{
							if (strMsg.Length > 0) strMsg += "\n";
							strMsg += "Pedido " + _listaCtrlNumeracaoVolumes[i].pedido + ": o contador de volumes (" + _listaCtrlNumeracaoVolumes[i].contador_numeracao_volume_pedido + ") não coincidiu com a quantidade total de volumes do pedido (" + _listaCtrlNumeracaoVolumes[i].qtde_volumes_pedido + ")!!";
						}
					}

					if (strMsg.Length > 0)
					{
						throw new Exception(strMsg);
					}
					#endregion

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdDados.Rows.Count; i++)
					{
						if (grdDados.Rows[i].Selected) grdDados.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					grdDados.ResumeLayout();
				}
				#endregion

				#region [ Totais ]
				lblTotalRegistros.Text = Global.formataInteiro(_dtbConsulta.Rows.Count);
				lblTotalPedidos.Text = Global.formataInteiro(intQtdePedidos);
				lblTotalVolumes.Text = Global.formataInteiro(intQtdeTotalVolumes);
				#endregion

				#region [ Resultado: Completo ou Parcial? ]
				lblDadosRelatorioCompletoOuParcial.Text = "COMPLETO";
				btnImprimirCompleto.Enabled = true;
				btnImprimirParcial.Enabled = false;
				btnImprimirVolume.Enabled = true;
				limpaCamposFiltro();
				#endregion

				grdDados.Focus();

				// Feedback da conclusão da consulta
				SystemSounds.Exclamation.Play();

				_dtHrNsuUltimoSelecionado = DateTime.Now;
				_idNsuUltimoSelecionado = idNsuSelecionado;

				return true;
			}
			catch (Exception ex)
			{
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

		#region [ executaFiltragemDados ]
		private bool executaFiltragemDados(eFiltroPreenchimentoObrigatorio opcaoFiltro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaFiltragemDados()";
			bool blnTemFiltro = false;
			bool blnNumeroNF = false;
			bool blnPedido = false;
			bool blnSequenciaInicio = false;
			bool blnSequenciaFim = false;
			bool blnZona = false;
			int intNumeroNF = 0;
			int intSequencia = 0;
			int intQtdePedidos = 0;
			int intQtdeTotalVolumes = 0;
			string strPedido = "";
			int intSequenciaInicio = 0;
			int intSequenciaFim = 0;
			string strZona = "";
			EtiquetaDados etiqueta;
			StringBuilder sbPedidos = new StringBuilder("");
			#endregion

			try
			{
				#region [ Verifica se algum filtro foi especificado ]
				if (txtNumeroNF.Text.Trim().Length > 0)
				{
					blnTemFiltro = true;
					blnNumeroNF = true;
					intNumeroNF = (int)Global.converteInteiro(txtNumeroNF.Text);
				}
				if (txtPedido.Text.Trim().Length > 0)
				{
					blnTemFiltro = true;
					blnPedido = true;
					strPedido = Global.normalizaNumeroPedido(txtPedido.Text.Trim().ToUpper());
				}
				if (txtSequenciaInicio.Text.Trim().Length > 0)
				{
					blnTemFiltro = true;
					blnSequenciaInicio = true;
					intSequenciaInicio = (int)Global.converteInteiro(txtSequenciaInicio.Text);
				}
				if (txtSequenciaFim.Text.Trim().Length > 0)
				{
					blnTemFiltro = true;
					blnSequenciaFim = true;
					intSequenciaFim = (int)Global.converteInteiro(txtSequenciaFim.Text);
				}
				if (txtZona.Text.Trim().Length > 0)
				{
					blnTemFiltro = true;
					blnZona = true;
					strZona = txtZona.Text.Trim().ToUpper();
				}

				if (!blnTemFiltro)
				{
					if (opcaoFiltro == eFiltroPreenchimentoObrigatorio.OBRIGATORIO)
					{
						avisoErro("Nenhum filtro foi informado!!");
						return false;
					}
				}
				#endregion

				#region [ Intervalo de sequência está coerente? ]
				if ((txtSequenciaInicio.Text.Trim().Length > 0) && (txtSequenciaFim.Text.Trim().Length > 0))
				{
					if (Global.converteInteiro(txtSequenciaFim.Text) < Global.converteInteiro(txtSequenciaInicio.Text))
					{
						avisoErro("Valor final da faixa da sequência é menor que o valor inicial!!");
						txtSequenciaFim.Focus();
						return false;
					}
				}
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				_listaEtqParcial.Clear();

				for (int i = 0; i < _listaEtqCompleta.Count; i++)
				{
					etiqueta = _listaEtqCompleta[i];

					if (blnNumeroNF)
					{
						// O número armazenado no campo 'obs_3' tem prioridade
						if (etiqueta.obs_3.Length > 0)
						{
							if (intNumeroNF != (int)Global.converteInteiro(etiqueta.obs_3)) continue;
						}
						else
						{
							if (intNumeroNF != (int)Global.converteInteiro(etiqueta.obs_2)) continue;
						}
					}

					if (blnPedido)
					{
						if (!strPedido.Equals(etiqueta.pedido)) continue;
					}

					if (blnSequenciaInicio)
					{
						if (intSequenciaInicio > etiqueta.sequencia) continue;
					}

					if (blnSequenciaFim)
					{
						if (intSequenciaFim < etiqueta.sequencia) continue;
					}

					if (blnZona)
					{
						if (!strZona.Equals(etiqueta.zona_codigo)) continue;
					}

					_listaEtqParcial.Add(etiqueta);
				}

				#region [ Carrega dados no grid ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
					grdDados.SuspendLayout();

					grdDados.Rows.Clear();
					if (_listaEtqParcial.Count > 0) grdDados.Rows.Add(_listaEtqParcial.Count);

					for (int i = 0; i < _listaEtqParcial.Count; i++)
					{
						intSequencia++;
						etiqueta = _listaEtqParcial[i];
						strPedido = "|" + etiqueta.pedido + "|";
						if (!sbPedidos.ToString().Contains(strPedido))
						{
							sbPedidos.Append(strPedido);
							intQtdePedidos++;
						}
						preencheLinhaGrid(i, etiqueta);
						intQtdeTotalVolumes += etiqueta.qtde * etiqueta.qtde_volumes;
					}

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdDados.Rows.Count; i++)
					{
						if (grdDados.Rows[i].Selected) grdDados.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					grdDados.ResumeLayout();
				}
				#endregion

				#region [ Totais ]
				lblTotalRegistros.Text = Global.formataInteiro(_listaEtqParcial.Count);
				lblTotalPedidos.Text = Global.formataInteiro(intQtdePedidos);
				lblTotalVolumes.Text = Global.formataInteiro(intQtdeTotalVolumes);
				#endregion

				#region [ Resultado: Completo ou Parcial? ]
				if (_listaEtqParcial.Count == _listaEtqCompleta.Count)
				{
					lblDadosRelatorioCompletoOuParcial.Text = "COMPLETO";
					btnImprimirParcial.Enabled = false;
				}
				else
				{
					lblDadosRelatorioCompletoOuParcial.Text = "PARCIAL";
					btnImprimirParcial.Enabled = true;
				}
				#endregion

				grdDados.Focus();

				// Feedback da conclusão da consulta
				SystemSounds.Exclamation.Play();

				return true;
			}
			catch (Exception ex)
			{
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

		#region [ executaPesquisa ]
		private bool executaPesquisa()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaPesquisa()";
			String strSql;
			String strWhere = "";
			String strCriterioEmit = "";
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			try
			{
				#region [ Limpa campos dos dados de resposta ]
				limpaCamposDados();
				limpaCamposFiltro();
				#endregion

				#region [ Monta restrições da cláusula 'Where' ]
				if (dtpDataEmissao.Checked)
				{
					if ((dtpDataEmissao.Value != null) && (Global.converteInteiro(txtNsu.Text) == 0))
					{
						if (strWhere.Length > 0) strWhere += " AND";
						strWhere += " (dt_emissao = " + Global.sqlMontaDdMmYyyyParaSqlDateTime(dtpDataEmissao.Value.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador)) + ")";
					}
				}

				if (txtNsu.Text.Length > 0)
				{
					if (strWhere.Length > 0) strWhere += " AND";
					strWhere += " (tN1.id = " + Global.digitos(txtNsu.Text) + ")";
				}
				#endregion

				#region [ Há restrições definidas? ]
				if (strWhere.Length == 0)
				{
					avisoErro("É necessário informar a data ou o NSU para pesquisar!!");
					return false;
				}
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
                #endregion

				#region [ Monta critério de seleção do Emitente ]
				strCriterioEmit = " AND (EXISTS" +
								" (SELECT 1 FROM t_PEDIDO tPed" +
								"   INNER JOIN t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO tEtqPed ON tPed.pedido = tEtqPed.pedido" +
								"   WHERE tPed.id_nfe_emitente = " + Global.Usuario.emit_id +
								"   AND tEtqPed.id_wms_etq_n1 = tN1.id))";
				#endregion

				#region [ Monta o SQL ]
				strSql = "SELECT" +
							" tN1.id," +
							" tN1.dt_hr_emissao," +
							" tN1.usuario," +
							" (SELECT Count(*) FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO WHERE (id_wms_etq_n1=tN1.id)) AS qtde_pedidos," +
							" (SELECT Sum(qtde_volumes_pedido) FROM t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO WHERE (id_wms_etq_n1=tN1.id)) AS qtde_volumes" +
						" FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO tN1" +
						" WHERE" +
							strWhere +
							strCriterioEmit +
						" ORDER BY" +
							" tN1.id";
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
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_DATA_HORA].Value = Global.formataDataDdMmYyyyHhMmSsComSeparador(BD.readToDateTime(rowConsulta["dt_hr_emissao"]));
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_NSU].Value = BD.readToString(rowConsulta["id"]);
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_USUARIO].Value = BD.readToString(rowConsulta["usuario"]);
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_QTDE_PEDIDOS].Value = BD.readToInt(rowConsulta["qtde_pedidos"]).ToString();
						grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_QTDE_VOLUMES].Value = BD.readToInt(rowConsulta["qtde_volumes"]).ToString();
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

				grdPesquisa.Focus();

				// Feedback da conclusão da pesquisa
				SystemSounds.Exclamation.Play();

				return true;
			}
			catch (Exception ex)
			{
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

		#region [ montaDadosImpressaoEtiqueta ]
		private bool montaDadosImpressaoEtiqueta(EtiquetaDados etiqueta, out string textoImpressaoEtiqueta, out string strMsgErro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "montaDadosImpressaoEtiqueta()";
			StringBuilder sbEtiqueta = new StringBuilder("");
			string textoEtiqueta;
			string strTranportadora;
			string strZona;
			string strNumeroNF;
			string strDestinoUF;
			string strDestinoCidade;
			string strCliente;
			string strProduto;
			string strProdutoAux;
			string strNsuEtiqueta;
			string strVolume;
			string strIdentVolProdComposto;
			#endregion

			textoImpressaoEtiqueta = "";
			strMsgErro = "";
			try
			{
				if (etiqueta == null)
				{
					strMsgErro = "Não há dados para a impressão da etiqueta!!";
					return false;
				}

				strTranportadora = Global.filtraAcentuacao(etiqueta.transportadora_id == null ? "" : etiqueta.transportadora_id.ToUpper());
				strZona = (etiqueta.zona_codigo == null ? "" : etiqueta.zona_codigo.ToUpper());
				if (etiqueta.obs_3.Length > 0)
				{
					strNumeroNF = Global.filtraAcentuacao(etiqueta.obs_3 == null ? "" : etiqueta.obs_3);
				}
				else
				{
					strNumeroNF = Global.filtraAcentuacao(etiqueta.obs_2 == null ? "" : etiqueta.obs_2);
				}
				strDestinoUF = etiqueta.destino_uf.ToUpper();
				strDestinoCidade = Texto.iniciaisEmMaiusculas(Global.filtraAcentuacao(etiqueta.destino_cidade == null ? "" : etiqueta.destino_cidade));
				strCliente = Global.filtraAcentuacao(etiqueta.nome_cliente == null ? "" : etiqueta.nome_cliente);
				strProduto = Global.filtraAcentuacao(etiqueta.descricao == null ? "" : etiqueta.descricao);

				for (int i = 0; i < etiqueta.ctrlNumeracaoVolume.Count; i++)
				{
					strNsuEtiqueta = etiqueta.id_N1.ToString().PadLeft(3, '0') + "-" + etiqueta.sequencia.ToString().PadLeft(4, '0') + " / " + etiqueta.id_N3.ToString().PadLeft(4, '0');
					strVolume = etiqueta.ctrlNumeracaoVolume[i].numeracaoVolume.ToString() + '/' + etiqueta.qtde_volumes_pedido.ToString();
					strIdentVolProdComposto = etiqueta.ctrlNumeracaoVolume[i].identificacaoVolumeProdutoComposto;
					if (strIdentVolProdComposto.Length > 0) strIdentVolProdComposto = strIdentVolProdComposto + " - ";
					strProdutoAux = strIdentVolProdComposto + strProduto;

					if (!montaDadosImpressaoEtiquetaIndividual(strTranportadora, strZona, strNumeroNF, strVolume, strCliente, strProdutoAux, strNsuEtiqueta, strDestinoUF, strDestinoCidade, out textoEtiqueta, out strMsgErro))
					{
						throw new Exception("Falha ao preparar os dados para a etiqueta do pedido " + etiqueta.pedido + ", produto (" + etiqueta.fabricante + ")" + etiqueta.produto + "!!");
					}

					sbEtiqueta.Append(textoEtiqueta);
				}

				textoImpressaoEtiqueta = sbEtiqueta.ToString();
				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ montaDadosImpressaoEtiquetaIndividual ]
		private bool montaDadosImpressaoEtiquetaIndividual(string transportadora, string zona, string numeroNF, string volume, string cliente, string produto, string nsuEtiqueta, string destinoUF, string destinoCidade, out string textoImpressaoEtiqueta, out string strMsgErro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "montaDadosImpressaoEtiquetaIndividual()";
			const string COORDENADAS_MARGEM_X = "0030";
			#region [ Padrão ]
			const string FONTE_PADRAO = "9";
			const string SUB_FONTE_PADRAO = "004";
			const string MULTIPLICADOR_HORIZONTAL_PADRAO = "1";
			const string MULTIPLICADOR_VERTICAL_PADRAO = "1";
			#endregion
			#region [ Reduzida ]
			const string FONTE_REDUZIDA = "9";
			const string SUB_FONTE_REDUZIDA = "002";
			const string MULTIPLICADOR_HORIZONTAL_REDUZIDA = "1";
			const string MULTIPLICADOR_VERTICAL_REDUZIDA = "1";
			#endregion
			#region [ NF ]
			const string FONTE_NF = "9";
			const string SUB_FONTE_NF = "006";
			const string MULTIPLICADOR_HORIZONTAL_NF = "2";
			const string MULTIPLICADOR_VERTICAL_NF = "2";
			#endregion
			#region [ Volume ]
			const string FONTE_VOLUME = "9";
			const string SUB_FONTE_VOLUME = "004";
			const string MULTIPLICADOR_HORIZONTAL_VOLUME = "2";
			const string MULTIPLICADOR_VERTICAL_VOLUME = "2";
			#endregion
			#region [ Transportadora ]
			const string FONTE_TRANSPORTADORA = "9";
			const string SUB_FONTE_TRANSPORTADORA = "003";
			const string MULTIPLICADOR_HORIZONTAL_TRANSPORTADORA = "2";
			const string MULTIPLICADOR_VERTICAL_TRANSPORTADORA = "2";
			#endregion
			#region [ UF/Cidade destino ]
			const string FONTE_LOCALIDADE_DESTINO = "9";
			const string SUB_FONTE_LOCALIDADE_DESTINO = "003";
			const string MULTIPLICADOR_HORIZONTAL_LOCALIDADE_DESTINO = "2";
			const string MULTIPLICADOR_VERTICAL_LOCALIDADE_DESTINO = "2";
			#endregion
			#region [ Zona ]
			const string FONTE_ZONA = "9";
			const string SUB_FONTE_ZONA = "003";
			const string MULTIPLICADOR_HORIZONTAL_ZONA = "2";
			const string MULTIPLICADOR_VERTICAL_ZONA = "2";
			#endregion

			string strUfCidadeDestino;
			#endregion

			textoImpressaoEtiqueta = "";
			strMsgErro = "";
			try
			{
				strUfCidadeDestino = (destinoUF == null ? "" : destinoUF) + " / " + (destinoCidade == null ? "" : destinoCidade);

				// IMPORTANTE
				// ==========
				// 1) A impressão da etiqueta é feita usando a linguagem PPLA
				// 2) A origem das coordenadas é no canto inferior esquerdo
				// 3) As medidas estão em milímetros
				// 4) Fonte:
				//    '0','1','2','3','4','5','6','7','8': Fontes internas (Subtipo fixo em '000')
				//    '9': fonte ASD smooth (Subtipos '000' a '006')
				//         '000': 4 points
				//         '001': 6 points
				//         '002': 8 points
				//         '003': 10 points
				//         '004': 12 points
				//         '005': 14 points
				//         '006': 18 points
				textoImpressaoEtiqueta =
					CODIGO_STX + "L" + CODIGO_CR + // Enters label formatting state
					"H12" + CODIGO_CR + // Temperatura da cabeça p/ controlar o contraste (padrão: H10, máximo: H20, máximo recomendável: H16)
					"D11" + CODIGO_CR + // Sets width and height pixel size (default: D22)
					// NSU etiqueta
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_REDUZIDA + // Fonte
					MULTIPLICADOR_HORIZONTAL_REDUZIDA + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_REDUZIDA + // Multiplicador Vertical
					SUB_FONTE_REDUZIDA + // Subtipo da fonte
					"0020" + // Coordenadas Y
					COORDENADAS_MARGEM_X + // Coordenadas X
					(nsuEtiqueta == null ? "" : nsuEtiqueta) + CODIGO_CR + // Texto a ser impresso
					// Produto
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_PADRAO + // Fonte
					MULTIPLICADOR_HORIZONTAL_PADRAO + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_PADRAO + // Multiplicador Vertical
					SUB_FONTE_PADRAO + // Subtipo da fonte
					"0070" + // Coordenadas Y
					COORDENADAS_MARGEM_X + // Coordenadas X
					(produto == null ? "" : produto) + CODIGO_CR + // Texto a ser impresso
					// Cliente
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_PADRAO + // Fonte
					MULTIPLICADOR_HORIZONTAL_PADRAO + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_PADRAO + // Multiplicador Vertical
					SUB_FONTE_PADRAO + // Subtipo da fonte
					"0130" + // Coordenadas Y
					COORDENADAS_MARGEM_X + // Coordenadas X
					(cliente == null ? "" : cliente) + CODIGO_CR + // Texto a ser impresso
					// Zona
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_ZONA + // Fonte
					MULTIPLICADOR_HORIZONTAL_ZONA + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_ZONA + // Multiplicador Vertical
					SUB_FONTE_ZONA + // Subtipo da fonte
					"0380" + // Coordenadas Y
					"0870" + // Coordenadas X
					(zona == null ? "" : zona) + CODIGO_CR + // Texto a ser impresso
					// Transportadora
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_TRANSPORTADORA + // Fonte
					MULTIPLICADOR_HORIZONTAL_TRANSPORTADORA + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_TRANSPORTADORA + // Multiplicador Vertical
					SUB_FONTE_TRANSPORTADORA + // Subtipo da fonte
					"0380" + // Coordenadas Y
					COORDENADAS_MARGEM_X + // Coordenadas X
					(transportadora == null ? "" : transportadora) + CODIGO_CR + // Texto a ser impresso
					// UF/Cidade de destino
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_LOCALIDADE_DESTINO + // Fonte
					MULTIPLICADOR_HORIZONTAL_LOCALIDADE_DESTINO + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_LOCALIDADE_DESTINO + // Multiplicador Vertical
					SUB_FONTE_LOCALIDADE_DESTINO + // Subtipo da fonte
					"0180" + // Coordenadas Y
					COORDENADAS_MARGEM_X + // Coordenadas X
					strUfCidadeDestino + CODIGO_CR + // Texto a ser impresso
					// NF
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_NF + // Fonte
					MULTIPLICADOR_HORIZONTAL_NF + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_NF + // Multiplicador Vertical
					SUB_FONTE_NF + // Subtipo da fonte
					"0250" + // Coordenadas Y
					COORDENADAS_MARGEM_X + // Coordenadas X
					"NF." + (numeroNF == null ? "" : numeroNF) + CODIGO_CR + // Texto a ser impresso
					// Volume
					"1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
					FONTE_VOLUME + // Fonte
					MULTIPLICADOR_HORIZONTAL_VOLUME + // Multiplicador Horizontal
					MULTIPLICADOR_VERTICAL_VOLUME + // Multiplicador Vertical
					SUB_FONTE_VOLUME + // Subtipo da fonte
					"0250" + // Coordenadas Y
					"0800" + // Coordenadas X
					(volume == null ? "" : volume) + CODIGO_CR + // Texto a ser impresso
					// Comandos de finalização da etiqueta
					"Q0001" + CODIGO_CR + // Sets the quantity of labels to print
					"E" + CODIGO_CR // Ends the job and exit from label formatting mode
					;

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ limpaCamposPesquisa ]
		private void limpaCamposPesquisa()
		{
			grdPesquisa.Rows.Clear();
			dtpDataEmissao.Value = DateTime.Now.Date;
			dtpDataEmissao.Checked = false;
			txtNsu.Text = "";
		}
		#endregion

		#region [ limpaCamposDados ]
		private void limpaCamposDados()
		{
			lblNsuDadosRelatorio.Text = "";
			lblDadosRelatorioCompletoOuParcial.Text = "";
			grdDados.Rows.Clear();
			lblTotalRegistros.Text = "";
			lblTotalPedidos.Text = "";
			lblTotalVolumes.Text = "";
			btnImprimirCompleto.Enabled = false;
			btnImprimirParcial.Enabled = false;
			btnImprimirVolume.Enabled = false;
		}
		#endregion

		#region [ limpaCamposFiltro ]
		private void limpaCamposFiltro()
		{
			txtNumeroNF.Text = "";
			txtPedido.Text = "";
			txtSequenciaInicio.Text = "";
			txtSequenciaFim.Text = "";
			txtZona.Text = "";
			btnImprimirParcial.Enabled = false;
		}
		#endregion

		#region [ trataBotaoPrinterDialog ]
		private void trataBotaoPrinterDialog()
		{
			printDialog.ShowDialog();
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

			if (grdPesquisa.Rows.Count == 1)
			{
				grdPesquisa.Rows[0].Selected = true;
				return executaConsulta();
			}

			return true;
		}
		#endregion

		#region [ trataBotaoConsultar ]
		private bool trataBotaoConsultar()
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

			return executaConsulta();
		}
		#endregion

		#region [ trataBotaoFiltrar ]
		private bool trataBotaoFiltrar()
		{
			return executaFiltragemDados(eFiltroPreenchimentoObrigatorio.OBRIGATORIO);
		}
		#endregion

		#region [ trataBotaoLimparFiltro ]
		private bool trataBotaoLimparFiltro()
		{
			limpaCamposFiltro();
			return executaFiltragemDados(eFiltroPreenchimentoObrigatorio.OPCIONAL);
		}
		#endregion

		#region [ trataBotaoImprimirCompleto ]
		private void trataBotaoImprimirCompleto()
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "trataBotaoImprimirCompleto()";
			int totalVolumes;
			int intPedidosAlertaStEntrega = 0;
			int intPedidosAlertaOpTriangular = 0;
			String strAux;
			String strDescricaoLog;
			String strMsgErroLog = "";
			String strMsgErro;
			String textoEtiqueta;
			String textoEtiquetaBatch;
			StringBuilder sbEtiqueta = new StringBuilder("");
			StringBuilder sbAlertaStEntrega = new StringBuilder("");
			StringBuilder sbAlertaOpTriangular = new StringBuilder("");
			StringBuilder sbListaPedidosAlertaStEntrega = new StringBuilder("");
			StringBuilder sbListaPedidosAlertaOpTriangular = new StringBuilder("");
			Log log = new Log();
			WmsEtqN1SepZonaRel wmsEtqN1SepZonaRel;
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
				if (_listaEtqCompleta.Count == 0)
				{
					avisoErro("Não há dados para imprimir!!");
					return;
				}

				if (grdDados.Rows.Count == 0)
				{
					avisoErro("Consulte os dados detalhados do relatório!!");
					return;
				}

				#region [ Verifica se o status de entrega do pedido permite a impressão ]
				for (int i = 0; i < _listaEtqCompleta.Count; i++)
				{
					if (Global.isStEntregaPedidoBloqueadoParaImpressaoEtiqueta(_listaEtqCompleta[i].st_entrega))
					{
						if (sbListaPedidosAlertaStEntrega.ToString().IndexOf("|" + _listaEtqCompleta[i].pedido + "|") == -1)
						{
							intPedidosAlertaStEntrega++;
							sbListaPedidosAlertaStEntrega.Append("|" + _listaEtqCompleta[i].pedido + "|");
							if (sbAlertaStEntrega.Length > 0) sbAlertaStEntrega.Append("\n");
							sbAlertaStEntrega.Append(_listaEtqCompleta[i].pedido + ": " + Global.stEntregaPedidoDescricao(_listaEtqCompleta[i].st_entrega));
						}
					}
				}

				if (intPedidosAlertaStEntrega > 0)
				{
					if (intPedidosAlertaStEntrega == 1)
					{
						strMsgErro = "A etiqueta do seguinte pedido NÃO pode ser impressa devido ao status de entrega:\n" + sbAlertaStEntrega.ToString();
					}
					else
					{
						strMsgErro = "As etiquetas dos seguintes pedidos NÃO podem ser impressas devido ao status de entrega:\n" + sbAlertaStEntrega.ToString();
					}

					Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Verifica se existe pedido com emissão de nota triangular ]
				for (int i = 0; i < _listaEtqCompleta.Count; i++)
				{
					if (_listaEtqCompleta[i].origem_uf != _listaEtqCompleta[i].destino_uf)
					{
						if (sbListaPedidosAlertaOpTriangular.ToString().IndexOf("|" + _listaEtqCompleta[i].pedido + "|") == -1)
						{
							intPedidosAlertaOpTriangular++;
							sbListaPedidosAlertaOpTriangular.Append("|" + _listaEtqCompleta[i].pedido + "|");
							if (sbAlertaOpTriangular.Length > 0) sbAlertaOpTriangular.Append("\n");
							sbAlertaOpTriangular.Append(_listaEtqCompleta[i].pedido + ": origem " + _listaEtqCompleta[i].origem_uf + ", destino " + _listaEtqCompleta[i].destino_uf);
						}
					}
				}

				if (intPedidosAlertaOpTriangular > 0)
				{
					if (intPedidosAlertaOpTriangular == 1)
					{
						strMsgErro = "O seguinte pedido está relacionado a notas fiscais de venda e de remessa:\n" + sbAlertaOpTriangular.ToString() + "\n\nContinuar com a impressão?";
					}
					else
					{
						strMsgErro = "Os seguintes pedidos estão relacionados a notas fiscais de venda e de remessa:\n" + sbAlertaOpTriangular.ToString() + "\n\nContinuar com a impressão?";
					}

					if (!confirma(strMsgErro)) return;

				}
				#endregion

				while (true)
				{
					if (!printDialog.PrinterSettings.PrinterName.ToUpper().Contains("ARGOX"))
					{
						if (confirma("A impressora selecionada (" + printDialog.PrinterSettings.PrinterName + ") não parece ser a impressora de etiquetas!!\nDeseja selecionar outra impressora?"))
						{
							printDialog.ShowDialog();
						}
						else break;
					}
					else break;
				}

				if (!calculaTotalVolumes(_listaEtqCompleta, out totalVolumes, out strMsgErro))
				{
					avisoErro("Erro ao obter o total de etiquetas a serem impressas!!\n" + strMsgErro);
					return;
				}

				if (!confirma("Confirma a impressão da listagem COMPLETA (" + totalVolumes.ToString() + " etiquetas)?")) return;

				#region [ Verifica se já foi feita alguma impressão completa ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
				wmsEtqN1SepZonaRel = ComumDAO.getWmsEtqN1SepZonaRel(_idNsuUltimoSelecionado);
				if (wmsEtqN1SepZonaRel.etiqueta_impressao_status != 0)
				{
					strAux = "Já foi realizada uma impressão completa em " + Global.formataDataDdMmYyyyHhMmComSeparador(wmsEtqN1SepZonaRel.etiqueta_impressao_ultima_vez_data_hora) + " por " + wmsEtqN1SepZonaRel.etiqueta_impressao_ultima_vez_usuario + "!!" +
							"\n" +
							"Continua mesmo assim?";
					if (!confirma(strAux)) return;
				}
				#endregion

				info(ModoExibicaoMensagemRodape.EmExecucao, "preparando dados das etiquetas");

				#region [ Monta os dados e imprime ]
				for (int i = 0; i < _listaEtqCompleta.Count; i++)
				{
					if (!montaDadosImpressaoEtiqueta(_listaEtqCompleta[i], out textoEtiqueta, out strMsgErro))
					{
						throw new Exception(strMsgErro);
					}

					sbEtiqueta.Append(textoEtiqueta);
				}

				textoEtiquetaBatch =
					CODIGO_STX + "m" + CODIGO_CR + // Sets measurement to metric
					CODIGO_STX + "r" + CODIGO_CR + // Selects reflective sensor for gap
					CODIGO_STX + "V0" + CODIGO_CR + // Sets cutter and dispenser configuration ('0': no cutter and peeler function; '1': Enables cutter function; '4': Enables peeler function)
					CODIGO_SOH + "D" + CODIGO_CR + // Disables the interaction command
					CODIGO_STX + "f920" + CODIGO_CR + // Sets stop position and automatic back-feed for the label stock (Back-feed will not be activated if xxx is less than 220)
					sbEtiqueta.ToString();

				info(ModoExibicaoMensagemRodape.EmExecucao, "enviando dados para a impressora");

				RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, textoEtiquetaBatch);
				#endregion

				#region [ Grava log no BD ]
				strDescricaoLog = "Relatório Separação (Zona): impressão completa das etiquetas do relatório nº " + _idNsuUltimoSelecionado.ToString() + " (" + totalVolumes.ToString() + " etiquetas impressas)";
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.EtqWms.LogOperacao.ETIQUETA_WMS_IMPRESSAO_COMPLETA;
				log.complemento = strDescricaoLog;
				LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				#region [ Registra a impressão na tabela do relatório ]
				if (!ComumDAO.atualizaWmsEtqN1SepZonaRelEmissaoCompleta(_idNsuUltimoSelecionado, Global.Usuario.usuario, out strMsgErro))
				{
					avisoErro("Falha ao tentar registrar no banco de dados a impressão completa das etiquetas!!");
				}
				#endregion

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

		#region [ trataBotaoImprimirParcial ]
		private void trataBotaoImprimirParcial()
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "trataBotaoImprimirParcial()";
			int totalVolumes = 0;
			int intPedidosAlertaStEntrega = 0;
			int intPedidosAlertaOpTriangular = 0;
			String strDescricaoLog;
			String strMsgErroLog = "";
			String strMsgErro;
			String textoEtiqueta;
			String textoEtiquetaBatch;
			StringBuilder sbEtiqueta = new StringBuilder("");
			StringBuilder sbAlertaStEntrega = new StringBuilder("");
			StringBuilder sbAlertaOpTriangular = new StringBuilder("");
			StringBuilder sbListaPedidosAlertaStEntrega = new StringBuilder("");
			StringBuilder sbListaPedidosAlertaOpTriangular = new StringBuilder("");
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
				if (_listaEtqParcial.Count == 0)
				{
					avisoErro("Não há dados para imprimir!!");
					return;
				}

				if (_listaEtqParcial.Count == _listaEtqCompleta.Count)
				{
					strMsgErro = "Nenhuma filtragem foi realizada!!\nPara imprimir a listagem completa, por favor utilize o botão 'Imprimir Completo'!!";
					avisoErro(strMsgErro);
					return;
				}

				#region [ Verifica se o status de entrega do pedido permite a impressão ]
				for (int i = 0; i < _listaEtqParcial.Count; i++)
				{
					if (Global.isStEntregaPedidoBloqueadoParaImpressaoEtiqueta(_listaEtqParcial[i].st_entrega))
					{
						if (sbListaPedidosAlertaStEntrega.ToString().IndexOf("|" + _listaEtqParcial[i].pedido + "|") == -1)
						{
							intPedidosAlertaStEntrega++;
							sbListaPedidosAlertaStEntrega.Append("|" + _listaEtqParcial[i].pedido + "|");
							if (sbAlertaStEntrega.Length > 0) sbAlertaStEntrega.Append("\n");
							sbAlertaStEntrega.Append(_listaEtqParcial[i].pedido + ": " + Global.stEntregaPedidoDescricao(_listaEtqParcial[i].st_entrega));
						}
					}
				}

				if (intPedidosAlertaStEntrega > 0)
				{
					if (intPedidosAlertaStEntrega == 1)
					{
						strMsgErro = "A etiqueta do seguinte pedido NÃO pode ser impressa devido ao status de entrega:\n" + sbAlertaStEntrega.ToString();
					}
					else
					{
						strMsgErro = "As etiquetas dos seguintes pedidos NÃO podem ser impressas devido ao status de entrega:\n" + sbAlertaStEntrega.ToString();
					}

					Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion


				#region [ Verifica se existe pedido com emissão de nota triangular ]
				for (int i = 0; i < _listaEtqParcial.Count; i++)
				{
					if (_listaEtqParcial[i].origem_uf != _listaEtqParcial[i].destino_uf)
					{
						if (sbListaPedidosAlertaOpTriangular.ToString().IndexOf("|" + _listaEtqParcial[i].pedido + "|") == -1)
						{
							intPedidosAlertaOpTriangular++;
							sbListaPedidosAlertaOpTriangular.Append("|" + _listaEtqParcial[i].pedido + "|");
							if (sbAlertaOpTriangular.Length > 0) sbAlertaOpTriangular.Append("\n");
							sbAlertaOpTriangular.Append(_listaEtqParcial[i].pedido + ": origem " + _listaEtqParcial[i].origem_uf + ", destino " + _listaEtqParcial[i].destino_uf);
						}
					}
				}

				if (intPedidosAlertaOpTriangular > 0)
				{
					if (intPedidosAlertaOpTriangular == 1)
					{
						strMsgErro = "O seguinte pedido está relacionado a notas fiscais de venda e de remessa:\n" + sbAlertaOpTriangular.ToString() + "\n\nContinuar com a impressão?";
					}
					else
					{
						strMsgErro = "Os seguintes pedidos estão relacionados a notas fiscais de venda e de remessa:\n" + sbAlertaOpTriangular.ToString() + "\n\nContinuar com a impressão?";
					}

					if (!confirma(strMsgErro)) return;

				}
				#endregion


				while (true)
				{
					if (!printDialog.PrinterSettings.PrinterName.ToUpper().Contains("ARGOX"))
					{
						if (confirma("A impressora selecionada (" + printDialog.PrinterSettings.PrinterName + ") não parece ser a impressora de etiquetas!!\nDeseja selecionar outra impressora?"))
						{
							printDialog.ShowDialog();
						}
						else break;
					}
					else break;
				}

				if (!calculaTotalVolumes(_listaEtqParcial, out totalVolumes, out strMsgErro))
				{
					avisoErro("Erro ao obter o total de etiquetas a serem impressas!!\n" + strMsgErro);
					return;
				}

				if (!confirma("Confirma a impressão da listagem PARCIAL (" + totalVolumes.ToString() + " etiquetas)?")) return;

				info(ModoExibicaoMensagemRodape.EmExecucao, "preparando dados das etiquetas");

				#region [ Monta os dados e imprime ]
				for (int i = 0; i < _listaEtqParcial.Count; i++)
				{
					if (!montaDadosImpressaoEtiqueta(_listaEtqParcial[i], out textoEtiqueta, out strMsgErro))
					{
						throw new Exception(strMsgErro);
					}

					sbEtiqueta.Append(textoEtiqueta);
				}

				textoEtiquetaBatch =
					CODIGO_STX + "m" + CODIGO_CR + // Sets measurement to metric
					CODIGO_STX + "r" + CODIGO_CR + // Selects reflective sensor for gap
					CODIGO_STX + "V0" + CODIGO_CR + // Sets cutter and dispenser configuration ('0': no cutter and peeler function; '1': Enables cutter function; '4': Enables peeler function)
					CODIGO_SOH + "D" + CODIGO_CR + // Disables the interaction command
					CODIGO_STX + "f920" + CODIGO_CR + // Sets stop position and automatic back-feed for the label stock (Back-feed will not be activated if xxx is less than 220)
					sbEtiqueta.ToString();

				info(ModoExibicaoMensagemRodape.EmExecucao, "enviando dados para a impressora");

				RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, textoEtiquetaBatch);
				#endregion

				#region [ Grava log no BD ]
				strDescricaoLog = "Relatório Separação (Zona): impressão parcial das etiquetas do relatório nº " + _idNsuUltimoSelecionado.ToString() + " (" + totalVolumes.ToString() + " etiquetas impressas)";
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.EtqWms.LogOperacao.ETIQUETA_WMS_IMPRESSAO_PARCIAL;
				log.complemento = strDescricaoLog;
				LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

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

		#region [ trataBotaoImprimirVolume ]
		private void trataBotaoImprimirVolume()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "trataBotaoImprimirVolume()";
			StringBuilder sbEtiqueta = new StringBuilder("");
			string textoEtiquetaBatch;
			string strAux = "";
			string strMsg;
			string strMsgErro;
			string strDescricaoLog;
			string strMsgErroLog = "";
			FVolumeSeleciona fVolumeSeleciona;
			DialogResult drResultado;
			FVolumeSeleciona.eOpcaoSelecao opcaoSelecionada;
			int numeroNF;
			int numeroVolumeUnico;
			int numeroVolumeIntervaloInicio;
			int numeroVolumeIntervaloFim;
			int totalVolumesImpressos = 0;
			string textoEtiqueta;
			string strTranportadora;
			string strZona;
			string strNumeroNF;
			string strDestinoUF;
			string strDestinoCidade;
			string strCliente;
			string strProduto;
			string strProdutoAux;
			string strNsuEtiqueta;
			string strVolume;
			string strIdentVolProdComposto;
			EtiquetaDados etiqueta;
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
				if (_listaEtqParcial.Count == 0)
				{
					avisoErro("Não há dados para imprimir!!");
					return;
				}

				#region [ Exibe painel p/ obter o nº volume a ser impresso ]
				fVolumeSeleciona = new FVolumeSeleciona(_listaEtqParcial);
				drResultado = fVolumeSeleciona.ShowDialog(this);
				if (drResultado != DialogResult.OK)
				{
					strMsgErro = "Impressão cancelada!!";
					avisoErro(strMsgErro);
					return;
				}

				opcaoSelecionada = fVolumeSeleciona.opcaoSelecionada;
				numeroNF = fVolumeSeleciona.numNF;
				numeroVolumeUnico = fVolumeSeleciona.numVolumeUnico;
				numeroVolumeIntervaloInicio = (fVolumeSeleciona.numVolumeIntervaloInicio == 0 ? 1 : fVolumeSeleciona.numVolumeIntervaloInicio);
				numeroVolumeIntervaloFim = fVolumeSeleciona.numVolumeIntervaloFim;
				#endregion

				while (true)
				{
					if (!printDialog.PrinterSettings.PrinterName.ToUpper().Contains("ARGOX"))
					{
						if (confirma("A impressora selecionada (" + printDialog.PrinterSettings.PrinterName + ") não parece ser a impressora de etiquetas!!\nDeseja selecionar outra impressora?"))
						{
							printDialog.ShowDialog();
						}
						else break;
					}
					else break;
				}

				if (opcaoSelecionada == FVolumeSeleciona.eOpcaoSelecao.VOLUME_UNICO)
				{
					strMsg = "Confirma a impressão da etiqueta do volume nº " + numeroVolumeUnico.ToString() + " (NF: " + numeroNF.ToString() + ")?";
					if (!confirma(strMsg)) return;
				}
				else if (opcaoSelecionada == FVolumeSeleciona.eOpcaoSelecao.VOLUME_INTERVALO)
				{
					if (numeroVolumeIntervaloFim == 0)
					{
						strMsg = "Confirma a impressão das etiquetas do volume nº " + numeroVolumeIntervaloInicio.ToString() + " até o final (NF: " + numeroNF.ToString() + ")?";
					}
					else
					{
						strMsg = "Confirma a impressão das etiquetas do volume nº " + numeroVolumeIntervaloInicio.ToString() + " até nº " + numeroVolumeIntervaloFim.ToString() + " (NF: " + numeroNF.ToString() + ")?";
					}
					if (!confirma(strMsg)) return;
				}
				else
				{
					avisoErro("Opção inválida (" + ((int)opcaoSelecionada).ToString() + ")!!");
					return;
				}

				info(ModoExibicaoMensagemRodape.EmExecucao, "preparando dados da(s) etiqueta(s)");

				#region [ Monta os dados e imprime ]
				for (int ic = 0; ic < _listaEtqParcial.Count; ic++)
				{
					if (_listaEtqParcial[ic].obs_3.Length > 0)
					{
						if ((int)Global.converteInteiro(_listaEtqParcial[ic].obs_3) != numeroNF) continue;
					}
					else
					{
						if ((int)Global.converteInteiro(_listaEtqParcial[ic].obs_2) != numeroNF) continue;
					}

					etiqueta = _listaEtqParcial[ic];

					#region [ Verifica se o status de entrega do pedido permite a impressão ]
					if (Global.isStEntregaPedidoBloqueadoParaImpressaoEtiqueta(etiqueta.st_entrega))
					{
						strMsgErro = "A etiqueta do pedido " + etiqueta.pedido + " NÃO pode ser impressa devido ao status de entrega: " + Global.stEntregaPedidoDescricao(etiqueta.st_entrega);
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#region [ Verifica se o pedido tem emissão de nota triangular ]
					if (etiqueta.origem_uf != etiqueta.destino_uf)
					{
						strMsgErro = "A etiqueta do pedido " + etiqueta.pedido + " está relacionado a notas fiscais de venda e de remessa: origem " + etiqueta.origem_uf + ", destino " + etiqueta.destino_uf + "\nContinuar com a impressão?";
						if (!confirma(strMsgErro)) return;
					}
					#endregion


					strTranportadora = Global.filtraAcentuacao(etiqueta.transportadora_id == null ? "" : etiqueta.transportadora_id.ToUpper());
					strZona = (etiqueta.zona_codigo == null ? "" : etiqueta.zona_codigo.ToUpper());
					if (etiqueta.obs_3.Length > 0)
					{
						strNumeroNF = Global.filtraAcentuacao(etiqueta.obs_3 == null ? "" : etiqueta.obs_3);
					}
					else
					{
						strNumeroNF = Global.filtraAcentuacao(etiqueta.obs_2 == null ? "" : etiqueta.obs_2);
					}
					strDestinoUF = etiqueta.destino_uf.ToUpper();
					strDestinoCidade = Texto.iniciaisEmMaiusculas(Global.filtraAcentuacao(etiqueta.destino_cidade == null ? "" : etiqueta.destino_cidade));
					strCliente = Global.filtraAcentuacao(etiqueta.nome_cliente == null ? "" : etiqueta.nome_cliente);
					strProduto = Global.filtraAcentuacao(etiqueta.descricao == null ? "" : etiqueta.descricao);

					if (opcaoSelecionada == FVolumeSeleciona.eOpcaoSelecao.VOLUME_UNICO)
					{
						for (int i = 0; i < etiqueta.ctrlNumeracaoVolume.Count; i++)
						{
							if (etiqueta.ctrlNumeracaoVolume[i].numeracaoVolume == numeroVolumeUnico)
							{
								strNsuEtiqueta = etiqueta.id_N1.ToString().PadLeft(3, '0') + "-" + etiqueta.sequencia.ToString().PadLeft(4, '0') + " / " + etiqueta.id_N3.ToString().PadLeft(4, '0');
								strVolume = etiqueta.ctrlNumeracaoVolume[i].numeracaoVolume.ToString() + '/' + etiqueta.qtde_volumes_pedido.ToString();
								strIdentVolProdComposto = etiqueta.ctrlNumeracaoVolume[i].identificacaoVolumeProdutoComposto;
								if (strIdentVolProdComposto.Length > 0) strIdentVolProdComposto = strIdentVolProdComposto + " - ";
								strProdutoAux = strIdentVolProdComposto + strProduto;

								if (!montaDadosImpressaoEtiquetaIndividual(strTranportadora, strZona, strNumeroNF, strVolume, strCliente, strProdutoAux, strNsuEtiqueta, strDestinoUF, strDestinoCidade, out textoEtiqueta, out strMsgErro))
								{
									throw new Exception("Falha ao preparar os dados para a etiqueta do pedido " + etiqueta.pedido + ", produto (" + etiqueta.fabricante + ")" + etiqueta.produto + "!!");
								}

								totalVolumesImpressos++;
								sbEtiqueta.Append(textoEtiqueta);
								break;
							}
						}
					}
					else if (opcaoSelecionada == FVolumeSeleciona.eOpcaoSelecao.VOLUME_INTERVALO)
					{
						for (int i = 0; i < etiqueta.ctrlNumeracaoVolume.Count; i++)
						{
							// Verifica se o volume está dentro do intervalo especificado
							// Lembrando que é permitido omitir o campo (ex: 3 até 0, ou seja, reimprimir a partir do volume nº 3 até o final)
							if (numeroVolumeIntervaloInicio > 0)
							{
								if (etiqueta.ctrlNumeracaoVolume[i].numeracaoVolume < numeroVolumeIntervaloInicio) continue;
							}

							if (numeroVolumeIntervaloFim > 0)
							{
								if (etiqueta.ctrlNumeracaoVolume[i].numeracaoVolume > numeroVolumeIntervaloFim) continue;
							}

							strNsuEtiqueta = etiqueta.id_N1.ToString().PadLeft(3, '0') + "-" + etiqueta.sequencia.ToString().PadLeft(4, '0') + " / " + etiqueta.id_N3.ToString().PadLeft(4, '0');
							strVolume = etiqueta.ctrlNumeracaoVolume[i].numeracaoVolume.ToString() + '/' + etiqueta.qtde_volumes_pedido.ToString();
							strIdentVolProdComposto = etiqueta.ctrlNumeracaoVolume[i].identificacaoVolumeProdutoComposto;
							if (strIdentVolProdComposto.Length > 0) strIdentVolProdComposto = strIdentVolProdComposto + " - ";
							strProdutoAux = strIdentVolProdComposto + strProduto;

							if (!montaDadosImpressaoEtiquetaIndividual(strTranportadora, strZona, strNumeroNF, strVolume, strCliente, strProdutoAux, strNsuEtiqueta, strDestinoUF, strDestinoCidade, out textoEtiqueta, out strMsgErro))
							{
								throw new Exception("Falha ao preparar os dados para a etiqueta do pedido " + etiqueta.pedido + ", produto (" + etiqueta.fabricante + ")" + etiqueta.produto + "!!");
							}

							totalVolumesImpressos++;
							sbEtiqueta.Append(textoEtiqueta);
						}
					}
				}

				if (sbEtiqueta.ToString().Trim().Length == 0)
				{
					throw new Exception("Falha ao tentar obter os dados para a etiqueta da NF " + numeroNF.ToString() + "!!");
				}

				textoEtiquetaBatch =
					CODIGO_STX + "m" + CODIGO_CR + // Sets measurement to metric
					CODIGO_STX + "r" + CODIGO_CR + // Selects reflective sensor for gap
					CODIGO_STX + "V0" + CODIGO_CR + // Sets cutter and dispenser configuration ('0': no cutter and peeler function; '1': Enables cutter function; '4': Enables peeler function)
					CODIGO_SOH + "D" + CODIGO_CR + // Disables the interaction command
					CODIGO_STX + "f920" + CODIGO_CR + // Sets stop position and automatic back-feed for the label stock (Back-feed will not be activated if xxx is less than 220)
					sbEtiqueta.ToString();

				info(ModoExibicaoMensagemRodape.EmExecucao, "enviando dados para a impressora");

				RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, textoEtiquetaBatch);
				#endregion

				#region [ Grava log no BD ]
				if (opcaoSelecionada == FVolumeSeleciona.eOpcaoSelecao.VOLUME_UNICO)
				{
					strAux = "Volume único: nº " + numeroVolumeUnico.ToString() + " (NF: " + numeroNF.ToString() + ")";
				}
				else if (opcaoSelecionada == FVolumeSeleciona.eOpcaoSelecao.VOLUME_INTERVALO)
				{
					if (numeroVolumeIntervaloFim == 0)
					{
						strAux = "Intervalo de volumes: do nº " + numeroVolumeIntervaloInicio.ToString() + " até o final (NF: " + numeroNF.ToString() + ")";
					}
					else
					{
						strAux = "Intervalo de volumes: do nº " + numeroVolumeIntervaloInicio.ToString() + " até nº " + numeroVolumeIntervaloFim.ToString() + " (NF: " + numeroNF.ToString() + ")";
					}
				}

				strDescricaoLog = "Relatório Separação (Zona): impressão de etiquetas de volumes específicos do relatório nº " + _idNsuUltimoSelecionado.ToString() + " (" + totalVolumesImpressos.ToString() + " etiquetas impressas); " + strAux;
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.EtqWms.LogOperacao.ETIQUETA_WMS_IMPRESSAO_VOLUME;
				log.complemento = strDescricaoLog;
				LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

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

		#region [ preencheLinhaGrid ]
		private bool preencheLinhaGrid(int rowIndex, EtiquetaDados etiqueta)
		{
			try
			{
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_NSU_N1].Value = etiqueta.id_N1;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_NSU_N2].Value = etiqueta.id_N2;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_NSU_N3].Value = etiqueta.id_N3;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_SEQUENCIA].Value = etiqueta.sequencia;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_PEDIDO].Value = etiqueta.pedido;
				if (etiqueta.obs_3.Length > 0)
				{
					grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_NUMERO_NF].Value = etiqueta.obs_3;
				}
				else
				{
					grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_NUMERO_NF].Value = etiqueta.obs_2;
				}
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_TRANSPORTADORA].Value = etiqueta.transportadora_id;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_CLIENTE].Value = etiqueta.nome_cliente;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_ZONA].Value = etiqueta.zona_codigo;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_QTDE].Value = etiqueta.qtde;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_QTDE_VOLUMES].Value = etiqueta.qtde * etiqueta.qtde_volumes;
				grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_PRODUTO].Value = "(" + etiqueta.fabricante + ")" + etiqueta.produto + " - " + etiqueta.descricao;
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FEtiquetaImprime ]

		#region [ FEtiquetaImprime_Load ]
		private void FEtiquetaImprime_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCamposPesquisa();
				limpaCamposDados();
				limpaCamposFiltro();

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

		#region [ FEtiquetaImprime_Shown ]
		private void FEtiquetaImprime_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{

					#region [ Informa Emitente ]
					lblEmitente.Text = Global.Usuario.emit;
					#endregion


					#region [ Ajusta layout do header do grid (resultado da pesquisa) ]
					grdPesquisa.Columns[GRID_PESQ_COL_CHECKBOX].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdPesquisa.Columns[GRID_PESQ_COL_DATA_HORA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdPesquisa.Columns[GRID_PESQ_COL_DATA_HORA].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_NSU].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdPesquisa.Columns[GRID_PESQ_COL_NSU].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_QTDE_PEDIDOS].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdPesquisa.Columns[GRID_PESQ_COL_QTDE_PEDIDOS].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_QTDE_VOLUMES].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdPesquisa.Columns[GRID_PESQ_COL_QTDE_VOLUMES].ReadOnly = true;
					grdPesquisa.Columns[GRID_PESQ_COL_USUARIO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdPesquisa.Columns[GRID_PESQ_COL_USUARIO].ReadOnly = true;
					#endregion

					#region [ Ajusta layout do header do grid (dados do relatório) ]
					grdDados.Columns[GRID_DADOS_COL_CHECKBOX].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdDados.Columns[GRID_DADOS_COL_SEQUENCIA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdDados.Columns[GRID_DADOS_COL_SEQUENCIA].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_PEDIDO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdDados.Columns[GRID_DADOS_COL_PEDIDO].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_NUMERO_NF].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdDados.Columns[GRID_DADOS_COL_NUMERO_NF].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_TRANSPORTADORA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdDados.Columns[GRID_DADOS_COL_TRANSPORTADORA].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_CLIENTE].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdDados.Columns[GRID_DADOS_COL_CLIENTE].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_ZONA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
					grdDados.Columns[GRID_DADOS_COL_ZONA].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_QTDE].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdDados.Columns[GRID_DADOS_COL_QTDE].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_QTDE_VOLUMES].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
					grdDados.Columns[GRID_DADOS_COL_QTDE_VOLUMES].ReadOnly = true;
					grdDados.Columns[GRID_DADOS_COL_PRODUTO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
					grdDados.Columns[GRID_DADOS_COL_PRODUTO].ReadOnly = true;
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

		#region [ FEtiquetaImprime_FormClosing ]
		private void FEtiquetaImprime_FormClosing(object sender, FormClosingEventArgs e)
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

		#region [ btnConsultar ]

		#region [ btnConsultar_Click ]
		private void btnConsultar_Click(object sender, EventArgs e)
		{
			trataBotaoConsultar();
		}
		#endregion

		#endregion

		#region [ btnFiltrar ]

		#region [ btnFiltrar_Click ]
		private void btnFiltrar_Click(object sender, EventArgs e)
		{
			trataBotaoFiltrar();
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

		#region [ btnImprimirCompleto ]

		#region [ btnImprimirCompleto_Click ]
		private void btnImprimirCompleto_Click(object sender, EventArgs e)
		{
			trataBotaoImprimirCompleto();
		}
		#endregion

		#endregion

		#region [ btnImprimirParcial ]

		#region [ btnImprimirParcial_Click ]
		private void btnImprimirParcial_Click(object sender, EventArgs e)
		{
			trataBotaoImprimirParcial();
		}
		#endregion

		#endregion

		#region [ btnImprimirVolume ]
		
		#region [ btnImprimirVolume_Click ]
		private void btnImprimirVolume_Click(object sender, EventArgs e)
		{
			trataBotaoImprimirVolume();
		}
		#endregion

		#endregion

		#region [ grdPesquisa ]

		#region [ grdPesquisa_SortCompare ]
		private void grdPesquisa_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			if (e.Column.Name.Equals(GRID_PESQ_COL_DATA_HORA))
			{
				e.SortResult = System.DateTime.Compare(Global.converteDdMmYyyyHhMmSsParaDateTime(e.CellValue1.ToString()), Global.converteDdMmYyyyHhMmSsParaDateTime(e.CellValue2.ToString()));
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_PESQ_COL_NSU))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_PESQ_COL_QTDE_PEDIDOS))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_PESQ_COL_QTDE_VOLUMES))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
		}
		#endregion

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

		#region [ grdPesquisa_CellDoubleClick ]
		private void grdPesquisa_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
		{
			trataBotaoConsultar();
		}
		#endregion

		#endregion

		#region [ grdDados ]

		#region [ grdDados_SortCompare ]
		private void grdDados_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			if (e.Column.Name.Equals(GRID_DADOS_COL_SEQUENCIA))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_DADOS_COL_NUMERO_NF))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_DADOS_COL_QTDE))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
			else if (e.Column.Name.Equals(GRID_DADOS_COL_QTDE_VOLUMES))
			{
				e.SortResult = (int)Global.converteInteiro(e.CellValue1.ToString()) - (int)Global.converteInteiro(e.CellValue2.ToString());
				e.Handled = true;
			}
		}
		#endregion

		#region [ grdDados_CellContentClick ]
		private void grdDados_CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e == null) return;
			if (e.ColumnIndex == 0)
			{
				DataGridViewCheckBoxCell chkBox = (DataGridViewCheckBoxCell)this.grdDados[e.ColumnIndex, e.RowIndex];
				if (chkBox.EditingCellFormattedValue.ToString().ToUpper().Equals("TRUE"))
				{
					this.grdDados.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSkyBlue;
				}
				else
				{
					this.grdDados.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Empty;
				}
			}
		}
		#endregion

		#endregion

		#region [ txtNsu ]

		#region [ txtNsu_Enter ]
		private void txtNsu_Enter(object sender, EventArgs e)
		{
			txtNsu.Select(0, txtNsu.Text.Length);
		}
		#endregion

		#region [ txtNsu_Leave ]
		private void txtNsu_Leave(object sender, EventArgs e)
		{
			txtNsu.Text = Global.digitos(txtNsu.Text);
		}
		#endregion

		#region [ txtNsu_KeyPress ]
		private void txtNsu_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#region [ txtNsu_KeyDown ]
		private void txtNsu_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (txtNsu.Text.Length > 0)
				{
					trataBotaoPesquisar();
					return;
				}
			}

			Global.trataTextBoxKeyDown(sender, e, btnPesquisar);
		}
		#endregion

		#endregion

		#region [ txtNumeroNF ]

		#region [ txtNumeroNF_Enter ]
		private void txtNumeroNF_Enter(object sender, EventArgs e)
		{
			txtNumeroNF.Select(0, txtNumeroNF.Text.Length);
		}
		#endregion

		#region [ txtNumeroNF_Leave ]
		private void txtNumeroNF_Leave(object sender, EventArgs e)
		{
			txtNumeroNF.Text = Global.digitos(txtNumeroNF.Text);
		}
		#endregion

		#region [ txtNumeroNF_KeyPress ]
		private void txtNumeroNF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#region [ txtNumeroNF_KeyDown ]
		private void txtNumeroNF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtPedido);
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
			Global.trataTextBoxKeyDown(sender, e, txtSequenciaInicio);
		}
		#endregion

		#endregion

		#region [ txtSequenciaInicio ]

		#region [ txtSequenciaInicio_Enter ]
		private void txtSequenciaInicio_Enter(object sender, EventArgs e)
		{
			txtSequenciaInicio.Select(0, txtSequenciaInicio.Text.Length);
		}
		#endregion

		#region [ txtSequenciaInicio_Leave ]
		private void txtSequenciaInicio_Leave(object sender, EventArgs e)
		{
			txtSequenciaInicio.Text = Global.digitos(txtSequenciaInicio.Text);
		}
		#endregion

		#region [ txtSequenciaInicio_KeyPress ]
		private void txtSequenciaInicio_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#region [ txtSequenciaInicio_KeyDown ]
		private void txtSequenciaInicio_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtSequenciaFim);
		}
		#endregion

		#endregion

		#region [ txtSequenciaFim ]

		#region [ txtSequenciaFim_Enter ]
		private void txtSequenciaFim_Enter(object sender, EventArgs e)
		{
			txtSequenciaFim.Select(0, txtSequenciaFim.Text.Length);
		}
		#endregion

		#region [ txtSequenciaFim_Leave ]
		private void txtSequenciaFim_Leave(object sender, EventArgs e)
		{
			txtSequenciaFim.Text = Global.digitos(txtSequenciaFim.Text);
		}
		#endregion

		#region [ txtSequenciaFim_KeyPress ]
		private void txtSequenciaFim_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#region [ txtSequenciaFim_KeyDown ]
		private void txtSequenciaFim_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtZona);
		}
		#endregion

		#endregion

		#region [ txtZona ]

		#region [ txtZona_Enter ]
		private void txtZona_Enter(object sender, EventArgs e)
		{
			txtZona.Select(0, txtZona.Text.Length);
		}
		#endregion

		#region [ txtZona_Leave ]
		private void txtZona_Leave(object sender, EventArgs e)
		{
			txtZona.Text = txtZona.Text.Trim().ToUpper();
		}
		#endregion

		#region [ txtZona_KeyPress ]
		private void txtZona_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoSomenteLetras(e.KeyChar);
			if (e.KeyChar != '\0') e.KeyChar = e.KeyChar.ToString().ToUpper()[0];
		}
		#endregion

		#region [ txtZona_KeyDown ]
		private void txtZona_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnFiltrar);
		}
		#endregion

		#endregion

		#endregion
	}
	#endregion

	#region [ Declaração de Classes Auxiliares ]

	#region [ EtiquetaDados ]
	public class EtiquetaDados
	{
		#region [ Getters/Setters ]
		private int _id_N1;
		public int id_N1
		{
			get { return _id_N1; }
			set { _id_N1 = value; }
		}

		private int _id_N2;
		public int id_N2
		{
			get { return _id_N2; }
			set { _id_N2 = value; }
		}

		private int _id_N3;
		public int id_N3
		{
			get { return _id_N3; }
			set { _id_N3 = value; }
		}

		private string _pedido;
		public string pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private string _obs_2;
		public string obs_2
		{
			get { return _obs_2; }
			set { _obs_2 = value; }
		}

		private string _obs_3;
		public string obs_3
		{
			get { return _obs_3; }
			set { _obs_3 = value; }
		}

		private string _loja;
		public string loja
		{
			get { return _loja; }
			set { _loja = value; }
		}

		private string _transportadora_id;
		public string transportadora_id
		{
			get { return _transportadora_id; }
			set { _transportadora_id = value; }
		}

		private string _id_cliente;
		public string id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private string _cnpj_cpf_cliente;
		public string cnpj_cpf_cliente
		{
			get { return _cnpj_cpf_cliente; }
			set { _cnpj_cpf_cliente = value; }
		}

		private string _nome_cliente;
		public string nome_cliente
		{
			get { return _nome_cliente; }
			set { _nome_cliente = value; }
		}

		private string _nome_fabricante;
		public string nome_fabricante
		{
			get { return _nome_fabricante; }
			set { _nome_fabricante = value; }
		}

		private string _razao_social_fabricante;
		public string razao_social_fabricante
		{
			get { return _razao_social_fabricante; }
			set { _razao_social_fabricante = value; }
		}

		private string _fabricante;
		public string fabricante
		{
			get { return _fabricante; }
			set { _fabricante = value; }
		}

		private string _produto;
		public string produto
		{
			get { return _produto; }
			set { _produto = value; }
		}

		private int _qtde;
		public int qtde
		{
			get { return _qtde; }
			set { _qtde = value; }
		}

		private int _qtde_volumes;
		public int qtde_volumes
		{
			get { return _qtde_volumes; }
			set { _qtde_volumes = value; }
		}

		private string _descricao;
		public string descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private string _descricao_html;
		public string descricao_html
		{
			get { return _descricao_html; }
			set { _descricao_html = value; }
		}

		private int _zona_id;
		public int zona_id
		{
			get { return _zona_id; }
			set { _zona_id = value; }
		}

		private string _zona_codigo;
		public string zona_codigo
		{
			get { return _zona_codigo; }
			set { _zona_codigo = value; }
		}

		private string _numeroNFe;
		public string numeroNFe
		{
			get { return _numeroNFe; }
			set { _numeroNFe = value; }
		}

		private int _qtde_volumes_pedido;
		public int qtde_volumes_pedido
		{
			get { return _qtde_volumes_pedido; }
			set { _qtde_volumes_pedido = value; }
		}

		private int _sequencia;
		public int sequencia
		{
			get { return _sequencia; }
			set { _sequencia = value; }
		}

		private string _st_entrega;
		public string st_entrega
		{
			get { return _st_entrega; }
			set { _st_entrega = value; }
		}

		private string _destino_tipo_endereco;
		public string destino_tipo_endereco
		{
			get { return _destino_tipo_endereco; }
			set { _destino_tipo_endereco = value; }
		}

		private string _destino_endereco;
		public string destino_endereco
		{
			get { return _destino_endereco; }
			set { _destino_endereco = value; }
		}

		private string _destino_endereco_numero;
		public string destino_endereco_numero
		{
			get { return _destino_endereco_numero; }
			set { _destino_endereco_numero = value; }
		}

		private string _destino_endereco_complemento;
		public string destino_endereco_complemento
		{
			get { return _destino_endereco_complemento; }
			set { _destino_endereco_complemento = value; }
		}

		private string _destino_bairro;
		public string destino_bairro
		{
			get { return _destino_bairro; }
			set { _destino_bairro = value; }
		}

		private string _destino_cidade;
		public string destino_cidade
		{
			get { return _destino_cidade; }
			set { _destino_cidade = value; }
		}

		private string _destino_uf;
		public string destino_uf
		{
			get { return _destino_uf; }
			set { _destino_uf = value; }
		}

		private string _destino_cep;
		public string destino_cep
		{
			get { return _destino_cep; }
			set { _destino_cep = value; }
		}

		private string _origem_uf;
		public string origem_uf
		{
			get { return _origem_uf; }
			set { _origem_uf = value; }
		}

		private List<CtrlNumeracaoVolume> _ctrlNumeracaoVolume;
		public List<CtrlNumeracaoVolume> ctrlNumeracaoVolume
		{
			get { return _ctrlNumeracaoVolume; }
			set { _ctrlNumeracaoVolume = value; }
		}
		#endregion

		#region [ Construtor ]
		public EtiquetaDados()
		{
			_id_N1 = 0;
			_id_N2 = 0;
			_id_N3 = 0;
			_pedido = "";
			_obs_2 = "";
			_obs_3 = "";
			_loja = "";
			_transportadora_id = "";
			_id_cliente = "";
			_cnpj_cpf_cliente = "";
			_nome_cliente = "";
			_nome_fabricante = "";
			_razao_social_fabricante = "";
			_fabricante = "";
			_produto = "";
			_qtde = 0;
			_qtde_volumes = 0;
			_descricao = "";
			_descricao_html = "";
			_zona_id = 0;
			_zona_codigo = "";
			_numeroNFe = "";
			_qtde_volumes_pedido = 0;
			_sequencia = 0;
			_st_entrega = "";
			_destino_tipo_endereco = "";
			_destino_endereco = "";
			_destino_endereco_numero = "";
			_destino_endereco_complemento = "";
			_destino_bairro = "";
			_destino_cidade = "";
			_destino_uf = "";
			_destino_cep = "";
			_ctrlNumeracaoVolume = new List<CtrlNumeracaoVolume>();
		}
		#endregion
	}
	#endregion

	#region [ CtrlNumeracaoVolume ]
	public class CtrlNumeracaoVolume
	{
		#region [ Getters/Setters ]
		private int _numeracaoVolume;
		public int numeracaoVolume
		{
			get { return _numeracaoVolume; }
			set { _numeracaoVolume = value; }
		}

		private string _identificacaoVolumeProdutoComposto;
		public string identificacaoVolumeProdutoComposto
		{
			get { return _identificacaoVolumeProdutoComposto; }
			set { _identificacaoVolumeProdutoComposto = value; }
		}
		#endregion

		#region [ Construtor ]
		public CtrlNumeracaoVolume()
		{
			_identificacaoVolumeProdutoComposto = "";
		}
		#endregion
	}
	#endregion

	#region [ EtiquetaCtrlNumeracaoVolumes ]
	public class EtiquetaCtrlNumeracaoVolumes
	{
		#region [ Getters/Setters ]
		private string _pedido;
		public string pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private int _qtde_volumes_pedido;
		public int qtde_volumes_pedido
		{
			get { return _qtde_volumes_pedido; }
			set { _qtde_volumes_pedido = value; }
		}

		private int _contador_numeracao_volume_pedido;
		public int contador_numeracao_volume_pedido
		{
			get { return _contador_numeracao_volume_pedido; }
			set { _contador_numeracao_volume_pedido = value; }
		}
		#endregion

		#region [ Construtor ]
		public EtiquetaCtrlNumeracaoVolumes()
		{
			_pedido = "";
			_qtde_volumes_pedido = 0;
			_contador_numeracao_volume_pedido = 0;
		}
		#endregion
	}
	#endregion

	#region [ RawPrinterHelper ]
	public class RawPrinterHelper
	{
		#region [ Structure and API declarions ]
		// Structure and API declarions:
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
		public class DOCINFOA
		{
			[MarshalAs(UnmanagedType.LPStr)]
			public string pDocName;
			[MarshalAs(UnmanagedType.LPStr)]
			public string pOutputFile;
			[MarshalAs(UnmanagedType.LPStr)]
			public string pDataType;
		}
		[DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

		[DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool ClosePrinter(IntPtr hPrinter);

		[DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

		[DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool EndDocPrinter(IntPtr hPrinter);

		[DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool StartPagePrinter(IntPtr hPrinter);

		[DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool EndPagePrinter(IntPtr hPrinter);

		[DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
		public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);
		#endregion

		#region [ SendBytesToPrinter ]
		// SendBytesToPrinter()
		// When the function is given a printer name and an unmanaged array
		// of bytes, the function sends those bytes to the print queue.
		// Returns true on success, false on failure.
		public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
		{
			Int32 dwError = 0, dwWritten = 0;
			IntPtr hPrinter = new IntPtr(0);
			DOCINFOA di = new DOCINFOA();
			bool bSuccess = false; // Assume failure unless you specifically succeed.

			di.pDocName = "Etiqueta-WMS";
			di.pDataType = "RAW";

			// Open the printer.
			if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
			{
				// Start a document.
				if (StartDocPrinter(hPrinter, 1, di))
				{
					// Start a page.
					if (StartPagePrinter(hPrinter))
					{
						// Write your bytes.
						bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
						EndPagePrinter(hPrinter);
					}
					EndDocPrinter(hPrinter);
				}
				ClosePrinter(hPrinter);
			}
			// If you did not succeed, GetLastError may give more information
			// about why not.
			if (bSuccess == false)
			{
				dwError = Marshal.GetLastWin32Error();
			}
			return bSuccess;
		}
		#endregion

		#region [ SendFileToPrinter ]
		public static bool SendFileToPrinter(string szPrinterName, string szFileName)
		{
			// Open the file.
			FileStream fs = new FileStream(szFileName, FileMode.Open);
			try
			{
				// Create a BinaryReader on the file.
				BinaryReader br = new BinaryReader(fs);
				// Dim an array of bytes big enough to hold the file's contents.
				Byte[] bytes = new Byte[fs.Length];
				bool bSuccess = false;
				// Your unmanaged pointer.
				IntPtr pUnmanagedBytes = new IntPtr(0);
				int nLength;

				nLength = Convert.ToInt32(fs.Length);
				// Read the contents of the file into the array.
				bytes = br.ReadBytes(nLength);
				// Allocate some unmanaged memory for those bytes.
				pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
				// Copy the managed byte array into the unmanaged array.
				Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
				// Send the unmanaged bytes to the printer.
				bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
				// Free the unmanaged memory that you allocated earlier.
				Marshal.FreeCoTaskMem(pUnmanagedBytes);
				return bSuccess;
			}
			finally
			{
				fs.Close();
			}
		}
		#endregion

		#region [ SendStringToPrinter ]
		public static bool SendStringToPrinter(string szPrinterName, string szString)
		{
			IntPtr pBytes;
			Int32 dwCount;
			// How many characters are in the string?
			dwCount = szString.Length;
			// Assume that the printer is expecting ANSI text, and then convert
			// the string to ANSI text.
			pBytes = Marshal.StringToCoTaskMemAnsi(szString);
			// Send the converted ANSI string to the printer.
			SendBytesToPrinter(szPrinterName, pBytes, dwCount);
			Marshal.FreeCoTaskMem(pBytes);
			return true;
		}
		#endregion
	}
	#endregion

	#endregion
}
