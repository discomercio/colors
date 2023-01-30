#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace ADM2
{
	public class PedidoFreteDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
		private SqlCommand cmInsertPedidoFreteViaCsv;
		#endregion

		#region [ inicializaConstrutorEstatico ]
		public static void inicializaConstrutorEstatico()
		{
			// NOP
			// 1) The static constructor for a class executes before any instance of the class is created.
			// 2) The static constructor for a class executes before any of the static members for the class are referenced.
			// 3) The static constructor for a class executes after the static field initializers (if any) for the class.
			// 4) The static constructor for a class executes at most one time during a single program instantiation
			// 5) A static constructor does not take access modifiers or have parameters.
			// 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
			// 7) A static constructor cannot be called directly.
			// 8) The user has no control on when the static constructor is executed in the program.
			// 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
		}
		#endregion

		#region [ Construtor ]
		public PedidoFreteDAO(ref BancoDados bd)
		{
			_bd = bd;
			inicializaObjetos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetos ]
		public void inicializaObjetos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmInsertPedidoFreteViaCsv ]
			strSql = "INSERT INTO t_PEDIDO_FRETE (" +
						"id" +
						", pedido" +
						", codigo_tipo_frete" +
						", vl_frete" +
						", transportadora_id" +
						", transportadora_cnpj" +
						", id_nfe_emitente" +
						", serie_NF" +
						", numero_NF" +
						", tipo_preenchimento" +
						", usuario_cadastro" +
						", usuario_ult_atualizacao" +
						", vl_NF" +
						", emissor_cnpj" +
						", id_editrp_arq_input_linha_processada_n1" +
					") VALUES (" +
						"@id" +
						", @pedido" +
						", @codigo_tipo_frete" +
						", @vl_frete" +
						", @transportadora_id" +
						", @transportadora_cnpj" +
						", @id_nfe_emitente" +
						", @serie_NF" +
						", @numero_NF" +
						", @tipo_preenchimento" +
						", @usuario_cadastro" +
						", @usuario_ult_atualizacao" +
						", @vl_NF" +
						", @emissor_cnpj" +
						", @id_editrp_arq_input_linha_processada_n1" +
					")";
			cmInsertPedidoFreteViaCsv = _bd.criaSqlCommand();
			cmInsertPedidoFreteViaCsv.CommandText = strSql;
			cmInsertPedidoFreteViaCsv.Parameters.Add("@id", SqlDbType.Int);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@codigo_tipo_frete", SqlDbType.VarChar, 3);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@vl_frete", SqlDbType.Money);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@transportadora_id", SqlDbType.VarChar, 10);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@transportadora_cnpj", SqlDbType.VarChar, 14);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@id_nfe_emitente", SqlDbType.SmallInt);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@serie_NF", SqlDbType.Int);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@numero_NF", SqlDbType.Int);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@tipo_preenchimento", SqlDbType.SmallInt);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@vl_NF", SqlDbType.Money);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@emissor_cnpj", SqlDbType.VarChar, 14);
			cmInsertPedidoFreteViaCsv.Parameters.Add("@id_editrp_arq_input_linha_processada_n1", SqlDbType.Int);
			cmInsertPedidoFreteViaCsv.Prepare();
			#endregion
		}
		#endregion

		#region [ PedidoFreteLoadFromDataRow ]
		public PedidoFrete PedidoFreteLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			PedidoFrete frete = new PedidoFrete();
			#endregion

			frete.id = BD.readToInt(rowDados["id"]);
			frete.pedido = BD.readToString(rowDados["pedido"]);
			frete.codigo_tipo_frete = BD.readToString(rowDados["codigo_tipo_frete"]);
			frete.descricao_tipo_frete = BD.readToString(rowDados["descricao_tipo_frete"]);
			frete.vl_frete = BD.readToDecimal(rowDados["vl_frete"]);
			frete.vl_frete_original_EDI = BD.readToDecimal(rowDados["vl_frete_original_EDI"]);
			frete.vl_NF = BD.readToDecimal(rowDados["vl_NF"]);
			frete.transportadora_id = BD.readToString(rowDados["transportadora_id"]);
			frete.transportadora_cnpj = BD.readToString(rowDados["transportadora_cnpj"]);
			frete.emissor_cnpj = BD.readToString(rowDados["emissor_cnpj"]);
			frete.id_nfe_emitente = BD.readToInt(rowDados["id_nfe_emitente"]);
			frete.serie_NF = BD.readToInt(rowDados["serie_NF"]);
			frete.numero_NF = BD.readToInt(rowDados["numero_NF"]);
			frete.tipo_preenchimento = BD.readToInt(rowDados["tipo_preenchimento"]);
			frete.id_editrp_arq_input_linha_processada_n1 = BD.readToInt(rowDados["id_editrp_arq_input_linha_processada_n1"]);
			frete.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			frete.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			frete.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			frete.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			frete.dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["dt_hr_ult_atualizacao"]);
			frete.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);

			return frete;
		}
		#endregion

		#region [ getPedidoFrete ]
		public List<PedidoFrete> getPedidoFrete(string numeroPedido, out string msg_erro)
		{
			#region [ Declarações ]
			String strSql;
			List<PedidoFrete> listaPedidoFrete = new List<PedidoFrete>();
			PedidoFrete frete;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";

			#region [ Consistências ]
			if (numeroPedido == null) throw new Exception("Nº do pedido não foi fornecido!!");
			if (numeroPedido.Length == 0) throw new Exception("Nº do pedido não foi informado!!");
			#endregion

			#region [ Inicialização ]
			numeroPedido = numeroPedido.Trim();
			#endregion

			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta Select ]
				strSql = "SELECT " +
							"t_PEDIDO_FRETE.*" +
							", t_CODIGO_DESCRICAO.descricao AS descricao_tipo_frete" +
						" FROM t_PEDIDO_FRETE" +
							" LEFT JOIN t_CODIGO_DESCRICAO ON (t_CODIGO_DESCRICAO.grupo = '" + Global.Cte.GruposCodigoDescricao.ID_GRUPO__PEDIDO_TIPO_FRETE + "') AND (t_CODIGO_DESCRICAO.codigo = t_PEDIDO_FRETE.codigo_tipo_frete)" +
						" WHERE" +
							" (t_PEDIDO_FRETE.pedido = '" + numeroPedido + "')" +
						" ORDER BY" +
							" t_PEDIDO_FRETE.id";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return listaPedidoFrete;

				foreach (DataRow row in dtbResultado.Rows)
				{
					frete = PedidoFreteLoadFromDataRow(row);
					listaPedidoFrete.Add(frete);
				}

				return listaPedidoFrete;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ InsertPedidoFreteViaCsv ]
		public bool InsertPedidoFreteViaCsv(PedidoFrete pedidoFrete, string usuario, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoFreteDAO.InsertPedidoFreteViaCsv()";
			string strMsg;
			string strMsgErroParam;
			string strMsgErroLog = "";
			int intNsuNovoPedidoFrete;
			int intRetorno;
			Log log = new Log();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if (pedidoFrete == null)
				{
					msg_erro = "Dados do frete do pedido não foram informados!";
					return false;
				}

				if ((pedidoFrete.pedido ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o número do pedido para registrar o frete!";
					return false;
				}

				if (pedidoFrete.vl_frete <= 0)
				{
					msg_erro = "Valor de frete inválido!";
					return false;
				}

				if ((usuario ?? "").Length == 0)
				{
					msg_erro = "Usuário responsável pela atualização não foi informado!";
					return false;
				}
				#endregion

				#region [ Já foi definido o ID para o novo registro? ]
				if (pedidoFrete.id == 0)
				{
					if (!_bd.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.ID_T_FIN_CONTROLE.T_PEDIDO_FRETE, out intNsuNovoPedidoFrete, out strMsgErroParam))
					{
						msg_erro = "Falha ao tentar gerar o ID para gravar o registro de frete no pedido " + pedidoFrete.pedido + " (NF: " + pedidoFrete.numero_NF.ToString() + ")";
						return false;
					}
					pedidoFrete.id = intNsuNovoPedidoFrete;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertPedidoFreteViaCsv.Parameters["@id"].Value = pedidoFrete.id;
				cmInsertPedidoFreteViaCsv.Parameters["@pedido"].Value = pedidoFrete.pedido;
				cmInsertPedidoFreteViaCsv.Parameters["@codigo_tipo_frete"].Value = pedidoFrete.codigo_tipo_frete;
				cmInsertPedidoFreteViaCsv.Parameters["@vl_frete"].Value = pedidoFrete.vl_frete;
				cmInsertPedidoFreteViaCsv.Parameters["@transportadora_id"].Value = pedidoFrete.transportadora_id;
				cmInsertPedidoFreteViaCsv.Parameters["@transportadora_cnpj"].Value = pedidoFrete.transportadora_cnpj;
				cmInsertPedidoFreteViaCsv.Parameters["@id_nfe_emitente"].Value = pedidoFrete.id_nfe_emitente;
				cmInsertPedidoFreteViaCsv.Parameters["@serie_NF"].Value = pedidoFrete.serie_NF;
				cmInsertPedidoFreteViaCsv.Parameters["@numero_NF"].Value = pedidoFrete.numero_NF;
				cmInsertPedidoFreteViaCsv.Parameters["@tipo_preenchimento"].Value = pedidoFrete.tipo_preenchimento;
				cmInsertPedidoFreteViaCsv.Parameters["@usuario_cadastro"].Value = usuario;
				cmInsertPedidoFreteViaCsv.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				cmInsertPedidoFreteViaCsv.Parameters["@vl_NF"].Value = pedidoFrete.vl_NF;
				cmInsertPedidoFreteViaCsv.Parameters["@emissor_cnpj"].Value = pedidoFrete.emissor_cnpj;
				cmInsertPedidoFreteViaCsv.Parameters["@id_editrp_arq_input_linha_processada_n1"].Value = pedidoFrete.id_editrp_arq_input_linha_processada_n1;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = _bd.executaNonQuery(ref cmInsertPedidoFreteViaCsv);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;
					strMsg = NOME_DESTA_ROTINA + " - Exception ao tentar registrar o frete no pedido " + pedidoFrete.pedido + ": numero_NF=" + pedidoFrete.numero_NF.ToString() + ", vl_frete=" + Global.formataMoeda(pedidoFrete.vl_frete) + ", codigo_tipo_frete=" + pedidoFrete.codigo_tipo_frete + ", transportadora_id=" + pedidoFrete.transportadora_id + ", usuario_cadastro=" + usuario + "\r\n" + ex.ToString();
					Global.gravaLogAtividade(strMsg);
					return false;
				}
				#endregion

				if (intRetorno == 0)
				{
					msg_erro = NOME_DESTA_ROTINA + " - Falha ao tentar registrar o frete no pedido " + pedidoFrete.pedido + ": numero_NF=" + pedidoFrete.numero_NF.ToString() + ", vl_frete=" + Global.formataMoeda(pedidoFrete.vl_frete) + ", codigo_tipo_frete=" + pedidoFrete.codigo_tipo_frete + ", transportadora_id=" + pedidoFrete.transportadora_id + ", usuario_cadastro=" + usuario;
					return false;
				}

				#region [ Registra log ]
				strMsg = "[Módulo ADM2] Anotação de frete via CSV no pedido " + pedidoFrete.pedido + ": vl_frete=" + Global.formataMoeda(pedidoFrete.vl_frete) + ", NF=" + pedidoFrete.numero_NF.ToString() + ", codigo_tipo_frete=" + pedidoFrete.codigo_tipo_frete + ", transportadora_id=" + pedidoFrete.transportadora_id;
				log.usuario = usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_ANOTA_FRETE_PEDIDO;
				log.pedido = pedidoFrete.pedido;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				strMsg = NOME_DESTA_ROTINA + " - Sucesso na gravação do frete no pedido " + pedidoFrete.pedido + ": vl_frete=" + Global.formataMoeda(pedidoFrete.vl_frete) + ", NF=" + pedidoFrete.numero_NF.ToString() + ", codigo_tipo_frete=" + pedidoFrete.codigo_tipo_frete + ", transportadora_id=" + pedidoFrete.transportadora_id + ", usuario_cadastro=" + usuario;
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return false;
			}
		}
		#endregion

		#endregion
	}
}
