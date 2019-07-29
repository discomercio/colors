#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class BoletoPreCadastradoDAO
	{
		#region [ obtemBoletoPlanoContasDestino ]
		/// <summary>
		/// Consulta os pedidos relacionados na tabela de rateio e localiza as lojas às quais 
		/// os pedidos pertencem.
		/// Para cada loja, obtém o plano de contas para o qual deverá ser lançado o lançamento 
		/// do fluxo de caixa gerado em decorrência do boleto.
		/// Gera uma exceção no caso de não encontrar nenhum plano de contas ou se houver mais
		/// do que 1 plano de contas.
		/// </summary>
		/// <param name="id_nf_parcela_pagto">Nº identificação do registro principal em t_FIN_NF_PARCELA_PAGTO</param>
		/// <returns>Retorna um objeto do tipo BoletoPlanoContasDestino com os dados do plano de contas</returns>
		public static BoletoPlanoContasDestino obtemBoletoPlanoContasDestino(int id_nf_parcela_pagto)
		{
			#region [ Declarações ]
			BoletoPlanoContasDestino boletoPlanoContasDestino = new BoletoPlanoContasDestino();
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Consulta SQL ]
			strSql = "SELECT DISTINCT" +
						" id_plano_contas_empresa," +
						" id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" natureza" +
					" FROM t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tPPIR" +
						" INNER JOIN t_PEDIDO tP ON (tPPIR.pedido=tP.pedido)" +
						" INNER JOIN t_LOJA tL ON (tP.loja=tL.loja)" +
					" WHERE" +
						" (id_nf_parcela_pagto = " + id_nf_parcela_pagto.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Consistência ]
			if (dtbResultado.Rows.Count == 0)
			{
				throw new FinanceiroException("Não há informações do plano de contas (id_nf_parcela_pagto=" + id_nf_parcela_pagto.ToString() + ")");
			}
			else if (dtbResultado.Rows.Count > 1)
			{
				throw new FinanceiroException("Há mais de 1 plano de contas (id_nf_parcela_pagto=" + id_nf_parcela_pagto.ToString() + ")");
			}
			rowResultado = dtbResultado.Rows[0];
			if (BD.readToInt(rowResultado["id_plano_contas_conta"]) == 0)
			{
				throw new FinanceiroException("A informação do plano de contas não foi preenchida adequadamente no cadastro de lojas (id_nf_parcela_pagto=" + id_nf_parcela_pagto.ToString() + ")!!");
			}
			#endregion

			#region [ Carrega os dados ]
			boletoPlanoContasDestino.id_plano_contas_empresa = BD.readToByte(rowResultado["id_plano_contas_empresa"]);
			boletoPlanoContasDestino.id_plano_contas_grupo = BD.readToShort(rowResultado["id_plano_contas_grupo"]);
			boletoPlanoContasDestino.id_plano_contas_conta = BD.readToInt(rowResultado["id_plano_contas_conta"]);
			boletoPlanoContasDestino.natureza = BD.readToChar(rowResultado["natureza"]);
			#endregion

			return boletoPlanoContasDestino;
		}
		#endregion

		#region [ getBoletoPreCadastrado ]
		/// <summary>
		/// Retorna um objeto BoletoPreCadastrado contendo os dados lidos do BD
		/// </summary>
		/// <param name="id">
		/// Identificação do registro
		/// </param>
		/// <returns>
		/// Retorna um objeto BoletoPreCadastrado contendo os dados lidos do BD
		/// </returns>
		public static BoletoPreCadastrado getBoletoPreCadastrado(int id)
		{
			#region [ Declarações ]
			String strSql;
			String strWhere = "";
			BoletoPreCadastrado boleto = new BoletoPreCadastrado();
			BoletoPreCadastradoItem boletoItem;
			BoletoPreCadastradoItemRateio boletoRateio;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistência ]
			if (id == 0) throw new FinanceiroException("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do boleto ]
			strWhere = " (id = " + id.ToString() + ")";
			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_NF_PARCELA_PAGTO" +
					strWhere;
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Registro id=" + id.ToString() + " não localizado na tabela t_FIN_NF_PARCELA_PAGTO!!");

			rowResultado = dtbResultado.Rows[0];

			boleto.id = (int)rowResultado["id"];
			boleto.id_cliente = rowResultado["id_cliente"].ToString();
			boleto.numero_NF = (int)rowResultado["numero_NF"];
			boleto.qtde_parcelas = (byte)rowResultado["qtde_parcelas"];
			boleto.qtde_parcelas_boleto = (byte)rowResultado["qtde_parcelas_boleto"];
			boleto.status = (byte)rowResultado["status"];
			boleto.dt_cadastro = (DateTime)rowResultado["dt_cadastro"];
			boleto.usuario_cadastro = rowResultado["usuario_cadastro"].ToString();
			boleto.dt_ult_atualizacao = (DateTime)rowResultado["dt_ult_atualizacao"];
			boleto.usuario_ult_atualizacao = rowResultado["usuario_ult_atualizacao"].ToString();
			#endregion

			#region [ Dados das parcelas ]
			boleto.listaItem = new List<BoletoPreCadastradoItem>();
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_NF_PARCELA_PAGTO_ITEM" +
					" WHERE" +
						" (id_nf_parcela_pagto = " + id.ToString() + ")" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				boletoItem = new BoletoPreCadastradoItem();
				rowResultado = dtbResultado.Rows[i];
				boletoItem.id = (int)rowResultado["id"];
				boletoItem.id_nf_parcela_pagto = (int)rowResultado["id_nf_parcela_pagto"];
				boletoItem.num_parcela = (byte)rowResultado["num_parcela"];
				boletoItem.forma_pagto = (short)rowResultado["forma_pagto"];
				boletoItem.dt_vencto = (DateTime)rowResultado["dt_vencto"];
				boletoItem.valor = (decimal)rowResultado["valor"];
				boleto.listaItem.Add(boletoItem);
			}
			#endregion

			#region [ Dados do rateio de pedidos ]
			foreach (BoletoPreCadastradoItem item in boleto.listaItem)
			{
				item.listaRateio = new List<BoletoPreCadastradoItemRateio>();
				strSql = "SELECT " +
							"*" +
						" FROM t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO" +
						" WHERE" +
							" (id_nf_parcela_pagto_item = " + item.id.ToString() + ")" +
						" ORDER BY" +
							" pedido";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					boletoRateio = new BoletoPreCadastradoItemRateio();
					rowResultado = dtbResultado.Rows[i];
					boletoRateio.id_nf_parcela_pagto_item = (int)rowResultado["id_nf_parcela_pagto_item"];
					boletoRateio.pedido = rowResultado["pedido"].ToString();
					boletoRateio.id_nf_parcela_pagto = (int)rowResultado["id_nf_parcela_pagto"];
					boletoRateio.valor = (decimal)rowResultado["valor"];
					item.listaRateio.Add(boletoRateio);
				}
			}
			#endregion

			return boleto;
		}
		#endregion

		#region [ anula ]
		/// <summary>
		/// Anula o registro em t_FIN_NF_PARCELA_PAGTO, que contém o registro principal dos dados
		/// das parcelas de pagamento gerados na emissão da NF e que são usados como base p/ o
		/// cadastramento de boletos.
		/// </summary>
		/// <param name="usuario">
		/// Identificação do usuário que está executando a operação de anular
		/// </param>
		/// <param name="intIdSelecionado">
		/// Identificação do registro a ser anulado
		/// </param>
		/// <param name="strDescricaoLog">
		/// Retorna a mensagem para o log
		/// </param>
		/// <param name="strMsgErro">
		/// Retorna mensagem em caso de ocorrer erro
		/// </param>
		/// <returns>
		/// true: registro foi anulado
		/// false: falha ao tentar anular o registro
		/// </returns>
		public static bool anula( String usuario,
								  int intIdSelecionado,
								  ref String strDescricaoLog,
								  ref String strMsgErro)
		{
			#region [ Declarações ]
			SqlCommand cmCommand;
			String strSql;
			#endregion

			try
			{
				#region [ Inicialização ]
				strDescricaoLog = "";
				strMsgErro = "";
				#endregion

				#region [ Prepara acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				#endregion

				#region [ Anula o registro ]
				strSql = "UPDATE" +
							" t_FIN_NF_PARCELA_PAGTO" +
						" SET" +
							" status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.CANCELADO.ToString() + ", " +
							" dt_ult_atualizacao = getdate(), " +
							" usuario_ult_atualizacao = '" + Global.Usuario.usuario + "'" +
						" WHERE" +
							" (id = " + intIdSelecionado.ToString() + ")" +
							" AND (status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")";

				cmCommand.CommandText = strSql;
				if (BD.executaNonQuery(ref cmCommand) == 1)
				{
					strDescricaoLog = "Registro id=" + intIdSelecionado.ToString() + " anulado com sucesso pelo usuário " + usuario;
					return true;
				}
				else
				{
					strMsgErro = "Falha ao anular o registro id=" + intIdSelecionado.ToString() + ": registro não encontrado ou status inválido!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = "Falha ao anular o registro id=" + intIdSelecionado.ToString() +
							 "\n" +
							 ex.Message;
				return false;
			}
		}
		#endregion

		#region [ marcaComoTratado ]
		/// <summary>
		/// Marca como já tratado o registro em t_FIN_NF_PARCELA_PAGTO, que contém o registro principal
		/// dos dados das parcelas de pagamento gerados na emissão da NF e que são usados como base p/ o
		/// cadastramento de boletos.
		/// O registro de t_FIN_NF_PARCELA_PAGTO deve ser marcado como tratado quando ele foi usado como
		/// base p/ o cadastramento de um boleto e deve ser marcado como anulado quando for descartado
		/// sem gerar nenhum boleto.
		/// </summary>
		/// <param name="usuario">
		/// Identificação do usuário que está executando a operação
		/// </param>
		/// <param name="intIdSelecionado">
		/// Identificação do registro a ser marcado
		/// </param>
		/// <param name="strDescricaoLog">
		/// Retorna a mensagem para o log
		/// </param>
		/// <param name="strMsgErro">
		/// Retorna mensagem em caso de ocorrer erro
		/// </param>
		/// <returns>
		/// true: registro foi alterado com sucesso
		/// false: falha ao tentar alterar o registro
		/// </returns>
		public static bool marcaComoTratado(String usuario,
											int intIdSelecionado,
											ref String strDescricaoLog,
											ref String strMsgErro)
		{
			#region [ Declarações ]
			String strNomeTabela = "t_FIN_NF_PARCELA_PAGTO";
			SqlCommand cmCommand;
			String strSql;
			#endregion

			try
			{
				#region [ Inicialização ]
				strDescricaoLog = "";
				strMsgErro = "";
				#endregion

				#region [ Prepara acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				#endregion

				#region [ Altera o registro ]
				strSql = "UPDATE" +
							" t_FIN_NF_PARCELA_PAGTO" +
						" SET" +
							" status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.TRATADO.ToString() + ", " +
							" dt_ult_atualizacao = getdate(), " +
							" usuario_ult_atualizacao = '" + Global.Usuario.usuario + "'" +
						" WHERE" +
							" (id = " + intIdSelecionado.ToString() + ")" +
							" AND (status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")";

				cmCommand.CommandText = strSql;
				if (BD.executaNonQuery(ref cmCommand) == 1)
				{
					strDescricaoLog = "Registro " + strNomeTabela + ".id=" + intIdSelecionado.ToString() + " marcado como já tratado com sucesso pelo usuário " + usuario;
					return true;
				}
				else
				{
					strMsgErro = "Falha ao marcar como já tratado o registro " + strNomeTabela + ".id=" + intIdSelecionado.ToString() + ": registro não encontrado ou status inválido!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = "Falha ao marcar como já tratado o registro " + strNomeTabela + ".id=" + intIdSelecionado.ToString() +
							 "\n" +
							 ex.Message;
				return false;
			}
		}
		#endregion

		#region [ obtemListaNumeroPedidoRateio ]
		public static List<String> obtemListaNumeroPedidoRateio(int idNfParcelaPagto)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<String> listaPedido = new List<String>();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT DISTINCT" +
						" pedido" +
					" FROM t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO" +
					" WHERE" +
						" (id_nf_parcela_pagto = " + idNfParcelaPagto.ToString() + ")" +
					" ORDER BY" +
						" pedido";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				listaPedido.Add(BD.readToString(dtbResultado.Rows[i]["pedido"]));
			}

			return listaPedido;
		}
		#endregion

		#region [ restauraStatusInicial ]
		public static bool restauraStatusInicial(String usuario,
												 int id_nf_parcela_pagto,
												 ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Restaura status do registro em t_FIN_NF_PARCELA_PAGTO";
			String strSql;
			bool blnSucesso = false;
			int intRetorno;
			SqlCommand cmComando;
			#endregion

			strMsgErro = "";
			try
			{
				cmComando = BD.criaSqlCommand();

				#region [ Restaura status ]
				strSql = "UPDATE" +
							" t_FIN_NF_PARCELA_PAGTO" +
						 " SET" +
							" status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() +
						 " WHERE" +
							" (id = " + id_nf_parcela_pagto.ToString() + ")";
				cmComando.CommandText = strSql;
				intRetorno = BD.executaNonQuery(ref cmComando);
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar restaurar o status do registro em t_FIN_NF_PARCELA_PAGTO!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion
	}
}
