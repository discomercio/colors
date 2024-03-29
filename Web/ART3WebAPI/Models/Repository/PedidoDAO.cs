﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Domains;
using ART3WebAPI.Models.Entities;
using System.Data.SqlClient;
using System.Data;

namespace ART3WebAPI.Models.Repository
{
	public class PedidoDAO
	{
		#region [ pesquisaPedidosByCpfCnpj ]
		public static List<Pedido> pesquisaPedidosByCpfCnpj(string cpfCnpj)
		{
			#region [ Declarações ]
			SqlConnection cn;
			String strSql;
			List<Pedido> listaPedidos = new List<Pedido>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			Pedido pedido;
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try // finally: BD.fechaConexao(ref cn);
			{
				#region [ Monta Select ]
				strSql = "SELECT" +
							" pedido" +
						" FROM t_PEDIDO tP" +
							" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" +
						" WHERE" +
							" (tC.cnpj_cpf = '" + Global.digitos(cpfCnpj) + "')" +
						" ORDER BY" +
							" tP.data_hora DESC,"+
							" tP.pedido DESC";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					pedido = getPedido(BD.readToString(rowResultado["pedido"]));
					listaPedidos.Add(pedido);
				}
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return listaPedidos;
		}
		#endregion

		#region [ pesquisaPedidoValidoByNumPedidoMagento ]
		public static List<string> pesquisaPedidoValidoByNumPedidoMagento(string numeroPedidoMagento)
		{
			#region [ Declarações ]
			SqlConnection cn;
			String strSql;
			List<string> listaPedidos = new List<string>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try // finally: BD.fechaConexao(ref cn);
			{
				#region [ Monta Select ]
				strSql = "SELECT DISTINCT" +
							" pedido_base" +
						" FROM t_PEDIDO" +
						" WHERE" +
							" (pedido_bs_x_ac = '" + numeroPedidoMagento + "')" +
							" AND (st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
						" ORDER BY" +
							" pedido_base DESC";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					listaPedidos.Add(BD.readToString(rowResultado["pedido_base"]));
				}
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return listaPedidos;
		}
		#endregion

		#region [ getPedido ]
		/// <summary>
		/// Retorna um objeto Pedido contendo os dados lidos do BD
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido
		/// </param>
		/// <returns>
		/// Retorna um objeto Pedido contendo os dados lidos do BD
		/// </returns>
		public static Pedido getPedido(String numeroPedido)
		{
			#region [ Declarações ]
			String strSql;
			String numeroPedidoBase;
			decimal razaoValorPedidoFilhote = 0m;
			decimal vlBoletoDestePedido = 0m;
			decimal vlFormaPagtoDestePedido = 0m;
			decimal vlDiferencaArredondamento;
			decimal vlPagtoEmCartao = 0m;
			bool blnCartaoPagtoIntegral = false;
			Pedido pedido = new Pedido();
			PedidoItem pedidoItem;
			PedidoItemDevolvido pedidoItemDevolvido;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroPedido == null) throw new Exception("Nº do pedido a ser consultado não foi fornecido!!");
			if (numeroPedido.Length == 0) throw new Exception("Nº do pedido a ser consultado não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try // finally: BD.fechaConexao(ref cn);
			{
				#region [ Inicialização ]
				numeroPedido = numeroPedido.Trim();
				numeroPedidoBase = Global.retornaNumeroPedidoBase(numeroPedido);
				#endregion

				#region [ Pesquisa pedido ]

				#region [ Monta Select ]
				strSql = "SELECT" +
							" t_PEDIDO.*," +
							" t_LOJA.razao_social AS loja_razao_social," +
							" t_LOJA.nome AS loja_nome," +
							" t_USUARIO_VENDEDOR.nome AS vendedor_nome," +
							" t_USUARIO_CADASTRO.nome AS usuario_cadastro_nome," +
							" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota AS indicador_desempenho_nota" +
						" FROM t_PEDIDO" +
							" INNER JOIN t_LOJA ON (t_PEDIDO.loja=t_LOJA.loja)" +
							" INNER JOIN t_USUARIO AS t_USUARIO_VENDEDOR ON (t_PEDIDO.vendedor=t_USUARIO_VENDEDOR.usuario)" +
							" LEFT JOIN t_USUARIO AS t_USUARIO_CADASTRO ON (t_PEDIDO.usuario_cadastro=t_USUARIO_CADASTRO.usuario)" +
							" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Pedido nº " + numeroPedido + " não foi encontrado!!");

				#region [ Carrega os dados ]
				rowResultado = dtbResultado.Rows[0];
				pedido.pedido = BD.readToString(rowResultado["pedido"]);
				pedido.loja = BD.readToString(rowResultado["loja"]);
				pedido.loja_razao_social = BD.readToString(rowResultado["loja_razao_social"]);
				pedido.loja_nome = BD.readToString(rowResultado["loja_nome"]);
				pedido.data = BD.readToDateTime(rowResultado["data"]);
				pedido.hora = BD.readToString(rowResultado["hora"]);
				pedido.data_hora = BD.readToDateTime(rowResultado["data_hora"]);
				pedido.id_cliente = BD.readToString(rowResultado["id_cliente"]);
				pedido.midia = BD.readToString(rowResultado["midia"]);
				pedido.servicos = BD.readToString(rowResultado["servicos"]);
				pedido.vl_servicos = BD.readToDecimal(rowResultado["vl_servicos"]);
				pedido.vendedor = BD.readToString(rowResultado["vendedor"]);
				pedido.vendedor_nome = BD.readToString(rowResultado["vendedor_nome"]);
				pedido.usuario_cadastro = BD.readToString(rowResultado["usuario_cadastro"]);
				pedido.usuario_cadastro_nome = BD.readToString(rowResultado["usuario_cadastro_nome"]);
				pedido.st_entrega = BD.readToString(rowResultado["st_entrega"]);
				pedido.entregue_data = BD.readToDateTime(rowResultado["entregue_data"]);
				pedido.entregue_usuario = BD.readToString(rowResultado["entregue_usuario"]);
				pedido.cancelado_data = BD.readToDateTime(rowResultado["cancelado_data"]);
				pedido.cancelado_usuario = BD.readToString(rowResultado["cancelado_usuario"]);
				pedido.st_pagto = BD.readToString(rowResultado["st_pagto"]);
				pedido.st_recebido = BD.readToString(rowResultado["st_recebido"]);
				pedido.obs_1 = BD.readToString(rowResultado["obs_1"]);
				pedido.obs_2 = BD.readToString(rowResultado["obs_2"]);
				pedido.obs_3 = BD.readToString(rowResultado["obs_3"]);
				pedido.qtde_parcelas = BD.readToShort(rowResultado["qtde_parcelas"]);
				pedido.forma_pagto = BD.readToString(rowResultado["forma_pagto"]);
				pedido.vl_total_familia = BD.readToDecimal(rowResultado["vl_total_familia"]);
				pedido.vl_pago_familia = BD.readToDecimal(rowResultado["vl_pago_familia"]);
				pedido.split_status = BD.readToShort(rowResultado["split_status"]);
				pedido.split_data = BD.readToDateTime(rowResultado["split_data"]);
				pedido.split_hora = BD.readToString(rowResultado["split_hora"]);
				pedido.split_usuario = BD.readToString(rowResultado["split_usuario"]);
				pedido.a_entregar_status = BD.readToShort(rowResultado["a_entregar_status"]);
				pedido.a_entregar_data_marcada = BD.readToDateTime(rowResultado["a_entregar_data_marcada"]);
				pedido.a_entregar_data = BD.readToDateTime(rowResultado["a_entregar_data"]);
				pedido.a_entregar_hora = BD.readToString(rowResultado["a_entregar_hora"]);
				pedido.a_entregar_usuario = BD.readToString(rowResultado["a_entregar_usuario"]);
				pedido.loja_indicou = BD.readToString(rowResultado["loja_indicou"]);
				pedido.comissao_loja_indicou = BD.readToSingle(rowResultado["comissao_loja_indicou"]);
				pedido.venda_externa = BD.readToShort(rowResultado["venda_externa"]);
				pedido.vl_frete = BD.readToDecimal(rowResultado["vl_frete"]);
				pedido.transportadora_id = BD.readToString(rowResultado["transportadora_id"]);
				pedido.transportadora_data = BD.readToDateTime(rowResultado["transportadora_data"]);
				pedido.transportadora_usuario = BD.readToString(rowResultado["transportadora_usuario"]);
				pedido.analise_credito = BD.readToShort(rowResultado["analise_credito"]);
				pedido.analise_credito_data = BD.readToDateTime(rowResultado["analise_credito_data"]);
				pedido.analise_credito_usuario = BD.readToString(rowResultado["analise_credito_usuario"]);
				pedido.tipo_parcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);
				pedido.av_forma_pagto = BD.readToShort(rowResultado["av_forma_pagto"]);
				pedido.pc_qtde_parcelas = BD.readToShort(rowResultado["pc_qtde_parcelas"]);
				pedido.pc_valor_parcela = BD.readToDecimal(rowResultado["pc_valor_parcela"]);
				pedido.pce_forma_pagto_entrada = BD.readToShort(rowResultado["pce_forma_pagto_entrada"]);
				pedido.pce_forma_pagto_prestacao = BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]);
				pedido.pce_entrada_valor = BD.readToDecimal(rowResultado["pce_entrada_valor"]);
				pedido.pce_prestacao_qtde = BD.readToShort(rowResultado["pce_prestacao_qtde"]);
				pedido.pce_prestacao_valor = BD.readToDecimal(rowResultado["pce_prestacao_valor"]);
				pedido.pce_prestacao_periodo = BD.readToShort(rowResultado["pce_prestacao_periodo"]);
				pedido.pse_forma_pagto_prim_prest = BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]);
				pedido.pse_forma_pagto_demais_prest = BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]);
				pedido.pse_prim_prest_valor = BD.readToDecimal(rowResultado["pse_prim_prest_valor"]);
				pedido.pse_prim_prest_apos = BD.readToShort(rowResultado["pse_prim_prest_apos"]);
				pedido.pse_demais_prest_qtde = BD.readToShort(rowResultado["pse_demais_prest_qtde"]);
				pedido.pse_demais_prest_valor = BD.readToDecimal(rowResultado["pse_demais_prest_valor"]);
				pedido.pse_demais_prest_periodo = BD.readToShort(rowResultado["pse_demais_prest_periodo"]);
				pedido.pu_forma_pagto = BD.readToShort(rowResultado["pu_forma_pagto"]);
				pedido.pu_valor = BD.readToDecimal(rowResultado["pu_valor"]);
				pedido.pu_vencto_apos = BD.readToShort(rowResultado["pu_vencto_apos"]);
				pedido.indicador = BD.readToString(rowResultado["indicador"]);
				pedido.indicador_desempenho_nota = BD.readToString(rowResultado["indicador_desempenho_nota"]);
				pedido.vl_total_NF = BD.readToDecimal(rowResultado["vl_total_NF"]);
				pedido.vl_total_RA = BD.readToDecimal(rowResultado["vl_total_RA"]);
				pedido.perc_RT = BD.readToSingle(rowResultado["perc_RT"]);
				pedido.st_orc_virou_pedido = BD.readToShort(rowResultado["st_orc_virou_pedido"]);
				pedido.orcamento = BD.readToString(rowResultado["orcamento"]);
				pedido.orcamentista = BD.readToString(rowResultado["orcamentista"]);
				pedido.comissao_paga = BD.readToShort(rowResultado["comissao_paga"]);
				pedido.comissao_paga_ult_op = BD.readToString(rowResultado["comissao_paga_ult_op"]);
				pedido.comissao_paga_data = BD.readToDateTime(rowResultado["comissao_paga_data"]);
				pedido.comissao_paga_usuario = BD.readToString(rowResultado["comissao_paga_usuario"]);
				pedido.perc_desagio_RA = BD.readToSingle(rowResultado["perc_desagio_RA"]);
				pedido.perc_limite_RA_sem_desagio = BD.readToSingle(rowResultado["perc_limite_RA_sem_desagio"]);
				pedido.vl_total_RA_liquido = BD.readToDecimal(rowResultado["vl_total_RA_liquido"]);
				pedido.st_tem_desagio_RA = BD.readToShort(rowResultado["st_tem_desagio_RA"]);
				pedido.qtde_parcelas_desagio_RA = BD.readToShort(rowResultado["qtde_parcelas_desagio_RA"]);
				pedido.transportadora_num_coleta = BD.readToString(rowResultado["transportadora_num_coleta"]);
				pedido.transportadora_contato = BD.readToString(rowResultado["transportadora_contato"]);
				pedido.st_end_entrega = BD.readToShort(rowResultado["st_end_entrega"]);
				pedido.endEtg_endereco = BD.readToString(rowResultado["EndEtg_endereco"]);
				pedido.endEtg_endereco_numero = BD.readToString(rowResultado["EndEtg_endereco_numero"]);
				pedido.endEtg_endereco_complemento = BD.readToString(rowResultado["EndEtg_endereco_complemento"]);
				pedido.endEtg_bairro = BD.readToString(rowResultado["EndEtg_bairro"]);
				pedido.endEtg_cidade = BD.readToString(rowResultado["EndEtg_cidade"]);
				pedido.endEtg_uf = BD.readToString(rowResultado["EndEtg_uf"]);
				pedido.endEtg_cep = BD.readToString(rowResultado["EndEtg_cep"]);
				pedido.st_etg_imediata = BD.readToShort(rowResultado["st_etg_imediata"]);
				pedido.etg_imediata_data = BD.readToDateTime(rowResultado["etg_imediata_data"]);
				pedido.etg_imediata_usuario = BD.readToString(rowResultado["etg_imediata_usuario"]);
				pedido.frete_status = BD.readToShort(rowResultado["frete_status"]);
				pedido.frete_valor = BD.readToDecimal(rowResultado["frete_valor"]);
				pedido.frete_data = BD.readToDateTime(rowResultado["frete_data"]);
				pedido.frete_usuario = BD.readToString(rowResultado["frete_usuario"]);
				pedido.stBemUsoConsumo = BD.readToShort(rowResultado["StBemUsoConsumo"]);
				pedido.pedidoRecebidoStatus = BD.readToShort(rowResultado["PedidoRecebidoStatus"]);
				pedido.pedidoRecebidoData = BD.readToDateTime(rowResultado["PedidoRecebidoData"]);
				pedido.pedidoRecebidoUsuarioUltAtualiz = BD.readToString(rowResultado["PedidoRecebidoUsuarioUltAtualiz"]);
				pedido.pedidoRecebidoDtHrUltAtualiz = BD.readToDateTime(rowResultado["PedidoRecebidoDtHrUltAtualiz"]);
				pedido.instaladorInstalaStatus = BD.readToShort(rowResultado["InstaladorInstalaStatus"]);
				pedido.instaladorInstalaUsuarioUltAtualiz = BD.readToString(rowResultado["InstaladorInstalaUsuarioUltAtualiz"]);
				pedido.instaladorInstalaDtHrUltAtualiz = BD.readToDateTime(rowResultado["InstaladorInstalaDtHrUltAtualiz"]);
				pedido.custoFinancFornecTipoParcelamento = BD.readToString(rowResultado["custoFinancFornecTipoParcelamento"]);
				pedido.custoFinancFornecQtdeParcelas = BD.readToShort(rowResultado["custoFinancFornecQtdeParcelas"]);
				pedido.tamanho_num_pedido = BD.readToInt(rowResultado["tamanho_num_pedido"]);
				pedido.pedido_base = BD.readToString(rowResultado["pedido_base"]);
				pedido.st_forma_pagto_somente_cartao = BD.readToByte(rowResultado["st_forma_pagto_somente_cartao"]);
				pedido.id_nfe_emitente = BD.readToInt(rowResultado["id_nfe_emitente"]);
				pedido.st_auto_split = BD.readToByte(rowResultado["st_auto_split"]);
				pedido.pedido_bs_x_ac = BD.readToString(rowResultado["pedido_bs_x_ac"]);
				pedido.pedido_bs_x_marketplace = BD.readToString(rowResultado["pedido_bs_x_marketplace"]);
				pedido.marketplace_codigo_origem = BD.readToString(rowResultado["marketplace_codigo_origem"]);
				#endregion

				#endregion

				#region [ Pesquisa pedido-base? ]
				if (Global.isPedidoFilhote(numeroPedido))
				{
					#region [ Monta Select do pedido-base ]
					strSql = "SELECT" +
								" t_PEDIDO.*," +
								" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota AS indicador_desempenho_nota" +
							" FROM t_PEDIDO" +
								" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" +
							" WHERE" +
								" (pedido = '" + numeroPedidoBase + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					dtbResultado.Reset();
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0) throw new Exception("Pedido-base nº " + numeroPedidoBase + " não foi encontrado!!");

					#region [ Carrega os dados do pedido-base ]
					rowResultado = dtbResultado.Rows[0];
					pedido.st_pagto = BD.readToString(rowResultado["st_pagto"]);
					pedido.st_recebido = BD.readToString(rowResultado["st_recebido"]);
					pedido.vl_total_familia = BD.readToDecimal(rowResultado["vl_total_familia"]);
					pedido.vl_pago_familia = BD.readToDecimal(rowResultado["vl_pago_familia"]);
					pedido.analise_credito = BD.readToShort(rowResultado["analise_credito"]);
					pedido.analise_credito_data = BD.readToDateTime(rowResultado["analise_credito_data"]);
					pedido.analise_credito_usuario = BD.readToString(rowResultado["analise_credito_usuario"]);
					pedido.tipo_parcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);
					pedido.av_forma_pagto = BD.readToShort(rowResultado["av_forma_pagto"]);
					pedido.pc_qtde_parcelas = BD.readToShort(rowResultado["pc_qtde_parcelas"]);
					pedido.pc_valor_parcela = BD.readToDecimal(rowResultado["pc_valor_parcela"]);
					pedido.pce_forma_pagto_entrada = BD.readToShort(rowResultado["pce_forma_pagto_entrada"]);
					pedido.pce_forma_pagto_prestacao = BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]);
					pedido.pce_entrada_valor = BD.readToDecimal(rowResultado["pce_entrada_valor"]);
					pedido.pce_prestacao_qtde = BD.readToShort(rowResultado["pce_prestacao_qtde"]);
					pedido.pce_prestacao_valor = BD.readToDecimal(rowResultado["pce_prestacao_valor"]);
					pedido.pce_prestacao_periodo = BD.readToShort(rowResultado["pce_prestacao_periodo"]);
					pedido.pse_forma_pagto_prim_prest = BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]);
					pedido.pse_forma_pagto_demais_prest = BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]);
					pedido.pse_prim_prest_valor = BD.readToDecimal(rowResultado["pse_prim_prest_valor"]);
					pedido.pse_prim_prest_apos = BD.readToShort(rowResultado["pse_prim_prest_apos"]);
					pedido.pse_demais_prest_qtde = BD.readToShort(rowResultado["pse_demais_prest_qtde"]);
					pedido.pse_demais_prest_valor = BD.readToDecimal(rowResultado["pse_demais_prest_valor"]);
					pedido.pse_demais_prest_periodo = BD.readToShort(rowResultado["pse_demais_prest_periodo"]);
					pedido.pu_forma_pagto = BD.readToShort(rowResultado["pu_forma_pagto"]);
					pedido.pu_valor = BD.readToDecimal(rowResultado["pu_valor"]);
					pedido.pu_vencto_apos = BD.readToShort(rowResultado["pu_vencto_apos"]);
					pedido.custoFinancFornecTipoParcelamento = BD.readToString(rowResultado["custoFinancFornecTipoParcelamento"]);
					pedido.custoFinancFornecQtdeParcelas = BD.readToShort(rowResultado["custoFinancFornecQtdeParcelas"]);
					pedido.indicador = BD.readToString(rowResultado["indicador"]);
					pedido.indicador_desempenho_nota = BD.readToString(rowResultado["indicador_desempenho_nota"]);
					pedido.vl_total_NF = BD.readToDecimal(rowResultado["vl_total_NF"]);
					pedido.vl_total_RA = BD.readToDecimal(rowResultado["vl_total_RA"]);
					pedido.perc_RT = BD.readToSingle(rowResultado["perc_RT"]);
					pedido.st_orc_virou_pedido = BD.readToShort(rowResultado["st_orc_virou_pedido"]);
					pedido.orcamento = BD.readToString(rowResultado["orcamento"]);
					pedido.orcamentista = BD.readToString(rowResultado["orcamentista"]);
					pedido.comissao_paga = BD.readToShort(rowResultado["comissao_paga"]);
					pedido.comissao_paga_ult_op = BD.readToString(rowResultado["comissao_paga_ult_op"]);
					pedido.comissao_paga_data = BD.readToDateTime(rowResultado["comissao_paga_data"]);
					pedido.comissao_paga_usuario = BD.readToString(rowResultado["comissao_paga_usuario"]);
					pedido.perc_desagio_RA = BD.readToSingle(rowResultado["perc_desagio_RA"]);
					pedido.perc_limite_RA_sem_desagio = BD.readToSingle(rowResultado["perc_limite_RA_sem_desagio"]);
					pedido.vl_total_RA_liquido = BD.readToDecimal(rowResultado["vl_total_RA_liquido"]);
					pedido.st_tem_desagio_RA = BD.readToShort(rowResultado["st_tem_desagio_RA"]);
					pedido.qtde_parcelas_desagio_RA = BD.readToShort(rowResultado["qtde_parcelas_desagio_RA"]);
					pedido.st_forma_pagto_somente_cartao = BD.readToByte(rowResultado["st_forma_pagto_somente_cartao"]);
					pedido.pedido_bs_x_ac = BD.readToString(rowResultado["pedido_bs_x_ac"]);
					pedido.pedido_bs_x_marketplace = BD.readToString(rowResultado["pedido_bs_x_marketplace"]);
					pedido.marketplace_codigo_origem = BD.readToString(rowResultado["marketplace_codigo_origem"]);
					#endregion
				}
				#endregion

				#region [ Pesquisa itens do pedido ]

				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PEDIDO_ITEM" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')" +
						" ORDER BY" +
							" sequencia";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Itens do pedido nº " + numeroPedido + " não foram encontrados!!");

				#region [ Carrega os dados ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					pedidoItem = new PedidoItem();
					pedidoItem.pedido = BD.readToString(rowResultado["pedido"]);
					pedidoItem.fabricante = BD.readToString(rowResultado["fabricante"]);
					pedidoItem.produto = BD.readToString(rowResultado["produto"]);
					pedidoItem.qtde = BD.readToShort(rowResultado["qtde"]);
					pedidoItem.desc_dado = BD.readToSingle(rowResultado["desc_dado"]);
					pedidoItem.preco_venda = BD.readToDecimal(rowResultado["preco_venda"]);
					pedidoItem.preco_fabricante = BD.readToDecimal(rowResultado["preco_fabricante"]);
					pedidoItem.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
					pedidoItem.margem = BD.readToSingle(rowResultado["margem"]);
					pedidoItem.desc_max = BD.readToSingle(rowResultado["desc_max"]);
					pedidoItem.comissao = BD.readToSingle(rowResultado["comissao"]);
					pedidoItem.descricao = BD.readToString(rowResultado["descricao"]);
					pedidoItem.ean = BD.readToString(rowResultado["ean"]);
					pedidoItem.grupo = BD.readToString(rowResultado["grupo"]);
					pedidoItem.peso = BD.readToSingle(rowResultado["peso"]);
					pedidoItem.qtde_volumes = BD.readToShort(rowResultado["qtde_volumes"]);
					pedidoItem.abaixo_min_status = BD.readToShort(rowResultado["abaixo_min_status"]);
					pedidoItem.abaixo_min_autorizacao = BD.readToString(rowResultado["abaixo_min_autorizacao"]);
					pedidoItem.abaixo_min_autorizador = BD.readToString(rowResultado["abaixo_min_autorizador"]);
					pedidoItem.sequencia = BD.readToShort(rowResultado["sequencia"]);
					pedidoItem.markup_fabricante = BD.readToSingle(rowResultado["markup_fabricante"]);
					pedidoItem.preco_NF = BD.readToDecimal(rowResultado["preco_NF"]);
					pedidoItem.abaixo_min_superv_autorizador = BD.readToString(rowResultado["abaixo_min_superv_autorizador"]);
					pedidoItem.vl_custo2 = BD.readToDecimal(rowResultado["vl_custo2"]);
					pedidoItem.descricao_html = BD.readToString(rowResultado["descricao_html"]);
					pedidoItem.custoFinancFornecCoeficiente = BD.readToSingle(rowResultado["custoFinancFornecCoeficiente"]);
					pedidoItem.custoFinancFornecPrecoListaBase = BD.readToDecimal(rowResultado["custoFinancFornecPrecoListaBase"]);
					pedido.listaPedidoItem.Add(pedidoItem);

					pedido.vlTotalPrecoNfDestePedido += pedidoItem.qtde * pedidoItem.preco_NF;
					pedido.vlTotalPrecoVendaDestePedido += pedidoItem.qtde * pedidoItem.preco_venda;
				}
				#endregion

				#endregion

				#region [ Pesquisa itens devolvidos ]

				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')" +
						" ORDER BY" +
							" devolucao_data," +
							" devolucao_hora";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				#region [ Carrega os dados ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					rowResultado = dtbResultado.Rows[i];
					pedidoItemDevolvido = new PedidoItemDevolvido();
					pedidoItemDevolvido.id = BD.readToString(rowResultado["id"]);
					pedidoItemDevolvido.devolucao_data = BD.readToDateTime(rowResultado["devolucao_data"]);
					pedidoItemDevolvido.devolucao_hora = BD.readToString(rowResultado["devolucao_hora"]);
					pedidoItemDevolvido.devolucao_usuario = BD.readToString(rowResultado["devolucao_usuario"]);
					pedidoItemDevolvido.pedido = BD.readToString(rowResultado["pedido"]);
					pedidoItemDevolvido.fabricante = BD.readToString(rowResultado["fabricante"]);
					pedidoItemDevolvido.produto = BD.readToString(rowResultado["produto"]);
					pedidoItemDevolvido.qtde = BD.readToShort(rowResultado["qtde"]);
					pedidoItemDevolvido.desc_dado = BD.readToSingle(rowResultado["desc_dado"]);
					pedidoItemDevolvido.preco_venda = BD.readToDecimal(rowResultado["preco_venda"]);
					pedidoItemDevolvido.preco_fabricante = BD.readToDecimal(rowResultado["preco_fabricante"]);
					pedidoItemDevolvido.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
					pedidoItemDevolvido.margem = BD.readToSingle(rowResultado["margem"]);
					pedidoItemDevolvido.desc_max = BD.readToSingle(rowResultado["desc_max"]);
					pedidoItemDevolvido.comissao = BD.readToSingle(rowResultado["comissao"]);
					pedidoItemDevolvido.descricao = BD.readToString(rowResultado["descricao"]);
					pedidoItemDevolvido.ean = BD.readToString(rowResultado["ean"]);
					pedidoItemDevolvido.grupo = BD.readToString(rowResultado["grupo"]);
					pedidoItemDevolvido.peso = BD.readToSingle(rowResultado["peso"]);
					pedidoItemDevolvido.qtde_volumes = BD.readToShort(rowResultado["qtde_volumes"]);
					pedidoItemDevolvido.abaixo_min_status = BD.readToShort(rowResultado["abaixo_min_status"]);
					pedidoItemDevolvido.abaixo_min_autorizacao = BD.readToString(rowResultado["abaixo_min_autorizacao"]);
					pedidoItemDevolvido.abaixo_min_autorizador = BD.readToString(rowResultado["abaixo_min_autorizador"]);
					pedidoItemDevolvido.markup_fabricante = BD.readToSingle(rowResultado["markup_fabricante"]);
					pedidoItemDevolvido.motivo = BD.readToString(rowResultado["motivo"]);
					pedidoItemDevolvido.preco_NF = BD.readToDecimal(rowResultado["preco_NF"]);
					pedidoItemDevolvido.comissao_descontada = BD.readToShort(rowResultado["comissao_descontada"]);
					pedidoItemDevolvido.comissao_descontada_ult_op = BD.readToString(rowResultado["comissao_descontada_ult_op"]);
					pedidoItemDevolvido.comissao_descontada_data = BD.readToDateTime(rowResultado["comissao_descontada_data"]);
					pedidoItemDevolvido.comissao_descontada_usuario = BD.readToString(rowResultado["comissao_descontada_usuario"]);
					pedidoItemDevolvido.abaixo_min_superv_autorizador = BD.readToString(rowResultado["abaixo_min_superv_autorizador"]);
					pedidoItemDevolvido.vl_custo2 = BD.readToDecimal(rowResultado["vl_custo2"]);
					pedidoItemDevolvido.descricao_html = BD.readToString(rowResultado["descricao_html"]);
					pedidoItemDevolvido.custoFinancFornecCoeficiente = BD.readToSingle(rowResultado["custoFinancFornecCoeficiente"]);
					pedidoItemDevolvido.custoFinancFornecPrecoListaBase = BD.readToDecimal(rowResultado["custoFinancFornecPrecoListaBase"]);
					pedido.listaPedidoItemDevolvido.Add(pedidoItemDevolvido);

					pedido.vlTotalPrecoNfDestePedido -= pedidoItemDevolvido.qtde * pedidoItemDevolvido.preco_NF;
					pedido.vlTotalPrecoVendaDestePedido -= pedidoItemDevolvido.qtde * pedidoItemDevolvido.preco_venda;
				}
				#endregion

				#endregion

				#region [ Calcula valor total já pago ]

				#region [ Monta Select ]
				strSql = "SELECT" +
							" Coalesce(SUM(valor), 0) AS vl_total" +
						" FROM t_PEDIDO_PAGAMENTO" +
						" WHERE" +
							" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total já pago!!");

				#region [ Carrega os dados ]
				rowResultado = dtbResultado.Rows[0];
				pedido.vlTotalFamiliaPago = BD.readToDecimal(rowResultado["vl_total"]);
				#endregion

				#endregion

				#region [ Calcula o valor total da família de pedidos ]

				#region [ Monta Select ]
				strSql = "SELECT" +
							" Coalesce(SUM(qtde*preco_venda), 0) AS vl_total," +
							" Coalesce(SUM(qtde*preco_NF), 0) AS vl_total_NF" +
						" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" +
							" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" +
						" WHERE" +
							" (st_entrega<>'" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
							" AND (t_PEDIDO.pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total da família de pedidos!!");

				#region [ Carrega os dados ]
				rowResultado = dtbResultado.Rows[0];
				pedido.vlTotalFamiliaPrecoVenda = BD.readToDecimal(rowResultado["vl_total"]);
				pedido.vlTotalFamiliaPrecoNF = BD.readToDecimal(rowResultado["vl_total_NF"]);
				#endregion

				#endregion

				#region [ Calcula o valor total em devoluções da família de pedidos ]

				#region [ Monta Select ]
				strSql = "SELECT" +
							" Coalesce(SUM(qtde*preco_venda), 0) AS vl_total," +
							" Coalesce(SUM(qtde*preco_NF), 0) AS vl_total_NF" +
						" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
						" WHERE" +
							" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total em devoluções da família de pedidos!!");

				#region [ Carrega os dados ]
				rowResultado = dtbResultado.Rows[0];
				pedido.vlTotalFamiliaDevolucaoPrecoVenda = BD.readToDecimal(rowResultado["vl_total"]);
				pedido.vlTotalFamiliaDevolucaoPrecoNF = BD.readToDecimal(rowResultado["vl_total_NF"]);

				pedido.vlTotalFamiliaPrecoVenda -= pedido.vlTotalFamiliaDevolucaoPrecoVenda;
				pedido.vlTotalFamiliaPrecoNF -= pedido.vlTotalFamiliaDevolucaoPrecoNF;
				#endregion

				#endregion

				#region [ Calcula o valor em boletos deste pedido ]
				if (pedido.vlTotalFamiliaPrecoNF == 0)
				{
					razaoValorPedidoFilhote = 0m;
				}
				else
				{
					razaoValorPedidoFilhote = pedido.vlTotalPrecoNfDestePedido / pedido.vlTotalFamiliaPrecoNF;
				}

				#region [ Calcula o valor proporcional, pois pode ser um pedido filhote ]
				if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
				{
					#region [ À vista ]
					vlFormaPagtoDestePedido = pedido.vlTotalPrecoNfDestePedido;
					if ((pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
						|| (pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV))
					{
						vlBoletoDestePedido = pedido.vlTotalPrecoNfDestePedido;
					}
					#endregion
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
				{
					#region [ Parcela única ]
					vlFormaPagtoDestePedido = Global.arredondaParaMonetario(pedido.pu_valor * razaoValorPedidoFilhote);
					if (pedido.pu_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						vlBoletoDestePedido = Global.arredondaParaMonetario(pedido.pu_valor * razaoValorPedidoFilhote);
					}
					#endregion
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
				{
					#region [ Parcelado com entrada ]
					vlFormaPagtoDestePedido = Global.arredondaParaMonetario(pedido.pce_entrada_valor * razaoValorPedidoFilhote);
					if (pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						vlBoletoDestePedido = Global.arredondaParaMonetario(pedido.pce_entrada_valor * razaoValorPedidoFilhote);
					}

					vlFormaPagtoDestePedido += Global.arredondaParaMonetario((pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor) * razaoValorPedidoFilhote);
					if (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						vlBoletoDestePedido += Global.arredondaParaMonetario((pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor) * razaoValorPedidoFilhote);
					}
					#endregion
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
				{
					#region [ Parcelado sem entrada ]
					vlFormaPagtoDestePedido = Global.arredondaParaMonetario(pedido.pse_prim_prest_valor * razaoValorPedidoFilhote);
					if (pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						vlBoletoDestePedido = Global.arredondaParaMonetario(pedido.pse_prim_prest_valor * razaoValorPedidoFilhote);
					}

					vlFormaPagtoDestePedido += Global.arredondaParaMonetario((pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor) * razaoValorPedidoFilhote);
					if (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						vlBoletoDestePedido += Global.arredondaParaMonetario((pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor) * razaoValorPedidoFilhote);
					}
					#endregion
				}
				#endregion

				vlDiferencaArredondamento = pedido.vlTotalPrecoNfDestePedido - vlFormaPagtoDestePedido;

				pedido.vlTotalFormaPagtoDestePedido = vlFormaPagtoDestePedido;
				pedido.vlTotalBoletoDestePedido = vlBoletoDestePedido;
				if (Math.Abs(vlDiferencaArredondamento) <= 1) pedido.vlTotalBoletoDestePedido += vlDiferencaArredondamento;
				#endregion

				#region [ Calcula o valor que será pago através de cartão ]
				if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
				{
					if (pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
					{
						blnCartaoPagtoIntegral = true;
						vlPagtoEmCartao = pedido.vl_total_NF;
					}
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO)
				{
					blnCartaoPagtoIntegral = true;
					vlPagtoEmCartao = pedido.pc_qtde_parcelas * pedido.pc_valor_parcela;
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
				{
					if (pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
					{
						vlPagtoEmCartao = pedido.pce_entrada_valor;
					}
					if (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
					{
						vlPagtoEmCartao += pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor;
					}
					if ((pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) && (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO))
					{
						blnCartaoPagtoIntegral = true;
					}
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
				{
					if (pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
					{
						vlPagtoEmCartao = pedido.pse_prim_prest_valor;
					}
					if (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
					{
						vlPagtoEmCartao += pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor;
					}
					if ((pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) && (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO))
					{
						blnCartaoPagtoIntegral = true;
					}
				}
				else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
				{
					if (pedido.pu_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
					{
						vlPagtoEmCartao = pedido.pu_valor;
						blnCartaoPagtoIntegral = true;
					}
				}

				if (blnCartaoPagtoIntegral)
				{
					pedido.vlPagtoEmCartao = pedido.vl_total_NF;
				}
				else
				{
					pedido.vlPagtoEmCartao = vlPagtoEmCartao;
				}
				#endregion
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return pedido;
		}
		#endregion
	}
}