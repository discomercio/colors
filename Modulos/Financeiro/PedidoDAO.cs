﻿#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class PedidoDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmPedidoMarcaStatusBoletoConfeccionado;
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

		#region [ Construtor estático ]
		static PedidoDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmPedidoMarcaStatusBoletoConfeccionado ]
			strSql = "UPDATE t_PEDIDO SET " +
						"BoletoConfeccionadoStatus = @BoletoConfeccionadoStatus, " +
						"BoletoConfeccionadoData = " + Global.sqlMontaGetdateSomenteData() +
					" WHERE" +
						" (pedido = @pedido)";
			cmPedidoMarcaStatusBoletoConfeccionado = BD.criaSqlCommand();
			cmPedidoMarcaStatusBoletoConfeccionado.CommandText = strSql;
			cmPedidoMarcaStatusBoletoConfeccionado.Parameters.Add("@BoletoConfeccionadoStatus", SqlDbType.TinyInt);
			cmPedidoMarcaStatusBoletoConfeccionado.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmPedidoMarcaStatusBoletoConfeccionado.Prepare();
			#endregion
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
			int flagPedidoUsarMemorizacaoCompletaEnderecos;
			String strSql;
			String numeroPedidoBase;
			decimal razaoValorPedidoFilhote = 0m;
			decimal vlBoletoDestePedido = 0m;
			decimal vlFormaPagtoDestePedido = 0m;
			decimal vlDiferencaArredondamento;
			Pedido pedido = new Pedido();
			PedidoItem pedidoItem;
			PedidoItemDevolvido pedidoItemDevolvido;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroPedido == null) throw new FinanceiroException("Nº do pedido a ser consultado não foi fornecido!!");
			if (numeroPedido.Length == 0) throw new FinanceiroException("Nº do pedido a ser consultado não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			numeroPedido = numeroPedido.Trim();
			numeroPedidoBase = Global.retornaNumeroPedidoBase(numeroPedido);
			flagPedidoUsarMemorizacaoCompletaEnderecos = ComumDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS, 0);
			#endregion

			#region [ Pesquisa pedido ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" t_PEDIDO.*," +
						" t_LOJA.razao_social AS loja_razao_social," +
						" t_LOJA.nome AS loja_nome," +
						" t_USUARIO_VENDEDOR.nome AS vendedor_nome," +
						" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota AS indicador_desempenho_nota" +
					" FROM t_PEDIDO" +
						" INNER JOIN t_LOJA ON (t_PEDIDO.loja=t_LOJA.loja)" +
						" INNER JOIN t_USUARIO AS t_USUARIO_VENDEDOR ON (t_PEDIDO.vendedor=t_USUARIO_VENDEDOR.usuario)" +
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

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Pedido nº " + numeroPedido + " não foi encontrado!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.pedido = BD.readToString(rowResultado["pedido"]);
			pedido.loja = BD.readToString(rowResultado["loja"]);
			pedido.loja_razao_social = BD.readToString(rowResultado["loja_razao_social"]);
			pedido.loja_nome = BD.readToString(rowResultado["loja_nome"]);
			pedido.data = BD.readToDateTime(rowResultado["data"]);
			pedido.hora = BD.readToString(rowResultado["hora"]);
			pedido.id_cliente = BD.readToString(rowResultado["id_cliente"]);
			pedido.midia = BD.readToString(rowResultado["midia"]);
			pedido.servicos = BD.readToString(rowResultado["servicos"]);
			pedido.vl_servicos = BD.readToDecimal(rowResultado["vl_servicos"]);
			pedido.vendedor = BD.readToString(rowResultado["vendedor"]);
			pedido.vendedor_nome = BD.readToString(rowResultado["vendedor_nome"]);
			pedido.st_entrega = BD.readToString(rowResultado["st_entrega"]);
			pedido.entregue_data = BD.readToDateTime(rowResultado["entregue_data"]);
			pedido.entregue_usuario = BD.readToString(rowResultado["entregue_usuario"]);
			pedido.cancelado_data = BD.readToDateTime(rowResultado["cancelado_data"]);
			pedido.cancelado_usuario = BD.readToString(rowResultado["cancelado_usuario"]);
			pedido.st_pagto = BD.readToString(rowResultado["st_pagto"]);
			pedido.st_recebido = BD.readToString(rowResultado["st_recebido"]);
			pedido.obs_1 = BD.readToString(rowResultado["obs_1"]);
			pedido.obs_2 = BD.readToString(rowResultado["obs_2"]);
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
			pedido.pc_maquineta_qtde_parcelas = BD.readToShort(rowResultado["pc_maquineta_qtde_parcelas"]);
			pedido.pc_maquineta_valor_parcela = BD.readToDecimal(rowResultado["pc_maquineta_valor_parcela"]);
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

			pedido.endereco_memorizado_status = BD.readToByte(rowResultado["endereco_memorizado_status"]);
			if (pedido.endereco_memorizado_status != 0)
			{
				pedido.endereco_logradouro = BD.readToString(rowResultado["endereco_logradouro"]);
				pedido.endereco_numero = BD.readToString(rowResultado["endereco_numero"]);
				pedido.endereco_complemento = BD.readToString(rowResultado["endereco_complemento"]);
				pedido.endereco_bairro = BD.readToString(rowResultado["endereco_bairro"]);
				pedido.endereco_cidade = BD.readToString(rowResultado["endereco_cidade"]);
				pedido.endereco_uf = BD.readToString(rowResultado["endereco_uf"]);
				pedido.endereco_cep = BD.readToString(rowResultado["endereco_cep"]);

				if (flagPedidoUsarMemorizacaoCompletaEnderecos != 0)
				{
					pedido.st_memorizacao_completa_enderecos = BD.readToByte(rowResultado["st_memorizacao_completa_enderecos"]);
					if (pedido.st_memorizacao_completa_enderecos != 0)
					{
						pedido.endereco_email = BD.readToString(rowResultado["endereco_email"]);
						pedido.endereco_email_xml = BD.readToString(rowResultado["endereco_email_xml"]);
						pedido.endereco_nome = BD.readToString(rowResultado["endereco_nome"]);
						pedido.endereco_ddd_res = BD.readToString(rowResultado["endereco_ddd_res"]);
						pedido.endereco_tel_res = BD.readToString(rowResultado["endereco_tel_res"]);
						pedido.endereco_ddd_com = BD.readToString(rowResultado["endereco_ddd_com"]);
						pedido.endereco_tel_com = BD.readToString(rowResultado["endereco_tel_com"]);
						pedido.endereco_ramal_com = BD.readToString(rowResultado["endereco_ramal_com"]);
						pedido.endereco_ddd_cel = BD.readToString(rowResultado["endereco_ddd_cel"]);
						pedido.endereco_tel_cel = BD.readToString(rowResultado["endereco_tel_cel"]);
						pedido.endereco_ddd_com_2 = BD.readToString(rowResultado["endereco_ddd_com_2"]);
						pedido.endereco_tel_com_2 = BD.readToString(rowResultado["endereco_tel_com_2"]);
						pedido.endereco_ramal_com_2 = BD.readToString(rowResultado["endereco_ramal_com_2"]);
						pedido.endereco_tipo_pessoa = BD.readToString(rowResultado["endereco_tipo_pessoa"]);
						pedido.endereco_cnpj_cpf = BD.readToString(rowResultado["endereco_cnpj_cpf"]);
						pedido.endereco_contribuinte_icms_status = BD.readToByte(rowResultado["endereco_contribuinte_icms_status"]);
						pedido.endereco_produtor_rural_status = BD.readToByte(rowResultado["endereco_produtor_rural_status"]);
						pedido.endereco_ie = BD.readToString(rowResultado["endereco_ie"]);
						pedido.endereco_rg = BD.readToString(rowResultado["endereco_rg"]);
						pedido.endereco_contato = BD.readToString(rowResultado["endereco_contato"]);
					}
				}
			}

			pedido.st_end_entrega = BD.readToShort(rowResultado["st_end_entrega"]);
			if (pedido.st_end_entrega != 0)
			{
				pedido.endEtg_endereco = BD.readToString(rowResultado["EndEtg_endereco"]);
				pedido.endEtg_endereco_numero = BD.readToString(rowResultado["EndEtg_endereco_numero"]);
				pedido.endEtg_endereco_complemento = BD.readToString(rowResultado["EndEtg_endereco_complemento"]);
				pedido.endEtg_bairro = BD.readToString(rowResultado["EndEtg_bairro"]);
				pedido.endEtg_cidade = BD.readToString(rowResultado["EndEtg_cidade"]);
				pedido.endEtg_uf = BD.readToString(rowResultado["EndEtg_uf"]);
				pedido.endEtg_cep = BD.readToString(rowResultado["EndEtg_cep"]);

				if (flagPedidoUsarMemorizacaoCompletaEnderecos != 0)
				{
					if (pedido.st_memorizacao_completa_enderecos != 0)
					{
						pedido.endEtg_email = BD.readToString(rowResultado["EndEtg_email"]);
						pedido.endEtg_email_xml = BD.readToString(rowResultado["EndEtg_email_xml"]);
						pedido.endEtg_nome = BD.readToString(rowResultado["EndEtg_nome"]);
						pedido.endEtg_ddd_res = BD.readToString(rowResultado["EndEtg_ddd_res"]);
						pedido.endEtg_tel_res = BD.readToString(rowResultado["EndEtg_tel_res"]);
						pedido.endEtg_ddd_com = BD.readToString(rowResultado["EndEtg_ddd_com"]);
						pedido.endEtg_tel_com = BD.readToString(rowResultado["EndEtg_tel_com"]);
						pedido.endEtg_ramal_com = BD.readToString(rowResultado["EndEtg_ramal_com"]);
						pedido.endEtg_ddd_cel = BD.readToString(rowResultado["EndEtg_ddd_cel"]);
						pedido.endEtg_tel_cel = BD.readToString(rowResultado["EndEtg_tel_cel"]);
						pedido.endEtg_ddd_com_2 = BD.readToString(rowResultado["EndEtg_ddd_com_2"]);
						pedido.endEtg_tel_com_2 = BD.readToString(rowResultado["EndEtg_tel_com_2"]);
						pedido.endEtg_ramal_com_2 = BD.readToString(rowResultado["EndEtg_ramal_com_2"]);
						pedido.endEtg_tipo_pessoa = BD.readToString(rowResultado["EndEtg_tipo_pessoa"]);
						pedido.endEtg_cnpj_cpf = BD.readToString(rowResultado["EndEtg_cnpj_cpf"]);
						pedido.endEtg_contribuinte_icms_status = BD.readToByte(rowResultado["EndEtg_contribuinte_icms_status"]);
						pedido.endEtg_produtor_rural_status = BD.readToByte(rowResultado["EndEtg_produtor_rural_status"]);
						pedido.endEtg_ie = BD.readToString(rowResultado["EndEtg_ie"]);
						pedido.endEtg_rg = BD.readToString(rowResultado["EndEtg_rg"]);
					}
				}
			}

			pedido.st_etg_imediata = BD.readToShort(rowResultado["st_etg_imediata"]);
			pedido.etg_imediata_data = BD.readToDateTime(rowResultado["etg_imediata_data"]);
			pedido.etg_imediata_usuario = BD.readToString(rowResultado["etg_imediata_usuario"]);
			pedido.PrevisaoEntregaData = BD.readToDateTime(rowResultado["PrevisaoEntregaData"]);
			pedido.PrevisaoEntregaUsuarioUltAtualiz = BD.readToString(rowResultado["PrevisaoEntregaUsuarioUltAtualiz"]);
			pedido.PrevisaoEntregaDtHrUltAtualiz = BD.readToDateTime(rowResultado["PrevisaoEntregaDtHrUltAtualiz"]);
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
			pedido.id_nfe_emitente = BD.readToInt(rowResultado["id_nfe_emitente"]);
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

				if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Pedido-base nº " + numeroPedidoBase + " não foi encontrado!!");

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
				pedido.pc_maquineta_qtde_parcelas = BD.readToShort(rowResultado["pc_maquineta_qtde_parcelas"]);
				pedido.pc_maquineta_valor_parcela = BD.readToDecimal(rowResultado["pc_maquineta_valor_parcela"]);
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

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Itens do pedido nº " + numeroPedido + " não foram encontrados!!");

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

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Falha ao calcular o valor total já pago!!");

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

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Falha ao calcular o valor total da família de pedidos!!");

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

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Falha ao calcular o valor total em devoluções da família de pedidos!!");

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

			return pedido;
		}
		#endregion

		#region [ getPedidoStPagto ]
		public static string getPedidoStPagto(string numeroPedido)
		{
			#region [ Declarações ]
			string strSql;
			string numeroPedidoBase;
			string st_pagto;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroPedido == null) throw new FinanceiroException("Nº do pedido a ser consultado não foi fornecido!!");
			if (numeroPedido.Length == 0) throw new FinanceiroException("Nº do pedido a ser consultado não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			numeroPedido = numeroPedido.Trim();
			numeroPedidoBase = Global.retornaNumeroPedidoBase(numeroPedido);
			#endregion

			#region [ Monta Select ]
			strSql = "SELECT" +
						" st_pagto" +
					" FROM t_PEDIDO" +
					" WHERE" +
						" (pedido = '" + numeroPedidoBase + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Pedido nº " + numeroPedido + " não foi encontrado!!");

			#region [ Processa o resultado ]
			rowResultado = dtbResultado.Rows[0];
			st_pagto = BD.readToString(rowResultado["st_pagto"]);
			#endregion

			return st_pagto;
		}
		#endregion

		#region [ marcaPedidoStatusBoletoConfeccionado ]
		public static bool marcaPedidoStatusBoletoConfeccionado(String usuario, 
													  String pedido, 
													  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca o pedido com o status de boleto já confeccionado";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoMarcaStatusBoletoConfeccionado.Parameters["@pedido"].Value = pedido;
				cmPedidoMarcaStatusBoletoConfeccionado.Parameters["@BoletoConfeccionadoStatus"].Value = Global.Cte.FIN.T_PEDIDO__BOLETO_CONFECCIONADO_STATUS.SIM;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoMarcaStatusBoletoConfeccionado);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
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
					strMsgErro = "Falha ao tentar marcar o pedido com o status de boleto já confeccionado!!";
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

		#endregion
	}
}
