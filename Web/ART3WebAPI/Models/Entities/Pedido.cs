using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	#region [ Pedido ]
	public class Pedido
	{
		#region [ Construtor ]
		public Pedido()
		{
			listaPedidoItem = new List<PedidoItem>();
			listaPedidoItemDevolvido = new List<PedidoItemDevolvido>();
		}
		#endregion

		public List<PedidoItem> listaPedidoItem;
		public List<PedidoItemDevolvido> listaPedidoItemDevolvido;

		public string pedido { get; set; }
		public string loja { get; set; }
		public string loja_razao_social { get; set; }
		public string loja_nome { get; set; }
		public DateTime data { get; set; }
		public DateTime data_hora { get; set; }
		public string hora { get; set; }
		public string id_cliente { get; set; }
		public string midia { get; set; }
		public string servicos { get; set; }
		public decimal vl_servicos { get; set; }
		public string vendedor { get; set; }
		public string vendedor_nome { get; set; }
		public string usuario_cadastro { get; set; }
		public string usuario_cadastro_nome { get; set; }
		public string st_entrega { get; set; }
		public DateTime entregue_data { get; set; }
		public string entregue_usuario { get; set; }
		public DateTime cancelado_data { get; set; }
		public string cancelado_usuario { get; set; }
		public string st_pagto { get; set; }
		public string st_recebido { get; set; }
		public string obs_1 { get; set; }
		public string obs_2 { get; set; }
		public string obs_3 { get; set; }
		public short qtde_parcelas { get; set; }
		public string forma_pagto { get; set; }
		public decimal vl_total_familia { get; set; }
		public decimal vl_pago_familia { get; set; }
		public short split_status { get; set; }
		public DateTime split_data { get; set; }
		public string split_hora { get; set; }
		public string split_usuario { get; set; }
		public short a_entregar_status { get; set; }
		public DateTime a_entregar_data_marcada { get; set; }
		public DateTime a_entregar_data { get; set; }
		public string a_entregar_hora { get; set; }
		public string a_entregar_usuario { get; set; }
		public string loja_indicou { get; set; }
		public double comissao_loja_indicou { get; set; }
		public short venda_externa { get; set; }
		public decimal vl_frete { get; set; }
		public string transportadora_id { get; set; }
		public DateTime transportadora_data { get; set; }
		public string transportadora_usuario { get; set; }
		public short analise_credito { get; set; }
		public DateTime analise_credito_data { get; set; }
		public string analise_credito_usuario { get; set; }
		public short tipo_parcelamento { get; set; }
		public short av_forma_pagto { get; set; }
		public short pc_qtde_parcelas { get; set; }
		public decimal pc_valor_parcela { get; set; }
		public short pce_forma_pagto_entrada { get; set; }
		public short pce_forma_pagto_prestacao { get; set; }
		public decimal pce_entrada_valor { get; set; }
		public short pce_prestacao_qtde { get; set; }
		public decimal pce_prestacao_valor { get; set; }
		public short pce_prestacao_periodo { get; set; }
		public short pse_forma_pagto_prim_prest { get; set; }
		public short pse_forma_pagto_demais_prest { get; set; }
		public decimal pse_prim_prest_valor { get; set; }
		public short pse_prim_prest_apos { get; set; }
		public short pse_demais_prest_qtde { get; set; }
		public decimal pse_demais_prest_valor { get; set; }
		public short pse_demais_prest_periodo { get; set; }
		public short pu_forma_pagto { get; set; }
		public decimal pu_valor { get; set; }
		public short pu_vencto_apos { get; set; }
		public string indicador { get; set; }
		public string indicador_desempenho_nota { get; set; }
		public decimal vl_total_NF { get; set; }
		public decimal vl_total_RA { get; set; }
		public double perc_RT { get; set; }
		public short st_orc_virou_pedido { get; set; }
		public string orcamento { get; set; }
		public string orcamentista { get; set; }
		public short comissao_paga { get; set; }
		public string comissao_paga_ult_op { get; set; }
		public DateTime comissao_paga_data { get; set; }
		public string comissao_paga_usuario { get; set; }
		public double perc_desagio_RA { get; set; }
		public double perc_limite_RA_sem_desagio { get; set; }
		public decimal vl_total_RA_liquido { get; set; }
		public short st_tem_desagio_RA { get; set; }
		public short qtde_parcelas_desagio_RA { get; set; }
		public string transportadora_num_coleta { get; set; }
		public string transportadora_contato { get; set; }
		public short st_end_entrega { get; set; }
		public string endEtg_endereco { get; set; }
		public string endEtg_endereco_numero { get; set; }
		public string endEtg_endereco_complemento { get; set; }
		public string endEtg_bairro { get; set; }
		public string endEtg_cidade { get; set; }
		public string endEtg_uf { get; set; }
		public string endEtg_cep { get; set; }
		public short st_etg_imediata { get; set; }
		public DateTime etg_imediata_data { get; set; }
		public string etg_imediata_usuario { get; set; }
		public short frete_status { get; set; }
		public decimal frete_valor { get; set; }
		public DateTime frete_data { get; set; }
		public string frete_usuario { get; set; }
		public short stBemUsoConsumo { get; set; }
		public short pedidoRecebidoStatus { get; set; }
		public DateTime pedidoRecebidoData { get; set; }
		public string pedidoRecebidoUsuarioUltAtualiz { get; set; }
		public DateTime pedidoRecebidoDtHrUltAtualiz { get; set; }
		public short instaladorInstalaStatus { get; set; }
		public string instaladorInstalaUsuarioUltAtualiz { get; set; }
		public DateTime instaladorInstalaDtHrUltAtualiz { get; set; }
		public string custoFinancFornecTipoParcelamento { get; set; }
		public short custoFinancFornecQtdeParcelas { get; set; }
		public int tamanho_num_pedido { get; set; }
		public string pedido_base { get; set; }
		public byte st_forma_pagto_somente_cartao { get; set; }
		public int id_nfe_emitente { get; set; }
		public byte st_auto_split { get; set; }
		public string pedido_bs_x_ac { get; set; }
		public string pedido_bs_x_marketplace { get; set; }
		public string marketplace_codigo_origem { get; set; }

		#region [ Campos calculados ]
		public decimal vlTotalPrecoNfDestePedido { get; set; }
		public decimal vlTotalBoletoDestePedido { get; set; }
		public decimal vlTotalFormaPagtoDestePedido { get; set; }
		public decimal vlTotalPrecoVendaDestePedido { get; set; }
		public decimal vlTotalFamiliaPago { get; set; }
		public decimal vlTotalFamiliaPrecoVenda { get; set; }
		public decimal vlTotalFamiliaPrecoNF { get; set; }
		public decimal vlTotalFamiliaDevolucaoPrecoVenda { get; set; }
		public decimal vlTotalFamiliaDevolucaoPrecoNF { get; set; }
		public decimal vlPagtoEmCartao { get; set; }
		#endregion
	}
	#endregion

	#region [ PedidoItem ]
	public class PedidoItem
	{
		public string pedido { get; set; }
		public string fabricante { get; set; }
		public string produto { get; set; }
		public short qtde { get; set; }
		public double desc_dado { get; set; }
		public decimal preco_venda { get; set; }
		public decimal preco_fabricante { get; set; }
		public decimal preco_lista { get; set; }
		public double margem { get; set; }
		public double desc_max { get; set; }
		public double comissao { get; set; }
		public string descricao { get; set; }
		public string ean { get; set; }
		public string grupo { get; set; }
		public double peso { get; set; }
		public short qtde_volumes { get; set; }
		public short abaixo_min_status { get; set; }
		public string abaixo_min_autorizacao { get; set; }
		public string abaixo_min_autorizador { get; set; }
		public short sequencia { get; set; }
		public double markup_fabricante { get; set; }
		public decimal preco_NF { get; set; }
		public string abaixo_min_superv_autorizador { get; set; }
		public decimal vl_custo2 { get; set; }
		public string descricao_html { get; set; }
		public double custoFinancFornecCoeficiente { get; set; }
		public decimal custoFinancFornecPrecoListaBase { get; set; }
	}
	#endregion

	#region [ PedidoItemDevolvido ]
	public class PedidoItemDevolvido
	{
		public string id { get; set; }
		public DateTime devolucao_data { get; set; }
		public string devolucao_hora { get; set; }
		public string devolucao_usuario { get; set; }
		public string pedido { get; set; }
		public string fabricante { get; set; }
		public string produto { get; set; }
		public short qtde { get; set; }
		public double desc_dado { get; set; }
		public decimal preco_venda { get; set; }
		public decimal preco_fabricante { get; set; }
		public decimal preco_lista { get; set; }
		public double margem { get; set; }
		public double desc_max { get; set; }
		public double comissao { get; set; }
		public string descricao { get; set; }
		public string ean { get; set; }
		public string grupo { get; set; }
		public double peso { get; set; }
		public short qtde_volumes { get; set; }
		public short abaixo_min_status { get; set; }
		public string abaixo_min_autorizacao { get; set; }
		public string abaixo_min_autorizador { get; set; }
		public double markup_fabricante { get; set; }
		public string motivo { get; set; }
		public decimal preco_NF { get; set; }
		public short comissao_descontada { get; set; }
		public string comissao_descontada_ult_op { get; set; }
		public DateTime comissao_descontada_data { get; set; }
		public string comissao_descontada_usuario { get; set; }
		public string abaixo_min_superv_autorizador { get; set; }
		public decimal vl_custo2 { get; set; }
		public string descricao_html { get; set; }
		public double custoFinancFornecCoeficiente { get; set; }
		public decimal custoFinancFornecPrecoListaBase { get; set; }
	}
	#endregion
}