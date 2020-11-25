using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinanceiroService
{
    #region [ Pedido ]
    class Pedido
    {
        #region [ Construtor ]
        public Pedido()
        {
            listaPedidoItem = new List<PedidoItem>();
            listaPedidoItemDevolvido = new List<PedidoItemDevolvido>();
        }
        #endregion

        #region [ Getters/Setters ]

        public List<PedidoItem> listaPedidoItem;
        public List<PedidoItemDevolvido> listaPedidoItemDevolvido;

        public String pedido { get; set; }

        public String loja { get; set; }

        public String loja_razao_social { get; set; }

        public String loja_nome { get; set; }

        public DateTime data { get; set; }

        public String hora { get; set; }

        public DateTime data_hora { get; set; }

        public String id_cliente { get; set; }

        public String midia { get; set; }

        public String servicos { get; set; }

        public decimal vl_servicos { get; set; }

        public String vendedor { get; set; }

        public String vendedor_nome { get; set; }

        public String st_entrega { get; set; }

        public DateTime entregue_data { get; set; }

        public String entregue_usuario { get; set; }

        public DateTime cancelado_data { get; set; }

        public String cancelado_usuario { get; set; }

        public String st_pagto { get; set; }

        public String st_recebido { get; set; }

        public String obs_1 { get; set; }

        public String obs_2 { get; set; }

        public String obs_3 { get; set; }

        public short qtde_parcelas { get; set; }

        public String forma_pagto { get; set; }

        public decimal vl_total_familia { get; set; }

        public decimal vl_pago_familia { get; set; }

        public short split_status { get; set; }

        public DateTime split_data { get; set; }

        public String split_hora { get; set; }

        public String split_usuario { get; set; }

        public short a_entregar_status { get; set; }

        public DateTime a_entregar_data_marcada { get; set; }

        public DateTime a_entregar_data { get; set; }

        public String a_entregar_hora { get; set; }

        public String a_entregar_usuario { get; set; }

        public String loja_indicou { get; set; }

        public double comissao_loja_indicou { get; set; }

        public short venda_externa { get; set; }

        public decimal vl_frete { get; set; }

        public String transportadora_id { get; set; }

        public DateTime transportadora_data { get; set; }

        public String transportadora_usuario { get; set; }

        public short analise_credito { get; set; }

        public DateTime analise_credito_data { get; set; }

        public String analise_credito_usuario { get; set; }

        public short tipo_parcelamento { get; set; }

        public short av_forma_pagto { get; set; }

        public short pc_qtde_parcelas { get; set; }

        public decimal pc_valor_parcela { get; set; }

        public short pc_maquineta_qtde_parcelas { get; set; }

        public decimal pc_maquineta_valor_parcela { get; set; }

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

        public String indicador { get; set; }

        public String indicador_desempenho_nota { get; set; }

        public decimal vl_total_NF { get; set; }

        public decimal vl_total_RA { get; set; }

        public double perc_RT { get; set; }

        public short st_orc_virou_pedido { get; set; }

        public String orcamento { get; set; }

        public String orcamentista { get; set; }

        public short comissao_paga { get; set; }

        public String comissao_paga_ult_op { get; set; }

        public DateTime comissao_paga_data { get; set; }

        public String comissao_paga_usuario { get; set; }

        public double perc_desagio_RA { get; set; }

        public double perc_limite_RA_sem_desagio { get; set; }

        public decimal vl_total_RA_liquido { get; set; }

        public short st_tem_desagio_RA { get; set; }

        public short qtde_parcelas_desagio_RA { get; set; }

        public String transportadora_num_coleta { get; set; }

        public String transportadora_contato { get; set; }

        public byte st_memorizacao_completa_enderecos { get; set; } = 0;

        public byte endereco_memorizado_status { get; set; } = 0;

        public string endereco_logradouro { get; set; } = "";

        public string endereco_bairro { get; set; } = "";

        public string endereco_cidade { get; set; } = "";

        public string endereco_uf { get; set; } = "";

        public string endereco_cep { get; set; } = "";

        public string endereco_numero { get; set; } = "";

        public string endereco_complemento { get; set; } = "";

        public string endereco_email { get; set; } = "";

        public string endereco_email_xml { get; set; } = "";

        public string endereco_nome { get; set; } = "";

        public string endereco_ddd_res { get; set; } = "";

        public string endereco_tel_res { get; set; } = "";

        public string endereco_ddd_com { get; set; } = "";

        public string endereco_tel_com { get; set; } = "";

        public string endereco_ramal_com { get; set; } = "";

        public string endereco_ddd_cel { get; set; } = "";

        public string endereco_tel_cel { get; set; } = "";

        public string endereco_ddd_com_2 { get; set; } = "";

        public string endereco_tel_com_2 { get; set; } = "";

        public string endereco_ramal_com_2 { get; set; } = "";

        public string endereco_tipo_pessoa { get; set; } = "";

        public string endereco_cnpj_cpf { get; set; } = "";

        public byte endereco_contribuinte_icms_status { get; set; } = 0;

        public byte endereco_produtor_rural_status { get; set; } = 0;

        public string endereco_ie { get; set; } = "";

        public string endereco_rg { get; set; } = "";

        public string endereco_contato { get; set; } = "";

        public short st_end_entrega { get; set; }

        public String endEtg_endereco { get; set; } = "";

        public String endEtg_endereco_numero { get; set; } = "";

        public String endEtg_endereco_complemento { get; set; } = "";

        public String endEtg_bairro { get; set; } = "";

        public String endEtg_cidade { get; set; } = "";

        public String endEtg_uf { get; set; } = "";

        public String endEtg_cep { get; set; } = "";

        public string endEtg_email { get; set; } = "";

        public string endEtg_email_xml { get; set; } = "";

        public string endEtg_nome { get; set; } = "";

        public string endEtg_ddd_res { get; set; } = "";

        public string endEtg_tel_res { get; set; } = "";

        public string endEtg_ddd_com { get; set; } = "";

        public string endEtg_tel_com { get; set; } = "";

        public string endEtg_ramal_com { get; set; } = "";

        public string endEtg_ddd_cel { get; set; } = "";

        public string endEtg_tel_cel { get; set; } = "";

        public string endEtg_ddd_com_2 { get; set; } = "";

        public string endEtg_tel_com_2 { get; set; } = "";

        public string endEtg_ramal_com_2 { get; set; } = "";

        public string endEtg_tipo_pessoa { get; set; } = "";

        public string endEtg_cnpj_cpf { get; set; } = "";

        public byte endEtg_contribuinte_icms_status { get; set; } = 0;

        public byte endEtg_produtor_rural_status { get; set; } = 0;

        public string endEtg_ie { get; set; } = "";

        public string endEtg_rg { get; set; } = "";

        public short st_etg_imediata { get; set; }

        public DateTime etg_imediata_data { get; set; }

        public String etg_imediata_usuario { get; set; }

        public DateTime PrevisaoEntregaData { get; set; }

        public string PrevisaoEntregaUsuarioUltAtualiz { get; set; }

        public DateTime PrevisaoEntregaDtHrUltAtualiz { get; set; }

        public short frete_status { get; set; }

        public decimal frete_valor { get; set; }

        public DateTime frete_data { get; set; }

        public String frete_usuario { get; set; }

        public short stBemUsoConsumo { get; set; }

        public short pedidoRecebidoStatus { get; set; }

        public DateTime pedidoRecebidoData { get; set; }

        public String pedidoRecebidoUsuarioUltAtualiz { get; set; }

        public DateTime pedidoRecebidoDtHrUltAtualiz { get; set; }

        public short instaladorInstalaStatus { get; set; }

        public String instaladorInstalaUsuarioUltAtualiz { get; set; }

        public DateTime instaladorInstalaDtHrUltAtualiz { get; set; }

        public String custoFinancFornecTipoParcelamento { get; set; }

        public short custoFinancFornecQtdeParcelas { get; set; }

        public int tamanho_num_pedido { get; set; }

        public string pedido_base { get; set; }

        public byte st_forma_pagto_somente_cartao { get; set; }

        public int id_nfe_emitente { get; set; }

        public byte st_auto_split { get; set; }

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

        #endregion
    }
    #endregion

    #region [ PedidoItem ]
    class PedidoItem
    {
        #region [ Getters/Setters ]

        public String pedido { get; set; }

        public String fabricante { get; set; }

        public String produto { get; set; }

        public short qtde { get; set; }

        public double desc_dado { get; set; }

        public decimal preco_venda { get; set; }

        public decimal preco_fabricante { get; set; }

        public decimal preco_lista { get; set; }

        public double margem { get; set; }

        public double desc_max { get; set; }

        public double comissao { get; set; }

        public String descricao { get; set; }

        public String ean { get; set; }

        public String grupo { get; set; }

        public double peso { get; set; }

        public short qtde_volumes { get; set; }

        public short abaixo_min_status { get; set; }

        public String abaixo_min_autorizacao { get; set; }

        public String abaixo_min_autorizador { get; set; }

        public short sequencia { get; set; }

        public double markup_fabricante { get; set; }

        public decimal preco_NF { get; set; }

        public String abaixo_min_superv_autorizador { get; set; }

        public decimal vl_custo2 { get; set; }

        public String descricao_html { get; set; }

        public double custoFinancFornecCoeficiente { get; set; }

        public decimal custoFinancFornecPrecoListaBase { get; set; }

        #endregion
    }
    #endregion

    #region [ PedidoItemDevolvido ]
    class PedidoItemDevolvido
    {
        #region [ Getters/Setters ]

        public String id { get; set; }

        public DateTime devolucao_data { get; set; }

        public String devolucao_hora { get; set; }

        public String devolucao_usuario { get; set; }

        public String pedido { get; set; }

        public String fabricante { get; set; }

        public String produto { get; set; }

        public short qtde { get; set; }

        public double desc_dado { get; set; }

        public decimal preco_venda { get; set; }

        public decimal preco_fabricante { get; set; }

        public decimal preco_lista { get; set; }

        public double margem { get; set; }

        public double desc_max { get; set; }

        public double comissao { get; set; }

        public String descricao { get; set; }

        public String ean { get; set; }

        public String grupo { get; set; }

        public double peso { get; set; }

        public short qtde_volumes { get; set; }

        public short abaixo_min_status { get; set; }

        public String abaixo_min_autorizacao { get; set; }

        public String abaixo_min_autorizador { get; set; }

        public double markup_fabricante { get; set; }

        public String motivo { get; set; }

        public decimal preco_NF { get; set; }

        public short comissao_descontada { get; set; }

        public String comissao_descontada_ult_op { get; set; }

        public DateTime comissao_descontada_data { get; set; }

        public String comissao_descontada_usuario { get; set; }

        public String abaixo_min_superv_autorizador { get; set; }

        public decimal vl_custo2 { get; set; }

        public String descricao_html { get; set; }

        public double custoFinancFornecCoeficiente { get; set; }

        public decimal custoFinancFornecPrecoListaBase { get; set; }

        #endregion
    }
    #endregion

    #region [ PedidoPagamento ]
    class PedidoPagamento
    {
        public string id { get; set; }

        public string pedido { get; set; }

        public DateTime data { get; set; }

        public string hora { get; set; }

        public decimal valor { get; set; }

        public string tipo_pagto { get; set; }

        public string usuario { get; set; }

        public int id_pedido_pagto_cielo { get; set; }

        public int id_pedido_pagto_braspag { get; set; }

        public int id_pagto_gw_pag_payment { get; set; }

        public int id_braspag_webhook_complementar { get; set; }
    }
    #endregion

    #region [ PedidoHistPagto ]
    class PedidoHistPagto
    {
        public int id { get; set; }

        public string pedido { get; set; }

        public byte status { get; set; }

        public int id_fluxo_caixa { get; set; }

        public int ctrl_pagto_id_parcela { get; set; }

        public byte ctrl_pagto_modulo { get; set; }

        public DateTime dt_vencto { get; set; }

        public decimal valor_total { get; set; }

        public decimal valor_rateado { get; set; }

        public decimal valor_pago { get; set; }

        public string descricao { get; set; }

        public DateTime dt_credito { get; set; }

        public DateTime dt_cadastro { get; set; }

        public string usuario_cadastro { get; set; }

        public DateTime dt_ult_atualizacao { get; set; }

        public string usuario_ult_atualizacao { get; set; }

        public decimal vl_abatimento_concedido { get; set; }

        public byte st_boleto_pago_cheque { get; set; }

        public DateTime dt_ocorrencia_banco_boleto_pago_cheque { get; set; }

        public byte st_boleto_ocorrencia_17 { get; set; }

        public DateTime dt_ocorrencia_banco_boleto_ocorrencia_17 { get; set; }

        public byte st_boleto_ocorrencia_15 { get; set; }

        public DateTime dt_ocorrencia_banco_boleto_ocorrencia_15 { get; set; }

        public byte st_boleto_ocorrencia_23 { get; set; }

        public DateTime dt_ocorrencia_banco_boleto_ocorrencia_23 { get; set; }

        public byte st_boleto_ocorrencia_34 { get; set; }

        public DateTime dt_ocorrencia_banco_boleto_ocorrencia_34 { get; set; }

        public byte st_boleto_baixado { get; set; }

        public DateTime dt_ocorrencia_banco_boleto_baixado { get; set; }

        public DateTime dt_operacao { get; set; }
    }
    #endregion
}
