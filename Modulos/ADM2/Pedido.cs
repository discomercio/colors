using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
	public class Pedido
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

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private String _loja;
		public String loja
		{
			get { return _loja; }
			set { _loja = value; }
		}

		private String _loja_razao_social;
		public String loja_razao_social
		{
			get { return _loja_razao_social; }
			set { _loja_razao_social = value; }
		}

		private String _loja_nome;
		public String loja_nome
		{
			get { return _loja_nome; }
			set { _loja_nome = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private String _hora;
		public String hora
		{
			get { return _hora; }
			set { _hora = value; }
		}

		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private String _midia;
		public String midia
		{
			get { return _midia; }
			set { _midia = value; }
		}

		private String _servicos;
		public String servicos
		{
			get { return _servicos; }
			set { _servicos = value; }
		}

		private decimal _vl_servicos;
		public decimal vl_servicos
		{
			get { return _vl_servicos; }
			set { _vl_servicos = value; }
		}

		private String _vendedor;
		public String vendedor
		{
			get { return _vendedor; }
			set { _vendedor = value; }
		}

		private String _vendedor_nome;
		public String vendedor_nome
		{
			get { return _vendedor_nome; }
			set { _vendedor_nome = value; }
		}

		private String _st_entrega;
		public String st_entrega
		{
			get { return _st_entrega; }
			set { _st_entrega = value; }
		}

		private DateTime _entregue_data;
		public DateTime entregue_data
		{
			get { return _entregue_data; }
			set { _entregue_data = value; }
		}

		private String _entregue_usuario;
		public String entregue_usuario
		{
			get { return _entregue_usuario; }
			set { _entregue_usuario = value; }
		}

		private DateTime _cancelado_data;
		public DateTime cancelado_data
		{
			get { return _cancelado_data; }
			set { _cancelado_data = value; }
		}

		private String _cancelado_usuario;
		public String cancelado_usuario
		{
			get { return _cancelado_usuario; }
			set { _cancelado_usuario = value; }
		}

		private String _st_pagto;
		public String st_pagto
		{
			get { return _st_pagto; }
			set { _st_pagto = value; }
		}

		private String _st_recebido;
		public String st_recebido
		{
			get { return _st_recebido; }
			set { _st_recebido = value; }
		}

		private String _obs_1;
		public String obs_1
		{
			get { return _obs_1; }
			set { _obs_1 = value; }
		}

		private String _obs_2;
		public String obs_2
		{
			get { return _obs_2; }
			set { _obs_2 = value; }
		}

		private short _qtde_parcelas;
		public short qtde_parcelas
		{
			get { return _qtde_parcelas; }
			set { _qtde_parcelas = value; }
		}

		private String _forma_pagto;
		public String forma_pagto
		{
			get { return _forma_pagto; }
			set { _forma_pagto = value; }
		}

		private decimal _vl_total_familia;
		public decimal vl_total_familia
		{
			get { return _vl_total_familia; }
			set { _vl_total_familia = value; }
		}

		private decimal _vl_pago_familia;
		public decimal vl_pago_familia
		{
			get { return _vl_pago_familia; }
			set { _vl_pago_familia = value; }
		}

		private short _split_status;
		public short split_status
		{
			get { return _split_status; }
			set { _split_status = value; }
		}

		private DateTime _split_data;
		public DateTime split_data
		{
			get { return _split_data; }
			set { _split_data = value; }
		}

		private String _split_hora;
		public String split_hora
		{
			get { return _split_hora; }
			set { _split_hora = value; }
		}

		private String _split_usuario;
		public String split_usuario
		{
			get { return _split_usuario; }
			set { _split_usuario = value; }
		}

		private short _a_entregar_status;
		public short a_entregar_status
		{
			get { return _a_entregar_status; }
			set { _a_entregar_status = value; }
		}

		private DateTime _a_entregar_data_marcada;
		public DateTime a_entregar_data_marcada
		{
			get { return _a_entregar_data_marcada; }
			set { _a_entregar_data_marcada = value; }
		}

		private DateTime _a_entregar_data;
		public DateTime a_entregar_data
		{
			get { return _a_entregar_data; }
			set { _a_entregar_data = value; }
		}

		private String _a_entregar_hora;
		public String a_entregar_hora
		{
			get { return _a_entregar_hora; }
			set { _a_entregar_hora = value; }
		}

		private String _a_entregar_usuario;
		public String a_entregar_usuario
		{
			get { return _a_entregar_usuario; }
			set { _a_entregar_usuario = value; }
		}

		private String _loja_indicou;
		public String loja_indicou
		{
			get { return _loja_indicou; }
			set { _loja_indicou = value; }
		}

		private double _comissao_loja_indicou;
		public double comissao_loja_indicou
		{
			get { return _comissao_loja_indicou; }
			set { _comissao_loja_indicou = value; }
		}

		private short _venda_externa;
		public short venda_externa
		{
			get { return _venda_externa; }
			set { _venda_externa = value; }
		}

		private decimal _vl_frete;
		public decimal vl_frete
		{
			get { return _vl_frete; }
			set { _vl_frete = value; }
		}

		private String _transportadora_id;
		public String transportadora_id
		{
			get { return _transportadora_id; }
			set { _transportadora_id = value; }
		}

		private DateTime _transportadora_data;
		public DateTime transportadora_data
		{
			get { return _transportadora_data; }
			set { _transportadora_data = value; }
		}

		private String _transportadora_usuario;
		public String transportadora_usuario
		{
			get { return _transportadora_usuario; }
			set { _transportadora_usuario = value; }
		}

		private short _analise_credito;
		public short analise_credito
		{
			get { return _analise_credito; }
			set { _analise_credito = value; }
		}

		private DateTime _analise_credito_data;
		public DateTime analise_credito_data
		{
			get { return _analise_credito_data; }
			set { _analise_credito_data = value; }
		}

		private String _analise_credito_usuario;
		public String analise_credito_usuario
		{
			get { return _analise_credito_usuario; }
			set { _analise_credito_usuario = value; }
		}

		private short _tipo_parcelamento;
		public short tipo_parcelamento
		{
			get { return _tipo_parcelamento; }
			set { _tipo_parcelamento = value; }
		}

		private short _av_forma_pagto;
		public short av_forma_pagto
		{
			get { return _av_forma_pagto; }
			set { _av_forma_pagto = value; }
		}

		private short _pc_qtde_parcelas;
		public short pc_qtde_parcelas
		{
			get { return _pc_qtde_parcelas; }
			set { _pc_qtde_parcelas = value; }
		}

		private decimal _pc_valor_parcela;
		public decimal pc_valor_parcela
		{
			get { return _pc_valor_parcela; }
			set { _pc_valor_parcela = value; }
		}

		private short _pce_forma_pagto_entrada;
		public short pce_forma_pagto_entrada
		{
			get { return _pce_forma_pagto_entrada; }
			set { _pce_forma_pagto_entrada = value; }
		}

		private short _pce_forma_pagto_prestacao;
		public short pce_forma_pagto_prestacao
		{
			get { return _pce_forma_pagto_prestacao; }
			set { _pce_forma_pagto_prestacao = value; }
		}

		private decimal _pce_entrada_valor;
		public decimal pce_entrada_valor
		{
			get { return _pce_entrada_valor; }
			set { _pce_entrada_valor = value; }
		}

		private short _pce_prestacao_qtde;
		public short pce_prestacao_qtde
		{
			get { return _pce_prestacao_qtde; }
			set { _pce_prestacao_qtde = value; }
		}

		private decimal _pce_prestacao_valor;
		public decimal pce_prestacao_valor
		{
			get { return _pce_prestacao_valor; }
			set { _pce_prestacao_valor = value; }
		}

		private short _pce_prestacao_periodo;
		public short pce_prestacao_periodo
		{
			get { return _pce_prestacao_periodo; }
			set { _pce_prestacao_periodo = value; }
		}

		private short _pse_forma_pagto_prim_prest;
		public short pse_forma_pagto_prim_prest
		{
			get { return _pse_forma_pagto_prim_prest; }
			set { _pse_forma_pagto_prim_prest = value; }
		}

		private short _pse_forma_pagto_demais_prest;
		public short pse_forma_pagto_demais_prest
		{
			get { return _pse_forma_pagto_demais_prest; }
			set { _pse_forma_pagto_demais_prest = value; }
		}

		private decimal _pse_prim_prest_valor;
		public decimal pse_prim_prest_valor
		{
			get { return _pse_prim_prest_valor; }
			set { _pse_prim_prest_valor = value; }
		}

		private short _pse_prim_prest_apos;
		public short pse_prim_prest_apos
		{
			get { return _pse_prim_prest_apos; }
			set { _pse_prim_prest_apos = value; }
		}

		private short _pse_demais_prest_qtde;
		public short pse_demais_prest_qtde
		{
			get { return _pse_demais_prest_qtde; }
			set { _pse_demais_prest_qtde = value; }
		}

		private decimal _pse_demais_prest_valor;
		public decimal pse_demais_prest_valor
		{
			get { return _pse_demais_prest_valor; }
			set { _pse_demais_prest_valor = value; }
		}

		private short _pse_demais_prest_periodo;
		public short pse_demais_prest_periodo
		{
			get { return _pse_demais_prest_periodo; }
			set { _pse_demais_prest_periodo = value; }
		}

		private short _pu_forma_pagto;
		public short pu_forma_pagto
		{
			get { return _pu_forma_pagto; }
			set { _pu_forma_pagto = value; }
		}

		private decimal _pu_valor;
		public decimal pu_valor
		{
			get { return _pu_valor; }
			set { _pu_valor = value; }
		}

		private short _pu_vencto_apos;
		public short pu_vencto_apos
		{
			get { return _pu_vencto_apos; }
			set { _pu_vencto_apos = value; }
		}

		private String _indicador;
		public String indicador
		{
			get { return _indicador; }
			set { _indicador = value; }
		}

		private String _indicador_desempenho_nota;
		public String indicador_desempenho_nota
		{
			get { return _indicador_desempenho_nota; }
			set { _indicador_desempenho_nota = value; }
		}

		private decimal _vl_total_NF;
		public decimal vl_total_NF
		{
			get { return _vl_total_NF; }
			set { _vl_total_NF = value; }
		}

		private decimal _vl_total_RA;
		public decimal vl_total_RA
		{
			get { return _vl_total_RA; }
			set { _vl_total_RA = value; }
		}

		private double _perc_RT;
		public double perc_RT
		{
			get { return _perc_RT; }
			set { _perc_RT = value; }
		}

		private short _st_orc_virou_pedido;
		public short st_orc_virou_pedido
		{
			get { return _st_orc_virou_pedido; }
			set { _st_orc_virou_pedido = value; }
		}

		private String _orcamento;
		public String orcamento
		{
			get { return _orcamento; }
			set { _orcamento = value; }
		}

		private String _orcamentista;
		public String orcamentista
		{
			get { return _orcamentista; }
			set { _orcamentista = value; }
		}

		private short _comissao_paga;
		public short comissao_paga
		{
			get { return _comissao_paga; }
			set { _comissao_paga = value; }
		}

		private String _comissao_paga_ult_op;
		public String comissao_paga_ult_op
		{
			get { return _comissao_paga_ult_op; }
			set { _comissao_paga_ult_op = value; }
		}

		private DateTime _comissao_paga_data;
		public DateTime comissao_paga_data
		{
			get { return _comissao_paga_data; }
			set { _comissao_paga_data = value; }
		}

		private String _comissao_paga_usuario;
		public String comissao_paga_usuario
		{
			get { return _comissao_paga_usuario; }
			set { _comissao_paga_usuario = value; }
		}

		private double _perc_desagio_RA;
		public double perc_desagio_RA
		{
			get { return _perc_desagio_RA; }
			set { _perc_desagio_RA = value; }
		}

		private double _perc_limite_RA_sem_desagio;
		public double perc_limite_RA_sem_desagio
		{
			get { return _perc_limite_RA_sem_desagio; }
			set { _perc_limite_RA_sem_desagio = value; }
		}

		private decimal _vl_total_RA_liquido;
		public decimal vl_total_RA_liquido
		{
			get { return _vl_total_RA_liquido; }
			set { _vl_total_RA_liquido = value; }
		}

		private short _st_tem_desagio_RA;
		public short st_tem_desagio_RA
		{
			get { return _st_tem_desagio_RA; }
			set { _st_tem_desagio_RA = value; }
		}

		private short _qtde_parcelas_desagio_RA;
		public short qtde_parcelas_desagio_RA
		{
			get { return _qtde_parcelas_desagio_RA; }
			set { _qtde_parcelas_desagio_RA = value; }
		}

		private String _transportadora_num_coleta;
		public String transportadora_num_coleta
		{
			get { return _transportadora_num_coleta; }
			set { _transportadora_num_coleta = value; }
		}

		private String _transportadora_contato;
		public String transportadora_contato
		{
			get { return _transportadora_contato; }
			set { _transportadora_contato = value; }
		}

		private short _st_end_entrega;
		public short st_end_entrega
		{
			get { return _st_end_entrega; }
			set { _st_end_entrega = value; }
		}

		private String _EndEtg_endereco;
		public String endEtg_endereco
		{
			get { return _EndEtg_endereco; }
			set { _EndEtg_endereco = value; }
		}

		private String _EndEtg_endereco_numero;
		public String endEtg_endereco_numero
		{
			get { return _EndEtg_endereco_numero; }
			set { _EndEtg_endereco_numero = value; }
		}

		private String _EndEtg_endereco_complemento;
		public String endEtg_endereco_complemento
		{
			get { return _EndEtg_endereco_complemento; }
			set { _EndEtg_endereco_complemento = value; }
		}

		private String _EndEtg_bairro;
		public String endEtg_bairro
		{
			get { return _EndEtg_bairro; }
			set { _EndEtg_bairro = value; }
		}

		private String _EndEtg_cidade;
		public String endEtg_cidade
		{
			get { return _EndEtg_cidade; }
			set { _EndEtg_cidade = value; }
		}

		private String _EndEtg_uf;
		public String endEtg_uf
		{
			get { return _EndEtg_uf; }
			set { _EndEtg_uf = value; }
		}

		private String _EndEtg_cep;
		public String endEtg_cep
		{
			get { return _EndEtg_cep; }
			set { _EndEtg_cep = value; }
		}

		private short _st_etg_imediata;
		public short st_etg_imediata
		{
			get { return _st_etg_imediata; }
			set { _st_etg_imediata = value; }
		}

		private DateTime _etg_imediata_data;
		public DateTime etg_imediata_data
		{
			get { return _etg_imediata_data; }
			set { _etg_imediata_data = value; }
		}

		private String _etg_imediata_usuario;
		public String etg_imediata_usuario
		{
			get { return _etg_imediata_usuario; }
			set { _etg_imediata_usuario = value; }
		}

		private short _frete_status;
		public short frete_status
		{
			get { return _frete_status; }
			set { _frete_status = value; }
		}

		private decimal _frete_valor;
		public decimal frete_valor
		{
			get { return _frete_valor; }
			set { _frete_valor = value; }
		}

		private DateTime _frete_data;
		public DateTime frete_data
		{
			get { return _frete_data; }
			set { _frete_data = value; }
		}

		private String _frete_usuario;
		public String frete_usuario
		{
			get { return _frete_usuario; }
			set { _frete_usuario = value; }
		}

		private short _StBemUsoConsumo;
		public short stBemUsoConsumo
		{
			get { return _StBemUsoConsumo; }
			set { _StBemUsoConsumo = value; }
		}

		private short _PedidoRecebidoStatus;
		public short pedidoRecebidoStatus
		{
			get { return _PedidoRecebidoStatus; }
			set { _PedidoRecebidoStatus = value; }
		}

		private DateTime _PedidoRecebidoData;
		public DateTime pedidoRecebidoData
		{
			get { return _PedidoRecebidoData; }
			set { _PedidoRecebidoData = value; }
		}

		private String _PedidoRecebidoUsuarioUltAtualiz;
		public String pedidoRecebidoUsuarioUltAtualiz
		{
			get { return _PedidoRecebidoUsuarioUltAtualiz; }
			set { _PedidoRecebidoUsuarioUltAtualiz = value; }
		}

		private DateTime _PedidoRecebidoDtHrUltAtualiz;
		public DateTime pedidoRecebidoDtHrUltAtualiz
		{
			get { return _PedidoRecebidoDtHrUltAtualiz; }
			set { _PedidoRecebidoDtHrUltAtualiz = value; }
		}

		private short _InstaladorInstalaStatus;
		public short instaladorInstalaStatus
		{
			get { return _InstaladorInstalaStatus; }
			set { _InstaladorInstalaStatus = value; }
		}

		private String _InstaladorInstalaUsuarioUltAtualiz;
		public String instaladorInstalaUsuarioUltAtualiz
		{
			get { return _InstaladorInstalaUsuarioUltAtualiz; }
			set { _InstaladorInstalaUsuarioUltAtualiz = value; }
		}

		private DateTime _InstaladorInstalaDtHrUltAtualiz;
		public DateTime instaladorInstalaDtHrUltAtualiz
		{
			get { return _InstaladorInstalaDtHrUltAtualiz; }
			set { _InstaladorInstalaDtHrUltAtualiz = value; }
		}

		private String _custoFinancFornecTipoParcelamento;
		public String custoFinancFornecTipoParcelamento
		{
			get { return _custoFinancFornecTipoParcelamento; }
			set { _custoFinancFornecTipoParcelamento = value; }
		}

		private short _custoFinancFornecQtdeParcelas;
		public short custoFinancFornecQtdeParcelas
		{
			get { return _custoFinancFornecQtdeParcelas; }
			set { _custoFinancFornecQtdeParcelas = value; }
		}

		#region [ Campos calculados ]
		private decimal _vlTotalPrecoNfDestePedido;
		public decimal vlTotalPrecoNfDestePedido
		{
			get { return _vlTotalPrecoNfDestePedido; }
			set { _vlTotalPrecoNfDestePedido = value; }
		}

		private decimal _vlTotalBoletoDestePedido;
		public decimal vlTotalBoletoDestePedido
		{
			get { return _vlTotalBoletoDestePedido; }
			set { _vlTotalBoletoDestePedido = value; }
		}

		private decimal _vlTotalFormaPagtoDestePedido;
		public decimal vlTotalFormaPagtoDestePedido
		{
			get { return _vlTotalFormaPagtoDestePedido; }
			set { _vlTotalFormaPagtoDestePedido = value; }
		}

		private decimal _vlTotalPrecoVendaDestePedido;
		public decimal vlTotalPrecoVendaDestePedido
		{
			get { return _vlTotalPrecoVendaDestePedido; }
			set { _vlTotalPrecoVendaDestePedido = value; }
		}

		private decimal _vlTotalFamiliaPago;
		public decimal vlTotalFamiliaPago
		{
			get { return _vlTotalFamiliaPago; }
			set { _vlTotalFamiliaPago = value; }
		}

		private decimal _vlTotalFamiliaPrecoVenda;
		public decimal vlTotalFamiliaPrecoVenda
		{
			get { return _vlTotalFamiliaPrecoVenda; }
			set { _vlTotalFamiliaPrecoVenda = value; }
		}

		private decimal _vlTotalFamiliaPrecoNF;
		public decimal vlTotalFamiliaPrecoNF
		{
			get { return _vlTotalFamiliaPrecoNF; }
			set { _vlTotalFamiliaPrecoNF = value; }
		}

		private decimal _vlTotalFamiliaDevolucaoPrecoVenda;
		public decimal vlTotalFamiliaDevolucaoPrecoVenda
		{
			get { return _vlTotalFamiliaDevolucaoPrecoVenda; }
			set { _vlTotalFamiliaDevolucaoPrecoVenda = value; }
		}

		private decimal _vlTotalFamiliaDevolucaoPrecoNF;
		public decimal vlTotalFamiliaDevolucaoPrecoNF
		{
			get { return _vlTotalFamiliaDevolucaoPrecoNF; }
			set { _vlTotalFamiliaDevolucaoPrecoNF = value; }
		}
		#endregion

		private int _id_nfe_emitente;
		public int id_nfe_emitente
		{
			get { return _id_nfe_emitente; }
			set { _id_nfe_emitente = value; }
		}

		private string _marketplace_codigo_origem;
		public string marketplace_codigo_origem
		{
			get { return _marketplace_codigo_origem; }
			set { _marketplace_codigo_origem = value; }
		}

		private byte _MarketplacePedidoRecebidoRegistrarStatus;
		public byte MarketplacePedidoRecebidoRegistrarStatus
		{
			get { return _MarketplacePedidoRecebidoRegistrarStatus; }
			set { _MarketplacePedidoRecebidoRegistrarStatus = value; }
		}

		private DateTime _MarketplacePedidoRecebidoRegistrarDataRecebido;
		public DateTime MarketplacePedidoRecebidoRegistrarDataRecebido
		{
			get { return _MarketplacePedidoRecebidoRegistrarDataRecebido; }
			set { _MarketplacePedidoRecebidoRegistrarDataRecebido = value; }
		}

		private DateTime _MarketplacePedidoRecebidoRegistrarDataHora;
		public DateTime MarketplacePedidoRecebidoRegistrarDataHora
		{
			get { return _MarketplacePedidoRecebidoRegistrarDataHora; }
			set { _MarketplacePedidoRecebidoRegistrarDataHora = value; }
		}

		private string _MarketplacePedidoRecebidoRegistrarUsuario;
		public string MarketplacePedidoRecebidoRegistrarUsuario
		{
			get { return _MarketplacePedidoRecebidoRegistrarUsuario; }
			set { _MarketplacePedidoRecebidoRegistrarUsuario = value; }
		}

		private byte _MarketplacePedidoRecebidoRegistradoStatus;
		public byte MarketplacePedidoRecebidoRegistradoStatus
		{
			get { return _MarketplacePedidoRecebidoRegistradoStatus; }
			set { _MarketplacePedidoRecebidoRegistradoStatus = value; }
		}

		private DateTime _MarketplacePedidoRecebidoRegistradoDataHora;
		public DateTime MarketplacePedidoRecebidoRegistradoDataHora
		{
			get { return _MarketplacePedidoRecebidoRegistradoDataHora; }
			set { _MarketplacePedidoRecebidoRegistradoDataHora = value; }
		}

		private string _MarketplacePedidoRecebidoRegistradoUsuario;
		public string MarketplacePedidoRecebidoRegistradoUsuario
		{
			get { return _MarketplacePedidoRecebidoRegistradoUsuario; }
			set { _MarketplacePedidoRecebidoRegistradoUsuario = value; }
		}
		#endregion
	}

	public class PedidoItem
	{
		#region [ Getters/Setters ]

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private String _fabricante;
		public String fabricante
		{
			get { return _fabricante; }
			set { _fabricante = value; }
		}

		private String _produto;
		public String produto
		{
			get { return _produto; }
			set { _produto = value; }
		}

		private short _qtde;
		public short qtde
		{
			get { return _qtde; }
			set { _qtde = value; }
		}

		private double _desc_dado;
		public double desc_dado
		{
			get { return _desc_dado; }
			set { _desc_dado = value; }
		}

		private decimal _preco_venda;
		public decimal preco_venda
		{
			get { return _preco_venda; }
			set { _preco_venda = value; }
		}

		private decimal _preco_fabricante;
		public decimal preco_fabricante
		{
			get { return _preco_fabricante; }
			set { _preco_fabricante = value; }
		}

		private decimal _preco_lista;
		public decimal preco_lista
		{
			get { return _preco_lista; }
			set { _preco_lista = value; }
		}

		private double _margem;
		public double margem
		{
			get { return _margem; }
			set { _margem = value; }
		}

		private double _desc_max;
		public double desc_max
		{
			get { return _desc_max; }
			set { _desc_max = value; }
		}

		private double _comissao;
		public double comissao
		{
			get { return _comissao; }
			set { _comissao = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private String _ean;
		public String ean
		{
			get { return _ean; }
			set { _ean = value; }
		}

		private String _grupo;
		public String grupo
		{
			get { return _grupo; }
			set { _grupo = value; }
		}

		private double _peso;
		public double peso
		{
			get { return _peso; }
			set { _peso = value; }
		}

		private short _qtde_volumes;
		public short qtde_volumes
		{
			get { return _qtde_volumes; }
			set { _qtde_volumes = value; }
		}

		private short _abaixo_min_status;
		public short abaixo_min_status
		{
			get { return _abaixo_min_status; }
			set { _abaixo_min_status = value; }
		}

		private String _abaixo_min_autorizacao;
		public String abaixo_min_autorizacao
		{
			get { return _abaixo_min_autorizacao; }
			set { _abaixo_min_autorizacao = value; }
		}

		private String _abaixo_min_autorizador;
		public String abaixo_min_autorizador
		{
			get { return _abaixo_min_autorizador; }
			set { _abaixo_min_autorizador = value; }
		}

		private short _sequencia;
		public short sequencia
		{
			get { return _sequencia; }
			set { _sequencia = value; }
		}

		private double _markup_fabricante;
		public double markup_fabricante
		{
			get { return _markup_fabricante; }
			set { _markup_fabricante = value; }
		}

		private decimal _preco_NF;
		public decimal preco_NF
		{
			get { return _preco_NF; }
			set { _preco_NF = value; }
		}

		private String _abaixo_min_superv_autorizador;
		public String abaixo_min_superv_autorizador
		{
			get { return _abaixo_min_superv_autorizador; }
			set { _abaixo_min_superv_autorizador = value; }
		}

		private decimal _vl_custo2;
		public decimal vl_custo2
		{
			get { return _vl_custo2; }
			set { _vl_custo2 = value; }
		}

		private String _descricao_html;
		public String descricao_html
		{
			get { return _descricao_html; }
			set { _descricao_html = value; }
		}

		private double _custoFinancFornecCoeficiente;
		public double custoFinancFornecCoeficiente
		{
			get { return _custoFinancFornecCoeficiente; }
			set { _custoFinancFornecCoeficiente = value; }
		}

		private decimal _custoFinancFornecPrecoListaBase;
		public decimal custoFinancFornecPrecoListaBase
		{
			get { return _custoFinancFornecPrecoListaBase; }
			set { _custoFinancFornecPrecoListaBase = value; }
		}

		#endregion
	}

	public class PedidoItemDevolvido
	{
		#region [ Getters/Setters ]

		private String _id;
		public String id
		{
			get { return _id; }
			set { _id = value; }
		}

		private DateTime _devolucao_data;
		public DateTime devolucao_data
		{
			get { return _devolucao_data; }
			set { _devolucao_data = value; }
		}

		private String _devolucao_hora;
		public String devolucao_hora
		{
			get { return _devolucao_hora; }
			set { _devolucao_hora = value; }
		}

		private String _devolucao_usuario;
		public String devolucao_usuario
		{
			get { return _devolucao_usuario; }
			set { _devolucao_usuario = value; }
		}

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private String _fabricante;
		public String fabricante
		{
			get { return _fabricante; }
			set { _fabricante = value; }
		}

		private String _produto;
		public String produto
		{
			get { return _produto; }
			set { _produto = value; }
		}

		private short _qtde;
		public short qtde
		{
			get { return _qtde; }
			set { _qtde = value; }
		}

		private double _desc_dado;
		public double desc_dado
		{
			get { return _desc_dado; }
			set { _desc_dado = value; }
		}

		private decimal _preco_venda;
		public decimal preco_venda
		{
			get { return _preco_venda; }
			set { _preco_venda = value; }
		}

		private decimal _preco_fabricante;
		public decimal preco_fabricante
		{
			get { return _preco_fabricante; }
			set { _preco_fabricante = value; }
		}

		private decimal _preco_lista;
		public decimal preco_lista
		{
			get { return _preco_lista; }
			set { _preco_lista = value; }
		}

		private double _margem;
		public double margem
		{
			get { return _margem; }
			set { _margem = value; }
		}

		private double _desc_max;
		public double desc_max
		{
			get { return _desc_max; }
			set { _desc_max = value; }
		}

		private double _comissao;
		public double comissao
		{
			get { return _comissao; }
			set { _comissao = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private String _ean;
		public String ean
		{
			get { return _ean; }
			set { _ean = value; }
		}

		private String _grupo;
		public String grupo
		{
			get { return _grupo; }
			set { _grupo = value; }
		}

		private double _peso;
		public double peso
		{
			get { return _peso; }
			set { _peso = value; }
		}

		private short _qtde_volumes;
		public short qtde_volumes
		{
			get { return _qtde_volumes; }
			set { _qtde_volumes = value; }
		}

		private short _abaixo_min_status;
		public short abaixo_min_status
		{
			get { return _abaixo_min_status; }
			set { _abaixo_min_status = value; }
		}

		private String _abaixo_min_autorizacao;
		public String abaixo_min_autorizacao
		{
			get { return _abaixo_min_autorizacao; }
			set { _abaixo_min_autorizacao = value; }
		}

		private String _abaixo_min_autorizador;
		public String abaixo_min_autorizador
		{
			get { return _abaixo_min_autorizador; }
			set { _abaixo_min_autorizador = value; }
		}

		private double _markup_fabricante;
		public double markup_fabricante
		{
			get { return _markup_fabricante; }
			set { _markup_fabricante = value; }
		}

		private String _motivo;
		public String motivo
		{
			get { return _motivo; }
			set { _motivo = value; }
		}

		private decimal _preco_NF;
		public decimal preco_NF
		{
			get { return _preco_NF; }
			set { _preco_NF = value; }
		}

		private short _comissao_descontada;
		public short comissao_descontada
		{
			get { return _comissao_descontada; }
			set { _comissao_descontada = value; }
		}

		private String _comissao_descontada_ult_op;
		public String comissao_descontada_ult_op
		{
			get { return _comissao_descontada_ult_op; }
			set { _comissao_descontada_ult_op = value; }
		}

		private DateTime _comissao_descontada_data;
		public DateTime comissao_descontada_data
		{
			get { return _comissao_descontada_data; }
			set { _comissao_descontada_data = value; }
		}

		private String _comissao_descontada_usuario;
		public String comissao_descontada_usuario
		{
			get { return _comissao_descontada_usuario; }
			set { _comissao_descontada_usuario = value; }
		}

		private String _abaixo_min_superv_autorizador;
		public String abaixo_min_superv_autorizador
		{
			get { return _abaixo_min_superv_autorizador; }
			set { _abaixo_min_superv_autorizador = value; }
		}

		private decimal _vl_custo2;
		public decimal vl_custo2
		{
			get { return _vl_custo2; }
			set { _vl_custo2 = value; }
		}

		private String _descricao_html;
		public String descricao_html
		{
			get { return _descricao_html; }
			set { _descricao_html = value; }
		}

		private double _custoFinancFornecCoeficiente;
		public double custoFinancFornecCoeficiente
		{
			get { return _custoFinancFornecCoeficiente; }
			set { _custoFinancFornecCoeficiente = value; }
		}

		private decimal _custoFinancFornecPrecoListaBase;
		public decimal custoFinancFornecPrecoListaBase
		{
			get { return _custoFinancFornecPrecoListaBase; }
			set { _custoFinancFornecPrecoListaBase = value; }
		}

		#endregion
	}
}
