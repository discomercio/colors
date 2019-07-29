#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	#region [ Boleto ]
	class Boleto
	{
		#region [ Construtor ]
		public Boleto()
		{
			listaBoletoItem = new List<BoletoItem>();
		}
		#endregion

		#region [ Getters/Setters ]

		public List<BoletoItem> listaBoletoItem;

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private String _id_cliente = "";
		public String id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private int _id_nf_parcela_pagto;
		public int id_nf_parcela_pagto
		{
			get { return _id_nf_parcela_pagto; }
			set { _id_nf_parcela_pagto = value; }
		}

		private byte _tipo_vinculo;
		public byte tipo_vinculo
		{
			get { return _tipo_vinculo; }
			set { _tipo_vinculo = value; }
		}

		private short _status;
		public short status
		{
			get { return _status; }
			set { _status = value; }
		}

		private DateTime _dt_remessa;
		public DateTime dt_remessa
		{
			get { return _dt_remessa; }
			set { _dt_remessa = value; }
		}

		private int _id_boleto_arq_remessa;
		public int id_boleto_arq_remessa
		{
			get { return _id_boleto_arq_remessa; }
			set { _id_boleto_arq_remessa = value; }
		}

		private int _numero_NF;
		public int numero_NF
		{
			get { return _numero_NF; }
			set { _numero_NF = value; }
		}

		private byte _qtde_parcelas;
		public byte qtde_parcelas
		{
			get { return _qtde_parcelas; }
			set { _qtde_parcelas = value; }
		}

		private short _id_boleto_cedente;
		public short id_boleto_cedente
		{
			get { return _id_boleto_cedente; }
			set { _id_boleto_cedente = value; }
		}

		private String _codigo_empresa = "";
		public String codigo_empresa
		{
			get { return _codigo_empresa; }
			set { _codigo_empresa = value; }
		}

		private String _nome_empresa = "";
		public String nome_empresa
		{
			get { return _nome_empresa; }
			set { _nome_empresa = value; }
		}

		private String _num_banco = "";
		public String num_banco
		{
			get { return _num_banco; }
			set { _num_banco = value; }
		}

		private String _nome_banco = "";
		public String nome_banco
		{
			get { return _nome_banco; }
			set { _nome_banco = value; }
		}

		private String _agencia = "";
		public String agencia
		{
			get { return _agencia; }
			set { _agencia = value; }
		}

		private String _digito_agencia = "";
		public String digito_agencia
		{
			get { return _digito_agencia; }
			set { _digito_agencia = value; }
		}

		private String _conta = "";
		public String conta
		{
			get { return _conta; }
			set { _conta = value; }
		}

		private String _digito_conta = "";
		public String digito_conta
		{
			get { return _digito_conta; }
			set { _digito_conta = value; }
		}

		private String _carteira = "";
		public String carteira
		{
			get { return _carteira; }
			set { _carteira = value; }
		}

		private double _juros_mora;
		public double juros_mora
		{
			get { return _juros_mora; }
			set { _juros_mora = value; }
		}

		private double _perc_multa;
		public double perc_multa
		{
			get { return _perc_multa; }
			set { _perc_multa = value; }
		}

		private String _primeira_instrucao = "";
		public String primeira_instrucao
		{
			get { return _primeira_instrucao; }
			set { _primeira_instrucao = value; }
		}

		private String _segunda_instrucao = "";
		public String segunda_instrucao
		{
			get { return _segunda_instrucao; }
			set { _segunda_instrucao = value; }
		}

		private short _qtde_dias_protesto;
		public short qtde_dias_protesto
		{
			get { return _qtde_dias_protesto; }
			set { _qtde_dias_protesto = value; }
		}

		private short _qtde_dias_decurso_prazo;
		public short qtde_dias_decurso_prazo
		{
			get { return _qtde_dias_decurso_prazo; }
			set { _qtde_dias_decurso_prazo = value; }
		}

		private String _tipo_sacado = "";
		public String tipo_sacado
		{
			get { return _tipo_sacado; }
			set { _tipo_sacado = value; }
		}

		private String _num_inscricao_sacado = "";
		public String num_inscricao_sacado
		{
			get { return _num_inscricao_sacado; }
			set { _num_inscricao_sacado = value; }
		}

		private String _nome_sacado = "";
		public String nome_sacado
		{
			get { return _nome_sacado; }
			set { _nome_sacado = value; }
		}

		private String _endereco_sacado = "";
		public String endereco_sacado
		{
			get { return _endereco_sacado; }
			set { _endereco_sacado = value; }
		}

		private String _cep_sacado = "";
		public String cep_sacado
		{
			get { return _cep_sacado; }
			set { _cep_sacado = value; }
		}

		private String _bairro_sacado = "";
		public String bairro_sacado
		{
			get { return _bairro_sacado; }
			set { _bairro_sacado = value; }
		}

		private String _cidade_sacado = "";
		public String cidade_sacado
		{
			get { return _cidade_sacado; }
			set { _cidade_sacado = value; }
		}

		private String _uf_sacado = "";
		public String uf_sacado
		{
			get { return _uf_sacado; }
			set { _uf_sacado = value; }
		}

		private String _email_sacado = "";
		public String email_sacado
		{
			get { return _email_sacado; }
			set { _email_sacado = value; }
		}

		private String _segunda_mensagem = "";
		public String segunda_mensagem
		{
			get { return _segunda_mensagem; }
			set { _segunda_mensagem = value; }
		}

		private String _mensagem_1 = "";
		public String mensagem_1
		{
			get { return _mensagem_1; }
			set { _mensagem_1 = value; }
		}

		private String _mensagem_2 = "";
		public String mensagem_2
		{
			get { return _mensagem_2; }
			set { _mensagem_2 = value; }
		}

		private String _mensagem_3 = "";
		public String mensagem_3
		{
			get { return _mensagem_3; }
			set { _mensagem_3 = value; }
		}

		private String _mensagem_4 = "";
		public String mensagem_4
		{
			get { return _mensagem_4; }
			set { _mensagem_4 = value; }
		}

		private DateTime _dt_cadastro;
		public DateTime dt_cadastro
		{
			get { return _dt_cadastro; }
			set { _dt_cadastro = value; }
		}

		private DateTime _dt_hr_cadastro;
		public DateTime dt_hr_cadastro
		{
			get { return _dt_hr_cadastro; }
			set { _dt_hr_cadastro = value; }
		}

		private String _usuario_cadastro = "";
		public String usuario_cadastro
		{
			get { return _usuario_cadastro; }
			set { _usuario_cadastro = value; }
		}

		private DateTime _dt_ult_atualizacao;
		public DateTime dt_ult_atualizacao
		{
			get { return _dt_ult_atualizacao; }
			set { _dt_ult_atualizacao = value; }
		}

		private DateTime _dt_hr_ult_atualizacao;
		public DateTime dt_hr_ult_atualizacao
		{
			get { return _dt_hr_ult_atualizacao; }
			set { _dt_hr_ult_atualizacao = value; }
		}

		private String _usuario_ult_atualizacao = "";
		public String usuario_ult_atualizacao
		{
			get { return _usuario_ult_atualizacao; }
			set { _usuario_ult_atualizacao = value; }
		}

		private int _num_documento_boleto_avulso;
		public int num_documento_boleto_avulso
		{
			get { return _num_documento_boleto_avulso; }
			set { _num_documento_boleto_avulso = value; }
		}

		#endregion
	}
	#endregion

	#region [ BoletoItem ]
	class BoletoItem
	{
		#region [ Construtor ]
		public BoletoItem()
		{
			listaBoletoItemRateio = new List<BoletoItemRateio>();
		}
		#endregion

		#region [ Getters / Setters ]

		public List<BoletoItemRateio> listaBoletoItemRateio;

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_boleto;
		public int id_boleto
		{
			get { return _id_boleto; }
			set { _id_boleto = value; }
		}

		private byte _num_parcela;
		public byte num_parcela
		{
			get { return _num_parcela; }
			set { _num_parcela = value; }
		}

		private short _status;
		public short status
		{
			get { return _status; }
			set { _status = value; }
		}

		private byte _tipo_vencimento;
		public byte tipo_vencimento
		{
			get { return _tipo_vencimento; }
			set { _tipo_vencimento = value; }
		}

		private DateTime _dt_vencto;
		public DateTime dt_vencto
		{
			get { return _dt_vencto; }
			set { _dt_vencto = value; }
		}

		private decimal _valor;
		public decimal valor
		{
			get { return _valor; }
			set { _valor = value; }
		}

		private decimal _bonificacao_por_dia;
		public decimal bonificacao_por_dia
		{
			get { return _bonificacao_por_dia; }
			set { _bonificacao_por_dia = value; }
		}

		private decimal _valor_por_dia_atraso;
		public decimal valor_por_dia_atraso
		{
			get { return _valor_por_dia_atraso; }
			set { _valor_por_dia_atraso = value; }
		}

		private DateTime _dt_limite_desconto;
		public DateTime dt_limite_desconto
		{
			get { return _dt_limite_desconto; }
			set { _dt_limite_desconto = value; }
		}

		private decimal _valor_desconto;
		public decimal valor_desconto
		{
			get { return _valor_desconto; }
			set { _valor_desconto = value; }
		}

		private String _numero_documento = "";
		public String numero_documento
		{
			get { return _numero_documento; }
			set { _numero_documento = value; }
		}

		private String _nosso_numero = "";
		public String nosso_numero
		{
			get { return _nosso_numero; }
			set { _nosso_numero = value; }
		}

		private String _digito_nosso_numero = "";
		public String digito_nosso_numero
		{
			get { return _digito_nosso_numero; }
			set { _digito_nosso_numero = value; }
		}

		private String _primeira_mensagem = "";
		public String primeira_mensagem
		{
			get { return _primeira_mensagem; }
			set { _primeira_mensagem = value; }
		}

		private String _num_controle_participante = "";
		public String num_controle_participante
		{
			get { return _num_controle_participante; }
			set { _num_controle_participante = value; }
		}

		private int _num_sequencial_registro;
		public int num_sequencial_registro
		{
			get { return _num_sequencial_registro; }
			set { _num_sequencial_registro = value; }
		}

		private byte _st_instrucao_protesto;
		public byte st_instrucao_protesto
		{
			get { return _st_instrucao_protesto; }
			set { _st_instrucao_protesto = value; }
		}
		#endregion
	}
	#endregion

	#region [ BoletoItemRateio ]
	class BoletoItemRateio
	{
		#region [ Getters / Setters ]

		private int _id_boleto_item;
		public int id_boleto_item
		{
			get { return _id_boleto_item; }
			set { _id_boleto_item = value; }
		}

		private String _pedido = "";
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private int _id_boleto;
		public int id_boleto
		{
			get { return _id_boleto; }
			set { _id_boleto = value; }
		}

		private decimal _valor;
		public decimal valor
		{
			get { return _valor; }
			set { _valor = value; }
		}

		#endregion
	}
	#endregion

	#region [ BoletoArqRemessa ]
	class BoletoArqRemessa
	{
		#region [ Getters / Setters ]

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _nsu_arq_remessa;
		public int nsu_arq_remessa
		{
			get { return _nsu_arq_remessa; }
			set { _nsu_arq_remessa = value; }
		}

		private DateTime _dt_geracao;
		public DateTime dt_geracao
		{
			get { return _dt_geracao; }
			set { _dt_geracao = value; }
		}

		private DateTime _dt_hr_geracao;
		public DateTime dt_hr_geracao
		{
			get { return _dt_hr_geracao; }
			set { _dt_hr_geracao = value; }
		}

		private String _usuario_geracao;
		public String usuario_geracao
		{
			get { return _usuario_geracao; }
			set { _usuario_geracao = value; }
		}

		private int _qtde_registros;
		public int qtde_registros
		{
			get { return _qtde_registros; }
			set { _qtde_registros = value; }
		}

		private int _qtde_serie_boletos;
		public int qtde_serie_boletos
		{
			get { return _qtde_serie_boletos; }
			set { _qtde_serie_boletos = value; }
		}

		private short _id_boleto_cedente;
		public short id_boleto_cedente
		{
			get { return _id_boleto_cedente; }
			set { _id_boleto_cedente = value; }
		}

		private String _codigo_empresa = "";
		public String codigo_empresa
		{
			get { return _codigo_empresa; }
			set { _codigo_empresa = value; }
		}

		private String _nome_empresa = "";
		public String nome_empresa
		{
			get { return _nome_empresa; }
			set { _nome_empresa = value; }
		}

		private String _num_banco = "";
		public String num_banco
		{
			get { return _num_banco; }
			set { _num_banco = value; }
		}

		private String _nome_banco = "";
		public String nome_banco
		{
			get { return _nome_banco; }
			set { _nome_banco = value; }
		}

		private String _agencia = "";
		public String agencia
		{
			get { return _agencia; }
			set { _agencia = value; }
		}

		private String _digito_agencia = "";
		public String digito_agencia
		{
			get { return _digito_agencia; }
			set { _digito_agencia = value; }
		}

		private String _conta = "";
		public String conta
		{
			get { return _conta; }
			set { _conta = value; }
		}

		private String _digito_conta = "";
		public String digito_conta
		{
			get { return _digito_conta; }
			set { _digito_conta = value; }
		}

		private String _carteira = "";
		public String carteira
		{
			get { return _carteira; }
			set { _carteira = value; }
		}

		private decimal _vl_total;
		public decimal vl_total
		{
			get { return _vl_total; }
			set { _vl_total = value; }
		}

		private int _duracao_proc_em_seg;
		public int duracao_proc_em_seg
		{
			get { return _duracao_proc_em_seg; }
			set { _duracao_proc_em_seg = value; }
		}

		private String _nome_arq_remessa = "";
		public String nome_arq_remessa
		{
			get { return _nome_arq_remessa; }
			set { _nome_arq_remessa = value; }
		}

		private String _caminho_arq_remessa = "";
		public String caminho_arq_remessa
		{
			get { return _caminho_arq_remessa; }
			set { _caminho_arq_remessa = value; }
		}

		private short _st_geracao;
		public short st_geracao
		{
			get { return _st_geracao; }
			set { _st_geracao = value; }
		}

		private String _msg_erro_geracao = "";
		public String msg_erro_geracao
		{
			get { return _msg_erro_geracao; }
			set { _msg_erro_geracao = value; }
		}

		#endregion
	}
	#endregion

	#region [ BoletoArqRetorno ]
	class BoletoArqRetorno
	{
		#region [ Getters / Setters ]

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_boleto_cedente;
		public int id_boleto_cedente
		{
			get { return _id_boleto_cedente; }
			set { _id_boleto_cedente = value; }
		}

		private DateTime _dt_processamento;
		public DateTime dt_processamento
		{
			get { return _dt_processamento; }
			set { _dt_processamento = value; }
		}

		private DateTime _dt_hr_processamento;
		public DateTime dt_hr_processamento
		{
			get { return _dt_hr_processamento; }
			set { _dt_hr_processamento = value; }
		}

		private String _usuario_processamento;
		public String usuario_processamento
		{
			get { return _usuario_processamento; }
			set { _usuario_processamento = value; }
		}

		private int _qtde_registros;
		public int qtde_registros
		{
			get { return _qtde_registros; }
			set { _qtde_registros = value; }
		}

		private String _codigo_empresa = "";
		public String codigo_empresa
		{
			get { return _codigo_empresa; }
			set { _codigo_empresa = value; }
		}

		private String _nome_empresa = "";
		public String nome_empresa
		{
			get { return _nome_empresa; }
			set { _nome_empresa = value; }
		}

		private String _num_banco = "";
		public String num_banco
		{
			get { return _num_banco; }
			set { _num_banco = value; }
		}

		private String _nome_banco = "";
		public String nome_banco
		{
			get { return _nome_banco; }
			set { _nome_banco = value; }
		}

		private String _data_gravacao_arquivo = "";
		public String data_gravacao_arquivo
		{
			get { return _data_gravacao_arquivo; }
			set { _data_gravacao_arquivo = value; }
		}

		private String _numero_aviso_bancario = "";
		public String numero_aviso_bancario
		{
			get { return _numero_aviso_bancario; }
			set { _numero_aviso_bancario = value; }
		}

		private String _data_credito = "";
		public String data_credito
		{
			get { return _data_credito; }
			set { _data_credito = value; }
		}

		private String _qtdeTitulosEmCobranca = "";
		public String qtdeTitulosEmCobranca
		{
			get { return _qtdeTitulosEmCobranca; }
			set { _qtdeTitulosEmCobranca = value; }
		}

		private String _valorTotalEmCobranca = "";
		public String valorTotalEmCobranca
		{
			get { return _valorTotalEmCobranca; }
			set { _valorTotalEmCobranca = value; }
		}

		private String _qtdeRegsOcorrencia02ConfirmacaoEntradas = "";
		public String qtdeRegsOcorrencia02ConfirmacaoEntradas
		{
			get { return _qtdeRegsOcorrencia02ConfirmacaoEntradas; }
			set { _qtdeRegsOcorrencia02ConfirmacaoEntradas = value; }
		}

		private String _valorRegsOcorrencia02ConfirmacaoEntradas = "";
		public String valorRegsOcorrencia02ConfirmacaoEntradas
		{
			get { return _valorRegsOcorrencia02ConfirmacaoEntradas; }
			set { _valorRegsOcorrencia02ConfirmacaoEntradas = value; }
		}

		private String _valorRegsOcorrencia06Liquidacao = "";
		public String valorRegsOcorrencia06Liquidacao
		{
			get { return _valorRegsOcorrencia06Liquidacao; }
			set { _valorRegsOcorrencia06Liquidacao = value; }
		}

		private String _qtdeRegsOcorrencia06Liquidacao = "";
		public String qtdeRegsOcorrencia06Liquidacao
		{
			get { return _qtdeRegsOcorrencia06Liquidacao; }
			set { _qtdeRegsOcorrencia06Liquidacao = value; }
		}

		private String _valorRegsOcorrencia06 = "";
		public String valorRegsOcorrencia06
		{
			get { return _valorRegsOcorrencia06; }
			set { _valorRegsOcorrencia06 = value; }
		}

		private String _qtdeRegsOcorrencia09e10TitulosBaixados = "";
		public String qtdeRegsOcorrencia09e10TitulosBaixados
		{
			get { return _qtdeRegsOcorrencia09e10TitulosBaixados; }
			set { _qtdeRegsOcorrencia09e10TitulosBaixados = value; }
		}

		private String _valorRegsOcorrencia09e10TitulosBaixados = "";
		public String valorRegsOcorrencia09e10TitulosBaixados
		{
			get { return _valorRegsOcorrencia09e10TitulosBaixados; }
			set { _valorRegsOcorrencia09e10TitulosBaixados = value; }
		}

		private String _qtdeRegsOcorrencia13AbatimentoCancelado = "";
		public String qtdeRegsOcorrencia13AbatimentoCancelado
		{
			get { return _qtdeRegsOcorrencia13AbatimentoCancelado; }
			set { _qtdeRegsOcorrencia13AbatimentoCancelado = value; }
		}

		private String _valorRegsOcorrencia13AbatimentoCancelado = "";
		public String valorRegsOcorrencia13AbatimentoCancelado
		{
			get { return _valorRegsOcorrencia13AbatimentoCancelado; }
			set { _valorRegsOcorrencia13AbatimentoCancelado = value; }
		}

		private String _qtdeRegsOcorrencia14VenctoAlterado = "";
		public String qtdeRegsOcorrencia14VenctoAlterado
		{
			get { return _qtdeRegsOcorrencia14VenctoAlterado; }
			set { _qtdeRegsOcorrencia14VenctoAlterado = value; }
		}

		private String _valorRegsOcorrencia14VenctoAlterado = "";
		public String valorRegsOcorrencia14VenctoAlterado
		{
			get { return _valorRegsOcorrencia14VenctoAlterado; }
			set { _valorRegsOcorrencia14VenctoAlterado = value; }
		}

		private String _qtdeRegsOcorrencia12AbatimentoConcedido = "";
		public String qtdeRegsOcorrencia12AbatimentoConcedido
		{
			get { return _qtdeRegsOcorrencia12AbatimentoConcedido; }
			set { _qtdeRegsOcorrencia12AbatimentoConcedido = value; }
		}

		private String _valorRegsOcorrencia12AbatimentoConcedido = "";
		public String valorRegsOcorrencia12AbatimentoConcedido
		{
			get { return _valorRegsOcorrencia12AbatimentoConcedido; }
			set { _valorRegsOcorrencia12AbatimentoConcedido = value; }
		}

		private String _qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto = "";
		public String qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto
		{
			get { return _qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto; }
			set { _qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto = value; }
		}

		private String _valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto = "";
		public String valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto
		{
			get { return _valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto; }
			set { _valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto = value; }
		}

		private String _valorTotalRateiosEfetuados = "";
		public String valorTotalRateiosEfetuados
		{
			get { return _valorTotalRateiosEfetuados; }
			set { _valorTotalRateiosEfetuados = value; }
		}

		private String _qtdeTotalRateiosEfetuados = "";
		public String qtdeTotalRateiosEfetuados
		{
			get { return _qtdeTotalRateiosEfetuados; }
			set { _qtdeTotalRateiosEfetuados = value; }
		}

		private int _duracao_proc_em_seg;
		public int duracao_proc_em_seg
		{
			get { return _duracao_proc_em_seg; }
			set { _duracao_proc_em_seg = value; }
		}

		private String _nome_arq_retorno = "";
		public String nome_arq_retorno
		{
			get { return _nome_arq_retorno; }
			set { _nome_arq_retorno = value; }
		}

		private String _caminho_arq_retorno = "";
		public String caminho_arq_retorno
		{
			get { return _caminho_arq_retorno; }
			set { _caminho_arq_retorno = value; }
		}

		private short _st_processamento;
		public short st_processamento
		{
			get { return _st_processamento; }
			set { _st_processamento = value; }
		}

		private String _msg_erro_processamento = "";
		public String msg_erro_processamento
		{
			get { return _msg_erro_processamento; }
			set { _msg_erro_processamento = value; }
		}

		#endregion
	}
	#endregion

	#region [ BoletoPlanoContasDestino ]
	class BoletoPlanoContasDestino
	{
		private byte _id_plano_contas_empresa;
		public byte id_plano_contas_empresa
		{
			get { return _id_plano_contas_empresa; }
			set { _id_plano_contas_empresa = value; }
		}

		private short _id_plano_contas_grupo;
		public short id_plano_contas_grupo
		{
			get { return _id_plano_contas_grupo; }
			set { _id_plano_contas_grupo = value; }
		}

		private int _id_plano_contas_conta;
		public int id_plano_contas_conta
		{
			get { return _id_plano_contas_conta; }
			set { _id_plano_contas_conta = value; }
		}

		private char _natureza;
		public char natureza
		{
			get { return _natureza; }
			set { _natureza = value; }
		}
	}
	#endregion
}
