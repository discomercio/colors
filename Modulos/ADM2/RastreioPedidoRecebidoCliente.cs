using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
	#region [ RastreioPedidoRecebidoCliente ]
	class RastreioPedidoRecebidoCliente
	{
		public RastreioPedidoRecebidoClienteRaw dadosRaw { get; set; }
		public RastreioPedidoRecebidoClienteNormalizado dadosNormalizado { get; set; }
		public RastreioPedidoRecebidoClienteProcesso processo { get; set; }

		#region [ Construtor ]
		public RastreioPedidoRecebidoCliente()
		{
			dadosRaw = new RastreioPedidoRecebidoClienteRaw();
			dadosNormalizado = new RastreioPedidoRecebidoClienteNormalizado();
			processo = new RastreioPedidoRecebidoClienteProcesso();
		}
		#endregion
	}
	#endregion

	#region [ RastreioPedidoRecebidoClienteProcesso ]
	public enum eRastreioPedidoRecebidoClienteProcessoStatus
	{
		StatusInicial = 0,
		ErroInconsistencia = 1,
		LiberadoParaRegistrarPedidoRecebidoCliente = 2,
		SucessoRegistroPedidoRecebidoCliente = 3,
		FalhaRegistroPedidoRecebidoCliente = 4,
		LiberadoParaRegistrarPedidoRecebidoCliente_E_LiberadoParaRegistrarDataPrevisaoEntrega = 5,
		LiberadoParaRegistrarDataPrevisaoEntrega = 6,
		FalhaRegistroDataPrevisaoEntrega = 7,
		FalhaRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega = 8,
		SucessoRegistroDataPrevisaoEntrega = 9,
		SucessoRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega = 10,
		SucessoRegistroPedidoRecebidoCliente_E_FalhaRegistroDataPrevisaoEntrega = 11,
		FalhaRegistroPedidoRecebidoCliente_E_SucessoRegistroDataPrevisaoEntrega = 12
	}

	public enum eRastreioPedidoRecebidoClienteProcessoCodigoErro
	{
		// A numeração dos códigos influencia na ordenação em que os dados são exibidos (números menores são mais prioritários, com exceção do zero)
		SEM_ERRO = 0,
		OCORRENCIA_SEM_DATA_RECEBIMENTO = 1,
		MULTIPLOS_PEDIDOS_LOCALIZADOS_PARA_NF = 2,
		CNPJ_CPF_DIVERGENTE = 3,
		OCORRENCIA_DEVOLUCAO = 4,
		PEDIDO_NAO_LOCALIZADO_POR_NF = 5,
		PEDIDO_ST_ENTREGA_INVALIDO = 6,
		PEDIDO_SEM_TRANSPORTADORA_CADASTRADA = 7,
		DATA_RECEBIMENTO_ANTERIOR_DATA_PEDIDO_ENTREGUE = 8,
		NUMERO_NF_NAO_INFORMADO = 9,
		NUMERO_NF_NAO_ENCONTRADO = 10,
		NUMERO_NF_FORMATO_INVALIDO = 11,
		OCORRENCIA_REPETIDA = 12,
		OCORRENCIA_COM_SITUACAO_INVALIDA = 13,
		PEDIDO_RECEBIDO_JA_REGISTRADO = 14,
		OCORRENCIA_SEM_DATA_RECEBIMENTO_E_SEM_DATA_PREVISAO_ENTREGA = 15
	}

	class RastreioPedidoRecebidoClienteProcesso
	{
		public string Guid { get; set; } = "";
		public string Pedido { get; set; } = "";
		public string marketplace_codigo_origem { get; set; } = "";
		public byte MarketplacePedidoRecebidoRegistrarStatus { get; set; } = 0;
		public DateTime PrevisaoEntregaTranspDataOriginal { get; set; } = DateTime.MinValue;
		public string MensagemInformativa { get; set; } = "";
		public string MensagemErro { get; set; } = "";
		public eRastreioPedidoRecebidoClienteProcessoStatus Status { get; set; } = eRastreioPedidoRecebidoClienteProcessoStatus.StatusInicial;
		public eRastreioPedidoRecebidoClienteProcessoCodigoErro CodigoErro { get; set; } = eRastreioPedidoRecebidoClienteProcessoCodigoErro.SEM_ERRO;
		public string campoOrdenacao { get; set; } = string.Empty;
	}
	#endregion

	#region [ RastreioPedidoRecebidoClienteNormalizado ]
	class RastreioPedidoRecebidoClienteNormalizado
	{
		public string CnpjCpfRemetente { get; set; } = "";
		public string Remetente { get; set; } = "";
		public string CnpjCpfDestinatario { get; set; } = "";
		public string Destinatario { get; set; } = "";
		public string CTRC { get; set; } = "";
		public string SerieNF { get; set; } = "";
		public int numSerieNF { get; set; } = 0;
		public string NF { get; set; } = "";
		public int numNF { get; set; } = 0;
		public string NroPedido { get; set; } = "";
		public string DataInclusao { get; set; } = "";
		public DateTime dtDataInclusao { get; set; } = DateTime.MinValue;
		public string CidadeDestino { get; set; } = "";
		public string UfDestino { get; set; } = "";
		public string Unidade { get; set; } = "";
		public string DataHoraOcorrencia { get; set; } = "";
		public DateTime dtDataHoraOcorrencia { get; set; } = DateTime.MinValue;
		public string Situacao { get; set; } = "";
		public string Empresa { get; set; } = "";
		public string Detalhe { get; set; } = "";
		public string DataEntrega { get; set; } = "";
		public DateTime dtDataEntrega { get; set; } = DateTime.MinValue;
		public string PrevisaoEntrega { get; set; } = "";
		public DateTime dtPrevisaoEntrega { get; set; } = DateTime.MinValue;

		public bool isSituacaoMercadoriaEntregue
		{
			get { return Global.listaCodigosRastreioSituacaoMercadoriaEntregue.Contains(Situacao.Trim().ToUpper()); }
		}

		public bool hasDataInclusao
		{
			get { return (DataInclusao.Trim().Length > 0); }
		}

		public bool hasDataHoraOcorrencia
		{
			get { return (DataHoraOcorrencia.Trim().Length > 0); }
		}

		public bool hasDataEntrega
		{
			get { return (DataEntrega.Trim().Length > 0); }
		}

		public bool hasPrevisaoEntrega
		{
			get { return (PrevisaoEntrega.Trim().Length > 0); }
		}

		public bool hasNF
		{
			get { return (NF.Trim().Length > 0); }
		}
	}
	#endregion

	#region [ RastreioPedidoRecebidoClienteRaw ]
	class RastreioPedidoRecebidoClienteRaw
	{
		public string CnpjCpfRemetente { get; set; } = "";
		public string Remetente { get; set; } = "";
		public string CnpjCpfDestinatario { get; set; } = "";
		public string Destinatario { get; set; } = "";
		public string CTRC { get; set; } = "";
		public string NF { get; set; } = "";
		public string NroPedido { get; set; } = "";
		public string DataInclusao { get; set; } = "";
		public string CidadeDestino { get; set; } = "";
		public string UfDestino { get; set; } = "";
		public string Unidade { get; set; } = "";
		public string DataHoraOcorrencia { get; set; } = "";
		public string Situacao { get; set; } = "";
		public string Empresa { get; set; } = "";
		public string Detalhe { get; set; } = "";
		public string DataEntrega { get; set; } = "";
		public string PrevisaoEntrega { get; set; } = "";

		public bool isSituacaoMercadoriaEntregue
		{
			get { return Global.listaCodigosRastreioSituacaoMercadoriaEntregue.Contains(Situacao.Trim().ToUpper()); }
		}

		public bool hasDataInclusao
		{
			get { return (DataInclusao.Trim().Length > 0); }
		}

		public bool hasDataHoraOcorrencia
		{
			get { return (DataHoraOcorrencia.Trim().Length > 0); }
		}

		public bool hasDataEntrega
		{
			get { return (DataEntrega.Trim().Length > 0); }
		}

		public bool hasPrevisaoEntrega
		{
			get { return (PrevisaoEntrega.Trim().Length > 0); }
		}

		public bool hasNF
		{
			get { return (NF.Trim().Length > 0); }
		}
	}
	#endregion

	#region [ ColunaHeaderRastreioPedidoRecebidoCliente ]
	class ColunaHeaderRastreioPedidoRecebidoCliente
	{
		public string tituloColuna { get; set; } = "";
		public int? indexColuna { get; set; } = null;

		#region [ Constutor ]
		public ColunaHeaderRastreioPedidoRecebidoCliente(string TituloColuna)
		{
			tituloColuna = TituloColuna;
		}
		#endregion
	}
	#endregion

	#region [ HeaderRastreioPedidoRecebidoCliente ]
	class HeaderRastreioPedidoRecebidoCliente
	{
		public List<ColunaHeaderRastreioPedidoRecebidoCliente> listaCamposHeader;

		public ColunaHeaderRastreioPedidoRecebidoCliente CnpjCpfRemetente { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente Remetente { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente CnpjCpfDestinatario { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente Destinatario { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente CTRC { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente NF { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente NroPedido { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente DataInclusao { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente CidadeDestino { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente UfDestino { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente Unidade { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente DataHoraOcorrencia { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente Situacao { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente Empresa { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente Detalhe { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente DataEntrega { get; set; }
		public ColunaHeaderRastreioPedidoRecebidoCliente PrevisaoEntrega { get; set; }

		#region [ Construtor ]
		public HeaderRastreioPedidoRecebidoCliente()
		{
			listaCamposHeader = new List<ColunaHeaderRastreioPedidoRecebidoCliente>();

			CnpjCpfRemetente = new ColunaHeaderRastreioPedidoRecebidoCliente("CNPJ/CPF Remetente");
			listaCamposHeader.Add(CnpjCpfRemetente);

			Remetente = new ColunaHeaderRastreioPedidoRecebidoCliente("Remetente");
			listaCamposHeader.Add(Remetente);

			CnpjCpfDestinatario = new ColunaHeaderRastreioPedidoRecebidoCliente("CNPJ/CPF Destinatario");
			listaCamposHeader.Add(CnpjCpfDestinatario);

			Destinatario = new ColunaHeaderRastreioPedidoRecebidoCliente("Destinatario");
			listaCamposHeader.Add(Destinatario);

			CTRC = new ColunaHeaderRastreioPedidoRecebidoCliente("CTRC");
			listaCamposHeader.Add(CTRC);

			NF = new ColunaHeaderRastreioPedidoRecebidoCliente("Nota Fiscal/Nro Coleta");
			listaCamposHeader.Add(NF);

			NroPedido = new ColunaHeaderRastreioPedidoRecebidoCliente("Nro Pedido");
			listaCamposHeader.Add(NroPedido);

			DataInclusao = new ColunaHeaderRastreioPedidoRecebidoCliente("Data Inclusao");
			listaCamposHeader.Add(DataInclusao);

			CidadeDestino = new ColunaHeaderRastreioPedidoRecebidoCliente("Cidade Destino");
			listaCamposHeader.Add(CidadeDestino);

			UfDestino = new ColunaHeaderRastreioPedidoRecebidoCliente("UF Destino");
			listaCamposHeader.Add(UfDestino);

			Unidade = new ColunaHeaderRastreioPedidoRecebidoCliente("Unidade");
			listaCamposHeader.Add(Unidade);

			DataHoraOcorrencia = new ColunaHeaderRastreioPedidoRecebidoCliente("Data/Hora da Ocorrencia");
			listaCamposHeader.Add(DataHoraOcorrencia);

			Situacao = new ColunaHeaderRastreioPedidoRecebidoCliente("Situacao");
			listaCamposHeader.Add(Situacao);

			Empresa = new ColunaHeaderRastreioPedidoRecebidoCliente("Empresa");
			listaCamposHeader.Add(Empresa);

			Detalhe = new ColunaHeaderRastreioPedidoRecebidoCliente("Detalhe");
			listaCamposHeader.Add(Detalhe);

			DataEntrega = new ColunaHeaderRastreioPedidoRecebidoCliente("Data Entrega");
			listaCamposHeader.Add(DataEntrega);

			PrevisaoEntrega = new ColunaHeaderRastreioPedidoRecebidoCliente("Previsao de Entrega");
			listaCamposHeader.Add(PrevisaoEntrega);
		}
		#endregion

	}
	#endregion
}
