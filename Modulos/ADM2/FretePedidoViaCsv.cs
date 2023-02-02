using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
	#region [ FretePedidoViaCsv ]
	class FretePedidoViaCsv
	{
		public FretePedidoViaCsvRaw dadosRaw { get; set; }
		public FretePedidoViaCsvNormalizado dadosNormalizado { get; set; }
		public FretePedidoViaCsvProcesso processo { get; set; }

		#region [ Construtor ]
		public FretePedidoViaCsv()
		{
			dadosRaw = new FretePedidoViaCsvRaw();
			dadosNormalizado = new FretePedidoViaCsvNormalizado();
			processo = new FretePedidoViaCsvProcesso();
		}
		#endregion
	}
	#endregion

	#region [ FretePedidoViaCsvRaw ]
	class FretePedidoViaCsvRaw
	{
		public int linhaArquivoCsv { get; set; } = 0;
		public string CnpjRemetente { get; set; } = "";
		public string NF { get; set; } = "";
		public string ValorFrete { get; set; } = "";
		public string TransportadoraCsv { get; set; } = "";
		public string TipoFrete { get; set; } = "";

		public bool hasCnpjRemetente
		{
			get { return (CnpjRemetente.Trim().Length > 0); }
		}

		public bool hasNF
		{
			get { return (NF.Trim().Length > 0); }
		}

		public bool hasValorFrete
		{
			get { return (ValorFrete.Trim().Length > 0); }
		}

		public bool hasTransportadoraCsv
		{
			get { return (TransportadoraCsv.Trim().Length > 0); }
		}

		public bool hasTipoFrete
		{
			get { return (TipoFrete.Trim().Length > 0); }
		}
	}
	#endregion

	#region [ FretePedidoViaCsvNormalizado ]
	class FretePedidoViaCsvNormalizado
	{
		public string CnpjRemetente { get; set; } = "";
		public string NF { get; set; } = "";
		public int serieNF { get; set; } = 0;
		public int numNF { get; set; } = 0;
		public string ValorFrete { get; set; } = "";
		public decimal vlFrete { get; set; } = 0m;
		public string TransportadoraCsv { get; set; } = "";
		public string TransportadoraCnpjCsv { get; set; } = "";
		public string TipoFrete { get; set; } = "";
		public string TipoFreteCodigoSistema { get; set; } = "";
		public string TipoFreteDescricaoSistema { get; set; } = "";
		public int TipoFreteOrdenacaoSistema { get; set; } = 0;
		public bool hasCnpjRemetente
		{
			get { return (CnpjRemetente.Trim().Length > 0); }
		}

		public bool hasNF
		{
			get { return (NF.Trim().Length > 0); }
		}

		public bool hasValorFrete
		{
			get { return (ValorFrete.Trim().Length > 0); }
		}

		public bool hasTransportadoraCsv
		{
			get { return (TransportadoraCsv.Trim().Length > 0); }
		}

		public bool hasTipoFrete
		{
			get { return (TipoFrete.Trim().Length > 0); }
		}
	}
	#endregion

	#region [ FretePedidoViaCsvProcesso ]
	public enum eFretePedidoViaCsvProcessoStatus
	{
		StatusInicial = 0,
		ErroInconsistencia = 1,
		LiberadoComRessalvasParaRegistrarFretePedido = 2,
		LiberadoComObsParaRegistrarFretePedido = 3,
		LiberadoParaRegistrarFretePedido = 4,
		SucessoRegistroFretePedido = 5,
		FalhaRegistroFretePedido = 6
	}

	public enum eFretePedidoViaCsvProcessoCodigoErro
	{
		// A numeração dos códigos influencia na ordenação em que os dados são exibidos (números menores são mais prioritários, com exceção do zero)
		SEM_ERRO = 0,
		CNPJ_REMETENTE_DESCONHECIDO = 1,
		NUMERO_NF_NAO_INFORMADO = 2,
		NUMERO_NF_NAO_ENCONTRADO = 3,
		NUMERO_NF_FORMATO_INVALIDO = 4,
		TIPO_FRETE_DESCONHECIDO = 5,
		VALOR_FRETE_INVALIDO = 6,
		PEDIDO_NAO_LOCALIZADO_POR_NF = 7,
		MULTIPLOS_PEDIDOS_LOCALIZADOS_PARA_NF = 8,
		TRANSPORTADORA_DESCONHECIDA = 9,
		PEDIDO_SEM_TRANSPORTADORA_CADASTRADA = 10,
		PEDIDO_ST_ENTREGA_INVALIDO = 11,
		FRETE_JA_REGISTRADO = 12
	}

	class FretePedidoViaCsvProcesso
	{
		public string Guid { get; set; } = "";
		public Pedido pedido { get; set; } = null;
		public string MensagemInformativa { get; set; } = "";
		public string MensagemErro { get; set; } = "";
		public eFretePedidoViaCsvProcessoStatus Status { get; set; } = eFretePedidoViaCsvProcessoStatus.StatusInicial;
		public eFretePedidoViaCsvProcessoCodigoErro CodigoErro { get; set; } = eFretePedidoViaCsvProcessoCodigoErro.SEM_ERRO;
		public string campoOrdenacao { get; set; } = string.Empty;
	}
	#endregion

	#region [ ColunaHeaderFretePedidoViaCsv ]
	class ColunaHeaderFretePedidoViaCsv
	{
		public string tituloColuna { get; set; } = "";
		public int? indexColuna { get; set; } = null;

		#region [ Construtor ]
		public ColunaHeaderFretePedidoViaCsv(string TituloColuna)
		{
			tituloColuna = TituloColuna;
		}
		#endregion
	}
	#endregion

	#region [ HeaderFretePedidoViaCsv ]
	class HeaderFretePedidoViaCsv
	{
		public List<ColunaHeaderFretePedidoViaCsv> listaCamposHeader;

		public ColunaHeaderFretePedidoViaCsv CnpjRemetente { get; set; }
		public ColunaHeaderFretePedidoViaCsv NF { get; set; }
		public ColunaHeaderFretePedidoViaCsv ValorFrete { get; set; }
		public ColunaHeaderFretePedidoViaCsv TransportadoraCsv { get; set; }
		public ColunaHeaderFretePedidoViaCsv TipoFrete { get; set; }

		#region [ Construtor ]
		public HeaderFretePedidoViaCsv()
		{
			listaCamposHeader = new List<ColunaHeaderFretePedidoViaCsv>();

			CnpjRemetente = new ColunaHeaderFretePedidoViaCsv("CNPJ REMETENTE");
			listaCamposHeader.Add(CnpjRemetente);

			NF = new ColunaHeaderFretePedidoViaCsv("NF");
			listaCamposHeader.Add(NF);

			ValorFrete = new ColunaHeaderFretePedidoViaCsv("FRETE");
			listaCamposHeader.Add(ValorFrete);

			TransportadoraCsv = new ColunaHeaderFretePedidoViaCsv("TRANSPORTADORA");
			listaCamposHeader.Add(TransportadoraCsv);

			TipoFrete = new ColunaHeaderFretePedidoViaCsv("TIPO DE FRETE");
			listaCamposHeader.Add(TipoFrete);
		}
		#endregion
	}
	#endregion

	#region [ CodigoDescricaoTipoFrete ]
	class CodigoDescricaoTipoFrete
	{
		public string grupo { get; set; } = "";
		public string codigo { get; set; } = "";
		public string descricao { get; set; } = "";
		public int ordenacao { get; set; } = 0;
		public List<string> listaCodigosAceitosCsv { get; set; } = new List<string>();
	}
	#endregion
}
