#region [ using ]
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;
#endregion

namespace Reciprocidade
{
	class ArqRemessa
	{
		#region [ Atributos ]
		private LinhaHeader _linhaHeader = null;
		public LinhaHeader linhaHeader
		{
			get { return _linhaHeader; }
			set { _linhaHeader = value; }
		}

		private List<DetalheTempoRelacionamento> _detTempoRelactoList = new List<DetalheTempoRelacionamento>();
		public List<DetalheTempoRelacionamento> detTempoRelactoList
		{
			get { return _detTempoRelactoList; }
			set { _detTempoRelactoList = value; }
		}

		private List<DetalheTitulo> _detTituloList = new List<DetalheTitulo>();
		public List<DetalheTitulo> detTituloList
		{
			get { return _detTituloList; }
			set { _detTituloList = value; }
		}

		private LinhaTrailler _linhaTrailler = null;
		public LinhaTrailler linhaTrailler
		{
			get { return _linhaTrailler; }
			set { _linhaTrailler = value; }
		}
		#endregion

		#region [ Métodos das listas que contém os registros ]
		public void addDetalheTempoRelacionamento(DetalheTempoRelacionamento detTempoRelacto)
		{
			_detTempoRelactoList.Add(detTempoRelacto);
		}

		public void addDetalheTitulo(DetalheTitulo detTitulo)
		{
			_detTituloList.Add(detTitulo);
		}
		#endregion

		#region [ LinhaHeader ]
		public class LinhaHeader
		{
			const String ID_HEADER = "00";
			const String RELATO_COMP_NEGOCIOS = "RELATO COMP NEGOCIOS";
			public static readonly string CNPJ_EMPRESA_CONVENIADA = Global.Cte.SerasaReciprocidade.CNPJ_EMPRESA_CONVENIADA;
			public const String PERIODICIDADE_REMESSA = Global.Cte.SerasaReciprocidade.CODIGO_PERIODICIDADE_ARQ_REMESSA;
			public static String RESERVADO_SERASA = new String(' ', 15);
			public static String BRANCOS_1 = new String(' ', 29);
			public const String ID_VERSAO_LAYOUT = "V.";
			public const String NUM_VERSAO_LAYOUT = "01";
			public static String BRANCOS_2 = new String(' ', 26);

			private DateTime _dataInicio;
			public DateTime dataInicio
			{
				get { return _dataInicio; }
				set { _dataInicio = value; }
			}

			private DateTime _dataFim;
			public DateTime dataFim
			{
				get { return _dataFim; }
				set { _dataFim = value; }
			}

			public static String ID_GRUPO_RELATO_SEGMENTO = new String(' ', 3);

			private const String CONCILIA = "CONCILIA";
			private bool _isConciliacao = false;

			public LinhaHeader(DateTime dataInicio, DateTime dataFim)
			{
				this._dataInicio = dataInicio;
				this._dataFim = dataFim;
			}

			public LinhaHeader(DateTime dataFim)
			{
				this._dataFim = dataFim;
				this._isConciliacao = true;
			}

			public override string ToString()
			{
				String dataInicioFormatada = "";

				if (_isConciliacao)
				{
					dataInicioFormatada = CONCILIA;
				}
				else
				{
					dataInicioFormatada = Global.formataDataYyyyMmDdSemSeparador(this._dataInicio);
				}

				String dataFimFormatada = Global.formataDataYyyyMmDdSemSeparador(this._dataFim);

				return ID_HEADER + RELATO_COMP_NEGOCIOS + CNPJ_EMPRESA_CONVENIADA + dataInicioFormatada + dataFimFormatada
					+ PERIODICIDADE_REMESSA + RESERVADO_SERASA + ID_GRUPO_RELATO_SEGMENTO + BRANCOS_1 + ID_VERSAO_LAYOUT
					+ NUM_VERSAO_LAYOUT + BRANCOS_2;
			}
		}
		#endregion

		#region [ DetalheTempoRelacionamento ]
		public class DetalheTempoRelacionamento
		{
			public const String ID = "01";
			public const String TIPO_DADOS = "01";
			public static String BRANCOS = new String(' ', 103);
			const String CLIENTE_ANTIGO = "1";
			const String CLIENTE_MENOS_DE_UM_ANO = "2";
			const int ANO = 365;

			private int _clienteId;
			public int clienteId
			{
				get { return _clienteId; }
				set { _clienteId = value; }
			}

			private String _cnpjCliente = "";
			public String cnpjCliente
			{
				get { return _cnpjCliente; }
				set { _cnpjCliente = value; }
			}

			private DateTime _clienteDesde;
			public DateTime clienteDesde
			{
				get { return _clienteDesde; }
				set { _clienteDesde = value; }
			}

			private String _tipoCliente;
			public String tipoCliente
			{
				get { return _tipoCliente; }
				set { _tipoCliente = value; }
			}

			public DetalheTempoRelacionamento(int clienteId, String cnpjCliente, DateTime clienteDesde)
			{
				this._clienteId = clienteId;
				this._cnpjCliente = cnpjCliente;
				this._clienteDesde = clienteDesde;

				long tempo = DateAndTime.DateDiff(DateInterval.Day, _clienteDesde, DateTime.Now);
				if (tempo < ANO)
				{
					_tipoCliente = CLIENTE_MENOS_DE_UM_ANO;
				}
				else
				{
					_tipoCliente = CLIENTE_ANTIGO;
				}
			}

			public override string ToString()
			{

				String clienteDesdeFormatado = Global.formataDataYyyyMmDdSemSeparador(_clienteDesde);
				return ID + _cnpjCliente + TIPO_DADOS + clienteDesdeFormatado + _tipoCliente + BRANCOS;
			}
		}
		#endregion

		#region [ DetalheTitulo ]
		public class DetalheTitulo
		{
			public const String ID = "01";
			public const String TIPO_DADOS = "05";
			public static String BRANCOS = new String(' ', 1);
			public static String RESERVADO_SERASA = new String(' ', 30);

			#region [ Atributos ]
			private int _tituloMovimentoId;
			public int tituloMovimentoId
			{
				get { return _tituloMovimentoId; }
				set { _tituloMovimentoId = value; }
			}

			private int _clienteId;
			public int clienteId
			{
				get { return _clienteId; }
				set { _clienteId = value; }
			}

			private String _cnpjSacado;
			public String cnpjSacado
			{
				get { return _cnpjSacado; }
				set { _cnpjSacado = value; }
			}

			private String _numeroTitulo;
			public String numeroTitulo
			{
				get { return _numeroTitulo; }
				set { _numeroTitulo = value; }
			}

			private DateTime _dataEmissao;
			public DateTime dataEmissao
			{
				get { return _dataEmissao; }
				set { _dataEmissao = value; }
			}

			private Decimal _valorTitulo;
			public Decimal valorTitulo
			{
				get { return _valorTitulo; }
				set { _valorTitulo = value; }
			}

			private DateTime _dataVencimento;
			public DateTime dataVencimento
			{
				get { return _dataVencimento; }
				set { _dataVencimento = value; }
			}

			private DateTime _dataPagamento;
			public DateTime dataPagamento
			{
				get { return _dataPagamento; }
				set { _dataPagamento = value; }
			}

			private String _numeroTitulo10Posicoes;
			public String numeroTitulo10Posicoes
			{
				get { return _numeroTitulo10Posicoes; }
				set { _numeroTitulo10Posicoes = value; }
			}

			private String _numeroTituloEstendido;
			public String numeroTituloEstendido
			{
				get { return _numeroTituloEstendido; }
				set { _numeroTituloEstendido = value; }
			}

			private bool _isTituloBaixado;
			public bool isTituloBaixado
			{
				get { return _isTituloBaixado; }
				set { _isTituloBaixado = value; }
			}

			private bool _isTituloExcluido;
			public bool isTituloExcluido
			{
				get { return _isTituloExcluido; }
				set { _isTituloExcluido = value; }
			}
			#endregion

			public DetalheTitulo() { }

			public DetalheTitulo(int tituloMovimentoId, int clienteId, String cnpjSacado, String numeroTitulo, DateTime dataEmissao, Decimal valorTitulo, DateTime dataVencimento, DateTime dataPagamento)
			{
				this._tituloMovimentoId = tituloMovimentoId;
				this._clienteId = clienteId;
				this._cnpjSacado = cnpjSacado;
				this._numeroTitulo = numeroTitulo;
				this._dataEmissao = dataEmissao;
				this._valorTitulo = valorTitulo;
				this._dataVencimento = dataVencimento;
				this._dataPagamento = dataPagamento;
			}

			private void trataNumeroTitulos()
			{
				String nTitulo = "#D" + _numeroTitulo;
				StringBuilder aux = new StringBuilder();

				_numeroTitulo10Posicoes = _numeroTitulo.Substring(0, 10);

				if (_numeroTitulo.Length > 10)
				{
					aux.Append(nTitulo);
					aux.Insert(nTitulo.Length, " ", 34 - nTitulo.Length);
					_numeroTituloEstendido = aux.ToString();
				}
				else
				{
					aux.Append(" ", 0, 34);
					_numeroTituloEstendido = aux.ToString();
				}
			}

			public override string ToString()
			{
				String dtEmissao = Global.formataDataYyyyMmDdSemSeparador(this._dataEmissao);
				String vlTitulo = Global.formataMoedaSemSeparador(this._valorTitulo, 13);

				if (_isTituloBaixado || _isTituloExcluido)
				{
					vlTitulo = "9999999999999"; //formato para a exclusão do título junto a Serasa
				}

				String dtVcto = Global.formataDataYyyyMmDdSemSeparador(this._dataVencimento);

				trataNumeroTitulos();

				String dtPgto;
				if (this._dataPagamento != DateTime.MinValue)
				{
					dtPgto = Global.formataDataYyyyMmDdSemSeparador(this._dataPagamento);
				}
				else
				{
					dtPgto = new String(' ', 8);
				}

				return ID + _cnpjSacado + TIPO_DADOS + _numeroTitulo10Posicoes + dtEmissao + vlTitulo + dtVcto + dtPgto +
					_numeroTituloEstendido + BRANCOS + RESERVADO_SERASA;
			}

			public void carrega(String linha)
			{
				this._cnpjSacado = linha.Substring(2, 14);
				this._numeroTitulo10Posicoes = linha.Substring(18, 10);

				String indicaTituloEstendido = linha.Substring(65, 2);
				if (indicaTituloEstendido.Equals("#D"))
				{
					this._numeroTituloEstendido = linha.Substring(67, 32);
					this._numeroTitulo = _numeroTituloEstendido;
				}
				else
				{
					this._numeroTitulo = _numeroTitulo10Posicoes;
				}

				this._dataEmissao = Global.converteYyyyMmDdParaDateTime(linha.Substring(28, 8));
				this._valorTitulo = Global.converteMoedaSemSeparadorParaDecimal(linha.Substring(36, 13));
				this._dataVencimento = Global.converteYyyyMmDdParaDateTime(linha.Substring(49, 8));

				String dtPgto = linha.Substring(57, 8);
				if ((dtPgto.Trim().Length > 0) && (!dtPgto.Equals("99999999")))
				{
					this._dataPagamento = Global.converteYyyyMmDdParaDateTime(dtPgto);
				}
			}
		}
		#endregion

		#region [ LinhaTrailler ]
		public class LinhaTrailler
		{
			const String ID_TRAILLER = "99";
			public static String BRANCOS_1 = new String(' ', 44);
			public static String RESERVADO_SERASA = new String(' ', 32);
			public static String BRANCOS_2 = new String(' ', 30);

			private int _qtdeRegTempoRelacionamento = 0;
			public int qtdeRegTempoRelacionamento
			{
				get { return _qtdeRegTempoRelacionamento; }
				set { _qtdeRegTempoRelacionamento = value; }
			}

			private int _qtdeRegTitulo = 0;
			public int qtdeRegTitulo
			{
				get { return _qtdeRegTitulo; }
				set { _qtdeRegTitulo = value; }
			}

			public LinhaTrailler(int qtdeRegTempoRelacionamento, int qtdeRegTitulo)
			{
				this._qtdeRegTempoRelacionamento = qtdeRegTempoRelacionamento;
				this._qtdeRegTitulo = qtdeRegTitulo;
			}

			//Layout exige que tais campos tenham 11 posicoes
			private String formataQtdeRegistros(int qtdeRegistro)
			{
				string strQtdeRegistro = qtdeRegistro.ToString();
				int qtdeDigitos = strQtdeRegistro.Length;
				StringBuilder registroFormatado = new StringBuilder(strQtdeRegistro);
				registroFormatado.Insert(0, "0", 11 - qtdeDigitos);

				return registroFormatado.ToString();
			}

			public override string ToString()
			{
				String qtdRegTempoRelac = formataQtdeRegistros(this._qtdeRegTempoRelacionamento);
				String qtdRegTitulo = formataQtdeRegistros(this._qtdeRegTitulo);

				return ID_TRAILLER + qtdRegTempoRelac + BRANCOS_1 + qtdRegTitulo + RESERVADO_SERASA + BRANCOS_2;
			}
		}
		#endregion
	}
}
