#region [ using ]
using System;
using System.Collections.Generic;
#endregion

namespace Reciprocidade
{
	class ArqRemessaRetorno
	{
		#region [ Atributos ]
		private List<TabelaErro> _registrosErro = new List<TabelaErro>();
		public List<TabelaErro> registrosErro
		{
			get { return _registrosErro; }
			set { _registrosErro = value; }
		}

		private List<RegistroDetalhe> _registrosDetalhe = new List<RegistroDetalhe>();
		public List<RegistroDetalhe> registrosDetalhe
		{
			get { return _registrosDetalhe; }
			set { _registrosDetalhe = value; }
		}

		private Dictionary<String, String> _dicionarioErros = new Dictionary<String, String>();
		public Dictionary<String, String> dicionarioErros
		{
			get { return _dicionarioErros; }
			set { _dicionarioErros = value; }
		}
		#endregion

		#region [ Métodos ]
		public void adicionaRegistroErro(TabelaErro registro)
		{
			_registrosErro.Add(registro);
		}

		public void adicionaRegistroDetalhe(RegistroDetalhe registro)
		{
			_registrosDetalhe.Add(registro);
		}

		public String getMsgErro(String codErro)
		{
			if (_dicionarioErros.Count == 0)
			{
				converteTabErrosParaDictionary();
			}

			String ret = _dicionarioErros[codErro];

			return ret;
		}

		private void converteTabErrosParaDictionary()
		{
			foreach (TabelaErro t in _registrosErro)
			{
				_dicionarioErros.Add(t.numeroMsg, t.descricao);
			}
		}

		public int getTotalRegistros()
		{
			return _registrosDetalhe.Count;
		}

		public int getTotalRegistrosRejeitados()
		{
			int contador = 0;
			foreach (RegistroDetalhe registro in _registrosDetalhe)
			{
				String erros = registro.erros;
				if (erros.Trim().Length > 0)
				{
					contador++;
				}
			}
			return contador;
		}

		public int getTotalRegistrosSemRejeicao()
		{
			return getTotalRegistros() - getTotalRegistrosRejeitados();
		}
		#endregion

		#region [ Mensagem de Processamento da Remessa ]
		public class MensagemProcessamento
		{
			private String _numero;
			public String numero
			{
				get { return _numero; }
				set { _numero = value; }
			}

			private String _mensagem;
			public String mensagem
			{
				get { return _mensagem; }
				set { _mensagem = value; }
			}

			public MensagemProcessamento() { }

			public MensagemProcessamento(String numero, String mensagem)
			{
				this._numero = numero;
				this._mensagem = mensagem;
			}

			public void carrega(String linha)
			{
				String num = linha.Substring(2, 2);
				String msg = linha.Substring(4);

				this._numero = num;
				this._mensagem = msg;
			}
		}
		#endregion

		#region [ Texto de Relatório de Totalização da Remessa ]
		public class RelatorioTotalizacao
		{
			private String _descricao;
			public String descricao
			{
				get { return _descricao; }
				set { _descricao = value; }
			}

			public RelatorioTotalizacao() { }

			public RelatorioTotalizacao(String descricao)
			{
				this._descricao = descricao;
			}

			public void carrega(String linha)
			{
				String desc;
				if (linha.Length >= 75)
				{
					desc = linha.Substring(3, 72);
				}
				else
				{
					desc = linha.Substring(3);
				}

				this._descricao = desc;
			}
		}
		#endregion

		#region [ Tabela de Erros ]
		public class TabelaErro
		{
			private String _numeroMsg;
			public String numeroMsg
			{
				get { return _numeroMsg; }
				set { _numeroMsg = value; }
			}

			private String _descricao;
			public String descricao
			{
				get { return _descricao; }
				set { _descricao = value; }
			}

			public TabelaErro() { }

			public TabelaErro(String numeroMsg, String descricao)
			{
				this._numeroMsg = numeroMsg;
				this._descricao = descricao;
			}

			public void carrega(String linha)
			{
				String num = linha.Substring(2, 3);
				String desc;

				if (linha.Length >= 77)
				{
					desc = linha.Substring(7, 70);
				}
				else
				{
					desc = linha.Substring(7);
				}

				this._numeroMsg = num;
				this._descricao = desc;
			}
		}
		#endregion

		#region [ Registro Detalhe ]
		public class RegistroDetalhe
		{
			private ArqRemessa.DetalheTitulo _linhaDetalhe;
			public ArqRemessa.DetalheTitulo linhaDetalhe
			{
				get { return _linhaDetalhe; }
				set { _linhaDetalhe = value; }
			}

			private String _erros;
			public String erros
			{
				get { return _erros; }
				set { _erros = value; }
			}

			public RegistroDetalhe(ArqRemessa.DetalheTitulo linhaDetalhe, String erros)
			{
				this._linhaDetalhe = linhaDetalhe;
				this._erros = erros;
			}
		}
		#endregion

		#region [ Totalizador de Clientes ]
		public class TotalizadorCliente
		{
			private int _qtde;
			public int qtde
			{
				get { return _qtde; }
				set { _qtde = value; }
			}

			public TotalizadorCliente() { }

			public TotalizadorCliente(int qtde)
			{
				this._qtde = qtde;
			}

			public void carrega(String linha)
			{
				String qtde = linha.Substring(2, 6);
				this._qtde = Convert.ToInt32(qtde);
			}
		}
		#endregion

		#region [ Totalizador Pagamento ]
		public class TotalizadorPagamento
		{
			private int _qtde;
			public int qtde
			{
				get { return _qtde; }
				set { _qtde = value; }
			}

			private decimal _somatoria;
			public decimal somatoria
			{
				get { return _somatoria; }
				set { _somatoria = value; }
			}

			public TotalizadorPagamento() { }

			public TotalizadorPagamento(int qtde, decimal somatoria)
			{
				this._qtde = qtde;
				this._somatoria = somatoria;
			}

			public void carrega(String linha)
			{
				String qtde = linha.Substring(2, 8);
				String soma = linha.Substring(10, 18);

				this._qtde = Convert.ToInt32(qtde);
				this._somatoria = Convert.ToInt32(soma);
			}
		}
		#endregion
	}
}
