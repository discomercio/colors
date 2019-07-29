#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace ADM2
{
	#region [ Classe: IbptDados ]
	class IbptDados
	{
		#region [ Getters/Setters ]
		private String _codigo;
		public String codigo
		{
			get { return _codigo; }
			set { _codigo = value; }
		}

		private String _ex;
		public String ex
		{
			get { return _ex; }
			set { _ex = value; }
		}

		private String _tabela;
		public String tabela
		{
			get { return _tabela; }
			set { _tabela = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private String _aliqNac;
		public String aliqNac
		{
			get { return _aliqNac; }
			set { _aliqNac = value; }
		}

		private String _aliqImp;
		public String aliqImp
		{
			get { return _aliqImp; }
			set { _aliqImp = value; }
		}

		private double _percAliqNac;
		public double percAliqNac
		{
			get { return _percAliqNac; }
			set { _percAliqNac = value; }
		}

		private double _percAliqImp;
		public double percAliqImp
		{
			get { return _percAliqImp; }
			set { _percAliqImp = value; }
		}
		#endregion
	}
	#endregion

	#region [ Classe: LinhaHeaderArquivoIbptCsv ]
	public class LinhaHeaderArquivoIbptCsv
	{
		#region [ Getters / Setters ]
		private String _codigo;
		public String codigo
		{
			get { return _codigo; }
			set { _codigo = value; }
		}

		private String _ex;
		public String ex
		{
			get { return _ex; }
			set { _ex = value; }
		}

		private String _tabela;
		public String tabela
		{
			get { return _tabela; }
			set { _tabela = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private String _aliqNac;
		public String aliqNac
		{
			get { return _aliqNac; }
			set { _aliqNac = value; }
		}

		private String _aliqImp;
		public String aliqImp
		{
			get { return _aliqImp; }
			set { _aliqImp = value; }
		}

		private String _versao;
		public String versao
		{
			get { return _versao; }
			set { _versao = value; }
		}
		#endregion

		#region [ Construtor ]
		public LinhaHeaderArquivoIbptCsv()
		{
			inicializaCampos();
		}

		public LinhaHeaderArquivoIbptCsv(String linhaDados)
		{
			inicializaCampos();
			carregaDados(linhaDados);
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaCampos ]
		private void inicializaCampos()
		{
			_codigo = "";
			_ex = "";
			_tabela = "";
			_descricao = "";
			_aliqNac = "";
			_aliqImp = "";
			_versao = "";
		}
		#endregion

		#region [ carregaDados ]
		public void carregaDados(String linhaDados)
		{
			#region [ Declarações ]
			String[] v;
			#endregion

			inicializaCampos();

			v = linhaDados.Split(';');
			for (int i = 0; i < v.Length; i++)
			{
				switch (i)
				{
					case 0:
						_codigo = (v[i] == null ? "" : v[i]);
						break;
					case 1:
						_ex = (v[i] == null ? "" : v[i]);
						break;
					case 2:
						_tabela = (v[i] == null ? "" : v[i]);
						break;
					case 3:
						_descricao = (v[i] == null ? "" : v[i]);
						break;
					case 4:
						_aliqNac = (v[i] == null ? "" : v[i]);
						break;
					case 5:
						_aliqImp = (v[i] == null ? "" : v[i]);
						break;
					case 6:
						_versao = (v[i] == null ? "" : v[i]);
						break;
					default:
						break;
				}
			}
		}
		#endregion

		#endregion
	}
	#endregion

	#region [ Classe: LinhaDadosArquivoIbptCsv ]
	public class LinhaDadosArquivoIbptCsv
	{
		#region [ Getters / Setters ]
		private String _codigo;
		public String codigo
		{
			get { return _codigo; }
			set { _codigo = value; }
		}

		private String _ex;
		public String ex
		{
			get { return _ex; }
			set { _ex = value; }
		}

		private String _tabela;
		public String tabela
		{
			get { return _tabela; }
			set { _tabela = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private String _aliqNac;
		public String aliqNac
		{
			get { return _aliqNac; }
			set { _aliqNac = value; }
		}

		private String _aliqImp;
		public String aliqImp
		{
			get { return _aliqImp; }
			set { _aliqImp = value; }
		}
		#endregion

		#region [ Construtor ]
		public LinhaDadosArquivoIbptCsv()
		{
			inicializaCampos();
		}

		public LinhaDadosArquivoIbptCsv(String linhaDados)
		{
			inicializaCampos();
			carregaDados(linhaDados);
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaCampos ]
		private void inicializaCampos()
		{
			_codigo = "";
			_ex = "";
			_tabela = "";
			_descricao = "";
			_aliqNac = "";
			_aliqImp = "";
		}
		#endregion

		#region [ carregaDados ]
		public void carregaDados(String linhaDados)
		{
			#region [ Declarações ]
			String[] v;
			#endregion

			inicializaCampos();

			v = linhaDados.Split(';');
			for (int i = 0; i < v.Length; i++)
			{
				switch (i)
				{
					case 0:
						_codigo = (v[i] == null ? "" : v[i]);
						break;
					case 1:
						_ex = (v[i] == null ? "" : v[i]);
						break;
					case 2:
						_tabela = (v[i] == null ? "" : v[i]);
						break;
					case 3:
						_descricao = (v[i] == null ? "" : v[i]);
						break;
					case 4:
						_aliqNac = (v[i] == null ? "" : v[i]);
						break;
					case 5:
						_aliqImp = (v[i] == null ? "" : v[i]);
						break;
					default:
						break;
				}
			}
		}
		#endregion

		#endregion
	}
	#endregion
}
