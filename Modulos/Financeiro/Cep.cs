#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class Cep
	{
		#region [ Getters/Setters ]

		private String _cep;
		public String cep
		{
			get { return _cep; }
			set { _cep = value; }
		}

		private String _uf;
		public String uf
		{
			get { return _uf; }
			set { _uf = value; }
		}

		private String _cidade;
		public String cidade
		{
			get { return _cidade; }
			set { _cidade = value; }
		}

		private String _bairro;
		public String bairro
		{
			get { return _bairro; }
			set { _bairro = value; }
		}

		private String _logradouro;
		public String logradouro
		{
			get { return _logradouro; }
			set { _logradouro = value; }
		}

		private String _complemento;
		public String complemento
		{
			get { return _complemento; }
			set { _complemento = value; }
		}

		#endregion
	}
}
