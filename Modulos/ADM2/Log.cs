#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace ADM2
{
	public class Log
	{
		#region [ Getters/Setters ]
		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private String _usuario;
		public String usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private String _loja;
		public String loja
		{
			get { return _loja; }
			set { _loja = value; }
		}

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private String _operacao;
		public String operacao
		{
			get { return _operacao; }
			set { _operacao = value; }
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
