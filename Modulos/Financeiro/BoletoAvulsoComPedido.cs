#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	#region [ BoletoAvulsoComPedido ]
	class BoletoAvulsoComPedido
	{
		#region [ Getters/Setters ]

		private List<BoletoAvulsoComPedidoItem> _listaItem;
		public List<BoletoAvulsoComPedidoItem> listaItem
		{
			get { return _listaItem; }
			set { _listaItem = value; }
		}

		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private byte _qtde_parcelas;
		public byte qtde_parcelas
		{
			get { return _qtde_parcelas; }
			set { _qtde_parcelas = value; }
		}

		private byte _qtde_parcelas_boleto;
		public byte qtde_parcelas_boleto
		{
			get { return _qtde_parcelas_boleto; }
			set { _qtde_parcelas_boleto = value; }
		}

		#endregion
	}
	#endregion

	#region [ BoletoAvulsoComPedidoItem ]
	class BoletoAvulsoComPedidoItem
	{
		#region [ Getters/Setters ]
		
		private byte _num_parcela;
		public byte num_parcela
		{
			get { return _num_parcela; }
			set { _num_parcela = value; }
		}

		private short _forma_pagto;
		public short forma_pagto
		{
			get { return _forma_pagto; }
			set { _forma_pagto = value; }
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

		private List<BoletoAvulsoComPedidoItemRateio> _listaRateio;
		public List<BoletoAvulsoComPedidoItemRateio> listaRateio
		{
			get { return _listaRateio; }
			set { _listaRateio = value; }
		}
		
		#endregion
	}
	#endregion

	#region [ BoletoAvulsoComPedidoItemRateio ]
	class BoletoAvulsoComPedidoItemRateio
	{
		#region [ Getters/Setters ]
		
		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
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
}
