#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	#region [ BoletoPreCadastrado ]
	class BoletoPreCadastrado
	{
		#region [ Getters/Setters ]

		private List<BoletoPreCadastradoItem> _listaItem;
		public List<BoletoPreCadastradoItem> listaItem
		{
			get { return _listaItem; }
			set { _listaItem = value; }
		}

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private String _id_cliente;
		public String id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
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

		private byte _qtde_parcelas_boleto;
		public byte qtde_parcelas_boleto
		{
			get { return _qtde_parcelas_boleto; }
			set { _qtde_parcelas_boleto = value; }
		}

		private byte _status;
		public byte status
		{
			get { return _status; }
			set { _status = value; }
		}

		private DateTime _dt_cadastro;
		public DateTime dt_cadastro
		{
			get { return _dt_cadastro; }
			set { _dt_cadastro = value; }
		}

		private String _usuario_cadastro;
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

		private String _usuario_ult_atualizacao;
		public String usuario_ult_atualizacao
		{
			get { return _usuario_ult_atualizacao; }
			set { _usuario_ult_atualizacao = value; }
		}

		#endregion
	}
	#endregion

	#region [ BoletoPreCadastradoItem ]
	class BoletoPreCadastradoItem
	{
		#region [ Getters/Setters ]

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_nf_parcela_pagto;
		public int id_nf_parcela_pagto
		{
			get { return _id_nf_parcela_pagto; }
			set { _id_nf_parcela_pagto = value; }
		}

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

		private List<BoletoPreCadastradoItemRateio> _listaRateio;
		public List<BoletoPreCadastradoItemRateio> listaRateio
		{
			get { return _listaRateio; }
			set { _listaRateio = value; }
		}
		#endregion
	}
	#endregion

	#region [ BoletoPreCadastradoItemRateio ]
	class BoletoPreCadastradoItemRateio
	{
		#region [ Getters/Setters ]

		private int _id_nf_parcela_pagto_item;
		public int id_nf_parcela_pagto_item
		{
			get { return _id_nf_parcela_pagto_item; }
			set { _id_nf_parcela_pagto_item = value; }
		}

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private int _id_nf_parcela_pagto;
		public int id_nf_parcela_pagto
		{
			get { return _id_nf_parcela_pagto; }
			set { _id_nf_parcela_pagto = value; }
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
