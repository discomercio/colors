#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	class PedidoHistPagto
	{
		#region [ Getters/Setters ]

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private byte _status;
		public byte status
		{
			get { return _status; }
			set { _status = value; }
		}

		private int _id_fluxo_caixa;
		public int id_fluxo_caixa
		{
			get { return _id_fluxo_caixa; }
			set { _id_fluxo_caixa = value; }
		}

		private int _ctrl_pagto_id_parcela;
		public int ctrl_pagto_id_parcela
		{
			get { return _ctrl_pagto_id_parcela; }
			set { _ctrl_pagto_id_parcela = value; }
		}

		private byte _ctrl_pagto_modulo;
		public byte ctrl_pagto_modulo
		{
			get { return _ctrl_pagto_modulo; }
			set { _ctrl_pagto_modulo = value; }
		}

		private DateTime _dt_vencto;
		public DateTime dt_vencto
		{
			get { return _dt_vencto; }
			set { _dt_vencto = value; }
		}

		private decimal _valor_total;
		public decimal valor_total
		{
			get { return _valor_total; }
			set { _valor_total = value; }
		}

		private decimal _valor_rateado;
		public decimal valor_rateado
		{
			get { return _valor_rateado; }
			set { _valor_rateado = value; }
		}

		private String _descricao;
		public String descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private DateTime _dt_credito;
		public DateTime dt_credito
		{
			get { return _dt_credito; }
			set { _dt_credito = value; }
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

		private decimal _vl_abatimento_concedido;
		public decimal vl_abatimento_concedido
		{
			get { return _vl_abatimento_concedido; }
			set { _vl_abatimento_concedido = value; }
		}

		private byte _st_boleto_pago_cheque;
		public byte st_boleto_pago_cheque
		{
			get { return _st_boleto_pago_cheque; }
			set { _st_boleto_pago_cheque = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_pago_cheque;
		public DateTime dt_ocorrencia_banco_boleto_pago_cheque
		{
			get { return _dt_ocorrencia_banco_boleto_pago_cheque; }
			set { _dt_ocorrencia_banco_boleto_pago_cheque = value; }
		}

		private byte _st_boleto_ocorrencia_17;
		public byte st_boleto_ocorrencia_17
		{
			get { return _st_boleto_ocorrencia_17; }
			set { _st_boleto_ocorrencia_17 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_17;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_17
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_17; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_17 = value; }
		}

		private byte _st_boleto_ocorrencia_15;
		public byte st_boleto_ocorrencia_15
		{
			get { return _st_boleto_ocorrencia_15; }
			set { _st_boleto_ocorrencia_15 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_15;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_15
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_15; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_15 = value; }
		}

		private byte _st_boleto_ocorrencia_23;
		public byte st_boleto_ocorrencia_23
		{
			get { return _st_boleto_ocorrencia_23; }
			set { _st_boleto_ocorrencia_23 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_23;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_23
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_23; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_23 = value; }
		}

		private byte _st_boleto_ocorrencia_34;
		public byte st_boleto_ocorrencia_34
		{
			get { return _st_boleto_ocorrencia_34; }
			set { _st_boleto_ocorrencia_34 = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_ocorrencia_34;
		public DateTime dt_ocorrencia_banco_boleto_ocorrencia_34
		{
			get { return _dt_ocorrencia_banco_boleto_ocorrencia_34; }
			set { _dt_ocorrencia_banco_boleto_ocorrencia_34 = value; }
		}

		private byte _st_boleto_baixado;
		public byte st_boleto_baixado
		{
			get { return _st_boleto_baixado; }
			set { _st_boleto_baixado = value; }
		}

		private DateTime _dt_ocorrencia_banco_boleto_baixado;
		public DateTime dt_ocorrencia_banco_boleto_baixado
		{
			get { return _dt_ocorrencia_banco_boleto_baixado; }
			set { _dt_ocorrencia_banco_boleto_baixado = value; }
		}

		#endregion
	}
}
