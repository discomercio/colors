#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	#region [ SerasaCliente ]
	class SerasaCliente
	{
		#region [ Getters/Setters ]
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private string _id_cliente;
		public string id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private string _cnpj;
		public string cnpj
		{
			get { return _cnpj; }
			set { _cnpj = value; }
		}

		private string _raiz_cnpj;
		public string raiz_cnpj
		{
			get { return _raiz_cnpj; }
			set { _raiz_cnpj = value; }
		}

		private DateTime _dt_cliente_desde;
		public DateTime dt_cliente_desde
		{
			get { return _dt_cliente_desde; }
			set { _dt_cliente_desde = value; }
		}

		private byte _st_enviado_serasa;
		public byte st_enviado_serasa
		{
			get { return _st_enviado_serasa; }
			set { _st_enviado_serasa = value; }
		}

		private DateTime _dt_enviado_serasa;
		public DateTime dt_enviado_serasa
		{
			get { return _dt_enviado_serasa; }
			set { _dt_enviado_serasa = value; }
		}

		private int _id_serasa_arq_remessa_normal;
		public int id_serasa_arq_remessa_normal
		{
			get { return _id_serasa_arq_remessa_normal; }
			set { _id_serasa_arq_remessa_normal = value; }
		}
		#endregion
	}
	#endregion

	#region [ SerasaTituloMovimento ]
	class SerasaTituloMovimento
	{
		#region [ Getters/Setters ]
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_boleto_arq_retorno;
		public int id_boleto_arq_retorno
		{
			get { return _id_boleto_arq_retorno; }
			set { _id_boleto_arq_retorno = value; }
		}

		private int _id_boleto_item;
		public int id_boleto_item
		{
			get { return _id_boleto_item; }
			set { _id_boleto_item = value; }
		}

		private int _id_serasa_cliente;
		public int id_serasa_cliente
		{
			get { return _id_serasa_cliente; }
			set { _id_serasa_cliente = value; }
		}

		private string _cnpj;
		public string cnpj
		{
			get { return _cnpj; }
			set { _cnpj = value; }
		}

		private string _raiz_cnpj;
		public string raiz_cnpj
		{
			get { return _raiz_cnpj; }
			set { _raiz_cnpj = value; }
		}

		private DateTime _dt_cadastro;
		public DateTime dt_cadastro
		{
			get { return _dt_cadastro; }
			set { _dt_cadastro = value; }
		}

		private DateTime _dt_hr_cadastro;
		public DateTime dt_hr_cadastro
		{
			get { return _dt_hr_cadastro; }
			set { _dt_hr_cadastro = value; }
		}

		private string _identificacao_ocorrencia_boleto;
		public string identificacao_ocorrencia_boleto
		{
			get { return _identificacao_ocorrencia_boleto; }
			set { _identificacao_ocorrencia_boleto = value; }
		}

		private byte _st_envio_serasa_cancelado;
		public byte st_envio_serasa_cancelado
		{
			get { return _st_envio_serasa_cancelado; }
			set { _st_envio_serasa_cancelado = value; }
		}

		private DateTime _dt_envio_serasa_cancelado;
		public DateTime dt_envio_serasa_cancelado
		{
			get { return _dt_envio_serasa_cancelado; }
			set { _dt_envio_serasa_cancelado = value; }
		}

		private DateTime _dt_hr_envio_serasa_cancelado;
		public DateTime dt_hr_envio_serasa_cancelado
		{
			get { return _dt_hr_envio_serasa_cancelado; }
			set { _dt_hr_envio_serasa_cancelado = value; }
		}

		private string _usuario_envio_serasa_cancelado;
		public string usuario_envio_serasa_cancelado
		{
			get { return _usuario_envio_serasa_cancelado; }
			set { _usuario_envio_serasa_cancelado = value; }
		}

		private byte _st_enviado_serasa;
		public byte st_enviado_serasa
		{
			get { return _st_enviado_serasa; }
			set { _st_enviado_serasa = value; }
		}

		private int _id_serasa_arq_remessa_normal;
		public int id_serasa_arq_remessa_normal
		{
			get { return _id_serasa_arq_remessa_normal; }
			set { _id_serasa_arq_remessa_normal = value; }
		}

		private byte _st_retorno_serasa;
		public byte st_retorno_serasa
		{
			get { return _st_retorno_serasa; }
			set { _st_retorno_serasa = value; }
		}

		private int _id_serasa_arq_retorno_normal;
		public int id_serasa_arq_retorno_normal
		{
			get { return _id_serasa_arq_retorno_normal; }
			set { _id_serasa_arq_retorno_normal = value; }
		}

		private byte _st_processado_serasa_sucesso;
		public byte st_processado_serasa_sucesso
		{
			get { return _st_processado_serasa_sucesso; }
			set { _st_processado_serasa_sucesso = value; }
		}

		private byte _st_editado_manual;
		public byte st_editado_manual
		{
			get { return _st_editado_manual; }
			set { _st_editado_manual = value; }
		}

		private DateTime _dt_editado_manual;
		public DateTime dt_editado_manual
		{
			get { return _dt_editado_manual; }
			set { _dt_editado_manual = value; }
		}

		private DateTime _dt_hr_editado_manual;
		public DateTime dt_hr_editado_manual
		{
			get { return _dt_hr_editado_manual; }
			set { _dt_hr_editado_manual = value; }
		}

		private string _usuario_editado_manual;
		public string usuario_editado_manual
		{
			get { return _usuario_editado_manual; }
			set { _usuario_editado_manual = value; }
		}

		private int _qtde_vezes_editado_manual;
		public int qtde_vezes_editado_manual
		{
			get { return _qtde_vezes_editado_manual; }
			set { _qtde_vezes_editado_manual = value; }
		}

		private string _numero_documento;
		public string numero_documento
		{
			get { return _numero_documento; }
			set { _numero_documento = value; }
		}

		private string _nosso_numero;
		public string nosso_numero
		{
			get { return _nosso_numero; }
			set { _nosso_numero = value; }
		}

		private string _digito_nosso_numero;
		public string digito_nosso_numero
		{
			get { return _digito_nosso_numero; }
			set { _digito_nosso_numero = value; }
		}

		private DateTime _dt_emissao;
		public DateTime dt_emissao
		{
			get { return _dt_emissao; }
			set { _dt_emissao = value; }
		}

		private decimal _vl_titulo;
		public decimal vl_titulo
		{
			get { return _vl_titulo; }
			set { _vl_titulo = value; }
		}

		private DateTime _dt_vencto;
		public DateTime dt_vencto
		{
			get { return _dt_vencto; }
			set { _dt_vencto = value; }
		}

		private DateTime _dt_pagto;
		public DateTime dt_pagto
		{
			get { return _dt_pagto; }
			set { _dt_pagto = value; }
		}

		private decimal _vl_pago;
		public decimal vl_pago
		{
			get { return _vl_pago; }
			set { _vl_pago = value; }
		}

		private string _retorno_codigos_erro;
		public string retorno_codigos_erro
		{
			get { return _retorno_codigos_erro; }
			set { _retorno_codigos_erro = value; }
		}
		#endregion
	}
	#endregion
}
