using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FinanceiroService
{
	class EmailCtrl
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_emailsndsvc_mensagem;
		public int id_emailsndsvc_mensagem
		{
			get { return _id_emailsndsvc_mensagem; }
			set { _id_emailsndsvc_mensagem = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _pedido;
		public string pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private string _id_cliente;
		public string id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private string _cnpj_cpf_cliente;
		public string cnpj_cpf_cliente
		{
			get { return _cnpj_cpf_cliente; }
			set { _cnpj_cpf_cliente = value; }
		}

		private string _tipo_destinatario;
		public string tipo_destinatario
		{
			get { return _tipo_destinatario; }
			set { _tipo_destinatario = value; }
		}

		private string _modulo;
		public string modulo
		{
			get { return _modulo; }
			set { _modulo = value; }
		}

		private string _tipo_msg;
		public string tipo_msg
		{
			get { return _tipo_msg; }
			set { _tipo_msg = value; }
		}

		private string _codigo_msg;
		public string codigo_msg
		{
			get { return _codigo_msg; }
			set { _codigo_msg = value; }
		}

		private string _rotina;
		public string rotina
		{
			get { return _rotina; }
			set { _rotina = value; }
		}

		private string _remetente;
		public string remetente
		{
			get { return _remetente; }
			set { _remetente = value; }
		}

		private string _destinatario;
		public string destinatario
		{
			get { return _destinatario; }
			set { _destinatario = value; }
		}
	}
}
