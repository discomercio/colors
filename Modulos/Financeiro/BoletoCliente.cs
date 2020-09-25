using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.ComponentModel.Com2Interop;

namespace Financeiro
{
	class BoletoCliente
	{
		public string nome { get; set; }
		public string cnpj_cpf { get; set; }
		public string email { get; set; }
		public string endereco_logradouro { get; set; }
		public string endereco_numero { get; set; }
		public string endereco_complemento { get; set; }
		public string endereco_cep { get; set; }
		public string endereco_bairro { get; set; }
		public string endereco_cidade { get; set; }
		public string endereco_uf { get; set; }
	}
}
