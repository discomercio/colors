using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	#region [ Loja ]
	public class Loja
	{
		public string loja { get; set; } = "";
		public string cnpj { get; set; } = "";
		public string ie { get; set; } = "";
		public string nome { get; set; } = "";
		public string razao_social { get; set; } = "";
		public string endereco { get; set; } = "";
		public string endereco_numero { get; set; } = "";
		public string endereco_complemento { get; set; } = "";
		public string bairro { get; set; } = "";
		public string cidade { get; set; } = "";
		public string uf { get; set; } = "";
		public string cep { get; set; } = "";
		public string ddd { get; set; } = "";
		public string telefone { get; set; } = "";
		public string fax { get; set; } = "";
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public double comissao_indicacao { get; set; } = 0d;
		public double PercMaxSenhaDesconto { get; set; } = 0d;
		public byte id_plano_contas_empresa { get; set; }
		public int id_plano_contas_grupo { get; set; }
		public int id_plano_contas_conta { get; set; }
		public string natureza { get; set; } = "";
		public double PercMaxDescSemZerarRT { get; set; } = 0d;
		public double perc_max_comissao { get; set; } = 0d;
		public double perc_max_comissao_e_desconto { get; set; } = 0d;
		public double perc_max_comissao_e_desconto_nivel2 { get; set; } = 0d;
		public double perc_max_comissao_e_desconto_nivel2_pj { get; set; } = 0d;
		public double perc_max_comissao_e_desconto_pj { get; set; } = 0d;
		public string unidade_negocio { get; set; } = "";
		public int magento_api_versao { get; set; } = 0;
		public string magento_api_urlWebService { get; set; } = "";
		public string magento_api_username { get; set; } = "";
		public string magento_api_password { get; set; } = "";
		public string magento_api_rest_endpoint { get; set; } = "";
		public string magento_api_rest_access_token { get; set; } = "";
		public byte magento_api_rest_force_get_sales_order_by_entity_id { get; set; } = 0;
	}
	#endregion
}
