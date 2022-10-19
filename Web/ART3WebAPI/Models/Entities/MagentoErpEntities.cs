using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	#region [ MagentoErpSalesOrder ]
	public class MagentoErpSalesOrder
	{
		public string numeroPedidoMagento { get; set; } = "";
		public string cpfCnpjIdentificado { get; set; } = "";
		public MagentoErpSalesOrderJaCadastrado erpSalesOrderJaCadastrado { get; set; } = new MagentoErpSalesOrderJaCadastrado();
		public MagentoErpSalesOrderCliente erpCliente { get; set; } = new MagentoErpSalesOrderCliente();
		public MagentoSoapApiSalesOrderInfo magentoSalesOrderInfo { get; set; } = null;
		public Magento2SalesOrderInfo magento2SalesOrderInfo { get; set; } = null;
	}
	#endregion

	#region [ MagentoErpSalesOrderCliente ]
	public class MagentoErpSalesOrderCliente
	{
		public string id_cliente { get; set; } = "";
		public string cnpj_cpf { get; set; } = "";
		public string nome { get; set; } = "";
	}
	#endregion

	#region [ MagentoErpSalesOrderJaCadastrado ]
	public class MagentoErpSalesOrderJaCadastrado
	{
		public string pedido { get; set; } = "";
		public string dt_cadastro { get; set; } = "";
		public string dt_cadastro_formatado { get; set; }
		public string dt_hr_cadastro { get; set; } = "";
		public string dt_hr_cadastro_formatado { get; set; }
		public string vendedor { get; set; } = "";
		public string usuario_cadastro { get; set; } = "";
	}
	#endregion

	#region [ Class MagentoErpPedidoXml ]
	public class MagentoErpPedidoXml
	{
		public int id { get; set; }
		public string operationControlTicket { get; set; }
		public string loja { get; set; }
		public string pedido_magento { get; set; }
		public string pedido_erp { get; set; }
		public string pedido_marketplace { get; set; }
		public string pedido_marketplace_completo { get; set; }
		public string marketplace_codigo_origem { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public int magento_api_versao { get; set; }
		public string pedido_xml { get; set; }
		public string pedido_json { get; set; }
		public string cpfCnpjIdentificado { get; set; }
		public int increment_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public int customer_id { get; set; }
		public int billing_address_id { get; set; }
		public int shipping_address_id { get; set; }
		public string status { get; set; }
		public string status_descricao { get; set; }
		public string state { get; set; }
		public string state_descricao { get; set; }
		public string customer_email { get; set; }
		public string customer_firstname { get; set; }
		public string customer_lastname { get; set; }
		public string customer_middlename { get; set; }
		public int quote_id { get; set; }
		public int customer_group_id { get; set; }
		public int order_id { get; set; }
		public string customer_dob { get; set; }
		public string clearsale_status_code { get; set; }
		public string clearSale_status { get; set; }
		public string clearSale_score { get; set; }
		public string clearSale_packageID { get; set; }
		public decimal shipping_amount { get; set; }
		public decimal shipping_discount_amount { get; set; }
		public decimal discount_amount { get; set; }
		public decimal subtotal { get; set; }
		public decimal grand_total { get; set; }
		public string installer_document { get; set; }
		public int installer_id { get; set; }
		public decimal commission_value { get; set; }
		public decimal commission_discount { get; set; }
		public decimal commission_final_discount { get; set; }
		public decimal commission_final_value { get; set; }
		public string commission_discount_type { get; set; }
		public byte mktp_datasource_status { get; set; } = 0;
		public decimal mktp_datasource_discount { get; set; } = 0m;
		public decimal mktp_datasource_total_ordered { get; set; } = 0m;
		public decimal mktp_datasource_shipping_cost { get; set; } = 0m;
		public string b2b_installer_name { get; set; }
		public int b2b_installer_id { get; set; }
		public decimal b2b_installer_commission_value { get; set; }
		public double b2b_installer_commission_percentage { get; set; } = 0d;
		public string b2b_type_order { get; set; }
	}
	#endregion

	#region [ MagentoErpPedidoXmlDecodeEndereco ]
	public class MagentoErpPedidoXmlDecodeEndereco
	{
		public int id { get; set; }
		public int id_magento_api_pedido_xml { get; set; }
		public string tipo_endereco { get; set; }
		public string endereco { get; set; }
		public string endereco_numero { get; set; }
		public string endereco_complemento { get; set; }
		public string bairro { get; set; }
		public string cidade { get; set; }
		public string uf { get; set; }
		public string cep { get; set; }
		public int address_id { get; set; }
		public int parent_id { get; set; }
		public int customer_address_id { get; set; }
		public int quote_address_id { get; set; }
		public int region_id { get; set; }
		public string address_type { get; set; }
		public string street { get; set; }
		public string city { get; set; }
		public string region { get; set; }
		public string postcode { get; set; }
		public string country_id { get; set; }
		public string firstname { get; set; }
		public string middlename { get; set; }
		public string lastname { get; set; }
		public string email { get; set; }
		public string telephone { get; set; }
		public string celular { get; set; }
		public string fax { get; set; }
		public string tipopessoa { get; set; }
		public string rg { get; set; }
		public string ie { get; set; }
		public string cpfcnpj { get; set; }
		public string empresa { get; set; }
		public string nomefantasia { get; set; }
		public string street_detail { get; set; }
	}
	#endregion

	#region [ MagentoErpPedidoXmlDecodeItem ]
	public class MagentoErpPedidoXmlDecodeItem
	{
		public int id { get; set; }
		public int id_magento_api_pedido_xml { get; set; }
		public string sku { get; set; }
		public decimal qty_ordered { get; set; }
		public int product_id { get; set; }
		public int item_id { get; set; }
		public int order_id { get; set; }
		public int quote_item_id { get; set; }
		public decimal price { get; set; }
		public decimal base_price { get; set; }
		public decimal original_price { get; set; }
		public decimal base_original_price { get; set; }
		public decimal discount_percent { get; set; }
		public decimal discount_amount { get; set; }
		public decimal base_discount_amount { get; set; }
		public string name { get; set; }
		public string product_type { get; set; }
		public string has_children { get; set; }
		public int parent_item_id { get; set; }
		public double weight { get; set; } = 0d;
		public int is_virtual { get; set; } = 0;
		public int free_shipping { get; set; } = 0;
		public int is_qty_decimal { get; set; } = 0;
		public int no_discount { get; set; } = 0;
		public decimal qty_canceled { get; set; } = 0m;
		public decimal qty_invoiced { get; set; } = 0m;
		public decimal qty_refunded { get; set; } = 0m;
		public decimal qty_shipped { get; set; } = 0m;
		public double tax_percent { get; set; } = 0d;
		public decimal tax_amount { get; set; } = 0m;
		public decimal base_tax_amount { get; set; } = 0m;
		public decimal tax_invoiced { get; set; } = 0m;
		public decimal base_tax_invoiced { get; set; } = 0m;
		public decimal discount_invoiced { get; set; } = 0m;
		public decimal base_discount_invoiced { get; set; } = 0m;
		public decimal amount_refunded { get; set; } = 0m;
		public decimal base_amount_refunded { get; set; } = 0m;
		public decimal row_total { get; set; } = 0m;
		public decimal base_row_total { get; set; } = 0m;
		public decimal row_invoiced { get; set; } = 0m;
		public decimal base_row_invoiced { get; set; } = 0m;
		public double row_weight { get; set; } = 0d;
		public decimal price_incl_tax { get; set; } = 0m;
		public decimal base_price_incl_tax { get; set; } = 0m;
		public decimal row_total_incl_tax { get; set; } = 0m;
		public decimal base_row_total_incl_tax { get; set; } = 0m;
		public decimal mktp_datasource_special_price { get; set; } = 0m;
		public decimal mktp_datasource_shipping_cost { get; set; } = 0m;
		public decimal mktp_datasource_original_price { get; set; } = 0m;
	}
	#endregion

	#region [ MagentoErpPedidoXmlDecodeStatusHistory ]
	public class MagentoErpPedidoXmlDecodeStatusHistory
	{
		public int id { get; set; }
		public int id_magento_api_pedido_xml { get; set; }
		public int parent_id { get; set; }
		public byte is_customer_notified { get; set; }
		public byte is_visible_on_front { get; set; }
		public string comment { get; set; }
		public string status { get; set; }
		public string created_at { get; set; }
		public string entity_name { get; set; }
		public int store_id { get; set; }
	}
	#endregion

	#region [ MagentoErpPedidoXmlDecodeSkyhubMktpPayment ]
	public class MagentoErpPedidoXmlDecodeSkyhubMktpPayment
	{
		public int id { get; set; }
		public int id_magento_api_pedido_xml { get; set; }
		public decimal? value { get; set; } = null;
		public string type { get; set; }
		public string transaction_date { get; set; }
		public string status { get; set; }
		public int? parcels { get; set; } = null;
		public string method { get; set; }
		public string description { get; set; }
		public string card_issuer { get; set; }
		public string autorization_id { get; set; }
		public string sefaz_type_integration { get; set; }
		public string sefaz_payment_indicator { get; set; }
		public string sefaz_name_payment { get; set; }
		public string sefaz_name_card_issuer { get; set; }
		public string sefaz_id_payment { get; set; }
		public string sefaz_id_card_issuer { get; set; }
	}
	#endregion
}