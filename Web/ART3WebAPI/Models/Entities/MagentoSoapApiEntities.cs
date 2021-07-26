using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	#region [ MagentoSoapApiSalesOrderInfo ]
	public class MagentoSoapApiSalesOrderInfo
	{
		public string increment_id { get; set; }
		public string parent_id { get; set; }
		public string store_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string is_active { get; set; }
		public string customer_id { get; set; }
		public string tax_amount { get; set; }
		public string tax_canceled { get; set; }
		public string tax_invoiced { get; set; }
		public string tax_refunded { get; set; }
		public string shipping_amount { get; set; }
		public string shipping_canceled { get; set; }
		public string shipping_invoiced { get; set; }
		public string shipping_refunded { get; set; }
		public string shipping_tax_amount { get; set; }
		public string shipping_tax_refunded { get; set; }
		public string shipping_discount_amount { get; set; }
		public string discount_amount { get; set; }
		public string discount_canceled { get; set; }
		public string discount_invoiced { get; set; }
		public string discount_refunded { get; set; }
		public string subtotal { get; set; }
		public string subtotal_canceled { get; set; }
		public string subtotal_invoiced { get; set; }
		public string subtotal_refunded { get; set; }
		public string subtotal_incl_tax { get; set; }
		public string grand_total { get; set; }
		public string total_paid { get; set; }
		public string total_refunded { get; set; }
		public string total_qty_ordered { get; set; }
		public string total_canceled { get; set; }
		public string total_invoiced { get; set; }
		public string total_due { get; set; }
		public string total_online_refunded { get; set; }
		public string total_offline_refunded { get; set; }
		public string base_tax_amount { get; set; }
		public string base_tax_canceled { get; set; }
		public string base_tax_invoiced { get; set; }
		public string base_tax_refunded { get; set; }
		public string base_shipping_amount { get; set; }
		public string base_shipping_canceled { get; set; }
		public string base_shipping_invoiced { get; set; }
		public string base_shipping_refunded { get; set; }
		public string base_shipping_tax_amount { get; set; }
		public string base_shipping_tax_refunded { get; set; }
		public string base_discount_amount { get; set; }
		public string base_discount_canceled { get; set; }
		public string base_discount_invoiced { get; set; }
		public string base_discount_refunded { get; set; }
		public string base_subtotal { get; set; }
		public string base_subtotal_canceled { get; set; }
		public string base_subtotal_invoiced { get; set; }
		public string base_subtotal_refunded { get; set; }
		public string base_grand_total { get; set; }
		public string base_total_paid { get; set; }
		public string base_total_refunded { get; set; }
		public string base_total_qty_ordered { get; set; }
		public string base_total_canceled { get; set; }
		public string base_total_invoiced { get; set; }
		public string base_total_invoiced_cost { get; set; }
		public string base_total_online_refunded { get; set; }
		public string base_total_offline_refunded { get; set; }
		public string billing_address_id { get; set; }
		public string billing_firstname { get; set; }
		public string billing_lastname { get; set; }
		public string shipping_address_id { get; set; }
		public string shipping_firstname { get; set; }
		public string shipping_lastname { get; set; }
		public string billing_name { get; set; }
		public string shipping_name { get; set; }
		public string store_to_base_rate { get; set; }
		public string store_to_order_rate { get; set; }
		public string base_to_global_rate { get; set; }
		public string base_to_order_rate { get; set; }
		public string weight { get; set; }
		public string store_name { get; set; }
		public string remote_ip { get; set; }
		public string status { get; set; }
		public string state { get; set; }
		public string applied_rule_ids { get; set; }
		public string global_currency_code { get; set; }
		public string base_currency_code { get; set; }
		public string store_currency_code { get; set; }
		public string order_currency_code { get; set; }
		public string shipping_method { get; set; }
		public string shipping_description { get; set; }
		public string customer_email { get; set; }
		public string customer_firstname { get; set; }
		public string customer_lastname { get; set; }
		public string customer_middlename { get; set; }
		public string customer_prefix { get; set; }
		public string customer_suffix { get; set; }
		public string customer_taxvat { get; set; }
		public string quote_id { get; set; }
		public string is_virtual { get; set; }
		public string customer_group_id { get; set; }
		public string customer_note { get; set; }
		public string customer_note_notify { get; set; }
		public string customer_is_guest { get; set; }
		public string email_sent { get; set; }
		public string order_id { get; set; }
		public string gift_message_id { get; set; }
		public string gift_message { get; set; }
		public string coupon_code { get; set; }
		public string protect_code { get; set; }
		public string can_ship_partially { get; set; }
		public string can_ship_partially_item { get; set; }
		public string edit_increment { get; set; }
		public string forced_shipment_with_invoice { get; set; }
		public string forced_do_shipment_with_invoice { get; set; }
		public string payment_auth_expiration { get; set; }
		public string quote_address_id { get; set; }
		public string adjustment_negative { get; set; }
		public string adjustment_positive { get; set; }
		public string base_adjustment_negative { get; set; }
		public string base_adjustment_positive { get; set; }
		public string base_shipping_discount_amount { get; set; }
		public string base_subtotal_incl_tax { get; set; }
		public string base_total_due { get; set; }
		public string payment_authorization_amount { get; set; }
		public string customer_dob { get; set; }
		public string discount_description { get; set; }
		public string ext_customer_id { get; set; }
		public string ext_order_id { get; set; }
		public string hold_before_state { get; set; }
		public string hold_before_status { get; set; }
		public string original_increment_id { get; set; }
		public string relation_child_id { get; set; }
		public string relation_child_real_id { get; set; }
		public string relation_parent_id { get; set; }
		public string relation_parent_real_id { get; set; }
		public string x_forwarded_for { get; set; }
		public string total_item_count { get; set; }
		public string customer_gender { get; set; }
		public string hidden_tax_amount { get; set; }
		public string base_hidden_tax_amount { get; set; }
		public string shipping_hidden_tax_amount { get; set; }
		public string base_shipping_hidden_tax_amnt { get; set; }
		public string hidden_tax_invoiced { get; set; }
		public string base_hidden_tax_invoiced { get; set; }
		public string hidden_tax_refunded { get; set; }
		public string base_hidden_tax_refunded { get; set; }
		public string shipping_incl_tax { get; set; }
		public string base_shipping_incl_tax { get; set; }
		public string coupon_rule_name { get; set; }
		public string paypal_ipn_customer_notified { get; set; }
		public string firecheckout_delivery_date { get; set; }
		public string firecheckout_delivery_timerange { get; set; }
		public string firecheckout_customer_comment { get; set; }
		public string tm_field1 { get; set; }
		public string tm_field2 { get; set; }
		public string tm_field3 { get; set; }
		public string tm_field4 { get; set; }
		public string tm_field5 { get; set; }
		public string from_lengow { get; set; }
		public string order_id_lengow { get; set; }
		public string fees_lengow { get; set; }
		public string xml_node_lengow { get; set; }
		public string feed_id_lengow { get; set; }
		public string message_lengow { get; set; }
		public string marketplace_lengow { get; set; }
		public string total_paid_lengow { get; set; }
		public string carrier_lengow { get; set; }
		public string carrier_method_lengow { get; set; }
		public string clearsale_status_code { get; set; }
		public string session_id { get; set; }
		public string skyhub_code { get; set; }
		public string commission_value { get; set; }
		public string installer_document { get; set; }
		public string installer_id { get; set; }
		public string commission_discount { get; set; }
		public string commission_final_discount { get; set; }
		public string commission_discount_type { get; set; }
		public string commission_final_value { get; set; }
		public string base_bseller_payment_total_tax_rate { get; set; }
		public string bseller_payment_total_tax_rate { get; set; }
		public string payment_authorization_expiration { get; set; }
		public string base_shipping_hidden_tax_amount { get; set; }
		public string clearSale_status { get; set; }
		public string clearSale_score { get; set; }
		public string clearSale_packageID { get; set; }
		public string clearSale_fingerPrintSessionId { get; set; }
		public string integracommerce_id { get; set; }
		public string bseller_skyhub { get; set; }
		public string bseller_skyhub_code { get; set; }
		public string bseller_skyhub_channel { get; set; }
		public string bseller_skyhub_invoice_key { get; set; }
		public string bseller_skyhub_interest { get; set; }
		public string bseller_skyhub_json { get; set; }
		public MagentoSoapApiShippingAddress shipping_address { get; set; } = new MagentoSoapApiShippingAddress();
		public MagentoSoapApiBillingAddress billing_address { get; set; } = new MagentoSoapApiBillingAddress();
		public List<MagentoSoapApiSalesOrderItem> items { get; set; } = new List<MagentoSoapApiSalesOrderItem>();
		public MagentoSoapApiSalesOrderPayment payment { get; set; } = new MagentoSoapApiSalesOrderPayment();
		public List<MagentoSoapApiStatusHistory> status_history { get; set; } = new List<MagentoSoapApiStatusHistory>();
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();
		public MagentoSoapApiFaultResponse faultResponse { get; set; } = new MagentoSoapApiFaultResponse();

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			int iCounter;
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiSalesOrderInfo).GetProperties())
			{
				#region [ shipping_address ]
				if (prop.Name.Equals("shipping_address"))
				{
					sbResp.AppendLine("");
					sbResp.AppendLine(margem + "shipping_address");
					if (this.shipping_address != null) sbResp.Append(this.shipping_address.FormataDados(margem + "\t"));
					continue;
				}
				#endregion

				#region [ billing_address ]
				if (prop.Name.Equals("billing_address"))
				{
					sbResp.AppendLine("");
					sbResp.AppendLine(margem + "billing_address");
					if (this.billing_address != null) sbResp.Append(this.billing_address.FormataDados(margem + "\t"));
					continue;
				}
				#endregion

				#region [ items ]
				iCounter = 0;
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("items"))
				{
					foreach (MagentoSoapApiSalesOrderItem item in this.items)
					{
						iCounter++;
						if (iCounter > 1) sbAux.AppendLine("");
						sbAux.AppendLine(margem + "item (" + iCounter.ToString() + "/" + this.items.Count.ToString() + ")");
						if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "items [" + this.items.Count.ToString() + "]");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				#region [ payment ]
				if (prop.Name.Equals("payment"))
				{
					sbResp.AppendLine("");
					sbResp.AppendLine(margem + "payment");
					if (this.payment != null) sbResp.Append(this.payment.FormataDados(margem + "\t"));
					continue;
				}
				#endregion

				#region [ status_history ]
				iCounter = 0;
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("status_history"))
				{
					foreach (MagentoSoapApiStatusHistory item in this.status_history)
					{
						iCounter++;
						if (iCounter > 1) sbAux.AppendLine("");
						sbAux.AppendLine(margem + "status_history (" + iCounter.ToString() + "/" + this.status_history.Count.ToString() + ")");
						if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "status_history [" + this.status_history.Count.ToString() + "]");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				#region [ faultResponse ]
				if (prop.Name.Equals("faultResponse"))
				{
					if (this.faultResponse != null)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "faultResponse");
						sbResp.Append(this.faultResponse.FormataDados(margem + "\t"));
					}
					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + prop.GetValue(this, null);
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiSalesOrderItem ]
	public class MagentoSoapApiSalesOrderItem
	{
		public string item_id { get; set; }
		public string order_id { get; set; }
		public string parent_item_id { get; set; }
		public string quote_item_id { get; set; }
		public string store_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string product_id { get; set; }
		public string product_type { get; set; }
		public string product_options { get; set; }
		public string weight { get; set; }
		public string is_virtual { get; set; }
		public string sku { get; set; }
		public string name { get; set; }
		public string description { get; set; }
		public string applied_rule_ids { get; set; }
		public string additional_data { get; set; }
		public string free_shipping { get; set; }
		public string is_qty_decimal { get; set; }
		public string no_discount { get; set; }
		public string qty_backordered { get; set; }
		public string qty_canceled { get; set; }
		public string qty_invoiced { get; set; }
		public string qty_ordered { get; set; }
		public string qty_refunded { get; set; }
		public string qty_shipped { get; set; }
		public string base_cost { get; set; }
		public string price { get; set; }
		public string base_price { get; set; }
		public string original_price { get; set; }
		public string base_original_price { get; set; }
		public string tax_percent { get; set; }
		public string tax_amount { get; set; }
		public string base_tax_amount { get; set; }
		public string tax_invoiced { get; set; }
		public string base_tax_invoiced { get; set; }
		public string discount_percent { get; set; }
		public string discount_amount { get; set; }
		public string base_discount_amount { get; set; }
		public string discount_invoiced { get; set; }
		public string base_discount_invoiced { get; set; }
		public string amount_refunded { get; set; }
		public string base_amount_refunded { get; set; }
		public string row_total { get; set; }
		public string base_row_total { get; set; }
		public string row_invoiced { get; set; }
		public string base_row_invoiced { get; set; }
		public string row_weight { get; set; }
		public string base_tax_before_discount { get; set; }
		public string tax_before_discount { get; set; }
		public string ext_order_item_id { get; set; }
		public string locked_do_invoice { get; set; }
		public string locked_do_ship { get; set; }
		public string price_incl_tax { get; set; }
		public string base_price_incl_tax { get; set; }
		public string row_total_incl_tax { get; set; }
		public string base_row_total_incl_tax { get; set; }
		public string hidden_tax_amount { get; set; }
		public string base_hidden_tax_amount { get; set; }
		public string hidden_tax_invoiced { get; set; }
		public string base_hidden_tax_invoiced { get; set; }
		public string hidden_tax_refunded { get; set; }
		public string base_hidden_tax_refunded { get; set; }
		public string is_nominal { get; set; }
		public string tax_canceled { get; set; }
		public string hidden_tax_canceled { get; set; }
		public string tax_refunded { get; set; }
		public string base_tax_refunded { get; set; }
		public string discount_refunded { get; set; }
		public string base_discount_refunded { get; set; }
		public string gift_message_id { get; set; }
		public string gift_message_available { get; set; }
		public string base_weee_tax_applied_amount { get; set; }
		public string base_weee_tax_applied_row_amnt { get; set; }
		public string base_weee_tax_applied_row_amount { get; set; }
		public string weee_tax_applied_amount { get; set; }
		public string weee_tax_applied_row_amount { get; set; }
		public string weee_tax_applied { get; set; }
		public string weee_tax_disposition { get; set; }
		public string weee_tax_row_disposition { get; set; }
		public string base_weee_tax_disposition { get; set; }
		public string base_weee_tax_row_disposition { get; set; }
		public string installer_document { get; set; }
		public string commission_type { get; set; }
		public string commission_value { get; set; }
		public string has_children { get; set; }
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiSalesOrderItem).GetProperties())
			{
				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiBillingAddress ]
	public class MagentoSoapApiBillingAddress
	{
		public string parent_id { get; set; }
		public string customer_address_id { get; set; }
		public string quote_address_id { get; set; }
		public string region_id { get; set; }
		public string customer_id { get; set; }
		public string fax { get; set; } = "";
		public string region { get; set; }
		public string postcode { get; set; }
		public string firstname { get; set; }
		public string middlename { get; set; }
		public string lastname { get; set; }
		public string street { get; set; }
		public string city { get; set; }
		public string email { get; set; }
		public string telephone { get; set; } = "";
		public string country_id { get; set; }
		public string address_type { get; set; }
		public string prefix { get; set; }
		public string suffix { get; set; }
		public string company { get; set; }
		public string vat_id { get; set; }
		public string vat_is_valid { get; set; }
		public string vat_request_id { get; set; }
		public string vat_request_date { get; set; }
		public string vat_request_success { get; set; }
		public string tipopessoa { get; set; }
		public string rg { get; set; }
		public string ie { get; set; }
		public string cpfcnpj { get; set; }
		public string celular { get; set; } = "";
		public string empresa { get; set; }
		public string nomefantasia { get; set; }
		public string cpf { get; set; }
		public string address_id { get; set; }
		public string street_detail { get; set; }
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiBillingAddress).GetProperties())
			{
				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiShippingAddress ]
	public class MagentoSoapApiShippingAddress
	{
		public string parent_id { get; set; }
		public string customer_address_id { get; set; }
		public string quote_address_id { get; set; }
		public string region_id { get; set; }
		public string customer_id { get; set; }
		public string fax { get; set; } = "";
		public string region { get; set; }
		public string postcode { get; set; }
		public string firstname { get; set; }
		public string middlename { get; set; }
		public string lastname { get; set; }
		public string street { get; set; }
		public string city { get; set; }
		public string email { get; set; }
		public string telephone { get; set; } = "";
		public string country_id { get; set; }
		public string address_type { get; set; }
		public string prefix { get; set; }
		public string suffix { get; set; }
		public string company { get; set; }
		public string vat_id { get; set; }
		public string vat_is_valid { get; set; }
		public string vat_request_id { get; set; }
		public string vat_request_date { get; set; }
		public string vat_request_success { get; set; }
		public string tipopessoa { get; set; }
		public string rg { get; set; }
		public string ie { get; set; }
		public string cpfcnpj { get; set; }
		public string celular { get; set; } = "";
		public string empresa { get; set; }
		public string nomefantasia { get; set; }
		public string cpf { get; set; }
		public string address_id { get; set; }
		public string street_detail { get; set; }
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiShippingAddress).GetProperties())
			{
				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiSalesOrderPayment ]
	public class MagentoSoapApiSalesOrderPayment
	{
		public string parent_id { get; set; }
		public string base_shipping_captured { get; set; }
		public string shipping_captured { get; set; }
		public string amount_refunded { get; set; }
		public string base_amount_paid { get; set; }
		public string amount_canceled { get; set; }
		public string base_amount_authorized { get; set; }
		public string base_amount_paid_online { get; set; }
		public string base_amount_refunded_online { get; set; }
		public string base_shipping_amount { get; set; }
		public string shipping_amount { get; set; }
		public string amount_paid { get; set; }
		public string amount_authorized { get; set; }
		public string base_amount_ordered { get; set; }
		public string base_shipping_refunded { get; set; }
		public string shipping_refunded { get; set; }
		public string base_amount_refunded { get; set; }
		public string amount_ordered { get; set; }
		public string base_amount_canceled { get; set; }
		public string quote_payment_id { get; set; }
		public string additional_data { get; set; }
		public string cc_exp_month { get; set; }
		public string cc_ss_start_year { get; set; }
		public string echeck_bank_name { get; set; }
		public string method { get; set; }
		public string cc_debug_request_body { get; set; }
		public string cc_secure_verify { get; set; }
		public string protection_eligibility { get; set; }
		public string cc_approval { get; set; }
		public string cc_last4 { get; set; }
		public string cc_status_description { get; set; }
		public string echeck_type { get; set; }
		public string cc_debug_response_serialized { get; set; }
		public string cc_ss_start_month { get; set; }
		public string echeck_account_type { get; set; }
		public string last_trans_id { get; set; }
		public string cc_cid_status { get; set; }
		public string cc_owner { get; set; }
		public string cc_type { get; set; }
		public string po_number { get; set; }
		public string cc_exp_year { get; set; }
		public string cc_status { get; set; }
		public string echeck_routing_number { get; set; }
		public string account_status { get; set; }
		public string anet_trans_method { get; set; }
		public string cc_debug_response_body { get; set; }
		public string cc_ss_issue { get; set; }
		public string echeck_account_name { get; set; }
		public string cc_avs_status { get; set; }
		public string cc_number_enc { get; set; }
		public string cc_trans_id { get; set; }
		public string paybox_request_number { get; set; }
		public string address_status { get; set; }
		public string cc_parcelamento { get; set; }
		public string cc_type2 { get; set; }
		public string cc_owner2 { get; set; }
		public string cc_last42 { get; set; }
		public string cc_number_enc2 { get; set; }
		public string cc_exp_month2 { get; set; }
		public string cc_exp_year2 { get; set; }
		public string cc_ss_issue2 { get; set; }
		public string cc_cid2 { get; set; }
		public string cc_parcelamento2 { get; set; }
		public string bseller_payment_in_cash { get; set; }
		public string bseller_payment_installment { get; set; }
		public string payment_id { get; set; }
		public string integracommerce_name { get; set; }
		public string integracommerce_installments { get; set; }
		public MagentoSoapApiSalesOrderPaymentAdditionalInformation additional_information { get; set; } = new MagentoSoapApiSalesOrderPaymentAdditionalInformation();
		public MagentoSoapApiSalesOrderPaymentAdditionalInformation additional_information2 { get; set; } = new MagentoSoapApiSalesOrderPaymentAdditionalInformation();
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiSalesOrderPayment).GetProperties())
			{
				#region [ additional_information ]
				if (prop.Name.Equals("additional_information"))
				{
					sbResp.AppendLine("");
					sbResp.AppendLine(margem + "additional_information");
					if (this.additional_information != null) sbResp.Append(this.additional_information.FormataDados(margem + "\t"));
					continue;
				}
				#endregion

				#region [ additional_information2 ]
				if (prop.Name.Equals("additional_information2"))
				{
					sbResp.AppendLine("");
					sbResp.AppendLine(margem + "additional_information2");
					if (this.additional_information2 != null) sbResp.Append(this.additional_information2.FormataDados(margem + "\t"));
					continue;
				}
				#endregion

				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiSalesOrderPaymentAdditionalInformation ]
	public class MagentoSoapApiSalesOrderPaymentAdditionalInformation
	{
		public string PaymentMethod { get; set; }
		public string InstallmentsCount { get; set; }
		public string BraspagOrderId { get; set; }
		public string ErrorDescription { get; set; }
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiSalesOrderPaymentAdditionalInformation).GetProperties())
			{
				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiStatusHistory ]
	public class MagentoSoapApiStatusHistory
	{
		public string parent_id { get; set; }
		public string is_customer_notified { get; set; }
		public string is_visible_on_front { get; set; }
		public string comment { get; set; }
		public string status { get; set; }
		public string created_at { get; set; }
		public string entity_name { get; set; }
		public string store_id { get; set; }
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiStatusHistory).GetProperties())
			{
				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ MagentoSoapApiFaultResponse ]
	public class MagentoSoapApiFaultResponse
	{
		public bool isFaultResponse { get; set; } = false;
		public string faultcode { get; set; }
		public string faultstring { get; set; }
		public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();

		#region [ FormataDados ]
		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(MagentoSoapApiFaultResponse).GetProperties())
			{
				#region [ UnknownFields ]
				sbAux = new StringBuilder("");
				if (prop.Name.Equals("UnknownFields"))
				{
					foreach (KeyValuePair<string, string> item in this.UnknownFields)
					{
						linha = margem + "\t" + item.Key + " = " + item.Value;
						sbAux.AppendLine(linha);
					}
					if (sbAux.Length > 0)
					{
						sbResp.AppendLine("");
						sbResp.AppendLine(margem + "UnknownFields (" + this.UnknownFields.Count.ToString() + ")");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + prop.GetValue(this, null));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion
}