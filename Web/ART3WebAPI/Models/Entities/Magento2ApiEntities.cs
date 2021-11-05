using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;
using ART3WebAPI.Models.Domains;

namespace ART3WebAPI.Models.Entities
{
	#region [ Magento2SalesOrderInfo ]
	public class Magento2SalesOrderInfo
	{
		public string increment_id { get; set; }
		public string entity_id { get; set; } // Campo novo do Magento 2
		public string store_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string status { get; set; }
		public string state { get; set; }
		public string base_currency_code { get; set; }
		public string base_discount_amount { get; set; }
		public string base_grand_total { get; set; }
		public string base_discount_tax_compensation_amount { get; set; } // Campo novo do Magento 2
		public string base_shipping_amount { get; set; }
		public string base_shipping_discount_amount { get; set; }
		public string base_shipping_discount_tax_compensation_amnt { get; set; } // Campo novo do Magento 2
		public string base_shipping_incl_tax { get; set; }
		public string base_shipping_tax_amount { get; set; }
		public string base_subtotal { get; set; }
		public string base_subtotal_incl_tax { get; set; }
		public string base_tax_amount { get; set; }
		public string base_total_due { get; set; }
		public string base_to_global_rate { get; set; }
		public string base_to_order_rate { get; set; }
		public string billing_address_id { get; set; }
		public string customer_email { get; set; }
		public string customer_firstname { get; set; }
		public string customer_lastname { get; set; }
		public string customer_group_id { get; set; }
		public string customer_id { get; set; }
		public string customer_is_guest { get; set; }
		public string customer_note_notify { get; set; }
		public string customer_taxvat { get; set; }
		public string discount_amount { get; set; }
		public string email_sent { get; set; }
		public string global_currency_code { get; set; }
		public string grand_total { get; set; }
		public string discount_tax_compensation_amount { get; set; } // Campo novo do Magento 2
		public string is_virtual { get; set; }
		public string order_currency_code { get; set; }
		public string protect_code { get; set; }
		public string quote_id { get; set; }
		public string remote_ip { get; set; }
		public string shipping_amount { get; set; }
		public string shipping_description { get; set; }
		public string shipping_discount_amount { get; set; }
		public string shipping_discount_tax_compensation_amount { get; set; } // Campo novo do Magento 2
		public string shipping_incl_tax { get; set; }
		public string shipping_tax_amount { get; set; }
		public string store_currency_code { get; set; }
		public string store_name { get; set; }
		public string store_to_base_rate { get; set; }
		public string store_to_order_rate { get; set; }
		public string subtotal { get; set; }
		public string subtotal_incl_tax { get; set; }
		public string tax_amount { get; set; }
		public string total_due { get; set; }
		public string total_item_count { get; set; }
		public string total_qty_ordered { get; set; }
		public string weight { get; set; }
		public string x_forwarded_for { get; set; }
		public List<Magento2SalesOrderItem> items { get; set; }
		public Magento2BillingAddress billing_address { get; set; } = new Magento2BillingAddress();
		public Magento2SalesOrderPayment payment { get; set; } = new Magento2SalesOrderPayment();
		public List<Magento2StatusHistory> status_histories { get; set; } = new List<Magento2StatusHistory>();
		public Magento2ExtensionAttributes extension_attributes { get; set; } = new Magento2ExtensionAttributes();

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
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderInfo).GetProperties())
			{
				#region [ items ]
				if (prop.Name.Equals("items"))
				{
					iCounter = 0;
					sbAux = new StringBuilder("");
					foreach (Magento2SalesOrderItem item in this.items)
					{
						iCounter++;
						// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
						sbAux.AppendLine(margem + "item (" + iCounter.ToString() + "/" + this.items.Count.ToString() + ")");
						if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
					}

					if (sbAux.Length > 0)
					{
						// TODO-SEP-ARRAY						sbResp.AppendLine("");
						// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "items [" + this.items.Count.ToString() + "]");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				#region [ billing_address ]
				if (prop.Name.Equals("billing_address"))
				{
					if (this.billing_address != null)
					{
						sbResp.AppendLine(margem + "billing_address");
						sbResp.Append(this.billing_address.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "billing_address = null");
					}
					continue;
				}
				#endregion

				#region [ payment ]
				if (prop.Name.Equals("payment"))
				{
					if (this.payment != null)
					{
						sbResp.AppendLine(margem + "payment");
						sbResp.Append(this.payment.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "payment = null");
					}
					continue;
				}
				#endregion

				#region [ status_histories ]
				if (prop.Name.Equals("status_histories"))
				{
					if (this.status_histories.Count == 0)
					{
						sbResp.AppendLine(margem + "status_histories = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2StatusHistory item in this.status_histories)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "status_histories (" + iCounter.ToString() + "/" + this.status_histories.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "status_histories [" + this.status_histories.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ extension_attributes ]
				if (prop.Name.Equals("extension_attributes"))
				{
					if (this.extension_attributes != null)
					{
						sbResp.AppendLine(margem + "extension_attributes");
						sbResp.Append(this.extension_attributes.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "extension_attributes = null");
					}
					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderItem ]
	public class Magento2SalesOrderItem
	{
		public string sku { get; set; }
		public string item_id { get; set; }
		public string order_id { get; set; }
		public string parent_item_id { get; set; }
		public string quote_item_id { get; set; }
		public string store_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string product_id { get; set; }
		public string amount_refunded { get; set; }
		public string base_amount_refunded { get; set; }
		public string base_discount_amount { get; set; }
		public string base_discount_invoiced { get; set; }
		public string base_discount_tax_compensation_amount { get; set; } // Campo novo do Magento 2
		public string base_original_price { get; set; }
		public string base_price { get; set; }
		public string base_price_incl_tax { get; set; }
		public string base_row_invoiced { get; set; }
		public string base_row_total { get; set; }
		public string base_row_total_incl_tax { get; set; }
		public string base_tax_amount { get; set; }
		public string base_tax_invoiced { get; set; }
		public string discount_amount { get; set; }
		public string discount_invoiced { get; set; }
		public string discount_percent { get; set; }
		public string free_shipping { get; set; }
		public string discount_tax_compensation_amount { get; set; } // Campo novo do Magento 2
		public string is_qty_decimal { get; set; }
		public string is_virtual { get; set; }
		public string name { get; set; }
		public string no_discount { get; set; }
		public string original_price { get; set; }
		public string price { get; set; }
		public string price_incl_tax { get; set; }
		public string product_type { get; set; }
		public string qty_canceled { get; set; }
		public string qty_invoiced { get; set; }
		public string qty_ordered { get; set; }
		public string qty_refunded { get; set; }
		public string qty_returned { get; set; } // Campo novo do Magento 2
		public string qty_shipped { get; set; }
		public string row_invoiced { get; set; }
		public string row_total { get; set; }
		public string row_total_incl_tax { get; set; }
		public string row_weight { get; set; }
		public string tax_amount { get; set; }
		public string tax_invoiced { get; set; }
		public string tax_percent { get; set; }
		public string weight { get; set; }
		public Magento2SalesOrderItemProductOption product_option { get; set; }
		public Magento2SalesOrderItemParentItem parent_item { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderItem).GetProperties())
			{
				#region [ product_option ]
				if (prop.Name.Equals("product_option"))
				{
					if (this.product_option != null)
					{
						sbResp.AppendLine(margem + "product_option");
						sbResp.Append(this.product_option.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "product_option = null");
					}
					continue;
				}
				#endregion

				#region [ parent_item ]
				if (prop.Name.Equals("parent_item"))
				{
					if (this.parent_item != null)
					{
						sbResp.AppendLine(margem + "parent_item");
						sbResp.Append(this.parent_item.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "parent_item = null");
					}
					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString()));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderItemParentItem ]
	public class Magento2SalesOrderItemParentItem
	{
		public string sku { get; set; }
		public string item_id { get; set; }
		public string product_id { get; set; }
		public string order_id { get; set; }
		public string store_id { get; set; }
		public string quote_item_id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
		public string amount_refunded { get; set; }
		public string base_amount_refunded { get; set; }
		public string base_discount_amount { get; set; }
		public string base_discount_invoiced { get; set; }
		public string base_discount_tax_compensation_amount { get; set; }
		public string base_original_price { get; set; }
		public string base_price { get; set; }
		public string base_price_incl_tax { get; set; }
		public string base_row_invoiced { get; set; }
		public string base_row_total { get; set; }
		public string base_row_total_incl_tax { get; set; }
		public string base_tax_amount { get; set; }
		public string base_tax_invoiced { get; set; }
		public string discount_amount { get; set; }
		public string discount_invoiced { get; set; }
		public string discount_percent { get; set; }
		public string free_shipping { get; set; }
		public string discount_tax_compensation_amount { get; set; }
		public string is_qty_decimal { get; set; }
		public string is_virtual { get; set; }
		public string name { get; set; }
		public string no_discount { get; set; }
		public string original_price { get; set; }
		public string price { get; set; }
		public string price_incl_tax { get; set; }
		public string product_type { get; set; }
		public string qty_canceled { get; set; }
		public string qty_invoiced { get; set; }
		public string qty_ordered { get; set; }
		public string qty_refunded { get; set; }
		public string qty_returned { get; set; }
		public string qty_shipped { get; set; }
		public string row_invoiced { get; set; }
		public string row_total { get; set; }
		public string row_total_incl_tax { get; set; }
		public string row_weight { get; set; }
		public string tax_amount { get; set; }
		public string tax_invoiced { get; set; }
		public string tax_percent { get; set; }
		public string weight { get; set; }
		public Magento2SalesOrderItemProductOption product_option { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderItemParentItem).GetProperties())
			{
				#region [ product_option ]
				if (prop.Name.Equals("product_option"))
				{
					if (this.product_option != null)
					{
						sbResp.AppendLine(margem + "product_option");
						sbResp.Append(this.product_option.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "product_option = null");
					}
					continue;
				}
				#endregion

				sbResp.AppendLine(margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString()));
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderItemProductOption ]
	public class Magento2SalesOrderItemProductOption
	{
		public Magento2SalesOrderItemProductOptionExtensionAttributes extension_attributes { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderItemProductOption).GetProperties())
			{
				#region [ extension_attributes ]
				if (prop.Name.Equals("extension_attributes"))
				{
					sbResp.AppendLine(margem + "extension_attributes");
					if (this.extension_attributes != null) sbResp.Append(this.extension_attributes.FormataDados(margem + "\t"));
					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderItemProductOptionExtensionAttributes ]
	public class Magento2SalesOrderItemProductOptionExtensionAttributes
	{
		public List<Magento2SalesOrderItemProductOptionExtensionAttributesConfigurableItemOptions> configurable_item_options { get; set; }

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
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderItemProductOptionExtensionAttributes).GetProperties())
			{
				#region [ configurable_item_options ]
				if (prop.Name.Equals("configurable_item_options"))
				{
					iCounter = 0;
					sbAux = new StringBuilder("");
					foreach (Magento2SalesOrderItemProductOptionExtensionAttributesConfigurableItemOptions option in this.configurable_item_options)
					{
						iCounter++;
						// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
						sbAux.AppendLine(margem + "configurable_item_options (" + iCounter.ToString() + "/" + this.configurable_item_options.Count.ToString() + ")");
						if (option != null) sbAux.Append(option.FormataDados(margem + "\t"));
					}

					if (sbAux.Length > 0)
					{
						// TODO-SEP-ARRAY						sbResp.AppendLine("");
						// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "configurable item options [" + this.configurable_item_options.Count.ToString() + "]");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderItemProductOptionExtensionAttributesConfigurableItemOptions ]
	public class Magento2SalesOrderItemProductOptionExtensionAttributesConfigurableItemOptions
	{
		public string option_id { get; set; }
		public string option_value { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderItemProductOptionExtensionAttributesConfigurableItemOptions).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2BillingAddress ]
	public class Magento2BillingAddress
	{
		public string entity_id { get; set; }
		public string customer_address_id { get; set; }
		public string parent_id { get; set; }
		public string address_type { get; set; }
		public string city { get; set; }
		public string country_id { get; set; }
		public string email { get; set; }
		public string firstname { get; set; }
		public string lastname { get; set; }
		public string postcode { get; set; }
		public string region { get; set; }
		public string region_code { get; set; }
		public string region_id { get; set; }
		public List<string> street { get; set; }
		public string telephone { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbAux;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2BillingAddress).GetProperties())
			{
				#region [ street ]
				if (prop.Name.Equals("street"))
				{
					if (this.street == null)
					{
						sbResp.AppendLine(margem + "street = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.street)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "street");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderPayment ]
	public class Magento2SalesOrderPayment
	{
		public string entity_id { get; set; }
		public string parent_id { get; set; }
		public string account_status { get; set; }
		public string amount_ordered { get; set; }
		public string base_amount_ordered { get; set; }
		public string base_amount_authorized { get; set; }
		public string base_shipping_amount { get; set; }
		public string amount_authorized { get; set; }
		public string cc_last4 { get; set; }
		public string cc_ss_start_month { get; set; }
		public string cc_ss_start_year { get; set; }
		public string method { get; set; }
		public string shipping_amount { get; set; }
		public string last_trans_id { get; set; }
		public List<string> additional_information { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbAux;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2SalesOrderPayment).GetProperties())
			{
				#region [ additional_information ]
				if (prop.Name.Equals("additional_information"))
				{
					if (this.additional_information == null)
					{
						sbResp.AppendLine(margem + "additional_information = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.additional_information)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "additional_information");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributes ]
	public class Magento2ExtensionAttributes
	{
		public List<Magento2ExtensionAttributesShippingAssignments> shipping_assignments { get; set; }
		public List<Magento2ExtensionAttributesPaymentAdditionalInfo> payment_additional_info { get; set; }
		public List<Magento2ExtensionAttributesGiftCards> gift_cards { get; set; }
		public string base_gift_cards_amount { get; set; }
		public string gift_cards_amount { get; set; }
		public List<Magento2ExtensionAttributesAppliedTaxes> applied_taxes { get; set; }
		public List<Magento2ExtensionAttributesItemAppliedTaxes> item_applied_taxes { get; set; }
		public string gw_base_price { get; set; }
		public string gw_price { get; set; }
		public string gw_items_base_price { get; set; }
		public string gw_items_price { get; set; }
		public string gw_card_base_price { get; set; }
		public string gw_card_price { get; set; }
		public Magento2ExtensionAttributesSkyhubInfo skyhub_info { get; set; }

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
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributes).GetProperties())
			{
				#region [ shipping_assignments ]
				if (prop.Name.Equals("shipping_assignments"))
				{
					if (this.shipping_assignments.Count == 0)
					{
						sbResp.AppendLine(margem + "shipping_assignments = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ExtensionAttributesShippingAssignments item in this.shipping_assignments)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "shipping_assignments (" + iCounter.ToString() + "/" + this.shipping_assignments.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "shipping_assignments [" + this.shipping_assignments.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ payment_additional_info ]
				if (prop.Name.Equals("payment_additional_info"))
				{
					if (this.payment_additional_info.Count == 0)
					{
						sbResp.AppendLine(margem + "payment_additional_info = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ExtensionAttributesPaymentAdditionalInfo item in this.payment_additional_info)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "payment_additional_info (" + iCounter.ToString() + "/" + this.payment_additional_info.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "payment_additional_info [" + this.payment_additional_info.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ gift_cards ]
				if (prop.Name.Equals("gift_cards"))
				{
					if (this.gift_cards.Count == 0)
					{
						sbResp.AppendLine(margem + "gift_cards = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ExtensionAttributesGiftCards item in this.gift_cards)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "gift_cards (" + iCounter.ToString() + "/" + this.gift_cards.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "gift_cards [" + this.gift_cards.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ applied_taxes ]
				if (prop.Name.Equals("applied_taxes"))
				{
					if (this.applied_taxes.Count == 0)
					{
						sbResp.AppendLine(margem + "applied_taxes = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ExtensionAttributesAppliedTaxes item in this.applied_taxes)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "applied_taxes (" + iCounter.ToString() + "/" + this.applied_taxes.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "applied_taxes [" + this.applied_taxes.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ item_applied_taxes ]
				if (prop.Name.Equals("item_applied_taxes"))
				{
					if (this.item_applied_taxes.Count == 0)
					{
						sbResp.AppendLine(margem + "item_applied_taxes = null");
					}
					else
					{
						iCounter = 0;
						sbAux = new StringBuilder("");
						foreach (Magento2ExtensionAttributesItemAppliedTaxes item in this.item_applied_taxes)
						{
							iCounter++;
							// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
							sbAux.AppendLine(margem + "item_applied_taxes (" + iCounter.ToString() + "/" + this.item_applied_taxes.Count.ToString() + ")");
							if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
						}

						if (sbAux.Length > 0)
						{
							// TODO-SEP-ARRAY						sbResp.AppendLine("");
							// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "item_applied_taxes [" + this.item_applied_taxes.Count.ToString() + "]");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				#region [ skyhub_info ]
				if (prop.Name.Equals("skyhub_info"))
				{
					if (this.skyhub_info != null)
					{
						sbResp.AppendLine(margem + "skyhub_info");
						sbResp.Append(this.skyhub_info.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "skyhub_info = null");
					}
					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesShippingAssignments ]
	public class Magento2ExtensionAttributesShippingAssignments
	{
		public Magento2ExtensionAttributesShippingAssignmentsShipping shipping { get; set; }
		public List<Magento2SalesOrderItem> items { get; set; }

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
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesShippingAssignments).GetProperties())
			{
				#region [ shipping ]
				if (prop.Name.Equals("shipping"))
				{
					if (this.shipping != null)
					{
						sbResp.AppendLine(margem + "shipping");
						sbResp.Append(this.shipping.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "shipping = null");
					}
					continue;
				}
				#endregion

				#region [ items ]
				if (prop.Name.Equals("items"))
				{
					iCounter = 0;
					sbAux = new StringBuilder("");
					foreach (Magento2SalesOrderItem item in this.items)
					{
						iCounter++;
						// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
						sbAux.AppendLine(margem + "items (" + iCounter.ToString() + "/" + this.items.Count.ToString() + ")");
						if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
					}

					if (sbAux.Length > 0)
					{
						// TODO-SEP-ARRAY						sbResp.AppendLine("");
						// TODO-SEP-ARRAY						sbResp.AppendLine(margem + "items [" + this.items.Count.ToString() + "]");
						sbResp.Append(sbAux.ToString());
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesShippingAssignmentsShipping ]
	public class Magento2ExtensionAttributesShippingAssignmentsShipping
	{
		public Magento2ExtensionAttributesShippingAssignmentsShippingAddress address { get; set; }
		public string method { get; set; }
		public Magento2ExtensionAttributesShippingAssignmentsShippingTotal total { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesShippingAssignmentsShipping).GetProperties())
			{
				#region [ address ]
				if (prop.Name.Equals("address"))
				{
					if (this.address != null)
					{
						sbResp.AppendLine(margem + "address");
						sbResp.Append(this.address.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "address = null");
					}
					continue;
				}
				#endregion

				#region [ total ]
				if (prop.Name.Equals("total"))
				{
					if (this.total != null)
					{
						sbResp.AppendLine(margem + "total");
						sbResp.Append(this.total.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "total = null");
					}
					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesShippingAssignmentsShippingAddress ]
	public class Magento2ExtensionAttributesShippingAssignmentsShippingAddress
	{
		public string entity_id { get; set; }
		public string customer_address_id { get; set; }
		public string address_type { get; set; }
		public string parent_id { get; set; }
		public string email { get; set; }
		public string firstname { get; set; }
		public string lastname { get; set; }
		public List<string> street { get; set; }
		public string city { get; set; }
		public string region { get; set; }
		public string region_code { get; set; }
		public string region_id { get; set; }
		public string postcode { get; set; }
		public string country_id { get; set; }
		public string telephone { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			StringBuilder sbAux;
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesShippingAssignmentsShippingAddress).GetProperties())
			{
				#region [ street ]
				if (prop.Name.Equals("street"))
				{
					if (this.street == null)
					{
						sbResp.AppendLine(margem + "street = null");
					}
					else
					{
						sbAux = new StringBuilder("");
						foreach (string linhaTexto in this.street)
						{
							linha = margem + "\t" + linhaTexto;
							sbAux.AppendLine(linha);
						}

						if (sbAux.Length > 0)
						{
							sbResp.AppendLine(margem + "street");
							sbResp.Append(sbAux.ToString());
						}
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesShippingAssignmentsShippingTotal ]
	public class Magento2ExtensionAttributesShippingAssignmentsShippingTotal
	{
		public string base_shipping_amount { get; set; }
		public string base_shipping_discount_amount { get; set; }
		public string base_shipping_discount_tax_compensation_amnt { get; set; }
		public string base_shipping_incl_tax { get; set; }
		public string base_shipping_tax_amount { get; set; }
		public string shipping_amount { get; set; }
		public string shipping_discount_amount { get; set; }
		public string shipping_discount_tax_compensation_amount { get; set; }
		public string shipping_incl_tax { get; set; }
		public string shipping_tax_amount { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesShippingAssignmentsShippingTotal).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesPaymentAdditionalInfo ]
	public class Magento2ExtensionAttributesPaymentAdditionalInfo
	{
		public string key { get; set; }
		public string value { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesPaymentAdditionalInfo).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesGiftCards ]
	public class Magento2ExtensionAttributesGiftCards
	{
		public string id { get; set; }
		public string code { get; set; }
		public string amount { get; set; }
		public string base_amount { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesGiftCards).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesAppliedTaxes ]
	public class Magento2ExtensionAttributesAppliedTaxes
	{
		public string code { get; set; }
		public string title { get; set; }
		public string percent { get; set; }
		public string amount { get; set; }
		public string base_amount { get; set; }
		public Magento2ExtensionAttributesAppliedTaxesExtensionAttributes extension_attributes { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesAppliedTaxes).GetProperties())
			{
				if (prop.Name.Equals("extension_attributes"))
				{
					if (this.extension_attributes != null)
					{
						sbResp.AppendLine(margem + "extension_attributes");
						sbResp.Append(this.extension_attributes.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "extension_attributes = null");
					}
					continue;
				}

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesAppliedTaxesExtensionAttributes ]
	public class Magento2ExtensionAttributesAppliedTaxesExtensionAttributes
	{
		public List<Magento2ExtensionAttributesAppliedTaxesExtensionAttributesRates> rates { get; set; }

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
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesItemAppliedTaxes).GetProperties())
			{
				#region [ rates ]
				if (prop.Name.Equals("rates"))
				{
					iCounter = 0;
					sbAux = new StringBuilder("");
					foreach (Magento2ExtensionAttributesAppliedTaxesExtensionAttributesRates item in this.rates)
					{
						iCounter++;
						// TODO-SEP-ARRAY						if (iCounter > 1) sbAux.AppendLine("");
						sbAux.AppendLine(margem + "rates (" + iCounter.ToString() + "/" + this.rates.Count.ToString() + ")");
						if (item != null) sbAux.Append(item.FormataDados(margem + "\t"));
					}

					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesAppliedTaxesExtensionAttributesRates ]
	public class Magento2ExtensionAttributesAppliedTaxesExtensionAttributesRates
	{
		public string code { get; set; }
		public string title { get; set; }
		public string percent { get; set; }
		// extension_attributes: campo ignorado por ter estrutura desconhecida

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesItemAppliedTaxes).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesItemAppliedTaxes ]
	public class Magento2ExtensionAttributesItemAppliedTaxes
	{
		public string type { get; set; }
		public string item_id { get; set; }
		public string associated_item_id { get; set; }
		public List<Magento2ExtensionAttributesAppliedTaxes> applied_taxes { get; set; }
		// extension_attributes: campo ignorado por ter estrutura desconhecida

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesItemAppliedTaxes).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesSkyhubInfo ]
	public class Magento2ExtensionAttributesSkyhubInfo
	{
		public string id { get; set; }
		public string store_id { get; set; }
		public Magento2ExtensionAttributesSkyhubInfoStore store { get; set; }
		public string order_id { get; set; }
		public string code { get; set; }
		public string channel { get; set; }
		public string invoice_key { get; set; }
		public string data_source { get; set; }
		public string interest { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesSkyhubInfo).GetProperties())
			{
				#region [ store ]
				if (prop.Name.Equals("store"))
				{
					if (this.store != null)
					{
						sbResp.AppendLine(margem + "store");
						sbResp.Append(this.store.FormataDados(margem + "\t"));
					}
					else
					{
						sbResp.AppendLine(margem + "store = null");
					}
					continue;
				}
				#endregion

				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2ExtensionAttributesSkyhubInfoStore ]
	public class Magento2ExtensionAttributesSkyhubInfoStore
	{
		public string id { get; set; }
		public string code { get; set; }
		public string name { get; set; }
		public string website_id { get; set; }
		public string store_group_id { get; set; }
		public string is_active { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2ExtensionAttributesSkyhubInfoStore).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2StatusHistory ]
	public class Magento2StatusHistory
	{
		public string entity_id { get; set; }
		public string parent_id { get; set; }
		public string is_customer_notified { get; set; }
		public string is_visible_on_front { get; set; }
		public string created_at { get; set; }
		public string entity_name { get; set; }
		public string status { get; set; }
		public string comment { get; set; }

		#region [ FormataDados ]
		public string FormataDados()
		{
			return FormataDados("");
		}

		public string FormataDados(string margem)
		{
			#region [ Declarações ]
			string linha;
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			margem = (margem ?? "");

			#region [ Formata os dados das propriedades ]
			foreach (PropertyInfo prop in typeof(Magento2StatusHistory).GetProperties())
			{
				linha = margem + prop.Name + " = " + Magento2RestApi.formataDadosCampoComQuebraLinha(margem, prop.Name, (prop.GetValue(this, null) ?? "").ToString());
				sbResp.AppendLine(linha);
			}
			#endregion

			return sbResp.ToString();
		}
		#endregion
	}
	#endregion

	#region [ Magento2SalesOrderSearchResponse ]
	public class Magento2SalesOrderSearchResponse
	{
		public List<Magento2SalesOrderInfo> items { get; set; }
		public Magento2SearchCriteria search_criteria { get; set; }
		public int total_count { get; set; }
	}
	#endregion

	#region [ Magento2SearchCriteria ]
	public class Magento2SearchCriteria
	{
		public List<Magento2SearchCriteriaFilterGroups> filter_groups { get; set; }
	}
	#endregion

	#region [ Magento2SearchCriteriaFilterGroups ]
	public class Magento2SearchCriteriaFilterGroups
	{
		public List<Magento2SearchCriteriaFilterGroupsFilters> filters { get; set; }
	}
	#endregion

	#region [ Magento2SearchCriteriaFilterGroupsFilters ]
	public class Magento2SearchCriteriaFilterGroupsFilters
	{
		public string field { get; set; }
		public string value { get; set; }
		public string condition_type { get; set; }
	}
	#endregion

	#region [ Magento2 Skyhub (DataSource/JSON) ]

	#region [ Magento2SkyhubDataSource ]
	public class Magento2SkyhubDataSource
	{
		public string code { get; set; }
		public string delivery_contract_type { get; set; }
		public string channel { get; set; }
		public string shipping_cost { get; set; }
		public string shipping_method { get; set; }
		public string calculation_type { get; set; }
		public string shipping_carrier { get; set; }
		public string estimated_delivery { get; set; }
		public string updated_at { get; set; }
		public string placed_at { get; set; }
		public string available_to_sync { get; set; }
		public string expedition_limit_date { get; set; }
		public string shipping_method_id { get; set; }
		public string sync_status { get; set; }
		public string discount { get; set; }
		public string total_ordered { get; set; }
		public string imported_at { get; set; }
		public string approved_date { get; set; }
		public Magento2SkyhubDataSourceBillingAddress billing_address { get; set; } = new Magento2SkyhubDataSourceBillingAddress();
		public Magento2SkyhubDataSourceShippingAddress shipping_address { get; set; } = new Magento2SkyhubDataSourceShippingAddress();
		public Magento2SkyhubDataSourceImportInfo import_info { get; set; } = new Magento2SkyhubDataSourceImportInfo();
		public Magento2SkyhubDataSourceStatus status { get; set; } = new Magento2SkyhubDataSourceStatus();
		public List<Magento2SkyhubDataSourceItem> items { get; set; }
		public List<Magento2SkyhubDataSourcePayment> payments { get; set; }
		public Magento2SkyhubDataSourceCustomer customer { get; set; } = new Magento2SkyhubDataSourceCustomer();
	}
	#endregion

	#region [ Magento2SkyhubDataSourceBillingAddress ]
	public class Magento2SkyhubDataSourceBillingAddress
	{
		public string street { get; set; }
		public string secondary_phone { get; set; }
		public string region { get; set; }
		public string reference { get; set; }
		public string postcode { get; set; }
		public string phone { get; set; }
		public string number { get; set; }
		public string neighborhood { get; set; }
		public string full_name { get; set; }
		public string detail { get; set; }
		public string country { get; set; }
		public string complement { get; set; }
		public string city { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourceShippingAddress ]
	public class Magento2SkyhubDataSourceShippingAddress
	{
		public string street { get; set; }
		public string secondary_phone { get; set; }
		public string region { get; set; }
		public string reference { get; set; }
		public string postcode { get; set; }
		public string phone { get; set; }
		public string number { get; set; }
		public string neighborhood { get; set; }
		public string full_name { get; set; }
		public string detail { get; set; }
		public string country { get; set; }
		public string complement { get; set; }
		public string city { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourceImportInfo ]
	public class Magento2SkyhubDataSourceImportInfo
	{
		public string ss_name { get; set; }
		public string remote_id { get; set; }
		public string remote_code { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourceStatus ]
	public class Magento2SkyhubDataSourceStatus
	{
		public string type { get; set; }
		public string label { get; set; }
		public string code { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourceItem ]
	public class Magento2SkyhubDataSourceItem
	{
		public string special_price { get; set; }
		public string shipping_cost { get; set; }
		public string sale_fee { get; set; }
		public string remote_store_id { get; set; }
		public string qty { get; set; }
		public string product_id { get; set; }
		public string original_price { get; set; }
		public string name { get; set; }
		public string listing_type_id { get; set; }
		public string id { get; set; }
		public string gift_wrap { get; set; }
		public string detail { get; set; }
		public string delivery_line_id { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourcePayment ]
	public class Magento2SkyhubDataSourcePayment
	{
		public string value { get; set; }
		public string type { get; set; }
		public string transaction_date { get; set; }
		public string status { get; set; }
		public string parcels { get; set; }
		public string method { get; set; }
		public string description { get; set; }
		public string card_issuer { get; set; }
		public string autorization_id { get; set; }
		public Magento2SkyhubDataSourcePaymentSefaz sefaz { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourcePaymentSefaz ]
	public class Magento2SkyhubDataSourcePaymentSefaz
	{
		public string type_integration { get; set; }
		public string payment_indicator { get; set; }
		public string name_payment { get; set; }
		public string name_card_issuer { get; set; }
		public string id_payment { get; set; }
		public string id_card_issuer { get; set; }
	}
	#endregion

	#region [ Magento2SkyhubDataSourceCustomer ]
	public class Magento2SkyhubDataSourceCustomer
	{
		public string vat_number { get; set; }
		public List<string> phones { get; set; }
		public string name { get; set; }
		public string gender { get; set; }
		public string email { get; set; }
		public string date_of_birth { get; set; }
	}
	#endregion

	#endregion
}