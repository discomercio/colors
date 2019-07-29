using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ConsolidadorXlsEC
{
    public class SalesOrderInfo
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
        public ShippingAddress shipping_address { get; set; } = new ShippingAddress();
        public BillingAddress billing_address { get; set; } = new BillingAddress();
        public List<SalesOrderItem> items { get; set; } = new List<SalesOrderItem>();
        public SalesOrderPayment payment { get; set; } = new SalesOrderPayment();
        public List<StatusHistory> status_history { get; set; } = new List<StatusHistory>();
        public List<KeyValuePair<string, string>> UnknownFields { get; set; } = new List<KeyValuePair<string, string>>();
        public FaultResponse faultResponse { get; set; } = new FaultResponse();

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
            foreach (PropertyInfo prop in typeof(SalesOrderInfo).GetProperties())
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
                    foreach (SalesOrderItem item in this.items)
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
                    foreach (StatusHistory item in this.status_history)
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

    public class SalesOrderItem
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
            foreach (PropertyInfo prop in typeof(SalesOrderItem).GetProperties())
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
}
