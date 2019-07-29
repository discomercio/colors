using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ConsolidadorXlsEC
{
    public class SalesOrderPayment
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
        public SalesOrderPaymentAdditionalInformation additional_information { get; set; } = new SalesOrderPaymentAdditionalInformation();
        public SalesOrderPaymentAdditionalInformation additional_information2 { get; set; } = new SalesOrderPaymentAdditionalInformation();
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
            foreach (PropertyInfo prop in typeof(SalesOrderPayment).GetProperties())
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

    public class SalesOrderPaymentAdditionalInformation
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
            foreach (PropertyInfo prop in typeof(SalesOrderPaymentAdditionalInformation).GetProperties())
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
