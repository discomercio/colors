using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ConsolidadorXlsEC
{
    public class BillingAddress
    {
        public string parent_id { get; set; }
        public string customer_address_id { get; set; }
        public string quote_address_id { get; set; }
        public string region_id { get; set; }
        public string customer_id { get; set; }
        public string fax { get; set; }
        public string region { get; set; }
        public string postcode { get; set; }
        public string firstname { get; set; }
        public string middlename { get; set; }
        public string lastname { get; set; }
        public string street { get; set; }
        public string city { get; set; }
        public string email { get; set; }
        public string telephone { get; set; }
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
        public string celular { get; set; }
        public string empresa { get; set; }
        public string nomefantasia { get; set; }
        public string cpf { get; set; }
        public string address_id { get; set; }
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
            foreach (PropertyInfo prop in typeof(BillingAddress).GetProperties())
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
