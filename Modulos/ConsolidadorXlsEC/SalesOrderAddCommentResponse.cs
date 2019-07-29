using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ConsolidadorXlsEC
{
    public class SalesOrderAddCommentResponse
    {
        public string callReturn { get; set; }
        public FaultResponse faultResponse { get; set; } = new FaultResponse();
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
            foreach (PropertyInfo prop in typeof(SalesOrderAddCommentResponse).GetProperties())
            {
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
