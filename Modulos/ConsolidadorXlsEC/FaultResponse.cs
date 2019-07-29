using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ConsolidadorXlsEC
{
    public class FaultResponse
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
            foreach (PropertyInfo prop in typeof(FaultResponse).GetProperties())
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
