using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;

namespace ART3WebAPI.Models.Domains
{
    public class CadIndicadoresGeradorRelatorio
    {
        public static Task GerarListagemCSV(List<Indicador> datasource, string filePath)
        {
            return Task.Run(() =>
            {
                Encoding encode = Encoding.GetEncoding("Windows-1252");
                using (StreamWriter sw = new StreamWriter(new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Read), encode))
                {

                    string delimitador = ";";
                    int length = datasource.Count;
                    StringBuilder sb = new StringBuilder();

                    //cria cabeçalho
                    sb.Append("Nome" + delimitador);
                    sb.Append("E-mail" + delimitador);
                    sb.Append("UF" + delimitador);
                    sb.Append("Vendedor" + delimitador);

                    sw.WriteLine(sb);

                    for (int i = 0; i < length; i++)
                    {
                        sw.WriteLine(datasource.ElementAt(i).Nome.ToUpper().Replace(";",",") + delimitador +
                            datasource.ElementAt(i).Email + delimitador +    
                            datasource.ElementAt(i).Uf + delimitador +
                            datasource.ElementAt(i).Vendedor + delimitador +
                            "" + delimitador);                          
                    }

                }
            });
        }
    }
}