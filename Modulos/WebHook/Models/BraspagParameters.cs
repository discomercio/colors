using System.Reflection;
using System.Text;

namespace WebHook.Models
{
	public class BraspagParameters
	{
		public string NumPedido { get; set; }

		public string Status { get; set; }

		public string CODPAGAMENTO { get; set; }

		#region [ FormataDados ]
		public string FormataDados(bool showOnePropertyPerLine = false, string inlinePropertySeparator = ", ")
		{
			#region [ Declarações ]
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			foreach (PropertyInfo prop in typeof(BraspagParameters).GetProperties())
			{
				if (showOnePropertyPerLine)
				{
					sbResp.AppendLine(prop.Name + " = " + (prop.GetValue(this, null) ?? "").ToString());
				}
				else
				{
					if (sbResp.Length > 0) sbResp.Append(inlinePropertySeparator ?? "");
					sbResp.Append(prop.Name + " = " + (prop.GetValue(this, null) ?? "").ToString());
				}
			}

			return sbResp.ToString();
		}
		#endregion
	}
}