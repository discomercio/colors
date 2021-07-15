using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;

namespace WebHookV2.Models.Entities
{
	public class BraspagPostNotificacao
	{
		public string RecurrentPaymentId { get; set; }
		public string PaymentId { get; set; }
		public byte ChangeType { get; set; }

		#region [ FormataDados ]
		public string FormataDados(bool showOnePropertyPerLine = false, string inlinePropertySeparator = ", ")
		{
			#region [ Declarações ]
			StringBuilder sbResp = new StringBuilder("");
			#endregion

			foreach (PropertyInfo prop in typeof(BraspagPostNotificacao).GetProperties())
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