using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Domains
{
	public class WebApiException : Exception
	{
		#region [ Construtor ]
		public WebApiException() : base() { }
		public WebApiException(string mensagem) : base(mensagem) { }
		public WebApiException(string mensagem, Exception innerException) : base(mensagem, innerException) { }
		#endregion
	}
}