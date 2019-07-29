#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace Financeiro
{
	public class FinanceiroException : Exception
	{
		#region [ Construtor ]
		public FinanceiroException() : base() { }
		public FinanceiroException(String mensagem) : base(mensagem) { }
		public FinanceiroException(String mensagem, Exception innerException) : base(mensagem, innerException) { }
		#endregion
	}
}
