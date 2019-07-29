using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	public class Usuario
	{
		public string usuario { get; set; }
		public string nome { get; set; }
		public string senhaDescriptografada { get; set; }
		public string datastamp { get; set; }
		public bool bloqueado { get; set; }
		public bool senhaExpirada { get; set; }
		public string fin_email_remetente { get; set; }
		public string fin_servidor_smtp { get; set; }
		public int fin_servidor_smtp_porta { get; set; }
		public string fin_usuario_smtp { get; set; }
		public string fin_senha_smtp { get; set; }
		public string fin_display_name_remetente { get; set; }
	}
}
