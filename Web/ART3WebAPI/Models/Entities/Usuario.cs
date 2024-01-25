using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class Usuario
	{
		public string usuario { get; set; }
		public int Id { get; set; }
		public string nome { get; set; }
		public string senhaDescriptografada { get; set; }
		public string datastamp { get; set; }
		public bool bloqueado { get; set; }
		public bool senhaExpirada { get; set; }
		public string SessionTokenModuloCentral { get; set; }
		public string SessionTokenModuloLoja { get; set; }
	}
}