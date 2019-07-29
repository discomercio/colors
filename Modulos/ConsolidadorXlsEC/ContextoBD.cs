using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	#region [ ContextoBD ]
	public class ContextoBD
	{
		#region [ Atributos ]
		public List<AmbienteBD> Ambientes = new List<AmbienteBD>();
		public AmbienteBD AmbienteBase;
		public readonly string DescricaoAmbiente = Global.GetConfigurationValue("AmbienteExecucao");
		#endregion

		#region [ Constructor ]
		public ContextoBD()
		{
		}
		#endregion
	}
	#endregion
}
