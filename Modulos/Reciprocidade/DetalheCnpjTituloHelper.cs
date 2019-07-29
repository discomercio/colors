#region [ using ]
using System.Collections.Generic;
#endregion

namespace Reciprocidade
{
	class DetalheCnpjTituloHelper
	{
		#region [ Atributos ]
		private ArqRemessa.DetalheTempoRelacionamento _detalheRelacto;
		public ArqRemessa.DetalheTempoRelacionamento detalheRelacto
		{
			get { return _detalheRelacto; }
		}

		private List<ArqRemessa.DetalheTitulo> _titulos = new List<ArqRemessa.DetalheTitulo>();
		public List<ArqRemessa.DetalheTitulo> titulos
		{
			get { return _titulos; }
		}
		#endregion

		#region [ Construtor ]
		public DetalheCnpjTituloHelper() { }

		public DetalheCnpjTituloHelper(ArqRemessa.DetalheTempoRelacionamento detalheRelacto)
		{
			_detalheRelacto = detalheRelacto;
		}
		#endregion

		#region [ Métodos ]
		public void adicionaDetalheTitulo(ArqRemessa.DetalheTitulo titulo)
		{
			_titulos.Add(titulo);
		}
		#endregion
	}
}
