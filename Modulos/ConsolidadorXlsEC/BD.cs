#region[ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Configuration;
#endregion

namespace ConsolidadorXlsEC
{
	#region [ Classe BD ]
	class BD
	{
		#region[ Constantes ]
		public const int MAX_TAMANHO_VARCHAR = 8000;
		public const int MAX_TENTATIVAS_INSERT_BD = 3;
		public const int MAX_TENTATIVAS_UPDATE_BD = 2;
		public const int MAX_TENTATIVAS_DELETE_BD = 2;
		public const int intCommandTimeoutEmSegundos = 5 * 60;
		public const char CARACTER_CURINGA_TODOS = '%';
		#endregion

		#region[ Métodos ]

		#region [ readToString ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo texto
		/// </param>
		/// <returns>
		/// Retorna o texto armazenado no campo. Caso o conteúdo seja DBNull, retorna uma String vazia.
		/// </returns>
		public static String readToString(object campo)
		{
			return !Convert.IsDBNull(campo) ? campo.ToString() : "";
		}
		#endregion

		#region [ readToDateTime ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo data
		/// </param>
		/// <returns>
		/// Retorna a data armazenada no campo. Caso o conteúdo seja DBNull, retorna DateTime.MinValue
		/// </returns>
		public static DateTime readToDateTime(object campo)
		{
			return !Convert.IsDBNull(campo) ? (DateTime)campo : DateTime.MinValue;
		}
		#endregion

		#region [ readToSingle ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo real
		/// </param>
		/// <returns>
		/// Retorna o número real armazenado no campo
		/// </returns>
		public static Single readToSingle(object campo)
		{
			return (Single)campo;
		}
		#endregion

		#region [ readToByte ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo byte
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static byte readToByte(object campo)
		{
			return (byte)campo;
		}
		#endregion

		#region [ readToShort ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo short
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static short readToShort(object campo)
		{
			return (short)campo;
		}
		#endregion

		#region [ readToInt ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo int
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static int readToInt(object campo)
		{
			if (campo.GetType().Name.Equals("Int16"))
			{
				return (int)(Int16)campo;
			}
			else
			{
				return (int)campo;
			}
		}
		#endregion

		#region [ readToInt16 ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo System.Int16
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static System.Int16 readToInt16(object campo)
		{
			return (System.Int16)campo;
		}
		#endregion

		#region [ readToChar ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo char
		/// </param>
		/// <returns>
		/// Retorna o caracter armazenado no campo. Caso o conteúdo seja DBNull, retorna um caracter nulo.
		/// </returns>
		public static char readToChar(object campo)
		{
			String s;
			char c = '\0';

			if (!Convert.IsDBNull(campo))
			{
				s = campo.ToString();
				if (s.Length > 0) c = s[0];
			}

			return c;
		}
		#endregion

		#region [ readToDecimal ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo decimal
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static decimal readToDecimal(object campo)
		{
			return (decimal)campo;
		}
		#endregion

		#endregion
	}
	#endregion

	#region [ Classe VersaoModulo ]
	public class VersaoModulo
	{
		private string _modulo;
		public string modulo
		{
			get { return _modulo; }
			set { _modulo = value; }
		}

		private string _versao;
		public string versao
		{
			get { return _versao; }
			set { _versao = value; }
		}

		private string _mensagem;
		public string mensagem
		{
			get { return _mensagem; }
			set { _mensagem = value; }
		}

		private string _cor_fundo_padrao;
		public string cor_fundo_padrao
		{
			get { return _cor_fundo_padrao; }
			set { _cor_fundo_padrao = value; }
		}

		private string _identificador_ambiente;
		public string identificador_ambiente
		{
			get { return _identificador_ambiente; }
			set { _identificador_ambiente = value; }
		}
	}
	#endregion
}
