#region [ using ]
using System;
using System.Text;
using System.Data;
using System.Data.Sql;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.Diagnostics;
#endregion

namespace SqlClrUtil
{
	public class SqlClr
	{
		public static readonly string VERSAO = "1.04 - 12.JUL.2019";

		#region [ Comentário sobre as versões ]
		/*================================================================================================
		* v 1.03 - 27.01.2018 - por HHO
		*		  Implementação de novas funções:
		*				codificaTextoHex
		*				decodificaTextoHex
		*				codificaTextoUnicodeHex
		*				decodificaTextoUnicodeHex
		*				digitos
		*				converteInteiro
		*				formataCep
		*				formataCnpjCpf
		*				formataInteiro
		*				formataMoeda
		* -----------------------------------------------------------------------------------------------
		* v 1.04 - 12.07.2019 - por HHO
		*		  Implementação de tratamento para o novo meio de pagamento 'cartão (maquineta)'.
		* -----------------------------------------------------------------------------------------------
		* v 1.05 - XX.XX.20XX - por XXX
		*		  
		* -----------------------------------------------------------------------------------------------
		* v 1.06 - XX.XX.20XX - por XXX
		*		  
		* -----------------------------------------------------------------------------------------------
		* v 1.07 - XX.XX.20XX - por XXX
		*		  
		* -----------------------------------------------------------------------------------------------
		* v 1.08 - XX.XX.20XX - por XXX
		*		  
		* -----------------------------------------------------------------------------------------------
		* v 1.09 - XX.XX.20XX - por XXX
		*		  
		* -----------------------------------------------------------------------------------------------
		* v 1.10 - XX.XX.20XX - por XXX
		*		  
		* ===============================================================================================
		*/
		#endregion

		#region [ versao ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString versao()
		{
			return (SqlString)VERSAO;
		}
		#endregion

		#region [ decodificaDadoHex ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString decodificaDadoHex(SqlString dadoParaDecodificar)
		{
			#region [ Declarações ]
			string strResposta;
			string msg_erro;
			#endregion

			if (!CriptoHex.decodificaDado(dadoParaDecodificar.ToString(), out strResposta, out msg_erro))
			{
				strResposta = "";
				if ((msg_erro ?? "").Trim().Length > 0) Debug.Print(msg_erro);
			}
			return (SqlString)strResposta;
		}
		#endregion

		#region [ codificaDadoHex ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString codificaDadoHex(SqlString dadoParaCodificar)
		{
			#region [ Declarações ]
			string strResposta;
			string msg_erro;
			#endregion

			if (!CriptoHex.codificaDado(dadoParaCodificar.ToString(), out strResposta, out msg_erro))
			{
				strResposta = "";
				if ((msg_erro ?? "").Trim().Length > 0) Debug.Print(msg_erro);
			}
			return (SqlString)strResposta;
		}
		#endregion

		#region [ codificaTextoHex ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString codificaTextoHex(SqlString textoParaCodificar)
		{
			#region [ Declarações ]
			string strResposta;
			string msg_erro;
			#endregion

			if (!CriptoHex.codificaTexto(textoParaCodificar.ToString(), out strResposta, out msg_erro))
			{
				strResposta = "";
				if ((msg_erro ?? "").Trim().Length > 0) Debug.Print(msg_erro);
			}
			return (SqlString)strResposta;
		}
		#endregion

		#region [ codificaTextoUnicodeHex ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString codificaTextoUnicodeHex(SqlString textoParaCodificar)
		{
			#region [ Declarações ]
			string strResposta;
			string msg_erro;
			#endregion

			if (!CriptoHex.codificaTextoUnicode(textoParaCodificar.ToString(), out strResposta, out msg_erro))
			{
				strResposta = "";
				if ((msg_erro ?? "").Trim().Length > 0) Debug.Print(msg_erro);
			}
			return (SqlString)strResposta;
		}
		#endregion

		#region [ decodificaTextoHex ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString decodificaTextoHex(SqlString textoParaDecodificar)
		{
			#region [ Declarações ]
			string strResposta;
			string msg_erro;
			#endregion

			if (!CriptoHex.decodificaTexto(textoParaDecodificar.ToString(), out strResposta, out msg_erro))
			{
				strResposta = "";
				if ((msg_erro ?? "").Trim().Length > 0) Debug.Print(msg_erro);
			}
			return (SqlString)strResposta;
		}
		#endregion

		#region [ decodificaTextoUnicodeHex ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString decodificaTextoUnicodeHex(SqlString textoParaDecodificar)
		{
			#region [ Declarações ]
			string strResposta;
			string msg_erro;
			#endregion

			if (!CriptoHex.decodificaTextoUnicode(textoParaDecodificar.ToString(), out strResposta, out msg_erro))
			{
				strResposta = "";
				if ((msg_erro ?? "").Trim().Length > 0) Debug.Print(msg_erro);
			}
			return (SqlString)strResposta;
		}
		#endregion

		#region [ converteInteiro ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlInt32 converteInteiro(SqlString numero)
		{
			int n;
			n = Global.converteInteiro(numero.ToString());
			return (SqlInt32)n;
		}
		#endregion

		#region [ digitos ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString digitos(SqlString texto)
		{
			string strResposta;
			strResposta = Texto.digitos(texto.ToString());
			return (SqlString)strResposta;
		}
		#endregion

		#region [ formataCep ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString formataCep(SqlString cep)
		{
			string strResposta;
			strResposta = Global.formataCep(cep.ToString());
			return (SqlString)strResposta;
		}
		#endregion

		#region [ formataCnpjCpf ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString formataCnpjCpf(SqlString cnpj_cpf)
		{
			string strResposta;
			strResposta = Global.formataCnpjCpf(cnpj_cpf.ToString());
			return (SqlString)strResposta;
		}
		#endregion

		#region [ formataInteiro ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString formataInteiro(SqlInt32 numero)
		{
			string strResposta;
			strResposta = Global.formataInteiro((int)numero);
			return (SqlString)strResposta;
		}
		#endregion

		#region [ formataMoeda ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString formataMoeda(SqlMoney valor)
		{
			string strResposta;
			strResposta = Global.formataMoeda((decimal)valor);
			return (SqlString)strResposta;
		}
		#endregion

		#region [ iniciaisEmMaiusculas ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlString iniciaisEmMaiusculas(SqlString texto)
		{
			return (SqlString)Texto.iniciaisEmMaiusculas(texto.ToString());
		}
		#endregion

		#region [ calculaValorMeioPagtoEspecificadoFormaPagtoPedido ]
		[Microsoft.SqlServer.Server.SqlFunction(IsDeterministic = true, IsPrecise = true)]
		public static SqlMoney calculaValorMeioPagtoEspecificadoFormaPagtoPedido(SqlInt16 meio_pagto_especificado,
																			 SqlMoney vlTotalPedido,
																			 SqlInt16 tipo_parcelamento,
																			 SqlInt16 av_forma_pagto,
																			 SqlInt16 pu_forma_pagto, SqlMoney pu_valor,
																			 SqlInt16 pc_qtde_parcelas, SqlMoney pc_valor_parcela,
																			 SqlInt16 pc_maquineta_qtde_parcelas, SqlMoney pc_maquineta_valor_parcela,
																			 SqlInt16 pce_forma_pagto_entrada, SqlInt16 pce_forma_pagto_prestacao, SqlMoney pce_entrada_valor, SqlInt16 pce_prestacao_qtde, SqlMoney pce_prestacao_valor,
																			 SqlInt16 pse_forma_pagto_prim_prest, SqlInt16 pse_forma_pagto_demais_prest, SqlMoney pse_prim_prest_valor, SqlInt16 pse_demais_prest_qtde, SqlMoney pse_demais_prest_valor)
		{
			#region [ Declarações ]
			SqlMoney vlMeioPagto = 0;
			#endregion

			if (tipo_parcelamento == (SqlInt16)Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				if (av_forma_pagto == meio_pagto_especificado) vlMeioPagto = vlTotalPedido;
			}
			else if (tipo_parcelamento == (SqlInt16)Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
			{
				if (pu_forma_pagto == meio_pagto_especificado) vlMeioPagto = pu_valor;
			}
			else if (tipo_parcelamento == (SqlInt16)Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO)
			{
				if (meio_pagto_especificado == (SqlInt16)Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) vlMeioPagto = pc_qtde_parcelas * pc_valor_parcela;
			}
			else if (tipo_parcelamento == (SqlInt16)Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA)
			{
				if (meio_pagto_especificado == (SqlInt16)Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO_MAQUINETA) vlMeioPagto = pc_maquineta_qtde_parcelas * pc_maquineta_valor_parcela;
			}
			else if (tipo_parcelamento == (SqlInt16)Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
			{
				if (pce_forma_pagto_entrada == meio_pagto_especificado) vlMeioPagto = pce_entrada_valor;
				if (pce_forma_pagto_prestacao == meio_pagto_especificado) vlMeioPagto += pce_prestacao_qtde * pce_prestacao_valor;
			}
			else if (tipo_parcelamento == (SqlInt16)Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
			{
				if (pse_forma_pagto_prim_prest == meio_pagto_especificado) vlMeioPagto = pse_prim_prest_valor;
				if (pse_forma_pagto_demais_prest == meio_pagto_especificado) vlMeioPagto += pse_demais_prest_qtde * pse_demais_prest_valor;
			}

			return vlMeioPagto;
		}
		#endregion
	}
}
