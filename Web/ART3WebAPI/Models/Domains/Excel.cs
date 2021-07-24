using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ART3WebAPI.Models.Domains
{
	public class Excel
	{
		#region [ ConverteNumeracaoDigitoParaLetra ]
		public static string ConverteNumeracaoDigitoParaLetra(int numeracaoDigito)
		{
			#region [ Declarações ]
			const int TOTAL_LETRAS_ALFABETO = 26;
			string strResp;
			int intQuoc;
			int intResto;
			#endregion

			strResp = "";
			if (numeracaoDigito <= 0) return "";
			intQuoc = (int)(numeracaoDigito - 1) / TOTAL_LETRAS_ALFABETO;
			intResto = numeracaoDigito - (intQuoc * TOTAL_LETRAS_ALFABETO);
			if (intQuoc > TOTAL_LETRAS_ALFABETO) return "";
			if (intQuoc > 0) strResp = ((char)(65 - 1 + intQuoc)).ToString();
			strResp += ((char)(65 - 1 + intResto)).ToString();
			return strResp;
		}
		#endregion

		#region [ CellAddress ]
		public static string CellAddress(int colNumber, int rowNumber)
		{
			string sCol, sCellAddress;
			sCol = ConverteNumeracaoDigitoParaLetra(colNumber);
			sCellAddress = sCol + rowNumber.ToString();
			return sCellAddress;
		}

		public static string CellAddress(string colLetter, int rowNumber)
		{
			string sCellAddress;
			sCellAddress = colLetter + rowNumber.ToString();
			return sCellAddress;
		}
		#endregion

		#region [ RangeAddress ]
		public static string RangeAddress(int colNumberBegin, int rowNumberBegin, int colNumberEnd, int rowNumberEnd)
		{
			string sColBegin, sColEnd, sRangeAddress;
			sColBegin = ConverteNumeracaoDigitoParaLetra(colNumberBegin);
			sColEnd = ConverteNumeracaoDigitoParaLetra(colNumberEnd);
			sRangeAddress = sColBegin + rowNumberBegin.ToString() + ":" + sColEnd + rowNumberEnd.ToString();
			return sRangeAddress;
		}

		public static string RangeAddress(string colLetterBegin, int rowNumberBegin, string colLetterEnd, int rowNumberEnd)
		{
			string sRangeAddress;
			sRangeAddress = colLetterBegin + rowNumberBegin.ToString() + ":" + colLetterEnd + rowNumberEnd.ToString();
			return sRangeAddress;
		}
		#endregion

		public static void SetRangeBold(ExcelRangeBase range, bool bold)
		{
			range.Style.Font.Bold = bold;
		}

		public static void SetFontSize(ExcelRangeBase range, float size)
		{
			range.Style.Font.Size = size;
		}
	}
}