#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
#endregion

namespace ADM2
{
	class ExcelAutomation
	{
		/*
		 * Classe contendo as constantes e métodos wrappers para implementar a automação
		 * do Excel através de late binding.
		 * Observações:
		 * ============
		 * 1) As constantes foram obtidas através de um interop do Excel 12.0
		 * 2) O early binding é mais rápido no processamento e fornece as classes
		 *    já prontas para uso durante o desenvolvimento, porém, é mais difícil
		 *    administrar as diferentes versões de Excel durante o uso em ambiente
		 *    de produção. É necessário notar que o Interop usado durante o desen-
		 *    volvimento deve ser mais antigo ou igual à versão mais antiga do
		 *    Excel que será usado em produção.
		 * 3) Portanto, apesar de ser mais lento e ser mais trabalhoso durante o
		 *    desenvolvimento, o late binding se mostra mais flexível com relação
		 *    ao uso de diferentes versões de Excel.
		 */

		#region [ Constantes ]
		// Estas constantes foram definidas manualmente
		public static class PropertyType
		{
			public const string Value = "Value";
			public const string Range = "Range";
			public const string Interior = "Interior";
			public const string Name = "Name";
			public const string Visible = "Visible";
			public const string DisplayAlerts = "DisplayAlerts";
			public const string Item = "Item";
			public const string Workbooks = "Workbooks";
			public const string Worksheets = "Worksheets";
			public const string ActiveSheet = "ActiveSheet";
			public const string Pattern = "Pattern";
			public const string TintAndShade = "TintAndShade";
			public const string PatternTintAndShade = "PatternTintAndShade";
			public const string PatternColorIndex = "PatternColorIndex";
			public const string ThemeColor = "ThemeColor";
			public const string Color = "Color";
		}

		public static class MethodType
		{
			public const string Open = "Open";
			public const string Select = "Select";
			public const string Save = "Save";
			public const string Close = "Close";
			public const string Quit = "Quit";
		}
		#endregion

		#region [ Constantes (enum) ]

		#region [ XlBordersIndex ]
		public enum XlBordersIndex
		{
			xlDiagonalDown = 5,
			xlDiagonalUp = 6,
			xlEdgeLeft = 7,
			xlEdgeTop = 8,
			xlEdgeBottom = 9,
			xlEdgeRight = 10,
			xlInsideVertical = 11,
			xlInsideHorizontal = 12,
		}
		#endregion

		#region [ XlBorderWeight ]
		public enum XlBorderWeight
		{
			xlMedium = -4138,
			xlHairline = 1,
			xlThin = 2,
			xlThick = 4,
		}
		#endregion

		#region [ XlColorIndex ]
		// Estes códigos foram definidos manualmente
		public enum XlColorIndex
		{
			xlColorIndexNone = -4142,
			xlColorIndexAutomatic = -4105,
		}
		#endregion

		#region [ XlHAlign ]
		public enum XlHAlign
		{
			xlHAlignRight = -4152,
			xlHAlignLeft = -4131,
			xlHAlignJustify = -4130,
			xlHAlignDistributed = -4117,
			xlHAlignCenter = -4108,
			xlHAlignGeneral = 1,
			xlHAlignFill = 5,
			xlHAlignCenterAcrossSelection = 7,
		}
		#endregion

		#region [ XlVAlign ]
		public enum XlVAlign
		{
			xlVAlignTop = -4160,
			xlVAlignJustify = -4130,
			xlVAlignDistributed = -4117,
			xlVAlignCenter = -4108,
			xlVAlignBottom = -4107,
		}
		#endregion

		#region [ XlLineStyle ]
		public enum XlLineStyle
		{
			xlLineStyleNone = -4142,
			xlDouble = -4119,
			xlDot = -4118,
			xlDash = -4115,
			xlContinuous = 1,
			xlDashDot = 4,
			xlDashDotDot = 5,
			xlSlantDashDot = 13,
		}
		#endregion

		#region [ XlPaperSize ]
		public enum XlPaperSize
		{
			xlPaperLetter = 1,
			xlPaperLetterSmall = 2,
			xlPaperTabloid = 3,
			xlPaperLedger = 4,
			xlPaperLegal = 5,
			xlPaperStatement = 6,
			xlPaperExecutive = 7,
			xlPaperA3 = 8,
			xlPaperA4 = 9,
			xlPaperA4Small = 10,
			xlPaperA5 = 11,
			xlPaperB4 = 12,
			xlPaperB5 = 13,
			xlPaperFolio = 14,
			xlPaperQuarto = 15,
			xlPaper10x14 = 16,
			xlPaper11x17 = 17,
			xlPaperNote = 18,
			xlPaperEnvelope9 = 19,
			xlPaperEnvelope10 = 20,
			xlPaperEnvelope11 = 21,
			xlPaperEnvelope12 = 22,
			xlPaperEnvelope14 = 23,
			xlPaperCsheet = 24,
			xlPaperDsheet = 25,
			xlPaperEsheet = 26,
			xlPaperEnvelopeDL = 27,
			xlPaperEnvelopeC5 = 28,
			xlPaperEnvelopeC3 = 29,
			xlPaperEnvelopeC4 = 30,
			xlPaperEnvelopeC6 = 31,
			xlPaperEnvelopeC65 = 32,
			xlPaperEnvelopeB4 = 33,
			xlPaperEnvelopeB5 = 34,
			xlPaperEnvelopeB6 = 35,
			xlPaperEnvelopeItaly = 36,
			xlPaperEnvelopeMonarch = 37,
			xlPaperEnvelopePersonal = 38,
			xlPaperFanfoldUS = 39,
			xlPaperFanfoldStdGerman = 40,
			xlPaperFanfoldLegalGerman = 41,
			xlPaperUser = 256,
		}
		#endregion

		#region [ XlPageOrientation ]
		public enum XlPageOrientation
		{
			xlPortrait = 1,
			xlLandscape = 2,
		}
		#endregion

		#region [ XlPattern ]
		// Estes códigos foram definidos manualmente
		public enum XlPattern
		{
			xlNone = -4142,
			xlSolid = 1,
		}
		#endregion

		#region [ XlPatternColorIndex ]
		// Estes códigos foram definidos manualmente
		public enum XlPatternColorIndex
		{
			xlAutomatic = -4105,
		}
		#endregion

		#region [ XlThemeColor ]
		// Estes códigos foram definidos manualmente
		public enum XlThemeColor
		{
			xlThemeColorDark1 = 1,
			xlThemeColorDark2 = 3,
			xlThemeColorLight1 = 2,
			xlThemeColorLight2 = 4,
			xlThemeColorAccent1 = 5,
			xlThemeColorAccent2 = 6,
			xlThemeColorAccent3 = 7,
			xlThemeColorAccent4 = 8,
			xlThemeColorAccent5 = 9,
			xlThemeColorAccent6 = 10,
			xlThemeColorHyperlink = 11,
			xlThemeColorFollowedHyperlink = 12,
		}
		#endregion

		#region [ XlUnderlineStyle ]
		public enum XlUnderlineStyle
		{
			xlUnderlineStyleNone = -4142,
			xlUnderlineStyleDouble = -4119,
			xlUnderlineStyleSingle = 2,
			xlUnderlineStyleSingleAccounting = 4,
			xlUnderlineStyleDoubleAccounting = 5,
		}
		#endregion

		#region [ XlWindowState ]
		public enum XlWindowState
		{
			xlNormal = -4143,
			xlMinimized = -4140,
			xlMaximized = -4137,
		}
		#endregion

		#endregion

		#region [ Métodos Públicos ]

		#region [ CriaInstanciaExcel ]
		public static object CriaInstanciaExcel()
		{
			Type objClassType;
			objClassType = Type.GetTypeFromProgID("Excel.Application");
			return Activator.CreateInstance(objClassType);
		}
		#endregion

		#region [ SetProperty ]
		public static void SetProperty(object obj, string sProperty, object oValue)
		{
			object[] oParam = new object[1];
			oParam[0] = oValue;
			obj.GetType().InvokeMember(sProperty, BindingFlags.SetProperty, null, obj, oParam);
		}
		#endregion

		#region [ GetProperty (1 parâmetro) ]
		public static object GetProperty(object obj, string sProperty, object oValue)
		{
			object[] oParam = new object[1];
			oParam[0] = oValue;
			if (oValue == null)
			{
				return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, null);
			}
			else
			{
				return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oParam);
			}
		}
		#endregion

		#region [ GetProperty (2 parâmetros) ]
		public static object GetProperty(object obj, string sProperty, object oValue1, object oValue2)
		{
			object[] oParam = new object[2];
			oParam[0] = oValue1;
			oParam[1] = oValue2;
			return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oParam);
		}
		#endregion

		#region [ GetProperty (sem parâmetros) ]
		public static object GetProperty(object obj, string sProperty)
		{
			return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, null);
		}
		#endregion

		#region [ InvokeMethod ]
		public static object InvokeMethod(object obj, string sProperty, object[] oParam)
		{
			return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
		}
		#endregion

		#region [ InvokeMethod ]
		public static object InvokeMethod(object obj, string sProperty, object oValue)
		{
			object[] oParam = new object[1];
			oParam[0] = oValue;
			return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
		}
		#endregion

		#region [ NAR ]
		// https://support.microsoft.com/en-us/kb/317109
		public static void NAR(object o)
		{
			if (o == null) return;

			try
			{
				while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0) ;
			}
			catch { }
			finally
			{
				o = null;
			}
		}
		#endregion

		#endregion
	}
}
