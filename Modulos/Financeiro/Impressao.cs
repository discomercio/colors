#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
#endregion

namespace Financeiro
{
	class Impressao
	{
		#region [ Constantes ]
		public const float INCH_TO_MILLIMETER = 25.4f;
		private const float MARGEM_ESQUERDA_EM_MM_DEFAULT = 20.0f;
		private const float MARGEM_DIREITA_EM_MM_DEFAULT = 10.0f;
		private const float MARGEM_SUPERIOR_EM_MM_DEFAULT = 5.0f;
		private const float MARGEM_INFERIOR_EM_MM_DEFAULT = 15.0f;
		#endregion

		#region [ Atributos ]
		private float _margemEsquerdaEmMm;
		public float margemEsquerdaEmMm
		{
			get { return _margemEsquerdaEmMm; }
			set { _margemEsquerdaEmMm = value; }
		}

		private float _margemDireitaEmMm;
		public float margemDireitaEmMm
		{
			get { return _margemDireitaEmMm; }
			set { _margemDireitaEmMm = value; }
		}

		private float _margemSuperiorEmMm;
		public float margemSuperiorEmMm
		{
			get { return _margemSuperiorEmMm; }
			set { _margemSuperiorEmMm = value; }
		}

		private float _margemInferiorEmMm;
		public float margemInferiorEmMm
		{
			get { return _margemInferiorEmMm; }
			set { _margemInferiorEmMm = value; }
		}
		#endregion

		#region [ Construtor ]
		public Impressao()
		{
			_margemEsquerdaEmMm = MARGEM_ESQUERDA_EM_MM_DEFAULT;
			_margemDireitaEmMm = MARGEM_DIREITA_EM_MM_DEFAULT;
			_margemSuperiorEmMm = MARGEM_SUPERIOR_EM_MM_DEFAULT;
			_margemInferiorEmMm = MARGEM_INFERIOR_EM_MM_DEFAULT;
		}

		public Impressao(bool landscape)
		{
			if (landscape)
			{
				_margemEsquerdaEmMm = MARGEM_ESQUERDA_EM_MM_DEFAULT;
				_margemDireitaEmMm = MARGEM_DIREITA_EM_MM_DEFAULT;
				_margemSuperiorEmMm = MARGEM_SUPERIOR_EM_MM_DEFAULT;
				_margemInferiorEmMm = 5f;
			}
			else
			{
				_margemEsquerdaEmMm = MARGEM_ESQUERDA_EM_MM_DEFAULT;
				_margemDireitaEmMm = MARGEM_DIREITA_EM_MM_DEFAULT;
				_margemSuperiorEmMm = MARGEM_SUPERIOR_EM_MM_DEFAULT;
				_margemInferiorEmMm = MARGEM_INFERIOR_EM_MM_DEFAULT;
			}
		}

		public Impressao(float margemEsquerdaEmMilimetros, float margemDireitaEmMilimetros, float margemSuperiorEmMilimetros, float margemInferiorEmMilimetros)
		{
			this._margemEsquerdaEmMm = margemEsquerdaEmMilimetros;
			this._margemDireitaEmMm = margemDireitaEmMilimetros;
			this._margemSuperiorEmMm = margemSuperiorEmMilimetros;
			this._margemInferiorEmMm = margemInferiorEmMilimetros;
		}
		#endregion

		#region [ Métodos ]

		#region [ converteParaMm ]
		/// <summary>
		/// Dado um valor em "hundredths of an inch", converte para milímetros
		/// </summary>
		/// <param name="valor">
		/// Valor expresso em "hundredths of an inch"
		/// </param>
		/// <returns>
		/// Retorna o valor convertido para milímetros
		/// </returns>
		public static float converteParaMm(float valor)
		{
			// As medidas são expressas em "hundredths of an inch", portanto, 100 = 1 polegada (ou 100dpi)
			return (valor / 100) * INCH_TO_MILLIMETER;
		}
		#endregion

		#region [ getLeftMarginInMm ]
		/// <summary>
		/// Retorna a largura da margem esquerda em milímetros. É retornado o maior valor entre:
		///		1) Mínimo da impressora
		///		2) Valor definido no atributo "margemEsquerdaEmMm"
		/// </summary>
		/// <param name="e">
		/// Parâmetro System.Drawing.Printing.PrintPageEventArgs do evento PrintPage
		/// </param>
		/// <returns>
		/// Retorna a largura da margem esquerda em milímetros
		/// </returns>
		public float getLeftMarginInMm(System.Drawing.Printing.PrintPageEventArgs e)
		{
			float resultado;

			if (!e.PageSettings.Landscape)
				resultado = Math.Max(converteParaMm(e.PageSettings.PrintableArea.Left), _margemEsquerdaEmMm) - converteParaMm(e.PageSettings.HardMarginX);
			else
				resultado = Math.Max(converteParaMm(e.PageSettings.PrintableArea.Top), _margemEsquerdaEmMm) - converteParaMm(e.PageSettings.HardMarginY);
			
			return resultado;
		}
		#endregion

		#region [ getRightMarginInMm ]
		/// <summary>
		/// Retorna a largura da margem direita em milímetros. É retornado o maior valor entre:
		///		1) Mínimo da impressora
		///		2) Valor definido no atributo "margemDireitaEmMm"
		/// </summary>
		/// <param name="e">
		/// Parâmetro System.Drawing.Printing.PrintPageEventArgs do evento PrintPage
		/// </param>
		/// <returns>
		/// Retorna a largura da margem direita em milímetros
		/// </returns>
		public float getRightMarginInMm(System.Drawing.Printing.PrintPageEventArgs e)
		{
			float resultado;

			if (!e.PageSettings.Landscape)
				resultado = Math.Max(converteParaMm(e.PageSettings.PaperSize.Width - e.PageSettings.PrintableArea.Right), _margemDireitaEmMm) - converteParaMm(e.PageSettings.HardMarginX);
			else
				resultado = Math.Max(converteParaMm(e.PageSettings.PaperSize.Height - e.PageSettings.PrintableArea.Bottom), _margemDireitaEmMm) - converteParaMm(e.PageSettings.HardMarginY);

			return resultado;
		}
		#endregion

		#region [ getWidthInMm ]
		/// <summary>
		/// Retorna a largura disponível para impressão em milímetros, já descontando as margens esquerda e direita calculadas por getLeftMarginInMm() e getRightMarginInMm()
		/// </summary>
		/// <param name="e">
		/// Parâmetro System.Drawing.Printing.PrintPageEventArgs do evento PrintPage
		/// </param>
		/// <returns>
		/// Retorna a largura disponível para impressão em milímetros, já descontando as margens esquerda e direita calculadas por getLeftMarginInMm() e getRightMarginInMm()
		/// </returns>
		public float getWidthInMm(System.Drawing.Printing.PrintPageEventArgs e)
		{
			float resultado;

			if (!e.PageSettings.Landscape)
				resultado = converteParaMm(e.PageSettings.PaperSize.Width) - getLeftMarginInMm(e) - getRightMarginInMm(e) - 2 * converteParaMm(e.PageSettings.HardMarginX);
			else
				resultado = converteParaMm(e.PageSettings.PaperSize.Height) - getLeftMarginInMm(e) - getRightMarginInMm(e) - 2 * converteParaMm(e.PageSettings.HardMarginY);

			return resultado;
		}
		#endregion

		#region [ getTopMarginInMm ]
		/// <summary>
		/// Retorna a margem superior em milímetros. É retornado o maior valor entre:
		///		1) Mínimo da impressora
		///		2) Valor definido no atributo "margemSuperiorEmMm"
		/// </summary>
		/// <param name="e">
		/// Parâmetro System.Drawing.Printing.PrintPageEventArgs do evento PrintPage
		/// </param>
		/// <returns>
		/// Retorna a altura da margem superior em milímetros
		/// </returns>
		public float getTopMarginInMm(System.Drawing.Printing.PrintPageEventArgs e)
		{
			float resultado;

			if (!e.PageSettings.Landscape)
				resultado = Math.Max(converteParaMm(e.PageSettings.PrintableArea.Top), _margemSuperiorEmMm) - converteParaMm(e.PageSettings.HardMarginY);
			else
				resultado = Math.Max(converteParaMm(e.PageSettings.PaperSize.Width - e.PageSettings.PrintableArea.Right), _margemSuperiorEmMm) - converteParaMm(e.PageSettings.HardMarginX);

			return resultado;
		}
		#endregion

		#region [ getBottomMarginInMm ]
		/// <summary>
		/// Retorna a margem inferior em milímetros. É retornado o maior valor entre:
		///		1) Mínimo da impressora
		///		2) Valor definido no atributo "margemInferiorEmMm"
		/// </summary>
		/// <param name="e">
		/// Parâmetro System.Drawing.Printing.PrintPageEventArgs do evento PrintPage
		/// </param>
		/// <returns>
		/// Retorna a altura da margem inferior em milímetros
		/// </returns>
		public float getBottomMarginInMm(System.Drawing.Printing.PrintPageEventArgs e)
		{
			float resultado;

			if (!e.PageSettings.Landscape)
				resultado = Math.Max(converteParaMm(e.PageSettings.PaperSize.Height - e.PageSettings.PrintableArea.Bottom), _margemInferiorEmMm) - converteParaMm(e.PageSettings.HardMarginY);
			else
				resultado = Math.Max(converteParaMm(e.PageSettings.PrintableArea.Left), _margemInferiorEmMm) - converteParaMm(e.PageSettings.HardMarginX);
			return resultado;
		}
		#endregion

		#region [ getHeightInMm ]
		/// <summary>
		/// Retorna a altura disponível para impressão em milímetros, já descontando as margens superior e inferior calculadas por getTopMarginInMm() e getBottomMarginInMm()
		/// </summary>
		/// <param name="e">
		/// Parâmetro System.Drawing.Printing.PrintPageEventArgs do evento PrintPage
		/// </param>
		/// <returns>
		/// Retorna a altura disponível para impressão em milímetros, já descontando as margens superior e inferior calculadas por getTopMarginInMm() e getBottomMarginInMm()
		/// </returns>
		public float getHeightInMm(System.Drawing.Printing.PrintPageEventArgs e)
		{
			float resultado;

			if (!e.PageSettings.Landscape)
				resultado = converteParaMm(e.PageSettings.PaperSize.Height) - getTopMarginInMm(e) - getBottomMarginInMm(e) - 2 * converteParaMm(e.PageSettings.HardMarginX);
			else
				resultado = converteParaMm(e.PageSettings.PaperSize.Width) - getTopMarginInMm(e) - getBottomMarginInMm(e) - 2 * converteParaMm(e.PageSettings.HardMarginX);

			return resultado;
		}
		#endregion

		#region [ criaPenTracoPontilhado ]
		public static Pen criaPenTracoPontilhado()
		{
			Pen penTracoPontilhado;
			penTracoPontilhado = new Pen(new SolidBrush(Color.Black), .15f);
			// Define um padrão de pontilhado
			penTracoPontilhado.DashPattern = new float[] { 1f, 6f };
			return penTracoPontilhado;
		}
		#endregion

		#endregion
	}
}
