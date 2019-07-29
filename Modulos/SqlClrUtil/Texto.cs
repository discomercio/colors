#region [ using ]
using System;
using System.Text;
#endregion

namespace SqlClrUtil
{
	static class Texto
	{
		#region [ leftStr ]
		/// <summary>
		/// Retorna a quantidade de caracteres especificada mais à esquerda do texto.
		/// Se o texto for null, retorna um texto vazio.
		/// Se a quantidade for menor ou igual a zero, retorna um texto vazio.
		/// Se a quantidade for maior que o tamanho do texto, retorna o próprio texto.
		/// </summary>
		/// <param name="texto">Texto a partir do qual será retornado um trecho</param>
		/// <param name="qtde">Quantidade de caracteres a ser retornada</param>
		/// <returns>
		/// Retorna um trecho mais à esquerda do texto especificado
		/// </returns>
		public static String leftStr(String texto, int qtde)
		{
			if (texto == null) return "";
			if (qtde <= 0) return "";

			if (qtde >= texto.Length) return texto;
			return texto.Substring(0, qtde);
		}
		#endregion

		#region [ rightStr ]
		/// <summary>
		/// Retorna a quantidade de caracteres especificada mais à direita do texto.
		/// Se o texto for null, retorna um texto vazio.
		/// Se a quantidade for menor ou igual a zero, retorna um texto vazio.
		/// Se a quantidade for maior que o tamanho do texto, retorna o próprio texto.
		/// </summary>
		/// <param name="texto">Texto a partir do qual será retornado um trecho</param>
		/// <param name="qtde">Quantidade de caracteres a ser retornada</param>
		/// <returns>
		/// Retorna um trecho mais à direita do texto especificado
		/// </returns>
		public static String rightStr(String texto, int qtde)
		{
			if (texto == null) return "";
			if (qtde <= 0) return "";

			if (qtde >= texto.Length) return texto;
			return texto.Substring(texto.Length - qtde, qtde);
		}
		#endregion

		#region [ midStr ]
		/// <summary>
		/// Retorna o texto a partir da posição inicial especificada.
		/// Se o texto for null, retorna um texto vazio.
		/// </summary>
		/// <param name="texto">Texto do qual será extraído um trecho</param>
		/// <param name="posicaoInicial">Posição inicial a partir da qual o texto restante será retornado. O primeiro caracter é a posição 1</param>
		/// <returns>Trecho mais à direita do texto a partir da posição inicial especificada</returns>
		public static String midStr(String texto, int posicaoInicialBase1)
		{
			int posicaoInicial;
			if (texto == null) return "";
			if (posicaoInicialBase1 > texto.Length) return "";
			posicaoInicial = posicaoInicialBase1 - 1;
			if (posicaoInicial < 0) posicaoInicial = 0;
			return texto.Substring(posicaoInicial);
		}
		#endregion

		#region [ chr ]
		/// <summary>
		/// Converte um código ASCII para char
		/// </summary>
		/// <param name="n">
		/// Código ASCII
		/// </param>
		/// <returns>
		/// Retorna o código ASCII convertido para char
		/// </returns>
		public static char chr(short n)
		{
			try
			{
				return Convert.ToChar(n);
			}
			catch (Exception)
			{
				return '\0';
			}
		}

		public static char chr(ushort n)
		{
			try
			{
				return Convert.ToChar(n);
			}
			catch (Exception)
			{
				return '\0';
			}
		}
		#endregion

		#region [ asc ]
		/// <summary>
		/// Converte um caracter char para o código ASCII
		/// </summary>
		/// <param name="c">
		/// Caracter char para converter
		/// </param>
		/// <returns>
		/// Retorna o código ASCII
		/// </returns>
		public static byte asc(char c)
		{
			return Convert.ToByte(c);
		}

		public static ushort ascUshort(char c)
		{
			return Convert.ToUInt16(c);
		}
		#endregion

		#region [ hex ]
		/// <summary>
		/// Converte um número decimal para um texto com sua representação em hexadecimal
		/// </summary>
		/// <param name="n">
		/// Número a ser convertido
		/// </param>
		/// <returns>
		/// Texto contendo a representação em hexadecimal
		/// </returns>
		public static string hex(int n)
		{
			return Convert.ToString(n, 16);
		}

		public static string hex(ushort n)
		{
			return Convert.ToString(n, 16);
		}
		#endregion

		#region [ haDigito ]
		public static bool haDigito(String texto)
		{
			if (texto == null) return false;

			for (int i = 0; i < texto.Length; i++)
			{
				if (isDigit(texto[i])) return true;
			}
			return false;
		}
		#endregion

		#region [ haVogal ]
		public static bool haVogal(String texto)
		{
			#region [ Declarações ]
			char c;
			#endregion

			if (texto == null) return false;

			for (int i = 0; i < texto.Length; i++)
			{
				c = texto[i];
				if ((c == 'A') || (c == 'E') || (c == 'I') || (c == 'O') || (c == 'U')) return true;
				if ((c == 'a') || (c == 'e') || (c == 'i') || (c == 'o') || (c == 'u')) return true;
			}
			return false;
		}
		#endregion

		#region [ iniciaisEmMaiusculas ]
		/// <summary>
		/// Retorna o texto formatado apenas com as iniciais em maiúsculas.
		/// </summary>
		/// <param name="texto">Texto a ser formatado</param>
		/// <returns>Retorna o texto informado apenas com as iniciais em maiúsculas</returns>
		public static string iniciaisEmMaiusculas(String texto)
		{
			#region [ Declarações ]
			const string PALAVRAS_MINUSCULAS = "|A|AS|AO|AOS|À|ÀS|E|O|OS|UM|UNS|UMA|UMAS|DA|DAS|DE|DO|DOS|EM|NA|NAS|NO|NOS|COM|SEM|POR|PELO|PELA|PARA|PRA|P/|S/|C/|TEM|OU|E/OU|ATE|ATÉ|QUE|SE|QUAL|";
			const string PALAVRAS_MAIUSCULAS = "|II|III|IV|VI|VII|VIII|IX|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX|XXI|XXII|XXIII|S/A|S/C|AC|AL|AM|AP|BA|CE|DF|ES|GO|MA|MG|MS|MT|PA|PB|PE|PI|PR|RJ|RN|RO|RR|RS|SC|SE|SP|TO|ME|EPP|";
			String strAspasDuplas = '\x0022'.ToString();
			String letra;
			StringBuilder palavra = new StringBuilder("");
			StringBuilder frase = new StringBuilder("");
			String s;
			String strPalavra;
			String[] v;
			bool blnAltera;
			#endregion

			if (texto == null) return "";

			for (int i = 0; i < texto.Length; i++)
			{
				letra = texto[i].ToString();
				palavra.Append(letra);
				if ((letra.Equals(" ")) || (i == (texto.Length - 1)) || (letra.Equals("(")) || (letra.Equals(")")) || (letra.Equals("[")) || (letra.Equals("]")) || (letra.Equals("'")) || (letra.Equals(strAspasDuplas)) || (letra.Equals("-")))
				{
					s = "|" + palavra.ToString().Trim().ToUpper() + "|";
					if ((PALAVRAS_MINUSCULAS.IndexOf(s) > -1) && (frase.ToString().Length > 0))
					{
						//  SE FOR FINAL DA FRASE, DEIXA INALTERADO (EX: BLOCO A)
						if (i < (texto.Length - 1)) palavra = new StringBuilder(palavra.ToString().ToLower());
					}
					else if (PALAVRAS_MAIUSCULAS.IndexOf(s) > -1)
					{
						palavra = new StringBuilder(palavra.ToString().ToUpper());
					}
					else
					{
						//  ANALISA SE CONVERTE O TEXTO OU NÃO
						blnAltera = true;
						strPalavra = palavra.ToString();
						if (haDigito(strPalavra))
						{
							//  ENDEREÇOS CUJO Nº DA RESIDÊNCIA SÃO SEPARADOS POR VÍRGULA, SEM NENHUM ESPAÇO EM BRANCO
							//  CASO CONTRÁRIO, CONSIDERA QUE É ALGUM TIPO DE CÓDIGO
							if (strPalavra.IndexOf(",") == -1) blnAltera = false;
						}

						if (blnAltera)
						{
							if (strPalavra.IndexOf(".") > -1)
							{
								v = strPalavra.Split('.');
								for (int k = 0; k < v.Length; k++)
								{
									if (PALAVRAS_MAIUSCULAS.IndexOf(v[k]) > -1)
									{
										v[k] = v[k].ToUpper();
									}
									else if (!haVogal(v[k]))
									{
										// NOP
									}
									else if (v[k].Length == 1)
									{
										// NOP
									}
									else
									{
										v[k] = v[k].Substring(0, 1).ToUpper() + v[k].Substring(1).ToLower();
									}
								}

								strPalavra = String.Join(".", v);
								//  A alteração na formatação já foi tratada aqui neste bloco
								blnAltera = false;
								palavra = new StringBuilder(strPalavra);
							}

							if (strPalavra.IndexOf("/") > -1)
							{
								//  S/C, S/A, S/C., S/A.
								if (strPalavra.Length <= 4)
								{
									blnAltera = false;
								}
								else
								{
									v = strPalavra.Split('/');
									for (int k = 0; k < v.Length; k++)
									{
										if (PALAVRAS_MAIUSCULAS.IndexOf(v[k]) > -1)
										{
											v[k] = v[k].ToUpper();
										}
										else if (!haVogal(v[k]))
										{
											// NOP
										}
										else if (v[k].Length == 1)
										{
											// NOP
										}
										else
										{
											v[k] = v[k].Substring(0, 1).ToUpper() + v[k].Substring(1).ToLower();
										}
									}

									strPalavra = String.Join("/", v);
									//  A alteração na formatação já foi tratada aqui neste bloco
									blnAltera = false;
									palavra = new StringBuilder(strPalavra);
								}
							}
						}

						if (blnAltera)
						{
							// Sigla?
							if (!haVogal(strPalavra)) blnAltera = false;
						}

						if (blnAltera) palavra = new StringBuilder(strPalavra.Substring(0, 1).ToUpper() + strPalavra.Substring(1).ToLower());
					}

					frase.Append(palavra.ToString());
					palavra = new StringBuilder("");
				}
			}
			return frase.ToString();
		}
		#endregion

		#region[ isDigit ]
		public static bool isDigit(char c)
		{
			if ((c >= '0') && (c <= '9')) return true;
			return false;
		}
		#endregion

		#region[ replica ]
		public static string replica(string texto, int qtdeRepeticoes)
		{
			StringBuilder sbTexto = new StringBuilder();
			for (int i = 0; i < qtdeRepeticoes; i++)
			{
				sbTexto.Append(texto);
			}
			return sbTexto.ToString();
		}
		#endregion

		#region[ digitos ]
		public static string digitos(string texto)
		{
			StringBuilder d = new StringBuilder();
			if (texto == null) return "";
			for (int i = 0; i < texto.Length; i++)
			{
				if ((texto[i] >= '0') && (texto[i] <= '9')) d.Append(texto[i]);
			}
			return d.ToString();
		}
		#endregion
	}
}
