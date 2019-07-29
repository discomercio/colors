#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace EtqWms
{
	class CriptoHex
	{
		/*
		 * ========================== IMPORTANTE ==================================================================
		 * Estas rotinas de criptografia, além de criptografar/descriptografar a senha,
		 * convertem os caracteres da senha criptograda para códigos em hexadecimal.
		 * Com isso evita-se problemas de acentuação e/ou conversão de idiomas no banco
		 * de dados e dificulta-se ainda mais a interpretação dos dados.
		 * Obviamente, as rotinas são 'case sensitive', ou seja, letras maiúsculas e
		 * minúsculas geram resultados diferentes.
		 * A senha digitada pelo usuário nunca poderá ultrapassar o MENOR dos seguintes
		 * limites:
		 *		a) 255 caracteres
		 *		b) ((TAMANHO_SENHA_FORMATADA / 2) - 2) caracteres
		 * ========================================================================================================
		 */

		#region [ Contantes ]
		// FATOR (CRIPTOGRAFIA): ATÉ 9999
		private static int FATOR_CRIPTO = 1209;
		private static int TAMANHO_SENHA_FORMATADA = 32;  // Procurar usar sempre potência de 2
		private static String PREFIXO_SENHA_FORMATADA = "0x";
		private static int TAMANHO_CAMPO_COMPRIMENTO_SENHA = 2;
		#endregion

		#region [ converte_bin_para_dec ]
		/// <summary>
		/// Converte um número binário para sua forma decimal
		/// </summary>
		/// <param name="strNumero">
		/// Texto expressando um número binário (ex: "01010001")
		/// </param>
		/// <returns>
		/// Retorna um byte representando o número binário convertido para decimal
		/// </returns>
		private static byte converte_bin_para_dec(String strNumero)
		{
			try
			{
				return Convert.ToByte(strNumero, 2);
			}
			catch (Exception)
			{
				return 0;
			}
		}
		#endregion

		#region [ converte_dec_para_bin ]
		/// <summary>
		/// Converte um número decimal para sua forma binária
		/// O número é preenchido c/ zeros à esquerda, se necessário
		/// </summary>
		/// <param name="byteNumero">
		/// Byte informando um número em formato decimal
		/// </param>
		/// <returns>
		/// Retorna um texto expressando o número na forma binária
		/// </returns>
		private static String converte_dec_para_bin(byte byteNumero)
		{
			String s;
			s = Convert.ToString(byteNumero, 2);
			s = s.PadLeft(8, '0');
			return s;
		}
		#endregion

		#region [ gera_chave_codificacao ]
		/// <summary>
		/// Gera a chave para criptografia
		/// </summary>
		/// <param name="fator">
		/// Fator usado para geração da chave de criptografia
		/// </param>
		/// <param name="chave_gerada">
		/// Chave de criptografia gerada
		/// </param>
		/// <returns>
		/// true: sucesso
		/// false: falha
		/// </returns>
		private static bool gera_chave_codificacao(Int32 fator, ref String chave_gerada)
		{
			int COD_MINIMO = 35;
			int COD_MAXIMO = 96;
			int TAMANHO_CHAVE = 128;
			int i;
			Int64 k;
			StringBuilder s = new StringBuilder("");

			for (i = 1; i <= TAMANHO_CHAVE; i++)
			{
				k = COD_MAXIMO - COD_MINIMO + 1;
				k = k * fator;
				k = (k * i) + COD_MINIMO;
				k = k % 128;
				s.Append(Texto.chr((short)k));
			}
			chave_gerada = s.ToString();
			return true;
		}
		#endregion

		#region [ rotaciona_direita ]
		/// <summary>
		/// Rotaciona o byte para direita 'casas' posições
		/// Importante: o último bit da direita será colocado na 1ª casa da esquerda
		/// </summary>
		/// <param name="byteNumero">
		/// Número contido em um byte que será rotacionado
		/// </param>
		/// <param name="byteCasas">
		/// Número de casas a rotacionar
		/// </param>
		private static void rotaciona_direita(ref byte byteNumero, byte byteCasas)
		{
			int i;
			String s;
			String s_byte;

			// Transforma decimal -> binário ('0101...')
			s_byte = converte_dec_para_bin(byteNumero);

			//Rotaciona
			for (i = 1; i <= byteCasas; i++)
			{
				s = Texto.rightStr(s_byte, 1);
				s_byte = Texto.leftStr(s_byte, s_byte.Length - 1);
				s_byte = s + s_byte;
			}
			// Transforma binário -> decimal
			byteNumero = converte_bin_para_dec(s_byte);
		}
		#endregion

		#region [ rotaciona_esquerda ]
		/// <summary>
		/// Rotaciona o byte para esquerda 'casas' posições
		/// Importante: o 1º bit da esquerda será colocado na última casa da direita
		/// </summary>
		/// <param name="byteNumero">
		/// Número contido em um byte que será rotacionado
		/// </param>
		/// <param name="byteCasas">
		/// Número de casas a rotacionar
		/// </param>
		private static void rotaciona_esquerda(ref byte byteNumero, byte byteCasas)
		{
			int i;
			String s;
			String s_byte;

			// Transforma decimal -> binário ('0101...')
			s_byte = converte_dec_para_bin(byteNumero);

			// Rotaciona
			for (i = 1; i <= byteCasas; i++)
			{
				s = Texto.leftStr(s_byte, 1);
				s_byte = Texto.rightStr(s_byte, s_byte.Length - 1);
				s_byte += s;
			}

			// Transforma binário -> decimal
			byteNumero = converte_bin_para_dec(s_byte);
		}
		#endregion

		#region [ shift_direita ]
		/// <summary>
		/// Desloca o byte para direita 'casas' posições
		/// Importante: as casas da esquerda serão preenchidas com zeros
		/// </summary>
		/// <param name="byteNumero">
		/// Número contido em um byte que será deslocado
		/// </param>
		/// <param name="byteCasas">
		/// Número de casas a deslocar
		/// </param>
		private static void shift_direita(ref byte byteNumero, byte byteCasas)
		{
			int i;
			String s_byte;

			// Transforma decimal -> binário ('0101...')
			s_byte = converte_dec_para_bin(byteNumero);

			// Rotaciona
			for (i = 1; i <= byteCasas; i++)
			{
				s_byte = Texto.leftStr(s_byte, s_byte.Length - 1);
				s_byte = "0" + s_byte;
			}

			// Transforma binário -> decimal
			byteNumero = converte_bin_para_dec(s_byte);
		}
		#endregion

		#region [ shift_esquerda ]
		/// <summary>
		/// Desloca o byte para esquerda 'casas' posições
		/// Importante: as casas da direita serão preenchidas com zeros
		/// </summary>
		/// <param name="byteNumero">
		/// Número contido em um byte que será deslocado
		/// </param>
		/// <param name="byteCasas">
		/// Número de casas a deslocar
		/// </param>
		private static void shift_esquerda(ref byte byteNumero, byte byteCasas)
		{
			int i;
			String s_byte;

			// Transforma decimal -> binário ('0101...')
			s_byte = converte_dec_para_bin(byteNumero);

			// Rotaciona
			for (i = 1; i <= byteCasas; i++)
			{
				s_byte = Texto.rightStr(s_byte, s_byte.Length - 1);
				s_byte += "0";
			}

			// Transforma binário -> decimal
			byteNumero = converte_bin_para_dec(s_byte);
		}
		#endregion

		#region [ codificaDado ]
		/// <summary>
		/// Codifica o valor dado por 'strOrigem', utilizando a chave pré-definida.
		/// </summary>
		/// <param name="strOrigem">
		/// Senha sem criptografia
		/// </param>
		/// <param name="strDestino">
		/// Senha após criptografia
		/// </param>
		/// <param name="blnIncluiPreenchimento">
		/// Flag que indica se deve haver preenchimento dos campos não usados
		/// </param>
		/// <returns>
		/// true = sucesso
		/// false = falha
		/// </returns>
		/// <remarks>
		/// Esta função gera a senha criptografada, depois converte cada um dos caracteres
		/// criptografados para seu respectivo código hexadecimal e adiciona o prefixo '0xNN',
		/// sendo que 'NN' é um número hexadecimal indicando o tamanho da senha.
		/// O tamanho da senha indica os 'NN' caracteres da direita que devem ser utilizados
		/// para descriptografar a senha. Os caracteres restantes da esquerda são apenas para
		/// preenchimento e devem ser ignorados.
		/// Lembre-se de que a senha ocupa no BD, no mínimo (sem os caracteres de preenchimento):
		/// 2 x (tamanho descriptografado) + 2 bytes do '0x' + 2 bytes do 'NN'
		/// Por exemplo:
		/// 'AbCdEf' -> '0x0c34330210feccf497b2907e4d61ac7ad0be04ac09a3cd679061bb9d7fd923'
		///			 => '0x' -> prefixo a ser descartado.
		///			 =>   '0c' -> os 12 caracteres da direita contém a senha criptografada.
		///			 =>     '34330210feccf497b2907e4d61ac7ad0be04ac09a3cd6790' -> caracteres preenchimento que devem ser descartados.
		///			 =>                                                     '61bb9d7fd923' -> caracteres da senha criptografada.
		///			 
		/// A senha criptografada, portanto, é gerada em hexadecimal, com tamanho
		/// formatado para que seu comprimento total fique sempre com TAMANHO_SENHA_FORMATADA
		/// bytes (incluindo o '0xNN').
		/// Deve-se lembrar que a senha (descriptografada) em si poderá ter no máximo:
		/// ((TAMANHO_SENHA_FORMATADA / 2) - 2) caracteres
		/// </remarks>
		public static bool codificaDado(String strOrigem, ref String strDestino)
		{
			return codificaDado(strOrigem, ref strDestino, false);
		}

		public static bool codificaDado(String strOrigem, ref String strDestino, bool blnIncluiPreenchimento)
		{
			byte i;
			int i_tam_senha;
			byte i_chave;
			byte i_dado;
			byte k;
			String s_origem;
			String s_destino;
			String s;
			String chave;

			strDestino = "";

			// Senha de origem está vazia
			if (strOrigem == null) return false;
			if (strOrigem.Trim().Length == 0) return false;

			// Senha excede tamanho
			if (strOrigem.Trim().Length > ((TAMANHO_SENHA_FORMATADA - PREFIXO_SENHA_FORMATADA.Length) - TAMANHO_CAMPO_COMPRIMENTO_SENHA) / 2) return false;

			// Gera chave de criptografia
			chave = "";
			if (!gera_chave_codificacao(FATOR_CRIPTO, ref chave)) return false;

			s_destino = "";
			s_origem = strOrigem.Trim();

			// Criptografa pela chave
			for (i = 1; i <= s_origem.Length; i++)
			{
				i_chave = Texto.asc(chave.ToCharArray()[i - 1]);
				shift_esquerda(ref i_chave, 1);
				i_chave++;

				i_dado = Texto.asc(s_origem.ToCharArray()[i - 1]);
				rotaciona_esquerda(ref i_dado, 1);

				// XOR
				k = (byte)(i_chave ^ i_dado);

				s_destino = s_destino + Texto.chr(k);
			}

			// ASCII -> Hexadecimal
			s_origem = s_destino;
			s_destino = "";
			for (i = 1; i <= s_origem.Length; i++)
			{
				k = Texto.asc(s_origem.ToCharArray()[i - 1]);
				s = Texto.hex(k);
				s = s.PadLeft(2, '0');
				s_destino = s_destino + s;
			}

			// Guarda o tamanho real da senha
			i_tam_senha = s_destino.Length;

			if (blnIncluiPreenchimento)
			{
				// Coloca máscara (imita formato timestamp)
				i = 0;
				while (s_destino.Length < (TAMANHO_SENHA_FORMATADA - PREFIXO_SENHA_FORMATADA.Length - TAMANHO_CAMPO_COMPRIMENTO_SENHA))
				{
					// Ao invés de preencheer com zeros, gera código p/ preenchimento
					i++;
					s = Texto.hex(i ^ (Convert.ToInt16("0x" + s_destino.Substring(s_destino.Length - (i - 1) - 1, 1), 16)) ^ (Convert.ToInt16("0x" + s_destino.Substring(s_destino.Length - i - 1, 1), 16)));
					// Adiciona um caracter por vez p/ não ter o risco de ultrapassar o tamanho máximo
					s_destino = Texto.rightStr(s, 1) + s_destino;
				}

				// Adiciona prefixo e tamanho real da senha
				s = Texto.hex(i_tam_senha);
				s = s.PadLeft(2, '0');
			}
			else
			{
				while (s_destino.Length < (TAMANHO_SENHA_FORMATADA - PREFIXO_SENHA_FORMATADA.Length - TAMANHO_CAMPO_COMPRIMENTO_SENHA))
				{
					s_destino = "0" + s_destino;
				}
				s = "00";
			}

			s_destino = PREFIXO_SENHA_FORMATADA + s + s_destino;
			strDestino = s_destino.ToLower();
			return true;
		}
		#endregion

		#region [ decodificaDado ]
		/// <summary>
		/// Decodifica o valor dado por 'strOrigem', utilizando a chave pré-definida.
		/// </summary>
		/// <param name="strOrigem">
		/// Senha com criptografia
		/// </param>
		/// <param name="strDestino">
		/// Senha sem criptografia
		/// </param>
		/// <returns>
		/// true = sucesso
		/// false = falha
		/// </returns>
		/// <remarks>
		/// Esta função descriptografa a senha, convertendo os códigos hexadecimais
		/// de volta para os caracteres ASCII criptografados e, depois, descriptografando
		/// a senha.
		/// As 4 primeiras posições ('0xNN') formam o prefixo da senha, sendo que '0x'
		/// deve ser descartado e 'NN' indica o tamanho da senha.
		/// O tamanho da senha indica os 'NN' caracteres da direita que devem ser utilizados
		/// para descriptografar a senha. Os caracteres restantes da esquerda são apenas para
		/// preenchimento e devem ser ignorados.
		/// Lembre-se de que a senha ocupa no BD, no mínimo (sem os caracteres de preenchimento):
		/// 2 x (tamanho descriptografado) + 2 bytes do '0x' + 2 bytes do 'NN'
		/// Por exemplo:
		/// 'AbCdEf' -> '0x0c34330210feccf497b2907e4d61ac7ad0be04ac09a3cd679061bb9d7fd923'
		///			 => '0x' -> prefixo a ser descartado.
		///			 =>   '0c' -> os 12 caracteres da direita contém a senha criptografada.
		///			 =>     '34330210feccf497b2907e4d61ac7ad0be04ac09a3cd6790' -> caracteres preenchimento que devem ser descartados.
		///			 =>                                                     '61bb9d7fd923' -> caracteres da senha criptografada.
		///			 
		/// A senha criptografada é gerada com formatação para que seu tamanho
		/// total fique sempre com TAMANHO_SENHA_FORMATADA bytes (incluindo o '0xNN').
		/// Deve-se lembrar que a senha (descriptografada) em si poderá ter no máximo:
		/// ((TAMANHO_SENHA_FORMATADA / 2) - 2) caracteres
		/// </remarks>
		public static bool decodificaDado(String strOrigem, ref String strDestino)
		{
			byte i;
			byte i_chave;
			byte i_dado;
			byte k;
			String s_origem;
			String s_destino;
			String s;
			String chave = "";

			strDestino = "";

			if (strOrigem == null) return false;
			if (strOrigem.Trim().Length == 0) return false;

			if (!gera_chave_codificacao(FATOR_CRIPTO, ref chave)) return false;

			s_destino = "";
			s_origem = strOrigem.Trim();

			// Possui prefixo '0x'?
			if (!Texto.leftStr(s_origem, PREFIXO_SENHA_FORMATADA.Length).Equals(PREFIXO_SENHA_FORMATADA)) return false;

			// Retira prefixo '0x' da máscara (imita formato timestamp)
			s_origem = Texto.rightStr(s_origem, s_origem.Length - PREFIXO_SENHA_FORMATADA.Length);
			s_origem = s_origem.ToUpper();

			// Retira caracteres de preenchimento (imita formato timestamp)
			s = Texto.leftStr(s_origem, TAMANHO_CAMPO_COMPRIMENTO_SENHA);
			s = "0x" + s;
			try
			{
				i = Convert.ToByte(s, 16);
			}
			catch (Exception)
			{
				i = 0;
			}

			if (i != 0)
			{
				s_origem = Texto.rightStr(s_origem, i);
			}
			else
			{
				while (s_origem.Substring(0, 2).Equals("00"))
				{
					s_origem = Texto.rightStr(s_origem, s_origem.Length - 2);
				}
			}

			// Hexadecimal -> ASCII
			for (i = 1; i <= s_origem.Length; i += 2)
			{
				s = s_origem.Substring(i - 1, 2);
				s = "0x" + s;
				s_destino += Texto.chr(Convert.ToByte(s, 16));
			}

			// Descriptografa pela chave
			s_origem = s_destino;
			s_destino = "";
			for (i = 1; i <= s_origem.Length; i++)
			{
				i_chave = Texto.asc(chave.ToCharArray()[i - 1]);
				shift_esquerda(ref i_chave, 1);
				i_chave++;

				i_dado = Texto.asc(s_origem.ToCharArray()[i - 1]);
				// XOR
				k = (byte)(i_chave ^ i_dado);

				rotaciona_direita(ref k, 1);
				s_destino = s_destino + Texto.chr(k);
			}

			strDestino = s_destino;
			return true;
		}
		#endregion
	}
}
