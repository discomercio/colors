#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;
#endregion

namespace EmailSenderService
{
	class Criptografia
	{
		#region[ Globais ]
		private byte[] key;
		private byte[] iv;
		#endregion

		#region[ (constructor} ]
		public Criptografia()
		{
			string chave = "bnZJ69V2fB8Z2iFd";
			string vetor_inicial = "9fgDR8xxJ90aQFB3";
			key = new byte[16];
			iv = new byte[16];

			int i;
			for (i = 0; i < chave.Length; i++) key[i] = Convert.ToByte(chave[i]);
			for (i = 0; i < vetor_inicial.Length; i++) iv[i] = Convert.ToByte(vetor_inicial[i]);
		}
		#endregion

		#region[ in (codifica) ]
		public string entra(string original)
		{
			try
			{
				RijndaelManaged algoritmo = new RijndaelManaged();
				algoritmo.BlockSize = 128;
				algoritmo.KeySize = 128;

				MemoryStream memStream = new MemoryStream();
				CryptoStream crStream = new CryptoStream(memStream, algoritmo.CreateEncryptor(key, iv), CryptoStreamMode.Write);
				StreamWriter strWriter = new StreamWriter(crStream);

				//	envia a frase original para criptografia através do stream
				strWriter.Write(original);
				strWriter.Flush();
				crStream.FlushFinalBlock();

				//	lê a matriz de bytes com o código resultante
				byte[] codigo = new byte[memStream.Length];
				memStream.Position = 0;
				memStream.Read(codigo, 0, codigo.Length);
				return Convert.ToBase64String(codigo);
			}
			catch
			{
				return "";
			}
		}
		#endregion

		#region[ sai (decodifica) ]
		public string sai(string codigo)
		{
			try
			{
				if (codigo == "") return "";
				RijndaelManaged algoritmo = new RijndaelManaged();
				algoritmo.BlockSize = 128;
				algoritmo.KeySize = 128;

				MemoryStream memStream = new MemoryStream(Convert.FromBase64String(codigo));

				memStream.Position = 0;
				CryptoStream crStream = new CryptoStream(memStream, algoritmo.CreateDecryptor(key, iv), CryptoStreamMode.Read);
				StreamReader strReader = new StreamReader(crStream);
				return strReader.ReadToEnd();
			}
			catch
			{
				return "";
			}
		}
		#endregion

		#region[ Criptografa ]
		public static string Criptografa(char[] Valor)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < Valor.Length; i++)
			{
				sb.Append(Valor[i]);

			}
			return Criptografa(sb.ToString());
		}
		public static string Criptografa(string Valor)
		{
			if (Valor == null) return "";
			if (Valor.Length == 0) return "";
			Criptografia c = new Criptografia();
			return c.entra(Valor);
		}
		#endregion

		#region[ Descriptografa ]
		public static string Descriptografa(string Valor)
		{
			if (Valor == null) return "";
			if (Valor.Length == 0) return "";
			Criptografia c = new Criptografia();
			return c.sai(Valor);
		}
		#endregion
	}
}
