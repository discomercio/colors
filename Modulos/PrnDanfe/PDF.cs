#region [ using ]
using System;
using Acrobat;
using System.Windows.Forms;
using System.IO;
#endregion

namespace PrnDANFE
{
	class PDF
	{
		#region [ testaAcrobat ]
		public static bool testaAcrobat()
		{
			#region [ Declarações ]
			CAcroApp AccroApp;
			bool ok;
			#endregion

			ok = false;

			try
			{
				AccroApp = new AcroApp();
				AccroApp.Exit();
				ok = true;
			}
			catch (Exception)
			{
				return false;
			}

			return ok;
		}
		#endregion

		#region [ abrePDF ]
		public static bool abrePDF(String ArqPDF)
		{
			#region [ Declarações ]
			CAcroApp AccroApp;
			CAcroPDDoc p;
			AcroAVDoc a;
			bool ok;
			#endregion

			if (!File.Exists(ArqPDF))
			{
				return false;
			}

			ok = true;

			try
			{
				AccroApp = new AcroApp();
				p = new AcroPDDoc();
				a = new AcroAVDoc();
				ok = p.Open(ArqPDF);
				a = p.OpenAVDoc(ArqPDF);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Problemas na abertura do arquivo " + ArqPDF + " - " + ex.ToString(), Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			return ok;
		}
		#endregion

		#region [ imprimePDF ]
		public static bool imprimePDF(String ArqPDF, bool modoInvisivel, bool closePDFAfterPrint)
		{
			#region [ Declarações ]
			int i;
			CAcroApp AccroApp;
			CAcroPDDoc p;
			AcroAVDoc a;
			bool ok;
			#endregion

			if (!File.Exists(ArqPDF))
			{
				return false;
			}

			ok = true;

			try
			{
				AccroApp = new AcroApp();
				if (modoInvisivel)
				{
					AccroApp.Hide();
				}
				p = new AcroPDDoc();
				a = new AcroAVDoc();
				ok = p.Open(ArqPDF);
				a = p.OpenAVDoc(ArqPDF);
				i = p.GetNumPages() - 1;
				a.PrintPagesSilent(0, i, 0, 0, 0);
				if (modoInvisivel)
				{
					AccroApp.Show();
				}
				if (closePDFAfterPrint)
				{
					a.Close(0);
					p.Close();
					AccroApp.Exit();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Problemas na impressão do arquivo " + ArqPDF + ": " + ex.ToString(), Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			return ok;
		}
		#endregion

		#region [ concatenaPDFs ]
		public static bool concatenaPDFs(ref string NovoArquivo)
		{
			#region [ Declarações ]
			int i;
			CAcroApp AccroApp;
			CAcroPDDoc p;
			bool ok;
			CAcroPDDoc InsertPDDoc;
			int iNumberOfPagesToInsert;
			int iLastPage;
			string[] ListaPDFs;
			string YYYYMMDD_HHMMSS = Global.formataDataYyyyMmDdComSeparador(DateTime.Now) + "_" + Global.formataHhMmSsComSigla(DateTime.Now);
			string strPastaAgrupados = "";
			#endregion

			ok = true;

			try
			{
				//se existe o arquivo com a concatenação de PDFs anteriores, apagar
				if (File.Exists(Global.Cte.Etc.ArqMergePDF))
				{
					File.Delete(Global.Cte.Etc.ArqMergePDF);
				}

				//obter os arquivos presentes no diretório de manipulação de PDFs
				ListaPDFs = Directory.GetFiles(Global.Cte.Etc.PathManipulaPDF, "*.pdf", SearchOption.TopDirectoryOnly);

				if (ListaPDFs.Length > 0)
				{
					Array.Sort(ListaPDFs);

					AccroApp = new AcroApp();
					//o esquema de ocultar/reexibir está dando problema, pois não fecha o Acrobat ao final;
					//por enquanto, não vamos ocultar
					//AccroApp.Hide();
					p = new AcroPDDoc();
					ok = p.Create();

					for (i = 0; i < ListaPDFs.Length; i++)
					{
						InsertPDDoc = new AcroPDDoc();

						ok = InsertPDDoc.Open(ListaPDFs[i]);
						iNumberOfPagesToInsert = InsertPDDoc.GetNumPages();
						iLastPage = p.GetNumPages() - 1;
						ok = p.InsertPages(iLastPage, InsertPDDoc, 0, iNumberOfPagesToInsert, 1);
						InsertPDDoc.Close();
					}

					p.Save(1, Global.Cte.Etc.ArqMergePDF);
					p.Close();
					//AccroApp.Show();

					//copiando o arquivo concatenado para a devida pasta
					strPastaAgrupados = Application.StartupPath + "\\" + Global.Usuario.strPastaEmitente + "\\PDF_AGRUPADO";
					if (!Directory.Exists(strPastaAgrupados)) Directory.CreateDirectory(strPastaAgrupados);
					NovoArquivo = strPastaAgrupados + "\\NFEs_" + YYYYMMDD_HHMMSS + ".pdf";
					File.Copy(Global.Cte.Etc.ArqMergePDF, NovoArquivo);

					AccroApp.Exit();
				}
				else
				{
					return false;
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Falha na concatenação de arquivos: " + ex.ToString(), Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			return ok;
		}
		#endregion

		#region [ limpaPastaManipulaPDFs ]
		public static bool limpaPastaManipulaPDFs()
		{
			#region [ Declarações ]
			string[] ArqsPDF;
			int i;
			bool ok;
			#endregion

			ok = false;

			if (!Directory.Exists(Global.Cte.Etc.PathManipulaPDF))
			{
				return false;
			}

			try
			{
				ArqsPDF = Directory.GetFiles(Global.Cte.Etc.PathManipulaPDF);
				for (i = 0; i < ArqsPDF.Length; i++)
				{
					File.Delete(ArqsPDF[i]);
				}
				ok = true;
			}
			catch (Exception)
			{
				return false;
			}

			return ok;
		}
		#endregion

		#region [ copiaPastaManipulaPDFs ]
		public static bool copiaPastaManipulaPDFs()
		{
			#region [ Declarações ]
			string[] ArqsPDF;
			string ArqDestino;
			string strPastaIndividuais = "";
			int i;
			bool ok;
			#endregion

			ok = false;

			if (!Directory.Exists(Global.Cte.Etc.PathManipulaPDF))
			{
				return false;
			}

			try
			{
				strPastaIndividuais = Application.StartupPath + "\\" + Global.Usuario.strPastaEmitente + "\\PDF_INDIVIDUAL";
				if (!Directory.Exists(strPastaIndividuais)) Directory.CreateDirectory(strPastaIndividuais);
				ArqsPDF = Directory.GetFiles(Global.Cte.Etc.PathManipulaPDF);
				for (i = 0; i < ArqsPDF.Length; i++)
				{
					if (ArqsPDF[i] == Global.Cte.Etc.ArqMergePDF) continue;
					ArqDestino = ArqsPDF[i];
					ArqDestino = ArqDestino.Replace(Global.Cte.Etc.PathManipulaPDF, strPastaIndividuais);
					File.Copy(ArqsPDF[i], ArqDestino, true);
				}
				ok = true;
			}
			catch (Exception)
			{
				return false;
			}

			return ok;
		}
		#endregion
	}
}
