#region [ using ]
using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
#endregion

namespace PrnDANFE
{
	class PDFBaixar
	{
		#region [ NFeFormataSerieNF ]
		public static String NFeFormataSerieNF(String numeroSerieNF)
		{
			#region [ Declarações ]
			String s_resp;
			#endregion

			s_resp = numeroSerieNF.Trim();
			while (s_resp.Length < 3)
			{
				s_resp = "0" + s_resp;
			}
			return s_resp;
		}
		#endregion

		#region [ NFeFormataNumeroNF ]
		public static String NFeFormataNumeroNF(String numeroNF)
		{
			#region [ Declarações ]
			String s_resp;
			#endregion

			s_resp = numeroNF.Trim();
			while (s_resp.Length < 9)
			{
				s_resp = "0" + s_resp;
			}
			return s_resp;
		}
        #endregion

		#region [ executa_download_pdf_danfe ]
		public static bool executa_download_pdf_danfe(String cedenteNFe, String numeroNFe, String serieNFe, ref String mensagemRetorno)
		{
			#region [ Declarações ]
			const String NomeDestaRotina = "executa_download_pdf_danfe()";
			const String BD_OLEDB_PROVIDER = "SQLOLEDB";
			String s;
			String s_aux;
			String strLogNF = "";
			String strLogNfSemDadosPdf = "";
			String strDiretorioPdfDanfe;
			String strNomeArqDanfe;
			String strNomeArqCompletoDanfe;
			String strNumeroNfNormalizado;
			String strSerieNfNormalizado;
			String strNomeEmitente = "";
			String strNfeT1ServidorBd;
			String strNfeT1NomeBd;
			String strNfeT1UsuarioBd;
			String strNfeT1SenhaCriptografadaBd;
			int id_boleto_cedente;
			int id_boleto_cedente_anterior;
			long intQtdeArqDownload = 0;
			long intQtdeNFeSemDadosPdf = 0;
			long intContadorRegistros;
			long intQtdeTotalRegistros;
			int intNfeRetornoSP;
			long lngFileSize;
			byte[] bytFile;
			int i;

			// BANCO DE DADOS
			OleDbConnection dbcNFe = null;
			DataTable t_FIN_BOLETO_CEDENTE = new DataTable();
			DataTable t_NFe_EMISSAO = new DataTable();
			OleDbCommand cmdNFeSituacao = new OleDbCommand();
			OleDbCommand cmdNFeDanfe = new OleDbCommand();
			DataTable rsNFeRetornoSPSituacao = new DataTable();
			DataTable rsNFeRetornoSPDanfe = new DataTable();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataRow rowConsultaEmissao;
			DataRow rowConsultaBoleto;
			OleDbDataReader readerNFeRetornoSPSituacao;
			OleDbDataReader readerNFeRetornoSPDanfe;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();

				cmdNFeSituacao.CommandType = CommandType.StoredProcedure;
				cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao";
				cmdNFeSituacao.Parameters.Add("NFe", OleDbType.Char, 9);
				cmdNFeSituacao.Parameters.Add("Serie", OleDbType.Char, 3);

				cmdNFeDanfe.CommandType = CommandType.StoredProcedure;
				cmdNFeDanfe.CommandText = "Proc_NFe_Danfe";
				cmdNFeDanfe.Parameters.Add("NFe", OleDbType.Char, 9);
				cmdNFeDanfe.Parameters.Add("Serie", OleDbType.Char, 3);

				s = "SELECT DISTINCT" +
						" id_boleto_cedente," +
						" id," +
						" NFe_serie_NF," +
						" NFe_numero_NF" +
					" FROM t_NFe_EMISSAO" +
					" WHERE" +
						" (id_boleto_cedente = " + cedenteNFe + ")" +
						" AND (NFe_numero_NF = '" + numeroNFe + "')" +
						" AND (NFe_serie_NF = '" + serieNFe + "')" +
					" ORDER BY" +
						" id_boleto_cedente," +
						" id DESC," +
						" NFe_serie_NF," +
						" NFe_numero_NF";
				cmCommand.CommandText = s;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(t_NFe_EMISSAO);

				if (t_NFe_EMISSAO.Rows.Count == 0)
				{
					mensagemRetorno = "NFe " + numeroNFe + " (série " + serieNFe + ") não localizada!!";
					return false;
				}

				id_boleto_cedente_anterior = -1;
				intContadorRegistros = 0;
				intQtdeTotalRegistros = t_NFe_EMISSAO.Rows.Count;
				for (i = 0; i < t_NFe_EMISSAO.Rows.Count; i++)
				{
					rowConsultaEmissao = t_NFe_EMISSAO.Rows[i];
					intContadorRegistros = intContadorRegistros + 1;
					id_boleto_cedente = BD.readToInt(rowConsultaEmissao["id_boleto_cedente"]);
					if (id_boleto_cedente != id_boleto_cedente_anterior)
					{
						s = "SELECT" +
								" nome_empresa," +
								" NFe_T1_servidor_BD," +
								" NFe_T1_nome_BD," +
								" NFe_T1_usuario_BD," +
								" NFe_T1_senha_BD" +
							" FROM t_FIN_BOLETO_CEDENTE" +
							" WHERE" +
								" (id = " + id_boleto_cedente.ToString() + ")";
						t_FIN_BOLETO_CEDENTE.Clear();
						cmCommand.CommandText = s;
						daAdapter.SelectCommand = cmCommand;
						daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
						daAdapter.Fill(t_FIN_BOLETO_CEDENTE);
						if (t_FIN_BOLETO_CEDENTE.Rows.Count == 0)
						{
							mensagemRetorno = "Falha ao localizar o registro em t_FIN_BOLETO_CEDENTE (id=" + id_boleto_cedente.ToString() + ")!!";
							return false;
						}

						rowConsultaBoleto = t_FIN_BOLETO_CEDENTE.Rows[0];
						strNomeEmitente = BD.readToString(rowConsultaBoleto["nome_empresa"]).ToUpper();
						strNfeT1ServidorBd = BD.readToString(rowConsultaBoleto["NFe_T1_servidor_BD"]);
						strNfeT1NomeBd = BD.readToString(rowConsultaBoleto["NFe_T1_nome_BD"]);
						strNfeT1UsuarioBd = BD.readToString(rowConsultaBoleto["NFe_T1_usuario_BD"]);
						strNfeT1SenhaCriptografadaBd = BD.readToString(rowConsultaBoleto["NFe_T1_senha_BD"]);

						s_aux = "";
						CriptoHex.decodificaDado(strNfeT1SenhaCriptografadaBd, ref s_aux);
						s = "Provider=" + BD_OLEDB_PROVIDER +
							";Data Source=" + strNfeT1ServidorBd +
							";Initial Catalog=" + strNfeT1NomeBd +
							";User Id=" + strNfeT1UsuarioBd +
							";Password=" + s_aux;

						id_boleto_cedente_anterior = id_boleto_cedente;

						if (dbcNFe != null)
						{
							if (dbcNFe.State != ConnectionState.Closed) dbcNFe.Close();
						}
						dbcNFe = new OleDbConnection();
						dbcNFe.ConnectionString = s;
						dbcNFe.Open();
					}

					strNumeroNfNormalizado = NFeFormataNumeroNF(BD.readToString(rowConsultaEmissao["NFe_numero_NF"]));
					strSerieNfNormalizado = NFeFormataSerieNF(BD.readToString(rowConsultaEmissao["NFe_serie_NF"]));

					//'   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
					cmdNFeSituacao.Connection = dbcNFe;
					cmdNFeDanfe.Connection = dbcNFe;

					cmdNFeSituacao.Parameters["NFe"].Value = strNumeroNfNormalizado;
					cmdNFeSituacao.Parameters["Serie"].Value = strSerieNfNormalizado;

					readerNFeRetornoSPSituacao = cmdNFeSituacao.ExecuteReader();
					try
					{
						intNfeRetornoSP = 0;
						if (readerNFeRetornoSPSituacao.Read()) intNfeRetornoSP = (byte)readerNFeRetornoSPSituacao[0];
					}
					finally
					{
						readerNFeRetornoSPSituacao.Close();
					}

					//se a DANFE não for localizada no banco de dados, retornar false
					//(controle na hora da impressão no programa)
					if (intNfeRetornoSP != 1)
					{
						mensagemRetorno = "NFe nº " + numeroNFe + " (série " + serieNFe + ") não localizada.";
						return false;
					}
					else
					{
						cmdNFeDanfe.Parameters["NFe"].Value = strNumeroNfNormalizado;
						cmdNFeDanfe.Parameters["Serie"].Value = strSerieNfNormalizado;

						readerNFeRetornoSPDanfe = cmdNFeDanfe.ExecuteReader();
						try
						{
							while (readerNFeRetornoSPDanfe.Read())
							{
								lngFileSize = readerNFeRetornoSPDanfe.GetBytes(0, 0, null, 0, 0);

								if (lngFileSize <= 0)
								{
									intQtdeNFeSemDadosPdf = intQtdeNFeSemDadosPdf + 1;
									if (strLogNfSemDadosPdf != "") strLogNfSemDadosPdf = strLogNfSemDadosPdf + ", ";
									strLogNfSemDadosPdf = strLogNfSemDadosPdf + id_boleto_cedente.ToString() + "/" + strSerieNfNormalizado + "/" + strNumeroNfNormalizado;
								}

								if (lngFileSize > 0) intQtdeArqDownload = intQtdeArqDownload + 1;

								//'   LOG
								if (strLogNF != "") strLogNF = strLogNF + ", ";
								strLogNF = strLogNF + id_boleto_cedente.ToString() + "/" + strSerieNfNormalizado + "/" + strNumeroNfNormalizado;

								//'   ARQUIVO DE DANFE
								strNomeArqDanfe = "NFe_" + strSerieNfNormalizado + "_" + strNumeroNfNormalizado + ".pdf";
								strDiretorioPdfDanfe = Global.Cte.Etc.PathManipulaPDF;

								if (!Directory.Exists(strDiretorioPdfDanfe))
								{
									try
									{
										Directory.CreateDirectory(strDiretorioPdfDanfe);
									}
									catch
									{
										mensagemRetorno = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" + strDiretorioPdfDanfe + ")";
										return false;
									}
								}

								strNomeArqCompletoDanfe = strDiretorioPdfDanfe + "\\" + strNomeArqDanfe;
								if (!File.Exists(strNomeArqCompletoDanfe))
								{
									try
									{
										File.Delete(strNomeArqCompletoDanfe);
									}
									catch (Exception)
									{
										mensagemRetorno = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" + strNomeArqCompletoDanfe + ")";
										return false;
									}
								}

								try
								{
									bytFile = new Byte[lngFileSize];
									readerNFeRetornoSPDanfe.GetBytes(0, 0, bytFile, 0, (int)lngFileSize);
									File.WriteAllBytes(strNomeArqCompletoDanfe, bytFile);
								}
								catch (Exception)
								{
									mensagemRetorno = "Falha ao tentar gravar o arquivo PDF (" + strNomeArqCompletoDanfe + ")";
									return false;
								}
							}  // while (readerNFeRetornoSPDanfe.Read())
						}
						finally
						{
							readerNFeRetornoSPDanfe.Close();
						}
					}
				}  // for
			}
			catch (Exception ex)
			{
				mensagemRetorno = ex.ToString();
				Global.gravaLogAtividade(NomeDestaRotina + '\n' + mensagemRetorno);
				return false;
			}
			finally
			{
				if (strLogNF.Length > 0)
				{
					Global.gravaLogAtividade("NFe com PDF transferido com sucesso: " + strLogNF);
				}

				if (strLogNfSemDadosPdf.Length > 0)
				{
					Global.gravaLogAtividade("NFe sem dados de PDF: " + strLogNfSemDadosPdf);
				}
			}

			return true;
		}
		#endregion

		#region [ executa_download_pdf_danfe_parametro_emitente ]
		public static bool executa_download_pdf_danfe_parametro_emitente(String emitenteNFe, String numeroNFe, String serieNFe, ref String mensagemRetorno)
		{
			#region [ Declarações ]
			const String NomeDestaRotina = "executa_download_pdf_danfe_parametro_emitente()";
			const String BD_OLEDB_PROVIDER = "SQLOLEDB";
			String s;
			String s_aux;
			String strLogNF = "";
			String strLogNfSemDadosPdf = "";
			String strDiretorioPdfDanfe;
			String strNomeArqDanfe;
			String strNomeArqCompletoDanfe;
			String strNumeroNfNormalizado;
			String strSerieNfNormalizado;
			String strNomeEmitente = "";
			String strNfeT1ServidorBd;
			String strNfeT1NomeBd;
			String strNfeT1UsuarioBd;
			String strNfeT1SenhaCriptografadaBd;
			int id_nfe_emitente;
			int id_nfe_emitente_anterior;
			long intQtdeArqDownload = 0;
			long intQtdeNFeSemDadosPdf = 0;
			long intContadorRegistros;
			long intQtdeTotalRegistros;
			int intNfeRetornoSP;
			long lngFileSize;
			byte[] bytFile;
			int i;

			// BANCO DE DADOS
			OleDbConnection dbcNFe = null;
			DataTable t_NFE_EMITENTE = new DataTable();
			DataTable t_NFe_EMISSAO = new DataTable();
			OleDbCommand cmdNFeSituacao = new OleDbCommand();
			OleDbCommand cmdNFeDanfe = new OleDbCommand();
			DataTable rsNFeRetornoSPSituacao = new DataTable();
			DataTable rsNFeRetornoSPDanfe = new DataTable();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataRow rowConsultaEmissao;
			DataRow rowConsultaBoleto;
			OleDbDataReader readerNFeRetornoSPSituacao;
			OleDbDataReader readerNFeRetornoSPDanfe;
			#endregion

			try
			{
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();

				cmdNFeSituacao.CommandType = CommandType.StoredProcedure;
				cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao";
				cmdNFeSituacao.Parameters.Add("NFe", OleDbType.Char, 9);
				cmdNFeSituacao.Parameters.Add("Serie", OleDbType.Char, 3);

				cmdNFeDanfe.CommandType = CommandType.StoredProcedure;
				cmdNFeDanfe.CommandText = "Proc_NFe_Danfe";
				cmdNFeDanfe.Parameters.Add("NFe", OleDbType.Char, 9);
				cmdNFeDanfe.Parameters.Add("Serie", OleDbType.Char, 3);

				s = "SELECT DISTINCT" +
						" id_nfe_emitente," +
						" id," +
						" NFe_serie_NF," +
						" NFe_numero_NF" +
					" FROM t_NFe_EMISSAO" +
					" WHERE" +
						" (id_nfe_emitente = " + emitenteNFe + ")" +
						" AND (NFe_numero_NF = '" + numeroNFe + "')" +
						" AND (NFe_serie_NF = '" + serieNFe + "')" +
					" ORDER BY" +
						" id_nfe_emitente," +
						" id DESC," +
						" NFe_serie_NF," +
						" NFe_numero_NF";
				cmCommand.CommandText = s;
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daAdapter.Fill(t_NFe_EMISSAO);

				if (t_NFe_EMISSAO.Rows.Count == 0)
				{
					mensagemRetorno = "NFe " + numeroNFe + " (série " + serieNFe + ") não localizada!!";
					return false;
				}

				id_nfe_emitente_anterior = -1;
				intContadorRegistros = 0;
				intQtdeTotalRegistros = t_NFe_EMISSAO.Rows.Count;
				for (i = 0; i < t_NFe_EMISSAO.Rows.Count; i++)
				{
					rowConsultaEmissao = t_NFe_EMISSAO.Rows[i];
					intContadorRegistros = intContadorRegistros + 1;
					id_nfe_emitente = BD.readToInt(rowConsultaEmissao["id_nfe_emitente"]);
					if (id_nfe_emitente != id_nfe_emitente_anterior)
					{
						s = "SELECT" +
								" razao_social," +
								" NFe_T1_servidor_BD," +
								" NFe_T1_nome_BD," +
								" NFe_T1_usuario_BD," +
								" NFe_T1_senha_BD" +
							" FROM t_NFE_EMITENTE" +
							" WHERE" +
								" (id = " + id_nfe_emitente.ToString() + ")";
						t_NFE_EMITENTE.Clear();
						cmCommand.CommandText = s;
						daAdapter.SelectCommand = cmCommand;
						daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
						daAdapter.Fill(t_NFE_EMITENTE);
						if (t_NFE_EMITENTE.Rows.Count == 0)
						{
							mensagemRetorno = "Falha ao localizar o registro em t_NFE_EMITENTE (id=" + id_nfe_emitente.ToString() + ")!!";
							return false;
						}

						rowConsultaBoleto = t_NFE_EMITENTE.Rows[0];
						strNomeEmitente = BD.readToString(rowConsultaBoleto["razao_social"]).ToUpper();
						strNfeT1ServidorBd = BD.readToString(rowConsultaBoleto["NFe_T1_servidor_BD"]);
						strNfeT1NomeBd = BD.readToString(rowConsultaBoleto["NFe_T1_nome_BD"]);
						strNfeT1UsuarioBd = BD.readToString(rowConsultaBoleto["NFe_T1_usuario_BD"]);
						strNfeT1SenhaCriptografadaBd = BD.readToString(rowConsultaBoleto["NFe_T1_senha_BD"]);

						s_aux = "";
						CriptoHex.decodificaDado(strNfeT1SenhaCriptografadaBd, ref s_aux);
						s = "Provider=" + BD_OLEDB_PROVIDER +
							";Data Source=" + strNfeT1ServidorBd +
							";Initial Catalog=" + strNfeT1NomeBd +
							";User Id=" + strNfeT1UsuarioBd +
							";Password=" + s_aux;

						id_nfe_emitente_anterior = id_nfe_emitente;

						if (dbcNFe != null)
						{
							if (dbcNFe.State != ConnectionState.Closed) dbcNFe.Close();
						}
						dbcNFe = new OleDbConnection();
						dbcNFe.ConnectionString = s;
						dbcNFe.Open();
					}

					strNumeroNfNormalizado = NFeFormataNumeroNF(BD.readToString(rowConsultaEmissao["NFe_numero_NF"]));
					strSerieNfNormalizado = NFeFormataSerieNF(BD.readToString(rowConsultaEmissao["NFe_serie_NF"]));

					//'   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRAÇÃO C/ O SISTEMA DE NFe DA TARGET ONE
					cmdNFeSituacao.Connection = dbcNFe;
					cmdNFeDanfe.Connection = dbcNFe;

					cmdNFeSituacao.Parameters["NFe"].Value = strNumeroNfNormalizado;
					cmdNFeSituacao.Parameters["Serie"].Value = strSerieNfNormalizado;

					readerNFeRetornoSPSituacao = cmdNFeSituacao.ExecuteReader();
					try
					{
						intNfeRetornoSP = 0;
						if (readerNFeRetornoSPSituacao.Read()) intNfeRetornoSP = (byte)readerNFeRetornoSPSituacao[0];
					}
					finally
					{
						readerNFeRetornoSPSituacao.Close();
					}

					//se a DANFE não for localizada no banco de dados, retornar false
					//(controle na hora da impressão no programa)
					if (intNfeRetornoSP != 1)
					{
						mensagemRetorno = "NFe nº " + numeroNFe + " (série " + serieNFe + ") não localizada.";
						return false;
					}
					else
					{
						cmdNFeDanfe.Parameters["NFe"].Value = strNumeroNfNormalizado;
						cmdNFeDanfe.Parameters["Serie"].Value = strSerieNfNormalizado;

						readerNFeRetornoSPDanfe = cmdNFeDanfe.ExecuteReader();
						try
						{
							while (readerNFeRetornoSPDanfe.Read())
							{
								lngFileSize = readerNFeRetornoSPDanfe.GetBytes(0, 0, null, 0, 0);

								if (lngFileSize <= 0)
								{
									intQtdeNFeSemDadosPdf = intQtdeNFeSemDadosPdf + 1;
									if (strLogNfSemDadosPdf != "") strLogNfSemDadosPdf = strLogNfSemDadosPdf + ", ";
									strLogNfSemDadosPdf = strLogNfSemDadosPdf + id_nfe_emitente.ToString() + "/" + strSerieNfNormalizado + "/" + strNumeroNfNormalizado;
								}

								if (lngFileSize > 0) intQtdeArqDownload = intQtdeArqDownload + 1;

								//'   LOG
								if (strLogNF != "") strLogNF = strLogNF + ", ";
								strLogNF = strLogNF + id_nfe_emitente.ToString() + "/" + strSerieNfNormalizado + "/" + strNumeroNfNormalizado;

								//'   ARQUIVO DE DANFE
								strNomeArqDanfe = "NFe_" + strSerieNfNormalizado + "_" + strNumeroNfNormalizado + ".pdf";
								strDiretorioPdfDanfe = Global.Cte.Etc.PathManipulaPDF;

								if (!Directory.Exists(strDiretorioPdfDanfe))
								{
									try
									{
										Directory.CreateDirectory(strDiretorioPdfDanfe);
									}
									catch
									{
										mensagemRetorno = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" + strDiretorioPdfDanfe + ")";
										return false;
									}
								}

								strNomeArqCompletoDanfe = strDiretorioPdfDanfe + "\\" + strNomeArqDanfe;
								if (!File.Exists(strNomeArqCompletoDanfe))
								{
									try
									{
										File.Delete(strNomeArqCompletoDanfe);
									}
									catch (Exception)
									{
										mensagemRetorno = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" + strNomeArqCompletoDanfe + ")";
										return false;
									}
								}

								try
								{
									bytFile = new Byte[lngFileSize];
									readerNFeRetornoSPDanfe.GetBytes(0, 0, bytFile, 0, (int)lngFileSize);
									File.WriteAllBytes(strNomeArqCompletoDanfe, bytFile);
								}
								catch (Exception)
								{
									mensagemRetorno = "Falha ao tentar gravar o arquivo PDF (" + strNomeArqCompletoDanfe + ")";
									return false;
								}
							}  // while (readerNFeRetornoSPDanfe.Read())
						}
						finally
						{
							readerNFeRetornoSPDanfe.Close();
						}
					}
				}  // for
			}
			catch (Exception ex)
			{
				mensagemRetorno = ex.ToString();
				Global.gravaLogAtividade(NomeDestaRotina + '\n' + mensagemRetorno);
				return false;
			}
			finally
			{
				if (strLogNF.Length > 0)
				{
					Global.gravaLogAtividade("NFe com PDF transferido com sucesso: " + strLogNF);
				}

				if (strLogNfSemDadosPdf.Length > 0)
				{
					Global.gravaLogAtividade("NFe sem dados de PDF: " + strLogNfSemDadosPdf);
				}
			}

			return true;
		}
		#endregion

	}
}
