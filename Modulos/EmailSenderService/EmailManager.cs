using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.VisualBasic;
using System.Net.Mail;

namespace EmailSenderService
{
	class EmailManager
	{
		#region[ Atributos ]
		#region[ Parâmetros de Envio de Mensagens ]
		public static int intervaloMinEmSegundosEntreMsgs;
		public static int intervaloEmSegundosAposCicloOcioso;
		public static int intervaloMinEmSegundos_Tentativa_1_2;
		public static int intervaloMinEmSegundos_Tentativa_2_3;
		public static int intervaloMinEmSegundos_Tentativa_Demais;
		public static int qtdeMaxTentativas;
		public static String periodoSuspensao;
		public static int flagRotinaEnvioEmailHabilitado;
		public static bool blnAguardarPeriodoOcioso;
		public static DateTime dtInicioPeriodoOcioso;
		#endregion
		private static HashSet<Int32> idsRemetentesAptos;
		private static HashSet<Int32> idsMensagensSelecionadas;
		#endregion

		#region [ Construtor estático ]
		static EmailManager()
		{
			obtemParamDeEnvioDeMensagem();
			blnAguardarPeriodoOcioso = false;
		}
		#endregion

		#region [ Métodos ]
		public static void processaMensagens() //Entry Point da Rotina
		{
			#region [ Declaração ]
			const int HABILITADO = 1;
			#endregion

			if (flagRotinaEnvioEmailHabilitado == HABILITADO)
			{

				//enquanto estiver no período de ociosidade, não entrar no ciclo de tratamento de mensagens
				if (blnAguardarPeriodoOcioso)
				{
					long duracaoOciosidade = DateAndTime.DateDiff(DateInterval.Second, dtInicioPeriodoOcioso, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);
					if (duracaoOciosidade > intervaloEmSegundosAposCicloOcioso)
					{
						blnAguardarPeriodoOcioso = false;
						dtInicioPeriodoOcioso = DateTime.Now;
					}
					else
					{
						return;
					}
				}

				if (!isHorarioManutencaoSistema(periodoSuspensao))
				{
					preparaEnvioDasMensagens();
					if (idsMensagensSelecionadas != null && idsMensagensSelecionadas.Count > 0)
					{
						enviaMensagens();
						dtInicioPeriodoOcioso = DateTime.Now;
					}
					else
					{
						//se não houver mensagem para enviar, verificar se o tempo de envio entre mensagens foi atingido
						//em caso positivo, ativar flag e entrar no ciclo de ociosidade
						long duracaoCiclo = DateAndTime.DateDiff(DateInterval.Second, dtInicioPeriodoOcioso, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);
						if (duracaoCiclo > intervaloMinEmSegundosEntreMsgs) blnAguardarPeriodoOcioso = true;
					}
				}
			}
		}

		private static void preparaEnvioDasMensagens()
		{
			obtemListaIdsRemetentesAptosAEnviarMsg();
			if (idsRemetentesAptos != null && idsRemetentesAptos.Count > 0)
			{
				obtemListaIdsDeMensagensAEnviar();
			}
		}

		private static void obtemParamDeEnvioDeMensagem()
		{
			//Chama DAO e preenche os parametros de envio
			DataSet parametros = ParametroDAO.obtemParamDeEnvioDeMensagem();
			foreach (DataRow row in parametros.Tables["DtbParametro"].Rows)
			{
				String nomeDoParametro = BD.readToString(row["id"]).Trim();
				switch (nomeDoParametro)
				{
					case "EmailSndSvc_IntervaloMinEmSegundosEntreMsgs":
						intervaloMinEmSegundosEntreMsgs = BD.readToInt(row["campo_inteiro"]);
						break;
					case "EmailSndSvc_IntervaloEmSegundosAposCicloOcioso":
						intervaloEmSegundosAposCicloOcioso = BD.readToInt(row["campo_inteiro"]);
						break;
					case "EmailSndSvc_IntervaloMinEmSegundos_Tentativa_1_2":
						intervaloMinEmSegundos_Tentativa_1_2 = BD.readToInt(row["campo_inteiro"]);
						break;
					case "EmailSndSvc_IntervaloMinEmSegundos_Tentativa_2_3":
						intervaloMinEmSegundos_Tentativa_2_3 = BD.readToInt(row["campo_inteiro"]);
						break;
					case "EmailSndSvc_IntervaloMinEmSegundos_Tentativa_Demais":
						intervaloMinEmSegundos_Tentativa_Demais = BD.readToInt(row["campo_inteiro"]);
						break;
					case "EmailSndSvc_QtdeMaxTentativas":
						qtdeMaxTentativas = BD.readToInt(row["campo_inteiro"]);
						break;
					case "EmailSndSvc_PeriodoSuspensao":
						periodoSuspensao = BD.readToString(row["campo_texto"]);
						break;
					case "EmailSndSvc_FlagHabilitacao":
						flagRotinaEnvioEmailHabilitado = BD.readToInt(row["campo_inteiro"]);
						break;
				}
			}
		}

		private static void obtemListaIdsRemetentesAptosAEnviarMsg()
		{
			#region [ Declaração ]
			HashSet<Int32> ids = new HashSet<Int32>();
			#endregion

			//Esses remetentes já estão aptos a mandar mensagem
			DataSet remetentesQueNuncaMandaramEmail = RemetenteDAO.obtemRemetentesQueNuncaMandaramEmail();
			foreach (DataRow row in remetentesQueNuncaMandaramEmail.Tables["dtbRemetente"].Rows)
			{
				ids.Add(BD.readToInt(row["id"]));
			}

			//O remetente será escolhido se o intervalo de envio entre as mensagens for maior que intervaloMinEmSegundosEntreMsgs
			DataSet remetentesComSucessoOuFalhaDefinitivaNoEnvioAnterior = RemetenteDAO.obtemRemetentesQueJaEnviaramEmail();
			foreach (DataRow row in remetentesComSucessoOuFalhaDefinitivaNoEnvioAnterior.Tables["dtbRemetente"].Rows)
			{
				DateTime dtHrUltimaTentativaEnvio = BD.readToDateTime(row["dt_hr_ult_tentativa_envio"]);
				if (dtHrUltimaTentativaEnvio != DateTime.MinValue)
				{
					long duracao = DateAndTime.DateDiff(DateInterval.Second, dtHrUltimaTentativaEnvio, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);
					if (duracao > intervaloMinEmSegundosEntreMsgs)
					{
						ids.Add(BD.readToInt(row["id"]));
					}
				}
			}

			idsRemetentesAptos = ids;
		}

		private static void obtemListaIdsDeMensagensAEnviar()
		{
			#region [ Declaração ]
			HashSet<Int32> idsRemetentes = new HashSet<Int32>(idsRemetentesAptos);
			HashSet<Int32> idsRemetentesUtilizados = new HashSet<Int32>();
			HashSet<Int32> idsMensagensEscolhidos = new HashSet<Int32>();
			#endregion

			#region [ Mensagens novas sem data hora de agendamento ]
			DataSet mensagensNovasSemDataHoraAgendamento = MensagemDAO.obtemMensagensNovasSemDataHoraAgendamento(idsRemetentes);
			foreach (DataRow row in mensagensNovasSemDataHoraAgendamento.Tables["dtbMensagem"].Rows)
			{
				int idMsg = BD.readToInt(row["id"]);
				int idRemetente = BD.readToInt(row["id_remetente"]);

				if (!idsRemetentesUtilizados.Contains(idRemetente))
				{
					idsMensagensEscolhidos.Add(idMsg);
					idsRemetentesUtilizados.Add(idRemetente);
				}
			}
			#endregion

			//Remove ids de remetentes que já possuem mensagem atribuída
			idsRemetentes.ExceptWith(idsRemetentesUtilizados);

			if (idsRemetentes.Count > 0)
			{
				#region [ Mensagens novas com data hora de agendamento ]
				DataSet mensagensNovasComDataHoraAgendamento = MensagemDAO.obtemMensagensNovasComDataHoraAgendamento(idsRemetentes);
				foreach (DataRow row in mensagensNovasComDataHoraAgendamento.Tables["dtbMensagem"].Rows)
				{
					DateTime dtHrAgendamentoEnvio = BD.readToDateTime(row["dt_hr_agendamento_envio"]);
					if (dtHrAgendamentoEnvio <= DateTime.Now)
					{
						int idMsg = BD.readToInt(row["id"]);
						int idRemetente = BD.readToInt(row["id_remetente"]);

						if (!idsRemetentesUtilizados.Contains(idRemetente))
						{
							idsMensagensEscolhidos.Add(idMsg);
							idsRemetentesUtilizados.Add(idRemetente);
						}
					}
				}
				#endregion

				//Remove ids de remetentes que já possuem mensagem atribuída
				idsRemetentes.ExceptWith(idsRemetentesUtilizados);
			}

			if (idsRemetentes.Count > 0)
			{
				#region [ Mensagens que falharam na última tentativa de envio ]
				DataSet mensagensQueFalharam = MensagemDAO.obtemMensagensQueFalharam(idsRemetentes);
				foreach (DataRow row in mensagensQueFalharam.Tables["dtbMensagem"].Rows)
				{
					int idMsg = BD.readToInt(row["id"]);
					int idRemetente = BD.readToInt(row["id_remetente"]);
					int qtdeTentativasRealizadas = BD.readToInt(row["qtde_tentativas_realizadas"]);
					DateTime dtHrUltimaTentativaEnvio = BD.readToDateTime(row["dt_hr_ult_tentativa_envio"]);
					long duracao = 0;

					if (!idsRemetentesUtilizados.Contains(idRemetente))
					{
						if (dtHrUltimaTentativaEnvio != DateTime.MinValue)
						{
							duracao = DateAndTime.DateDiff(DateInterval.Second, dtHrUltimaTentativaEnvio, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);

							if (qtdeTentativasRealizadas == 1)
							{
								if (duracao > intervaloMinEmSegundos_Tentativa_1_2)
								{
									idsMensagensEscolhidos.Add(idMsg);
									idsRemetentesUtilizados.Add(idRemetente);
								}
							}

							if (qtdeTentativasRealizadas == 2)
							{
								if (duracao > intervaloMinEmSegundos_Tentativa_2_3)
								{
									idsMensagensEscolhidos.Add(idMsg);
									idsRemetentesUtilizados.Add(idRemetente);
								}
							}

							if (qtdeTentativasRealizadas > 2)
							{
								if (duracao > intervaloMinEmSegundos_Tentativa_Demais)
								{
									idsMensagensEscolhidos.Add(idMsg);
									idsRemetentesUtilizados.Add(idRemetente);
								}
							}
						}
					}
				}
				#endregion
			}

			idsMensagensSelecionadas = idsMensagensEscolhidos;
		}

		private static void enviaMensagens()
		{
			#region [ Declaração ]
			#region [ Constantes ]
			const int ST_ENVIADO_SUCESSO_TRUE = 1;
			const int ST_ENVIADO_SUCESSO_FALSE = 0;
			const int ST_FALHOU_EM_DEFINITIVO_TRUE = 1;
			const int ST_FALHOU_EM_DEFINITIVO_FALSE = 0;
			const String RESULTADO_ULT_TENTATIVA_ENVIO_SUCESSO = "S";
			const String RESULTADO_ULT_TENTATIVA_ENVIO_FD = "FD";
			const String RESULTADO_ULT_TENTATIVA_ENVIO_FALHA = "F";
			const int ST_PROCESSAMENTO_MSG_INICIO = 1;
			const int ST_PROCESSAMENTO_MSG_FIM = 0;
			const int SSL_HABILITADO = 1;
			#endregion

			String strEmailRemetente;
			String strReplyTo;
			String strReplyToMsg;
			String strDisplayNameRemetente;
			String strDestinatarioPara;
			String strDestinatarioCopia;
			String strDestinatarioCopiaOculta;
			String strAssunto;
			String strCorpoMsg;
			String[] v;
			String strUsuarioSmtp;
			String strSenhaSmtp;
			String strHost;
			String strEmailRemetenteFalha;
			String strListaDestinatariosFalha;
			String strCorpoMsgFalha;
			String strMsgErro = "";
			String strMsg;
			int intPorta;
			int intHabilitaSSL;
			MailMessage mailMensagem;
			SmtpClient smtpCliente;
			MailAddress mailAddressFrom;
			MailAddress mailAddressReplyTo;
			MailAddress mailAddressTo;
			MailAddress mailAddressCc;
			MailAddress mailAddressBcc;
			int intRemetenteId;
			int intMsgId;
			int intQtdeTentativasRealizadas;
			int idMsgFalha;
			int intUsarReplyToMsg;
			bool blnSucesso;
			bool blnGerouNSU;
			bool blnFalhouEmDefinitivo;
			int intNSU;
			#endregion

			DataSet mensagens = MensagemDAO.obtemMensagensParaEnvio(idsMensagensSelecionadas, idsRemetentesAptos);
			if (mensagens.Tables.Count == 0) return;
			foreach (DataRow row in mensagens.Tables["dtbMensagem"].Rows)
			{
				#region [ Obtém campos da mensagem ]
				intMsgId = BD.readToInt(row["id"]);
				intQtdeTentativasRealizadas = BD.readToInt(row["qtde_tentativas_realizadas"]);
				strAssunto = BD.readToString(row["assunto"]);
				strCorpoMsg = BD.readToString(row["corpo_mensagem"]);
				#endregion

				#region [ Obtém dados do remetente ]
				intRemetenteId = BD.readToInt(row.GetParentRow("dtbRemetente_dtbMensagem")["id"]);
				strEmailRemetente = BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["email_remetente"]).Trim();
				strReplyTo = BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["replyTo"]).Trim();
				strReplyToMsg = BD.readToString(row["replyToMsg"]).Trim();
				strDisplayNameRemetente = BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["display_name_remetente"]).Trim();
				strDestinatarioPara = BD.readToString(row["destinatario_To"]).Trim();
				strDestinatarioCopia = BD.readToString(row["destinatario_Cc"]).Trim();
				strDestinatarioCopiaOculta = BD.readToString(row["destinatario_CCo"]).Trim();
				strHost = BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["servidor_smtp"]).Trim();
				intPorta = Convert.ToInt32(BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["servidor_smtp_porta"]));
				intHabilitaSSL = Convert.ToInt32(BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["st_habilita_ssl"]));
				intUsarReplyToMsg = Convert.ToInt32(BD.readToString(row["st_replyToMsg"]));
				strUsuarioSmtp = BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["usuario_smtp"]).Trim();
				strSenhaSmtp = Criptografia.Descriptografa(BD.readToString(row.GetParentRow("dtbRemetente_dtbMensagem")["senha_smtp"]).Trim());
				#endregion

				#region [ Seta flag indicando o início do processamento ]
				if (!MensagemDAO.atualizaStatusProcessamentoMensagem(ST_PROCESSAMENTO_MSG_INICIO, intMsgId))
				{
					//grava log em disco
					Global.gravaLogAtividade("Não foi possível setar a flag st_processamento_mensagem do registro " +
									"com ID: " + intMsgId + " da tabela T_EMAILSNDSVC_MENSAGEM. Processamento desse registro " +
									"e consequente envio da mensagem ignorado.");
					//pula o processamento da mensagem corrente
					continue;
				}
                #endregion

				try
				{
					#region [ Prepara o e-mail ]
					mailMensagem = new MailMessage();
					smtpCliente = new SmtpClient();

					#region [ Preenche remetente ]
					if (strDisplayNameRemetente.Length == 0)
					{
						mailAddressFrom = new MailAddress(strEmailRemetente);
					}
					else
					{
						mailAddressFrom = new MailAddress(strEmailRemetente, strDisplayNameRemetente);
					}

					mailMensagem.From = mailAddressFrom;
					#endregion


					#region[ Se houver, preenche ReplyTo ]
					if (intUsarReplyToMsg == EmailSndSvcDAO.USAR_REPLYTOMSG_SIM)
					{
						if (strReplyToMsg.Length > 0)
						{
							mailAddressReplyTo = new MailAddress(strReplyToMsg);
							mailMensagem.ReplyToList.Clear();
							mailMensagem.ReplyToList.Add(mailAddressReplyTo);
						}
					}
					else
					{
						if (strReplyTo.Length > 0)
						{
							mailAddressReplyTo = new MailAddress(strReplyTo);
							mailMensagem.ReplyToList.Clear();
							mailMensagem.ReplyToList.Add(mailAddressReplyTo);
						}
					}
					#endregion


					#region[ Preenche destinatário ]
					if (strDestinatarioPara.Length > 0)
					{
						strDestinatarioPara = strDestinatarioPara.Replace("\n", " ");
						strDestinatarioPara = strDestinatarioPara.Replace("\r", " ");
						strDestinatarioPara = strDestinatarioPara.Replace(",", " ");
						strDestinatarioPara = strDestinatarioPara.Replace(";", " ");
						v = strDestinatarioPara.Split(' ');
						for (int i = 0; i < v.Length; i++)
						{
							if (v[i].Trim().Length > 0)
							{
								mailAddressTo = new MailAddress(v[i].Trim());
								mailMensagem.To.Add(mailAddressTo);
							}
						}
					}
					#endregion

					#region [ Se houver Cópia Para, preenche campos ]
					if (strDestinatarioCopia.Length > 0)
					{
						strDestinatarioCopia = strDestinatarioCopia.Replace("\n", " ");
						strDestinatarioCopia = strDestinatarioCopia.Replace("\r", " ");
						strDestinatarioCopia = strDestinatarioCopia.Replace(",", " ");
						strDestinatarioCopia = strDestinatarioCopia.Replace(";", " ");
						v = strDestinatarioCopia.Split(' ');
						for (int i = 0; i < v.Length; i++)
						{
							if (v[i].Trim().Length > 0)
							{
								mailAddressCc = new MailAddress(v[i].Trim());
								mailMensagem.CC.Add(mailAddressCc);
							}
						}
					}
					#endregion

					#region [ Se houver Cópia Oculta, preenche campos ]
					if (strDestinatarioCopiaOculta.Length > 0)
					{
						strDestinatarioCopiaOculta = strDestinatarioCopiaOculta.Replace("\n", " ");
						strDestinatarioCopiaOculta = strDestinatarioCopiaOculta.Replace("\r", " ");
						strDestinatarioCopiaOculta = strDestinatarioCopiaOculta.Replace(",", " ");
						strDestinatarioCopiaOculta = strDestinatarioCopiaOculta.Replace(";", " ");
						v = strDestinatarioCopiaOculta.Split(' ');
						for (int i = 0; i < v.Length; i++)
						{
							if (v[i].Trim().Length > 0)
							{
								mailAddressBcc = new MailAddress(v[i].Trim());
								mailMensagem.Bcc.Add(mailAddressBcc);
							}
						}
					}
					#endregion

					#region [ Preenche o assunto e o corpo do email ]
					mailMensagem.Subject = strAssunto;
					mailMensagem.Body = strCorpoMsg;
					#endregion

					#endregion

					#region [ Transmite o e-mail ]
					smtpCliente.Host = strHost;
					if (intPorta > 0) smtpCliente.Port = intPorta;
					if (intHabilitaSSL == SSL_HABILITADO) smtpCliente.EnableSsl = true;
					smtpCliente.Credentials = new System.Net.NetworkCredential(strUsuarioSmtp, strSenhaSmtp);
					smtpCliente.Send(mailMensagem);
					#endregion

					strMsg = "Mensagem enviada com sucesso: t_EMAILSNDSVC_MENSAGEM.id=" + intMsgId.ToString() +
							", remetente=" + strEmailRemetente +
							", destinatário (TO)=" + (strDestinatarioPara.Length > 0 ? strDestinatarioPara : "(nenhum)") +
							", destinatário (CC)=" + (strDestinatarioCopia.Length > 0 ? strDestinatarioCopia : "(nenhum)") +
							", destinatário (BCC)=" + (strDestinatarioCopiaOculta.Length > 0 ? strDestinatarioCopiaOculta : "(nenhum)") +
							"\r\nSubject:" +
							"\r\n" + strAssunto +
							"\r\n" +
							"\r\nBody:" +
							"\r\n" + strCorpoMsg +
							"\r\n" +
							"\r\n";
					Global.gravaLogAtividade(strMsg);

					blnSucesso = true;
				}
				catch (Exception e)
				{
					strMsgErro = "Ocorreu um erro ao tentar enviar a mensagem associada ao registro com ID: " +
									intMsgId + " da tabela T_EMAILSNDSVC_MENSAGEM." +
									"\r\n" + e.ToString() +
									"\r\nPilha de exceção:" +
									"\r\n" + e.StackTrace;
					//grava log em disco
					Global.gravaLogAtividade(strMsgErro);
					blnSucesso = false;
				}

				try
				{
					if (blnSucesso)
					{
						#region [ Sucesso no envio da mensagem ]
						if (!MensagemDAO.atualizaStatusDaMensagem(intQtdeTentativasRealizadas + 1,
																	ST_ENVIADO_SUCESSO_TRUE,
																	DateTime.Now,
																	ST_FALHOU_EM_DEFINITIVO_FALSE,
																	DateTime.MinValue,
																	RESULTADO_ULT_TENTATIVA_ENVIO_SUCESSO,
																	DateTime.Now,
																	null,
																	intMsgId))
						{
							throw new Exception("Houve uma falha ao tentar atualizar os campos da mensagem ID: " + intMsgId);
						}

						if (!RemetenteDAO.atualizaStatusDoRemetente(RESULTADO_ULT_TENTATIVA_ENVIO_SUCESSO,
																	DateTime.Now,
																	intMsgId,
																	intRemetenteId))
						{
							throw new Exception("Houve uma falha ao tentar atualizar os campos do remetente ID: " + intRemetenteId);
						}

						blnGerouNSU = false;
						blnGerouNSU = BD.geraNsu(Global.Cte.Nsu.T_EMAILSNDSVC_LOG, out intNSU, out strMsgErro);
						if (!blnGerouNSU)
						{
							throw new Exception("Houve uma falha ao tentar gerar o NSU da tabela T_EMAILSNDSVC_LOG");
						}

						if (!LogDAO.insereLog(intNSU,
												intMsgId,
												DateTime.Now.Date,
												DateTime.Now,
												RESULTADO_ULT_TENTATIVA_ENVIO_SUCESSO,
												null))
						{
							throw new Exception("Houve uma falha ao tentar criar um registro da tabela T_EMAILSNDSVC_LOG");
						}
						#endregion
					}
					else
					{
						#region [ Falha no envio da mensagem ]
						blnFalhouEmDefinitivo = intQtdeTentativasRealizadas + 1 >= qtdeMaxTentativas ? true : false;

						if (!MensagemDAO.atualizaStatusDaMensagem(intQtdeTentativasRealizadas + 1,
																	ST_ENVIADO_SUCESSO_FALSE,
																	DateTime.MinValue,
																	blnFalhouEmDefinitivo ? ST_FALHOU_EM_DEFINITIVO_TRUE : ST_FALHOU_EM_DEFINITIVO_FALSE,
																	blnFalhouEmDefinitivo ? DateTime.Now : DateTime.MinValue,
																	blnFalhouEmDefinitivo ? RESULTADO_ULT_TENTATIVA_ENVIO_FD : RESULTADO_ULT_TENTATIVA_ENVIO_FALHA,
																	DateTime.Now,
																	strMsgErro,
																	intMsgId))
						{
							throw new Exception("Houve uma falha ao tentar atualizar os campos da mensagem ID: " + intMsgId);
						}

						if (!RemetenteDAO.atualizaStatusDoRemetente(blnFalhouEmDefinitivo ? RESULTADO_ULT_TENTATIVA_ENVIO_FD : RESULTADO_ULT_TENTATIVA_ENVIO_FALHA,
																	DateTime.Now,
																	intMsgId,
																	intRemetenteId))
						{
							throw new Exception("Houve uma falha ao tentar atualizar os campos do remetente ID: " + intRemetenteId);
						}

						blnGerouNSU = false;
						blnGerouNSU = BD.geraNsu(Global.Cte.Nsu.T_EMAILSNDSVC_LOG, out intNSU, out strMsgErro);
						if (!blnGerouNSU)
						{
							throw new Exception("Houve uma falha ao tentar gerar o NSU da tabela T_EMAILSNDSVC_LOG");
						}

						if (!LogDAO.insereLog(intNSU,
												intMsgId,
												DateTime.Now.Date,
												DateTime.Now,
												blnFalhouEmDefinitivo ? RESULTADO_ULT_TENTATIVA_ENVIO_FD : RESULTADO_ULT_TENTATIVA_ENVIO_FALHA,
												strMsgErro))
						{
							throw new Exception("Houve uma falha ao tentar criar um registro da tabela T_EMAILSNDSVC_LOG");
						}

						if (blnFalhouEmDefinitivo)
						{
							blnGerouNSU = false;
							blnGerouNSU = BD.geraNsu(Global.Cte.Nsu.T_EMAILSNDSVC_LOG_ERRO, out intNSU, out strMsgErro);
							if (!blnGerouNSU)
							{
								throw new Exception("Houve uma falha ao tentar gerar o NSU da tabela T_EMAILSNDSVC_LOG_ERRO");
							}

							if (!LogErroDAO.insereLogErro(intNSU,
															intMsgId,
															DateTime.Now.Date,
															DateTime.Now,
															strMsgErro))
							{
								throw new Exception("Houve uma falha ao tentar criar um registro da tabela T_EMAILSNDSVC_LOG");
							}

							strEmailRemetenteFalha = RemetenteDAO.emailRemetenteQueEnviaFalhas();
							strListaDestinatariosFalha = MensagemDAO.listaEmailsDestinatariosFalhas();
							if ((strEmailRemetente != strEmailRemetenteFalha) && (strEmailRemetenteFalha != "") && (strListaDestinatariosFalha != ""))
							{
								strCorpoMsgFalha = "Remetente: " + strEmailRemetente + "\r\n" + "\r\n";
								strCorpoMsgFalha = strCorpoMsgFalha + "Servidor SMTP: " + strHost + "\r\n" + "\r\n";
								strCorpoMsgFalha = strCorpoMsgFalha + "Porta: " + intPorta.ToString() + "\r\n" + "\r\n";
								strCorpoMsgFalha = strCorpoMsgFalha + "Usuário SMTP: " + strUsuarioSmtp + "\r\n" + "\r\n";
								if (intHabilitaSSL == EmailSndSvcDAO.ESS_SSL_NAO_HABILITADO)
								{
									strCorpoMsgFalha = strCorpoMsgFalha + "SSL: Não" + "\r\n" + "\r\n";
								}
								else
								{
									strCorpoMsgFalha = strCorpoMsgFalha + "SSL: Sim" + "\r\n" + "\r\n";
								}
								strCorpoMsgFalha = strCorpoMsgFalha + "Para: " + strDestinatarioPara + "\r\n" + "\r\n";
								if (strDestinatarioCopia.Trim() != "") strCorpoMsgFalha = strCorpoMsgFalha + "Com Cópia: " + strDestinatarioCopia + "\r\n" + "\r\n";
								if (strDestinatarioCopiaOculta.Trim() != "") strCorpoMsgFalha = strCorpoMsgFalha + "Com Cópia Oculta: " + strDestinatarioCopiaOculta + "\r\n" + "\r\n";
								strCorpoMsgFalha = strCorpoMsgFalha + "Assunto: " + strAssunto + "\r\n" + "\r\n";
								strCorpoMsgFalha = strCorpoMsgFalha + "Texto: " + strCorpoMsg;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(
										strEmailRemetenteFalha,
										MensagemDAO.listaEmailsDestinatariosFalhas(), //To
										"", // Cc
										"", // Cco
										"Falha no envio de mensagem (ID: " + intMsgId.ToString() + ")",
										strCorpoMsgFalha,
										DateTime.Now,
										"",
										EmailSndSvcDAO.USAR_REPLYTOMSG_NAO,
										out idMsgFalha,
										out strMsgErro))
								{
									throw new Exception("Não foi possível gravar uma comunicação da falha na entrega da mensagem id " + intMsgId.ToString() + ": " + strMsgErro);
								}
							}
						}
						#endregion
					}
				}
				catch (Exception e)
				{
					Global.gravaLogAtividade("Ocorreu um erro ao modificar as tabelas decorrentes do processamento do registro " +
									"da mensagem ID: " + intMsgId + "." +
									"\r\nDescrição do erro:" +
									"\r\n" + e.Message +
									"\r\nPilha de exceção:" +
									"\r\n" + e.StackTrace);
				}

				#region [ Seta flag indicando o término do processamento ]
				if (!MensagemDAO.atualizaStatusProcessamentoMensagem(ST_PROCESSAMENTO_MSG_FIM, intMsgId))
				{
					//grava log em disco
					Global.gravaLogAtividade("Não foi possível setar a flag st_processamento_mensagem do registro " +
									"com ID: " + intMsgId + " da tabela T_EMAILSNDSVC_MENSAGEM para indicar o término do processamento. ");
				}
				#endregion
			}
		}

		private static bool isHorarioManutencaoSistema(String horarioManutencao) //horarioManutencao no formato HH:MM|HH:MM
		{
			#region [ Declaração ]
			String[] v;
			String strDtInicio;
			String strDtTermino;
			DateTime dtInicio;
			DateTime dtTermino;
            #endregion

			if (horarioManutencao.Trim() == "") return false;

			v = horarioManutencao.Split('|');
			strDtInicio = DateTime.Now.ToShortDateString() + " " + v[0].Trim();
			strDtTermino = DateTime.Now.ToShortDateString() + " " + v[1].Trim();

			dtInicio = DateTime.Parse(strDtInicio);
			dtTermino = DateTime.Parse(strDtTermino);

			if (dtTermino < dtInicio)
			{
				if (DateTime.Now <= dtTermino)
				{
					dtInicio = dtInicio.AddDays(-1);
				}
				else
				{
					dtTermino = dtTermino.AddDays(1);
				}
			}

			if (DateTime.Now > dtInicio && DateTime.Now < dtTermino) return true;

			return false;
		}
		#endregion
	}
}
