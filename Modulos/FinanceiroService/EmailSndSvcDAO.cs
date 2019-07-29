using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Data;

namespace FinanceiroService
{
	class EmailSndSvcDAO
	{
		#region [ Constantes ]
		public const int ESS_SSL_NAO_HABILITADO = 0;
		public const int ESS_SSL_HABILITADO = 1;
		public const int ESS_ENVIO_NAO_HABILITADO = 0;
		public const int ESS_ENVIO_HABILITADO = 1;
		public const int ESS_ST_ENVIADO_SUCESSO_TRUE = 1;
		public const int ESS_ST_ENVIADO_SUCESSO_FALSE = 0;
		public const int ESS_ST_FALHOU_EM_DEFINITIVO_TRUE = 1;
		public const int ESS_ST_FALHOU_EM_DEFINITIVO_FALSE = 0;
		public const String ESS_RESULTADO_ULT_TENTATIVA_ENVIO_SUCESSO = "S";
		public const String ESS_RESULTADO_ULT_TENTATIVA_ENVIO_FD = "FD";
		public const String ESS_RESULTADO_ULT_TENTATIVA_ENVIO_FALHA = "F";
		public const int ESS_ST_PROCESSAMENTO_MSG_INICIO = 1;
		public const int ESS_ST_PROCESSAMENTO_MSG_FIM = 0;
		public const int ESS_ST_ENVIO_CANCELADO_TRUE = 1;
		public const int ESS_ST_ENVIO_CANCELADO_FALSE = 0;
		#endregion

		#region [ EssNsu ]
		public class EssNsu
		{
			public const string T_EMAILSNDSVC_MENSAGEM = "T_EMAILSNDSVC_MENSAGEM";
			public const string T_EMAILSNDSVC_REMETENTE = "T_EMAILSNDSVC_REMETENTE";
		}
		#endregion

		#region [ Atributos ]
		private static SqlCommand cmInsereRemetente;
		private static SqlCommand cmAlteraHabilitacaoEnvioRemetente;
		private static SqlCommand cmInsereMensagem;
		private static SqlCommand cmCancelaEnvioMensagem;
		#endregion

		#region [ inicializaConstrutorEstatico ]
		public static void inicializaConstrutorEstatico()
		{
			// NOP
			// 1) The static constructor for a class executes before any instance of the class is created.
			// 2) The static constructor for a class executes before any of the static members for the class are referenced.
			// 3) The static constructor for a class executes after the static field initializers (if any) for the class.
			// 4) The static constructor for a class executes at most one time during a single program instantiation
			// 5) A static constructor does not take access modifiers or have parameters.
			// 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
			// 7) A static constructor cannot be called directly.
			// 8) The user has no control on when the static constructor is executed in the program.
			// 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
		}
		#endregion

		#region [ Construtor Estático ]
		static EmailSndSvcDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmInsereRemetente ]
			strSql = "INSERT INTO T_EMAILSNDSVC_REMETENTE (" +
							"id, " +
							"email_remetente, " +
							"display_name_remetente, " +
							"servidor_smtp, " +
							"servidor_smtp_porta, " +
							"usuario_smtp, " +
							"senha_smtp, " +
							"replyTo, " +
							"st_habilita_ssl, " +
							"st_envio_mensagem_habilitado " +
							") VALUES (" +
							"@id, " +
							"@email_remetente, " +
							"@display_name_remetente, " +
							"@servidor_smtp, " +
							"@servidor_smtp_porta, " +
							"@usuario_smtp, " +
							"@senha_smtp, " +
							"@replyTo, " +
							"@st_habilita_ssl, " +
							"@st_envio_mensagem_habilitado " +
							") ";

			cmInsereRemetente = BD.criaSqlCommand();
			cmInsereRemetente.CommandText = strSql;
			cmInsereRemetente.Parameters.Add("@id", SqlDbType.Int);
			cmInsereRemetente.Parameters.Add("@email_remetente", SqlDbType.VarChar, 80);
			cmInsereRemetente.Parameters.Add("@display_name_remetente", SqlDbType.VarChar, 80);
			cmInsereRemetente.Parameters.Add("@servidor_smtp", SqlDbType.VarChar, 80);
			cmInsereRemetente.Parameters.Add("@servidor_smtp_porta", SqlDbType.VarChar, 5);
			cmInsereRemetente.Parameters.Add("@usuario_smtp", SqlDbType.VarChar, 80);
			cmInsereRemetente.Parameters.Add("@senha_smtp", SqlDbType.VarChar, 160);
			cmInsereRemetente.Parameters.Add("@replyTo", SqlDbType.VarChar, 1024);
			cmInsereRemetente.Parameters.Add("@st_habilita_ssl", SqlDbType.Int);
			cmInsereRemetente.Parameters.Add("@st_envio_mensagem_habilitado", SqlDbType.Int);
			cmInsereRemetente.Prepare();
			#endregion

			#region [ cmAlteraHabilitacaoEnvioRemetente ]
			strSql = "UPDATE T_EMAILSNDSVC_REMETENTE " +
						"SET st_envio_mensagem_habilitado = @st_envio_mensagem_habilitado " +
						"WHERE id = @id ";

			cmAlteraHabilitacaoEnvioRemetente = BD.criaSqlCommand();
			cmAlteraHabilitacaoEnvioRemetente.CommandText = strSql;
			cmAlteraHabilitacaoEnvioRemetente.Parameters.Add("@id", SqlDbType.Int);
			cmAlteraHabilitacaoEnvioRemetente.Parameters.Add("@st_envio_mensagem_habilitado", SqlDbType.Int);
			cmAlteraHabilitacaoEnvioRemetente.Prepare();
			#endregion

			#region [ cmInsereMensagem ]
			strSql = "INSERT INTO T_EMAILSNDSVC_MENSAGEM (" +
							"id, " +
							"id_remetente, " +
							"dt_cadastro, " +
							"dt_hr_cadastro, " +
							"assunto, " +
							"corpo_mensagem, " +
							"destinatario_To, " +
							"destinatario_Cc, " +
							"destinatario_Cco, " +
							"dt_hr_agendamento_envio " +
							") VALUES (" +
							"@id, " +
							"@id_remetente, " +
							"convert(datetime, convert(varchar(10),getdate(), 121), 121), " +
							"getdate(), " +
							"@assunto, " +
							"@corpo_mensagem, " +
							"@destinatario_To, " +
							"@destinatario_Cc, " +
							"@destinatario_Cco, " +
							"@dt_hr_agendamento_envio " +
							") ";

			cmInsereMensagem = BD.criaSqlCommand();
			cmInsereMensagem.CommandText = strSql;
			cmInsereMensagem.Parameters.Add("@id", SqlDbType.Int);
			cmInsereMensagem.Parameters.Add("@id_remetente", SqlDbType.Int);
			cmInsereMensagem.Parameters.Add("@assunto", SqlDbType.VarChar, 240);
			cmInsereMensagem.Parameters.Add("@corpo_mensagem", SqlDbType.VarChar, -1);
			cmInsereMensagem.Parameters.Add("@destinatario_To", SqlDbType.VarChar, 1024);
			cmInsereMensagem.Parameters.Add("@destinatario_Cc", SqlDbType.VarChar, 1024);
			cmInsereMensagem.Parameters.Add("@destinatario_Cco", SqlDbType.VarChar, 1024);
			cmInsereMensagem.Parameters.Add("@dt_hr_agendamento_envio", SqlDbType.DateTime);
			cmInsereMensagem.Prepare();
			#endregion

			#region [ cmCancelaEnvioMensagem ]
			strSql = "UPDATE T_EMAILSNDSVC_MENSAGEM " +
						"SET st_envio_cancelado = @st_envio_cancelado, " +
						"dt_hr_envio_cancelado = getdate(), " +
						"usuario_envio_cancelado = @usuario_envio_cancelado " +
						"WHERE id = @id ";

			cmCancelaEnvioMensagem = BD.criaSqlCommand();
			cmCancelaEnvioMensagem.CommandText = strSql;
			cmCancelaEnvioMensagem.Parameters.Add("@id", SqlDbType.Int);
			cmCancelaEnvioMensagem.Parameters.Add("@st_envio_cancelado", SqlDbType.Int);
			cmCancelaEnvioMensagem.Parameters.Add("@usuario_envio_cancelado", SqlDbType.VarChar, 10);
			cmCancelaEnvioMensagem.Prepare();
			#endregion

		}
		#endregion

		#region [ Métodos Privados ]

		#region [ isEmailValido ]
		/// <summary>
		/// Indica se o e-mail possui sintaxe válida. Se for uma lista de e-mails, testa cada um dos e-mails.
		/// </summary>
		/// <param name="email">
		/// Um ou mais e-mails que devem ser analisados. Os e-mails podem ser separados por espaço em branco,
		/// vírgula ou ponto e vírgula.
		/// </param>
		/// <param name="relacaoEmailInvalido">
		/// Informa os e-mails inválidos separados por espaço em branco.
		/// </param>
		/// <returns>
		/// true: todos os e-mails são válidos
		/// false: um ou mais e-mails inválidos
		/// </returns>
		private static bool isEmailValido(String email, ref String relacaoEmailInvalido)
		{
			string strRegExEmailValidacao = "^([0-9a-zA-Z]([-.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";
			bool blnSucesso;
			int intQtdeEmail = 0;
			String[] v;
			String strEmail;
			Regex rgex = new Regex(strRegExEmailValidacao);

			relacaoEmailInvalido = "";
			if (email == null) return false;
			if (email.Trim().Length == 0) return false;

			blnSucesso = true;
			strEmail = email.Trim();
			strEmail = strEmail.Replace(',', ' ');
			strEmail = strEmail.Replace(';', ' ');
			strEmail = strEmail.Replace("\n", " ");
			strEmail = strEmail.Replace("\r", " ");
			v = strEmail.Split(' ');
			for (int i = 0; i < v.Length; i++)
			{
				if (v[i].Trim().Length > 0)
				{
					intQtdeEmail++;
					if (!rgex.IsMatch(v[i].Trim()))
					{
						if (relacaoEmailInvalido.Length > 0) relacaoEmailInvalido += " ";
						relacaoEmailInvalido += v[i];
						blnSucesso = false;
					}
				}
			}
			if (intQtdeEmail <= 0) return false;
			return blnSucesso;
		}
		#endregion

		#endregion

		#region [ Métodos Públicos ]

		#region [ obtemRemetentePeloEmail ]
		/// <summary>
		/// Obtém as informações de um remetente através do fornecimento do endereço de e-mail.
		/// </summary>
		/// <param name="emailRemetente">
		/// Endereço de e-mail a ser pesquisado.
		/// </param>
		/// <returns>
		/// Dataset: registro(s) com o endereço de e-mail informado.
		/// </returns>
		public static DataSet obtemRemetentePeloEmail(String emailRemetente)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbRemetente = new DataTable("dtbRemetente");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT * " +
						"FROM T_EMAILSNDSVC_REMETENTE " +
						"WHERE email_remetente = '" + emailRemetente.Trim() + "' ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbRemetente);
			dsResultado.Tables.Add(dtbRemetente);

			return dsResultado;
		}
		#endregion

		#region [ gravaRemetente ]
		/// <summary>
		/// Grava as informações de um remetente de e-mail.
		/// </summary>
		/// <param name="email_remetente">
		/// Endereço de e-mail do remetente.
		/// </param>
		/// <param name="display_name_remetente">
		/// Nome que será informado quando o e-mail for gerado (opcional).
		/// </param>
		/// <param name="servidor_smtp">
		/// Endereço do servidor SMTP através do qual a mensagem será enviada.
		/// </param>
		/// <param name="servidor_smtp_porta">
		/// Número da porta utilizada no servidor SMTP.
		/// </param>
		/// <param name="usuario_smtp">
		/// Usuário cadastrado no servidor SMTP, autorizado a efetuar o envio de mensagens.
		/// </param>
		/// <param name="senha_smtp">
		/// Senha não criptografada do usuário SMTP (a senha será criptografada no momento da gravação).
		/// </param>
		/// <param name="replyTo">
		/// Endereço de e-mail para o qual as respostas serão enviadas (opcional).
		/// </param>
		/// <param name="habilita_ssl">
		/// Informa se será utilizado um canal SSL para envio da mensagem (0 = não; 1 = sim).
		/// </param>
		/// <param name="envio_mensagem_habilitado">
		/// Informa se o remetente estará autorizado a enviar mensagens (0 = não; 1 = sim).
		/// </param>
		/// <param name="id_remetente">
		/// Retorna o id do remetente na tabela, caso a gravação tenha ocorrido.
		/// </param>
		/// <param name="msg_erro_grava_rem">
		/// Retorna uma mensagem de erro, caso a gravação não tenha ocorrido.
		/// </param>
		/// <returns>
		/// true: a gravação foi realizada
		/// false: a gravação não foi realizada
		/// </returns>
		public static bool gravaRemetente(String email_remetente,
												String display_name_remetente,
												String servidor_smtp,
												String servidor_smtp_porta,
												String usuario_smtp,
												String senha_smtp,
												String replyTo,
												int habilita_ssl,
												int envio_mensagem_habilitado,
												out int id_remetente,
												out String msg_erro_grava_rem)
		{

			#region [ Declarações ]
			DataSet dsRemetente;
			String strEmailListaErros = "";
			int intRetorno;
			bool blnSucesso = false;
			#endregion

			id_remetente = 0;
			msg_erro_grava_rem = "";
			try
			{
				#region [ Verifica remetente ]
				//verifica se o e-mail do remetente está preenchido
				if (email_remetente.Trim() == "")
				{
					msg_erro_grava_rem = "E-mail do remetente não preenchido";
					return false;
				}
				//verifica se o e-mail do remetente está preenchido com caracteres válidos
				if (!isEmailValido(email_remetente, ref strEmailListaErros))
				{
					msg_erro_grava_rem = "E-mail do remetente preenchido com caracteres inválidos";
					return false;
				}

				dsRemetente = obtemRemetentePeloEmail(email_remetente);
				if (dsRemetente.Tables["dtbRemetente"].Rows.Count > 0)
				{
					msg_erro_grava_rem = "Endereço de e-mail já cadastrado no banco de dados";
					return false;
				}

				#endregion

				#region [ Verifica demais campos ]

				//verifica se o endereço do servidor smtp está preenchido
				if (servidor_smtp.Trim() == "")
				{
					msg_erro_grava_rem = "Servidor SMTP não informado";
					return false;
				}

				//verifica a porta do servidor SMTP está preenchida apenas com números
				if (Global.converteInteiro(servidor_smtp_porta) == 0)
				{
					msg_erro_grava_rem = "Porta do servidor preenchida incorretamente";
					return false;
				}

				//verificar se usuário do e-mail foi preenchido
				if (usuario_smtp.Trim() == "")
				{
					msg_erro_grava_rem = "Não foi fornecido o usuário SMTP do e-mail";
					return false;
				}

				//verifica se o e-mail de retorno está preenchido com caracteres válidos
				if (replyTo.Trim() != "")
				{
					if (!isEmailValido(replyTo, ref strEmailListaErros))
					{
						msg_erro_grava_rem = "E-mail do campo <<Responder para>> preenchido com caracteres inválidos";
						return false;
					}
					else
					{
						if (strEmailListaErros.Trim() != "")
						{
							msg_erro_grava_rem = "Há e-mail do campo <<Responder para>> preenchido com caracteres inválidos: " + strEmailListaErros;
							return false;
						}
					}
				}

				//verificar se a habilitação SSL está com valores corretos
				if ((habilita_ssl != ESS_SSL_NAO_HABILITADO) && (habilita_ssl != ESS_SSL_HABILITADO))
				{
					msg_erro_grava_rem = "Informação inconsistente sobre utilização de SSL";
					return false;
				}

				//verificar se a habilitação para o envio de e-mails está com valores corretos
				if ((envio_mensagem_habilitado != ESS_ENVIO_NAO_HABILITADO) && (envio_mensagem_habilitado != ESS_ENVIO_HABILITADO))
				{
					msg_erro_grava_rem = "Informação inconsistente sobre habilitação para envio de e-mail";
					return false;
				}

				#endregion

				#region [ Transação Gravação ]
				try
				{
					BD.iniciaTransacao();

					if (!BD.geraNsuUsandoTabelaFinControle(EssNsu.T_EMAILSNDSVC_REMETENTE, out id_remetente, out msg_erro_grava_rem))
					{
						msg_erro_grava_rem = "Problema na geração do NSU do remetente";
					}

					#region [ Preencher dados ]
					cmInsereRemetente.Parameters["@id"].Value = id_remetente;
					cmInsereRemetente.Parameters["@email_remetente"].Value = email_remetente;
					if (display_name_remetente.Trim() == "")
					{
						cmInsereRemetente.Parameters["@display_name_remetente"].Value = DBNull.Value;
					}
					else
					{
						cmInsereRemetente.Parameters["@display_name_remetente"].Value = display_name_remetente;
					}
					cmInsereRemetente.Parameters["@servidor_smtp"].Value = servidor_smtp;
					cmInsereRemetente.Parameters["@servidor_smtp_porta"].Value = servidor_smtp_porta;
					cmInsereRemetente.Parameters["@usuario_smtp"].Value = usuario_smtp;
					cmInsereRemetente.Parameters["@senha_smtp"].Value = Criptografia.Criptografa(senha_smtp);
					if (replyTo.Trim() == "")
					{
						cmInsereRemetente.Parameters["@replyTo"].Value = DBNull.Value;
					}
					else
					{
						cmInsereRemetente.Parameters["@replyTo"].Value = replyTo;
					}
					cmInsereRemetente.Parameters["@st_habilita_ssl"].Value = habilita_ssl;
					cmInsereRemetente.Parameters["@st_envio_mensagem_habilitado"].Value = envio_mensagem_habilitado;
					#endregion

					#region [ Tenta alterar o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsereRemetente);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						msg_erro_grava_rem = ex.ToString();
					}

					if (intRetorno == 1)
					{
						blnSucesso = true;
					}
					else
					{
						blnSucesso = false;
					}
					#endregion

				}
				finally
				{
					#region [ Commit / Rollback ]
					if (blnSucesso)
					{
						#region [ Commit ]
						try
						{
							BD.commitTransacao();
						}
						catch (Exception ex)
						{
							blnSucesso = false;
							msg_erro_grava_rem = ex.ToString();
						}
						#endregion
					}
					else
					{
						#region [ Rollback ]
						try
						{
							BD.rollbackTransacao();
						}
						catch (Exception ex)
						{
							msg_erro_grava_rem = ex.ToString();
						}
						#endregion
					}
					#endregion
				}

				if (!blnSucesso)
				{
					msg_erro_grava_rem = "Problema na sequência de gravação do remetente: " + msg_erro_grava_rem;
					return false;
				}

				#endregion

				return true;
			}
			catch (Exception e)
			{
				msg_erro_grava_rem = "Ocorreu um erro durante a gravação do remetente: " + e.Message;
				return false;
			}
		}
		#endregion

		#region [ alteraHabilitacaoEnvioRemetente ]
		/// <summary>
		/// Habilita ou desabilita o envio de mensagens para um remetente.
		/// </summary>
		/// <param name="id_remetente">
		/// A identificação do remetente a ser alterado.
		/// </param>
		/// <param name="envio_mensagem_habilitado">
		/// Campo que determina se o envio de mensagens está habilitado para o remetente (0 = não, 1 = sim).
		/// </param>
		/// <param name="msg_erro_grava_rem">
		/// Retorna uma mensagem de erro, caso a alteração não tenha ocorrido.
		/// </param>
		/// <returns>
		/// true: a alteração da habilitação de envio ocorreu
		/// false: a alteração da habilitação de envio não ocorreu
		/// </returns>
		public static bool alteraHabilitacaoEnvioRemetente(int id_remetente,
												int envio_mensagem_habilitado,
												out String msg_erro_grava_rem)
		{

			#region [ Declarações ]
			int intRetorno;
			bool blnSucesso = false;
			#endregion

			msg_erro_grava_rem = "";
			try
			{

				#region [ Verifica status ]

				//verificar se a habilitação para o envio de e-mails está com valores corretos
				if ((envio_mensagem_habilitado != ESS_ENVIO_NAO_HABILITADO) && (envio_mensagem_habilitado != ESS_ENVIO_HABILITADO))
				{
					msg_erro_grava_rem = "Informação inconsistente sobre habilitação para envio de e-mail";
					return false;
				}

				#endregion

				#region [ Alteração Habilitação ]

				#region [ Preencher dados ]
				cmAlteraHabilitacaoEnvioRemetente.Parameters["@id"].Value = id_remetente;
				cmAlteraHabilitacaoEnvioRemetente.Parameters["@st_envio_mensagem_habilitado"].Value = envio_mensagem_habilitado;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmAlteraHabilitacaoEnvioRemetente);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					msg_erro_grava_rem = ex.ToString();
				}

				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				if (!blnSucesso)
				{
					msg_erro_grava_rem = "Problema na alteração da habilitação de envio: " + msg_erro_grava_rem;
					return false;
				}

				#endregion

				return true;
			}
			catch (Exception e)
			{
				msg_erro_grava_rem = "Ocorreu um erro durante a alteração da habilitação de envio: " + e.Message;
				return false;
			}
		}
		#endregion

		#region [ gravaMensagemParaEnvio ]
		/// <summary>
		/// Grava as informações de uma mensagem a ser enviada.
		/// </summary>
		/// <param name="email_remetente">
		/// Endereço de e-mail do remetente que enviará a mensagem.
		/// </param>
		/// <param name="destinatario_To">
		/// Um ou mais e-mails de destinatários da mensagem. Os e-mails podem ser separados por espaço em branco,
		/// vírgula ou ponto e vírgula.
		/// </param>
		/// <param name="destinatario_Cc">
		/// Um ou mais e-mails que receberão cópia da mensagem. Os e-mails podem ser separados por espaço em branco,
		/// vírgula ou ponto e vírgula.
		/// </param>
		/// <param name="destinatario_Cco">
		/// Um ou mais e-mails que receberão cópia oculta da mensagem. Os e-mails podem ser separados por espaço em branco,
		/// vírgula ou ponto e vírgula.
		/// </param>
		/// <param name="assunto">
		/// Texto que aparecerá no Subject da mensagem.
		/// </param>
		/// <param name="corpo_mensagem">
		/// Texto com o conteúdo da mensagem.
		/// </param>
		/// <param name="dt_hr_agendamento_envio">
		/// Data e horário quando a mensagem será enviada.
		/// </param>
		/// <param name="id_mensagem">
		/// Retorna o id da mensagem na tabela, caso a gravação tenha ocorrido.
		/// </param>
		/// <param name="msg_erro_grava_msg">
		/// Retorna uma mensagem de erro, caso a gravação não tenha ocorrido.
		/// </param>
		/// <returns>
		/// true: a gravação foi realizada
		/// false: a gravação não foi realizada
		/// </returns>
		public static bool gravaMensagemParaEnvio(String email_remetente,
													String destinatario_To,
													String destinatario_Cc,
													String destinatario_Cco,
													String assunto,
													String corpo_mensagem,
													DateTime dt_hr_agendamento_envio,
													out int id_mensagem,
													out String msg_erro_grava_msg)
		{

			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EmailSndSvcDAO.gravaMensagemParaEnvio()";
			DataSet dsRemetente;
			String strEmailListaErros = "";
			string msg_erro_aux;
			int id_remetente = 0;
			int intRetorno;
			bool existeEmailDeDestino = false;
			bool blnSucesso = false;
			#endregion

			id_mensagem = 0;
			msg_erro_grava_msg = "";
			try
			{
				if (email_remetente == null) email_remetente = "";
				if (destinatario_To == null) destinatario_To = "";
				if (destinatario_Cc == null) destinatario_Cc = "";
				if (destinatario_Cco == null) destinatario_Cco = "";
				if (assunto == null) assunto = "";
				if (corpo_mensagem == null) corpo_mensagem = "";

				#region [ Verifica remetente ]
				//verifica se o e-mail do remetente está preenchido
				if (email_remetente.Trim() == "")
				{
					msg_erro_grava_msg = "E-mail do remetente não preenchido";
					return false;
				}
				//verifica se o e-mail do remetente está preenchido com caracteres válidos
				if (!isEmailValido(email_remetente, ref strEmailListaErros))
				{
					msg_erro_grava_msg = "E-mail do remetente preenchido com caracteres inválidos";
					return false;
				}

				dsRemetente = obtemRemetentePeloEmail(email_remetente);
				foreach (DataRow row in dsRemetente.Tables["dtbRemetente"].Rows)
				{
					if (BD.readToInt(row["st_envio_mensagem_habilitado"]) == ESS_ENVIO_HABILITADO)
					{
						id_remetente = BD.readToInt(row["id"]);
						break;
					}
				}
				if (id_remetente == 0)
				{
					msg_erro_grava_msg = "O remetente informado não está cadastrado ou está com o envio de mensagens desabilitado";
					return false;
				}

				#endregion

				#region [ Verifica mensagem ]

				//verifica se o e-mail do destinatário está preenchido com caracteres válidos
				if (destinatario_To.Trim() != "")
				{
					if (!isEmailValido(destinatario_To, ref strEmailListaErros))
					{
						msg_erro_grava_msg = "E-mail do campo <<Para>> preenchido com caracteres inválidos";
						return false;
					}
					else
					{
						if (strEmailListaErros.Trim() != "")
						{
							msg_erro_grava_msg = "Há e-mail do campo <<Para>> preenchido com caracteres inválidos: " + strEmailListaErros;
							return false;
						}
					}
					existeEmailDeDestino = true;
				}

				//verifica se o e-mail de cópia está preenchido com caracteres válidos
				if (destinatario_Cc.Trim() != "")
				{
					if (!isEmailValido(destinatario_Cc, ref strEmailListaErros))
					{
						msg_erro_grava_msg = "E-mail do campo <<Com cópia>> preenchido com caracteres inválidos";
						return false;
					}
					else
					{
						if (strEmailListaErros.Trim() != "")
						{
							msg_erro_grava_msg = "Há e-mail  do campo <<Com cópia>> preenchido com caracteres inválidos: " + strEmailListaErros;
							return false;
						}
					}
					existeEmailDeDestino = true;
				}

				//verifica se o e-mail de cópia oculta está preenchido com caracteres válidos
				if (destinatario_Cco.Trim() != "")
				{
					if (!isEmailValido(destinatario_Cco, ref strEmailListaErros))
					{
						msg_erro_grava_msg = "E-mail do campo <<Com cópia oculta>> preenchido com caracteres inválidos";
						return false;
					}
					else
					{
						if (strEmailListaErros.Trim() != "")
						{
							msg_erro_grava_msg = "Há e-mail  do campo <<Com cópia oculta>> preenchido com caracteres inválidos: " + strEmailListaErros;
							return false;
						}
					}
					existeEmailDeDestino = true;
				}

				//verificar se algum dos campos de destinatários está preenchido
				if (!existeEmailDeDestino)
				{
					msg_erro_grava_msg = "Não foi fornecido e-mail de nenhum destinatário para a mensagem";
					return false;
				}

				//verificar se o campo de assunto está preenchido
				if (assunto.Trim() == "")
				{
					msg_erro_grava_msg = "O campo Assunto não foi preenchido";
					return false;
				}

				//verificar se o corpo da mensagem está preenchido
				if (corpo_mensagem.Trim() == "")
				{
					msg_erro_grava_msg = "O corpo da mensagem não foi preenchido";
					return false;
				}

				#endregion

				#region [ Transação Gravação ]
				try
				{
					BD.iniciaTransacao();

					if (!BD.geraNsuUsandoTabelaFinControle(EssNsu.T_EMAILSNDSVC_MENSAGEM, out id_mensagem, out msg_erro_grava_msg))
					{
						msg_erro_grava_msg = "Problema na geração do NSU da mensagem";
					}

					#region [ Preencher dados ]
					cmInsereMensagem.Parameters["@id"].Value = id_mensagem;
					cmInsereMensagem.Parameters["@id_remetente"].Value = id_remetente;
					if (destinatario_To.Trim() == "")
					{
						cmInsereMensagem.Parameters["@destinatario_To"].Value = DBNull.Value;
					}
					else
					{
						cmInsereMensagem.Parameters["@destinatario_To"].Value = destinatario_To;
					}
					if (destinatario_Cc.Trim() == "")
					{
						cmInsereMensagem.Parameters["@destinatario_Cc"].Value = DBNull.Value;
					}
					else
					{
						cmInsereMensagem.Parameters["@destinatario_Cc"].Value = destinatario_Cc;
					}
					if (destinatario_Cco.Trim() == "")
					{
						cmInsereMensagem.Parameters["@destinatario_Cco"].Value = DBNull.Value;
					}
					else
					{
						cmInsereMensagem.Parameters["@destinatario_Cco"].Value = destinatario_Cco;
					}
					cmInsereMensagem.Parameters["@assunto"].Value = assunto;
					cmInsereMensagem.Parameters["@corpo_mensagem"].Value = corpo_mensagem;
					if (dt_hr_agendamento_envio == DateTime.MinValue)
					{
						cmInsereMensagem.Parameters["@dt_hr_agendamento_envio"].Value = DBNull.Value;
					}
					else
					{
						cmInsereMensagem.Parameters["@dt_hr_agendamento_envio"].Value = dt_hr_agendamento_envio;
					}
					#endregion

					#region [ Tenta alterar o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsereMensagem);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						msg_erro_grava_msg = ex.ToString();
					}

					if (intRetorno == 1)
					{
						blnSucesso = true;
					}
					else
					{
						blnSucesso = false;
					}
					#endregion

				}
				finally
				{
					#region [ Commit / Rollback ]
					if (blnSucesso)
					{
						#region [ Commit ]
						try
						{
							BD.commitTransacao();
						}
						catch (Exception ex)
						{
							blnSucesso = false;
							msg_erro_grava_msg = ex.ToString();
						}
						#endregion
					}
					else
					{
						#region [ Rollback ]
						try
						{
							BD.rollbackTransacao();
						}
						catch (Exception ex)
						{
							msg_erro_grava_msg = ex.ToString();
						}
						#endregion
					}
					#endregion
				}

				if (!blnSucesso)
				{
					msg_erro_grava_msg = "Problema na sequência de gravação da mensagem: " + msg_erro_grava_msg;
					return false;
				}

				#endregion

				return true;
			}
			catch (Exception e)
			{
				msg_erro_grava_msg = "Ocorreu um erro durante a gravação da mensagem para envio: " + e.Message;
				return false;
			}
			finally
			{
				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = "Gravação de email na fila de mensagens: " + (blnSucesso ? "sucesso" : "falha") + " (id=" + id_mensagem.ToString() + ", dt_hr_agendamento_envio=" + (dt_hr_agendamento_envio == DateTime.MinValue ? "Imediato" : Global.formataDataDdMmYyyyHhMmSsComSeparador(dt_hr_agendamento_envio)) + ")";
				svcLog.complemento_1 = "To: " + (destinatario_To == null ? "(null)" : (destinatario_To == "" ? "(vazio)" : destinatario_To)) +
										"\n" +
										"Cc: " + (destinatario_Cc == null ? "(null)" : (destinatario_Cc == "" ? "(vazio)" : destinatario_Cc)) +
										"\n" +
										"Cco: " + (destinatario_Cco == null ? "(null)" : (destinatario_Cco == "" ? "(vazio)" : destinatario_Cco));
				svcLog.complemento_2 = "Assunto:\n" + assunto;
				svcLog.complemento_3 = "Corpo:\n" + corpo_mensagem;
				if (msg_erro_grava_msg != null)
				{
					if (msg_erro_grava_msg.Length > 0) svcLog.complemento_4 = msg_erro_grava_msg;
				}
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion
			}
		}
		#endregion

		#region [ cancelaEnvioMensagem ]
		/// <summary>
		/// Cancela o envio de uma mensagem.
		/// </summary>
		/// <param name="id_mensagem">
		/// A identificação da mensagem a ser cancelada.
		/// </param>
		/// <param name="usuario_envio_cancelado">
		/// O usuário que efetuou o cancelamento.
		/// </param>
		/// <param name="msg_erro_grava_msg">
		/// Retorna uma mensagem de erro, caso o cancelamento não tenha ocorrido.
		/// </param>
		/// <returns>
		/// true: o cancelamento do envio ocorreu
		/// false: o cancelamento do envio não ocorreu
		/// </returns>
		public static bool cancelaEnvioMensagem(int id_mensagem,
												String usuario_envio_cancelado,
												out String msg_erro_grava_msg)
		{

			#region [ Declarações ]
			int intRetorno;
			bool blnSucesso = false;
			#endregion

			msg_erro_grava_msg = "";
			try
			{

				#region [ Alteração Habilitação ]

				#region [ Preencher dados ]
				cmCancelaEnvioMensagem.Parameters["@id"].Value = id_mensagem;
				cmCancelaEnvioMensagem.Parameters["@st_envio_cancelado"].Value = ESS_ST_ENVIO_CANCELADO_TRUE;
				cmCancelaEnvioMensagem.Parameters["@usuario_envio_cancelado"].Value = usuario_envio_cancelado;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmCancelaEnvioMensagem);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					msg_erro_grava_msg = ex.ToString();
				}

				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				if (!blnSucesso)
				{
					msg_erro_grava_msg = "Problema no cancelamento de envio: " + msg_erro_grava_msg;
					return false;
				}

				#endregion

				return true;
			}
			catch (Exception e)
			{
				msg_erro_grava_msg = "Ocorreu um erro durante o cancelamento de envio: " + e.Message;
				return false;
			}
		}
		#endregion

		#endregion
	}
}
