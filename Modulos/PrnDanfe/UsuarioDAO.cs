#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

#endregion

namespace PrnDANFE
{
	class UsuarioDAO
	{
		#region [ Atributos ]

		#region [ SqlCommand ]
		private static SqlCommand cmUsuarioAtualizaFinEmail;
		#endregion

		#region [ Getters / Setters ]

		#region [ cadastrado ]
		private bool _cadastrado;
		public bool cadastrado
		{
			get { return _cadastrado; }
			set { _cadastrado = value; }
		}
		#endregion

		#region [ usuario ]
		private String _usuario;
		public String usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}
		#endregion

		#region [ senhaDescriptografada ]
		private String _senhaDescriptografada;
		public String senhaDescriptografada
		{
			get { return _senhaDescriptografada; }
			set { _senhaDescriptografada = value; }
		}
		#endregion

		#region [ nome ]
		private String _nome;
		public String nome
		{
			get { return _nome; }
			set { _nome = value; }
		}
		#endregion

		#region [ datastamp ]
		private String _datastamp;
		public String datastamp
		{
			get { return _datastamp; }
			set { _datastamp = value; }
		}
		#endregion

		#region [ bloqueado ]
		private bool _bloqueado;
		public bool bloqueado
		{
			get { return _bloqueado; }
			set { _bloqueado = value; }
		}
		#endregion

		#region [ senhaExpirada ]
		private bool _senhaExpirada;
		public bool senhaExpirada
		{
			get { return _senhaExpirada; }
			set { _senhaExpirada = value; }
		}
		#endregion

		#region [ fin_email_remetente ]
		private String _fin_email_remetente;
		public String fin_email_remetente
		{
			get { return _fin_email_remetente; }
			set { _fin_email_remetente = value; }
		}
		#endregion

		#region [ fin_servidor_smtp ]
		private String _fin_servidor_smtp;
		public String fin_servidor_smtp
		{
			get { return _fin_servidor_smtp; }
			set { _fin_servidor_smtp = value; }
		}
		#endregion

		#region [ fin_servidor_smtp_porta ]
		private int _fin_servidor_smtp_porta;
		public int fin_servidor_smtp_porta
		{
			get { return _fin_servidor_smtp_porta; }
			set { _fin_servidor_smtp_porta = value; }
		}
		#endregion

		#region [ fin_usuario_smtp ]
		private String _fin_usuario_smtp;
		public String fin_usuario_smtp
		{
			get { return _fin_usuario_smtp; }
			set { _fin_usuario_smtp = value; }
		}
		#endregion

		#region [ fin_senha_smtp ]
		private String _fin_senha_smtp;
		public String fin_senha_smtp
		{
			get { return _fin_senha_smtp; }
			set { _fin_senha_smtp = value; }
		}
		#endregion

		#region [ fin_display_name_remetente ]
		private String _fin_display_name_remetente;
		public String fin_display_name_remetente
		{
			get { return _fin_display_name_remetente; }
			set { _fin_display_name_remetente = value; }
		}
		#endregion

		#endregion

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

		#region [ Construtor estático ]
		static UsuarioDAO()
		{
			String strSql;

			#region [ cmUsuarioAtualizaFinEmail ]
			strSql = "UPDATE t_USUARIO SET " +
						"fin_email_remetente = @fin_email_remetente, " +
						"fin_display_name_remetente = @fin_display_name_remetente, " +
						"fin_servidor_smtp = @fin_servidor_smtp, " +
						"fin_servidor_smtp_porta = @fin_servidor_smtp_porta, " +
						"fin_usuario_smtp = @fin_usuario_smtp, " +
						"fin_senha_smtp = @fin_senha_smtp " +
					"WHERE " +
						"(usuario = @usuario)";
			cmUsuarioAtualizaFinEmail = BD.criaSqlCommand();
			cmUsuarioAtualizaFinEmail.CommandText = strSql;
			cmUsuarioAtualizaFinEmail.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmUsuarioAtualizaFinEmail.Parameters.Add("@fin_email_remetente", SqlDbType.VarChar, 80);
			cmUsuarioAtualizaFinEmail.Parameters.Add("@fin_display_name_remetente", SqlDbType.VarChar, 80);
			cmUsuarioAtualizaFinEmail.Parameters.Add("@fin_servidor_smtp", SqlDbType.VarChar, 80);
			cmUsuarioAtualizaFinEmail.Parameters.Add("@fin_servidor_smtp_porta", SqlDbType.Int);
			cmUsuarioAtualizaFinEmail.Parameters.Add("@fin_usuario_smtp", SqlDbType.VarChar, 80);
			cmUsuarioAtualizaFinEmail.Parameters.Add("@fin_senha_smtp", SqlDbType.VarChar, 80);
			cmUsuarioAtualizaFinEmail.Prepare();
			#endregion
		}
		#endregion

		#region [ Construtor ]
		public UsuarioDAO(String usuario, ref List<String> listaOperacoesPermitidas)
		{
			#region [ Declarações ]
			String strIdOperacao;
			SqlCommand cmCommand;
			SqlDataReader drUsuario;
			SqlDataReader drOperacao;
			String strSql;
			#endregion

			this._usuario = usuario;

			#region [ Obtém os dados do usuário no BD ]
			strSql = "SELECT " +
						"*" +
					 " FROM t_USUARIO" +
					 " WHERE" +
						" usuario='" + usuario + "'";
			cmCommand = new SqlCommand(strSql, BD.cnConexao);
			cmCommand.CommandTimeout = BD.intCommandTimeoutEmSegundos;
			drUsuario = cmCommand.ExecuteReader();
			try
			{
				if (drUsuario.Read())
				{
					_cadastrado = true;
					// Usa no log a grafia de maiúsculas/minúsculas com que foi cadastrado
					_usuario = drUsuario["usuario"].ToString();
					_nome = drUsuario["nome"].ToString();
					_datastamp = drUsuario["datastamp"].ToString();

					if (drUsuario["bloqueado"].ToString().Equals("0"))
						_bloqueado = false;
					else
						_bloqueado = true;

					if (drUsuario["dt_ult_alteracao_senha"] == DBNull.Value)
						_senhaExpirada = true;
					else
						_senhaExpirada = false;

					if (drUsuario["fin_email_remetente"] == DBNull.Value)
						_fin_email_remetente = "";
					else
						_fin_email_remetente = drUsuario["fin_email_remetente"].ToString();

					if (drUsuario["fin_display_name_remetente"] == DBNull.Value)
						_fin_display_name_remetente = "";
					else
						_fin_display_name_remetente = drUsuario["fin_display_name_remetente"].ToString();

					if (drUsuario["fin_servidor_smtp"] == DBNull.Value)
						_fin_servidor_smtp = "";
					else
						_fin_servidor_smtp = drUsuario["fin_servidor_smtp"].ToString();

					if (drUsuario["fin_servidor_smtp_porta"] == DBNull.Value)
						_fin_servidor_smtp_porta = 0;
					else
						_fin_servidor_smtp_porta = BD.readToInt(drUsuario["fin_servidor_smtp_porta"]);

					if (drUsuario["fin_usuario_smtp"] == DBNull.Value)
						_fin_usuario_smtp = "";
					else
						_fin_usuario_smtp = drUsuario["fin_usuario_smtp"].ToString();

					if (drUsuario["fin_senha_smtp"] == DBNull.Value)
						_fin_senha_smtp = "";
					else
						_fin_senha_smtp = drUsuario["fin_senha_smtp"].ToString();
				}
				else
				{
					_cadastrado = false;
				}
			}
			finally
			{
				drUsuario.Close();
			}
			#endregion

			#region [ Carrega a lista de operações permitidas ]
			strSql = "SELECT DISTINCT" +
						" id_operacao" +
					 " FROM t_PERFIL p" +
						" INNER JOIN t_PERFIL_ITEM i" +
							" ON (p.id=i.id_perfil)" +
						" INNER JOIN t_PERFIL_X_USUARIO u" +
							" ON (p.id=u.id_perfil)" +
						" INNER JOIN t_OPERACAO o" +
							" ON (i.id_operacao=o.id)" +
					 " WHERE" +
						" (usuario='" + usuario + "')" +
						" AND (modulo='CENTR')" +
						" AND (tipo_operacao='CONS')" +
					 " ORDER BY" +
						" id_operacao";
			cmCommand.CommandText = strSql;
			drOperacao = cmCommand.ExecuteReader();
			try
			{
				if (listaOperacoesPermitidas == null) listaOperacoesPermitidas = new List<String>();
				if (listaOperacoesPermitidas.Count > 0) listaOperacoesPermitidas.Clear();
				while (drOperacao.Read())
				{
					strIdOperacao = drOperacao["id_operacao"].ToString().Trim();
					if (strIdOperacao.Length > 0) listaOperacoesPermitidas.Add(strIdOperacao);
				}
			}
			finally
			{
				drOperacao.Close();
			}
			#endregion
		}
		#endregion

		#region [ atualizaFinEmail ]
		/// <summary>
		/// Atualiza os dados no cadastro do usuário referentes aos campos que contém os
		/// parâmetros para o envio de e-mails através do módulo Financeiro.
		/// </summary>
		/// <param name="usuario">Usuário que está realizando a operação</param>
		/// /// <param name="usuarioSelecionado">Usuário cujo cadastro está sendo alterado</param>
		/// <param name="fin_email_remetente">Endereço de e-mail usado para enviar os e-mails</param>
		/// <param name="fin_display_name_remetente">Nome de exibição do remetente ao enviar os e-mails</param>
		/// <param name="fin_servidor_smtp">Endereço do servidor SMTP</param>
		/// <param name="fin_servidor_smtp_porta">Porta do servidor SMTP</param>
		/// <param name="fin_usuario_smtp">Usuário para fazer a autenticação no servidor SMTP</param>
		/// <param name="fin_senha_smtp">Senha para fazer a autenticação no servidor SMTP</param>
		/// <param name="strMsgErro">No caso de erro, retorna a mensagem de erro</param>
		/// <returns>
		/// true: sucesso na atualização dos dados
		/// false: falha na atualização dos dados
		/// </returns>
		public static bool atualizaFinEmail(String usuario,
											String usuarioSelecionado,
											String fin_email_remetente,
											String fin_display_name_remetente,
											String fin_servidor_smtp,
											int fin_servidor_smtp_porta,
											String fin_usuario_smtp,
											String fin_senha_smtp,
											ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza os parâmetros para envio de e-mails";
			String strSenhaCriptografada = "";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Criptografa a senha ]
				strSenhaCriptografada = Criptografia.Criptografa(fin_senha_smtp);
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUsuarioAtualizaFinEmail.Parameters["@usuario"].Value = usuarioSelecionado;
				cmUsuarioAtualizaFinEmail.Parameters["@fin_email_remetente"].Value = fin_email_remetente;
				cmUsuarioAtualizaFinEmail.Parameters["@fin_display_name_remetente"].Value = fin_display_name_remetente;
				cmUsuarioAtualizaFinEmail.Parameters["@fin_servidor_smtp"].Value = fin_servidor_smtp;
				cmUsuarioAtualizaFinEmail.Parameters["@fin_servidor_smtp_porta"].Value = fin_servidor_smtp_porta;
				cmUsuarioAtualizaFinEmail.Parameters["@fin_usuario_smtp"].Value = fin_usuario_smtp;
				cmUsuarioAtualizaFinEmail.Parameters["@fin_senha_smtp"].Value = strSenhaCriptografada;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUsuarioAtualizaFinEmail);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
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

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os parâmetros para envio de e-mails!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion
	}
}
