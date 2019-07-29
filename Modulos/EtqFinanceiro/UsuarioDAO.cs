#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace EtqFinanceiro
{
    class UsuarioDAO
    {
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
        }
        #endregion

        #region [ Construtor ]
        public UsuarioDAO(string usuario, ref List<string> listaOperacoesPermitidas)
        {
            #region [ Declarações ]
            string strIdOperacao;
            SqlCommand cmCommand;
            SqlDataReader drUsuario;
            SqlDataReader drOperacao;
            string strSql;
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
                        //" AND (tipo_operacao='ETQFIN')" +
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
    }
}
