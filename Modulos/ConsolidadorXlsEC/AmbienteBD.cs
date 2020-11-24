using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	#region [ Classe AmbienteBD ]
	public class AmbienteBD
	{
		#region [ Parâmetros de conexão ]
		public readonly string NomeAmbiente;
		public readonly string EnderecoServidor;
		public readonly string NomeBancoDados;
		public readonly string NomeUsuarioBD;
		public readonly string SenhaUsuarioCriptografada;
		public readonly string NumeroLojaArclube;
		#endregion

		#region [ Atributos ]
		public BancoDados BD;
		public LogDAO logDAO;
		public GeralDAO geralDAO;
		public PedidoDAO pedidoDAO;
		public UsuarioDAO usuarioDAO;
		public ProdutoDAO produtoDAO;
        public ComboDAO comboDAO;
		public List<PercentualCustoFinanceiroFornecedor> tabelaPercCustoFinanceiroFornecedor = null;
		public List<ProdutoLoja> tabelaProdutoLojaOriginal = null;
		#endregion

		#region [ Constructor ]
		public AmbienteBD(string nomeAmbiente, string enderecoServidor, string nomeBancoDados, string nomeUsuarioBD, string senhaUsuarioCriptografada, string numeroLojaArclube)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ConsolidadorXlsEC.AmbienteBD.Constructor()";
			string msg_erro_completo;
			string msg_erro_resumido;
			#endregion

			NomeAmbiente = nomeAmbiente;
			EnderecoServidor = enderecoServidor;
			NomeBancoDados = nomeBancoDados;
			NomeUsuarioBD = nomeUsuarioBD;
			SenhaUsuarioCriptografada = senhaUsuarioCriptografada;
			NumeroLojaArclube = numeroLojaArclube;
			if (!iniciaBancoDados(out msg_erro_completo, out msg_erro_resumido))
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg_erro_completo);
				throw new Exception(msg_erro_completo);
			}
		}
		#endregion

		#region [ iniciaBancoDados ]
		public bool iniciaBancoDados(out string msgErroCompleto, out string msgErroResumido)
		{
			msgErroCompleto = "";
			msgErroResumido = "";

			try
			{
				Global.gravaLogAtividade("Inicializando conexão com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")");

				BD = new BancoDados(NomeAmbiente, EnderecoServidor, NomeBancoDados, NomeUsuarioBD, SenhaUsuarioCriptografada);

				if (!BD.abreConexao(out msgErroCompleto, out msgErroResumido))
				{
					msgErroCompleto = "Falha ao tentar conectar com o banco de dados (Servidor: " + EnderecoServidor + ", BD: " + NomeBancoDados + ")" +
							((msgErroCompleto ?? "").Length > 0 ? "\r\n" + msgErroCompleto : "");
					msgErroResumido = "Falha ao tentar conectar com o banco de dados (Servidor: " + EnderecoServidor + ", BD: " + NomeBancoDados + ")" +
							((msgErroResumido ?? "").Length > 0 ? "\r\n" + msgErroResumido : "");
					Global.gravaLogAtividade("Falha ao inicializar a conexão com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")\r\n" + msgErroCompleto);
					return false;
				}

				if (!inicializaObjetosDAO(out msgErroCompleto, out msgErroResumido))
				{
					msgErroCompleto = "Falha ao tentar iniciar objetos DAO do banco de dados (Servidor: " + EnderecoServidor + ", BD: " + NomeBancoDados + ")" +
							((msgErroCompleto ?? "").Length > 0 ? "\r\n" + msgErroCompleto : "");
					msgErroResumido = "Falha ao tentar iniciar objetos DAO do banco de dados (Servidor: " + EnderecoServidor + ", BD: " + NomeBancoDados + ")" +
							((msgErroResumido ?? "").Length > 0 ? "\r\n" + msgErroResumido : "");
					Global.gravaLogAtividade("Falha ao inicializar a conexão com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")\r\n" + msgErroCompleto);
					return false;
				}

				Global.gravaLogAtividade("Sucesso ao inicializar a conexão com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")");

				return true;
			}
			catch (Exception ex)
			{
				msgErroCompleto = ex.ToString();
				msgErroResumido = ex.Message;
				Global.gravaLogAtividade("Falha ao inicializar a conexão com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")\r\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ reiniciaBancoDados ]
		public bool reiniciaBancoDados(out string msgErroCompleto, out string msgErroResumido)
		{
			msgErroCompleto = "";
			msgErroResumido = "";

			Global.gravaLogAtividade("Início da tentativa de reconectar com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")!!");

			#region [ Tenta fechar a conexão anterior, se houver ]
			try
			{
				if (BD.cnConexao != null)
				{
					if (BD.cnConexao.State != ConnectionState.Closed) BD.cnConexao.Close();
				}
			}
			catch (Exception)
			{
				// NOP
			}
			#endregion

			#region [ Tenta abrir nova conexão ]
			try
			{
				if (!iniciaBancoDados(out msgErroCompleto, out msgErroResumido))
				{
					msgErroCompleto = "Falha ao tentar reconectar com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")" +
									((msgErroCompleto ?? "").Length > 0 ? "\r\n" + msgErroCompleto : "");
					msgErroResumido = "Falha ao tentar reconectar com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")" +
									((msgErroResumido ?? "").Length > 0 ? "\r\n" + msgErroResumido : "");
					return false;
				}

				Global.gravaLogAtividade("Sucesso ao reconectar com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + "): processo concluído!!");

				return true;
			}
			catch (Exception ex)
			{
				msgErroCompleto = ex.ToString();
				msgErroResumido = ex.Message;
				Global.gravaLogAtividade("Falha ao tentar reconectar com o banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")!!\r\n" + ex.ToString());
				return false;
			}
			#endregion
		}
		#endregion

		#region [ inicializaObjetosDAO ]
		public bool inicializaObjetosDAO(out string msgErroCompleto, out string msgErroResumido)
		{
			#region [ Declarações ]
			string msg_erro_aux;
			#endregion

			msgErroCompleto = "";
			msgErroResumido = "";

			try
			{
				Global.gravaLogAtividade("Inicialização dos objetos DAO do banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")");

				logDAO = new LogDAO(ref BD);
				geralDAO = new GeralDAO(ref BD);
				pedidoDAO = new PedidoDAO(ref BD);
				usuarioDAO = new UsuarioDAO(ref BD);
				produtoDAO = new ProdutoDAO(ref BD);
                comboDAO = new ComboDAO(ref BD);
				tabelaPercCustoFinanceiroFornecedor = produtoDAO.GetTabelaPercentualCustoFinanceiroFornecedor(out msg_erro_aux);

				Global.gravaLogAtividade("Sucesso na inicialização dos objetos DAO do banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")");

				return true;
			}
			catch (Exception ex)
			{
				msgErroCompleto = ex.ToString();
				msgErroResumido = ex.Message;

				Global.gravaLogAtividade("Falha na inicialização dos objetos DAO do banco de dados '" + NomeBancoDados + "' (Servidor: " + EnderecoServidor + ")\r\n" + ex.ToString());

				return false;
			}
		}
		#endregion

		#region [ reinicializaObjetosDAO ]
		public bool reinicializaObjetosDAO(out string msgErroCompleto, out string msgErroResumido)
		{
			msgErroCompleto = "";
			msgErroResumido = "";

			if (!inicializaObjetosDAO(out msgErroCompleto, out msgErroResumido))
			{
				msgErroCompleto = "Falha ao tentar reiniciar objetos DAO do banco de dados (Servidor: " + EnderecoServidor + ", BD: " + NomeBancoDados + ")" +
						((msgErroCompleto ?? "").Length > 0 ? "\r\n" + msgErroCompleto : "");
				msgErroResumido = "Falha ao tentar reiniciar objetos DAO do banco de dados (Servidor: " + EnderecoServidor + ", BD: " + NomeBancoDados + ")" +
						((msgErroResumido ?? "").Length > 0 ? "\r\n" + msgErroResumido : "");
				return false;
			}

			return true;
		}
		#endregion
	}
	#endregion
}
