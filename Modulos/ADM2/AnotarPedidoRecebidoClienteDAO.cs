#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace ADM2
{
	public class AnotarPedidoRecebidoClienteDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
		private SqlCommand cmUpdatePedidoRecebidoData;
		private SqlCommand cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido;
		private SqlCommand cmUpdatePrevisaoEntregaTranspData;
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

		#region [ Construtor ]
		public AnotarPedidoRecebidoClienteDAO(ref BancoDados bd)
		{
			_bd = bd;
			inicializaObjetos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetos ]
		public void inicializaObjetos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmUpdatePedidoRecebidoData ]
			strSql = "UPDATE t_PEDIDO SET " +
						"PedidoRecebidoStatus = 1, " +
						"PedidoRecebidoData = @PedidoRecebidoData, " +
						"PedidoRecebidoDtHrUltAtualiz = getdate(), " +
						"PedidoRecebidoUsuarioUltAtualiz = @PedidoRecebidoUsuarioUltAtualiz" +
					" WHERE" +
						" (pedido = @pedido)" +
						" AND (PedidoRecebidoStatus = 0)";
			cmUpdatePedidoRecebidoData = _bd.criaSqlCommand();
			cmUpdatePedidoRecebidoData.CommandText = strSql;
			cmUpdatePedidoRecebidoData.Parameters.Add("@PedidoRecebidoData", SqlDbType.VarChar, 10);
			cmUpdatePedidoRecebidoData.Parameters.Add("@PedidoRecebidoUsuarioUltAtualiz", SqlDbType.VarChar, 10);
			cmUpdatePedidoRecebidoData.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoRecebidoData.Prepare();
			#endregion

			#region [ cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido ]
			strSql = "UPDATE t_PEDIDO SET " +
						"MarketplacePedidoRecebidoRegistrarStatus = 1, " +
						"MarketplacePedidoRecebidoRegistrarDataRecebido = @MarketplacePedidoRecebidoRegistrarDataRecebido, " +
						"MarketplacePedidoRecebidoRegistrarDataHora = getdate(), " +
						"MarketplacePedidoRecebidoRegistrarUsuario = @MarketplacePedidoRecebidoRegistrarUsuario" +
					" WHERE" +
						" (pedido = @pedido)" +
						" AND (LEN(Coalesce(marketplace_codigo_origem,'')) > 0)" +
						" AND (MarketplacePedidoRecebidoRegistrarStatus = 0)";
			cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = _bd.criaSqlCommand();
			cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.CommandText = strSql;
			cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Parameters.Add("@MarketplacePedidoRecebidoRegistrarDataRecebido", SqlDbType.VarChar, 10);
			cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Parameters.Add("@MarketplacePedidoRecebidoRegistrarUsuario", SqlDbType.VarChar, 10);
			cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Prepare();
			#endregion

			#region [ cmUpdatePrevisaoEntregaTranspData ]
			strSql = "UPDATE t_PEDIDO SET " +
						" PrevisaoEntregaTranspDataAnterior = PrevisaoEntregaTranspData" +
						", PrevisaoEntregaTranspData = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@PrevisaoEntregaTranspData") +
						", PrevisaoEntregaTranspDtHrUltAtualiz = getdate()" +
						", PrevisaoEntregaTranspUsuarioUltAtualiz = @PrevisaoEntregaTranspUsuarioUltAtualiz" +
					" WHERE" +
						" (pedido = @pedido)";
			cmUpdatePrevisaoEntregaTranspData = _bd.criaSqlCommand();
			cmUpdatePrevisaoEntregaTranspData.CommandText = strSql;
			cmUpdatePrevisaoEntregaTranspData.Parameters.Add("@PrevisaoEntregaTranspData", SqlDbType.VarChar, 10);
			cmUpdatePrevisaoEntregaTranspData.Parameters.Add("@PrevisaoEntregaTranspUsuarioUltAtualiz", SqlDbType.VarChar, 20);
			cmUpdatePrevisaoEntregaTranspData.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePrevisaoEntregaTranspData.Prepare();
			#endregion
		}
		#endregion

		#region [ UpdatePedidoRecebidoData ]
		public bool UpdatePedidoRecebidoData(string pedido, DateTime pedidoRecebidoData, string usuario, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "AnotarPedidoRecebidoClienteDAO.UpdatePedidoRecebidoData()";
			string strMsg;
			string strMsgErroLog = "";
			int intRetorno;
			Log log = new Log();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((pedido ?? "").Length == 0)
				{
					msg_erro = "Número do pedido não foi informado!";
					return false;
				}

				if (pedidoRecebidoData == null)
				{
					msg_erro = "Data de recebimento do pedido pelo cliente não foi informada!";
					return false;
				}

				if (pedidoRecebidoData == DateTime.MinValue)
				{
					msg_erro = "Data de recebimento do pedido pelo cliente não foi fornecida!";
					return false;
				}

				if ((usuario ?? "").Length == 0)
				{
					msg_erro = "Usuário responsável pela atualização não foi informado!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoRecebidoData.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoRecebidoData.Parameters["@PedidoRecebidoData"].Value = Global.formataDataYyyyMmDdComSeparador(pedidoRecebidoData);
				cmUpdatePedidoRecebidoData.Parameters["@PedidoRecebidoUsuarioUltAtualiz"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = _bd.executaNonQuery(ref cmUpdatePedidoRecebidoData);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;
					strMsg = NOME_DESTA_ROTINA + " - Exception ao tentar atualizar o pedido " + pedido + ": PedidoRecebidoData=" + Global.formataDataDdMmYyyyComSeparador(pedidoRecebidoData) + ", PedidoRecebidoUsuarioUltAtualiz=" + usuario + "\r\n" + ex.ToString();
					Global.gravaLogAtividade(strMsg);
					return false;
				}
				#endregion

				if (intRetorno == 0)
				{
					msg_erro = NOME_DESTA_ROTINA + " - Falha ao tentar atualizar o pedido " + pedido + ": PedidoRecebidoData=" + Global.formataDataDdMmYyyyComSeparador(pedidoRecebidoData) + ", PedidoRecebidoUsuarioUltAtualiz=" + usuario;
					return false;
				}

				#region [ Registra log com a alteração dos dados ]
				strMsg = "[Módulo ADM2] Atualização do pedido " + pedido + ": PedidoRecebidoData=" + Global.formataDataDdMmYyyyComSeparador(pedidoRecebidoData);
				log.usuario = usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_PEDIDO_RECEBIDO;
				log.pedido = pedido;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				strMsg = NOME_DESTA_ROTINA + " - Sucesso na atualização do pedido " + pedido + ": PedidoRecebidoData=" + Global.formataDataDdMmYyyyComSeparador(pedidoRecebidoData) + ", PedidoRecebidoUsuarioUltAtualiz=" + usuario;
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ UpdateMarketplacePedidoRecebidoRegistrarDataRecebido ]
		public bool UpdateMarketplacePedidoRecebidoRegistrarDataRecebido(string pedido, DateTime marketplacePedidoRecebidoRegistrarDataRecebido, string usuario, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "AnotarPedidoRecebidoClienteDAO.UpdateMarketplacePedidoRecebidoRegistrarDataRecebido()";
			string strMsg;
			string strMsgErroLog = "";
			int intRetorno;
			Log log = new Log();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((pedido ?? "").Length == 0)
				{
					msg_erro = "Número do pedido não foi informado!";
					return false;
				}

				if (marketplacePedidoRecebidoRegistrarDataRecebido == null)
				{
					msg_erro = "Data de recebimento do pedido pelo cliente não foi informada!";
					return false;
				}

				if (marketplacePedidoRecebidoRegistrarDataRecebido == DateTime.MinValue)
				{
					msg_erro = "Data de recebimento do pedido pelo cliente não foi fornecida!";
					return false;
				}

				if ((usuario ?? "").Length == 0)
				{
					msg_erro = "Usuário responsável pela atualização não foi informado!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Parameters["@pedido"].Value = pedido;
				cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Parameters["@MarketplacePedidoRecebidoRegistrarDataRecebido"].Value = Global.formataDataYyyyMmDdComSeparador(marketplacePedidoRecebidoRegistrarDataRecebido);
				cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Parameters["@MarketplacePedidoRecebidoRegistrarUsuario"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = _bd.executaNonQuery(ref cmUpdateMarketplacePedidoRecebidoRegistrarDataRecebido);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;
					strMsg = NOME_DESTA_ROTINA + " - Exception ao tentar atualizar o pedido " + pedido + ": MarketplacePedidoRecebidoRegistrarDataRecebido=" + Global.formataDataDdMmYyyyComSeparador(marketplacePedidoRecebidoRegistrarDataRecebido) + ", MarketplacePedidoRecebidoRegistrarUsuario=" + usuario + "\r\n" + ex.ToString();
					Global.gravaLogAtividade(strMsg);
					return false;
				}
				#endregion

				if (intRetorno == 0)
				{
					msg_erro = NOME_DESTA_ROTINA + " - Falha ao tentar atualizar o pedido " + pedido + ": MarketplacePedidoRecebidoRegistrarDataRecebido=" + Global.formataDataDdMmYyyyComSeparador(marketplacePedidoRecebidoRegistrarDataRecebido) + ", MarketplacePedidoRecebidoRegistrarUsuario=" + usuario;
					return false;
				}

				#region [ Registra log com a alteração dos dados ]
				strMsg = "[Módulo ADM2] Atualização do pedido " + pedido + ": MarketplacePedidoRecebidoRegistrarDataRecebido=" + Global.formataDataDdMmYyyyComSeparador(marketplacePedidoRecebidoRegistrarDataRecebido);
				log.usuario = usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_PEDIDO_RECEBIDO_MARKETPLACE;
				log.pedido = pedido;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				strMsg = NOME_DESTA_ROTINA + " - Sucesso na atualização do pedido " + pedido + ": MarketplacePedidoRecebidoRegistrarDataRecebido=" + Global.formataDataDdMmYyyyComSeparador(marketplacePedidoRecebidoRegistrarDataRecebido) + ", MarketplacePedidoRecebidoRegistrarUsuario=" + usuario;
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ UpdatePrevisaoEntregaTranspData ]
		public bool UpdatePrevisaoEntregaTranspData(string pedido, DateTime previsaoEntregaTranspData, DateTime previsaoEntregaTranspDataOriginal, string usuario, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "AnotarPedidoRecebidoClienteDAO.UpdatePrevisaoEntregaTranspData()";
			string strMsg;
			string strMsgErroLog = "";
			string sData;
			int intRetorno;
			Log log = new Log();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((pedido ?? "").Length == 0)
				{
					msg_erro = "Número do pedido não foi informado!";
					return false;
				}

				if ((usuario ?? "").Length == 0)
				{
					msg_erro = "Usuário responsável pela atualização não foi informado!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePrevisaoEntregaTranspData.Parameters["@pedido"].Value = pedido;
				cmUpdatePrevisaoEntregaTranspData.Parameters["@PrevisaoEntregaTranspUsuarioUltAtualiz"].Value = usuario;

				// Se a data for vazia, será atualizada com NULL
				sData = "";
				if (previsaoEntregaTranspData != null)
				{
					if (previsaoEntregaTranspData != DateTime.MinValue)
					{
						sData = Global.formataDataYyyyMmDdComSeparador(previsaoEntregaTranspData);
					}
				}
				cmUpdatePrevisaoEntregaTranspData.Parameters["@PrevisaoEntregaTranspData"].Value = sData;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = _bd.executaNonQuery(ref cmUpdatePrevisaoEntregaTranspData);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;
					strMsg = NOME_DESTA_ROTINA + " - Exception ao tentar atualizar o pedido " + pedido + ": PrevisaoEntregaTranspData=" + Global.formataDataDdMmYyyyComSeparador(previsaoEntregaTranspData) + ", PrevisaoEntregaTranspUsuarioUltAtualiz=" + usuario + "\r\n" + ex.ToString();
					Global.gravaLogAtividade(strMsg);
					return false;
				}
				#endregion

				if (intRetorno == 0)
				{
					msg_erro = NOME_DESTA_ROTINA + " - Falha ao tentar atualizar o pedido " + pedido + ": PrevisaoEntregaTranspData=" + Global.formataDataDdMmYyyyComSeparador(previsaoEntregaTranspData) + ", PrevisaoEntregaTranspUsuarioUltAtualiz=" + usuario;
					return false;
				}

				#region [ Registra log com a alteração dos dados ]
				strMsg = "[Módulo ADM2] Atualização do pedido " + pedido
						+ ": PrevisaoEntregaTranspData=" + (previsaoEntregaTranspData == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(previsaoEntregaTranspData))
						+ " (data anterior: " + (previsaoEntregaTranspDataOriginal == DateTime.MinValue ? "null" : Global.formataDataDdMmYyyyComSeparador(previsaoEntregaTranspDataOriginal)) + ")";
				log.usuario = usuario;
				log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_PEDIDO_ATUALIZA_PREVISAO_ENTREGA_TRANSP;
				log.pedido = pedido;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				strMsg = NOME_DESTA_ROTINA + " - Sucesso na atualização do pedido " + pedido + ": PrevisaoEntregaTranspData=" + Global.formataDataDdMmYyyyComSeparador(previsaoEntregaTranspData) + ", PrevisaoEntregaTranspUsuarioUltAtualiz=" + usuario;
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
				return false;
			}
		}
		#endregion

		#endregion
	}
}
