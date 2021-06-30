#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
#endregion

namespace FinanceiroService
{
	#region [ Braspag ]
	static class Braspag
	{
		#region [ decodifica_PaymentDataResponseStatus_para_GlobalStatus ]
		/// <summary>
		/// Decodifica o código de status retornado em PaymentDataResponse para um código global.
		/// Caso seja informado um código de status desconhecido, o mesmo será retornado com a seguinte formatação: 'PGnnn'
		/// 'PG' = Payment
		/// 'nnn' = código do status desconhecido formatado c/ zeros à esquerda
		/// </summary>
		/// <param name="codigoStatus">Status retornado em PaymentDataResponse</param>
		/// <returns>Código global de status</returns>
		public static string decodifica_PaymentDataResponseStatus_para_GlobalStatus(string codigoStatus)
		{
			#region [ Declarações ]
			string strResp = "";
			#endregion

			if (Global.Cte.Braspag.Pagador.PaymentDataResponseStatus.CAPTURADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.PaymentDataResponseStatus.AUTORIZADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.PaymentDataResponseStatus.NAO_AUTORIZADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.NAO_AUTORIZADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.PaymentDataResponseStatus.ERRO_DESQUALIFICANTE.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.ERRO_DESQUALIFICANTE.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.PaymentDataResponseStatus.AGUARDANDO_RESPOSTA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.AGUARDANDO_RESPOSTA.GetValue();
			}
			else
			{
				// CÓDIGO DESCONHECIDO
				strResp = "PG" + codigoStatus.PadLeft(3, '0');
			}

			return strResp;
		}
		#endregion

		#region [ decodifica_GetTransactionDataResponseStatus_para_GlobalStatus ]
		/// <summary>
		/// Decodifica o código de status retornado em GetTransactionDataResponse para um código global.
		/// Caso seja informado um código de status desconhecido, o mesmo será retornado com a seguinte formatação: 'QYnnn'
		/// 'QY' = Query
		/// 'nnn' = código do status desconhecido formatado c/ zeros à esquerda
		/// </summary>
		/// <param name="codigoStatus">Status retornado em GetTransactionDataResponse</param>
		/// <returns>Código global de status</returns>
		public static string decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(string codigoStatus)
		{
			#region [ Declarações ]
			string strResp = "";
			#endregion

			if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.INDEFINIDA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.INDEFINIDA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.CAPTURADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.AUTORIZADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.NAO_AUTORIZADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.NAO_AUTORIZADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.CAPTURA_CANCELADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURA_CANCELADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.ESTORNADA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNADA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.AGUARDANDO_RESPOSTA.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.AGUARDANDO_RESPOSTA.GetValue();
			}
			else if (Global.Cte.Braspag.Pagador.GetTransactionDataResponseStatus.ERRO_DESQUALIFICANTE.GetValue().Equals(codigoStatus))
			{
				strResp = Global.Cte.Braspag.Pagador.GlobalStatus.ERRO_DESQUALIFICANTE.GetValue();
			}
			else
			{
				// CÓDIGO DESCONHECIDO
				strResp = "QY" + codigoStatus.PadLeft(3, '0');
			}

			return strResp;
		}
		#endregion

		#region [ descricaoOperacaoRegistraPagto ]
		public static string descricaoOperacaoRegistraPagto(Global.Cte.Braspag.Pagador.Transacao tipoTransacao)
		{
			string strResp = "";

			if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.AuthorizeTransaction.GetCodOpLog()))
			{
				strResp = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.AUTORIZACAO.GetDescription();
			}
			else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog()))
			{
				strResp = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.CAPTURA.GetDescription();
			}
			else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()))
			{
				strResp = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.CANCELAMENTO.GetDescription();
			}
			else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction.GetCodOpLog()))
			{
				strResp = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.ESTORNO.GetDescription();
			}
			else
			{
				strResp = "Código desconhecido (" + tipoTransacao.GetCodOpLog() + ")";
			}

			return strResp;
		}
		#endregion

		#region [ criaGetOrderIdData ]
		public static BraspagGetOrderIdData criaGetOrderIdData(string MerchantId, string OrderId)
		{
			BraspagGetOrderIdData trx = new BraspagGetOrderIdData();
			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;
			trx.OrderId = OrderId;
			return trx;
		}
		#endregion

		#region [ criaGetOrderData ]
		public static BraspagGetOrderData criaGetOrderData(string MerchantId, string BraspagOrderId)
		{
			BraspagGetOrderData trx = new BraspagGetOrderData();
			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;
			trx.BraspagOrderId = BraspagOrderId;
			return trx;
		}
		#endregion

		#region [ criaGetBoletoData ]
		public static BraspagGetBoletoData criaGetBoletoData(string MerchantId, string BraspagTransactionId)
		{
			BraspagGetBoletoData trx = new BraspagGetBoletoData();
			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;
			trx.BraspagTransactionId = BraspagTransactionId;
			return trx;
		}
		#endregion

		#region [ criaGetTransactionData ]
		public static BraspagGetTransactionData criaGetTransactionData(string MerchantId, string BraspagTransactionId)
		{
			BraspagGetTransactionData trx = new BraspagGetTransactionData();
			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;
			trx.BraspagTransactionId = BraspagTransactionId;
			return trx;
		}
		#endregion

		#region [ criaCaptureCreditCardTransaction ]
		public static BraspagCaptureCreditCardTransaction criaCaptureCreditCardTransaction(string MerchantId, List<BraspagTransactionDataRequest> vBraspagTransactionId)
		{
			BraspagCaptureCreditCardTransaction trx = new BraspagCaptureCreditCardTransaction();
			BraspagTransactionDataRequest request;

			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;

			for (int i = 0; i < vBraspagTransactionId.Count; i++)
			{
				request = new BraspagTransactionDataRequest();
				request.BraspagTransactionId = vBraspagTransactionId[i].BraspagTransactionId;
				request.Amount = vBraspagTransactionId[i].Amount;
				request.ServiceTaxAmount = vBraspagTransactionId[i].ServiceTaxAmount;
				trx.TransactionDataCollection.Add(request);
			}
			return trx;
		}
		#endregion

		#region [ criaVoidCreditCardTransaction ]
		public static BraspagVoidCreditCardTransaction criaVoidCreditCardTransaction(string MerchantId, List<BraspagTransactionDataRequest> vBraspagTransactionId)
		{
			BraspagVoidCreditCardTransaction trx = new BraspagVoidCreditCardTransaction();
			BraspagTransactionDataRequest request;

			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;

			for (int i = 0; i < vBraspagTransactionId.Count; i++)
			{
				request = new BraspagTransactionDataRequest();
				request.BraspagTransactionId = vBraspagTransactionId[i].BraspagTransactionId;
				request.Amount = vBraspagTransactionId[i].Amount;
				request.ServiceTaxAmount = vBraspagTransactionId[i].ServiceTaxAmount;
				trx.TransactionDataCollection.Add(request);
			}
			return trx;
		}
		#endregion

		#region [ criaRefundCreditCardTransaction ]
		public static BraspagRefundCreditCardTransaction criaRefundCreditCardTransaction(string MerchantId, List<BraspagTransactionDataRequest> vBraspagTransactionId)
		{
			BraspagRefundCreditCardTransaction trx = new BraspagRefundCreditCardTransaction();
			BraspagTransactionDataRequest request;

			trx.Version = Global.Cte.Braspag.Pagador.Version;
			trx.RequestId = BD.gera_uid().ToLower();
			trx.MerchantId = MerchantId;

			for (int i = 0; i < vBraspagTransactionId.Count; i++)
			{
				request = new BraspagTransactionDataRequest();
				request.BraspagTransactionId = vBraspagTransactionId[i].BraspagTransactionId;
				request.Amount = vBraspagTransactionId[i].Amount;
				request.ServiceTaxAmount = vBraspagTransactionId[i].ServiceTaxAmount;
				trx.TransactionDataCollection.Add(request);
			}
			return trx;
		}
		#endregion

		#region [ montaXmlRequisicaoSoapGetOrderIdData ]
		public static string montaXmlRequisicaoSoapGetOrderIdData(BraspagGetOrderIdData trx)
		{
			StringBuilder sbTrx = new StringBuilder("");
			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<GetOrderIdData xmlns=\"https://www.pagador.com.br/query/pagadorquery\">");
			sbTrx.Append("<orderIdDataRequest>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<OrderId>" + trx.OrderId + "</OrderId>");
			sbTrx.Append("</orderIdDataRequest>");
			sbTrx.Append("</GetOrderIdData>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ montaXmlRequisicaoSoapGetOrderData ]
		public static string montaXmlRequisicaoSoapGetOrderData(BraspagGetOrderData trx)
		{
			StringBuilder sbTrx = new StringBuilder("");
			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<GetOrderData xmlns=\"https://www.pagador.com.br/query/pagadorquery\">");
			sbTrx.Append("<orderDataRequest>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<BraspagOrderId>" + trx.BraspagOrderId + "</BraspagOrderId>");
			sbTrx.Append("</orderDataRequest>");
			sbTrx.Append("</GetOrderData>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ montaXmlRequisicaoSoapGetBoletoData ]
		public static string montaXmlRequisicaoSoapGetBoletoData(BraspagGetBoletoData trx)
		{
			StringBuilder sbTrx = new StringBuilder("");
			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<GetBoletoData xmlns=\"https://www.pagador.com.br/query/pagadorquery\">");
			sbTrx.Append("<boletoDataRequest>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<BraspagTransactionId>" + trx.BraspagTransactionId + "</BraspagTransactionId>");
			sbTrx.Append("</boletoDataRequest>");
			sbTrx.Append("</GetBoletoData>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ montaXmlRequisicaoSoapGetTransactionData ]
		public static string montaXmlRequisicaoSoapGetTransactionData(BraspagGetTransactionData trx)
		{
			StringBuilder sbTrx = new StringBuilder("");
			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<GetTransactionData xmlns=\"https://www.pagador.com.br/query/pagadorquery\">");
			sbTrx.Append("<transactionDataRequest>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<BraspagTransactionId>" + trx.BraspagTransactionId + "</BraspagTransactionId>");
			sbTrx.Append("</transactionDataRequest>");
			sbTrx.Append("</GetTransactionData>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ montaXmlRequisicaoSoapCaptureCreditCardTransaction ]
		public static string montaXmlRequisicaoSoapCaptureCreditCardTransaction(BraspagCaptureCreditCardTransaction trx)
		{
			#region [ Declarações ]
			StringBuilder sbTrx = new StringBuilder("");
			BraspagTransactionDataRequest request;
			#endregion

			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<CaptureCreditCardTransaction xmlns=\"https://www.pagador.com.br/webservice/pagador\">");
			sbTrx.Append("<request>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<TransactionDataCollection>");
			for (int i = 0; i < trx.TransactionDataCollection.Count; i++)
			{
				request = trx.TransactionDataCollection[i];
				sbTrx.Append("<TransactionDataRequest>");
				sbTrx.Append("<BraspagTransactionId>" + request.BraspagTransactionId + "</BraspagTransactionId>");
				sbTrx.Append("<Amount>" + request.Amount + "</Amount>");
				sbTrx.Append("<ServiceTaxAmount>" + request.ServiceTaxAmount + "</ServiceTaxAmount>");
				sbTrx.Append("</TransactionDataRequest>");
			}
			sbTrx.Append("</TransactionDataCollection>");
			sbTrx.Append("</request>");
			sbTrx.Append("</CaptureCreditCardTransaction>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ montaXmlRequisicaoSoapVoidCreditCardTransaction ]
		public static string montaXmlRequisicaoSoapVoidCreditCardTransaction(BraspagVoidCreditCardTransaction trx)
		{
			#region [ Declarações ]
			StringBuilder sbTrx = new StringBuilder("");
			BraspagTransactionDataRequest request;
			#endregion

			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<VoidCreditCardTransaction xmlns=\"https://www.pagador.com.br/webservice/pagador\">");
			sbTrx.Append("<request>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<TransactionDataCollection>");
			for (int i = 0; i < trx.TransactionDataCollection.Count; i++)
			{
				request = trx.TransactionDataCollection[i];
				sbTrx.Append("<TransactionDataRequest>");
				sbTrx.Append("<BraspagTransactionId>" + request.BraspagTransactionId + "</BraspagTransactionId>");
				sbTrx.Append("<Amount>" + request.Amount + "</Amount>");
				sbTrx.Append("<ServiceTaxAmount>" + request.ServiceTaxAmount + "</ServiceTaxAmount>");
				sbTrx.Append("</TransactionDataRequest>");
			}
			sbTrx.Append("</TransactionDataCollection>");
			sbTrx.Append("</request>");
			sbTrx.Append("</VoidCreditCardTransaction>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ montaXmlRequisicaoSoapRefundCreditCardTransaction ]
		public static string montaXmlRequisicaoSoapRefundCreditCardTransaction(BraspagRefundCreditCardTransaction trx)
		{
			#region [ Declarações ]
			StringBuilder sbTrx = new StringBuilder("");
			BraspagTransactionDataRequest request;
			#endregion

			sbTrx.Append("<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">");
			sbTrx.Append("<soap:Body>");
			sbTrx.Append("<RefundCreditCardTransaction xmlns=\"https://www.pagador.com.br/webservice/pagador\">");
			sbTrx.Append("<request>");
			sbTrx.Append("<Version>" + trx.Version + "</Version>");
			sbTrx.Append("<RequestId>" + trx.RequestId + "</RequestId>");
			sbTrx.Append("<MerchantId>" + trx.MerchantId + "</MerchantId>");
			sbTrx.Append("<TransactionDataCollection>");
			for (int i = 0; i < trx.TransactionDataCollection.Count; i++)
			{
				request = trx.TransactionDataCollection[i];
				sbTrx.Append("<TransactionDataRequest>");
				sbTrx.Append("<BraspagTransactionId>" + request.BraspagTransactionId + "</BraspagTransactionId>");
				sbTrx.Append("<Amount>" + request.Amount + "</Amount>");
				sbTrx.Append("<ServiceTaxAmount>" + request.ServiceTaxAmount + "</ServiceTaxAmount>");
				sbTrx.Append("</TransactionDataRequest>");
			}
			sbTrx.Append("</TransactionDataCollection>");
			sbTrx.Append("</request>");
			sbTrx.Append("</RefundCreditCardTransaction>");
			sbTrx.Append("</soap:Body>");
			sbTrx.Append("</soap:Envelope>");
			return sbTrx.ToString();
		}
		#endregion

		#region [ decodificaXmlGetOrderIdDataResponse ]
		private static BraspagGetOrderIdDataResponse decodificaXmlGetOrderIdDataResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlGetOrderIdDataResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagGetOrderIdDataResponse rRESP = new BraspagGetOrderIdDataResponse();
			BraspagOrderIdTransactionResponse order;
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNodeList xmlNodeListL2;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/query/pagadorquery");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetOrderIdDataResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetOrderIdDataResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ OrderIdDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:GetOrderIdDataResult/pag:OrderIdDataCollection/pag:OrderIdTransactionResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					order = new BraspagOrderIdTransactionResponse();
					order.CorrelationId = Global.obtemXmlChildNodeValue(node, "CorrelationId");
					order.Success = Global.obtemXmlChildNodeValue(node, "Success");
					order.BraspagOrderId = Global.obtemXmlChildNodeValue(node, "BraspagOrderId");

					#region [ BraspagTransactionId ]
					// IMPORTANTE: QUANDO UM PAGAMENTO É FEITO ATRAVÉS DE 2 OU MAIS CARTÕES ENVIADOS NA MESMA REQUISIÇÃO, ESTA
					// ==========  CONSULTA RETORNA VÁRIOS 'BraspagTransactionId'
					// EX:
					// 		<BraspagTransactionId>
					//			<guid>2bdcc5f5-4331-4123-89f4-864d8203ca97</guid>
					//			<guid>7236d9e3-88f7-4627-856c-948302c4d41c</guid>
					//		</BraspagTransactionId>
					xmlNodeListL2 = node.SelectNodes("pag:BraspagTransactionId/pag:guid", nsmgr);
					foreach (XmlNode nodeL2 in xmlNodeListL2)
					{
						if (nodeL2 == null) continue;
						if (nodeL2.InnerText == null) continue;
						strValue = nodeL2.InnerText;
						order.BraspagTransactionId.Add(strValue);
					}
					#endregion

					#region [ ErrorReportDataCollection ]
					// Esta consulta também retorna uma coleção 'ErrorReportDataCollection' dentro do bloco 'OrderIdTransactionResponse'
					// Importante: o comando
					//				node.SelectNodes("//pag:GetOrderIdDataResult/pag:OrderIdDataCollection/pag:OrderIdTransactionResponse/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
					// retorna todos os erros contidos no XML, mesmo os que estão no nível mais externo ao 'node' selecionado e também os de outros nodes.
					// O comando correto para obter os erros que estão dentro do node selecionado é este:
					//				node.SelectNodes("pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
					xmlNodeListL2 = node.SelectNodes("pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
					foreach (XmlNode nodeL2 in xmlNodeListL2)
					{
						errorReportDataResponse = new BraspagErrorReportDataResponse();
						errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(nodeL2, "ErrorCode");
						errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(nodeL2, "ErrorMessage");
						order.ErrorReportDataCollection.Add(errorReportDataResponse);
					}
					#endregion

					rRESP.OrderIdDataCollection.Add(order);
				}
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:GetOrderIdDataResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ decodificaXmlGetOrderDataResponse ]
		private static BraspagGetOrderDataResponse decodificaXmlGetOrderDataResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlGetOrderDataResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagGetOrderDataResponse rRESP = new BraspagGetOrderDataResponse();
			BraspagOrderTransactionDataResponse order;
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/query/pagadorquery");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetOrderDataResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetOrderDataResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ TransactionDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:GetOrderDataResult/pag:TransactionDataCollection/pag:OrderTransactionDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					order = new BraspagOrderTransactionDataResponse();
					order.BraspagTransactionId = Global.obtemXmlChildNodeValue(node, "BraspagTransactionId");
					order.OrderId = Global.obtemXmlChildNodeValue(node, "OrderId");
					order.AcquirerTransactionId = Global.obtemXmlChildNodeValue(node, "AcquirerTransactionId");
					order.PaymentMethod = Global.obtemXmlChildNodeValue(node, "PaymentMethod");
					order.PaymentMethodName = Global.obtemXmlChildNodeValue(node, "PaymentMethodName");
					order.Amount = Global.obtemXmlChildNodeValue(node, "Amount");
					order.AuthorizationCode = Global.obtemXmlChildNodeValue(node, "AuthorizationCode");
					order.NumberOfPayments = Global.obtemXmlChildNodeValue(node, "NumberOfPayments");
					order.Currency = Global.obtemXmlChildNodeValue(node, "Currency");
					order.Country = Global.obtemXmlChildNodeValue(node, "Country");
					order.TransactionType = Global.obtemXmlChildNodeValue(node, "TransactionType");
					order.Status = Global.obtemXmlChildNodeValue(node, "Status");
					order.ReceivedDate = Global.obtemXmlChildNodeValue(node, "ReceivedDate");
					order.CapturedDate = Global.obtemXmlChildNodeValue(node, "CapturedDate");
					order.VoidedDate = Global.obtemXmlChildNodeValue(node, "VoidedDate");
					order.CreditCardToken = Global.obtemXmlChildNodeValue(node, "CreditCardToken");
					order.ProofOfSale = Global.obtemXmlChildNodeValue(node, "ProofOfSale");
					order.MaskedCardNumber = Global.obtemXmlChildNodeValue(node, "MaskedCardNumber");
					rRESP.TransactionDataCollection.Add(order);
				}
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:GetOrderDataResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ decodificaXmlGetBoletoDataResponse ]
		private static BraspagGetBoletoDataResponse decodificaXmlGetBoletoDataResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlGetBoletoDataResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagGetBoletoDataResponse rRESP = new BraspagGetBoletoDataResponse();
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/query/pagadorquery");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ BraspagTransactionId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BraspagTransactionId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BraspagTransactionId = strValue;
				#endregion

				#region [ PaymentMethod ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:PaymentMethod", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.PaymentMethod = strValue;
				#endregion

				#region [ DocumentNumber ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:DocumentNumber", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.DocumentNumber = strValue;
				#endregion

				#region [ DocumentDate ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:DocumentDate", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.DocumentDate = strValue;
				#endregion

				#region [ CustomerName ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:CustomerName", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CustomerName = strValue;
				#endregion

				#region [ BoletoNumber ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BoletoNumber", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BoletoNumber = strValue;
				#endregion

				#region [ BarCodeNumber ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BarCodeNumber", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BarCodeNumber = strValue;
				#endregion

				#region [ BoletoExpirationDate ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BoletoExpirationDate", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BoletoExpirationDate = strValue;
				#endregion

				#region [ BoletoInstructions ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BoletoInstructions", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BoletoInstructions = strValue;
				#endregion

				#region [ BoletoType ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BoletoType", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BoletoType = strValue;
				#endregion

				#region [ BoletoUrl ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BoletoUrl", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BoletoUrl = strValue;
				#endregion

				#region [ Amount ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:Amount", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Amount = strValue;
				#endregion

				#region [ PaidAmount ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:PaidAmount", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.PaidAmount = strValue;
				#endregion

				#region [ PaymentDate ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:PaymentDate", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.PaymentDate = strValue;
				#endregion

				#region [ BankNumber ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:BankNumber", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BankNumber = strValue;
				#endregion

				#region [ Agency ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:Agency", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Agency = strValue;
				#endregion

				#region [ Account ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:Account", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Account = strValue;
				#endregion

				#region [ Assignor ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetBoletoDataResult/pag:Assignor", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Assignor = strValue;
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:GetBoletoDataResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ decodificaXmlGetTransactionDataResponse ]
		private static BraspagGetTransactionDataResponse decodificaXmlGetTransactionDataResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlGetTransactionDataResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagGetTransactionDataResponse rRESP = new BraspagGetTransactionDataResponse();
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/query/pagadorquery");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ BraspagTransactionId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:BraspagTransactionId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.BraspagTransactionId = strValue;
				#endregion

				#region [ OrderId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:OrderId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.OrderId = strValue;
				#endregion

				#region [ AcquirerTransactionId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:AcquirerTransactionId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.AcquirerTransactionId = strValue;
				#endregion

				#region [ PaymentMethod ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:PaymentMethod", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.PaymentMethod = strValue;
				#endregion

				#region [ PaymentMethodName ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:PaymentMethodName", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.PaymentMethodName = strValue;
				#endregion

				#region [ Amount ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:Amount", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Amount = strValue;
				#endregion

				#region [ AuthorizationCode ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:AuthorizationCode", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.AuthorizationCode = strValue;
				#endregion

				#region [ NumberOfPayments ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:NumberOfPayments", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.NumberOfPayments = strValue;
				#endregion

				#region [ Currency ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:Currency", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Currency = strValue;
				#endregion

				#region [ Country ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:Country", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Country = strValue;
				#endregion

				#region [ TransactionType ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:TransactionType", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.TransactionType = strValue;
				#endregion

				#region [ Status ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:Status", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Status = strValue;
				#endregion

				#region [ ReceivedDate ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:ReceivedDate", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.ReceivedDate = strValue;
				#endregion

				#region [ CapturedDate ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:CapturedDate", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CapturedDate = strValue;
				#endregion

				#region [ VoidedDate ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:VoidedDate", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.VoidedDate = strValue;
				#endregion

				#region [ CreditCardToken ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:CreditCardToken", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CreditCardToken = strValue;
				#endregion

				#region [ ProofOfSale ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:ProofOfSale", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.ProofOfSale = strValue;
				#endregion

				#region [ MaskedCardNumber ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:GetTransactionDataResult/pag:MaskedCardNumber", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.MaskedCardNumber = strValue;
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:GetTransactionDataResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ decodificaXmlCaptureCreditCardTransactionResponse ]
		private static BraspagCaptureCreditCardTransactionResponse decodificaXmlCaptureCreditCardTransactionResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlCaptureCreditCardTransactionResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagCaptureCreditCardTransactionResponse rRESP = new BraspagCaptureCreditCardTransactionResponse();
			BraspagTransactionDataResponse rTrx;
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNodeList xmlNodeListL2;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/webservice/pagador");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:CaptureCreditCardTransactionResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:CaptureCreditCardTransactionResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ TransactionDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:CaptureCreditCardTransactionResult/pag:TransactionDataCollection/pag:TransactionDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					rTrx = new BraspagTransactionDataResponse();

					#region [ Obtém campos deste 'node' ]
					rTrx.BraspagTransactionId = Global.obtemXmlChildNodeValue(node, "BraspagTransactionId");
					rTrx.AcquirerTransactionId = Global.obtemXmlChildNodeValue(node, "AcquirerTransactionId");
					rTrx.Amount = Global.obtemXmlChildNodeValue(node, "Amount");
					rTrx.AuthorizationCode = Global.obtemXmlChildNodeValue(node, "AuthorizationCode");
					rTrx.ReturnCode = Global.obtemXmlChildNodeValue(node, "ReturnCode");
					rTrx.ReturnMessage = Global.obtemXmlChildNodeValue(node, "ReturnMessage");
					rTrx.Status = Global.obtemXmlChildNodeValue(node, "Status");
					rTrx.ProofOfSale = Global.obtemXmlChildNodeValue(node, "ProofOfSale");
					rTrx.ServiceTaxAmount = Global.obtemXmlChildNodeValue(node, "ServiceTaxAmount");
					#endregion

					#region [ Verifica se há mensagem de erro internas neste 'node' ]
					xmlNodeListL2 = node.SelectNodes("pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
					foreach (XmlNode nodeL2 in xmlNodeListL2)
					{
						errorReportDataResponse = new BraspagErrorReportDataResponse();
						errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(nodeL2, "ErrorCode");
						errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(nodeL2, "ErrorMessage");
						rTrx.ErrorReportDataCollection.Add(errorReportDataResponse);
					}
					#endregion

					rRESP.TransactionDataCollection.Add(rTrx);
				}
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:CaptureCreditCardTransactionResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ decodificaXmlVoidCreditCardTransactionResponse ]
		private static BraspagVoidCreditCardTransactionResponse decodificaXmlVoidCreditCardTransactionResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlVoidCreditCardTransactionResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagVoidCreditCardTransactionResponse rRESP = new BraspagVoidCreditCardTransactionResponse();
			BraspagTransactionDataResponse rTrx;
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNodeList xmlNodeListL2;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/webservice/pagador");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:VoidCreditCardTransactionResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:VoidCreditCardTransactionResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ TransactionDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:VoidCreditCardTransactionResult/pag:TransactionDataCollection/pag:TransactionDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					rTrx = new BraspagTransactionDataResponse();

					#region [ Obtém campos deste 'node' ]
					rTrx.BraspagTransactionId = Global.obtemXmlChildNodeValue(node, "BraspagTransactionId");
					rTrx.AcquirerTransactionId = Global.obtemXmlChildNodeValue(node, "AcquirerTransactionId");
					rTrx.Amount = Global.obtemXmlChildNodeValue(node, "Amount");
					rTrx.AuthorizationCode = Global.obtemXmlChildNodeValue(node, "AuthorizationCode");
					rTrx.ReturnCode = Global.obtemXmlChildNodeValue(node, "ReturnCode");
					rTrx.ReturnMessage = Global.obtemXmlChildNodeValue(node, "ReturnMessage");
					rTrx.Status = Global.obtemXmlChildNodeValue(node, "Status");
					rTrx.ProofOfSale = Global.obtemXmlChildNodeValue(node, "ProofOfSale");
					rTrx.ServiceTaxAmount = Global.obtemXmlChildNodeValue(node, "ServiceTaxAmount");
					#endregion

					#region [ Verifica se há mensagem de erro internas neste 'node' ]
					xmlNodeListL2 = node.SelectNodes("pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
					foreach (XmlNode nodeL2 in xmlNodeListL2)
					{
						errorReportDataResponse = new BraspagErrorReportDataResponse();
						errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(nodeL2, "ErrorCode");
						errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(nodeL2, "ErrorMessage");
						rTrx.ErrorReportDataCollection.Add(errorReportDataResponse);
					}
					#endregion

					rRESP.TransactionDataCollection.Add(rTrx);
				}
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:VoidCreditCardTransactionResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ decodificaXmlRefundCreditCardTransactionResponse ]
		private static BraspagRefundCreditCardTransactionResponse decodificaXmlRefundCreditCardTransactionResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.decodificaXmlRefundCreditCardTransactionResponse()";
			string msg_erro_aux;
			string strValue;
			BraspagRefundCreditCardTransactionResponse rRESP = new BraspagRefundCreditCardTransactionResponse();
			BraspagTransactionDataResponse rTrx;
			BraspagErrorReportDataResponse errorReportDataResponse;
			XmlDocument xmlDoc;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNodeList xmlNodeListL2;
			XmlNamespaceManager nsmgr;
			#endregion

			msg_erro = "";
			try
			{
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("pag", "https://www.pagador.com.br/webservice/pagador");

				#region [ CorrelationId ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:RefundCreditCardTransactionResult/pag:CorrelationId", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.CorrelationId = strValue;
				#endregion

				#region [ Success ]
				xmlNode = xmlDoc.SelectSingleNode("//pag:RefundCreditCardTransactionResult/pag:Success", nsmgr);
				strValue = Global.obtemXmlNodeFirstChildValue(xmlNode);
				if (strValue != null) rRESP.Success = strValue;
				#endregion

				#region [ TransactionDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:RefundCreditCardTransactionResult/pag:TransactionDataCollection/pag:TransactionDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					rTrx = new BraspagTransactionDataResponse();

					#region [ Obtém campos deste 'node' ]
					rTrx.BraspagTransactionId = Global.obtemXmlChildNodeValue(node, "BraspagTransactionId");
					rTrx.AcquirerTransactionId = Global.obtemXmlChildNodeValue(node, "AcquirerTransactionId");
					rTrx.Amount = Global.obtemXmlChildNodeValue(node, "Amount");
					rTrx.AuthorizationCode = Global.obtemXmlChildNodeValue(node, "AuthorizationCode");
					rTrx.ReturnCode = Global.obtemXmlChildNodeValue(node, "ReturnCode");
					rTrx.ReturnMessage = Global.obtemXmlChildNodeValue(node, "ReturnMessage");
					rTrx.Status = Global.obtemXmlChildNodeValue(node, "Status");
					rTrx.ProofOfSale = Global.obtemXmlChildNodeValue(node, "ProofOfSale");
					rTrx.ServiceTaxAmount = Global.obtemXmlChildNodeValue(node, "ServiceTaxAmount");
					#endregion

					#region [ Verifica se há mensagem de erro internas neste 'node' ]
					xmlNodeListL2 = node.SelectNodes("pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
					foreach (XmlNode nodeL2 in xmlNodeListL2)
					{
						errorReportDataResponse = new BraspagErrorReportDataResponse();
						errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(nodeL2, "ErrorCode");
						errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(nodeL2, "ErrorMessage");
						rTrx.ErrorReportDataCollection.Add(errorReportDataResponse);
					}
					#endregion

					rRESP.TransactionDataCollection.Add(rTrx);
				}
				#endregion

				#region [ ErrorReportDataCollection ]
				xmlNodeList = xmlDoc.SelectNodes("//pag:RefundCreditCardTransactionResult/pag:ErrorReportDataCollection/pag:ErrorReportDataResponse", nsmgr);
				foreach (XmlNode node in xmlNodeList)
				{
					errorReportDataResponse = new BraspagErrorReportDataResponse();
					errorReportDataResponse.ErrorCode = Global.obtemXmlChildNodeValue(node, "ErrorCode");
					errorReportDataResponse.ErrorMessage = Global.obtemXmlChildNodeValue(node, "ErrorMessage");
					rRESP.ErrorReportDataCollection.Add(errorReportDataResponse);
				}
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = xmlRespSoap;
				svcLog.complemento_2 = Global.serializaObjectToXml(rRESP);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ verificaPreRequisitoBraspagTransactionId ]
		/// <summary>
		/// Verifica se há a informação de 'BraspagTransactionId'. Caso não, executa a consulta 'GetOrderIdData' usando o campo 'OrderId' para obter o 'BraspagTransactionId', que é necessário para a maioria das requisições.
		/// </summary>
		/// <param name="payment"></param>
		/// <param name="msg_erro"></param>
		/// <returns></returns>
		public static bool verificaPreRequisitoBraspagTransactionId(BraspagPagPayment payment, out bool atualizouBraspagTransactionId, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.verificaPreRequisitoBraspagTransactionId()";
			bool blnBraspagTransactionIdOk = false;
			int intContagemPag;
			int intContagemPayment;
			char c;
			string msg_erro_aux;
			string strBraspagTransactionId = "";
			string strBraspagOrderId = "";
			string strPaymentMethod1;
			string strPaymentMethod2;
			string strAmount1;
			string strAmount2;
			string strNumberOfPayments1;
			string strNumberOfPayments2;
			string strCardNumberBegin1;
			string strCardNumberBegin2;
			string strCardNumberEnd1;
			string strCardNumberEnd2;
			BraspagPag pag;
			BraspagGetOrderIdDataResponse rGOID = null;
			BraspagGetOrderDataResponse rGOD = null;
			#endregion

			atualizouBraspagTransactionId = false;
			msg_erro = "";
			try
			{
				#region [ A informação 'BraspagTransactionId' já está armazenada? ]
				if (payment.resp_PaymentDataResponse_BraspagTransactionId != null)
				{
					if (payment.resp_PaymentDataResponse_BraspagTransactionId.Length > 0) return true;
				}
				#endregion

				#region [ Para fazer a consulta GetOrderIdData é necessário ter o campo 'OrderId' ]
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);
				if (pag.req_OrderData_OrderId == null)
				{
					msg_erro = "Não é possível realizar a consulta GetOrderIdData porque não há informação armazenada do campo 'OrderId'";
					return false;
				}
				if (pag.req_OrderData_OrderId.Trim().Length == 0)
				{
					msg_erro = "Não é possível realizar a consulta GetOrderIdData porque o campo 'OrderId' está vazio no banco de dados";
					return false;
				}
				#endregion

				#region [ 'OrderId' é único? ]
				// SE HOUVER MAIS DO QUE UMA TRANSAÇÃO C/ O MESMO VALOR DE 'OrderId' NÃO SERÁ POSSÍVEL DETERMINAR A QUAL DELAS SE
				// REFERE A RESPOSTA RETORNADA PELA CONSULTA 'GetOrderIdData'.
				// PORTANTO, NESTE CASO OPTOU-SE POR NÃO FAZER A CONSULTA AO INVÉS DE CORRER O RISCO DE EXIBIR UMA INFORMAÇÃO INCONSISTENTE.
				// EX: A PRIMEIRA TENTATIVA DE PAGAMENTO FALHOU DE FORMA QUE O CAMPO 'BraspagTransactionId' NÃO RETORNOU DA BRASPAG OU NÃO FOI GRAVADO CORRETAMENTE NO BD.
				//     A SEGUNDA TENTATIVA TAMBÉM FALHOU DA MESMA MANEIRA.
				//     A TERCEIRA TENTATIVA FOI BEM-SUCEDIDA.
				//     SE AS 3 TRANSAÇÕES POSSUÍREM O MESMO VALOR DE 'OrderId', A CONSULTA 'GetOrderIdData' FEITA P/ A TENTATIVA 1 OU 2 PODERÁ
				//     RETORNAR O 'BraspagTransactionId' DA TENTATIVA 3.
				//     O USO DESSE 'BraspagTransactionId' POSTERIORMENTE NA CONSULTA 'GetTransactionData' CAUSARIA UM ENTENDIMENTO ERRADO DE QUE HOUVE MAIS
				//     DO QUE UMA TRANSAÇÃO BEM-SUCEDIDA.
				intContagemPag = BraspagDAO.contagemRequisicoesPagByCampoOrderId(pag.req_OrderData_OrderId, out msg_erro_aux);
				if (intContagemPag > 1)
				{
					msg_erro = "Não é possível realizar a consulta GetOrderIdData porque há mais do que uma transação com o mesmo valor de OrderId";
					return false;
				}
				#endregion

				#region [ Executa consulta GetOrderIdData ]
				rGOID = executaConsultaGetOrderIdData(payment, out msg_erro_aux);
				#endregion

				#region [ Contagem de transações de pagamento para o mesmo OrderId ]
				intContagemPayment = BraspagDAO.contagemTransacoesPaymentByCampoOrderId(pag.req_OrderData_OrderId, out msg_erro_aux);
				#endregion

				#region [ Executa consulta GetOrderData? ]
				// Se houver mais do que uma transação de pagamento associada ao mesmo 'OrderId' será necessário realizar uma consulta 'GetOrderData'.
				// Isso ocorre porque a consulta 'GetOrderIdData' retorna uma coleção de respostas apenas com os campos BraspagOrderId e BraspagTransactionId,
				// portanto não é possível determinar a qual transação de pagamento se refere o BraspagTransactionId retornado.
				// Por outro lado, a consulta GetOrderData retorna uma coleção de repostas com vários dados suficientes p/ identificar a transação de pagamento,
				// mas para realizar a consulta é necessário informar o campo BraspagOrderId.
				if (rGOID != null)
				{
					if (rGOID.OrderIdDataCollection.Count == 1)
					{
						if (rGOID.OrderIdDataCollection[0].BraspagOrderId != null) strBraspagOrderId = rGOID.OrderIdDataCollection[0].BraspagOrderId.Trim();

						// IMPORTANTE: QUANDO UM PAGAMENTO É FEITO ATRAVÉS DE 2 OU MAIS CARTÕES ENVIADOS NA MESMA REQUISIÇÃO, ESTA
						// ==========  CONSULTA RETORNA VÁRIOS 'BraspagTransactionId'
						// EX:
						// 		<BraspagTransactionId>
						//			<guid>2bdcc5f5-4331-4123-89f4-864d8203ca97</guid>
						//			<guid>7236d9e3-88f7-4627-856c-948302c4d41c</guid>
						//		</BraspagTransactionId>
						if (rGOID.OrderIdDataCollection[0].BraspagTransactionId.Count == 1)
						{
							if (rGOID.OrderIdDataCollection[0].BraspagTransactionId[0] != null)
							{
								if (rGOID.OrderIdDataCollection[0].BraspagTransactionId[0].Trim().Length > 0)
								{
									blnBraspagTransactionIdOk = true;
									strBraspagTransactionId = rGOID.OrderIdDataCollection[0].BraspagTransactionId[0].Trim();
								}
							}
						}
					}
				}

				if (intContagemPayment > 1) blnBraspagTransactionIdOk = false;

				if ((!blnBraspagTransactionIdOk) && (strBraspagOrderId.Length > 0))
				{
					rGOD = executaConsultaGetOrderData(payment, strBraspagOrderId, out msg_erro_aux);
					if (rGOD != null)
					{
						if (rGOD.TransactionDataCollection.Count > 0)
						{
							foreach (BraspagOrderTransactionDataResponse item in rGOD.TransactionDataCollection)
							{
								if (item.BraspagTransactionId == null) continue;
								if (item.BraspagTransactionId.Trim().Length == 0) continue;

								strPaymentMethod1 = (item.PaymentMethod == null ? "" : item.PaymentMethod);
								strPaymentMethod2 = (payment.req_PaymentDataRequest_PaymentMethod == null ? "" : payment.req_PaymentDataRequest_PaymentMethod);

								strAmount1 = (item.Amount == null ? "" : item.Amount);
								strAmount2 = (payment.req_PaymentDataRequest_Amount == null ? "" : payment.req_PaymentDataRequest_Amount);

								strNumberOfPayments1 = (item.NumberOfPayments == null ? "" : item.NumberOfPayments);
								strNumberOfPayments2 = (payment.req_PaymentDataRequest_NumberOfPayments == null ? "" : payment.req_PaymentDataRequest_NumberOfPayments);

								strCardNumberBegin1 = "";
								strCardNumberEnd1 = "";
								if (item.MaskedCardNumber != null)
								{
									for (int i = 0; i < item.MaskedCardNumber.Length; i++)
									{
										c = item.MaskedCardNumber[i];
										if (!Global.isDigit(c)) break;
										strCardNumberBegin1 += c;
									}
									for (int i = (item.MaskedCardNumber.Length - 1); i >= 0; i--)
									{
										c = item.MaskedCardNumber[i];
										if (!Global.isDigit(c)) break;
										strCardNumberEnd1 = c + strCardNumberEnd1;
									}
								}

								strCardNumberBegin2 = "";
								strCardNumberEnd2 = "";
								if (payment.req_PaymentDataRequest_CardNumber != null)
								{
									if (strCardNumberBegin1.Length > 0) strCardNumberBegin2 = Texto.leftStr(payment.req_PaymentDataRequest_CardNumber, strCardNumberBegin1.Length);
									if (strCardNumberEnd1.Length > 0) strCardNumberEnd2 = Texto.rightStr(payment.req_PaymentDataRequest_CardNumber, strCardNumberEnd1.Length);
								}

								// É a mesma transação?
								if (strPaymentMethod1.Equals(strPaymentMethod2)
									&&
									strAmount1.Equals(strAmount2)
									&&
									strNumberOfPayments1.Equals(strNumberOfPayments2)
									&&
									strCardNumberBegin1.Equals(strCardNumberBegin2)
									&&
									strCardNumberEnd1.Equals(strCardNumberEnd2))
								{
									blnBraspagTransactionIdOk = true;
									strBraspagTransactionId = item.BraspagTransactionId;
									break;
								}
							}
						}
					}
				}
				#endregion

				#region [ Atualiza o campo 'BraspagTransactionId' no banco de dados ]
				if (blnBraspagTransactionIdOk)
				{
					if (BraspagDAO.updatePagPaymentBraspagTransactionId(payment.id, strBraspagTransactionId, out msg_erro_aux))
					{
						atualizouBraspagTransactionId = true;
					}
					else
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ consultaGetOrderIdData ]
		/// <summary>
		/// Executa a chamada ao método GetOrderIdData e retorna os dados no objeto da classe BraspagGetOrderIdDataResponse.
		/// Esta rotina NÃO grava dados nas tabelas t_PAGTO_GW_PAG_OP_COMPLEMENTAR e t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, pois foi desenvolvida para ser usada em pedidos do e-commerce,
		/// caso em que o ERP não possui os dados originais da transação.
		/// </summary>
		/// <param name="merchantId">Chave MerchantId</param>
		/// <param name="orderId">Nº do pedido</param>
		/// <param name="msg_erro">Mensagem do erro ocorrido, se houver</param>
		/// <returns>Retorna objeto da classe BraspagGetOrderIdDataResponse</returns>
		public static BraspagGetOrderIdDataResponse consultaGetOrderIdData(string merchantId, string orderId, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.consultaGetOrderIdData()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.GetOrderIdData;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			BraspagGetOrderIdData trx;
			BraspagGetOrderIdDataResponse rRESP = null;
			#endregion

			msg_erro = "";
			try
			{
				trx = criaGetOrderIdData(merchantId, orderId);
				xmlReqSoap = montaXmlRequisicaoSoapGetOrderIdData(trx);

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta vazia? ]
				if (xmlRespSoap.Trim().Length == 0)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta vazia";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				rRESP = decodificaXmlGetOrderIdDataResponse(xmlRespSoap, out msg_erro_aux);

				if (rRESP == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlRespSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Há mensagem de erro na resposta? ]
				if (rRESP.ErrorReportDataCollection.Count > 0)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlRespSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Status da resposta é de sucesso? ]
				if (rRESP.Success != null)
				{
					if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
				}

				if (!blnRespostaOk) return null;
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = $"MerchantId={merchantId}, OrderId={orderId}";
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ executaConsultaGetOrderIdData ]
		/// <summary>
		/// Executa a chamada ao método GetOrderIdData e retorna os dados no objeto da classe BraspagGetOrderIdDataResponse.
		/// Esta rotina grava dados nas tabelas t_PAGTO_GW_PAG_OP_COMPLEMENTAR e t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML
		/// </summary>
		/// <param name="payment">Objeto do tipo BraspagPagPayment com os dados da transação de pagamento</param>
		/// <param name="msg_erro">Mensagem do erro ocorrido, se houver</param>
		/// <returns></returns>
		public static BraspagGetOrderIdDataResponse executaConsultaGetOrderIdData(BraspagPagPayment payment, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.executaConsultaGetOrderIdData()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.GetOrderIdData;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			BraspagPag pag;
			BraspagGetOrderIdData trx;
			BraspagGetOrderIdDataResponse rRESP = null;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			#endregion

			msg_erro = "";
			try
			{
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);
				trx = criaGetOrderIdData(pag.req_OrderData_MerchantId, pag.req_OrderData_OrderId);
				xmlReqSoap = montaXmlRequisicaoSoapGetOrderIdData(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = payment.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = payment.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxSelecionada.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				opComplXmlRx = new BraspagPagOpComplementarXml();
				opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
				opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlGetOrderIdDataResponse(xmlRespSoap, out msg_erro_aux);

					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return null;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					#region [ BraspagTransactionId ]
					if (rRESP.OrderIdDataCollection.Count == 1)
					{
						if (rRESP.OrderIdDataCollection[0].BraspagTransactionId.Count == 1)
						{
							opCompl.resp_BraspagTransactionId = rRESP.OrderIdDataCollection[0].BraspagTransactionId[0];
						}
					}
					#endregion

					#region [ Há mensagem de erro na resposta? ]
					if (rRESP.ErrorReportDataCollection.Count > 0)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return null;
					}
					#endregion

					if (rRESP.Success != null)
					{
						if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
					}
				}
				#endregion

				if (blnRespostaOk)
				{
					#region [ Atualiza tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com dados da resposta ]
					if (!BraspagDAO.updatePagOpComplementarGetOrderIdDataResp(opCompl, out msg_erro_aux))
					{
						Global.gravaLogAtividade("Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR + " (id=" + opCompl.id.ToString() + ")\n" + msg_erro_aux);
					}
					#endregion
				}

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ executaConsultaGetOrderData ]
		public static BraspagGetOrderDataResponse executaConsultaGetOrderData(BraspagPagPayment payment, string BraspagOrderId, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.executaConsultaGetOrderData()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.GetOrderData;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			BraspagPag pag;
			BraspagGetOrderData trx;
			BraspagGetOrderDataResponse rRESP = null;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			#endregion

			msg_erro = "";
			try
			{
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);
				trx = criaGetOrderData(pag.req_OrderData_MerchantId, BraspagOrderId);
				xmlReqSoap = montaXmlRequisicaoSoapGetOrderData(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = payment.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = payment.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxSelecionada.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				opComplXmlRx = new BraspagPagOpComplementarXml();
				opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
				opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlGetOrderDataResponse(xmlRespSoap, out msg_erro_aux);

					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return null;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					#region [ BraspagTransactionId ]
					if (rRESP.TransactionDataCollection.Count == 1)
					{
						opCompl.resp_BraspagTransactionId = rRESP.TransactionDataCollection[0].BraspagTransactionId;
					}
					#endregion

					#region [ Há mensagem de erro na resposta? ]
					if (rRESP.ErrorReportDataCollection.Count > 0)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return null;
					}
					#endregion

					if (rRESP.Success != null)
					{
						if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
					}
				}
				#endregion

				if (blnRespostaOk)
				{
					#region [ Atualiza tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com dados da resposta ]
					if (!BraspagDAO.updatePagOpComplementarGetOrderDataResp(opCompl, out msg_erro_aux))
					{
						Global.gravaLogAtividade("Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR + " (id=" + opCompl.id.ToString() + ")\n" + msg_erro_aux);
					}
					#endregion
				}

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ consultaGetTransactionData ]
		/// <summary>
		/// Executa a chamada ao método GetTransactionData e retorna os dados no objeto da classe BraspagGetTransactionDataResponse.
		/// Esta rotina NÃO grava dados nas tabelas t_PAGTO_GW_PAG_PAYMENT, t_PAGTO_GW_PAG_OP_COMPLEMENTAR e t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML,
		/// pois foi desenvolvida para ser usada em pedidos do e-commerce, caso em que o ERP não possui os dados originais da transação.
		/// </summary>
		/// <param name="merchantId">Chave MerchantId</param>
		/// <param name="braspagTransactionId">BraspagTransactionId</param>
		/// <param name="msg_erro">Mensagem do erro ocorrido, se houver</param>
		/// <returns>Retorna objeto da classe BraspagGetTransactionDataResponse</returns>
		public static BraspagGetTransactionDataResponse consultaGetTransactionData(string merchantId, string braspagTransactionId, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.consultaGetTransactionData()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.GetTransactionData;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			DateTime dtCapturedDate = DateTime.MinValue;
			DateTime dtVoidedDate = DateTime.MinValue;
			BraspagGetTransactionData trx;
			BraspagGetTransactionDataResponse rRESP = null;
			#endregion

			msg_erro = "";
			try
			{
				trx = criaGetTransactionData(merchantId, braspagTransactionId);
				xmlReqSoap = montaXmlRequisicaoSoapGetTransactionData(trx);

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta vazia? ]
				if (xmlRespSoap.Trim().Length == 0)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta vazia";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				rRESP = decodificaXmlGetTransactionDataResponse(xmlRespSoap, out msg_erro_aux);

				if (rRESP == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlRespSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Há mensagem de erro na resposta? ]
				if (rRESP.ErrorReportDataCollection.Count > 0)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlRespSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Status da resposta é de sucesso? ]
				if (rRESP.Success != null)
				{
					if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
				}

				if (!blnRespostaOk) return null;
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = $"MerchantId={merchantId}, BraspagTransactionId={braspagTransactionId}";
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ consultaGetBoletoData ]
		/// <summary>
		/// Executa a chamada ao método GetBoletoData e retorna os dados no objeto da classe BraspagGetBoletoDataResponse.
		/// Esta rotina NÃO grava dados nas tabelas t_PAGTO_GW_PAG_OP_COMPLEMENTAR e t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, pois foi desenvolvida para ser usada em pedidos do e-commerce,
		/// caso em que o ERP não possui os dados originais da transação.
		/// </summary>
		/// <param name="merchantId">Chave MerchantId</param>
		/// <param name="braspagTransactionId">BraspagTransactionId</param>
		/// <param name="msg_erro">Mensagem do erro ocorrido, se houver</param>
		/// <returns>Retorna objeto da classe BraspagGetBoletoDataResponse</returns>
		public static BraspagGetBoletoDataResponse consultaGetBoletoData(string merchantId, string braspagTransactionId, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.consultaGetBoletoData()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.GetBoletoData;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			BraspagGetBoletoData trx;
			BraspagGetBoletoDataResponse rRESP = null;
			#endregion

			msg_erro = "";
			try
			{
				trx = criaGetBoletoData(merchantId, braspagTransactionId);
				xmlReqSoap = montaXmlRequisicaoSoapGetBoletoData(trx);

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Resposta vazia? ]
				if (xmlRespSoap.Trim().Length == 0)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta vazia";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				rRESP = decodificaXmlGetBoletoDataResponse(xmlRespSoap, out msg_erro_aux);

				if (rRESP == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlRespSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Há mensagem de erro na resposta? ]
				if (rRESP.ErrorReportDataCollection.Count > 0)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlRespSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return null;
				}
				#endregion

				#region [ Status da resposta é de sucesso? ]
				if (rRESP.Success != null)
				{
					if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
				}

				if (!blnRespostaOk) return null;
				#endregion

				return rRESP;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = $"MerchantId={merchantId}, BraspagTransactionId={braspagTransactionId}";
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ processaConsultaGetTransactionData ]
		/// <summary>
		/// Executa a chamada ao método GetTransactionData e atualiza os dados no banco de dados.
		/// </summary>
		/// <param name="payment">Objeto da classe BraspagPagPayment com os dados da transação</param>
		/// <param name="ult_GlobalStatus_original">Código GlobalStatus original, antes da consulta</param>
		/// <param name="ult_GlobalStatus_novo">Código GlobalStatus atualizado, após a consulta</param>
		/// <param name="msg_erro">Mensagem do erro ocorrido, se houver</param>
		/// <returns>Retorna true se o processamento for bem sucedido e false em caso de falha</returns>
		public static bool processaConsultaGetTransactionData(BraspagPagPayment payment, out string ult_GlobalStatus_original, out string ult_GlobalStatus_novo, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.processaConsultaGetTransactionData()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.GetTransactionData;
			string strMsg;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			bool blnBraspagTransactionIdOk = false;
			bool blnAtualizouBraspagTransactionId = false;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			DateTime dtCapturedDate = DateTime.MinValue;
			DateTime dtVoidedDate = DateTime.MinValue;
			BraspagPag pag;
			BraspagGetTransactionData trx;
			BraspagGetTransactionDataResponse rRESP = null;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			BraspagUpdatePagPaymentGetTransactionDataResponse rUpdate;
			BraspagPagPayment paymentAtualizado;
			#endregion

			ult_GlobalStatus_original = payment.ult_GlobalStatus;
			ult_GlobalStatus_novo = "";
			msg_erro = "";
			try
			{
				if (payment.resp_PaymentDataResponse_BraspagTransactionId != null)
				{
					if (payment.resp_PaymentDataResponse_BraspagTransactionId.Trim().Length > 0) blnBraspagTransactionIdOk = true;
				}

				if (!blnBraspagTransactionIdOk)
				{
					if (!verificaPreRequisitoBraspagTransactionId(payment, out blnAtualizouBraspagTransactionId, out msg_erro_aux))
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar obter o valor de 'BraspagTransactionId'\n" + msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
				}

				if (blnAtualizouBraspagTransactionId)
				{
					paymentAtualizado = BraspagDAO.getBraspagPagPaymentById(payment.id, out msg_erro_aux);
				}
				else
				{
					paymentAtualizado = payment;
				}

				if (paymentAtualizado.resp_PaymentDataResponse_BraspagTransactionId.Trim().Length == 0)
				{
					msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!";
					return false;
				}

				pag = BraspagDAO.getBraspagPagById(paymentAtualizado.id_pagto_gw_pag, out msg_erro_aux);
				trx = criaGetTransactionData(pag.req_OrderData_MerchantId, paymentAtualizado.resp_PaymentDataResponse_BraspagTransactionId);
				xmlReqSoap = montaXmlRequisicaoSoapGetTransactionData(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = paymentAtualizado.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = paymentAtualizado.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxSelecionada.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				opCompl.req_BraspagTransactionId = trx.BraspagTransactionId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				if (blnEnviouOk)
				{
					opComplXmlRx = new BraspagPagOpComplementarXml();
					opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
					opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
					opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
					opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
					if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
					}
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlGetTransactionDataResponse(xmlRespSoap, out msg_erro_aux);

					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					#region [ BraspagTransactionId ]
					if (rRESP.BraspagTransactionId != null) opCompl.resp_BraspagTransactionId = rRESP.BraspagTransactionId;
					#endregion

					#region [ AuthorizationCode ]
					if (rRESP.AuthorizationCode != null) opCompl.resp_AuthorizationCode = rRESP.AuthorizationCode;
					#endregion

					#region [ ProofOfSale ]
					if (rRESP.ProofOfSale != null) opCompl.resp_ProofOfSale = rRESP.ProofOfSale;
					#endregion

					#region [ Status ]
					if (rRESP.Status != null) opCompl.resp_Status = rRESP.Status;
					#endregion

					#region [ Há mensagem de erro na resposta? ]
					if (rRESP.ErrorReportDataCollection.Count > 0)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}
					#endregion

					if (rRESP.Success != null)
					{
						if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
					}
				}
				#endregion

				if (blnRespostaOk)
				{
					#region [ Atualiza tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com dados da resposta ]
					if (!BraspagDAO.updatePagOpComplementarGetTransactionDataResp(opCompl, out msg_erro_aux))
					{
						Global.gravaLogAtividade("Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR + " (id=" + opCompl.id.ToString() + ")\n" + msg_erro_aux);
					}
					#endregion

					#region [ Atualiza tabela T_PAGTO_GW_PAG_PAYMENT com dados da resposta ]
					if (rRESP.CapturedDate != null)
					{
						if (rRESP.CapturedDate.Trim().Length > 0)
						{
							dtCapturedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(rRESP.CapturedDate);
						}
					}

					if (rRESP.VoidedDate != null)
					{
						if (rRESP.VoidedDate.Trim().Length > 0)
						{
							dtVoidedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(rRESP.VoidedDate);
						}
					}

					rUpdate = new BraspagUpdatePagPaymentGetTransactionDataResponse();
					rUpdate.id_pagto_gw_pag_payment = paymentAtualizado.id;
					rUpdate.ult_GlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(rRESP.Status);
					ult_GlobalStatus_novo = rUpdate.ult_GlobalStatus;
					rUpdate.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
					rUpdate.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
					rUpdate.resp_CapturedDate = dtCapturedDate;
					rUpdate.resp_VoidedDate = dtVoidedDate;

					if (!BraspagDAO.updatePagPaymentGetTransactionDataResp(rUpdate, out msg_erro_aux))
					{
						Global.gravaLogAtividade("Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " (id=" + paymentAtualizado.id.ToString() + ")\n" + msg_erro_aux);
					}
					#endregion

					if (!ult_GlobalStatus_original.Equals(ult_GlobalStatus_novo))
					{
						strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + paymentAtualizado.id.ToString() + "): " + ult_GlobalStatus_original + "=>" + ult_GlobalStatus_novo;
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = strMsg;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
				}

				return blnRespostaOk;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ processaRequisicaoCaptureCreditCardTransaction ]
		public static bool processaRequisicaoCaptureCreditCardTransaction(BraspagPagPayment payment, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.processaRequisicaoCaptureCreditCardTransaction()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction;
			bool blnEnviouOk;
			bool blnSucesso;
			bool blnCapturaConfirmada = false;
			int id_emailsndsvc_mensagem;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string msg_erro_last_op;
			string xmlReqSoap;
			string xmlRespSoap;
			String strMsg;
			String strSubject;
			String strBody;
			BraspagPag pag;
			BraspagCaptureCreditCardTransaction trx;
			BraspagCaptureCreditCardTransactionResponse rRESP;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			BraspagTransactionDataRequest trxDataRequest;
			List<BraspagTransactionDataRequest> vTrxDataRequest = new List<BraspagTransactionDataRequest>();
			BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseSucesso rUpdatePaymentCaptureSucesso;
			BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseFalha rUpdatePaymentCaptureFalha = null;
			#endregion

			msg_erro = "";
			try
			{
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);

				trxDataRequest = new BraspagTransactionDataRequest(payment.resp_PaymentDataResponse_BraspagTransactionId, payment.req_PaymentDataRequest_Amount, payment.req_PaymentDataRequest_ServiceTaxAmount);
				vTrxDataRequest.Add(trxDataRequest);
				trx = criaCaptureCreditCardTransaction(pag.req_OrderData_MerchantId, vTrxDataRequest);
				xmlReqSoap = montaXmlRequisicaoSoapCaptureCreditCardTransaction(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = payment.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = payment.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxSelecionada.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				opCompl.req_BraspagTransactionId = payment.resp_PaymentDataResponse_BraspagTransactionId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				if (blnEnviouOk)
				{
					opComplXmlRx = new BraspagPagOpComplementarXml();
					opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
					opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
					opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
					opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
					if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
					}
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlCaptureCreditCardTransactionResponse(xmlRespSoap, out msg_erro_aux);
					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					if (rRESP.TransactionDataCollection.Count > 0)
					{
						#region [ BraspagTransactionId ]
						if (rRESP.TransactionDataCollection[0].BraspagTransactionId != null) opCompl.resp_BraspagTransactionId = rRESP.TransactionDataCollection[0].BraspagTransactionId;
						#endregion

						#region [ AuthorizationCode ]
						if (rRESP.TransactionDataCollection[0].AuthorizationCode != null) opCompl.resp_AuthorizationCode = rRESP.TransactionDataCollection[0].AuthorizationCode;
						#endregion

						#region [ ProofOfSale ]
						if (rRESP.TransactionDataCollection[0].ProofOfSale != null) opCompl.resp_ProofOfSale = rRESP.TransactionDataCollection[0].ProofOfSale;
						#endregion

						#region [ Status ]
						if (rRESP.TransactionDataCollection[0].Status != null) opCompl.resp_Status = rRESP.TransactionDataCollection[0].Status;
						#endregion
					}

					if (!BraspagDAO.updatePagOpComplementarCaptureCreditCardResp(opCompl, out msg_erro_aux))
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar atualizar o registro da tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;
						svcLog.complemento_1 = Global.serializaObjectToXml(opCompl);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}

					if (rRESP.TransactionDataCollection.Count > 0)
					{
						if (rRESP.TransactionDataCollection[0].Status != null)
						{
							if (rRESP.TransactionDataCollection[0].Status.Equals(Global.Cte.Braspag.Pagador.CaptureCreditCardTransactionResponseStatus.CAPTURE_CONFIRMED.GetValue()))
							{
								blnCapturaConfirmada = true;
								rUpdatePaymentCaptureSucesso = new BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseSucesso();
								rUpdatePaymentCaptureSucesso.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentCaptureSucesso.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentCaptureSucesso.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentCaptureSucesso.ult_GlobalStatus = Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue();
								rUpdatePaymentCaptureSucesso.captura_confirmada_status = 1;
								rUpdatePaymentCaptureSucesso.captura_confirmada_data = DateTime.Now.Date;
								rUpdatePaymentCaptureSucesso.captura_confirmada_data_hora = DateTime.Now;
								rUpdatePaymentCaptureSucesso.captura_confirmada_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentCaptureSucesso.resp_CapturedDate = DateTime.Now.Date;

								if (!BraspagDAO.updatePagPaymentCaptureCreditCardRespSucesso(rUpdatePaymentCaptureSucesso, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " após uma captura bem sucedida\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentCaptureSucesso);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar confirmar a captura [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar confirmar a captura do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
							}
							else
							{
								rUpdatePaymentCaptureFalha = new BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseFalha();
								rUpdatePaymentCaptureFalha.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentCaptureFalha.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentCaptureFalha.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentCaptureFalha.captura_confirmada_erro_status = 1;
								rUpdatePaymentCaptureFalha.captura_confirmada_erro_data = DateTime.Now.Date;
								rUpdatePaymentCaptureFalha.captura_confirmada_erro_data_hora = DateTime.Now;
								rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem = "Status: " + rRESP.TransactionDataCollection[0].Status;
								if (
										((rRESP.TransactionDataCollection[0].ReturnCode ?? "").Length > 0)
										||
										((rRESP.TransactionDataCollection[0].ReturnMessage ?? "").Length > 0)
									)
								{
									if (rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem.Length > 0) rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem += "\r\n";
									rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem += "ReturnCode: " + (rRESP.TransactionDataCollection[0].ReturnCode ?? "") + ", ReturnMessage: " + (rRESP.TransactionDataCollection[0].ReturnMessage ?? "");
								}

								if (rRESP.TransactionDataCollection[0].ErrorReportDataCollection.Count > 0)
								{
									if (rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem.Length > 0) rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem += "\r\n";
									rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem += "ErrorCode: " + (rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorCode ?? "")
																									+ ", ErrorMessage: " + (rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorMessage ?? "");
								}

								if (!BraspagDAO.updatePagPaymentCaptureCreditCardRespFalha(rUpdatePaymentCaptureFalha, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " após uma tentativa de captura com falha\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentCaptureFalha);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar dados sobre tentativa fracassada de captura [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar dados sobre tentativa fracassada de captura de pagamento do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}

								#region [ Envia e-mail informando a falha na captura ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar confirmar a captura [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha ao tentar confirmar a captura de pagamento do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + (rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem ?? "");
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion
							}
						}
					}
				}
				#endregion

				if (!blnCapturaConfirmada)
				{
					if (rUpdatePaymentCaptureFalha != null)
					{
						if ((msg_erro.Length > 0) && ((rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem ?? "").Length > 0)) msg_erro += "\n";
						msg_erro += (rUpdatePaymentCaptureFalha.captura_confirmada_erro_mensagem ?? "");
					}
					return false;
				}

				#region [ Registra o pagamento no pedido ]
				blnSucesso = false;
				BD.iniciaTransacao();
				try
				{
					blnSucesso = BraspagDAO.registraPagamentoNoPedido(trxSelecionada, payment.id, out msg_erro_aux);
					if (!blnSucesso) msg_erro = msg_erro_aux;
				}
				catch (Exception ex)
				{
					blnSucesso = false;
					msg_erro = ex.ToString();
				}
				finally
				{
					if (blnSucesso)
					{
						BD.commitTransacao();
					}
					else
					{
						BD.rollbackTransacao();
					}
				}

				if (!blnSucesso)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = "Falha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxSelecionada.GetMethodName() + "\r\n" + msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(payment);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o pagamento no pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxSelecionada.GetMethodName() + "\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ processaRequisicaoVoidCreditCardTransaction ]
		public static bool processaRequisicaoVoidCreditCardTransaction(BraspagPagPayment payment, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.processaRequisicaoVoidCreditCardTransaction()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction;
			bool blnEnviouOk;
			bool blnSucesso;
			int id_emailsndsvc_mensagem;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string msg_erro_last_op;
			string xmlReqSoap;
			string xmlRespSoap;
			string strSubject;
			string strBody;
			string strMsg;
			BraspagPag pag;
			BraspagVoidCreditCardTransaction trx;
			BraspagVoidCreditCardTransactionResponse rRESP;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			BraspagTransactionDataRequest trxDataRequest;
			List<BraspagTransactionDataRequest> vTrxDataRequest = new List<BraspagTransactionDataRequest>();
			BraspagUpdatePagPaymentVoidCreditCardTransactionResponseSucesso rUpdatePaymentVoidSucesso;
			BraspagUpdatePagPaymentVoidCreditCardTransactionResponseFalha rUpdatePaymentVoidFalha;
			#endregion

			msg_erro = "";
			try
			{
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);

				trxDataRequest = new BraspagTransactionDataRequest(payment.resp_PaymentDataResponse_BraspagTransactionId, payment.req_PaymentDataRequest_Amount, payment.req_PaymentDataRequest_ServiceTaxAmount);
				vTrxDataRequest.Add(trxDataRequest);
				trx = criaVoidCreditCardTransaction(pag.req_OrderData_MerchantId, vTrxDataRequest);
				xmlReqSoap = montaXmlRequisicaoSoapVoidCreditCardTransaction(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = payment.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = payment.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxSelecionada.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				opCompl.req_BraspagTransactionId = payment.resp_PaymentDataResponse_BraspagTransactionId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				if (blnEnviouOk)
				{
					opComplXmlRx = new BraspagPagOpComplementarXml();
					opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
					opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
					opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
					opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
					if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
					}
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlVoidCreditCardTransactionResponse(xmlRespSoap, out msg_erro_aux);
					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					if (rRESP.TransactionDataCollection.Count > 0)
					{
						#region [ BraspagTransactionId ]
						if (rRESP.TransactionDataCollection[0].BraspagTransactionId != null) opCompl.resp_BraspagTransactionId = rRESP.TransactionDataCollection[0].BraspagTransactionId;
						#endregion

						#region [ AuthorizationCode ]
						if (rRESP.TransactionDataCollection[0].AuthorizationCode != null) opCompl.resp_AuthorizationCode = rRESP.TransactionDataCollection[0].AuthorizationCode;
						#endregion

						#region [ ProofOfSale ]
						if (rRESP.TransactionDataCollection[0].ProofOfSale != null) opCompl.resp_ProofOfSale = rRESP.TransactionDataCollection[0].ProofOfSale;
						#endregion

						#region [ Status ]
						if (rRESP.TransactionDataCollection[0].Status != null) opCompl.resp_Status = rRESP.TransactionDataCollection[0].Status;
						#endregion
					}

					if (!BraspagDAO.updatePagOpComplementarVoidCreditCardResp(opCompl, out msg_erro_aux))
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar atualizar o registro da tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;
						svcLog.complemento_1 = Global.serializaObjectToXml(opCompl);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}

					if (rRESP.TransactionDataCollection.Count > 0)
					{
						if (rRESP.TransactionDataCollection[0].Status != null)
						{
							if (rRESP.TransactionDataCollection[0].Status.Equals(Global.Cte.Braspag.Pagador.VoidCreditCardTransactionResponseStatus.VOID_CONFIRMED.GetValue()))
							{
								rUpdatePaymentVoidSucesso = new BraspagUpdatePagPaymentVoidCreditCardTransactionResponseSucesso();
								rUpdatePaymentVoidSucesso.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentVoidSucesso.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentVoidSucesso.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentVoidSucesso.ult_GlobalStatus = Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURA_CANCELADA.GetValue();
								rUpdatePaymentVoidSucesso.voided_status = 1;
								rUpdatePaymentVoidSucesso.voided_data = DateTime.Now.Date;
								rUpdatePaymentVoidSucesso.voided_data_hora = DateTime.Now;
								rUpdatePaymentVoidSucesso.voided_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentVoidSucesso.resp_VoidedDate = DateTime.Now.Date;

								if (!BraspagDAO.updatePagPaymentVoidCreditCardRespSucesso(rUpdatePaymentVoidSucesso, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " após um cancelamento (void) bem sucedido\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentVoidSucesso);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar cancelar (void) a transação [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar cancelar (void) a transação do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
							}
							else
							{
								rUpdatePaymentVoidFalha = new BraspagUpdatePagPaymentVoidCreditCardTransactionResponseFalha();
								rUpdatePaymentVoidFalha.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentVoidFalha.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentVoidFalha.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentVoidFalha.voided_erro_status = 1;
								rUpdatePaymentVoidFalha.voided_erro_data = DateTime.Now.Date;
								rUpdatePaymentVoidFalha.voided_erro_data_hora = DateTime.Now;
								rUpdatePaymentVoidFalha.voided_erro_mensagem = "";
								if (
										((rRESP.TransactionDataCollection[0].ReturnCode ?? "").Length > 0)
										||
										((rRESP.TransactionDataCollection[0].ReturnMessage ?? "").Length > 0)
									)
								{
									rUpdatePaymentVoidFalha.voided_erro_mensagem = (rRESP.TransactionDataCollection[0].ReturnCode ?? "") + " - " + (rRESP.TransactionDataCollection[0].ReturnMessage ?? "");
								}

								if (rRESP.TransactionDataCollection[0].ErrorReportDataCollection.Count > 0)
								{
									if (rUpdatePaymentVoidFalha.voided_erro_mensagem.Length > 0) rUpdatePaymentVoidFalha.voided_erro_mensagem += "\r\n";
									rUpdatePaymentVoidFalha.voided_erro_mensagem += (rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorCode == null ? "" : rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorCode)
																					+ " - " +
																					(rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorMessage == null ? "" : rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorMessage);
								}

								if (!BraspagDAO.updatePagPaymentVoidCreditCardRespFalha(rUpdatePaymentVoidFalha, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " após uma tentativa de cancelamento (void) com falha\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentVoidFalha);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar cancelar (void) a transação [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar cancelar (void) a transação do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
							}
						}
					}
				}
				#endregion

				#region [ Registra o pagamento no pedido ]
				blnSucesso = false;
				BD.iniciaTransacao();
				try
				{
					blnSucesso = BraspagDAO.registraPagamentoNoPedido(trxSelecionada, payment.id, out msg_erro_aux);
					if (!blnSucesso) msg_erro = msg_erro_aux;
				}
				catch (Exception ex)
				{
					blnSucesso = false;
					msg_erro = ex.ToString();
				}
				finally
				{
					if (blnSucesso)
					{
						BD.commitTransacao();
					}
					else
					{
						BD.rollbackTransacao();
					}
				}

				if (!blnSucesso)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = "Falha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxSelecionada.GetMethodName() + "\r\n" + msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(payment);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o pagamento no pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxSelecionada.GetMethodName() + "\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ processaRequisicaoRefundCreditCardTransaction ]
		public static bool processaRequisicaoRefundCreditCardTransaction(BraspagPagPayment payment, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.processaRequisicaoRefundCreditCardTransaction()";
			Global.Cte.Braspag.Pagador.Transacao trxSelecionada = Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction;
			bool blnEnviouOk;
			bool blnSucesso;
			bool blnEstornoConfirmado = false;
			int id_emailsndsvc_mensagem;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string msg_erro_last_op;
			string xmlReqSoap;
			string xmlRespSoap;
			string strSubject;
			string strBody;
			string strMsg;
			BraspagPag pag;
			BraspagRefundCreditCardTransaction trx;
			BraspagRefundCreditCardTransactionResponse rRESP;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			BraspagTransactionDataRequest trxDataRequest;
			List<BraspagTransactionDataRequest> vTrxDataRequest = new List<BraspagTransactionDataRequest>();
			BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso rUpdatePaymentRefundSucesso;
			BraspagUpdatePagPaymentRefundCreditCardTransactionResponseRefundAccepted rUpdatePaymentRefundAccepted;
			BraspagUpdatePagPaymentRefundCreditCardTransactionResponseFalha rUpdatePaymentRefundFalha;
			#endregion

			msg_erro = "";
			try
			{
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);

				trxDataRequest = new BraspagTransactionDataRequest(payment.resp_PaymentDataResponse_BraspagTransactionId, payment.req_PaymentDataRequest_Amount, payment.req_PaymentDataRequest_ServiceTaxAmount);
				vTrxDataRequest.Add(trxDataRequest);
				trx = criaRefundCreditCardTransaction(pag.req_OrderData_MerchantId, vTrxDataRequest);
				xmlReqSoap = montaXmlRequisicaoSoapRefundCreditCardTransaction(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = payment.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = payment.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxSelecionada.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				opCompl.req_BraspagTransactionId = payment.resp_PaymentDataResponse_BraspagTransactionId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				if (blnEnviouOk)
				{
					opComplXmlRx = new BraspagPagOpComplementarXml();
					opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
					opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
					opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
					opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
					if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
					}
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlRefundCreditCardTransactionResponse(xmlRespSoap, out msg_erro_aux);
					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					if (rRESP.TransactionDataCollection.Count > 0)
					{
						#region [ BraspagTransactionId ]
						if (rRESP.TransactionDataCollection[0].BraspagTransactionId != null) opCompl.resp_BraspagTransactionId = rRESP.TransactionDataCollection[0].BraspagTransactionId;
						#endregion

						#region [ AuthorizationCode ]
						if (rRESP.TransactionDataCollection[0].AuthorizationCode != null) opCompl.resp_AuthorizationCode = rRESP.TransactionDataCollection[0].AuthorizationCode;
						#endregion

						#region [ ProofOfSale ]
						if (rRESP.TransactionDataCollection[0].ProofOfSale != null) opCompl.resp_ProofOfSale = rRESP.TransactionDataCollection[0].ProofOfSale;
						#endregion

						#region [ Status ]
						if (rRESP.TransactionDataCollection[0].Status != null) opCompl.resp_Status = rRESP.TransactionDataCollection[0].Status;
						#endregion
					}

					if (!BraspagDAO.updatePagOpComplementarRefundCreditCardResp(opCompl, out msg_erro_aux))
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar atualizar o registro da tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + "\n" + msg_erro_aux;
						svcLog.complemento_1 = Global.serializaObjectToXml(opCompl);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}

					if (rRESP.TransactionDataCollection.Count > 0)
					{
						if (rRESP.TransactionDataCollection[0].Status != null)
						{
							if (rRESP.TransactionDataCollection[0].Status.Equals(Global.Cte.Braspag.Pagador.RefundCreditCardTransactionResponseStatus.REFUND_CONFIRMED.GetValue()))
							{
								blnEstornoConfirmado = true;
								rUpdatePaymentRefundSucesso = new BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso();
								rUpdatePaymentRefundSucesso.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentRefundSucesso.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentRefundSucesso.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentRefundSucesso.ult_GlobalStatus = Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNADA.GetValue();
								rUpdatePaymentRefundSucesso.refunded_status = 1;
								rUpdatePaymentRefundSucesso.refunded_data = DateTime.Now.Date;
								rUpdatePaymentRefundSucesso.refunded_data_hora = DateTime.Now;
								rUpdatePaymentRefundSucesso.refunded_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentRefundSucesso.resp_VoidedDate = DateTime.Now.Date;

								if (!BraspagDAO.updatePagPaymentRefundCreditCardRespSucesso(rUpdatePaymentRefundSucesso, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " após um estorno (refund) bem sucedido\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentRefundSucesso);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar estornar (refund) a transação [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar estornar (refund) a transação do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
							}
							else if (rRESP.TransactionDataCollection[0].Status.Equals(Global.Cte.Braspag.Pagador.RefundCreditCardTransactionResponseStatus.REFUND_ACCEPTED.GetValue()))
							{
								rUpdatePaymentRefundAccepted = new BraspagUpdatePagPaymentRefundCreditCardTransactionResponseRefundAccepted();
								rUpdatePaymentRefundAccepted.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentRefundAccepted.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentRefundAccepted.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentRefundAccepted.ult_GlobalStatus = Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNO_PENDENTE.GetValue();
								if (!BraspagDAO.updatePagPaymentRefundCreditCardRespRefundAccepted(rUpdatePaymentRefundAccepted, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " com status 'refund accepted'\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentRefundAccepted);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar gravar os dados para estorno pendente [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar gravar os dados para estorno pendente do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
							}
							else
							{
								rUpdatePaymentRefundFalha = new BraspagUpdatePagPaymentRefundCreditCardTransactionResponseFalha();
								rUpdatePaymentRefundFalha.id_pagto_gw_pag_payment = payment.id;
								rUpdatePaymentRefundFalha.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
								rUpdatePaymentRefundFalha.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
								rUpdatePaymentRefundFalha.refunded_erro_status = 1;
								rUpdatePaymentRefundFalha.refunded_erro_data = DateTime.Now.Date;
								rUpdatePaymentRefundFalha.refunded_erro_data_hora = DateTime.Now;
								rUpdatePaymentRefundFalha.refunded_erro_mensagem = "";
								if (
										((rRESP.TransactionDataCollection[0].ReturnCode ?? "").Length > 0)
										||
										((rRESP.TransactionDataCollection[0].ReturnMessage ?? "").Length > 0)
									)
								{
									rUpdatePaymentRefundFalha.refunded_erro_mensagem = (rRESP.TransactionDataCollection[0].ReturnCode ?? "") + " - " + (rRESP.TransactionDataCollection[0].ReturnMessage ?? "");
								}

								if (rRESP.TransactionDataCollection[0].ErrorReportDataCollection.Count > 0)
								{
									if (rUpdatePaymentRefundFalha.refunded_erro_mensagem.Length > 0) rUpdatePaymentRefundFalha.refunded_erro_mensagem += "\r\n";
									rUpdatePaymentRefundFalha.refunded_erro_mensagem += (rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorCode == null ? "" : rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorCode)
																						+ " - " +
																						(rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorMessage == null ? "" : rRESP.TransactionDataCollection[0].ErrorReportDataCollection[0].ErrorMessage);
								}

								if (!BraspagDAO.updatePagPaymentRefundCreditCardRespFalha(rUpdatePaymentRefundFalha, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com os dados da resposta do método Braspag " + trxSelecionada.GetMethodName() + " após uma tentativa de estorno (refund) com falha\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentRefundFalha);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar estornar (refund) a transação [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar estornar (refund) a transação do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
							}
						}
					}
				}
				#endregion

				#region [ Registra o pagamento no pedido ]
				if (blnEstornoConfirmado)
				{
					blnSucesso = false;
					BD.iniciaTransacao();
					try
					{
						blnSucesso = BraspagDAO.registraPagamentoNoPedido(trxSelecionada, payment.id, out msg_erro_aux);
						if (!blnSucesso) msg_erro = msg_erro_aux;
					}
					catch (Exception ex)
					{
						blnSucesso = false;
						msg_erro = ex.ToString();
					}
					finally
					{
						if (blnSucesso)
						{
							BD.commitTransacao();
						}
						else
						{
							BD.rollbackTransacao();
						}
					}

					if (!blnSucesso)
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxSelecionada.GetMethodName() + "\r\n" + msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o estorno de pagamento no pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o estorno de pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxSelecionada.GetMethodName() + "\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ processaTransacaoEstornoPendente ]
		/// <summary>
		/// 
		/// </summary>
		/// <param name="payment"></param>
		/// <param name="ult_GlobalStatus_atualizado"></param>
		/// <param name="msg_erro"></param>
		/// <returns>
		///		Valores de retorno:
		///			true: (1) A transação de estorno pendente foi confirmada e o processamento no banco de dados foi realizado com sucesso; (2) A transação de estorno ainda continua pendente
		///			false: (1) Houve falha na consulta; (2) Houve falha no processamento no banco de dados; (3) Outras falhas
		/// </returns>
		public static bool processaTransacaoEstornoPendente(BraspagPagPayment payment, out bool st_estorno_confirmado, out string ult_GlobalStatus_atualizado, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.processaTransacaoEstornoPendente()";
			Global.Cte.Braspag.Pagador.Transacao trxGetTransactionData = Global.Cte.Braspag.Pagador.Transacao.GetTransactionData;
			Global.Cte.Braspag.Pagador.Transacao trxRefund = Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string xmlReqSoap;
			string xmlRespSoap;
			string msg_erro_last_op;
			string strSubject;
			string strBody;
			string strMsg;
			bool blnBraspagTransactionIdOk = false;
			bool blnAtualizouBraspagTransactionId = false;
			bool blnEnviouOk;
			bool blnRespostaOk = false;
			bool blnSucesso;
			int id_emailsndsvc_mensagem;
			DateTime dtCapturedDate = DateTime.MinValue;
			DateTime dtVoidedDate = DateTime.MinValue;
			BraspagGetTransactionData trx;
			BraspagGetTransactionDataResponse rRESP = null;
			BraspagPag pag;
			BraspagPagPayment paymentAtualizado;
			BraspagPagOpComplementar opCompl;
			BraspagPagOpComplementarXml opComplXmlTx;
			BraspagPagOpComplementarXml opComplXmlRx;
			BraspagUpdatePagPaymentGetTransactionDataResponse rUpdate;
			BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso rUpdatePaymentRefundSucesso;
			BraspagUpdatePagPaymentRefundPendingConfirmado rUpdateRefundPendingConfirmado;
			#endregion

			st_estorno_confirmado = false;
			ult_GlobalStatus_atualizado = "";
			msg_erro = "";
			try
			{
				if (payment.resp_PaymentDataResponse_BraspagTransactionId != null)
				{
					if (payment.resp_PaymentDataResponse_BraspagTransactionId.Trim().Length > 0) blnBraspagTransactionIdOk = true;
				}

				if (!blnBraspagTransactionIdOk)
				{
					if (!verificaPreRequisitoBraspagTransactionId(payment, out blnAtualizouBraspagTransactionId, out msg_erro_aux))
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar obter o valor de 'BraspagTransactionId'\n" + msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
				}

				if (blnAtualizouBraspagTransactionId)
				{
					paymentAtualizado = BraspagDAO.getBraspagPagPaymentById(payment.id, out msg_erro_aux);
				}
				else
				{
					paymentAtualizado = payment;
				}

				if (paymentAtualizado.resp_PaymentDataResponse_BraspagTransactionId.Trim().Length == 0)
				{
					msg_erro = "Não é possível consultar a Braspag porque não foi obtido o TransactionId quando a transação foi realizada inicialmente!!";
					return false;
				}

				pag = BraspagDAO.getBraspagPagById(paymentAtualizado.id_pagto_gw_pag, out msg_erro_aux);
				trx = criaGetTransactionData(pag.req_OrderData_MerchantId, paymentAtualizado.resp_PaymentDataResponse_BraspagTransactionId);
				xmlReqSoap = montaXmlRequisicaoSoapGetTransactionData(trx);

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR ]
				opCompl = new BraspagPagOpComplementar();
				opCompl.id_pagto_gw_pag = paymentAtualizado.id_pagto_gw_pag;
				opCompl.id_pagto_gw_pag_payment = paymentAtualizado.id;
				opCompl.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opCompl.operacao = trxGetTransactionData.GetCodOpLog();
				opCompl.req_RequestId = trx.RequestId;
				opCompl.req_Version = trx.Version;
				opCompl.req_MerchantId = trx.MerchantId;
				opCompl.req_BraspagTransactionId = trx.BraspagTransactionId;
				if (!BraspagDAO.inserePagOpComplementar(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR\n" + msg_erro_aux);
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new BraspagPagOpComplementarXml();
				opComplXmlTx.id_pagto_gw_pag_op_complementar = opCompl.id;
				opComplXmlTx.tipo_transacao = trxGetTransactionData.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (TX)\n" + msg_erro_aux);
				}
				#endregion

				#region [ Envia requisição para a Braspag ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxGetTransactionData, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX) ]
				if (blnEnviouOk)
				{
					opComplXmlRx = new BraspagPagOpComplementarXml();
					opComplXmlRx.id_pagto_gw_pag_op_complementar = opCompl.id;
					opComplXmlRx.tipo_transacao = trxGetTransactionData.GetCodOpLog();
					opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
					opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
					if (!BraspagDAO.inserePagOpComplementarXml(opComplXmlRx, out msg_erro_aux))
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (RX)\n" + msg_erro_aux);
					}
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar enviar transação para a Braspag: " + trxGetTransactionData.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Braspag " + trxGetTransactionData.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opCompl.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opCompl.trx_RX_vazio_status = 0;
					opCompl.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					rRESP = decodificaXmlGetTransactionDataResponse(xmlRespSoap, out msg_erro_aux);
					if (rRESP == null)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Falha ao tentar decodificar a resposta da requisição Braspag " + trxGetTransactionData.GetMethodName() + "\n" + msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}

					#region [ Success ]
					if (rRESP.Success != null)
					{
						if (rRESP.Success.Equals("true")) opCompl.st_sucesso = 1;
					}
					#endregion

					#region [ BraspagTransactionId ]
					if (rRESP.BraspagTransactionId != null) opCompl.resp_BraspagTransactionId = rRESP.BraspagTransactionId;
					#endregion

					#region [ AuthorizationCode ]
					if (rRESP.AuthorizationCode != null) opCompl.resp_AuthorizationCode = rRESP.AuthorizationCode;
					#endregion

					#region [ ProofOfSale ]
					if (rRESP.ProofOfSale != null) opCompl.resp_ProofOfSale = rRESP.ProofOfSale;
					#endregion

					#region [ Status ]
					if (rRESP.Status != null) opCompl.resp_Status = rRESP.Status;
					#endregion

					#region [ Há mensagem de erro na resposta? ]
					if (rRESP.ErrorReportDataCollection.Count > 0)
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Requisição ao método Braspag " + trxGetTransactionData.GetMethodName() + " retornou mensagem de erro: " + rRESP.ErrorReportDataCollection[0].ErrorCode + " - " + rRESP.ErrorReportDataCollection[0].ErrorMessage;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = xmlRespSoap;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						return false;
					}
					#endregion

					if (rRESP.Success != null)
					{
						if ((rRESP.Success.Equals("true")) && (rRESP.ErrorReportDataCollection.Count == 0)) blnRespostaOk = true;
					}
				}
				#endregion

				if (!blnRespostaOk) return false;

				#region [ Atualiza tabela T_PAGTO_GW_PAG_OP_COMPLEMENTAR com dados da resposta ]
				if (!BraspagDAO.updatePagOpComplementarGetTransactionDataResp(opCompl, out msg_erro_aux))
				{
					Global.gravaLogAtividade("Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR + " (id=" + opCompl.id.ToString() + ")\n" + msg_erro_aux);
				}
				#endregion

				#region [ Atualiza tabela T_PAGTO_GW_PAG_PAYMENT com dados da resposta ]
				if (rRESP.CapturedDate != null)
				{
					if (rRESP.CapturedDate.Trim().Length > 0)
					{
						dtCapturedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(rRESP.CapturedDate);
					}
				}

				if (rRESP.VoidedDate != null)
				{
					if (rRESP.VoidedDate.Trim().Length > 0)
					{
						dtVoidedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(rRESP.VoidedDate);
					}
				}

				rUpdate = new BraspagUpdatePagPaymentGetTransactionDataResponse();
				rUpdate.id_pagto_gw_pag_payment = paymentAtualizado.id;
				rUpdate.ult_GlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(rRESP.Status);
				ult_GlobalStatus_atualizado = rUpdate.ult_GlobalStatus;
				rUpdate.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				rUpdate.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
				rUpdate.resp_CapturedDate = dtCapturedDate;
				rUpdate.resp_VoidedDate = dtVoidedDate;

				if (!BraspagDAO.updatePagPaymentGetTransactionDataResp(rUpdate, out msg_erro_aux))
				{
					Global.gravaLogAtividade("Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " (id=" + paymentAtualizado.id.ToString() + ")\n" + msg_erro_aux);
				}
				#endregion

				#region [ Se transação consta como estornada na Braspag, realiza processamento de rotinas do ERP ]
				if (!ult_GlobalStatus_atualizado.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNADA.GetValue()))
				{
					return true;
				}
				else
				{
					st_estorno_confirmado = true;

					#region [ Registra dados referentes à confirmação do estorno ]
					rUpdatePaymentRefundSucesso = new BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso();
					rUpdatePaymentRefundSucesso.id_pagto_gw_pag_payment = payment.id;
					rUpdatePaymentRefundSucesso.ult_id_pagto_gw_pag_payment_op_complementar = opCompl.id;
					rUpdatePaymentRefundSucesso.ult_atualizacao_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
					rUpdatePaymentRefundSucesso.ult_GlobalStatus = Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNADA.GetValue();
					rUpdatePaymentRefundSucesso.refunded_status = 1;
					rUpdatePaymentRefundSucesso.refunded_data = dtVoidedDate;
					rUpdatePaymentRefundSucesso.refunded_data_hora = dtVoidedDate;
					rUpdatePaymentRefundSucesso.refunded_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
					rUpdatePaymentRefundSucesso.resp_VoidedDate = dtVoidedDate;

					if (!BraspagDAO.updatePagPaymentRefundCreditCardRespSucesso(rUpdatePaymentRefundSucesso, out msg_erro_aux))
					{
						msg_erro_last_op = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT para confirmar o estorno de um estorno pendente\n" + msg_erro_last_op;
						svcLog.complemento_1 = Global.serializaObjectToXml(rUpdatePaymentRefundSucesso);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar confirmar o estorno de um estorno pendente [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar confirmar o estorno de um estorno pendente do pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
					}
					#endregion

					#region [ Processamento no banco de dados ]
					blnSucesso = false;
					BD.iniciaTransacao();
					try
					{
						#region [ Processa registro de pagamento no pedido ]
						blnSucesso = BraspagDAO.registraPagamentoNoPedido(trxRefund, payment.id, out msg_erro_aux);
						if (!blnSucesso) msg_erro = msg_erro_aux;
						#endregion

						#region [ Registra no BD que este estorno pendente já foi processado ]
						if (blnSucesso)
						{
							rUpdateRefundPendingConfirmado = new BraspagUpdatePagPaymentRefundPendingConfirmado();
							rUpdateRefundPendingConfirmado.id_pagto_gw_pag_payment = payment.id;
							rUpdateRefundPendingConfirmado.refund_pending_confirmado_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;

							if (!BraspagDAO.updatePagPaymentRefundPendingConfirmado(rUpdateRefundPendingConfirmado, out msg_erro_aux))
							{
								blnSucesso = false;
								msg_erro_last_op = msg_erro_aux;

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com as informações de que o estorno pendente foi confirmado\n" + msg_erro_last_op;
								svcLog.complemento_1 = Global.serializaObjectToXml(rUpdateRefundPendingConfirmado);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar as informações de que o estorno pendente foi confirmado [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar as informações de que o estorno pendente foi confirmado para o pedido " + pag.pedido + " (" + pag.pedido_com_sufixo_nsu + "; t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_last_op;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
							}
						}
						#endregion
					}
					catch (Exception ex)
					{
						blnSucesso = false;
						msg_erro = ex.ToString();
					}
					finally
					{
						if (blnSucesso)
						{
							try
							{
								BD.commitTransacao();
							}
							catch (Exception ex)
							{
								if ((msg_erro ?? "").Length == 0) msg_erro = ex.ToString();
								blnSucesso = false;
							}
						}
						else
						{
							BD.rollbackTransacao();
						}
					}

					if (blnSucesso)
					{
						return true;
					}
					else
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxRefund.GetMethodName() + " pendente que foi confirmada\r\n" + msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o estorno de pagamento no pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o estorno de pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ") devido a uma transação " + trxRefund.GetMethodName() + " pendente que foi confirmada\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
					#endregion
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ enviaRequisicaoComRetry ]
		/// <summary>
		/// Método que executa o enviaRequisicao() dentro de um laço de tentativas até que a execução seja bem sucedida ou a quantidade máxima de tentativas seja atingida.
		/// Importante: este método pode ser utilizado livremente para requisições de consulta, entretanto, para requisições que alteram dados é importante avaliar antes
		/// as possíveis consequências que podem ocorrer no caso da requisição ter sido processada no web service e o erro ter ocorrido em algum estágio posterior durante
		/// o recebimento da resposta. Nesse caso, o uso deste método pode causar múltiplas execuções da requisição.
		/// </summary>
		/// <param name="xmlReqSoap"></param>
		/// <param name="trxParam"></param>
		/// <param name="xmlRespSoap"></param>
		/// <param name="msg_erro"></param>
		/// <returns></returns>
		private static bool enviaRequisicaoComRetry(string xmlReqSoap, Global.Cte.Braspag.Pagador.Transacao trxParam, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const int MAX_TENTATIVAS = 5;
			int qtdeTentativasRealizadas = 0;
			bool blnResposta;
			#endregion

			do
			{
				qtdeTentativasRealizadas++;

				blnResposta = enviaRequisicao(xmlReqSoap, trxParam, out xmlRespSoap, out msg_erro);
				if (blnResposta) break;

				Thread.Sleep(5 * 1000);
			} while (qtdeTentativasRealizadas < MAX_TENTATIVAS);

			return blnResposta;
		}
		#endregion

		#region [ enviaRequisicao ]
		private static bool enviaRequisicao(string xmlReqSoap, Global.Cte.Braspag.Pagador.Transacao trxParam, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Braspag.enviaRequisicao()";
			HttpWebRequest req;
			HttpWebResponse resp;
			#endregion

			xmlRespSoap = "";
			msg_erro = "";
			try
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - XML (TX)\n" + xmlReqSoap);

				req = (HttpWebRequest)WebRequest.Create(trxParam.GetEnderecoWebService());
				// The Timeout applies to the entire request and response, not individually to the GetRequestStream and GetResponse method calls
				req.Timeout = Global.Cte.Braspag.REQUEST_TIMEOUT_EM_MS;
				req.Headers.Add("SOAPAction", trxParam.GetSoapAction());
				req.ContentType = "text/xml;charset=\"utf-8\"";
				req.Method = "POST";
				using (Stream reqStm = req.GetRequestStream())
				{
					using (StreamWriter reqStmW = new StreamWriter(reqStm))
					{
						reqStmW.Write(xmlReqSoap);
					}
				}

				resp = (HttpWebResponse)req.GetResponse();
				using (Stream respStm = resp.GetResponseStream())
				{
					using (StreamReader respStmR = new StreamReader(respStm, Encoding.UTF8))
					{
						xmlRespSoap = respStmR.ReadToEnd();
					}
				}

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - XML (RX)\n" + xmlRespSoap);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - Exception\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ executaProcessamentoAtualizaStatusTransacoesPendentes ]
		/// <summary>
		/// Obtém a relação de transações que ainda constam como autorizadas, ou seja, ainda não tiveram a captura confirmada e obtém o status atual na Braspag para atualizar no banco de dados.
		/// Ressaltando que a transação que não tiver a captura confirmada dentro do período especificado (5 dias corridos), é cancelada automaticamente pela administradora.
		/// </summary>
		/// <returns></returns>
		public static bool executaProcessamentoAtualizaStatusTransacoesPendentes(out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaProcessamentoAtualizaStatusTransacoesPendentes()";
			int qtdeTrxTotal = 0;
			int qtdeTrxSucesso = 0;
			int qtdeTrxFalha = 0;
			string strSql;
			string msg_erro_aux;
			string strMsg;
			string ult_GlobalStatus_original;
			string ult_GlobalStatus_novo;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			List<BraspagPagPayment> listaTrx = new List<BraspagPagPayment>();
			BraspagPagPayment trx;
			BraspagPag pag;
			StringBuilder sbFalha = new StringBuilder("");
			StringBuilder sbSucesso = new StringBuilder("");
			#endregion

			strMsgInformativa = "";
			msg_erro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Obtém a relação de transações de pagamento que ainda não tiveram a captura confirmada ]

				#region [ Monta o SQL da consulta ]
				strSql = "SELECT" +
							" tPAG_PAY.id" +
						" FROM t_PAGTO_GW_PAG_PAYMENT tPAG_PAY" +
							" INNER JOIN t_PAGTO_GW_PAG tPAG ON (tPAG_PAY.id_pagto_gw_pag = tPAG.id)" +
						" WHERE" +
							" (tPAG.data < " + Global.sqlMontaGetdateSomenteData() + ")" +
							" AND (tPAG_PAY.ult_GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA + "')" +
						" ORDER BY" +
							" tPAG.data," +
							" tPAG.data_hora," +
							" tPAG_PAY.id";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					trx = BraspagDAO.getBraspagPagPaymentById(BD.readToInt(dtbResultado.Rows[i]["id"]), out msg_erro_aux);
					listaTrx.Add(trx);
				}
				#endregion

				for (int i = 0; i < listaTrx.Count; i++)
				{
					qtdeTrxTotal++;
					pag = BraspagDAO.getBraspagPagById(listaTrx[i].id_pagto_gw_pag, out msg_erro_aux);

					if (!processaConsultaGetTransactionData(listaTrx[i], out ult_GlobalStatus_original, out ult_GlobalStatus_novo, out msg_erro_aux))
					{
						qtdeTrxFalha++;
						strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + listaTrx[i].id.ToString() + "): " + msg_erro_aux;
						if (msg_erro_aux.Length > 0) sbFalha.AppendLine(strMsg);
					}
					else
					{
						qtdeTrxSucesso++;
						strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + listaTrx[i].id.ToString() + "): " + ult_GlobalStatus_original + (!ult_GlobalStatus_original.Equals(ult_GlobalStatus_novo) ? "=>" + ult_GlobalStatus_novo : "");
						sbSucesso.AppendLine(strMsg);
					}
				}

				strMsgInformativa = qtdeTrxTotal.ToString() + " consulta(s) realizada(s): " + qtdeTrxSucesso.ToString() + " com sucesso e " + qtdeTrxFalha.ToString() + " com falha" +
									"\n" +
									"Sucesso:\n" + (sbSucesso.ToString().Length > 0 ? sbSucesso.ToString() : "(nenhum)") +
									"\n" +
									"Falha:\n" + (sbFalha.ToString().Length > 0 ? sbFalha.ToString() : "(nenhum)");
				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
			finally
			{
				if (strMsgInformativa.Length > 0)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLogInfo = new FinSvcLog();
					svcLogInfo.operacao = NOME_DESTA_ROTINA;
					svcLogInfo.descricao = strMsgInformativa;
					GeralDAO.gravaFinSvcLog(svcLogInfo, out msg_erro_aux);
					#endregion
				}
			}
		}
		#endregion

		#region [ executaProcessamentoEnviarEmailAlertaTransacoesPendentesProxCancelAuto ]
		public static bool executaProcessamentoEnviarEmailAlertaTransacoesPendentesProxCancelAuto(out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaProcessamentoEnviarEmailAlertaTransacoesPendentesProxCancelAuto()";
			int qtdePrazoHoje = 0;
			int qtdePrazoAmanha = 0;
			int qtdePrazoOutros = 0;
			int id_emailsndsvc_mensagem;
			int qtdeDiasPrazoAlerta;
			string strSql;
			string strMsg;
			string msg_erro_aux;
			string strSubject;
			string strBody;
			StringBuilder sbPrazoHoje = new StringBuilder("");
			StringBuilder sbPrazoAmanha = new StringBuilder("");
			StringBuilder sbPrazoOutros = new StringBuilder("");
			StringBuilder sbEmail = new StringBuilder("");
			DateTime dtFinalCaptura;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			var culture = new System.Globalization.CultureInfo("pt-BR");
			#endregion

			strMsgInformativa = "";
			msg_erro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Prazo para alerta ]
				if (DateTime.Now.Date.DayOfWeek == DayOfWeek.Friday)
				{
					// Se hoje for 6ªf, inclui na mensagem as transações que expiram hoje, amanhã e domingo
					qtdeDiasPrazoAlerta = 2;
				}
				else
				{
					// Caso contrário, inlcui na mensagem as transações que expiram hoje e amanhã
					qtdeDiasPrazoAlerta = 1;
				}
				#endregion

				#region [ Monta SQL ]
				// Obtém a relação de transações de pagamento que ainda não tiveram a captura confirmada, ou seja, ainda estão c/ status 'Autorizada' e que
				// estão próximas do prazo limite de serem canceladas automaticamente
				strSql = "SELECT " +
							"*" +
						" FROM " +
							"(" +
								"SELECT" +
									" tPAG.id AS id_pagto_gw_pag," +
									" tPAG_PAY.id AS id_pagto_gw_pag_payment," +
									" Convert(datetime, Convert(varchar(10), Coalesce(tPAG_PAY.resp_AuthorizedDate, tPAG.data), 121), 121) AS AuthorizedDate," +
									" (Convert(datetime, Convert(varchar(10), Coalesce(tPAG_PAY.resp_AuthorizedDate, tPAG.data), 121), 121) + " + Global.Cte.Braspag.Pagador.PRAZO_CAPTURA_EM_DIAS_CORRIDOS.ToString() + ") AS DataFinalCaptura," +
									" tPAG.data," +
									" tPAG.data_hora," +
									" tPAG.pedido," +
									" tPAG.pedido_com_sufixo_nsu," +
									" tPAG_PAY.bandeira," +
									" tPAG_PAY.valor_transacao," +
									" tPAG_PAY.req_PaymentDataRequest_NumberOfPayments AS numero_parcelas," +
									" Coalesce(tCLI.nome_iniciais_em_maiusculas, '') AS nome_cliente" +
								" FROM t_PAGTO_GW_PAG_PAYMENT tPAG_PAY" +
									" INNER JOIN t_PAGTO_GW_PAG tPAG ON (tPAG_PAY.id_pagto_gw_pag = tPAG.id)" +
									" LEFT JOIN t_CLIENTE tCLI ON (tCLI.id = tPAG.id_cliente)" +
								" WHERE" +
									" (tPAG_PAY.ult_GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA + "')" +
									" AND (tPAG.data < " + Global.sqlMontaGetdateSomenteData() + ")" +
							") t" +
						" WHERE" +
							" (DataFinalCaptura <= (" + Global.sqlMontaGetdateSomenteData() + " + " + qtdeDiasPrazoAlerta.ToString() + "))" +
						" ORDER BY" +
							" AuthorizedDate," +
							" id_pagto_gw_pag," +
							" id_pagto_gw_pag_payment";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Processa o resultado ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					dtFinalCaptura = BD.readToDateTime(row["DataFinalCaptura"]);

					strMsg = "Pedido " + BD.readToString(row["pedido"]);
					if (!BD.readToString(row["pedido"]).Equals(BD.readToString(row["pedido_com_sufixo_nsu"]))) strMsg += " (" + BD.readToString(row["pedido_com_sufixo_nsu"]) + ")";
					strMsg += ", " + Texto.iniciaisEmMaiusculas(BD.readToString(row["bandeira"])) + ", " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(BD.readToDecimal(row["valor_transacao"]));
					strMsg += " em " + BD.readToString(row["numero_parcelas"]) + "x" + " (" + BD.readToString(row["nome_cliente"]) + ")";

					if (dtFinalCaptura == DateTime.Today)
					{
						qtdePrazoHoje++;
						sbPrazoHoje.AppendLine(strMsg);
					}
					else if (dtFinalCaptura == DateTime.Today.AddDays(1))
					{
						qtdePrazoAmanha++;
						sbPrazoAmanha.AppendLine(strMsg);
					}
					else
					{
						qtdePrazoOutros++;
						strMsg = "Prazo final em " + Global.formataDataDdMmYyyyComSeparador(dtFinalCaptura) + " (" + culture.DateTimeFormat.GetDayName(dtFinalCaptura.Date.DayOfWeek) + "): " + strMsg;
						sbPrazoOutros.AppendLine(strMsg);
					}
				}
				#endregion

				#region [ Prepara mensagem ]
				sbEmail.AppendLine("Transações pendentes com prazo final de cancelamento automático para HOJE (" + Global.formataDataDdMmYyyyComSeparador(DateTime.Today) + ")");
				if (sbPrazoHoje.Length > 0)
				{
					sbEmail.AppendLine(sbPrazoHoje.ToString());
				}
				else
				{
					sbEmail.AppendLine("(nenhuma transação)");
				}

				if (sbEmail.Length > 0)
				{
					sbEmail.AppendLine("");
					sbEmail.AppendLine("");
				}

				sbEmail.AppendLine("Transações pendentes com prazo final de cancelamento automático para AMANHÃ (" + Global.formataDataDdMmYyyyComSeparador(DateTime.Today.AddDays(1)) + ")");
				if (sbPrazoAmanha.Length > 0)
				{
					sbEmail.AppendLine(sbPrazoAmanha.ToString());
				}
				else
				{
					sbEmail.AppendLine("(nenhuma transação)");
				}

				if (sbPrazoOutros.Length > 0)
				{
					if (sbEmail.Length > 0)
					{
						sbEmail.AppendLine("");
						sbEmail.AppendLine("");
					}
					sbEmail.AppendLine("Transações pendentes próximas do prazo final de cancelamento automático");
					sbEmail.AppendLine(sbPrazoOutros.ToString());
				}
				#endregion

				#region [ Grava email para envio ]
				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: transações pendentes próximas do prazo de cancelamento automático";
				if ((qtdePrazoHoje + qtdePrazoAmanha + qtdePrazoOutros) == 0)
				{
					strMsg = " (nenhuma transação)";
					strSubject += strMsg;
					strMsgInformativa += strMsg;
				}
				else
				{
					strMsg = " (hoje: " + qtdePrazoHoje.ToString() + ", amanhã: " + qtdePrazoAmanha.ToString() + ", outros: " + qtdePrazoOutros.ToString() + ")";
					strSubject += strMsg;
					strMsgInformativa += strMsg;
				}
				strSubject += " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\nTransações pendentes próximas do prazo final para captura antes de serem canceladas automaticamente.\r\n\r\n" + sbEmail.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta sobre transações pendentes próximas do prazo de cancelamento automático na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}
				#endregion

				if (sbPrazoHoje.Length > 0) strMsgInformativa += "\r\nHoje:\r\n" + sbPrazoHoje.ToString();
				if (sbPrazoAmanha.Length > 0) strMsgInformativa += "\r\nAmanhã:\r\n" + sbPrazoAmanha.ToString();
				if (sbPrazoOutros.Length > 0) strMsgInformativa += "\r\nOutros:\r\n" + sbPrazoOutros.ToString();

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
			finally
			{
				if (strMsgInformativa.Length > 0)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLogInfo = new FinSvcLog();
					svcLogInfo.operacao = NOME_DESTA_ROTINA;
					svcLogInfo.descricao = strMsgInformativa;
					GeralDAO.gravaFinSvcLog(svcLogInfo, out msg_erro_aux);
					#endregion
				}
			}
		}
		#endregion

		#region [ executaCapturaTransacoesPendentesPrazoFinalCancelAuto ]
		public static bool executaCapturaTransacoesPendentesPrazoFinalCancelAuto(out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaCapturaTransacoesPendentesPrazoFinalCancelAuto()";
			int id_pagto_gw_pag_payment;
			int id_emailsndsvc_mensagem;
			int qtdeTrxTotal = 0;
			int qtdeTrxSucesso = 0;
			int qtdeTrxFalha = 0;
			string strSql;
			string strMsg;
			string msg_erro_aux;
			string msg_erro_temp;
			string strSubject;
			string strBody;
			string strPedidoInfo;
			string ult_GlobalStatus_original_aux;
			string ult_GlobalStatus_novo_aux;
			string ult_GlobalStatus_atualizado;
			StringBuilder sbEmail = new StringBuilder("");
			StringBuilder sbMsgSucesso = new StringBuilder("");
			StringBuilder sbMsgFalha = new StringBuilder("");
			StringBuilder sbPedidoSucesso = new StringBuilder("");
			StringBuilder sbPedidoFalha = new StringBuilder("");
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			var culture = new System.Globalization.CultureInfo("pt-BR");
			BraspagPagPayment payment;
			BraspagPagPayment paymentAtualizado;
			#endregion

			strMsgInformativa = "";
			msg_erro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Monta SQL ]
				// Obtém a relação de transações de pagamento que ainda não tiveram a captura confirmada, ou seja, ainda estão c/ status 'Autorizada' e que
				// o prazo final para captura seja hoje
				strSql = "SELECT " +
							"*" +
						" FROM " +
							"(" +
								"SELECT" +
									" tPAG.id AS id_pagto_gw_pag," +
									" tPAG_PAY.id AS id_pagto_gw_pag_payment," +
									" Convert(datetime, Convert(varchar(10), Coalesce(tPAG_PAY.resp_AuthorizedDate, tPAG.data), 121), 121) AS AuthorizedDate," +
									" (Convert(datetime, Convert(varchar(10), Coalesce(tPAG_PAY.resp_AuthorizedDate, tPAG.data), 121), 121) + " + Global.Cte.Braspag.Pagador.PRAZO_CAPTURA_EM_DIAS_CORRIDOS.ToString() + ") AS DataFinalCaptura," +
									" tPAG.data," +
									" tPAG.data_hora," +
									" tPAG.pedido," +
									" tPAG.pedido_com_sufixo_nsu," +
									" tPAG_PAY.bandeira," +
									" tPAG_PAY.valor_transacao," +
									" tPAG_PAY.req_PaymentDataRequest_NumberOfPayments AS numero_parcelas," +
									" Coalesce(tCLI.nome_iniciais_em_maiusculas, '') AS nome_cliente" +
								" FROM t_PAGTO_GW_PAG_PAYMENT tPAG_PAY" +
									" INNER JOIN t_PAGTO_GW_PAG tPAG ON (tPAG_PAY.id_pagto_gw_pag = tPAG.id)" +
									" LEFT JOIN t_CLIENTE tCLI ON (tCLI.id = tPAG.id_cliente)" +
								" WHERE" +
									" (tPAG_PAY.ult_GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA + "')" +
									" AND (tPAG.data < " + Global.sqlMontaGetdateSomenteData() + ")" +
							") t" +
						" WHERE" +
							" (DataFinalCaptura <= " + Global.sqlMontaGetdateSomenteData() + ")" +
						" ORDER BY" +
							" id_pagto_gw_pag," +
							" id_pagto_gw_pag_payment";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Processa o resultado ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					qtdeTrxTotal++;
					row = dtbResultado.Rows[i];
					id_pagto_gw_pag_payment = BD.readToInt(row["id_pagto_gw_pag_payment"]);

					strPedidoInfo = BD.readToString(row["pedido"]) + " (t_PAGTO_GW_PAG_PAYMENT.id=" + id_pagto_gw_pag_payment.ToString() + ")";

					strMsg = "Pedido " + BD.readToString(row["pedido"]);
					if (!BD.readToString(row["pedido"]).Equals(BD.readToString(row["pedido_com_sufixo_nsu"]))) strMsg += " (" + BD.readToString(row["pedido_com_sufixo_nsu"]) + ")";
					strMsg += ", " + Texto.iniciaisEmMaiusculas(BD.readToString(row["bandeira"])) + ", " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(BD.readToDecimal(row["valor_transacao"]));
					strMsg += " em " + BD.readToString(row["numero_parcelas"]) + "x" + " (" + BD.readToString(row["nome_cliente"]) + "): ";

					payment = BraspagDAO.getBraspagPagPaymentById(id_pagto_gw_pag_payment, out msg_erro_aux);
					if (payment != null)
					{
						if (processaRequisicaoCaptureCreditCardTransaction(payment, out msg_erro_aux))
						{
							qtdeTrxSucesso++;
							strMsg += "capturada com sucesso";
							sbMsgSucesso.AppendLine(strMsg);
							if (sbPedidoSucesso.Length > 0) sbPedidoSucesso.Append(", ");
							sbPedidoSucesso.Append(strPedidoInfo);
						}
						else
						{
							msg_erro_temp = msg_erro_aux;

							ult_GlobalStatus_atualizado = "";

							// Verifica se a transação chegou a ser capturada e o status atualizado no banco de dados
							paymentAtualizado = BraspagDAO.getBraspagPagPaymentById(id_pagto_gw_pag_payment, out msg_erro_aux);
							if (paymentAtualizado != null)
							{
								ult_GlobalStatus_atualizado = paymentAtualizado.ult_GlobalStatus;
							}

							if (!ult_GlobalStatus_atualizado.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue()))
							{
								// Realiza uma consulta para verificar se a transação chegou a ser capturada na Braspag
								if (processaConsultaGetTransactionData(payment, out ult_GlobalStatus_original_aux, out ult_GlobalStatus_novo_aux, out msg_erro_aux))
								{
									ult_GlobalStatus_atualizado = ult_GlobalStatus_novo_aux;
								}
								else
								{
									strMsg = "Falha ao tentar consultar o status atualizado da transação na Braspag (t_PAGTO_GW_PAG_PAYMENT.id=" + id_pagto_gw_pag_payment.ToString() + ")!!" +
										(((msg_erro_aux ?? "").Length > 0) ? "\r\n" : "") +
										(msg_erro_aux ?? "");
									Global.gravaLogAtividade(strMsg);
								}
							}

							if (ult_GlobalStatus_atualizado.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue()))
							{
								qtdeTrxSucesso++;
								strMsg += "a transação foi capturada, mas ocorreu um erro no processamento complementar";
								if ((msg_erro_temp ?? "").Length > 0) strMsg += " (" + msg_erro_temp + ")";
								sbMsgSucesso.AppendLine(strMsg);
								if (sbPedidoSucesso.Length > 0) sbPedidoSucesso.Append(", ");
								sbPedidoSucesso.Append(strPedidoInfo);
							}
							else
							{
								qtdeTrxFalha++;
								strMsg += "falha na captura";
								if ((msg_erro_temp ?? "").Length > 0) strMsg += " (" + msg_erro_temp + ")";
								sbMsgFalha.AppendLine(strMsg);
								if (sbPedidoFalha.Length > 0) sbPedidoFalha.Append(", ");
								sbPedidoFalha.Append(strPedidoInfo);
							}
						}
					}
					else
					{
						qtdeTrxFalha++;
						strMsg += "falha ao tentar obter os dados da transação";
						sbMsgFalha.AppendLine(strMsg);
						if (sbPedidoFalha.Length > 0) sbPedidoFalha.Append(", ");
						sbPedidoFalha.Append(strPedidoInfo);
					}
				}
				#endregion

				#region [ Envia email? ]
				// Envia email informativo somente se houve alguma transação c/ prazo final hoje
				if (qtdeTrxTotal == 0)
				{
					strMsgInformativa += "(nenhuma transação)";
				}
				else
				{
					#region [ Prepara mensagem ]
					sbEmail.AppendLine("Transações capturadas com sucesso: " + qtdeTrxSucesso.ToString());
					if (sbMsgSucesso.Length > 0)
					{
						sbEmail.AppendLine(sbMsgSucesso.ToString());
					}
					else
					{
						sbEmail.AppendLine("(nenhuma transação)");
					}

					if (sbEmail.Length > 0)
					{
						sbEmail.AppendLine("");
						sbEmail.AppendLine("");
					}

					sbEmail.AppendLine("Transações com falha na captura: " + qtdeTrxFalha.ToString());
					if (sbMsgFalha.Length > 0)
					{
						sbEmail.AppendLine(sbMsgFalha.ToString());
					}
					else
					{
						sbEmail.AppendLine("(nenhuma transação)");
					}
					#endregion

					#region [ Grava email para envio ]
					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: transações capturadas automaticamente devido ao prazo final de cancelamento automático";
					if ((qtdeTrxSucesso + qtdeTrxFalha) == 0)
					{
						strMsg = " (nenhuma transação)";
						strSubject += strMsg;
						strMsgInformativa += strMsg;
					}
					else
					{
						strMsg = " (sucesso: " + qtdeTrxSucesso.ToString() + ", falha: " + qtdeTrxFalha.ToString() + ")";
						strSubject += strMsg;
						strMsgInformativa += strMsg;
					}
					strSubject += " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nTransações capturadas automaticamente pelo sistema devido ao prazo final de cancelamento automático.\r\n\r\n" + sbEmail.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir na fila de mensagens o email informativo sobre transações capturadas automaticamente devido ao prazo final de cancelamento automático!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}
					#endregion
				}
				#endregion

				if (sbPedidoSucesso.Length > 0) strMsgInformativa += "\r\nSucesso na captura: " + sbPedidoSucesso.ToString();
				if (sbPedidoFalha.Length > 0) strMsgInformativa += "\r\nFalha na captura: " + sbPedidoFalha.ToString();

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
			finally
			{
				if (strMsgInformativa.Length > 0)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLogInfo = new FinSvcLog();
					svcLogInfo.operacao = NOME_DESTA_ROTINA;
					svcLogInfo.descricao = strMsgInformativa;
					GeralDAO.gravaFinSvcLog(svcLogInfo, out msg_erro_aux);
					#endregion
				}
			}
		}
		#endregion

		#region [ consultaDadosConsolidadosBoletoParaWebhook ]
		public static BraspagWebhookDadosConsolidadosBoleto consultaDadosConsolidadosBoletoParaWebhook(string merchantId, string orderId, out string msg_erro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "Braspag.consultaDadosConsolidadosBoletoParaWebhook()";
			bool blnSucesso = false;
			string msg_erro_aux;
			string strReceivedDate;
			string strCapturedDate;
			string strBoletoExpirationDate;
			DateTime? dtReceivedDate;
			DateTime? dtCapturedDate;
			DateTime? dtBoletoExpirationDate;
			BraspagWebhookDadosConsolidadosBoleto rRESP = null;
			BraspagGetOrderIdDataResponse rGetOrderIdData;
			BraspagGetTransactionDataResponse rGetTransactionData;
			BraspagGetBoletoDataResponse rGetBoletoData;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consulta ao GetOrderIdData para obter a relação de BraspagTransactionId do pedido ]
				rGetOrderIdData = consultaGetOrderIdData(merchantId, orderId, out msg_erro_aux);
				if (rGetOrderIdData == null)
				{
					msg_erro = "Falha na consulta ao método GetOrderIdData (MerchantId = " + merchantId + ", OrderId = " + orderId + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
					return null;
				}
				#endregion

				#region [ Analisa cada BraspagTransactionId vinculado ao pedido ]
				foreach (BraspagOrderIdTransactionResponse orderIdTransactionResponse in rGetOrderIdData.OrderIdDataCollection)
				{
					foreach (string braspagTransactionId in orderIdTransactionResponse.BraspagTransactionId)
					{
						#region [ Consulta GetTransactionData p/ obter 'Status', 'CapturedDate' e 'ReceivedDate' ]
						rGetTransactionData = consultaGetTransactionData(merchantId, braspagTransactionId, out msg_erro_aux);
						if (rGetTransactionData == null)
						{
							msg_erro = "Falha na consulta ao método GetTransactionData (MerchantId=" + merchantId + ", BraspagTransactionId=" + braspagTransactionId + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
							Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
							return null;
						}
						#endregion

						#region [ Consulta GetBoletoData p/ obter 'CustomerName', 'BoletoExpirationDate', 'Amount' e 'PaidAmount' ]
						if (
							(rGetTransactionData.PaymentMethod.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue()) || rGetTransactionData.PaymentMethod.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue()))
							&&
							((rGetTransactionData.CapturedDate ?? "").Length > 0)
							)
						{
							rGetBoletoData = consultaGetBoletoData(merchantId, braspagTransactionId, out msg_erro_aux);
							if (rGetBoletoData == null)
							{
								msg_erro = "Falha na consulta ao método GetBoletoData (MerchantId=" + merchantId + ", BraspagTransactionId=" + braspagTransactionId + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
								Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
								return null;
							}

							if (rGetBoletoData.Success.ToLower().Equals("true"))
							{
								dtReceivedDate = null;
								dtCapturedDate = null;
								dtBoletoExpirationDate = null;

								#region [ ReceivedDate ]
								strReceivedDate = (rGetTransactionData.ReceivedDate ?? "").Trim();
								if (strReceivedDate.Length > 0) dtReceivedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(strReceivedDate);
								#endregion

								#region [ CapturedDate ]
								strCapturedDate = (rGetTransactionData.CapturedDate ?? "").Trim();
								if (strCapturedDate.Length > 0) dtCapturedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(strCapturedDate);
								#endregion

								#region [ BoletoExpirationDate ]
								strBoletoExpirationDate = (rGetBoletoData.BoletoExpirationDate ?? "").Trim();
								if (strBoletoExpirationDate.Length > 0) dtBoletoExpirationDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(strBoletoExpirationDate);
								#endregion

								if (rRESP != null)
								{
									if ((rRESP.CapturedDate != null) && (dtCapturedDate != null))
									{
										// Já foi analisada uma transação da coleção e ela possui data de captura mais recente que a transação em análise atualmente
										if (rRESP.CapturedDate > dtCapturedDate) continue;
									}
								}

								rRESP = new BraspagWebhookDadosConsolidadosBoleto();
								rRESP.MerchantId = merchantId;
								rRESP.OrderId = orderId;
								rRESP.BraspagTransactionId = rGetTransactionData.BraspagTransactionId;
								rRESP.BraspagOrderId = orderIdTransactionResponse.BraspagOrderId;
								rRESP.PaymentMethod = rGetTransactionData.PaymentMethod;
								rRESP.GlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(rGetTransactionData.Status);
								rRESP.ReceivedDate = dtReceivedDate;
								rRESP.CapturedDate = dtCapturedDate;
								rRESP.CustomerName = rGetBoletoData.CustomerName;
								rRESP.BoletoExpirationDate = dtBoletoExpirationDate;
								rRESP.Amount = rGetBoletoData.Amount;
								rRESP.ValorAmount = ((decimal)Global.converteInteiro(rGetBoletoData.Amount)) / 100m;
								rRESP.PaidAmount = rGetBoletoData.PaidAmount;
								rRESP.ValorPaidAmount = ((decimal)Global.converteInteiro(rGetBoletoData.PaidAmount)) / 100m;

								blnSucesso = true;
							}
						}
						#endregion
					}
				}
				#endregion

				if (!blnSucesso) return null;

				return rRESP;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ consultaDadosConsolidadosBoletoParaWebhookV2 ]
		public static BraspagWebhookV2DadosConsolidadosBoleto consultaDadosConsolidadosBoletoParaWebhookV2(string merchantId, string braspagTransactionId, out string msg_erro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "Braspag.consultaDadosConsolidadosBoletoParaWebhookV2()";
			bool blnSucesso = false;
			string msg_erro_aux;
			string strReceivedDate;
			string strCapturedDate;
			string strBoletoExpirationDate;
			DateTime? dtReceivedDate;
			DateTime? dtCapturedDate;
			DateTime? dtBoletoExpirationDate;
			BraspagWebhookV2DadosConsolidadosBoleto rRESP = null;
			BraspagGetTransactionDataResponse rGetTransactionData;
			BraspagGetBoletoDataResponse rGetBoletoData;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consulta GetTransactionData p/ obter 'Status', 'CapturedDate' e 'ReceivedDate' ]
				rGetTransactionData = consultaGetTransactionData(merchantId, braspagTransactionId, out msg_erro_aux);
				if (rGetTransactionData == null)
				{
					msg_erro = "Falha na consulta ao método GetTransactionData (MerchantId=" + merchantId + ", BraspagTransactionId=" + braspagTransactionId + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
					return null;
				}
				#endregion

				#region [ Consulta GetBoletoData p/ obter 'CustomerName', 'BoletoExpirationDate', 'Amount' e 'PaidAmount' ]
				if (
					(rGetTransactionData.PaymentMethod.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue()) || rGetTransactionData.PaymentMethod.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue()))
					&&
					((rGetTransactionData.CapturedDate ?? "").Length > 0)
					)
				{
					rGetBoletoData = consultaGetBoletoData(merchantId, braspagTransactionId, out msg_erro_aux);
					if (rGetBoletoData == null)
					{
						msg_erro = "Falha na consulta ao método GetBoletoData (MerchantId=" + merchantId + ", BraspagTransactionId=" + braspagTransactionId + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
						return null;
					}

					if (rGetBoletoData.Success.ToLower().Equals("true"))
					{
						dtReceivedDate = null;
						dtCapturedDate = null;
						dtBoletoExpirationDate = null;

						#region [ ReceivedDate ]
						strReceivedDate = (rGetTransactionData.ReceivedDate ?? "").Trim();
						if (strReceivedDate.Length > 0) dtReceivedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(strReceivedDate);
						#endregion

						#region [ CapturedDate ]
						strCapturedDate = (rGetTransactionData.CapturedDate ?? "").Trim();
						if (strCapturedDate.Length > 0) dtCapturedDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(strCapturedDate);
						#endregion

						#region [ BoletoExpirationDate ]
						strBoletoExpirationDate = (rGetBoletoData.BoletoExpirationDate ?? "").Trim();
						if (strBoletoExpirationDate.Length > 0) dtBoletoExpirationDate = Global.converteMmDdYyyyHhMmSsAmPmParaDateTime(strBoletoExpirationDate);
						#endregion

						rRESP = new BraspagWebhookV2DadosConsolidadosBoleto();
						rRESP.MerchantId = merchantId;
						rRESP.OrderId = rGetTransactionData.OrderId;
						rRESP.BraspagTransactionId = rGetTransactionData.BraspagTransactionId;
						rRESP.BraspagOrderId = "";
						rRESP.PaymentMethod = rGetTransactionData.PaymentMethod;
						rRESP.GlobalStatus = decodifica_GetTransactionDataResponseStatus_para_GlobalStatus(rGetTransactionData.Status);
						rRESP.ReceivedDate = dtReceivedDate;
						rRESP.CapturedDate = dtCapturedDate;
						rRESP.CustomerName = rGetBoletoData.CustomerName;
						rRESP.BoletoExpirationDate = dtBoletoExpirationDate;
						rRESP.Amount = rGetBoletoData.Amount;
						rRESP.ValorAmount = ((decimal)Global.converteInteiro(rGetBoletoData.Amount)) / 100m;
						rRESP.PaidAmount = rGetBoletoData.PaidAmount;
						rRESP.ValorPaidAmount = ((decimal)Global.converteInteiro(rGetBoletoData.PaidAmount)) / 100m;

						blnSucesso = true;
					}
				}
				#endregion

				if (!blnSucesso) return null;

				return rRESP;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ executaProcessamentoWebhook ]
		public static bool executaProcessamentoWebhook(out bool blnEmailAlertaEnviado, out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "Braspag.executaProcessamentoWebhook()";
			int id_emailsndsvc_mensagem;
			decimal percDif;
			bool blnWebhookQueryComplSucesso;
			bool blnWebhookQueryComplFalhaDefinitiva;
			bool blnWebhookQueryComplFalhaTemporaria;
			bool blnRegistrouPagtoPedido;
			bool blnRegistrouLancamento;
			bool blnPagtoRegistradoManualmente;
			bool blnSucesso;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string msg_erro_send_email;
			string msg_erro_last_op;
			string strSql;
			string strMsg;
			string strSubject;
			string strBody;
			string strMerchantId;
			string strNumPedidoERP;
			string strNumPedidoERPAux;
			string s_log;
			String strLinhaSeparadora = new String('=', 80);
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			List<BraspagWebhook> listaWebhook = new List<BraspagWebhook>();
			BraspagWebhook pedidoWebhook;
			Global.Parametros.Braspag.WebhookBraspagMerchantId webhookBraspagMerchantId;
			Global.Parametros.Braspag.WebhookBraspagPlanoContasBoletoEC webhookBraspagPlanoContasBoletoEC;
			BraspagWebhookDadosConsolidadosBoleto rBoleto;
			BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva updateWebhookQueryComplFalhaDefinitiva = null;
			BraspagUpdateWebhookQueryDadosComplementaresFalhaTemporaria updateWebhookQueryComplFalhaTemporaria = null;
			BraspagUpdateWebhookQueryDadosComplementaresQtdeTentativas updateWebhookQueryComplQtdeTentativas = null;
			BraspagUpdateWebhookQueryDadosComplementaresSucesso updateWebhookQueryComplSucesso = null;
			BraspagInsertWebhookQueryDadosComplementares insertWebhookQueryCompl;
			BraspagWebhookComplementar braspagWebhookComplementarAtualizado;
			BraspagWebhookComplementar braspagWebhookComplementar;
			List<StringBuilder> vDadosEmail = new List<StringBuilder>();
			StringBuilder sbDadosEmail;
			StringBuilder sbBody;
			List<int> vBraspagWebhookIdEmailEnviadoStatusUpdate = new List<int>();
			LancamentoFluxoCaixaInsertDevidoBoletoEC lancamento;
			Pedido pedido;
			Cliente cliente;
			List<PedidoPagamento> listaPagto;
			PedidoPagamento pagtoManual;
			#endregion

			blnEmailAlertaEnviado = false;
			strMsgInformativa = "";
			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Monta SQL ]
				// A consulta seleciona os dados retornados pelo Webhook da Braspag referentes ao boleto Bradesco SPS e que ainda não foram tratados
				strSql = "SELECT " +
							"Id, " +
							"Empresa, " +
							"NumPedido, " +
							"Status, " +
							"CODPAGAMENTO, " +
							"BraspagDadosComplementaresQueryStatus, " +
							"BraspagDadosComplementaresQueryTentativas, " +
							"ProcessamentoErpStatus" +
						" FROM t_BRASPAG_WEBHOOK" +
						" WHERE" +
							" (EmailEnviadoStatus = 0)" +
							" AND (ProcessamentoErpStatus = 0)" +
							" AND (CODPAGAMENTO IN ('" + Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue() + "','" + Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue() + "'))" +
						" ORDER BY" +
							" DataHoraCadastro";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Carrega a lista de pedidos informados pela Braspag ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					pedidoWebhook = new BraspagWebhook();
					pedidoWebhook.Id = BD.readToInt(row["Id"]);
					pedidoWebhook.Empresa = BD.readToString(row["Empresa"]);
					pedidoWebhook.NumPedido = BD.readToString(row["NumPedido"]);
					pedidoWebhook.Status = BD.readToString(row["Status"]);
					pedidoWebhook.CODPAGAMENTO = BD.readToString(row["CODPAGAMENTO"]);
					pedidoWebhook.BraspagDadosComplementaresQueryStatus = BD.readToByte(row["BraspagDadosComplementaresQueryStatus"]);
					pedidoWebhook.BraspagDadosComplementaresQueryTentativas = BD.readToInt(row["BraspagDadosComplementaresQueryTentativas"]);
					pedidoWebhook.ProcessamentoErpStatus = BD.readToInt(row["ProcessamentoErpStatus"]);
					listaWebhook.Add(pedidoWebhook);
				}
				#endregion

				#region [ Há dados? ]
				if (listaWebhook.Count == 0)
				{
					strMsgInformativa = "Não há dados para processar";
					return true;
				}
				#endregion

				#region [ Obtém dados complementares consultando a Braspag ]
				// Importante: os pagamentos que foram realizados diretamente na plataforma de e-commerce precisam ter os dados
				// =========== recuperados através de consultas à Braspag, pois não há nenhuma informação no ERP.
				// O primeiro passo é obter o BraspagTransactionId através do número do pedido e isso é feito através da
				// consulta GetOrderIdData
				foreach (var pedidoWH in listaWebhook)
				{
					#region [ Contador de tentativas ]
					pedidoWH.BraspagDadosComplementaresQueryTentativas++;
					#endregion

					#region [ Inicialização de variáveis a cada iteração ]
					strMerchantId = "";
					webhookBraspagPlanoContasBoletoEC = null;
					blnWebhookQueryComplSucesso = false;
					blnWebhookQueryComplFalhaDefinitiva = false;
					blnWebhookQueryComplFalhaTemporaria = false;
					strNumPedidoERP = "";
					pedido = null;
					cliente = null;
					pagtoManual = null;
					blnRegistrouPagtoPedido = false;
					blnRegistrouLancamento = false;
					blnPagtoRegistradoManualmente = false;
					#endregion

					#region [ Consulta a Braspag para obter dados complementares ]
					// Try-Finally: tratamento para falhas, principalmente se excedeu quantidade máxima de tentativas
					try
					{
						#region [ Obtém o MerchantId p/ o pedido em questão ]
						if ((pedidoWH.Empresa ?? "").Length > 0)
						{
							try
							{
								webhookBraspagMerchantId = Global.Parametros.Braspag.webhookBraspagMerchantIdList.Single(p => p.Empresa.ToUpper().Equals(pedidoWH.Empresa.ToUpper()));
								if (webhookBraspagMerchantId != null)
								{
									strMerchantId = webhookBraspagMerchantId.MerchantId;
								}
							}
							catch (Exception ex)
							{
								strMerchantId = "";
								strMsg = NOME_DESTA_ROTINA + ": exception ao tentar obter o MerchantId para a empresa '" + pedidoWH.Empresa + "'!!" +
										"\r\n" + ex.ToString();
								Global.gravaLogAtividade(strMsg);
							}
						}

						if ((strMerchantId ?? "").Length == 0)
						{
							blnWebhookQueryComplFalhaDefinitiva = true;
							updateWebhookQueryComplFalhaDefinitiva = new BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva();
							updateWebhookQueryComplFalhaDefinitiva.id_braspag_webhook = pedidoWH.Id;
							updateWebhookQueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryTentativas = pedidoWH.BraspagDadosComplementaresQueryTentativas;
							updateWebhookQueryComplFalhaDefinitiva.EmailEnviadoStatus = Global.Cte.Braspag.Webhook.EmailEnviadoStatus.EmpresaInvalida;
							updateWebhookQueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.Webhook.BraspagDadosComplementaresQueryStatus.EmpresaInvalida;
							updateWebhookQueryComplFalhaDefinitiva.MsgErro = "Falha ao tentar obter o MerchantId para a empresa '" + pedidoWH.Empresa + "'";

							// Prossegue para o próximo pedido da lista (o bloco finally irá registrar o código e mensagem da falha)
							continue;
						}
						#endregion

						#region [ Obtém o plano de contas para gravar o lançamento no fluxo de caixa ]
						if ((pedidoWH.Empresa ?? "").Length > 0)
						{
							try
							{
								webhookBraspagPlanoContasBoletoEC = Global.Parametros.Braspag.webhookBraspagPlanoContasBoletoECList.Single(p => p.Empresa.ToUpper().Equals(pedidoWH.Empresa.ToUpper()));
							}
							catch (Exception ex)
							{
								strMsg = NOME_DESTA_ROTINA + ": exception ao tentar obter o plano de contas para gravação de lançamentos no fluxo de caixa dos boletos de e-commerce da empresa '" + pedidoWH.Empresa + "'!!" +
										"\r\n" + ex.ToString();
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Tenta localizar nº pedido ERP ]
						if (BraspagDAO.isPedidoERPDesteAmbiente(pedidoWH.NumPedido, strMerchantId, out strNumPedidoERPAux)) strNumPedidoERP = strNumPedidoERPAux;
						if (strNumPedidoERP.Length == 0)
						{
							if (GeralDAO.isPedidoECommerce(pedidoWH.NumPedido, out strNumPedidoERPAux)) strNumPedidoERP = strNumPedidoERPAux;
						}
						#endregion

						#region [ Se encontrou pedido ERP, carrega os dados ]
						if (strNumPedidoERP.Length > 0)
						{
							pedido = PedidoDAO.getPedido(strNumPedidoERP);
							if (pedido != null) cliente = ClienteDAO.getCliente(pedido.id_cliente);
						}
						#endregion

						#region [ Obtém dados consolidados do boleto ]
						rBoleto = consultaDadosConsolidadosBoletoParaWebhook(strMerchantId, pedidoWH.NumPedido, out msg_erro_aux);
						if (rBoleto == null)
						{
							msg_erro_requisicao = (msg_erro_aux ?? "");

							blnWebhookQueryComplFalhaTemporaria = true;
							updateWebhookQueryComplFalhaTemporaria = new BraspagUpdateWebhookQueryDadosComplementaresFalhaTemporaria();
							updateWebhookQueryComplFalhaTemporaria.id_braspag_webhook = pedidoWH.Id;
							updateWebhookQueryComplFalhaTemporaria.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.Webhook.BraspagDadosComplementaresQueryStatus.FalhaConsultaBraspag;
							updateWebhookQueryComplFalhaTemporaria.BraspagDadosComplementaresQueryTentativas = pedidoWH.BraspagDadosComplementaresQueryTentativas;
							updateWebhookQueryComplFalhaTemporaria.MsgErroTemporario = "Falha ao consultar dados complementares na Braspag (MerchantId = " + strMerchantId + ", OrderId = " + pedidoWH.NumPedido + ")" + (msg_erro_requisicao.Length > 0 ? ": " + msg_erro_requisicao : "");

							// Prossegue para o próximo pedido da lista (o bloco finally irá registrar o código e mensagem da falha)
							continue;
						}
						#endregion

						#region [ Se chegou até este ponto, o processamento do pedido foi bem sucedido ]
						blnWebhookQueryComplSucesso = true;
						#endregion

						#region [ Grava os dados complementares ]
						insertWebhookQueryCompl = new BraspagInsertWebhookQueryDadosComplementares();
						insertWebhookQueryCompl.id_braspag_webhook = pedidoWH.Id;
						insertWebhookQueryCompl.BraspagTransactionId = rBoleto.BraspagTransactionId;
						insertWebhookQueryCompl.BraspagOrderId = rBoleto.BraspagOrderId;
						insertWebhookQueryCompl.PaymentMethod = rBoleto.PaymentMethod;
						insertWebhookQueryCompl.GlobalStatus = rBoleto.GlobalStatus;
						insertWebhookQueryCompl.ReceivedDate = rBoleto.ReceivedDate;
						insertWebhookQueryCompl.CapturedDate = rBoleto.CapturedDate;
						insertWebhookQueryCompl.CustomerName = rBoleto.CustomerName;
						insertWebhookQueryCompl.BoletoExpirationDate = rBoleto.BoletoExpirationDate;
						insertWebhookQueryCompl.Amount = rBoleto.Amount;
						insertWebhookQueryCompl.ValorAmount = rBoleto.ValorAmount;
						insertWebhookQueryCompl.PaidAmount = rBoleto.PaidAmount;
						insertWebhookQueryCompl.ValorPaidAmount = rBoleto.ValorPaidAmount;
						insertWebhookQueryCompl.pedido = strNumPedidoERP;
						if (!BraspagDAO.insereWebhookQueryDadosComplementares(insertWebhookQueryCompl, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar gravar registro com dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar gravar registro com dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion
					}
					finally
					{
						#region [ Altera o status do registro na tabela t_BRASPAG_WEBHOOK ]
						if (blnWebhookQueryComplSucesso)
						{
							#region [ Atualiza c/ status de sucesso ]
							updateWebhookQueryComplSucesso = new BraspagUpdateWebhookQueryDadosComplementaresSucesso();
							updateWebhookQueryComplSucesso.id_braspag_webhook = pedidoWH.Id;
							updateWebhookQueryComplSucesso.BraspagDadosComplementaresQueryTentativas = pedidoWH.BraspagDadosComplementaresQueryTentativas;
							updateWebhookQueryComplSucesso.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.Webhook.BraspagDadosComplementaresQueryStatus.ProcessadoComSucesso;
							if (!BraspagDAO.updateWebhookQueryDadosComplementaresSucesso(updateWebhookQueryComplSucesso, out msg_erro_aux))
							{
								msg_erro_last_op = msg_erro_aux;

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando sucesso na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
								svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookQueryComplSucesso);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o banco de dados com o status de sucesso ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de sucesso ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
							}
							#endregion
						}
						else
						{
							#region [ Atualiza c/ status de falha (definitiva ou temporária) ]
							if (pedidoWH.BraspagDadosComplementaresQueryTentativas >= Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_MaxTentativasQueryDadosComplementares)
							{
								#region [ Excedeu quantidade máxima de tentativas (falha definitiva) ]

								#region [ Atualiza o registro c/ o status de falha definitiva ]
								updateWebhookQueryComplFalhaDefinitiva = new BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva();
								updateWebhookQueryComplFalhaDefinitiva.id_braspag_webhook = pedidoWH.Id;
								updateWebhookQueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryTentativas = pedidoWH.BraspagDadosComplementaresQueryTentativas;
								updateWebhookQueryComplFalhaDefinitiva.EmailEnviadoStatus = Global.Cte.Braspag.Webhook.EmailEnviadoStatus.ExcedeuMaxTentativasQueryDadosComplementares;
								updateWebhookQueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.Webhook.BraspagDadosComplementaresQueryStatus.ExcedeuMaxTentativasQueryDadosComplementares;
								updateWebhookQueryComplFalhaDefinitiva.MsgErro = "Excedeu quantidade máxima de tentativas: " + Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_MaxTentativasQueryDadosComplementares.ToString();
								// Altera o status e registra mensagem de erro
								if (!BraspagDAO.updateWebhookQueryDadosComplementaresFalhaDefinitiva(updateWebhookQueryComplFalhaDefinitiva, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha definitiva na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookQueryComplFalhaDefinitiva);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									#region [ Envia email de alerta sobre a falha na atualização do BD ]
									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o banco de dados com o status de falha definitiva ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha definitiva por exceder o limite máximo de " +
												Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_MaxTentativasQueryDadosComplementares.ToString() +
												" tentativas de obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
									#endregion
								}
								#endregion

								#region [ Envia email informando da falha definitiva na consulta dos dados complementares ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha definitiva ao tentar obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha definitiva por exceder o limite máximo de " +
											Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspag_MaxTentativasQueryDadosComplementares.ToString() +
											" tentativas de obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion

								#endregion
							}
							else if (blnWebhookQueryComplFalhaDefinitiva)
							{
								#region [ Ocorreu uma falha definitiva ]

								#region [ Atualiza o banco de dados c/ o status de falha definitiva ]
								// Altera o status e registra mensagem de erro
								if (!BraspagDAO.updateWebhookQueryDadosComplementaresFalhaDefinitiva(updateWebhookQueryComplFalhaDefinitiva, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha definitiva na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookQueryComplFalhaDefinitiva);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o banco de dados com o status de falha definitiva ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha definitiva ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
								#endregion

								#region [ Envia email de alerta sobre a falha definitiva ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha definitiva ao tentar obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha definitiva ao tentar obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion
								#endregion
							}
							else if (blnWebhookQueryComplFalhaTemporaria)
							{
								#region [ Falha temporária, apenas incrementa o contador de tentativas ]
								if (!BraspagDAO.updateWebhookQueryDadosComplementaresFalhaTemporaria(updateWebhookQueryComplFalhaTemporaria, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha temporária na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookQueryComplFalhaTemporaria);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o banco de dados com o status de falha temporária ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha temporária ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
								#endregion
							}
							else
							{
								#region [ Precaução: esta situação não deve ocorrer, mas caso ocorra, apenas atualiza o contador de tentativas ]
								updateWebhookQueryComplQtdeTentativas = new BraspagUpdateWebhookQueryDadosComplementaresQtdeTentativas();
								updateWebhookQueryComplQtdeTentativas.id_braspag_webhook = pedidoWH.Id;
								updateWebhookQueryComplQtdeTentativas.BraspagDadosComplementaresQueryTentativas = pedidoWH.BraspagDadosComplementaresQueryTentativas;
								if (!BraspagDAO.updateWebhookQueryDadosComplementaresQtdeTentativas(updateWebhookQueryComplQtdeTentativas, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha temporária desconhecida na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookQueryComplQtdeTentativas);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o banco de dados com o status de falha temporária desconhecida ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha temporária desconhecida ao obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}

								#region [ Envia email de alerta sobre falha desconhecida ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha temporária desconhecida ao tentar obter os dados complementares do pedido " + pedidoWH.NumPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha temporária desconhecida ao tentar obter os dados complementares do pedido " + pedidoWH.NumPedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion
								#endregion
							}
							#endregion
						}
						#endregion
					} // Finally
					#endregion

					#region [ Falha na obtenção dos dados complementares? ]
					if (!blnWebhookQueryComplSucesso)
					{
						// Prossegue para o próximo pedido da lista (o bloco finally anterior já registrou o código e mensagem da falha)
						continue;
					}
					#endregion

					#region [ Calcula variação percentual do valor pago ]
					percDif = 0m;
					if (rBoleto.ValorAmount > 0)
					{
						percDif = (rBoleto.ValorPaidAmount - rBoleto.ValorAmount) / rBoleto.ValorAmount;
					}
					#endregion

					#region [ Analisa e processa o registro do pagamento no pedido e alteração do status da análise de crédito ]

					#region [ Verifica se o BraspagTransactionId já foi processado anteriormente (status 'Capturado') ]
					if (BraspagDAO.transacaoJaRegistrouPagtoNoPedido(rBoleto.BraspagTransactionId, out braspagWebhookComplementar, out msg_erro_aux))
					{
						#region [ Atualiza o banco de dados c/ o status de falha definitiva ]
						// Altera o status e registra mensagem de erro
						updateWebhookQueryComplFalhaDefinitiva = new BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva();
						updateWebhookQueryComplFalhaDefinitiva.id_braspag_webhook = pedidoWH.Id;
						updateWebhookQueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryTentativas = pedidoWH.BraspagDadosComplementaresQueryTentativas;
						updateWebhookQueryComplFalhaDefinitiva.EmailEnviadoStatus = Global.Cte.Braspag.Webhook.EmailEnviadoStatus.TransacaoJaProcessadaAnteriormente;
						updateWebhookQueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.Webhook.BraspagDadosComplementaresQueryStatus.TransacaoJaProcessadaAnteriormente;
						updateWebhookQueryComplFalhaDefinitiva.MsgErro = "A transação BraspagTransactionId=" + braspagWebhookComplementar.BraspagTransactionId + " já foi processada anteriormente registrando o pagamento no pedido (data: " + Global.formataDataDdMmYyyyHhMmSsComSeparador(braspagWebhookComplementar.DataHora) + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + braspagWebhookComplementar.id_braspag_webhook.ToString() + ")";
						if (!BraspagDAO.updateWebhookQueryDadosComplementaresFalhaDefinitiva(updateWebhookQueryComplFalhaDefinitiva, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status indicando que a transação já foi processada anteriormente (pedido=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
							svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookQueryComplFalhaDefinitiva);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o banco de dados com o status indicando que a transação já foi processada anteriormente (pedido: " + pedidoWH.NumPedido + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status indicando que a transação já foi processada anteriormente (pedido=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Envia email de alerta sobre a falha definitiva ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): transação já foi processada anteriormente (pedido: " + pedidoWH.NumPedido + ", BraspagTransactionId=" + rBoleto.BraspagTransactionId + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nTransação já foi processada anteriormente (BraspagTransactionId=" + rBoleto.BraspagTransactionId + ")\r\n" +
									"Registro atual: pedido=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + "\r\n" +
									"Processamento anterior: data=" + Global.formataDataDdMmYyyyHhMmSsComSeparador(braspagWebhookComplementar.DataHora) + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + braspagWebhookComplementar.id_braspag_webhook.ToString();
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion

						#region [ Monta os dados para email informativo ]
						sbDadosEmail = new StringBuilder("");
						sbDadosEmail.AppendLine("Pedido: " + pedidoWH.NumPedido);
						if (!pedidoWH.NumPedido.Equals(strNumPedidoERP)) sbDadosEmail.AppendLine("Pedido (ERP): " + (strNumPedidoERP.Length > 0 ? strNumPedidoERP : "não localizado"));
						sbDadosEmail.AppendLine("Cliente: " + rBoleto.CustomerName.ToUpper());
						sbDadosEmail.AppendLine("Meio de Pagamento: " + pedidoWH.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(pedidoWH.CODPAGAMENTO));
						sbDadosEmail.AppendLine("Cedente: " + pedidoWH.Empresa);
						sbDadosEmail.AppendLine("Data Vencto:  " + Global.formataDataDdMmYyyyComSeparador(rBoleto.BoletoExpirationDate));
						sbDadosEmail.AppendLine("Data Crédito: " + Global.formataDataDdMmYyyyComSeparador(rBoleto.CapturedDate));
						sbDadosEmail.AppendLine("Valor Face: " + Global.formataMoeda(rBoleto.ValorAmount));
						sbDadosEmail.AppendLine("Valor Pago: " + Global.formataMoeda(rBoleto.ValorPaidAmount));
						sbDadosEmail.AppendLine("Variação Valor: " + Global.formataMoeda(rBoleto.ValorPaidAmount - rBoleto.ValorAmount) + "  (" + Global.formataPercentualCom2Decimais(100m * percDif) + "%)");
						sbDadosEmail.AppendLine("Observação: este pagamento já foi processado anteriormente em " + Global.formataDataDdMmYyyyHhMmSsComSeparador(braspagWebhookComplementar.DataHora));
						vDadosEmail.Add(sbDadosEmail);
						#endregion
					}
					else
					{
						if (strNumPedidoERP.Length > 0)
						{
							#region [ Verifica se o pagamento já foi registrado manualmente ]
							listaPagto = PedidoDAO.getPedidoPagamentoByPedido(strNumPedidoERP, out msg_erro_aux);
							if (listaPagto != null)
							{
								foreach (PedidoPagamento pagto in listaPagto)
								{
									// Analisa somente valores positivos
									if (pagto.valor > 0)
									{
										if (Math.Abs(rBoleto.ValorPaidAmount - pagto.valor) <= Global.Cte.Etc.MAX_VALOR_MARGEM_ERRO_PAGAMENTO)
										{
											blnPagtoRegistradoManualmente = true;
											pagtoManual = pagto;
											break;
										}
									}
								}
							}

							// Se não encontrou um registro de pagamento equivalente ao valor do boleto, analisa pelo status de pagamento do pedido
							if (!blnPagtoRegistradoManualmente)
							{
								if (pedido.st_pagto.Equals(Global.Cte.StPagtoPedido.ST_PAGTO_PAGO))
								{
									blnPagtoRegistradoManualmente = true;
								}
							}
							#endregion

							if (!blnPagtoRegistradoManualmente)
							{
								#region [ Registra o pagamento no pedido + lançamento no fluxo de caixa ]
								blnSucesso = false;
								BD.iniciaTransacao();
								try
								{
									#region [ Registra o pagamento no pedido ]
									// Obs: a própria rotina já grava um registro no log geral
									blnSucesso = BraspagDAO.registraPagamentoBoletoECNoPedido(insertWebhookQueryCompl.Id, out msg_erro_aux);
									if (blnSucesso)
									{
										blnRegistrouPagtoPedido = true;
									}
									else
									{
										msg_erro_last_op = msg_erro_aux;

										#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
										FinSvcLog svcLog = new FinSvcLog();
										svcLog.operacao = NOME_DESTA_ROTINA;
										svcLog.descricao = "Falha ao tentar registrar pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
										svcLog.complemento_1 = Global.serializaObjectToXml(pedidoWH);
										svcLog.complemento_2 = Global.serializaObjectToXml(insertWebhookQueryCompl);
										GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
										#endregion

										strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar registrar pagamento no pedido " + strNumPedidoERP + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
										strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
										if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
										{
											strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
											Global.gravaLogAtividade(strMsg);
										}
									}
									#endregion

									#region [ Registra lançamento no fluxo de caixa ]
									if (blnRegistrouPagtoPedido)
									{
										if (webhookBraspagPlanoContasBoletoEC != null)
										{
											lancamento = new LancamentoFluxoCaixaInsertDevidoBoletoEC();
											lancamento.id_conta_corrente = webhookBraspagPlanoContasBoletoEC.id_conta_corrente;
											lancamento.id_plano_contas_empresa = webhookBraspagPlanoContasBoletoEC.id_plano_contas_empresa;
											lancamento.id_plano_contas_grupo = webhookBraspagPlanoContasBoletoEC.id_plano_contas_grupo;
											lancamento.id_plano_contas_conta = webhookBraspagPlanoContasBoletoEC.id_plano_contas_conta;
											lancamento.dt_competencia = (DateTime)rBoleto.CapturedDate;
											lancamento.valor = rBoleto.ValorPaidAmount;
											lancamento.descricao = "PED " + strNumPedidoERP;
											lancamento.ctrl_pagto_id_parcela = insertWebhookQueryCompl.Id;
											lancamento.ctrl_pagto_modulo = Global.Cte.FIN.CtrlPagtoModulo.BRASPAG_WEBHOOK;
											if (pedido != null) lancamento.id_cliente = pedido.id_cliente;
											if (cliente != null) lancamento.cnpj_cpf = cliente.cnpj_cpf;

											blnSucesso = LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoEC(lancamento, out msg_erro_aux);
											if (blnSucesso)
											{
												blnRegistrouLancamento = true;

												#region [ Grava registro no log geral ]
												// Obs: a rotina de inserção do lançamento grava um registro no log financeiro
												s_log = "Inserção do registro em t_FIN_FLUXO_CAIXA.id=" + lancamento.id.ToString() + " devido a boleto de e-commerce (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + "): dt_competencia=" + Global.formataDataYyyyMmDdComSeparador(lancamento.dt_competencia) + ", valor=" + Global.formataMoeda(lancamento.valor) + ", descricao=" + lancamento.descricao;
												GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_ECOMMERCE, strNumPedidoERP, s_log, out msg_erro_aux);
												#endregion
											}
										}
									}
									#endregion
								}
								catch (Exception ex)
								{
									blnSucesso = false;
									msg_erro = ex.ToString();
								}
								finally
								{
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
											blnRegistrouPagtoPedido = false;
											blnRegistrouLancamento = false;

											msg_erro_last_op = ex.ToString();

											#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
											FinSvcLog svcLog = new FinSvcLog();
											svcLog.operacao = NOME_DESTA_ROTINA;
											svcLog.descricao = "Falha ao tentar executar o commit no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
											svcLog.complemento_1 = Global.serializaObjectToXml(pedidoWH);
											svcLog.complemento_2 = Global.serializaObjectToXml(insertWebhookQueryCompl);
											GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
											#endregion

											strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar executar o commit no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = "Mensagem de Financeiro Service\nFalha ao tentar executar o commit no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
												Global.gravaLogAtividade(strMsg);
											}
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
											msg_erro_last_op = ex.ToString();

											#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
											FinSvcLog svcLog = new FinSvcLog();
											svcLog.operacao = NOME_DESTA_ROTINA;
											svcLog.descricao = "Falha ao tentar executar o rollback no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
											svcLog.complemento_1 = Global.serializaObjectToXml(pedidoWH);
											svcLog.complemento_2 = Global.serializaObjectToXml(insertWebhookQueryCompl);
											GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
											#endregion

											strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar executar o rollback no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = "Mensagem de Financeiro Service\nFalha ao tentar executar o rollback no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWH.NumPedido + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
												Global.gravaLogAtividade(strMsg);
											}
										}
										finally
										{
											blnRegistrouPagtoPedido = false;
											blnRegistrouLancamento = false;
										}
										#endregion
									}
								}
								#endregion
							}
						}

						#region [ Atualiza campo t_BRASPAG_WEBHOOK.ProcessamentoErpStatus ]
						if (blnRegistrouPagtoPedido)
						{
							if (!BraspagDAO.updateWebhookProcessamentoErpStatusSucesso(pedidoWH.Id, out msg_erro_aux))
							{
								msg_erro_last_op = msg_erro_aux;

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de sucesso no processamento ERP (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\n" + msg_erro_last_op;
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de sucesso no processamento ERP (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de sucesso no processamento ERP (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + pedidoWH.Id.ToString() + ")\r\n" + msg_erro_last_op;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
							}
						}
						#endregion

						#region [ Obtém dados atualizados de t_BRASPAG_WEBHOOK_COMPLEMENTAR ]
						braspagWebhookComplementarAtualizado = BraspagDAO.getBraspagWebhookComplementarById(insertWebhookQueryCompl.Id, out msg_erro_aux);
						if (braspagWebhookComplementarAtualizado == null)
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar obter os dados complementares atualizados do banco de dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar obter os dados complementares atualizados do banco de dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar obter os dados complementares atualizados do banco de dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + insertWebhookQueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Monta os dados para email informativo ]
						sbDadosEmail = new StringBuilder("");
						sbDadosEmail.AppendLine("Pedido: " + pedidoWH.NumPedido);
						if (!pedidoWH.NumPedido.Equals(strNumPedidoERP)) sbDadosEmail.AppendLine("Pedido (ERP): " + (strNumPedidoERP.Length > 0 ? strNumPedidoERP : "não localizado"));
						sbDadosEmail.AppendLine("Cliente: " + rBoleto.CustomerName.ToUpper());
						sbDadosEmail.AppendLine("Meio de Pagamento: " + pedidoWH.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(pedidoWH.CODPAGAMENTO));
						sbDadosEmail.AppendLine("Cedente: " + pedidoWH.Empresa);
						sbDadosEmail.AppendLine("Data Vencto:  " + Global.formataDataDdMmYyyyComSeparador(rBoleto.BoletoExpirationDate));
						sbDadosEmail.AppendLine("Data Crédito: " + Global.formataDataDdMmYyyyComSeparador(rBoleto.CapturedDate));
						sbDadosEmail.AppendLine("Valor Face: " + Global.formataMoeda(rBoleto.ValorAmount));
						sbDadosEmail.AppendLine("Valor Pago: " + Global.formataMoeda(rBoleto.ValorPaidAmount));
						sbDadosEmail.AppendLine("Variação Valor: " + Global.formataMoeda(rBoleto.ValorPaidAmount - rBoleto.ValorAmount) + "  (" + Global.formataPercentualCom2Decimais(100m * percDif) + "%)");
						if (blnRegistrouPagtoPedido)
						{
							#region [ Informações referentes ao pagamento registrado automaticamente ]
							strMsg = "Pagamento registrado automaticamente no pedido: SIM";
							sbDadosEmail.AppendLine(strMsg);

							if (braspagWebhookComplementarAtualizado != null)
							{
								if (braspagWebhookComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoNovo.Equals(braspagWebhookComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoAnterior))
								{
									strMsg = "Não houve alteração do status de pagamento: '" + Global.stPagtoPedidoDescricao(braspagWebhookComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoNovo) + "'";
								}
								else
								{
									strMsg = "Alteração do status de pagamento: de '" + Global.stPagtoPedidoDescricao(braspagWebhookComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoAnterior) + "' para '" + Global.stPagtoPedidoDescricao(braspagWebhookComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoNovo) + "'";
								}
								sbDadosEmail.AppendLine(strMsg);

								if (braspagWebhookComplementarAtualizado.AnaliseCreditoStatusNovo == braspagWebhookComplementarAtualizado.AnaliseCreditoStatusAnterior)
								{
									strMsg = "Não houve alteração do status da análise de crédito: '" + Global.obtemDescricaoAnaliseCredito(braspagWebhookComplementarAtualizado.AnaliseCreditoStatusNovo) + "'";
								}
								else
								{
									strMsg = "Alteração do status da análise de crédito: de '" + Global.obtemDescricaoAnaliseCredito(braspagWebhookComplementarAtualizado.AnaliseCreditoStatusAnterior) + "' para '" + Global.obtemDescricaoAnaliseCredito(braspagWebhookComplementarAtualizado.AnaliseCreditoStatusNovo) + "'";
								}
								sbDadosEmail.AppendLine(strMsg);
							}
							#endregion
						}
						else if (blnPagtoRegistradoManualmente)
						{
							if (pagtoManual != null)
							{
								strMsg = "Pagamento já havia sido registrado no pedido pelo usuário '" + (pagtoManual.usuario ?? "") + "' em " + Global.formataDataDdMmYyyyComSeparador(pagtoManual.data) + " " + Global.formata_hhnnss_para_hh_nn(pagtoManual.hora) + " com o valor de " + Global.formataMoeda(pagtoManual.valor);
								sbDadosEmail.AppendLine(strMsg);
							}
							else
							{
								strMsg = "Pedido já estava com o status de pagamento '" + Global.stPagtoPedidoDescricao(pedido.st_pagto) + "'";
								sbDadosEmail.AppendLine(strMsg);
							}
						}
						else
						{
							strMsg = "Pagamento registrado automaticamente no pedido: NÃO";
							sbDadosEmail.AppendLine(strMsg);
						}

						if (blnRegistrouLancamento)
						{
							strMsg = "Lançamento do fluxo de caixa registrado automaticamente: SIM";
							sbDadosEmail.AppendLine(strMsg);
						}
						else
						{
							strMsg = "Lançamento do fluxo de caixa registrado automaticamente: NÃO";
							sbDadosEmail.AppendLine(strMsg);
						}

						vDadosEmail.Add(sbDadosEmail);
						vBraspagWebhookIdEmailEnviadoStatusUpdate.Add(pedidoWH.Id);
						#endregion
					}
					#endregion

					#endregion

				} // foreach (var pedidoWH in listaWebhook)
				#endregion

				#region [ Há dados? ]
				if (vDadosEmail.Count == 0)
				{
					strMsgInformativa = "Nenhum boleto processado";
					return true;
				}
				#endregion

				#region [ Envia o email ]
				sbBody = new StringBuilder("");
				strMsg = "Processamento automático dos boletos de e-commerce";
				sbBody.AppendLine(strMsg);
				sbBody.AppendLine("");
				sbBody.AppendLine(strLinhaSeparadora);
				sbBody.AppendLine("");
				for (int i = 0; i < vDadosEmail.Count; i++)
				{
					if (i > 0) sbBody.AppendLine("");
					sbBody.AppendLine(vDadosEmail[i].ToString());
					sbBody.AppendLine(strLinhaSeparadora);
				}

				sbBody.AppendLine("");
				strMsg = "Total de boletos: " + vDadosEmail.Count.ToString();
				sbBody.AppendLine(strMsg);

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Processamento de boletos de e-commerce [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG, null, null, strSubject, sbBody.ToString(), DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					msg_erro_send_email = "Falha ao tentar inserir email na fila de mensagens: " + msg_erro_aux;
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
					foreach (int id_braspag_webhook in vBraspagWebhookIdEmailEnviadoStatusUpdate)
					{
						if (!BraspagDAO.updateWebhookEmailEnviadoStatusFalha(id_braspag_webhook, Global.Cte.Braspag.Webhook.EmailEnviadoStatus.ErroERP, msg_erro_send_email, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de falha no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString() + ")\n" + msg_erro_last_op;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de falha no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de falha no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
					}
				}
				else
				{
					blnEmailAlertaEnviado = true;

					foreach (int id_braspag_webhook in vBraspagWebhookIdEmailEnviadoStatusUpdate)
					{
						if (!BraspagDAO.updateWebhookEmailEnviadoStatusSucesso(id_braspag_webhook, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de sucesso no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString() + ")\n" + msg_erro_last_op;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook): Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de sucesso no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com o status de sucesso no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
					}
				}
				#endregion

				strMsgInformativa = vDadosEmail.Count.ToString() + " boleto(s) processado(s)";

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}

		}
		#endregion

		#region [ executaProcessamentoWebhookV2 ]
		public static bool executaProcessamentoWebhookV2(out bool blnEmailAlertaEnviado, out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "Braspag.executaProcessamentoWebhookV2()";
			bool blnWebhookV2QueryComplSucesso;
			bool blnWebhookV2QueryComplFalhaDefinitiva;
			bool blnWebhookV2QueryComplFalhaTemporaria;
			bool blnRegistrouPagtoPedido;
			bool blnRegistrouLancamento;
			bool blnPagtoRegistradoManualmente;
			bool blnSucesso;
			byte processadoStatusResultado;
			int id_emailsndsvc_mensagem;
			decimal percDif;
			string msg_erro_aux;
			string msg_erro_requisicao;
			string strMsg;
			string strSql;
			string msg_erro_last_op;
			string msg_erro_send_email;
			string strSubject;
			string strBody;
			string strMerchantId;
			string strNumPedidoERP;
			string strNumPedidoERPAux;
			string s_log;
			String strLinhaSeparadora = new String('=', 80);
			List<StringBuilder> vDadosEmail = new List<StringBuilder>();
			StringBuilder sbDadosEmail;
			StringBuilder sbBody;
			List<int> vBraspagWebhookV2IdEmailEnviadoStatusUpdate = new List<int>();
			LancamentoFluxoCaixaInsertDevidoBoletoEC lancamento;
			Pedido pedido;
			Cliente cliente;
			List<PedidoPagamento> listaPagto;
			PedidoPagamento pagtoManual;
			Global.Parametros.Braspag.WebhookBraspagV2MerchantId webhookBraspagV2MerchantId;
			Global.Parametros.Braspag.WebhookBraspagV2PlanoContasBoletoEC webhookBraspagV2PlanoContasBoletoEC;
			BraspagWebhookV2DadosConsolidadosBoleto rBoleto;
			List<BraspagWebhookV2> listaWebhookV2 = new List<BraspagWebhookV2>();
			BraspagWebhookV2 pedidoWebhookV2;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			BraspagUpdateWebhookV2PaymentMethodIdentificado updateWebhookV2PaymentMethodIdentificado;
			BraspagUpdateWebhookV2QueryDadosComplementaresFalhaDefinitiva updateWebhookV2QueryComplFalhaDefinitiva = null;
			BraspagUpdateWebhookV2QueryDadosComplementaresFalhaTemporaria updateWebhookV2QueryComplFalhaTemporaria = null;
			BraspagUpdateWebhookV2QueryDadosComplementaresSucesso updateWebhookV2QueryComplSucesso = null;
			BraspagInsertWebhookV2QueryDadosComplementares insertWebhookV2QueryCompl;
			BraspagUpdateWebhookV2QueryDadosComplementaresQtdeTentativas updateWebhookV2QueryComplQtdeTentativas = null;
			BraspagWebhookV2Complementar braspagWebhookV2ComplementarAtualizado;
			BraspagWebhookV2Complementar braspagWebhookV2Complementar;
			#endregion

			blnEmailAlertaEnviado = false;
			strMsgInformativa = "";
			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Subsídios ]
				/*
				Documentação da Braspag em: https://braspag.github.io//manual/braspag-pagador#post-de-notificação
				Post de notificação
				PROPRIEDADE				DESCRIÇÃO																							TIPO	TAMANHO		OBRIGATÓRIO?
				RecurrentPaymentId		Identificador que representa o pedido recorrente (aplicável somente para ChangeType "2" ou "4").	GUID	36			Não
				PaymentId				Identificador que representa a transação.															GUID	36			Sim
				ChangeType				Especifica o tipo de notificação. Obs.: Consulte a tabela abaixo.									Número	1			Sim
				
				CHANGETYPE	DESCRIÇÃO
				"1"			Mudança de status do pagamento.
				"2"			Recorrência criada.
				"3"			Mudança de status do Antifraude.
				"4"			Mudança de status do pagamento recorrente (Ex.: desativação automática).
				"5"			Estorno negado (aplicável para Rede).
				"6"			Boleto registrado pago a menor.
				"7"			Notificação de chargeback. Para mais detalhes, consulte o manual de Risk Notification.
				"8"			Alerta de fraude.
				
				Portanto, para processar a notificação referente aos boletos, o primeiro passo é identificar o PaymentMethod de cada PaymentId
				Observações:
					1) O campo PaymentId se refere ao mesmo valor do campo BraspagTransactionId
					2) Na versão anterior do post de notificação, a Braspag informava o valor do PaymentMethod na própria notificação através do campo CODPAGAMENTO, mas,
					   por outro lado, não informava o PaymentId e sim o OrderId (nº pedido definido pelo lojista). Isso criava a necessidade de se realizar uma consulta
					   através do OrderId para identificar o BraspagTransactionId, o que deixou de ser necessário.
				*/
				#endregion

				#region [ Processa as novas transações notificadas ]

				#region [ Atualiza o status das notificações referentes a pagamento por cartão p/ que sejam ignoradas ]
				// Com base nos dados armazenados no sistema das transações de pagamento com cartão, identifica o PaymentMethod usando a tabela t_PAGTO_GW_PAG_PAYMENT
				if (!BraspagDAO.updateWebhookV2PaymentMethodIdentificadoCartao(out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha no processamento que tenta identificar o PaymentMethod das transações de cartão usando a tabela t_PAGTO_GW_PAG_PAYMENT";
					if ((msg_erro_aux ?? "").Length > 0) strMsg += "\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}
				#endregion

				#region [ Seleciona as novas transações notificadas que necessitam de análise individual ]
				strSql = "SELECT" +
							" *" +
						" FROM t_BRASPAG_WEBHOOK_V2" +
						" WHERE" +
							" (ProcessadoStatus = " + Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.Inicial.ToString() + ")" +
							" OR " +
							"(" +
								" (ProcessadoStatus = " + Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.PaymentMethodIdentificado.ToString() + ")" +
								" AND (ProcessamentoErpStatus = 0)" +
								" AND (PaymentMethodIdentificado IN ('" + Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue() + "','" + Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue() + "'))" +
							")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					pedidoWebhookV2 = new BraspagWebhookV2();
					pedidoWebhookV2.Id = BD.readToInt(row["Id"]);
					pedidoWebhookV2.Empresa = BD.readToString(row["Empresa"]);
					pedidoWebhookV2.RecurrentPaymentId = BD.readToString(row["RecurrentPaymentId"]);
					pedidoWebhookV2.PaymentId = BD.readToString(row["PaymentId"]);
					pedidoWebhookV2.ChangeType = BD.readToByte(row["ChangeType"]);
					pedidoWebhookV2.ProcessadoStatus = BD.readToByte(row["ProcessadoStatus"]);
					pedidoWebhookV2.BraspagDadosComplementaresQueryStatus = BD.readToByte(row["BraspagDadosComplementaresQueryStatus"]);
					pedidoWebhookV2.BraspagDadosComplementaresQueryTentativas = BD.readToInt(row["BraspagDadosComplementaresQueryTentativas"]);
					pedidoWebhookV2.ProcessamentoErpStatus = BD.readToInt(row["ProcessamentoErpStatus"]);
					listaWebhookV2.Add(pedidoWebhookV2);
				}
				#endregion

				#region [ Há dados? ]
				if (listaWebhookV2.Count == 0)
				{
					strMsgInformativa = "Não há dados para processar";
					return true;
				}
				#endregion

				#region [ Processa cada uma das novas notificações que necessitam de análise individual ]
				foreach (var pedidoWHV2 in listaWebhookV2)
				{
					#region [ Contador de tentativas ]
					pedidoWHV2.BraspagDadosComplementaresQueryTentativas++;
					#endregion

					#region [ Inicialização de variáveis a cada iteração ]
					strMerchantId = "";
					webhookBraspagV2PlanoContasBoletoEC = null;
					blnWebhookV2QueryComplSucesso = false;
					blnWebhookV2QueryComplFalhaDefinitiva = false;
					blnWebhookV2QueryComplFalhaTemporaria = false;
					strNumPedidoERP = "";
					pedido = null;
					cliente = null;
					pagtoManual = null;
					blnRegistrouPagtoPedido = false;
					blnRegistrouLancamento = false;
					blnPagtoRegistradoManualmente = false;
					processadoStatusResultado = Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.Inicial;
					#endregion

					#region [ Consulta a Braspag para obter dados complementares ]
					try // Try-Finally: tratamento para falhas, principalmente se excedeu quantidade máxima de tentativas
					{
						#region [ Obtém o MerchantId p/ o pedido em questão ]
						if ((pedidoWHV2.Empresa ?? "").Length > 0)
						{
							try
							{
								webhookBraspagV2MerchantId = Global.Parametros.Braspag.webhookBraspagV2MerchantIdList.Single(p => p.Empresa.ToUpper().Equals(pedidoWHV2.Empresa.ToUpper()));
								if (webhookBraspagV2MerchantId != null)
								{
									strMerchantId = webhookBraspagV2MerchantId.MerchantId;
								}
							}
							catch (Exception ex)
							{
								strMerchantId = "";
								strMsg = NOME_DESTA_ROTINA + ": exception ao tentar obter o MerchantId para a empresa '" + pedidoWHV2.Empresa + "'!!" +
										"\r\n" + ex.ToString();
								Global.gravaLogAtividade(strMsg);
							}
						}

						if ((strMerchantId ?? "").Length == 0)
						{
							blnWebhookV2QueryComplFalhaDefinitiva = true;
							updateWebhookV2QueryComplFalhaDefinitiva = new BraspagUpdateWebhookV2QueryDadosComplementaresFalhaDefinitiva();
							updateWebhookV2QueryComplFalhaDefinitiva.id_braspag_webhook_v2 = pedidoWHV2.Id;
							updateWebhookV2QueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryTentativas = pedidoWHV2.BraspagDadosComplementaresQueryTentativas;
							updateWebhookV2QueryComplFalhaDefinitiva.EmailEnviadoStatus = Global.Cte.Braspag.WebhookV2.EmailEnviadoStatus.EmpresaInvalida;
							updateWebhookV2QueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.WebhookV2.BraspagDadosComplementaresQueryStatus.EmpresaInvalida;
							updateWebhookV2QueryComplFalhaDefinitiva.MsgErro = "Falha ao tentar obter o MerchantId para a empresa '" + pedidoWHV2.Empresa + "'";

							// Prossegue para o próximo pedido da lista (o bloco finally irá registrar o código e mensagem da falha)
							continue;
						}
						#endregion

						#region [ Obtém o plano de contas para gravar o lançamento no fluxo de caixa ]
						if ((pedidoWHV2.Empresa ?? "").Length > 0)
						{
							try
							{
								webhookBraspagV2PlanoContasBoletoEC = Global.Parametros.Braspag.webhookBraspagV2PlanoContasBoletoECList.Single(p => p.Empresa.ToUpper().Equals(pedidoWHV2.Empresa.ToUpper()));
							}
							catch (Exception ex)
							{
								strMsg = NOME_DESTA_ROTINA + ": exception ao tentar obter o plano de contas para gravação de lançamentos no fluxo de caixa dos boletos de e-commerce da empresa '" + pedidoWHV2.Empresa + "'!!" +
										"\r\n" + ex.ToString();
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Obtém dados consolidados do boleto ]
						rBoleto = consultaDadosConsolidadosBoletoParaWebhookV2(strMerchantId, pedidoWHV2.PaymentId, out msg_erro_aux);
						if (rBoleto == null)
						{
							msg_erro_requisicao = (msg_erro_aux ?? "");

							blnWebhookV2QueryComplFalhaTemporaria = true;
							updateWebhookV2QueryComplFalhaTemporaria = new BraspagUpdateWebhookV2QueryDadosComplementaresFalhaTemporaria();
							updateWebhookV2QueryComplFalhaTemporaria.id_braspag_webhook_v2 = pedidoWHV2.Id;
							updateWebhookV2QueryComplFalhaTemporaria.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.WebhookV2.BraspagDadosComplementaresQueryStatus.FalhaConsultaBraspag;
							updateWebhookV2QueryComplFalhaTemporaria.BraspagDadosComplementaresQueryTentativas = pedidoWHV2.BraspagDadosComplementaresQueryTentativas;
							updateWebhookV2QueryComplFalhaTemporaria.MsgErroTemporario = "Falha ao consultar dados complementares na Braspag (MerchantId = " + strMerchantId + ", PaymentId = " + pedidoWHV2.PaymentId + ")" + (msg_erro_requisicao.Length > 0 ? ": " + msg_erro_requisicao : "");

							// Prossegue para o próximo pedido da lista (o bloco finally irá registrar o código e mensagem da falha)
							continue;
						}
						#endregion

						#region [ Atualiza os campos identificados (PaymentMethod e OrderId ]
						updateWebhookV2PaymentMethodIdentificado = new BraspagUpdateWebhookV2PaymentMethodIdentificado();
						updateWebhookV2PaymentMethodIdentificado.Id = pedidoWHV2.Id;
						updateWebhookV2PaymentMethodIdentificado.OrderIdIdentificado = rBoleto.OrderId;
						updateWebhookV2PaymentMethodIdentificado.PaymentMethodIdentificado = rBoleto.PaymentMethod;
						if (updateWebhookV2PaymentMethodIdentificado.PaymentMethodIdentificado.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue())
							|| updateWebhookV2PaymentMethodIdentificado.PaymentMethodIdentificado.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue()))
						{
							updateWebhookV2PaymentMethodIdentificado.ProcessadoStatus = Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.PaymentMethodIdentificado;
						}
						else
						{
							updateWebhookV2PaymentMethodIdentificado.ProcessadoStatus = Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.NaoProcessado;
						}

						if (!BraspagDAO.updateWebhookV2PaymentMethodIdentificado(updateWebhookV2PaymentMethodIdentificado, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o PaymentMethod e OrderId no registro da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o PaymentMethod e OrderId no registro da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}

							// Prossegue para o próximo da lista (o bloco finally irá registrar o código e mensagem da falha)
							continue;
						}

						pedidoWHV2.PaymentMethodIdentificado = updateWebhookV2PaymentMethodIdentificado.PaymentMethodIdentificado;
						pedidoWHV2.OrderIdIdentificado = updateWebhookV2PaymentMethodIdentificado.OrderIdIdentificado;
						#endregion

						#region [ Se não for boleto, segue p/ o próximo da lista ]
						if ((!pedidoWHV2.PaymentMethodIdentificado.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue()))
							&& (!pedidoWHV2.PaymentMethodIdentificado.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue())))
						{
							continue;
						}
						#endregion

						#region [ Tenta localizar nº pedido ERP ]
						if (BraspagDAO.isPedidoERPDesteAmbiente(rBoleto.OrderId, strMerchantId, out strNumPedidoERPAux)) strNumPedidoERP = strNumPedidoERPAux;
						if (strNumPedidoERP.Length == 0)
						{
							if (GeralDAO.isPedidoECommerce(rBoleto.OrderId, out strNumPedidoERPAux)) strNumPedidoERP = strNumPedidoERPAux;
						}
						#endregion

						#region [ Se encontrou pedido ERP, carrega os dados ]
						if (strNumPedidoERP.Length > 0)
						{
							pedido = PedidoDAO.getPedido(strNumPedidoERP);
							if (pedido != null) cliente = ClienteDAO.getCliente(pedido.id_cliente);
						}
						#endregion

						#region [ Grava os dados complementares ]
						insertWebhookV2QueryCompl = new BraspagInsertWebhookV2QueryDadosComplementares();
						insertWebhookV2QueryCompl.id_braspag_webhook_v2 = pedidoWHV2.Id;
						insertWebhookV2QueryCompl.BraspagTransactionId = rBoleto.BraspagTransactionId;
						insertWebhookV2QueryCompl.BraspagOrderId = rBoleto.BraspagOrderId;
						insertWebhookV2QueryCompl.PaymentMethod = rBoleto.PaymentMethod;
						insertWebhookV2QueryCompl.GlobalStatus = rBoleto.GlobalStatus;
						insertWebhookV2QueryCompl.ReceivedDate = rBoleto.ReceivedDate;
						insertWebhookV2QueryCompl.CapturedDate = rBoleto.CapturedDate;
						insertWebhookV2QueryCompl.CustomerName = rBoleto.CustomerName;
						insertWebhookV2QueryCompl.BoletoExpirationDate = rBoleto.BoletoExpirationDate;
						insertWebhookV2QueryCompl.Amount = rBoleto.Amount;
						insertWebhookV2QueryCompl.ValorAmount = rBoleto.ValorAmount;
						insertWebhookV2QueryCompl.PaidAmount = rBoleto.PaidAmount;
						insertWebhookV2QueryCompl.ValorPaidAmount = rBoleto.ValorPaidAmount;
						insertWebhookV2QueryCompl.pedido = strNumPedidoERP;
						if (!BraspagDAO.insereWebhookV2QueryDadosComplementares(insertWebhookV2QueryCompl, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar gravar registro com dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar gravar registro com dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Se chegou até este ponto, a consulta dos dados complementares foi bem sucedida ]
						blnWebhookV2QueryComplSucesso = true;
						#endregion
					}
					finally
					{
						#region [ Altera o status da consulta de dados complementares em t_BRASPAG_WEBHOOK_V2 (BraspagDadosComplementaresQueryStatus) ]
						if (blnWebhookV2QueryComplSucesso)
						{
							#region [ Atualiza c/ status de sucesso ]
							updateWebhookV2QueryComplSucesso = new BraspagUpdateWebhookV2QueryDadosComplementaresSucesso();
							updateWebhookV2QueryComplSucesso.id_braspag_webhook_v2 = pedidoWHV2.Id;
							updateWebhookV2QueryComplSucesso.BraspagDadosComplementaresQueryTentativas = pedidoWHV2.BraspagDadosComplementaresQueryTentativas;
							updateWebhookV2QueryComplSucesso.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.WebhookV2.BraspagDadosComplementaresQueryStatus.ProcessadoComSucesso;
							if (!BraspagDAO.updateWebhookV2QueryDadosComplementaresSucesso(updateWebhookV2QueryComplSucesso, out msg_erro_aux))
							{
								msg_erro_last_op = msg_erro_aux;

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com as informações indicando sucesso na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
								svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookV2QueryComplSucesso);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o banco de dados com o status de sucesso ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de sucesso ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
							}
							#endregion
						}
						else
						{
							#region [ Atualiza c/ status de falha (definitiva ou temporária) ]
							if (pedidoWHV2.BraspagDadosComplementaresQueryTentativas >= Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspagV2_MaxTentativasQueryDadosComplementares)
							{
								#region [ Excedeu quantidade máxima de tentativas (falha definitiva) ]

								#region [ Atualiza o registro c/ o status de falha definitiva ]
								updateWebhookV2QueryComplFalhaDefinitiva = new BraspagUpdateWebhookV2QueryDadosComplementaresFalhaDefinitiva();
								updateWebhookV2QueryComplFalhaDefinitiva.id_braspag_webhook_v2 = pedidoWHV2.Id;
								updateWebhookV2QueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryTentativas = pedidoWHV2.BraspagDadosComplementaresQueryTentativas;
								updateWebhookV2QueryComplFalhaDefinitiva.EmailEnviadoStatus = Global.Cte.Braspag.WebhookV2.EmailEnviadoStatus.ExcedeuMaxTentativasQueryDadosComplementares;
								updateWebhookV2QueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.WebhookV2.BraspagDadosComplementaresQueryStatus.ExcedeuMaxTentativasQueryDadosComplementares;
								updateWebhookV2QueryComplFalhaDefinitiva.MsgErro = "Excedeu quantidade máxima de tentativas: " + Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspagV2_MaxTentativasQueryDadosComplementares.ToString();
								// Altera o status e registra mensagem de erro
								if (!BraspagDAO.updateWebhookV2QueryDadosComplementaresFalhaDefinitiva(updateWebhookV2QueryComplFalhaDefinitiva, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com as informações indicando falha definitiva na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookV2QueryComplFalhaDefinitiva);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									#region [ Envia email de alerta sobre a falha na atualização do BD ]
									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o banco de dados com o status de falha definitiva ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha definitiva por exceder o limite máximo de " +
												Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspagV2_MaxTentativasQueryDadosComplementares.ToString() +
												" tentativas de obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
									#endregion
								}
								#endregion

								#region [ Envia email informando da falha definitiva na consulta dos dados complementares ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha definitiva ao tentar obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha definitiva por exceder o limite máximo de " +
											Global.Parametros.Geral.FinSvc_ProcessamentoWebhookBraspagV2_MaxTentativasQueryDadosComplementares.ToString() +
											" tentativas de obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion

								#endregion
							}
							else if (blnWebhookV2QueryComplFalhaDefinitiva)
							{
								#region [ Ocorreu uma falha definitiva ]

								#region [ Atualiza o banco de dados c/ o status de falha definitiva ]
								// Altera o status e registra mensagem de erro
								if (!BraspagDAO.updateWebhookV2QueryDadosComplementaresFalhaDefinitiva(updateWebhookV2QueryComplFalhaDefinitiva, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com as informações indicando falha definitiva na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookV2QueryComplFalhaDefinitiva);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o banco de dados com o status de falha definitiva ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha definitiva ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
								#endregion

								#region [ Envia email de alerta sobre a falha definitiva ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha definitiva ao tentar obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha definitiva ao tentar obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion
								#endregion
							}
							else if (blnWebhookV2QueryComplFalhaTemporaria)
							{
								#region [ Falha temporária, apenas incrementa o contador de tentativas ]
								if (!BraspagDAO.updateWebhookV2QueryDadosComplementaresFalhaTemporaria(updateWebhookV2QueryComplFalhaTemporaria, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com as informações indicando falha temporária na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookV2QueryComplFalhaTemporaria);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o banco de dados com o status de falha temporária ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha temporária ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}
								#endregion
							}
							else
							{
								#region [ Precaução: esta situação não deve ocorrer, mas caso ocorra, apenas atualiza o contador de tentativas ]
								updateWebhookV2QueryComplQtdeTentativas = new BraspagUpdateWebhookV2QueryDadosComplementaresQtdeTentativas();
								updateWebhookV2QueryComplQtdeTentativas.id_braspag_webhook_v2 = pedidoWHV2.Id;
								updateWebhookV2QueryComplQtdeTentativas.BraspagDadosComplementaresQueryTentativas = pedidoWHV2.BraspagDadosComplementaresQueryTentativas;
								if (!BraspagDAO.updateWebhookV2QueryDadosComplementaresQtdeTentativas(updateWebhookV2QueryComplQtdeTentativas, out msg_erro_aux))
								{
									msg_erro_last_op = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com as informações indicando falha temporária desconhecida na obtenção dos dados complementares (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
									svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookV2QueryComplQtdeTentativas);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o banco de dados com o status de falha temporária desconhecida ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status de falha temporária desconhecida ao obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
								}

								#region [ Envia email de alerta sobre falha desconhecida ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha temporária desconhecida ao tentar obter os dados complementares da transação " + pedidoWHV2.PaymentId + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha temporária desconhecida ao tentar obter os dados complementares da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion
								#endregion
							}
							#endregion
						}
						#endregion
					} // Finally
					#endregion

					#region [ Falha na obtenção dos dados complementares? ]
					if (!blnWebhookV2QueryComplSucesso)
					{
						// Prossegue para o próximo pedido da lista (o bloco finally anterior já registrou o código e mensagem da falha)
						continue;
					}
					#endregion

					#region [ Calcula variação percentual do valor pago ]
					percDif = 0m;
					if (rBoleto.ValorAmount > 0)
					{
						percDif = (rBoleto.ValorPaidAmount - rBoleto.ValorAmount) / rBoleto.ValorAmount;
					}
					#endregion

					#region [ Analisa e processa o registro do pagamento no pedido e alteração do status da análise de crédito ]
					if (BraspagDAO.transacaoJaRegistrouPagtoNoPedidoV2(rBoleto.BraspagTransactionId, out braspagWebhookV2Complementar, out msg_erro_aux))
					{
						#region [ BraspagTransactionId já foi processado anteriormente (status 'Capturado') ]
						processadoStatusResultado = Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.TransacaoJaProcessadaAnteriormente;

						#region [ Atualiza o banco de dados c/ o status de falha definitiva ]
						// Altera o status e registra mensagem de erro
						updateWebhookV2QueryComplFalhaDefinitiva = new BraspagUpdateWebhookV2QueryDadosComplementaresFalhaDefinitiva();
						updateWebhookV2QueryComplFalhaDefinitiva.id_braspag_webhook_v2 = pedidoWHV2.Id;
						updateWebhookV2QueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryTentativas = pedidoWHV2.BraspagDadosComplementaresQueryTentativas;
						updateWebhookV2QueryComplFalhaDefinitiva.EmailEnviadoStatus = Global.Cte.Braspag.WebhookV2.EmailEnviadoStatus.TransacaoJaProcessadaAnteriormente;
						updateWebhookV2QueryComplFalhaDefinitiva.BraspagDadosComplementaresQueryStatus = Global.Cte.Braspag.WebhookV2.BraspagDadosComplementaresQueryStatus.TransacaoJaProcessadaAnteriormente;
						updateWebhookV2QueryComplFalhaDefinitiva.MsgErro = "A transação BraspagTransactionId=" + braspagWebhookV2Complementar.BraspagTransactionId + " já foi processada anteriormente registrando o pagamento no pedido (data: " + Global.formataDataDdMmYyyyHhMmSsComSeparador(braspagWebhookV2Complementar.DataHora) + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + braspagWebhookV2Complementar.id_braspag_webhook_v2.ToString() + ")";
						if (!BraspagDAO.updateWebhookV2QueryDadosComplementaresFalhaDefinitiva(updateWebhookV2QueryComplFalhaDefinitiva, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status indicando que a transação já foi processada anteriormente (pedido=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
							svcLog.complemento_1 = Global.serializaObjectToXml(updateWebhookV2QueryComplFalhaDefinitiva);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o banco de dados com o status indicando que a transação já foi processada anteriormente (pedido: " + pedidoWHV2.OrderIdIdentificado + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o banco de dados com o status indicando que a transação já foi processada anteriormente (pedido=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Envia email de alerta sobre a falha definitiva ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): transação já foi processada anteriormente (pedido: " + pedidoWHV2.OrderIdIdentificado + ", BraspagTransactionId=" + rBoleto.BraspagTransactionId + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nTransação já foi processada anteriormente (BraspagTransactionId=" + rBoleto.BraspagTransactionId + ")\r\n" +
									"Registro atual: pedido=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + "\r\n" +
									"Processamento anterior: data=" + Global.formataDataDdMmYyyyHhMmSsComSeparador(braspagWebhookV2Complementar.DataHora) + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + braspagWebhookV2Complementar.id_braspag_webhook_v2.ToString();
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion

						#region [ Monta os dados para email informativo ]
						sbDadosEmail = new StringBuilder("");
						sbDadosEmail.AppendLine("Pedido: " + pedidoWHV2.OrderIdIdentificado);
						if (!pedidoWHV2.OrderIdIdentificado.Equals(strNumPedidoERP)) sbDadosEmail.AppendLine("Pedido (ERP): " + (strNumPedidoERP.Length > 0 ? strNumPedidoERP : "não localizado"));
						sbDadosEmail.AppendLine("Cliente: " + rBoleto.CustomerName.ToUpper());
						sbDadosEmail.AppendLine("Meio de Pagamento: " + pedidoWHV2.PaymentMethodIdentificado + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(pedidoWHV2.PaymentMethodIdentificado));
						sbDadosEmail.AppendLine("Cedente: " + pedidoWHV2.Empresa);
						sbDadosEmail.AppendLine("Data Vencto:  " + Global.formataDataDdMmYyyyComSeparador(rBoleto.BoletoExpirationDate));
						sbDadosEmail.AppendLine("Data Crédito: " + Global.formataDataDdMmYyyyComSeparador(rBoleto.CapturedDate));
						sbDadosEmail.AppendLine("Valor Face: " + Global.formataMoeda(rBoleto.ValorAmount));
						sbDadosEmail.AppendLine("Valor Pago: " + Global.formataMoeda(rBoleto.ValorPaidAmount));
						sbDadosEmail.AppendLine("Variação Valor: " + Global.formataMoeda(rBoleto.ValorPaidAmount - rBoleto.ValorAmount) + "  (" + Global.formataPercentualCom2Decimais(100m * percDif) + "%)");
						sbDadosEmail.AppendLine("Observação: este pagamento já foi processado anteriormente em " + Global.formataDataDdMmYyyyHhMmSsComSeparador(braspagWebhookV2Complementar.DataHora));
						vDadosEmail.Add(sbDadosEmail);
						#endregion
						#endregion
					}
					else
					{
						#region [ Processa o pagamento no pedido ]
						if (strNumPedidoERP.Length > 0)
						{
							#region [ Verifica se o pagamento já foi registrado manualmente ]
							listaPagto = PedidoDAO.getPedidoPagamentoByPedido(strNumPedidoERP, out msg_erro_aux);
							if (listaPagto != null)
							{
								foreach (PedidoPagamento pagto in listaPagto)
								{
									// Analisa somente valores positivos
									if (pagto.valor > 0)
									{
										if (Math.Abs(rBoleto.ValorPaidAmount - pagto.valor) <= Global.Cte.Etc.MAX_VALOR_MARGEM_ERRO_PAGAMENTO)
										{
											blnPagtoRegistradoManualmente = true;
											pagtoManual = pagto;
											break;
										}
									}
								}
							}

							// Se não encontrou um registro de pagamento equivalente ao valor do boleto, analisa pelo status de pagamento do pedido
							if (!blnPagtoRegistradoManualmente)
							{
								if (pedido.st_pagto.Equals(Global.Cte.StPagtoPedido.ST_PAGTO_PAGO))
								{
									blnPagtoRegistradoManualmente = true;
								}
							}

							if (blnPagtoRegistradoManualmente) processadoStatusResultado = Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.PagamentoJaRegistrado;
							#endregion

							if (!blnPagtoRegistradoManualmente)
							{
								#region [ Registra o pagamento no pedido + lançamento no fluxo de caixa ]
								blnSucesso = false;
								BD.iniciaTransacao();
								try
								{
									#region [ Registra o pagamento no pedido ]
									// Obs: a própria rotina já grava um registro no log geral
									blnSucesso = BraspagDAO.registraPagamentoBoletoECNoPedidoV2(insertWebhookV2QueryCompl.Id, out msg_erro_aux);
									if (blnSucesso)
									{
										blnRegistrouPagtoPedido = true;
									}
									else
									{
										msg_erro_last_op = msg_erro_aux;

										#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
										FinSvcLog svcLog = new FinSvcLog();
										svcLog.operacao = NOME_DESTA_ROTINA;
										svcLog.descricao = "Falha ao tentar registrar pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
										svcLog.complemento_1 = Global.serializaObjectToXml(pedidoWHV2);
										svcLog.complemento_2 = Global.serializaObjectToXml(insertWebhookV2QueryCompl);
										GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
										#endregion

										strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar registrar pagamento no pedido " + strNumPedidoERP + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
										strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
										if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
										{
											strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
											Global.gravaLogAtividade(strMsg);
										}
									}
									#endregion

									#region [ Registra lançamento no fluxo de caixa ]
									if (blnRegistrouPagtoPedido)
									{
										if (webhookBraspagV2PlanoContasBoletoEC != null)
										{
											lancamento = new LancamentoFluxoCaixaInsertDevidoBoletoEC();
											lancamento.id_conta_corrente = webhookBraspagV2PlanoContasBoletoEC.id_conta_corrente;
											lancamento.id_plano_contas_empresa = webhookBraspagV2PlanoContasBoletoEC.id_plano_contas_empresa;
											lancamento.id_plano_contas_grupo = webhookBraspagV2PlanoContasBoletoEC.id_plano_contas_grupo;
											lancamento.id_plano_contas_conta = webhookBraspagV2PlanoContasBoletoEC.id_plano_contas_conta;
											lancamento.dt_competencia = (DateTime)rBoleto.CapturedDate;
											lancamento.valor = rBoleto.ValorPaidAmount;
											lancamento.descricao = "PED " + strNumPedidoERP;
											lancamento.ctrl_pagto_id_parcela = insertWebhookV2QueryCompl.Id;
											lancamento.ctrl_pagto_modulo = Global.Cte.FIN.CtrlPagtoModulo.BRASPAG_WEBHOOK_V2;
											if (pedido != null) lancamento.id_cliente = pedido.id_cliente;
											if (cliente != null) lancamento.cnpj_cpf = cliente.cnpj_cpf;

											blnSucesso = LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoEC(lancamento, out msg_erro_aux);
											if (blnSucesso)
											{
												blnRegistrouLancamento = true;

												#region [ Grava registro no log geral ]
												// Obs: a rotina de inserção do lançamento grava um registro no log financeiro
												s_log = "Inserção do registro em t_FIN_FLUXO_CAIXA.id=" + lancamento.id.ToString() + " devido a boleto de e-commerce (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + "): dt_competencia=" + Global.formataDataYyyyMmDdComSeparador(lancamento.dt_competencia) + ", valor=" + Global.formataMoeda(lancamento.valor) + ", descricao=" + lancamento.descricao;
												GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_ECOMMERCE, strNumPedidoERP, s_log, out msg_erro_aux);
												#endregion
											}
										}
									}
									#endregion

									if (blnSucesso) processadoStatusResultado = Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.Sucesso;
								}
								catch (Exception ex)
								{
									blnSucesso = false;
									msg_erro = ex.ToString();
								}
								finally
								{
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
											blnRegistrouPagtoPedido = false;
											blnRegistrouLancamento = false;

											msg_erro_last_op = ex.ToString();

											#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
											FinSvcLog svcLog = new FinSvcLog();
											svcLog.operacao = NOME_DESTA_ROTINA;
											svcLog.descricao = "Falha ao tentar executar o commit no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
											svcLog.complemento_1 = Global.serializaObjectToXml(pedidoWHV2);
											svcLog.complemento_2 = Global.serializaObjectToXml(insertWebhookV2QueryCompl);
											GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
											#endregion

											strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar executar o commit no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = "Mensagem de Financeiro Service\nFalha ao tentar executar o commit no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
												Global.gravaLogAtividade(strMsg);
											}
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
											msg_erro_last_op = ex.ToString();

											#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
											FinSvcLog svcLog = new FinSvcLog();
											svcLog.operacao = NOME_DESTA_ROTINA;
											svcLog.descricao = "Falha ao tentar executar o rollback no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
											svcLog.complemento_1 = Global.serializaObjectToXml(pedidoWHV2);
											svcLog.complemento_2 = Global.serializaObjectToXml(insertWebhookV2QueryCompl);
											GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
											#endregion

											strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar executar o rollback no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
											strBody = "Mensagem de Financeiro Service\nFalha ao tentar executar o rollback no banco de dados ao registrar o pagamento no pedido " + strNumPedidoERP + " (OrderId=" + pedidoWHV2.OrderIdIdentificado + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
											if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
											{
												strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
												Global.gravaLogAtividade(strMsg);
											}
										}
										finally
										{
											blnRegistrouPagtoPedido = false;
											blnRegistrouLancamento = false;
										}
										#endregion
									}
								}
								#endregion
							}
						}

						#region [ Atualiza campo t_BRASPAG_WEBHOOK_V2.ProcessamentoErpStatus ]
						if (blnRegistrouPagtoPedido)
						{
							if (!BraspagDAO.updateWebhookV2ProcessamentoErpStatusSucesso(pedidoWHV2.Id, out msg_erro_aux))
							{
								msg_erro_last_op = msg_erro_aux;

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de sucesso no processamento ERP (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\n" + msg_erro_last_op;
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de sucesso no processamento ERP (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de sucesso no processamento ERP (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
							}
						}
						#endregion

						#region [ Obtém dados atualizados de t_BRASPAG_WEBHOOK_V2_COMPLEMENTAR ]
						braspagWebhookV2ComplementarAtualizado = BraspagDAO.getBraspagWebhookV2ComplementarById(insertWebhookV2QueryCompl.Id, out msg_erro_aux);
						if (braspagWebhookV2ComplementarAtualizado == null)
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar obter os dados complementares atualizados do banco de dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\n" + msg_erro_last_op;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar obter os dados complementares atualizados do banco de dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar obter os dados complementares atualizados do banco de dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR + ".Id=" + insertWebhookV2QueryCompl.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion

						#region [ Monta os dados para email informativo ]
						sbDadosEmail = new StringBuilder("");
						sbDadosEmail.AppendLine("Pedido: " + pedidoWHV2.OrderIdIdentificado);
						if (!pedidoWHV2.OrderIdIdentificado.Equals(strNumPedidoERP)) sbDadosEmail.AppendLine("Pedido (ERP): " + (strNumPedidoERP.Length > 0 ? strNumPedidoERP : "não localizado"));
						sbDadosEmail.AppendLine("Cliente: " + rBoleto.CustomerName.ToUpper());
						sbDadosEmail.AppendLine("Meio de Pagamento: " + pedidoWHV2.PaymentMethodIdentificado + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(pedidoWHV2.PaymentMethodIdentificado));
						sbDadosEmail.AppendLine("Cedente: " + pedidoWHV2.Empresa);
						sbDadosEmail.AppendLine("Data Vencto:  " + Global.formataDataDdMmYyyyComSeparador(rBoleto.BoletoExpirationDate));
						sbDadosEmail.AppendLine("Data Crédito: " + Global.formataDataDdMmYyyyComSeparador(rBoleto.CapturedDate));
						sbDadosEmail.AppendLine("Valor Face: " + Global.formataMoeda(rBoleto.ValorAmount));
						sbDadosEmail.AppendLine("Valor Pago: " + Global.formataMoeda(rBoleto.ValorPaidAmount));
						sbDadosEmail.AppendLine("Variação Valor: " + Global.formataMoeda(rBoleto.ValorPaidAmount - rBoleto.ValorAmount) + "  (" + Global.formataPercentualCom2Decimais(100m * percDif) + "%)");
						if (blnRegistrouPagtoPedido)
						{
							#region [ Informações referentes ao pagamento registrado automaticamente ]
							strMsg = "Pagamento registrado automaticamente no pedido: SIM";
							sbDadosEmail.AppendLine(strMsg);

							if (braspagWebhookV2ComplementarAtualizado != null)
							{
								if (braspagWebhookV2ComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoNovo.Equals(braspagWebhookV2ComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoAnterior))
								{
									strMsg = "Não houve alteração do status de pagamento: '" + Global.stPagtoPedidoDescricao(braspagWebhookV2ComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoNovo) + "'";
								}
								else
								{
									strMsg = "Alteração do status de pagamento: de '" + Global.stPagtoPedidoDescricao(braspagWebhookV2ComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoAnterior) + "' para '" + Global.stPagtoPedidoDescricao(braspagWebhookV2ComplementarAtualizado.PagtoRegistradoNoPedidoStPagtoNovo) + "'";
								}
								sbDadosEmail.AppendLine(strMsg);

								if (braspagWebhookV2ComplementarAtualizado.AnaliseCreditoStatusNovo == braspagWebhookV2ComplementarAtualizado.AnaliseCreditoStatusAnterior)
								{
									strMsg = "Não houve alteração do status da análise de crédito: '" + Global.obtemDescricaoAnaliseCredito(braspagWebhookV2ComplementarAtualizado.AnaliseCreditoStatusNovo) + "'";
								}
								else
								{
									strMsg = "Alteração do status da análise de crédito: de '" + Global.obtemDescricaoAnaliseCredito(braspagWebhookV2ComplementarAtualizado.AnaliseCreditoStatusAnterior) + "' para '" + Global.obtemDescricaoAnaliseCredito(braspagWebhookV2ComplementarAtualizado.AnaliseCreditoStatusNovo) + "'";
								}
								sbDadosEmail.AppendLine(strMsg);
							}
							#endregion
						}
						else if (blnPagtoRegistradoManualmente)
						{
							if (pagtoManual != null)
							{
								strMsg = "Pagamento já havia sido registrado no pedido pelo usuário '" + (pagtoManual.usuario ?? "") + "' em " + Global.formataDataDdMmYyyyComSeparador(pagtoManual.data) + " " + Global.formata_hhnnss_para_hh_nn(pagtoManual.hora) + " com o valor de " + Global.formataMoeda(pagtoManual.valor);
								sbDadosEmail.AppendLine(strMsg);
							}
							else
							{
								strMsg = "Pedido já estava com o status de pagamento '" + Global.stPagtoPedidoDescricao(pedido.st_pagto) + "'";
								sbDadosEmail.AppendLine(strMsg);
							}
						}
						else
						{
							strMsg = "Pagamento registrado automaticamente no pedido: NÃO";
							sbDadosEmail.AppendLine(strMsg);
						}

						if (blnRegistrouLancamento)
						{
							strMsg = "Lançamento do fluxo de caixa registrado automaticamente: SIM";
							sbDadosEmail.AppendLine(strMsg);
						}
						else
						{
							strMsg = "Lançamento do fluxo de caixa registrado automaticamente: NÃO";
							sbDadosEmail.AppendLine(strMsg);
						}

						vDadosEmail.Add(sbDadosEmail);
						vBraspagWebhookV2IdEmailEnviadoStatusUpdate.Add(pedidoWHV2.Id);
						#endregion

						#endregion
					}
					#endregion

					#region [ Atualiza o status de processamento da notificação ]
					if (processadoStatusResultado != Global.Cte.Braspag.WebhookV2.NotificacaoProcessadoStatus.Inicial)
					{
						if (!BraspagDAO.updateWebhookV2ProcessadoStatus(pedidoWHV2.Id, processadoStatusResultado, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o status do campo ProcessadoStatus em " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o status do campo ProcessadoStatus em " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " no registro da transação " + pedidoWHV2.PaymentId + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + pedidoWHV2.Id.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
					}
					#endregion
				} // foreach (var pedidoWHV2 in listaWebhookV2)
				#endregion

				#region [ Há dados? ]
				if (vDadosEmail.Count == 0)
				{
					strMsgInformativa = "Nenhum boleto processado";
					return true;
				}
				#endregion

				#region [ Envia o email ]
				sbBody = new StringBuilder("");
				strMsg = "Processamento automático dos boletos de e-commerce";
				sbBody.AppendLine(strMsg);
				sbBody.AppendLine("");
				sbBody.AppendLine(strLinhaSeparadora);
				sbBody.AppendLine("");
				for (int i = 0; i < vDadosEmail.Count; i++)
				{
					if (i > 0) sbBody.AppendLine("");
					sbBody.AppendLine(vDadosEmail[i].ToString());
					sbBody.AppendLine(strLinhaSeparadora);
				}

				sbBody.AppendLine("");
				strMsg = "Total de boletos: " + vDadosEmail.Count.ToString();
				sbBody.AppendLine(strMsg);

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Processamento de boletos de e-commerce [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG_V2, null, null, strSubject, sbBody.ToString(), DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					msg_erro_send_email = "Falha ao tentar inserir email na fila de mensagens: " + msg_erro_aux;
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
					foreach (int id_braspag_webhook_v2 in vBraspagWebhookV2IdEmailEnviadoStatusUpdate)
					{
						if (!BraspagDAO.updateWebhookV2EmailEnviadoStatusFalha(id_braspag_webhook_v2, Global.Cte.Braspag.WebhookV2.EmailEnviadoStatus.ErroERP, msg_erro_send_email, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de falha no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + id_braspag_webhook_v2.ToString() + ")\n" + msg_erro_last_op;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de falha no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + id_braspag_webhook_v2.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de falha no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + id_braspag_webhook_v2.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
					}
				}
				else
				{
					blnEmailAlertaEnviado = true;

					foreach (int id_braspag_webhook_v2 in vBraspagWebhookV2IdEmailEnviadoStatusUpdate)
					{
						if (!BraspagDAO.updateWebhookV2EmailEnviadoStatusSucesso(id_braspag_webhook_v2, out msg_erro_aux))
						{
							msg_erro_last_op = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de sucesso no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + id_braspag_webhook_v2.ToString() + ")\n" + msg_erro_last_op;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag (Webhook V2): Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de sucesso no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + id_braspag_webhook_v2.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + " com o status de sucesso no envio do email informativo (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_V2 + ".Id=" + id_braspag_webhook_v2.ToString() + ")\r\n" + msg_erro_last_op;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
					}
				}
				#endregion

				strMsgInformativa = vDadosEmail.Count.ToString() + " boleto(s) processado(s)";
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ executaProcessamentoEstornosPendentes ]
		public static bool executaProcessamentoEstornosPendentes(out int qtdeEstornosPendentesVerificados, out int qtdeEstornosConfirmados, out int qtdeEstornosAbortados, out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "Braspag.executaProcessamentoEstornosPendentes()";
			int qtdeVerificadoTotal = 0;
			int qtdeEstornoConfirmado = 0;
			int qtdeEstornoAindaPendente = 0;
			int qtdeStatusInvalido = 0;
			int qtdeFalha = 0;
			int qtdeAbortado = 0;
			int id_emailsndsvc_mensagem;
			bool st_estorno_confirmado;
			string msg_erro_aux;
			string msg_erro_last_op;
			string strSql;
			string strMsg;
			string ult_GlobalStatus_atualizado;
			string strSubject;
			string strBody;
			DateTime dtCorte;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			List<BraspagPagPayment> listaTrx = new List<BraspagPagPayment>();
			BraspagPagPayment trx;
			BraspagPag pag;
			BraspagUpdatePagPaymentRefundPendingFalha rUpdateRefundPendingFalha;
			StringBuilder sbFalha = new StringBuilder("");
			StringBuilder sbEstornoConfirmado = new StringBuilder("");
			StringBuilder sbEstornoAindaPendente = new StringBuilder("");
			StringBuilder sbStatusInvalido = new StringBuilder("");
			StringBuilder sbAbortado = new StringBuilder("");
			#endregion

			qtdeEstornosPendentesVerificados = 0;
			qtdeEstornosConfirmados = 0;
			qtdeEstornosAbortados = 0;
			strMsgInformativa = "";
			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Data de corte (timeout) para aguardar os estornos pendentes ]
				dtCorte = DateTime.Today.Date.AddDays(-Global.Parametros.Geral.ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS);
				#endregion

				#region [ Aborta a verificação de estornos pendentes que excedem o prazo definido ]

				#region [ Monta SQL ]
				// A consulta seleciona os estornos pendentes que precisam ser verificados se já foram finalizados pela adquirente, mas que já excederam o prazo máximo (timeout)
				// Observações: a Cielo retorna na própria requisição se o estorno foi realizado ou não, mas a Getnet e Redecard informam
				// inicialmente apenas que a requisição foi recebida e o processamento é realizado em até D+1 ou D+2, dependendo do horário
				// em que a requisição foi realizada.
				// Importante: a consulta deve se basear nos campos de status que indicam se a transação possui um estorno pendente e se a mesma ainda não foi confirmada e nem abortada ('refund_pending_status', 
				// 'refund_pending_confirmado_status' e 'refund_pending_falha_status'), pois o campo 'ult_GlobalStatus' pode ser alterado em várias rotinas diferentes de atualização de status.
				strSql = "SELECT" +
							" tPAG_PAY.id," +
							" tPAG.pedido," +
							" tPAG.pedido_com_sufixo_nsu," +
							" tPAG_PAY.valor_transacao," +
							" tPAG_PAY.bandeira," +
							" tPAG_PAY.req_PaymentDataRequest_NumberOfPayments AS numero_parcelas," +
							" Coalesce(tCLI.nome_iniciais_em_maiusculas, '') AS nome_cliente," +
							" tPAG_PAY.refund_pending_data," +
							" tPAG_PAY.refund_pending_data_hora" +
						" FROM t_PAGTO_GW_PAG tPAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT tPAG_PAY ON (tPAG.id = tPAG_PAY.id_pagto_gw_pag)" +
							" LEFT JOIN t_CLIENTE tCLI ON (tCLI.id = tPAG.id_cliente)" +
						" WHERE" +
							" (refund_pending_status = 1)" +
							" AND (refund_pending_confirmado_status = 0)" +
							" AND (refund_pending_falha_status = 0)" +
							" AND (refund_pending_data < " + Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtCorte) + ")" +
						" ORDER BY" +
							" tPAG_PAY.refund_pending_data_hora," +
							" tPAG_PAY.id";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Aborta a verificação dos estornos pendentes que excederam o prazo (timeout) ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					qtdeAbortado++;
					row = dtbResultado.Rows[i];

					#region [ Monta mensagem p/ email de alerta ]
					strMsg = "Pedido " + BD.readToString(row["pedido"]);
					if (!BD.readToString(row["pedido"]).Equals(BD.readToString(row["pedido_com_sufixo_nsu"]))) strMsg += " (" + BD.readToString(row["pedido_com_sufixo_nsu"]) + ")";
					strMsg += ", " + Texto.iniciaisEmMaiusculas(BD.readToString(row["bandeira"])) + ", " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(BD.readToDecimal(row["valor_transacao"]));
					strMsg += " em " + BD.readToString(row["numero_parcelas"]) + "x" + " (" + BD.readToString(row["nome_cliente"]) + ")";
					strMsg += ": estorno solicitado em " + Global.formataDataDdMmYyyyHhMmComSeparador(BD.readToDateTime(row["refund_pending_data_hora"])) + " [Payment Id=" + BD.readToInt(row["id"]).ToString() + "]";
					sbAbortado.AppendLine(strMsg);
					#endregion

					#region [ Atualiza registro no BD ]
					rUpdateRefundPendingFalha = new BraspagUpdatePagPaymentRefundPendingFalha();
					rUpdateRefundPendingFalha.id_pagto_gw_pag_payment = BD.readToInt(row["id"]);
					rUpdateRefundPendingFalha.refund_pending_falha_motivo = "Timeout aguardando confirmação do estorno pendente";
					if (!BraspagDAO.updatePagPaymentRefundPendingFalha(rUpdateRefundPendingFalha, out msg_erro_aux))
					{
						msg_erro_last_op = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = "Falha ao tentar atualizar o registro da tabela t_PAGTO_GW_PAG_PAYMENT com as informações de que a verificação do estorno pendente foi abortado definitivamente por timeout\n" + msg_erro_last_op;
						svcLog.complemento_1 = Global.serializaObjectToXml(rUpdateRefundPendingFalha);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar as informações de que a verificação do estorno pendente foi abortado definitivamente por timeout [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar as informações de que a verificação do estorno pendente foi abortado definitivamente por timeout: pedido " + BD.readToString(row["pedido"]) + " (" + BD.readToString(row["pedido_com_sufixo_nsu"]) + "; t_PAGTO_GW_PAG_PAYMENT.id=" + BD.readToInt(row["id"]).ToString() + ")\r\n" + msg_erro_last_op;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
					}
					#endregion
				}
				#endregion

				#region [ Envia mensagem de alerta ]
				if (qtdeAbortado > 0)
				{
					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: estornos pendentes sem confirmação da adquirente dentro do prazo máximo" + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nEstornos pendentes que não foram confirmados pela adquirente dentro do prazo máximo e que precisam de tratamento manual.\r\n\r\n" + sbAbortado.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				#endregion

				#region [ Processa estornos pendentes ]

				#region [ Monta SQL ]
				// A consulta seleciona os estornos pendentes que precisam ser verificados se já foram finalizados pela adquirente
				// Observações: a Cielo retorna na própria requisição se o estorno foi realizado ou não, mas a Getnet e Redecard informam
				// inicialmente apenas que a requisição foi recebida e o processamento é realizado em até D+1 ou D+2, dependendo do horário
				// em que a requisição foi realizada.
				// Importante: a consulta deve se basear nos campos de status que indicam se a transação possui um estorno pendente e se a mesma ainda não foi confirmada e nem abortada ('refund_pending_status', 
				// 'refund_pending_confirmado_status' e 'refund_pending_falha_status'), pois o campo 'ult_GlobalStatus' pode ser alterado em várias rotinas diferentes de atualização de status.
				strSql = "SELECT" +
							" tPAG_PAY.id" +
						" FROM t_PAGTO_GW_PAG tPAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT tPAG_PAY ON (tPAG.id = tPAG_PAY.id_pagto_gw_pag)" +
						" WHERE" +
							" (refund_pending_status = 1)" +
							" AND (refund_pending_confirmado_status = 0)" +
							" AND (refund_pending_falha_status = 0)"+
						" ORDER BY" +
							" tPAG_PAY.refund_pending_data_hora," +
							" tPAG_PAY.id";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Processa cada estorno pendente ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					trx = BraspagDAO.getBraspagPagPaymentById(BD.readToInt(dtbResultado.Rows[i]["id"]), out msg_erro_aux);
					listaTrx.Add(trx);
				}

				for (int i = 0; i < listaTrx.Count; i++)
				{
					qtdeVerificadoTotal++;
					pag = BraspagDAO.getBraspagPagById(listaTrx[i].id_pagto_gw_pag, out msg_erro_aux);

					if (!processaTransacaoEstornoPendente(listaTrx[i], out st_estorno_confirmado, out ult_GlobalStatus_atualizado, out msg_erro_aux))
					{
						qtdeFalha++;
						strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + listaTrx[i].id.ToString() + "): " + msg_erro_aux;
						sbFalha.AppendLine(strMsg);
					}
					else
					{
						if (ult_GlobalStatus_atualizado.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNADA.GetValue()))
						{
							qtdeEstornoConfirmado++;
							strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + listaTrx[i].id.ToString() + ")";
							sbEstornoConfirmado.AppendLine(strMsg);
						}
						else if (ult_GlobalStatus_atualizado.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue()))
						{
							qtdeEstornoAindaPendente++;
							strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + listaTrx[i].id.ToString() + ")";
							sbEstornoAindaPendente.AppendLine(strMsg);
						}
						else
						{
							qtdeStatusInvalido++;
							strMsg = pag.pedido + " (pedido_com_sufixo_nsu=" + pag.pedido_com_sufixo_nsu + ", t_PAGTO_GW_PAG_PAYMENT.id=" + listaTrx[i].id.ToString() + ", Status=" + Global.Cte.Braspag.Pagador.GlobalStatus.GetDescription(ult_GlobalStatus_atualizado) + ")";
							sbStatusInvalido.AppendLine(strMsg);
						}
					}
				}
				#endregion

				#region [ Atualiza valores dos parâmetros de retorno ]
				qtdeEstornosPendentesVerificados = qtdeVerificadoTotal;
				qtdeEstornosConfirmados = qtdeEstornoConfirmado;
				qtdeEstornosAbortados = qtdeAbortado;

				strMsgInformativa = qtdeEstornosAbortados.ToString() + " estorno(s) com verificação cancelada definitivamente por exceder o prazo máximo, " +
									qtdeVerificadoTotal.ToString() + " estorno(s) pendente(s) verificado(s): " +
									qtdeEstornoConfirmado.ToString() + " estorno(s) confirmado(s), " +
									qtdeEstornoAindaPendente.ToString() + " estorno(s) continua(m) pendente(s), " +
									qtdeStatusInvalido.ToString() + " com status inválido, " +
									qtdeFalha.ToString() + " com falha na verificação" +
									"\n" +
									"Estorno confirmado:\n" + (sbEstornoConfirmado.ToString().Length > 0 ? sbEstornoConfirmado.ToString() : "(nenhum)") +
									"\n" +
									"Estorno ainda pendente:\n" + (sbEstornoAindaPendente.ToString().Length > 0 ? sbEstornoAindaPendente.ToString() : "(nenhum)") +
									"\n" +
									"Status inválido:\n" + (sbStatusInvalido.ToString().Length > 0 ? sbStatusInvalido.ToString() : "(nenhum)") +
									"\n" +
									"Falha:\n" + (sbFalha.ToString().Length > 0 ? sbFalha.ToString() : "(nenhum)");
				#endregion

				#region [ Notificação sobre as transações com status inválido ]
				if (qtdeStatusInvalido > 0)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = "Transações com status inválido ao processar estornos pendentes";
					svcLog.complemento_1 = sbStatusInvalido.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Estornos Pendentes: transações com status inválido ao processar os estornos pendentes [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nForam encontradas transações com status inválido ao processar os estornos pendentes\r\n" + sbStatusInvalido.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				#region [ Notificações sobre as transações com falha no processamento ]
				if (qtdeFalha > 0)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = "Transações com falha no processamento de verificação dos estornos pendentes";
					svcLog.complemento_1 = sbFalha.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Estornos Pendentes: transações com falha ao processar os estornos pendentes [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nTransações com falha ao processar os estornos pendentes\r\n" + sbFalha.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
			finally
			{
				if (strMsgInformativa.Length > 0)
				{
					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLogInfo = new FinSvcLog();
					svcLogInfo.operacao = NOME_DESTA_ROTINA;
					svcLogInfo.descricao = strMsgInformativa;
					GeralDAO.gravaFinSvcLog(svcLogInfo, out msg_erro_aux);
					#endregion
				}
			}
		}
		#endregion
	}
	#endregion

	#region [ BraspagPag ]
	public class BraspagPag
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private int _owner;
		public int owner
		{
			get { return _owner; }
			set { _owner = value; }
		}

		private string _usuario;
		public string usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private string _loja;
		public string loja
		{
			get { return _loja; }
			set { _loja = value; }
		}

		private string _id_cliente;
		public string id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private string _pedido;
		public string pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private string _pedido_com_sufixo_nsu;
		public string pedido_com_sufixo_nsu
		{
			get { return _pedido_com_sufixo_nsu; }
			set { _pedido_com_sufixo_nsu = value; }
		}

		private decimal _valor_pedido;
		public decimal valor_pedido
		{
			get { return _valor_pedido; }
			set { _valor_pedido = value; }
		}

		private string _operacao;
		public string operacao
		{
			get { return _operacao; }
			set { _operacao = value; }
		}

		private byte _executado_pelo_cliente_status;
		public byte executado_pelo_cliente_status
		{
			get { return _executado_pelo_cliente_status; }
			set { _executado_pelo_cliente_status = value; }
		}

		private string _origem_endereco_IP;
		public string origem_endereco_IP
		{
			get { return _origem_endereco_IP; }
			set { _origem_endereco_IP = value; }
		}

		private string _FingerPrint_SessionID;
		public string FingerPrint_SessionID
		{
			get { return _FingerPrint_SessionID; }
			set { _FingerPrint_SessionID = value; }
		}

		private DateTime _trx_TX_data;
		public DateTime trx_TX_data
		{
			get { return _trx_TX_data; }
			set { _trx_TX_data = value; }
		}

		private DateTime _trx_TX_data_hora;
		public DateTime trx_TX_data_hora
		{
			get { return _trx_TX_data_hora; }
			set { _trx_TX_data_hora = value; }
		}

		private byte _trx_RX_status;
		public byte trx_RX_status
		{
			get { return _trx_RX_status; }
			set { _trx_RX_status = value; }
		}

		private DateTime _trx_RX_data;
		public DateTime trx_RX_data
		{
			get { return _trx_RX_data; }
			set { _trx_RX_data = value; }
		}

		private DateTime _trx_RX_data_hora;
		public DateTime trx_RX_data_hora
		{
			get { return _trx_RX_data_hora; }
			set { _trx_RX_data_hora = value; }
		}

		private byte _trx_RX_vazio_status;
		public byte trx_RX_vazio_status
		{
			get { return _trx_RX_vazio_status; }
			set { _trx_RX_vazio_status = value; }
		}

		private byte _trx_erro_status;
		public byte trx_erro_status
		{
			get { return _trx_erro_status; }
			set { _trx_erro_status = value; }
		}

		private string _trx_erro_codigo;
		public string trx_erro_codigo
		{
			get { return _trx_erro_codigo; }
			set { _trx_erro_codigo = value; }
		}

		private string _trx_erro_mensagem;
		public string trx_erro_mensagem
		{
			get { return _trx_erro_mensagem; }
			set { _trx_erro_mensagem = value; }
		}

		private int _trx_TX_id_pagto_gw_pag_xml;
		public int trx_TX_id_pagto_gw_pag_xml
		{
			get { return _trx_TX_id_pagto_gw_pag_xml; }
			set { _trx_TX_id_pagto_gw_pag_xml = value; }
		}

		private int _trx_RX_id_pagto_gw_pag_xml;
		public int trx_RX_id_pagto_gw_pag_xml
		{
			get { return _trx_RX_id_pagto_gw_pag_xml; }
			set { _trx_RX_id_pagto_gw_pag_xml = value; }
		}

		private string _req_RequestId;
		public string req_RequestId
		{
			get { return _req_RequestId; }
			set { _req_RequestId = value; }
		}

		private string _req_Version;
		public string req_Version
		{
			get { return _req_Version; }
			set { _req_Version = value; }
		}

		private string _req_OrderData_MerchantId;
		public string req_OrderData_MerchantId
		{
			get { return _req_OrderData_MerchantId; }
			set { _req_OrderData_MerchantId = value; }
		}

		private string _req_OrderData_OrderId;
		public string req_OrderData_OrderId
		{
			get { return _req_OrderData_OrderId; }
			set { _req_OrderData_OrderId = value; }
		}

		private string _req_CustomerData_CustomerIdentity;
		public string req_CustomerData_CustomerIdentity
		{
			get { return _req_CustomerData_CustomerIdentity; }
			set { _req_CustomerData_CustomerIdentity = value; }
		}

		private string _req_CustomerData_CustomerName;
		public string req_CustomerData_CustomerName
		{
			get { return _req_CustomerData_CustomerName; }
			set { _req_CustomerData_CustomerName = value; }
		}

		private string _resp_CorrelationId;
		public string resp_CorrelationId
		{
			get { return _resp_CorrelationId; }
			set { _resp_CorrelationId = value; }
		}

		private string _resp_Success;
		public string resp_Success
		{
			get { return _resp_Success; }
			set { _resp_Success = value; }
		}

		private string _resp_OrderData_OrderId;
		public string resp_OrderData_OrderId
		{
			get { return _resp_OrderData_OrderId; }
			set { _resp_OrderData_OrderId = value; }
		}

		private string _resp_OrderData_BraspagOrderId;
		public string resp_OrderData_BraspagOrderId
		{
			get { return _resp_OrderData_BraspagOrderId; }
			set { _resp_OrderData_BraspagOrderId = value; }
		}

		private string _recibo_url_css;
		public string recibo_url_css
		{
			get { return _recibo_url_css; }
			set { _recibo_url_css = value; }
		}

		private string _recibo_html;
		public string recibo_html
		{
			get { return _recibo_html; }
			set { _recibo_html = value; }
		}

		private string _msg_alerta_tela;
		public string msg_alerta_tela
		{
			get { return _msg_alerta_tela; }
			set { _msg_alerta_tela = value; }
		}

		private string _SessionCtrlInfo;
		public string SessionCtrlInfo
		{
			get { return _SessionCtrlInfo; }
			set { _SessionCtrlInfo = value; }
		}
	}
	#endregion

	#region [ BraspagPagPayment ]
	public class BraspagPagPayment
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_pag;
		public int id_pagto_gw_pag
		{
			get { return _id_pagto_gw_pag; }
			set { _id_pagto_gw_pag = value; }
		}

		private int _ordem;
		public int ordem
		{
			get { return _ordem; }
			set { _ordem = value; }
		}

		private byte _st_enviado_analise_AF;
		public byte st_enviado_analise_AF
		{
			get { return _st_enviado_analise_AF; }
			set { _st_enviado_analise_AF = value; }
		}

		private int _id_pagto_gw_af;
		public int id_pagto_gw_af
		{
			get { return _id_pagto_gw_af; }
			set { _id_pagto_gw_af = value; }
		}

		private string _bandeira;
		public string bandeira
		{
			get { return _bandeira; }
			set { _bandeira = value; }
		}

		private decimal _valor_transacao;
		public decimal valor_transacao
		{
			get { return _valor_transacao; }
			set { _valor_transacao = value; }
		}

		private string _checkout_opcao_parcelamento;
		public string checkout_opcao_parcelamento
		{
			get { return _checkout_opcao_parcelamento; }
			set { _checkout_opcao_parcelamento = value; }
		}

		private string _checkout_titular_nome;
		public string checkout_titular_nome
		{
			get { return _checkout_titular_nome; }
			set { _checkout_titular_nome = value; }
		}

		private string _checkout_titular_cpf_cnpj;
		public string checkout_titular_cpf_cnpj
		{
			get { return _checkout_titular_cpf_cnpj; }
			set { _checkout_titular_cpf_cnpj = value; }
		}

		private string _checkout_cartao_numero;
		public string checkout_cartao_numero
		{
			get { return _checkout_cartao_numero; }
			set { _checkout_cartao_numero = value; }
		}

		private string _checkout_cartao_validade_mes;
		public string checkout_cartao_validade_mes
		{
			get { return _checkout_cartao_validade_mes; }
			set { _checkout_cartao_validade_mes = value; }
		}

		private string _checkout_cartao_validade_ano;
		public string checkout_cartao_validade_ano
		{
			get { return _checkout_cartao_validade_ano; }
			set { _checkout_cartao_validade_ano = value; }
		}

		private string _checkout_cartao_codigo_seguranca;
		public string checkout_cartao_codigo_seguranca
		{
			get { return _checkout_cartao_codigo_seguranca; }
			set { _checkout_cartao_codigo_seguranca = value; }
		}

		private string _checkout_cartao_proprio;
		public string checkout_cartao_proprio
		{
			get { return _checkout_cartao_proprio; }
			set { _checkout_cartao_proprio = value; }
		}

		private string _checkout_fatura_end_logradouro;
		public string checkout_fatura_end_logradouro
		{
			get { return _checkout_fatura_end_logradouro; }
			set { _checkout_fatura_end_logradouro = value; }
		}

		private string _checkout_fatura_end_numero;
		public string checkout_fatura_end_numero
		{
			get { return _checkout_fatura_end_numero; }
			set { _checkout_fatura_end_numero = value; }
		}

		private string _checkout_fatura_end_complemento;
		public string checkout_fatura_end_complemento
		{
			get { return _checkout_fatura_end_complemento; }
			set { _checkout_fatura_end_complemento = value; }
		}

		private string _checkout_fatura_end_bairro;
		public string checkout_fatura_end_bairro
		{
			get { return _checkout_fatura_end_bairro; }
			set { _checkout_fatura_end_bairro = value; }
		}

		private string _checkout_fatura_end_cidade;
		public string checkout_fatura_end_cidade
		{
			get { return _checkout_fatura_end_cidade; }
			set { _checkout_fatura_end_cidade = value; }
		}

		private string _checkout_fatura_end_uf;
		public string checkout_fatura_end_uf
		{
			get { return _checkout_fatura_end_uf; }
			set { _checkout_fatura_end_uf = value; }
		}

		private string _checkout_fatura_end_cep;
		public string checkout_fatura_end_cep
		{
			get { return _checkout_fatura_end_cep; }
			set { _checkout_fatura_end_cep = value; }
		}

		private string _checkout_fatura_tel_pais;
		public string checkout_fatura_tel_pais
		{
			get { return _checkout_fatura_tel_pais; }
			set { _checkout_fatura_tel_pais = value; }
		}

		private string _checkout_fatura_tel_ddd;
		public string checkout_fatura_tel_ddd
		{
			get { return _checkout_fatura_tel_ddd; }
			set { _checkout_fatura_tel_ddd = value; }
		}

		private string _checkout_fatura_tel_numero;
		public string checkout_fatura_tel_numero
		{
			get { return _checkout_fatura_tel_numero; }
			set { _checkout_fatura_tel_numero = value; }
		}

		private string _checkout_email;
		public string checkout_email
		{
			get { return _checkout_email; }
			set { _checkout_email = value; }
		}

		private string _prim_GlobalStatus;
		public string prim_GlobalStatus
		{
			get { return _prim_GlobalStatus; }
			set { _prim_GlobalStatus = value; }
		}

		private DateTime _prim_atualizacao_data_hora;
		public DateTime prim_atualizacao_data_hora
		{
			get { return _prim_atualizacao_data_hora; }
			set { _prim_atualizacao_data_hora = value; }
		}

		private string _prim_atualizacao_usuario;
		public string prim_atualizacao_usuario
		{
			get { return _prim_atualizacao_usuario; }
			set { _prim_atualizacao_usuario = value; }
		}

		private string _ult_GlobalStatus;
		public string ult_GlobalStatus
		{
			get { return _ult_GlobalStatus; }
			set { _ult_GlobalStatus = value; }
		}

		private DateTime _ult_atualizacao_data_hora;
		public DateTime ult_atualizacao_data_hora
		{
			get { return _ult_atualizacao_data_hora; }
			set { _ult_atualizacao_data_hora = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}

		private DateTime _resp_AuthorizedDate;
		public DateTime resp_AuthorizedDate
		{
			get { return _resp_AuthorizedDate; }
			set { _resp_AuthorizedDate = value; }
		}

		private DateTime _resp_CapturedDate;
		public DateTime resp_CapturedDate
		{
			get { return _resp_CapturedDate; }
			set { _resp_CapturedDate = value; }
		}

		private DateTime _resp_VoidedDate;
		public DateTime resp_VoidedDate
		{
			get { return _resp_VoidedDate; }
			set { _resp_VoidedDate = value; }
		}

		private byte _tratado_manual_status;
		public byte tratado_manual_status
		{
			get { return _tratado_manual_status; }
			set { _tratado_manual_status = value; }
		}

		private string _tratado_manual_usuario;
		public string tratado_manual_usuario
		{
			get { return _tratado_manual_usuario; }
			set { _tratado_manual_usuario = value; }
		}

		private DateTime _tratado_manual_data;
		public DateTime tratado_manual_data
		{
			get { return _tratado_manual_data; }
			set { _tratado_manual_data = value; }
		}

		private DateTime _tratado_manual_data_hora;
		public DateTime tratado_manual_data_hora
		{
			get { return _tratado_manual_data_hora; }
			set { _tratado_manual_data_hora = value; }
		}

		private string _tratado_manual_obs;
		public string tratado_manual_obs
		{
			get { return _tratado_manual_obs; }
			set { _tratado_manual_obs = value; }
		}

		private string _tratado_manual_ult_atualiz_usuario;
		public string tratado_manual_ult_atualiz_usuario
		{
			get { return _tratado_manual_ult_atualiz_usuario; }
			set { _tratado_manual_ult_atualiz_usuario = value; }
		}

		private DateTime _tratado_manual_ult_atualiz_data;
		public DateTime tratado_manual_ult_atualiz_data
		{
			get { return _tratado_manual_ult_atualiz_data; }
			set { _tratado_manual_ult_atualiz_data = value; }
		}

		private DateTime _tratado_manual_ult_atualiz_data_hora;
		public DateTime tratado_manual_ult_atualiz_data_hora
		{
			get { return _tratado_manual_ult_atualiz_data_hora; }
			set { _tratado_manual_ult_atualiz_data_hora = value; }
		}

		private string _req_PaymentDataRequest_PaymentMethod;
		public string req_PaymentDataRequest_PaymentMethod
		{
			get { return _req_PaymentDataRequest_PaymentMethod; }
			set { _req_PaymentDataRequest_PaymentMethod = value; }
		}

		private string _req_PaymentDataRequest_Amount;
		public string req_PaymentDataRequest_Amount
		{
			get { return _req_PaymentDataRequest_Amount; }
			set { _req_PaymentDataRequest_Amount = value; }
		}

		private string _req_PaymentDataRequest_Currency;
		public string req_PaymentDataRequest_Currency
		{
			get { return _req_PaymentDataRequest_Currency; }
			set { _req_PaymentDataRequest_Currency = value; }
		}

		private string _req_PaymentDataRequest_Country;
		public string req_PaymentDataRequest_Country
		{
			get { return _req_PaymentDataRequest_Country; }
			set { _req_PaymentDataRequest_Country = value; }
		}

		private string _req_PaymentDataRequest_ServiceTaxAmount;
		public string req_PaymentDataRequest_ServiceTaxAmount
		{
			get { return _req_PaymentDataRequest_ServiceTaxAmount; }
			set { _req_PaymentDataRequest_ServiceTaxAmount = value; }
		}

		private string _req_PaymentDataRequest_NumberOfPayments;
		public string req_PaymentDataRequest_NumberOfPayments
		{
			get { return _req_PaymentDataRequest_NumberOfPayments; }
			set { _req_PaymentDataRequest_NumberOfPayments = value; }
		}

		private string _req_PaymentDataRequest_PaymentPlan;
		public string req_PaymentDataRequest_PaymentPlan
		{
			get { return _req_PaymentDataRequest_PaymentPlan; }
			set { _req_PaymentDataRequest_PaymentPlan = value; }
		}

		private string _req_PaymentDataRequest_TransactionType;
		public string req_PaymentDataRequest_TransactionType
		{
			get { return _req_PaymentDataRequest_TransactionType; }
			set { _req_PaymentDataRequest_TransactionType = value; }
		}

		private string _req_PaymentDataRequest_CardHolder;
		public string req_PaymentDataRequest_CardHolder
		{
			get { return _req_PaymentDataRequest_CardHolder; }
			set { _req_PaymentDataRequest_CardHolder = value; }
		}

		private string _req_PaymentDataRequest_CardNumber;
		public string req_PaymentDataRequest_CardNumber
		{
			get { return _req_PaymentDataRequest_CardNumber; }
			set { _req_PaymentDataRequest_CardNumber = value; }
		}

		private string _req_PaymentDataRequest_CardSecurityCode;
		public string req_PaymentDataRequest_CardSecurityCode
		{
			get { return _req_PaymentDataRequest_CardSecurityCode; }
			set { _req_PaymentDataRequest_CardSecurityCode = value; }
		}

		private string _req_PaymentDataRequest_CardExpirationDate;
		public string req_PaymentDataRequest_CardExpirationDate
		{
			get { return _req_PaymentDataRequest_CardExpirationDate; }
			set { _req_PaymentDataRequest_CardExpirationDate = value; }
		}

		private string _resp_PaymentDataResponse_BraspagTransactionId;
		public string resp_PaymentDataResponse_BraspagTransactionId
		{
			get { return _resp_PaymentDataResponse_BraspagTransactionId; }
			set { _resp_PaymentDataResponse_BraspagTransactionId = value; }
		}

		private string _resp_PaymentDataResponse_PaymentMethod;
		public string resp_PaymentDataResponse_PaymentMethod
		{
			get { return _resp_PaymentDataResponse_PaymentMethod; }
			set { _resp_PaymentDataResponse_PaymentMethod = value; }
		}

		private string _resp_PaymentDataResponse_Amount;
		public string resp_PaymentDataResponse_Amount
		{
			get { return _resp_PaymentDataResponse_Amount; }
			set { _resp_PaymentDataResponse_Amount = value; }
		}

		private string _resp_PaymentDataResponse_AcquirerTransactionId;
		public string resp_PaymentDataResponse_AcquirerTransactionId
		{
			get { return _resp_PaymentDataResponse_AcquirerTransactionId; }
			set { _resp_PaymentDataResponse_AcquirerTransactionId = value; }
		}

		private string _resp_PaymentDataResponse_AuthorizationCode;
		public string resp_PaymentDataResponse_AuthorizationCode
		{
			get { return _resp_PaymentDataResponse_AuthorizationCode; }
			set { _resp_PaymentDataResponse_AuthorizationCode = value; }
		}

		private string _resp_PaymentDataResponse_CreditCardToken;
		public string resp_PaymentDataResponse_CreditCardToken
		{
			get { return _resp_PaymentDataResponse_CreditCardToken; }
			set { _resp_PaymentDataResponse_CreditCardToken = value; }
		}

		private string _resp_PaymentDataResponse_ProofOfSale;
		public string resp_PaymentDataResponse_ProofOfSale
		{
			get { return _resp_PaymentDataResponse_ProofOfSale; }
			set { _resp_PaymentDataResponse_ProofOfSale = value; }
		}

		private string _resp_PaymentDataResponse_ReturnCode;
		public string resp_PaymentDataResponse_ReturnCode
		{
			get { return _resp_PaymentDataResponse_ReturnCode; }
			set { _resp_PaymentDataResponse_ReturnCode = value; }
		}

		private string _resp_PaymentDataResponse_ReturnMessage;
		public string resp_PaymentDataResponse_ReturnMessage
		{
			get { return _resp_PaymentDataResponse_ReturnMessage; }
			set { _resp_PaymentDataResponse_ReturnMessage = value; }
		}

		private string _resp_PaymentDataResponse_Status;
		public string resp_PaymentDataResponse_Status
		{
			get { return _resp_PaymentDataResponse_Status; }
			set { _resp_PaymentDataResponse_Status = value; }
		}

		private byte _captura_confirmada_status;
		public byte captura_confirmada_status
		{
			get { return _captura_confirmada_status; }
			set { _captura_confirmada_status = value; }
		}

		private DateTime _captura_confirmada_data;
		public DateTime captura_confirmada_data
		{
			get { return _captura_confirmada_data; }
			set { _captura_confirmada_data = value; }
		}

		private DateTime _captura_confirmada_data_hora;
		public DateTime captura_confirmada_data_hora
		{
			get { return _captura_confirmada_data_hora; }
			set { _captura_confirmada_data_hora = value; }
		}

		private string _captura_confirmada_usuario;
		public string captura_confirmada_usuario
		{
			get { return _captura_confirmada_usuario; }
			set { _captura_confirmada_usuario = value; }
		}

		private byte _voided_status;
		public byte voided_status
		{
			get { return _voided_status; }
			set { _voided_status = value; }
		}

		private DateTime _voided_data;
		public DateTime voided_data
		{
			get { return _voided_data; }
			set { _voided_data = value; }
		}

		private DateTime _voided_data_hora;
		public DateTime voided_data_hora
		{
			get { return _voided_data_hora; }
			set { _voided_data_hora = value; }
		}

		private string _voided_usuario;
		public string voided_usuario
		{
			get { return _voided_usuario; }
			set { _voided_usuario = value; }
		}

		private byte _refunded_status;
		public byte refunded_status
		{
			get { return _refunded_status; }
			set { _refunded_status = value; }
		}

		private DateTime _refunded_data;
		public DateTime refunded_data
		{
			get { return _refunded_data; }
			set { _refunded_data = value; }
		}

		private DateTime _refunded_data_hora;
		public DateTime refunded_data_hora
		{
			get { return _refunded_data_hora; }
			set { _refunded_data_hora = value; }
		}

		private string _refunded_usuario;
		public string refunded_usuario
		{
			get { return _refunded_usuario; }
			set { _refunded_usuario = value; }
		}

		private byte _refund_pending_status;
		public byte refund_pending_status
		{
			get { return _refund_pending_status; }
			set { _refund_pending_status = value; }
		}

		private DateTime _refund_pending_data;
		public DateTime refund_pending_data
		{
			get { return _refund_pending_data; }
			set { _refund_pending_data = value; }
		}

		private DateTime _refund_pending_data_hora;
		public DateTime refund_pending_data_hora
		{
			get { return _refund_pending_data_hora; }
			set { _refund_pending_data_hora = value; }
		}

		private string _refund_pending_usuario;
		public string refund_pending_usuario
		{
			get { return _refund_pending_usuario; }
			set { _refund_pending_usuario = value; }
		}

		private byte _refund_pending_confirmado_status;
		public byte refund_pending_confirmado_status
		{
			get { return _refund_pending_confirmado_status; }
			set { _refund_pending_confirmado_status = value; }
		}

		private DateTime _refund_pending_confirmado_data;
		public DateTime refund_pending_confirmado_data
		{
			get { return _refund_pending_confirmado_data; }
			set { _refund_pending_confirmado_data = value; }
		}

		private DateTime _refund_pending_confirmado_data_hora;
		public DateTime refund_pending_confirmado_data_hora
		{
			get { return _refund_pending_confirmado_data_hora; }
			set { _refund_pending_confirmado_data_hora = value; }
		}

		private string _refund_pending_confirmado_usuario;
		public string refund_pending_confirmado_usuario
		{
			get { return _refund_pending_confirmado_usuario; }
			set { _refund_pending_confirmado_usuario = value; }
		}

		private byte _refund_pending_falha_status;
		public byte refund_pending_falha_status
		{
			get { return _refund_pending_falha_status; }
			set { _refund_pending_falha_status = value; }
		}

		private DateTime _refund_pending_falha_data;
		public DateTime refund_pending_falha_data
		{
			get { return _refund_pending_falha_data; }
			set { _refund_pending_falha_data = value; }
		}

		private DateTime _refund_pending_falha_data_hora;
		public DateTime refund_pending_falha_data_hora
		{
			get { return _refund_pending_falha_data_hora; }
			set { _refund_pending_falha_data_hora = value; }
		}

		private string _refund_pending_falha_motivo;
		public string refund_pending_falha_motivo
		{
			get { return _refund_pending_falha_motivo; }
			set { _refund_pending_falha_motivo = value; }
		}

		private byte _captura_confirmada_erro_status;
		public byte captura_confirmada_erro_status
		{
			get { return _captura_confirmada_erro_status; }
			set { _captura_confirmada_erro_status = value; }
		}

		private DateTime _captura_confirmada_erro_data;
		public DateTime captura_confirmada_erro_data
		{
			get { return _captura_confirmada_erro_data; }
			set { _captura_confirmada_erro_data = value; }
		}

		private DateTime _captura_confirmada_erro_data_hora;
		public DateTime captura_confirmada_erro_data_hora
		{
			get { return _captura_confirmada_erro_data_hora; }
			set { _captura_confirmada_erro_data_hora = value; }
		}

		private string _captura_confirmada_erro_mensagem;
		public string captura_confirmada_erro_mensagem
		{
			get { return _captura_confirmada_erro_mensagem; }
			set { _captura_confirmada_erro_mensagem = value; }
		}

		private byte _voided_erro_status;
		public byte voided_erro_status
		{
			get { return _voided_erro_status; }
			set { _voided_erro_status = value; }
		}

		private DateTime _voided_erro_data;
		public DateTime voided_erro_data
		{
			get { return _voided_erro_data; }
			set { _voided_erro_data = value; }
		}

		private DateTime _voided_erro_data_hora;
		public DateTime voided_erro_data_hora
		{
			get { return _voided_erro_data_hora; }
			set { _voided_erro_data_hora = value; }
		}

		private string _voided_erro_mensagem;
		public string voided_erro_mensagem
		{
			get { return _voided_erro_mensagem; }
			set { _voided_erro_mensagem = value; }
		}

		private byte _refunded_erro_status;
		public byte refunded_erro_status
		{
			get { return _refunded_erro_status; }
			set { _refunded_erro_status = value; }
		}

		private DateTime _refunded_erro_data;
		public DateTime refunded_erro_data
		{
			get { return _refunded_erro_data; }
			set { _refunded_erro_data = value; }
		}

		private DateTime _refunded_erro_data_hora;
		public DateTime refunded_erro_data_hora
		{
			get { return _refunded_erro_data_hora; }
			set { _refunded_erro_data_hora = value; }
		}

		private string _refunded_erro_mensagem;
		public string refunded_erro_mensagem
		{
			get { return _refunded_erro_mensagem; }
			set { _refunded_erro_mensagem = value; }
		}

		private byte _pedido_hist_pagto_gravado_status;
		public byte pedido_hist_pagto_gravado_status
		{
			get { return _pedido_hist_pagto_gravado_status; }
			set { _pedido_hist_pagto_gravado_status = value; }
		}

		private DateTime _pedido_hist_pagto_gravado_data;
		public DateTime pedido_hist_pagto_gravado_data
		{
			get { return _pedido_hist_pagto_gravado_data; }
			set { _pedido_hist_pagto_gravado_data = value; }
		}

		private DateTime _pedido_hist_pagto_gravado_data_hora;
		public DateTime pedido_hist_pagto_gravado_data_hora
		{
			get { return _pedido_hist_pagto_gravado_data_hora; }
			set { _pedido_hist_pagto_gravado_data_hora = value; }
		}

		private byte _pagto_registrado_no_pedido_status;
		public byte pagto_registrado_no_pedido_status
		{
			get { return _pagto_registrado_no_pedido_status; }
			set { _pagto_registrado_no_pedido_status = value; }
		}

		private string _pagto_registrado_no_pedido_tipo_operacao;
		public string pagto_registrado_no_pedido_tipo_operacao
		{
			get { return _pagto_registrado_no_pedido_tipo_operacao; }
			set { _pagto_registrado_no_pedido_tipo_operacao = value; }
		}

		private DateTime _pagto_registrado_no_pedido_data;
		public DateTime pagto_registrado_no_pedido_data
		{
			get { return _pagto_registrado_no_pedido_data; }
			set { _pagto_registrado_no_pedido_data = value; }
		}

		private DateTime _pagto_registrado_no_pedido_data_hora;
		public DateTime pagto_registrado_no_pedido_data_hora
		{
			get { return _pagto_registrado_no_pedido_data_hora; }
			set { _pagto_registrado_no_pedido_data_hora = value; }
		}

		private string _pagto_registrado_no_pedido_usuario;
		public string pagto_registrado_no_pedido_usuario
		{
			get { return _pagto_registrado_no_pedido_usuario; }
			set { _pagto_registrado_no_pedido_usuario = value; }
		}

		private string _pagto_registrado_no_pedido_id_pedido_pagamento;
		public string pagto_registrado_no_pedido_id_pedido_pagamento
		{
			get { return _pagto_registrado_no_pedido_id_pedido_pagamento; }
			set { _pagto_registrado_no_pedido_id_pedido_pagamento = value; }
		}

		private string _pagto_registrado_no_pedido_st_pagto_anterior;
		public string pagto_registrado_no_pedido_st_pagto_anterior
		{
			get { return _pagto_registrado_no_pedido_st_pagto_anterior; }
			set { _pagto_registrado_no_pedido_st_pagto_anterior = value; }
		}

		private string _pagto_registrado_no_pedido_st_pagto_novo;
		public string pagto_registrado_no_pedido_st_pagto_novo
		{
			get { return _pagto_registrado_no_pedido_st_pagto_novo; }
			set { _pagto_registrado_no_pedido_st_pagto_novo = value; }
		}

		private byte _st_cancelado_envio_analise_AF;
		public byte st_cancelado_envio_analise_AF
		{
			get { return _st_cancelado_envio_analise_AF; }
			set { _st_cancelado_envio_analise_AF = value; }
		}

		private byte _st_processamento_AF_finalizado;
		public byte st_processamento_AF_finalizado
		{
			get { return _st_processamento_AF_finalizado; }
			set { _st_processamento_AF_finalizado = value; }
		}

		private DateTime _dt_hr_processamento_AF_finalizado;
		public DateTime dt_hr_processamento_AF_finalizado
		{
			get { return _dt_hr_processamento_AF_finalizado; }
			set { _dt_hr_processamento_AF_finalizado = value; }
		}

		private byte _st_processamento_PAG_finalizado;
		public byte st_processamento_PAG_finalizado
		{
			get { return _st_processamento_PAG_finalizado; }
			set { _st_processamento_PAG_finalizado = value; }
		}

		private DateTime _dt_hr_processamento_PAG_finalizado;
		public DateTime dt_hr_processamento_PAG_finalizado
		{
			get { return _dt_hr_processamento_PAG_finalizado; }
			set { _dt_hr_processamento_PAG_finalizado = value; }
		}
	}
	#endregion

	#region [ BraspagPagPaymentFinalizacao ]
	public class BraspagPagPaymentFinalizacao
	{
		public Global.Cte.Braspag.Pagador.OperacaoFinalizacao operacao;
		public BraspagPagPayment payment;

		public BraspagPagPaymentFinalizacao(Global.Cte.Braspag.Pagador.OperacaoFinalizacao operacao, BraspagPagPayment payment)
		{
			this.operacao = operacao;
			this.payment = payment;
		}
	}
	#endregion

	#region [ BraspagErrorReportDataResponse ]
	public class BraspagErrorReportDataResponse
	{
		private string _ErrorCode;
		public string ErrorCode
		{
			get { return _ErrorCode; }
			set { _ErrorCode = value; }
		}

		private string _ErrorMessage;
		public string ErrorMessage
		{
			get { return _ErrorMessage; }
			set { _ErrorMessage = value; }
		}
	}
	#endregion

	#region [ GetOrderIdData ]

	#region [ BraspagGetOrderIdData ]
	public class BraspagGetOrderIdData
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		private string _OrderId;
		public string OrderId
		{
			get { return _OrderId; }
			set { _OrderId = value; }
		}
	}
	#endregion

	#region [ BraspagOrderIdTransactionResponse ]
	public class BraspagOrderIdTransactionResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		private string _BraspagOrderId;
		public string BraspagOrderId
		{
			get { return _BraspagOrderId; }
			set { _BraspagOrderId = value; }
		}

		public List<String> BraspagTransactionId = new List<String>();
		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#region [ BraspagGetOrderIdDataResponse ]
	public class BraspagGetOrderIdDataResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		public List<BraspagOrderIdTransactionResponse> OrderIdDataCollection = new List<BraspagOrderIdTransactionResponse>();
		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#endregion

	#region [ GetOrderData ]

	#region [ BraspagGetOrderData ]
	public class BraspagGetOrderData
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		private string _BraspagOrderId;
		public string BraspagOrderId
		{
			get { return _BraspagOrderId; }
			set { _BraspagOrderId = value; }
		}
	}
	#endregion

	#region [ BraspagOrderTransactionDataResponse ]
	public class BraspagOrderTransactionDataResponse
	{
		private string _BraspagTransactionId;
		public string BraspagTransactionId
		{
			get { return _BraspagTransactionId; }
			set { _BraspagTransactionId = value; }
		}

		private string _OrderId;
		public string OrderId
		{
			get { return _OrderId; }
			set { _OrderId = value; }
		}

		private string _AcquirerTransactionId;
		public string AcquirerTransactionId
		{
			get { return _AcquirerTransactionId; }
			set { _AcquirerTransactionId = value; }
		}

		private string _PaymentMethod;
		public string PaymentMethod
		{
			get { return _PaymentMethod; }
			set { _PaymentMethod = value; }
		}

		private string _PaymentMethodName;
		public string PaymentMethodName
		{
			get { return _PaymentMethodName; }
			set { _PaymentMethodName = value; }
		}

		private string _ErrorCode;
		public string ErrorCode
		{
			get { return _ErrorCode; }
			set { _ErrorCode = value; }
		}

		private string _ErrorMessage;
		public string ErrorMessage
		{
			get { return _ErrorMessage; }
			set { _ErrorMessage = value; }
		}

		private string _Amount;
		public string Amount
		{
			get { return _Amount; }
			set { _Amount = value; }
		}

		private string _AuthorizationCode;
		public string AuthorizationCode
		{
			get { return _AuthorizationCode; }
			set { _AuthorizationCode = value; }
		}

		private string _NumberOfPayments;
		public string NumberOfPayments
		{
			get { return _NumberOfPayments; }
			set { _NumberOfPayments = value; }
		}

		private string _Currency;
		public string Currency
		{
			get { return _Currency; }
			set { _Currency = value; }
		}

		private string _Country;
		public string Country
		{
			get { return _Country; }
			set { _Country = value; }
		}

		private string _TransactionType;
		public string TransactionType
		{
			get { return _TransactionType; }
			set { _TransactionType = value; }
		}

		private string _Status;
		public string Status
		{
			get { return _Status; }
			set { _Status = value; }
		}

		private string _ReceivedDate;
		public string ReceivedDate
		{
			get { return _ReceivedDate; }
			set { _ReceivedDate = value; }
		}

		private string _CapturedDate;
		public string CapturedDate
		{
			get { return _CapturedDate; }
			set { _CapturedDate = value; }
		}

		private string _VoidedDate;
		public string VoidedDate
		{
			get { return _VoidedDate; }
			set { _VoidedDate = value; }
		}

		private string _CreditCardToken;
		public string CreditCardToken
		{
			get { return _CreditCardToken; }
			set { _CreditCardToken = value; }
		}

		private string _ProofOfSale;
		public string ProofOfSale
		{
			get { return _ProofOfSale; }
			set { _ProofOfSale = value; }
		}

		private string _MaskedCardNumber;
		public string MaskedCardNumber
		{
			get { return _MaskedCardNumber; }
			set { _MaskedCardNumber = value; }
		}
	}
	#endregion

	#region [ BraspagGetOrderDataResponse ]
	public class BraspagGetOrderDataResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		public List<BraspagOrderTransactionDataResponse> TransactionDataCollection = new List<BraspagOrderTransactionDataResponse>();
		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#endregion

	#region [ GetTransactionData ]

	#region [ BraspagGetTransactionData ]
	public class BraspagGetTransactionData
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		private string _BraspagTransactionId;
		public string BraspagTransactionId
		{
			get { return _BraspagTransactionId; }
			set { _BraspagTransactionId = value; }
		}
	}
	#endregion

	#region [ BraspagGetTransactionDataResponse ]
	public class BraspagGetTransactionDataResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		private string _BraspagTransactionId;
		public string BraspagTransactionId
		{
			get { return _BraspagTransactionId; }
			set { _BraspagTransactionId = value; }
		}

		private string _OrderId;
		public string OrderId
		{
			get { return _OrderId; }
			set { _OrderId = value; }
		}

		private string _AcquirerTransactionId;
		public string AcquirerTransactionId
		{
			get { return _AcquirerTransactionId; }
			set { _AcquirerTransactionId = value; }
		}

		private string _PaymentMethod;
		public string PaymentMethod
		{
			get { return _PaymentMethod; }
			set { _PaymentMethod = value; }
		}

		private string _PaymentMethodName;
		public string PaymentMethodName
		{
			get { return _PaymentMethodName; }
			set { _PaymentMethodName = value; }
		}

		private string _Amount;
		public string Amount
		{
			get { return _Amount; }
			set { _Amount = value; }
		}

		private string _AuthorizationCode;
		public string AuthorizationCode
		{
			get { return _AuthorizationCode; }
			set { _AuthorizationCode = value; }
		}

		private string _NumberOfPayments;
		public string NumberOfPayments
		{
			get { return _NumberOfPayments; }
			set { _NumberOfPayments = value; }
		}

		private string _Currency;
		public string Currency
		{
			get { return _Currency; }
			set { _Currency = value; }
		}

		private string _Country;
		public string Country
		{
			get { return _Country; }
			set { _Country = value; }
		}

		private string _TransactionType;
		public string TransactionType
		{
			get { return _TransactionType; }
			set { _TransactionType = value; }
		}

		private string _Status;
		public string Status
		{
			get { return _Status; }
			set { _Status = value; }
		}

		private string _ReceivedDate;
		public string ReceivedDate
		{
			get { return _ReceivedDate; }
			set { _ReceivedDate = value; }
		}

		private string _CapturedDate;
		public string CapturedDate
		{
			get { return _CapturedDate; }
			set { _CapturedDate = value; }
		}

		private string _VoidedDate;
		public string VoidedDate
		{
			get { return _VoidedDate; }
			set { _VoidedDate = value; }
		}

		private string _CreditCardToken;
		public string CreditCardToken
		{
			get { return _CreditCardToken; }
			set { _CreditCardToken = value; }
		}

		private string _ProofOfSale;
		public string ProofOfSale
		{
			get { return _ProofOfSale; }
			set { _ProofOfSale = value; }
		}

		private string _MaskedCardNumber;
		public string MaskedCardNumber
		{
			get { return _MaskedCardNumber; }
			set { _MaskedCardNumber = value; }
		}

		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();

		private string _faultcode;
		public string faultcode
		{
			get { return _faultcode; }
			set { _faultcode = value; }
		}

		private string _faultstring;
		public string faultstring
		{
			get { return _faultstring; }
			set { _faultstring = value; }
		}
	}
	#endregion

	#endregion

	#region [ GetBoletoData ]

	#region [ BraspagGetBoletoData ]
	public class BraspagGetBoletoData
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		private string _BraspagTransactionId;
		public string BraspagTransactionId
		{
			get { return _BraspagTransactionId; }
			set { _BraspagTransactionId = value; }
		}
	}
	#endregion

	#region [ BraspagGetBoletoDataResponse ]
	public class BraspagGetBoletoDataResponse
	{
		public string CorrelationId { get; set; }
		public string Success { get; set; }
		public string BraspagTransactionId { get; set; }
		public string PaymentMethod { get; set; }
		public string DocumentNumber { get; set; }
		public string DocumentDate { get; set; }
		public string CustomerName { get; set; }
		public string BoletoNumber { get; set; }
		public string BarCodeNumber { get; set; }
		public string BoletoExpirationDate { get; set; }
		public string BoletoInstructions { get; set; }
		public string BoletoType { get; set; }
		public string BoletoUrl { get; set; }
		public string Amount { get; set; }
		public string PaidAmount { get; set; }
		public string PaymentDate { get; set; }
		public string BankNumber { get; set; }
		public string Agency { get; set; }
		public string Account { get; set; }
		public string Assignor { get; set; }

		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#endregion

	#region [ CaptureCreditCardTransaction ]

	#region [ BraspagTransactionDataRequest ]
	public class BraspagTransactionDataRequest
	{
		public BraspagTransactionDataRequest() : this("", "", "") { }

		public BraspagTransactionDataRequest(string BraspagTransactionId, string Amount, string ServiceTaxAmount)
		{
			this._BraspagTransactionId = BraspagTransactionId;
			this._Amount = Amount;
			this._ServiceTaxAmount = ServiceTaxAmount;
		}

		private string _BraspagTransactionId;
		public string BraspagTransactionId
		{
			get { return _BraspagTransactionId; }
			set { _BraspagTransactionId = value; }
		}

		private string _Amount;
		public string Amount
		{
			get { return _Amount; }
			set { _Amount = value; }
		}

		private string _ServiceTaxAmount;
		public string ServiceTaxAmount
		{
			get { return _ServiceTaxAmount; }
			set { _ServiceTaxAmount = value; }
		}
	}
	#endregion

	#region [ BraspagCaptureCreditCardTransaction ]
	public class BraspagCaptureCreditCardTransaction
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		public List<BraspagTransactionDataRequest> TransactionDataCollection = new List<BraspagTransactionDataRequest>();
	}
	#endregion

	#region [ BraspagTransactionDataResponse ]
	public class BraspagTransactionDataResponse
	{
		private string _BraspagTransactionId;
		public string BraspagTransactionId
		{
			get { return _BraspagTransactionId; }
			set { _BraspagTransactionId = value; }
		}

		private string _AcquirerTransactionId;
		public string AcquirerTransactionId
		{
			get { return _AcquirerTransactionId; }
			set { _AcquirerTransactionId = value; }
		}

		private string _Amount;
		public string Amount
		{
			get { return _Amount; }
			set { _Amount = value; }
		}

		private string _AuthorizationCode;
		public string AuthorizationCode
		{
			get { return _AuthorizationCode; }
			set { _AuthorizationCode = value; }
		}

		private string _ReturnCode;
		public string ReturnCode
		{
			get { return _ReturnCode; }
			set { _ReturnCode = value; }
		}

		private string _ReturnMessage;
		public string ReturnMessage
		{
			get { return _ReturnMessage; }
			set { _ReturnMessage = value; }
		}

		private string _Status;
		public string Status
		{
			get { return _Status; }
			set { _Status = value; }
		}

		private string _ProofOfSale;
		public string ProofOfSale
		{
			get { return _ProofOfSale; }
			set { _ProofOfSale = value; }
		}

		private string _ServiceTaxAmount;
		public string ServiceTaxAmount
		{
			get { return _ServiceTaxAmount; }
			set { _ServiceTaxAmount = value; }
		}

		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#region [ BraspagCaptureCreditCardTransactionResponse ]
	public class BraspagCaptureCreditCardTransactionResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		public List<BraspagTransactionDataResponse> TransactionDataCollection = new List<BraspagTransactionDataResponse>();
		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#endregion

	#region [ VoidCreditCardTransaction ]

	#region [ BraspagVoidCreditCardTransaction ]
	public class BraspagVoidCreditCardTransaction
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		public List<BraspagTransactionDataRequest> TransactionDataCollection = new List<BraspagTransactionDataRequest>();
	}
	#endregion

	#region [ BraspagVoidCreditCardTransactionResponse ]
	public class BraspagVoidCreditCardTransactionResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		public List<BraspagTransactionDataResponse> TransactionDataCollection = new List<BraspagTransactionDataResponse>();
		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#endregion

	#region [ RefundCreditCardTransaction ]

	#region [ BraspagRefundCreditCardTransaction ]
	public class BraspagRefundCreditCardTransaction
	{
		private string _Version;
		public string Version
		{
			get { return _Version; }
			set { _Version = value; }
		}

		private string _RequestId;
		public string RequestId
		{
			get { return _RequestId; }
			set { _RequestId = value; }
		}

		private string _MerchantId;
		public string MerchantId
		{
			get { return _MerchantId; }
			set { _MerchantId = value; }
		}

		public List<BraspagTransactionDataRequest> TransactionDataCollection = new List<BraspagTransactionDataRequest>();
	}
	#endregion

	#region [ BraspagRefundCreditCardTransactionResponse ]
	public class BraspagRefundCreditCardTransactionResponse
	{
		private string _CorrelationId;
		public string CorrelationId
		{
			get { return _CorrelationId; }
			set { _CorrelationId = value; }
		}

		private string _Success;
		public string Success
		{
			get { return _Success; }
			set { _Success = value; }
		}

		public List<BraspagTransactionDataResponse> TransactionDataCollection = new List<BraspagTransactionDataResponse>();
		public List<BraspagErrorReportDataResponse> ErrorReportDataCollection = new List<BraspagErrorReportDataResponse>();
	}
	#endregion

	#endregion

	#region [ BraspagPagOpComplementar ]
	public class BraspagPagOpComplementar
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_pag;
		public int id_pagto_gw_pag
		{
			get { return _id_pagto_gw_pag; }
			set { _id_pagto_gw_pag = value; }
		}

		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _usuario;
		public string usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private string _operacao;
		public string operacao
		{
			get { return _operacao; }
			set { _operacao = value; }
		}

		private DateTime _trx_TX_data;
		public DateTime trx_TX_data
		{
			get { return _trx_TX_data; }
			set { _trx_TX_data = value; }
		}

		private DateTime _trx_TX_data_hora;
		public DateTime trx_TX_data_hora
		{
			get { return _trx_TX_data_hora; }
			set { _trx_TX_data_hora = value; }
		}

		private DateTime _trx_RX_data;
		public DateTime trx_RX_data
		{
			get { return _trx_RX_data; }
			set { _trx_RX_data = value; }
		}

		private DateTime _trx_RX_data_hora;
		public DateTime trx_RX_data_hora
		{
			get { return _trx_RX_data_hora; }
			set { _trx_RX_data_hora = value; }
		}

		private byte _trx_RX_status;
		public byte trx_RX_status
		{
			get { return _trx_RX_status; }
			set { _trx_RX_status = value; }
		}

		private byte _trx_RX_vazio_status;
		public byte trx_RX_vazio_status
		{
			get { return _trx_RX_vazio_status; }
			set { _trx_RX_vazio_status = value; }
		}

		private byte _st_sucesso;
		public byte st_sucesso
		{
			get { return _st_sucesso; }
			set { _st_sucesso = value; }
		}

		private string _req_RequestId;
		public string req_RequestId
		{
			get { return _req_RequestId; }
			set { _req_RequestId = value; }
		}

		private string _req_Version;
		public string req_Version
		{
			get { return _req_Version; }
			set { _req_Version = value; }
		}

		private string _req_MerchantId;
		public string req_MerchantId
		{
			get { return _req_MerchantId; }
			set { _req_MerchantId = value; }
		}

		private string _req_BraspagTransactionId;
		public string req_BraspagTransactionId
		{
			get { return _req_BraspagTransactionId; }
			set { _req_BraspagTransactionId = value; }
		}

		private string _req_OrderId;
		public string req_OrderId
		{
			get { return _req_OrderId; }
			set { _req_OrderId = value; }
		}

		private string _req_Amount;
		public string req_Amount
		{
			get { return _req_Amount; }
			set { _req_Amount = value; }
		}

		private string _req_ServiceTaxAmount;
		public string req_ServiceTaxAmount
		{
			get { return _req_ServiceTaxAmount; }
			set { _req_ServiceTaxAmount = value; }
		}

		private string _resp_BraspagTransactionId;
		public string resp_BraspagTransactionId
		{
			get { return _resp_BraspagTransactionId; }
			set { _resp_BraspagTransactionId = value; }
		}

		private string _resp_AuthorizationCode;
		public string resp_AuthorizationCode
		{
			get { return _resp_AuthorizationCode; }
			set { _resp_AuthorizationCode = value; }
		}

		private string _resp_ProofOfSale;
		public string resp_ProofOfSale
		{
			get { return _resp_ProofOfSale; }
			set { _resp_ProofOfSale = value; }
		}

		private string _resp_Status;
		public string resp_Status
		{
			get { return _resp_Status; }
			set { _resp_Status = value; }
		}
	}
	#endregion

	#region [ BraspagPagOpComplementarXml ]
	public class BraspagPagOpComplementarXml
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_pag_op_complementar;
		public int id_pagto_gw_pag_op_complementar
		{
			get { return _id_pagto_gw_pag_op_complementar; }
			set { _id_pagto_gw_pag_op_complementar = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _tipo_transacao;
		public string tipo_transacao
		{
			get { return _tipo_transacao; }
			set { _tipo_transacao = value; }
		}

		private string _fluxo_xml;
		public string fluxo_xml
		{
			get { return _fluxo_xml; }
			set { _fluxo_xml = value; }
		}

		private string _xml;
		public string xml
		{
			get { return _xml; }
			set { _xml = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentGetTransactionDataResponse ]
	public class BraspagUpdatePagPaymentGetTransactionDataResponse
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private string _ult_GlobalStatus;
		public string ult_GlobalStatus
		{
			get { return _ult_GlobalStatus; }
			set { _ult_GlobalStatus = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}

		private DateTime _resp_CapturedDate;
		public DateTime resp_CapturedDate
		{
			get { return _resp_CapturedDate; }
			set { _resp_CapturedDate = value; }
		}

		private DateTime _resp_VoidedDate;
		public DateTime resp_VoidedDate
		{
			get { return _resp_VoidedDate; }
			set { _resp_VoidedDate = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseSucesso ]
	public class BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseSucesso
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private string _ult_GlobalStatus;
		public string ult_GlobalStatus
		{
			get { return _ult_GlobalStatus; }
			set { _ult_GlobalStatus = value; }
		}

		private DateTime _resp_CapturedDate;
		public DateTime resp_CapturedDate
		{
			get { return _resp_CapturedDate; }
			set { _resp_CapturedDate = value; }
		}

		private byte _captura_confirmada_status;
		public byte captura_confirmada_status
		{
			get { return _captura_confirmada_status; }
			set { _captura_confirmada_status = value; }
		}

		private DateTime _captura_confirmada_data;
		public DateTime captura_confirmada_data
		{
			get { return _captura_confirmada_data; }
			set { _captura_confirmada_data = value; }
		}

		private DateTime _captura_confirmada_data_hora;
		public DateTime captura_confirmada_data_hora
		{
			get { return _captura_confirmada_data_hora; }
			set { _captura_confirmada_data_hora = value; }
		}

		private string _captura_confirmada_usuario;
		public string captura_confirmada_usuario
		{
			get { return _captura_confirmada_usuario; }
			set { _captura_confirmada_usuario = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseFalha ]
	public class BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseFalha
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private byte _captura_confirmada_erro_status;
		public byte captura_confirmada_erro_status
		{
			get { return _captura_confirmada_erro_status; }
			set { _captura_confirmada_erro_status = value; }
		}

		private DateTime _captura_confirmada_erro_data;
		public DateTime captura_confirmada_erro_data
		{
			get { return _captura_confirmada_erro_data; }
			set { _captura_confirmada_erro_data = value; }
		}

		private DateTime _captura_confirmada_erro_data_hora;
		public DateTime captura_confirmada_erro_data_hora
		{
			get { return _captura_confirmada_erro_data_hora; }
			set { _captura_confirmada_erro_data_hora = value; }
		}

		private string _captura_confirmada_erro_mensagem;
		public string captura_confirmada_erro_mensagem
		{
			get { return _captura_confirmada_erro_mensagem; }
			set { _captura_confirmada_erro_mensagem = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentVoidCreditCardTransactionResponseSucesso ]
	public class BraspagUpdatePagPaymentVoidCreditCardTransactionResponseSucesso
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private string _ult_GlobalStatus;
		public string ult_GlobalStatus
		{
			get { return _ult_GlobalStatus; }
			set { _ult_GlobalStatus = value; }
		}

		private DateTime _resp_VoidedDate;
		public DateTime resp_VoidedDate
		{
			get { return _resp_VoidedDate; }
			set { _resp_VoidedDate = value; }
		}

		private byte _voided_status;
		public byte voided_status
		{
			get { return _voided_status; }
			set { _voided_status = value; }
		}

		private DateTime _voided_data;
		public DateTime voided_data
		{
			get { return _voided_data; }
			set { _voided_data = value; }
		}

		private DateTime _voided_data_hora;
		public DateTime voided_data_hora
		{
			get { return _voided_data_hora; }
			set { _voided_data_hora = value; }
		}

		private string _voided_usuario;
		public string voided_usuario
		{
			get { return _voided_usuario; }
			set { _voided_usuario = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentVoidCreditCardTransactionResponseFalha ]
	public class BraspagUpdatePagPaymentVoidCreditCardTransactionResponseFalha
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private byte _voided_erro_status;
		public byte voided_erro_status
		{
			get { return _voided_erro_status; }
			set { _voided_erro_status = value; }
		}

		private DateTime _voided_erro_data;
		public DateTime voided_erro_data
		{
			get { return _voided_erro_data; }
			set { _voided_erro_data = value; }
		}

		private DateTime _voided_erro_data_hora;
		public DateTime voided_erro_data_hora
		{
			get { return _voided_erro_data_hora; }
			set { _voided_erro_data_hora = value; }
		}

		private string _voided_erro_mensagem;
		public string voided_erro_mensagem
		{
			get { return _voided_erro_mensagem; }
			set { _voided_erro_mensagem = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso ]
	public class BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private string _ult_GlobalStatus;
		public string ult_GlobalStatus
		{
			get { return _ult_GlobalStatus; }
			set { _ult_GlobalStatus = value; }
		}

		private DateTime _resp_VoidedDate;
		public DateTime resp_VoidedDate
		{
			get { return _resp_VoidedDate; }
			set { _resp_VoidedDate = value; }
		}

		private byte _refunded_status;
		public byte refunded_status
		{
			get { return _refunded_status; }
			set { _refunded_status = value; }
		}

		private DateTime _refunded_data;
		public DateTime refunded_data
		{
			get { return _refunded_data; }
			set { _refunded_data = value; }
		}

		private DateTime _refunded_data_hora;
		public DateTime refunded_data_hora
		{
			get { return _refunded_data_hora; }
			set { _refunded_data_hora = value; }
		}

		private string _refunded_usuario;
		public string refunded_usuario
		{
			get { return _refunded_usuario; }
			set { _refunded_usuario = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentRefundCreditCardTransactionResponseRefundAccepted ]
	public class BraspagUpdatePagPaymentRefundCreditCardTransactionResponseRefundAccepted
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private string _ult_GlobalStatus;
		public string ult_GlobalStatus
		{
			get { return _ult_GlobalStatus; }
			set { _ult_GlobalStatus = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentRefundCreditCardTransactionResponseFalha ]
	public class BraspagUpdatePagPaymentRefundCreditCardTransactionResponseFalha
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private byte _refunded_erro_status;
		public byte refunded_erro_status
		{
			get { return _refunded_erro_status; }
			set { _refunded_erro_status = value; }
		}

		private DateTime _refunded_erro_data;
		public DateTime refunded_erro_data
		{
			get { return _refunded_erro_data; }
			set { _refunded_erro_data = value; }
		}

		private DateTime _refunded_erro_data_hora;
		public DateTime refunded_erro_data_hora
		{
			get { return _refunded_erro_data_hora; }
			set { _refunded_erro_data_hora = value; }
		}

		private string _refunded_erro_mensagem;
		public string refunded_erro_mensagem
		{
			get { return _refunded_erro_mensagem; }
			set { _refunded_erro_mensagem = value; }
		}

		private string _ult_atualizacao_usuario;
		public string ult_atualizacao_usuario
		{
			get { return _ult_atualizacao_usuario; }
			set { _ult_atualizacao_usuario = value; }
		}

		private int _ult_id_pagto_gw_pag_payment_op_complementar;
		public int ult_id_pagto_gw_pag_payment_op_complementar
		{
			get { return _ult_id_pagto_gw_pag_payment_op_complementar; }
			set { _ult_id_pagto_gw_pag_payment_op_complementar = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentPagtoRegPedido ]
	public class BraspagUpdatePagPaymentPagtoRegPedido
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private byte _pagto_registrado_no_pedido_status;
		public byte pagto_registrado_no_pedido_status
		{
			get { return _pagto_registrado_no_pedido_status; }
			set { _pagto_registrado_no_pedido_status = value; }
		}

		private string _pagto_registrado_no_pedido_tipo_operacao;
		public string pagto_registrado_no_pedido_tipo_operacao
		{
			get { return _pagto_registrado_no_pedido_tipo_operacao; }
			set { _pagto_registrado_no_pedido_tipo_operacao = value; }
		}

		private string _pagto_registrado_no_pedido_usuario;
		public string pagto_registrado_no_pedido_usuario
		{
			get { return _pagto_registrado_no_pedido_usuario; }
			set { _pagto_registrado_no_pedido_usuario = value; }
		}

		private string _pagto_registrado_no_pedido_id_pedido_pagamento;
		public string pagto_registrado_no_pedido_id_pedido_pagamento
		{
			get { return _pagto_registrado_no_pedido_id_pedido_pagamento; }
			set { _pagto_registrado_no_pedido_id_pedido_pagamento = value; }
		}

		private string _pagto_registrado_no_pedido_st_pagto_anterior;
		public string pagto_registrado_no_pedido_st_pagto_anterior
		{
			get { return _pagto_registrado_no_pedido_st_pagto_anterior; }
			set { _pagto_registrado_no_pedido_st_pagto_anterior = value; }
		}

		private string _pagto_registrado_no_pedido_st_pagto_novo;
		public string pagto_registrado_no_pedido_st_pagto_novo
		{
			get { return _pagto_registrado_no_pedido_st_pagto_novo; }
			set { _pagto_registrado_no_pedido_st_pagto_novo = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentEstornoRegPedido ]
	public class BraspagUpdatePagPaymentEstornoRegPedido
	{
		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private byte _estorno_registrado_no_pedido_status;
		public byte estorno_registrado_no_pedido_status
		{
			get { return _estorno_registrado_no_pedido_status; }
			set { _estorno_registrado_no_pedido_status = value; }
		}

		private string _estorno_registrado_no_pedido_tipo_operacao;
		public string estorno_registrado_no_pedido_tipo_operacao
		{
			get { return _estorno_registrado_no_pedido_tipo_operacao; }
			set { _estorno_registrado_no_pedido_tipo_operacao = value; }
		}

		private string _estorno_registrado_no_pedido_usuario;
		public string estorno_registrado_no_pedido_usuario
		{
			get { return _estorno_registrado_no_pedido_usuario; }
			set { _estorno_registrado_no_pedido_usuario = value; }
		}

		private string _estorno_registrado_no_pedido_id_pedido_pagamento;
		public string estorno_registrado_no_pedido_id_pedido_pagamento
		{
			get { return _estorno_registrado_no_pedido_id_pedido_pagamento; }
			set { _estorno_registrado_no_pedido_id_pedido_pagamento = value; }
		}

		private string _estorno_registrado_no_pedido_st_pagto_anterior;
		public string estorno_registrado_no_pedido_st_pagto_anterior
		{
			get { return _estorno_registrado_no_pedido_st_pagto_anterior; }
			set { _estorno_registrado_no_pedido_st_pagto_anterior = value; }
		}

		private string _estorno_registrado_no_pedido_st_pagto_novo;
		public string estorno_registrado_no_pedido_st_pagto_novo
		{
			get { return _estorno_registrado_no_pedido_st_pagto_novo; }
			set { _estorno_registrado_no_pedido_st_pagto_novo = value; }
		}
	}
	#endregion

	#region [ BraspagUpdatePagPaymentRefundPendingConfirmado ]
	public class BraspagUpdatePagPaymentRefundPendingConfirmado
	{
		public int id_pagto_gw_pag_payment { get; set; }
		public string refund_pending_confirmado_usuario { get; set; }
	}
	#endregion

	#region [ BraspagUpdatePagPaymentRefundPendingFalha ]
	public class BraspagUpdatePagPaymentRefundPendingFalha
	{
		public int id_pagto_gw_pag_payment { get; set; }
		public string refund_pending_falha_motivo { get; set; }
	}
	#endregion

	#region [ BraspagWebhook ]
	public class BraspagWebhook
	{
		public int Id { get; set; }
		public DateTime DataCadastro { get; set; }
		public DateTime DataHoraCadastro { get; set; }
		public string Empresa { get; set; }
		public string NumPedido { get; set; }
		public string Status { get; set; }
		public string CODPAGAMENTO { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public DateTime? BraspagDadosComplementaresQueryDataHora { get; set; }
		public byte EmailEnviadoStatus { get; set; }
		public DateTime? EmailEnviadoDataHora { get; set; }
		public int ProcessamentoErpStatus { get; set; }
		public DateTime? ProcessamentoErpDataHora { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
		public DateTime? BraspagDadosComplementaresQueryDtHrUltTentativa { get; set; }
		public string MsgErro { get; set; }
		public string MsgErroTemporario { get; set; }
	}
	#endregion

	#region [ BraspagWebhookComplementar ]
	public class BraspagWebhookComplementar
	{
		public int Id { get; set; }
		public int id_braspag_webhook { get; set; }
		public DateTime Data { get; set; }
		public DateTime DataHora { get; set; }
		public string BraspagTransactionId { get; set; }
		public string BraspagOrderId { get; set; }
		public string PaymentMethod { get; set; }
		public string GlobalStatus { get; set; }
		public DateTime? ReceivedDate { get; set; }
		public DateTime? CapturedDate { get; set; }
		public string CustomerName { get; set; }
		public DateTime? BoletoExpirationDate { get; set; }
		public string Amount { get; set; }
		public decimal ValorAmount { get; set; }
		public string PaidAmount { get; set; }
		public decimal ValorPaidAmount { get; set; }
		public string pedido { get; set; }
		public byte PagtoRegistradoNoPedidoStatus { get; set; }
		public string PagtoRegistradoNoPedidoTipoOperacao { get; set; }
		public DateTime? PagtoRegistradoNoPedidoData { get; set; }
		public DateTime? PagtoRegistradoNoPedidoDataHora { get; set; }
		public string PagtoRegistradoNoPedido_id_pedido_pagamento { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoAnterior { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoNovo { get; set; }
		public int AnaliseCreditoStatusAnterior { get; set; }
		public int AnaliseCreditoStatusNovo { get; set; }
		public byte PedidoHistPagtoGravadoStatus { get; set; }
		public DateTime? PedidoHistPagtoGravadoData { get; set; }
		public DateTime? PedidoHistPagtoGravadoDataHora { get; set; }
		public string MsgErro { get; set; }
	}
	#endregion

	#region [ BraspagWebhookDadosConsolidadosBoleto ]
	public class BraspagWebhookDadosConsolidadosBoleto
	{
		public string MerchantId { get; set; }
		public string OrderId { get; set; }
		public string BraspagTransactionId { get; set; }
		public string BraspagOrderId { get; set; }
		public string PaymentMethod { get; set; }
		public string GlobalStatus { get; set; }
		public DateTime? ReceivedDate { get; set; }
		public DateTime? CapturedDate { get; set; }
		public string CustomerName { get; set; }
		public DateTime? BoletoExpirationDate { get; set; }
		public string Amount { get; set; }
		public decimal ValorAmount { get; set; }
		public string PaidAmount { get; set; }
		public decimal ValorPaidAmount { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva ]
	public class BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva
	{
		public int id_braspag_webhook { get; set; }
		public byte EmailEnviadoStatus { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
		public string MsgErro { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookQueryDadosComplementaresFalhaTemporaria ]
	public class BraspagUpdateWebhookQueryDadosComplementaresFalhaTemporaria
	{
		public int id_braspag_webhook { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
		public string MsgErroTemporario { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookQueryDadosComplementaresQtdeTentativas ]
	public class BraspagUpdateWebhookQueryDadosComplementaresQtdeTentativas
	{
		public int id_braspag_webhook { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookQueryDadosComplementaresSucesso ]
	public class BraspagUpdateWebhookQueryDadosComplementaresSucesso
	{
		public int id_braspag_webhook { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
	}
	#endregion

	#region [ BraspagInsertWebhookQueryDadosComplementares ]
	public class BraspagInsertWebhookQueryDadosComplementares
	{
		public int Id { get; set; }
		public int id_braspag_webhook { get; set; }
		public string BraspagTransactionId { get; set; }
		public string BraspagOrderId { get; set; }
		public string PaymentMethod { get; set; }
		public string GlobalStatus { get; set; }
		public DateTime? ReceivedDate { get; set; }
		public DateTime? CapturedDate { get; set; }
		public string CustomerName { get; set; }
		public DateTime? BoletoExpirationDate { get; set; }
		public string Amount { get; set; }
		public decimal ValorAmount { get; set; }
		public string PaidAmount { get; set; }
		public decimal ValorPaidAmount { get; set; }
		public string pedido { get; set; }
	}
	#endregion

	#region [ BraspagWebhookComplementarUpdatePagtoRegPedido ]
	public class BraspagWebhookComplementarUpdatePagtoRegPedido
	{
		public int id_braspag_webhook_complementar { get; set; }
		public byte PagtoRegistradoNoPedidoStatus { get; set; }
		public string PagtoRegistradoNoPedidoTipoOperacao { get; set; }
		public string PagtoRegistradoNoPedido_id_pedido_pagamento { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoAnterior { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoNovo { get; set; }
		public int AnaliseCreditoStatusAnterior { get; set; }
		public int AnaliseCreditoStatusNovo { get; set; }
	}
	#endregion

	#region [ BraspagWebhookV2 ]
	public class BraspagWebhookV2
	{
		public int Id { get; set; }
		public DateTime DataCadastro { get; set; }
		public DateTime DataHoraCadastro { get; set; }
		public string Empresa { get; set; }
		public string RecurrentPaymentId { get; set; }
		public string PaymentId { get; set; }
		public byte ChangeType { get; set; }
		public string OrderIdIdentificado { get; set; }
		public string PaymentMethodIdentificado { get; set; }
		public byte ProcessadoStatus { get; set; }
		public DateTime ProcessadoDataHora { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public DateTime? BraspagDadosComplementaresQueryDataHora { get; set; }
		public byte EmailEnviadoStatus { get; set; }
		public DateTime? EmailEnviadoDataHora { get; set; }
		public int ProcessamentoErpStatus { get; set; }
		public DateTime? ProcessamentoErpDataHora { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
		public DateTime? BraspagDadosComplementaresQueryDtHrUltTentativa { get; set; }
		public string MsgErro { get; set; }
		public string MsgErroTemporario { get; set; }
	}
	#endregion

	#region [ BraspagWebhookV2Complementar ]
	public class BraspagWebhookV2Complementar
	{
		public int Id { get; set; }
		public int id_braspag_webhook_v2 { get; set; }
		public DateTime Data { get; set; }
		public DateTime DataHora { get; set; }
		public string BraspagTransactionId { get; set; }
		public string BraspagOrderId { get; set; }
		public string PaymentMethod { get; set; }
		public string GlobalStatus { get; set; }
		public DateTime? ReceivedDate { get; set; }
		public DateTime? CapturedDate { get; set; }
		public string CustomerName { get; set; }
		public DateTime? BoletoExpirationDate { get; set; }
		public string Amount { get; set; }
		public decimal ValorAmount { get; set; }
		public string PaidAmount { get; set; }
		public decimal ValorPaidAmount { get; set; }
		public string pedido { get; set; }
		public byte PagtoRegistradoNoPedidoStatus { get; set; }
		public string PagtoRegistradoNoPedidoTipoOperacao { get; set; }
		public DateTime? PagtoRegistradoNoPedidoData { get; set; }
		public DateTime? PagtoRegistradoNoPedidoDataHora { get; set; }
		public string PagtoRegistradoNoPedido_id_pedido_pagamento { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoAnterior { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoNovo { get; set; }
		public int AnaliseCreditoStatusAnterior { get; set; }
		public int AnaliseCreditoStatusNovo { get; set; }
		public byte PedidoHistPagtoGravadoStatus { get; set; }
		public DateTime? PedidoHistPagtoGravadoData { get; set; }
		public DateTime? PedidoHistPagtoGravadoDataHora { get; set; }
		public string MsgErro { get; set; }
	}
	#endregion

	#region [ BraspagWebhookV2DadosConsolidadosBoleto ]
	public class BraspagWebhookV2DadosConsolidadosBoleto
	{
		public string MerchantId { get; set; }
		public string OrderId { get; set; }
		public string BraspagTransactionId { get; set; }
		public string BraspagOrderId { get; set; }
		public string PaymentMethod { get; set; }
		public string GlobalStatus { get; set; }
		public DateTime? ReceivedDate { get; set; }
		public DateTime? CapturedDate { get; set; }
		public string CustomerName { get; set; }
		public DateTime? BoletoExpirationDate { get; set; }
		public string Amount { get; set; }
		public decimal ValorAmount { get; set; }
		public string PaidAmount { get; set; }
		public decimal ValorPaidAmount { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookV2PaymentMethodIdentificado ]
	public class BraspagUpdateWebhookV2PaymentMethodIdentificado
	{
		public int Id { get; set; }
		public string OrderIdIdentificado { get; set; }
		public string PaymentMethodIdentificado { get; set; }
		public byte ProcessadoStatus { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookV2QueryDadosComplementaresFalhaDefinitiva ]
	public class BraspagUpdateWebhookV2QueryDadosComplementaresFalhaDefinitiva
	{
		public int id_braspag_webhook_v2 { get; set; }
		public byte EmailEnviadoStatus { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
		public string MsgErro { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookV2QueryDadosComplementaresFalhaTemporaria ]
	public class BraspagUpdateWebhookV2QueryDadosComplementaresFalhaTemporaria
	{
		public int id_braspag_webhook_v2 { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
		public string MsgErroTemporario { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookV2QueryDadosComplementaresQtdeTentativas ]
	public class BraspagUpdateWebhookV2QueryDadosComplementaresQtdeTentativas
	{
		public int id_braspag_webhook_v2 { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
	}
	#endregion

	#region [ BraspagUpdateWebhookV2QueryDadosComplementaresSucesso ]
	public class BraspagUpdateWebhookV2QueryDadosComplementaresSucesso
	{
		public int id_braspag_webhook_v2 { get; set; }
		public byte BraspagDadosComplementaresQueryStatus { get; set; }
		public int BraspagDadosComplementaresQueryTentativas { get; set; }
	}
	#endregion

	#region [ BraspagInsertWebhookV2QueryDadosComplementares ]
	public class BraspagInsertWebhookV2QueryDadosComplementares
	{
		public int Id { get; set; }
		public int id_braspag_webhook_v2 { get; set; }
		public string BraspagTransactionId { get; set; }
		public string BraspagOrderId { get; set; }
		public string PaymentMethod { get; set; }
		public string GlobalStatus { get; set; }
		public DateTime? ReceivedDate { get; set; }
		public DateTime? CapturedDate { get; set; }
		public string CustomerName { get; set; }
		public DateTime? BoletoExpirationDate { get; set; }
		public string Amount { get; set; }
		public decimal ValorAmount { get; set; }
		public string PaidAmount { get; set; }
		public decimal ValorPaidAmount { get; set; }
		public string pedido { get; set; }
	}
	#endregion

	#region [ BraspagWebhookV2ComplementarUpdatePagtoRegPedido ]
	public class BraspagWebhookV2ComplementarUpdatePagtoRegPedido
	{
		public int id_braspag_webhook_v2_complementar { get; set; }
		public byte PagtoRegistradoNoPedidoStatus { get; set; }
		public string PagtoRegistradoNoPedidoTipoOperacao { get; set; }
		public string PagtoRegistradoNoPedido_id_pedido_pagamento { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoAnterior { get; set; }
		public string PagtoRegistradoNoPedidoStPagtoNovo { get; set; }
		public int AnaliseCreditoStatusAnterior { get; set; }
		public int AnaliseCreditoStatusNovo { get; set; }
	}
	#endregion
}
