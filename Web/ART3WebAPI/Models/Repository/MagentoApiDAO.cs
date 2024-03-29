﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Data.SqlClient;
using System.Data;
using ART3WebAPI.Models.Domains;
using System.Text;
using System.Threading;

namespace ART3WebAPI.Models.Repository
{
	public class MagentoApiDAO
	{
		#region [ getLoginParameters ]
		public static MagentoApiLoginParameters getLoginParameters(string loja, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			string senha_criptografada;
			MagentoApiLoginParameters parameters;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			#endregion

			msg_erro = "";
			try
			{
				if ((loja ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o número da loja para obter os parâmetros de login da API do Magento!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT" +
							" magento_api_urlWebService," +
							" magento_api_username," +
							" magento_api_password," +
							" magento_api_versao," +
							" magento_api_rest_endpoint," +
							" magento_api_rest_access_token," +
							" magento_api_rest_force_get_sales_order_by_entity_id," +
							" magento_api_rest_prefixo_num_magento" +
						" FROM t_LOJA" +
						" WHERE" +
							" (loja = '" + loja + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado o registro da loja '" + loja + "'";
						return null;
					}

					row = dtbResultado.Rows[0];

					parameters = new MagentoApiLoginParameters();
					parameters.urlWebService = BD.readToString(row["magento_api_urlWebService"]);
					parameters.username = BD.readToString(row["magento_api_username"]);
					senha_criptografada = BD.readToString(row["magento_api_password"]);
					parameters.password = Domains.Criptografia.Descriptografa(senha_criptografada);
					parameters.api_versao = BD.readToInt(row["magento_api_versao"]);
					parameters.api_rest_endpoint = BD.readToString(row["magento_api_rest_endpoint"]);
					parameters.api_rest_access_token = BD.readToString(row["magento_api_rest_access_token"]);
					parameters.api_rest_force_get_sales_order_by_entity_id = BD.readToByte(row["magento_api_rest_force_get_sales_order_by_entity_id"]);
					parameters.magento_api_rest_prefixo_num_magento = BD.readToString(row["magento_api_rest_prefixo_num_magento"]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return parameters;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ magentoPedidoXmlLoadFromDataRow ]
		public static MagentoErpPedidoXml magentoPedidoXmlLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			MagentoErpPedidoXml pedidoXml = new MagentoErpPedidoXml();
			#endregion

			pedidoXml.id = BD.readToInt(rowDados["id"]);
			pedidoXml.operationControlTicket = BD.readToString(rowDados["operationControlTicket"]);
			pedidoXml.loja = BD.readToString(rowDados["loja"]);
			pedidoXml.pedido_magento = BD.readToString(rowDados["pedido_magento"]);
			pedidoXml.pedido_erp = BD.readToString(rowDados["pedido_erp"]);
			pedidoXml.pedido_marketplace = BD.readToString(rowDados["pedido_marketplace"]);
			pedidoXml.pedido_marketplace_completo = BD.readToString(rowDados["pedido_marketplace_completo"]);
			pedidoXml.marketplace_codigo_origem = BD.readToString(rowDados["marketplace_codigo_origem"]);
			pedidoXml.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			pedidoXml.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			pedidoXml.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			pedidoXml.magento_api_versao = BD.readToInt(rowDados["magento_api_versao"]);
			pedidoXml.pedido_xml = BD.readToString(rowDados["pedido_xml"]);
			pedidoXml.pedido_json = BD.readToString(rowDados["pedido_json"]);
			pedidoXml.cpfCnpjIdentificado = BD.readToString(rowDados["cpfCnpjIdentificado"]);
			pedidoXml.increment_id = BD.readToInt(rowDados["increment_id"]);
			pedidoXml.created_at = BD.readToString(rowDados["created_at"]);
			pedidoXml.updated_at = BD.readToString(rowDados["updated_at"]);
			pedidoXml.customer_id = BD.readToInt(rowDados["customer_id"]);
			pedidoXml.billing_address_id = BD.readToInt(rowDados["billing_address_id"]);
			pedidoXml.shipping_address_id = BD.readToInt(rowDados["shipping_address_id"]);
			pedidoXml.status = BD.readToString(rowDados["status"]);
			pedidoXml.status_descricao = BD.readToString(rowDados["status_descricao"]);
			pedidoXml.state = BD.readToString(rowDados["state"]);
			pedidoXml.state_descricao = BD.readToString(rowDados["state_descricao"]);
			pedidoXml.customer_email = BD.readToString(rowDados["customer_email"]);
			pedidoXml.customer_firstname = BD.readToString(rowDados["customer_firstname"]);
			pedidoXml.customer_lastname = BD.readToString(rowDados["customer_lastname"]);
			pedidoXml.customer_middlename = BD.readToString(rowDados["customer_middlename"]);
			pedidoXml.quote_id = BD.readToInt(rowDados["quote_id"]);
			pedidoXml.customer_group_id = BD.readToInt(rowDados["customer_group_id"]);
			pedidoXml.order_id = BD.readToInt(rowDados["order_id"]);
			pedidoXml.customer_dob = BD.readToString(rowDados["customer_dob"]);
			pedidoXml.clearsale_status_code = BD.readToString(rowDados["clearsale_status_code"]);
			pedidoXml.clearSale_status = BD.readToString(rowDados["clearSale_status"]);
			pedidoXml.clearSale_score = BD.readToString(rowDados["clearSale_score"]);
			pedidoXml.clearSale_packageID = BD.readToString(rowDados["clearSale_packageID"]);
			pedidoXml.shipping_amount = BD.readToDecimal(rowDados["shipping_amount"]);
			pedidoXml.discount_amount = BD.readToDecimal(rowDados["discount_amount"]);
			pedidoXml.subtotal = BD.readToDecimal(rowDados["subtotal"]);
			pedidoXml.grand_total = BD.readToDecimal(rowDados["grand_total"]);
			pedidoXml.installer_document = BD.readToString(rowDados["installer_document"]);
			pedidoXml.installer_id = BD.readToInt(rowDados["installer_id"]);
			pedidoXml.commission_value = BD.readToDecimal(rowDados["commission_value"]);
			pedidoXml.commission_discount = BD.readToDecimal(rowDados["commission_discount"]);
			pedidoXml.commission_final_discount = BD.readToDecimal(rowDados["commission_final_discount"]);
			pedidoXml.commission_final_value = BD.readToDecimal(rowDados["commission_final_value"]);
			pedidoXml.commission_discount_type = BD.readToString(rowDados["commission_discount_type"]);
			pedidoXml.mktp_datasource_status = BD.readToByte(rowDados["mktp_datasource_status"]);
			pedidoXml.mktp_datasource_discount = BD.readToDecimal(rowDados["mktp_datasource_discount"]);
			pedidoXml.mktp_datasource_total_ordered = BD.readToDecimal(rowDados["mktp_datasource_total_ordered"]);
			pedidoXml.mktp_datasource_shipping_cost = BD.readToDecimal(rowDados["mktp_datasource_shipping_cost"]);
			pedidoXml.b2b_installer_name = BD.readToString(rowDados["b2b_installer_name"]);
			pedidoXml.b2b_installer_id = BD.readToInt(rowDados["b2b_installer_id"]);
			pedidoXml.b2b_installer_commission_value = BD.readToDecimal(rowDados["b2b_installer_commission_value"]);
			pedidoXml.b2b_installer_commission_percentage = BD.readToSingle(rowDados["b2b_installer_commission_percentage"]);
			pedidoXml.b2b_type_order = BD.readToString(rowDados["b2b_type_order"]);

			return pedidoXml;
		}
		#endregion

		#region [ insertMagentoPedidoXml ]
		public static bool insertMagentoPedidoXml(Guid? httpRequestId, MagentoErpPedidoXml pedidoXml, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "insertMagentoPedidoXml()";
			const int TAMANHO_CAMPO_USUARIO_CADASTRO = 20;
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string msg;
			string msg_erro_aux = "";
			string strSql;
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			MagentoErpPedidoXml pedidoXmlBD;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_MAGENTO_API_PEDIDO_XML (" +
								"operationControlTicket, " +
								"loja, " +
								"pedido_magento, " +
								"pedido_erp, " +
								"pedido_marketplace, " +
								"pedido_marketplace_completo, " +
								"marketplace_codigo_origem, " +
								"usuario_cadastro, " +
								"magento_api_versao, " +
								"pedido_xml, " +
								"pedido_json, " +
								"cpfCnpjIdentificado, " +
								"increment_id, " +
								"created_at, " +
								"updated_at, " +
								"customer_id, " +
								"billing_address_id, " +
								"shipping_address_id, " +
								"status, " +
								"status_descricao, " +
								"state, " +
								"state_descricao, " +
								"customer_email, " +
								"customer_firstname, " +
								"customer_lastname, " +
								"customer_middlename, " +
								"quote_id, " +
								"customer_group_id, " +
								"order_id, " +
								"customer_dob, " +
								"clearsale_status_code, " +
								"clearSale_status, " +
								"clearSale_score, " +
								"clearSale_packageID, " +
								"shipping_amount, " +
								"shipping_discount_amount, " +
								"discount_amount, " +
								"subtotal, " +
								"grand_total, " +
								"installer_document, " +
								"installer_id, " +
								"commission_value, " +
								"commission_discount, " +
								"commission_final_discount, " +
								"commission_final_value, " +
								"commission_discount_type, " +
								"mktp_datasource_status, " +
								"mktp_datasource_discount, " +
								"mktp_datasource_total_ordered, " +
								"mktp_datasource_shipping_cost, " +
								"b2b_installer_name, " +
								"b2b_installer_id, " +
								"b2b_installer_commission_value, " +
								"b2b_installer_commission_percentage, " +
								"b2b_type_order" +
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@operationControlTicket, " +
								"@loja, " +
								"@pedido_magento, " +
								"@pedido_erp, " +
								"@pedido_marketplace, " +
								"@pedido_marketplace_completo, " +
								"@marketplace_codigo_origem, " +
								"@usuario_cadastro, " +
								"@magento_api_versao, " +
								"@pedido_xml, " +
								"@pedido_json, " +
								"@cpfCnpjIdentificado, " +
								"@increment_id, " +
								"@created_at, " +
								"@updated_at, " +
								"@customer_id, " +
								"@billing_address_id, " +
								"@shipping_address_id, " +
								"@status, " +
								"@status_descricao, " +
								"@state, " +
								"@state_descricao, " +
								"@customer_email, " +
								"@customer_firstname, " +
								"@customer_lastname, " +
								"@customer_middlename, " +
								"@quote_id, " +
								"@customer_group_id, " +
								"@order_id, " +
								"@customer_dob, " +
								"@clearsale_status_code, " +
								"@clearSale_status, " +
								"@clearSale_score, " +
								"@clearSale_packageID," +
								"@shipping_amount, " +
								"@shipping_discount_amount, " +
								"@discount_amount, " +
								"@subtotal, " +
								"@grand_total, " +
								"@installer_document, " +
								"@installer_id, " +
								"@commission_value, " +
								"@commission_discount, " +
								"@commission_final_discount, " +
								"@commission_final_value, " +
								"@commission_discount_type, " +
								"@mktp_datasource_status, " +
								"@mktp_datasource_discount, " +
								"@mktp_datasource_total_ordered, " +
								"@mktp_datasource_shipping_cost," +
								"@b2b_installer_name, " +
								"@b2b_installer_id, " +
								"@b2b_installer_commission_value, " +
								"@b2b_installer_commission_percentage, " +
								"@b2b_type_order" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@operationControlTicket", SqlDbType.UniqueIdentifier);
					cmInsert.Parameters.Add("@loja", SqlDbType.VarChar, 3);
					cmInsert.Parameters.Add("@pedido_magento", SqlDbType.VarChar, 9);
					cmInsert.Parameters.Add("@pedido_erp", SqlDbType.VarChar, 9);
					cmInsert.Parameters.Add("@pedido_marketplace", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@pedido_marketplace_completo", SqlDbType.VarChar, 30);
					cmInsert.Parameters.Add("@marketplace_codigo_origem", SqlDbType.VarChar, 3);
					cmInsert.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, TAMANHO_CAMPO_USUARIO_CADASTRO);
					cmInsert.Parameters.Add("@magento_api_versao", SqlDbType.Int);
					cmInsert.Parameters.Add("@pedido_xml", SqlDbType.VarChar, -1); // varchar(max)
					cmInsert.Parameters.Add("@pedido_json", SqlDbType.VarChar, -1); // varchar(max)
					cmInsert.Parameters.Add("@cpfCnpjIdentificado", SqlDbType.VarChar, 14);
					cmInsert.Parameters.Add("@increment_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@created_at", SqlDbType.VarChar, 19);
					cmInsert.Parameters.Add("@updated_at", SqlDbType.VarChar, 19);
					cmInsert.Parameters.Add("@customer_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@billing_address_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@shipping_address_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@status", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@status_descricao", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@state", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@state_descricao", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@customer_email", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@customer_firstname", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@customer_lastname", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@customer_middlename", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@quote_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@customer_group_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@order_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@customer_dob", SqlDbType.VarChar, 19);
					cmInsert.Parameters.Add("@clearsale_status_code", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@clearSale_status", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@clearSale_score", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@clearSale_packageID", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add(new SqlParameter("@shipping_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@shipping_discount_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@discount_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@subtotal", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@grand_total", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@installer_document", SqlDbType.VarChar, 18);
					cmInsert.Parameters.Add("@installer_id", SqlDbType.Int);
					cmInsert.Parameters.Add(new SqlParameter("@commission_value", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@commission_discount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@commission_final_discount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@commission_final_value", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@commission_discount_type", SqlDbType.VarChar, 10);
					cmInsert.Parameters.Add("@mktp_datasource_status", SqlDbType.TinyInt);
					cmInsert.Parameters.Add(new SqlParameter("@mktp_datasource_discount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@mktp_datasource_total_ordered", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@mktp_datasource_shipping_cost", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@b2b_installer_name", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@b2b_installer_id", SqlDbType.Int);
					cmInsert.Parameters.Add(new SqlParameter("@b2b_installer_commission_value", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@b2b_installer_commission_percentage", SqlDbType.Real);
					cmInsert.Parameters.Add("@b2b_type_order", SqlDbType.VarChar, 40);
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@operationControlTicket"].Value = Guid.Parse(pedidoXml.operationControlTicket);
							cmInsert.Parameters["@loja"].Value = (pedidoXml.loja ?? "");
							cmInsert.Parameters["@pedido_magento"].Value = (pedidoXml.pedido_magento ?? "");
							cmInsert.Parameters["@pedido_erp"].Value = (pedidoXml.pedido_erp ?? "");
							cmInsert.Parameters["@pedido_marketplace"].Value = (pedidoXml.pedido_marketplace ?? "");
							cmInsert.Parameters["@pedido_marketplace_completo"].Value = (pedidoXml.pedido_marketplace_completo ?? "");
							cmInsert.Parameters["@marketplace_codigo_origem"].Value = (pedidoXml.marketplace_codigo_origem ?? "");
							cmInsert.Parameters["@usuario_cadastro"].Value = Global.leftStr((pedidoXml.usuario_cadastro ?? ""), TAMANHO_CAMPO_USUARIO_CADASTRO);
							cmInsert.Parameters["@magento_api_versao"].Value = pedidoXml.magento_api_versao;
							cmInsert.Parameters["@pedido_xml"].Value = (pedidoXml.pedido_xml ?? "");
							cmInsert.Parameters["@pedido_json"].Value = (pedidoXml.pedido_json ?? "");
							cmInsert.Parameters["@cpfCnpjIdentificado"].Value = Global.digitos((pedidoXml.cpfCnpjIdentificado ?? ""));
							cmInsert.Parameters["@increment_id"].Value = pedidoXml.increment_id;
							cmInsert.Parameters["@created_at"].Value = (pedidoXml.created_at ?? "");
							cmInsert.Parameters["@updated_at"].Value = (pedidoXml.updated_at ?? "");
							cmInsert.Parameters["@customer_id"].Value = pedidoXml.customer_id;
							cmInsert.Parameters["@billing_address_id"].Value = pedidoXml.billing_address_id;
							cmInsert.Parameters["@shipping_address_id"].Value = pedidoXml.shipping_address_id;
							cmInsert.Parameters["@status"].Value = (pedidoXml.status ?? "");
							cmInsert.Parameters["@status_descricao"].Value = (pedidoXml.status_descricao ?? "");
							cmInsert.Parameters["@state"].Value = (pedidoXml.state ?? "");
							cmInsert.Parameters["@state_descricao"].Value = (pedidoXml.state_descricao ?? "");
							cmInsert.Parameters["@customer_email"].Value = (pedidoXml.customer_email ?? "");
							cmInsert.Parameters["@customer_firstname"].Value = (pedidoXml.customer_firstname ?? "");
							cmInsert.Parameters["@customer_lastname"].Value = (pedidoXml.customer_lastname ?? "");
							cmInsert.Parameters["@customer_middlename"].Value = (pedidoXml.customer_middlename ?? "");
							cmInsert.Parameters["@quote_id"].Value = pedidoXml.quote_id;
							cmInsert.Parameters["@customer_group_id"].Value = pedidoXml.customer_group_id;
							cmInsert.Parameters["@order_id"].Value = pedidoXml.order_id;
							cmInsert.Parameters["@customer_dob"].Value = (pedidoXml.customer_dob ?? "");
							cmInsert.Parameters["@clearsale_status_code"].Value = (pedidoXml.clearsale_status_code ?? "");
							cmInsert.Parameters["@clearSale_status"].Value = (pedidoXml.clearSale_status ?? "");
							cmInsert.Parameters["@clearSale_score"].Value = (pedidoXml.clearSale_score ?? "");
							cmInsert.Parameters["@clearSale_packageID"].Value = (pedidoXml.clearSale_packageID ?? "");
							cmInsert.Parameters["@shipping_amount"].Value = pedidoXml.shipping_amount;
							cmInsert.Parameters["@shipping_discount_amount"].Value = pedidoXml.shipping_discount_amount;
							cmInsert.Parameters["@discount_amount"].Value = pedidoXml.discount_amount;
							cmInsert.Parameters["@subtotal"].Value = pedidoXml.subtotal;
							cmInsert.Parameters["@grand_total"].Value = pedidoXml.grand_total;
							cmInsert.Parameters["@installer_document"].Value = (pedidoXml.installer_document ?? "");
							cmInsert.Parameters["@installer_id"].Value = pedidoXml.installer_id;
							cmInsert.Parameters["@commission_value"].Value = pedidoXml.commission_value;
							cmInsert.Parameters["@commission_discount"].Value = pedidoXml.commission_discount;
							cmInsert.Parameters["@commission_final_discount"].Value = pedidoXml.commission_final_discount;
							cmInsert.Parameters["@commission_final_value"].Value = pedidoXml.commission_final_value;
							cmInsert.Parameters["@commission_discount_type"].Value = (pedidoXml.commission_discount_type ?? "");
							cmInsert.Parameters["@mktp_datasource_status"].Value = pedidoXml.mktp_datasource_status;
							cmInsert.Parameters["@mktp_datasource_discount"].Value = pedidoXml.mktp_datasource_discount;
							cmInsert.Parameters["@mktp_datasource_total_ordered"].Value = pedidoXml.mktp_datasource_total_ordered;
							cmInsert.Parameters["@mktp_datasource_shipping_cost"].Value = pedidoXml.mktp_datasource_shipping_cost;
							cmInsert.Parameters["@b2b_installer_name"].Value = (pedidoXml.b2b_installer_name ?? "");
							cmInsert.Parameters["@b2b_installer_id"].Value = pedidoXml.b2b_installer_id;
							cmInsert.Parameters["@b2b_installer_commission_value"].Value = pedidoXml.b2b_installer_commission_value;
							cmInsert.Parameters["@b2b_installer_commission_percentage"].Value = pedidoXml.b2b_installer_commission_percentage;
							cmInsert.Parameters["@b2b_type_order"].Value = (pedidoXml.b2b_type_order ?? "");
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								pedidoXml.id = generatedId;
							}
							catch (Exception ex)
							{
								generatedId = 0;
								msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
								Global.gravaLogAtividade(httpRequestId, msg);
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								pedidoXmlBD = getMagentoPedidoXmlById(generatedId, out msg_erro_aux);
								pedidoXml.operationControlTicket = pedidoXmlBD.operationControlTicket;
								pedidoXml.dt_cadastro = pedidoXmlBD.dt_cadastro;
								pedidoXml.dt_hr_cadastro = pedidoXmlBD.dt_hr_cadastro;

								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion
						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados as informações do pedido Magento obtidos através da API após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ insertMagentoPedidoXmlDecodeEndereco ]
		public static bool insertMagentoPedidoXmlDecodeEndereco(Guid? httpRequestId, MagentoErpPedidoXmlDecodeEndereco endereco, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "insertMagentoPedidoXmlDecodeEndereco()";
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string msg;
			string strSql;
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_MAGENTO_API_PEDIDO_XML_DECODE_ENDERECO (" +
								"id_magento_api_pedido_xml, " +
								"tipo_endereco, " +
								"endereco, " +
								"endereco_numero, " +
								"endereco_complemento, " +
								"bairro, " +
								"cidade, " +
								"uf, " +
								"cep, " +
								"address_id, " +
								"parent_id, " +
								"customer_address_id, " +
								"quote_address_id, " +
								"region_id, " +
								"address_type, " +
								"street, " +
								"city, " +
								"region, " +
								"postcode, " +
								"country_id, " +
								"firstname, " +
								"middlename, " +
								"lastname, " +
								"email, " +
								"telephone, " +
								"celular, " +
								"fax, " +
								"tipopessoa, " +
								"rg, " +
								"ie, " +
								"cpfcnpj, " +
								"empresa, " +
								"nomefantasia, " +
								"street_detail"+
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@id_magento_api_pedido_xml, " +
								"@tipo_endereco, " +
								"@endereco, " +
								"@endereco_numero, " +
								"@endereco_complemento, " +
								"@bairro, " +
								"@cidade, " +
								"@uf, " +
								"@cep, " +
								"@address_id, " +
								"@parent_id, " +
								"@customer_address_id, " +
								"@quote_address_id, " +
								"@region_id, " +
								"@address_type, " +
								"@street, " +
								"@city, " +
								"@region, " +
								"@postcode, " +
								"@country_id, " +
								"@firstname, " +
								"@middlename, " +
								"@lastname, " +
								"@email, " +
								"@telephone, " +
								"@celular, " +
								"@fax, " +
								"@tipopessoa, " +
								"@rg, " +
								"@ie, " +
								"@cpfcnpj, " +
								"@empresa, " +
								"@nomefantasia, " +
								"@street_detail"+
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@id_magento_api_pedido_xml", SqlDbType.Int);
					cmInsert.Parameters.Add("@tipo_endereco", SqlDbType.VarChar, 3);
					cmInsert.Parameters.Add("@endereco", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@endereco_numero", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@endereco_complemento", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@bairro", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@cidade", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@uf", SqlDbType.VarChar, 2);
					cmInsert.Parameters.Add("@cep", SqlDbType.VarChar, 8);
					cmInsert.Parameters.Add("@address_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@parent_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@customer_address_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@quote_address_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@region_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@address_type", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@street", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@city", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@region", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@postcode", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@country_id", SqlDbType.VarChar, 10);
					cmInsert.Parameters.Add("@firstname", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@middlename", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@lastname", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@email", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@telephone", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@celular", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@fax", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@tipopessoa", SqlDbType.VarChar, 10);
					cmInsert.Parameters.Add("@rg", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@ie", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@cpfcnpj", SqlDbType.VarChar, 30);
					cmInsert.Parameters.Add("@empresa", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@nomefantasia", SqlDbType.VarChar, 200);
					cmInsert.Parameters.Add("@street_detail", SqlDbType.VarChar, 200);
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@id_magento_api_pedido_xml"].Value = endereco.id_magento_api_pedido_xml;
							cmInsert.Parameters["@tipo_endereco"].Value = (endereco.tipo_endereco ?? "");
							cmInsert.Parameters["@endereco"].Value = (endereco.endereco ?? "");
							cmInsert.Parameters["@endereco_numero"].Value = (endereco.endereco_numero ?? "");
							cmInsert.Parameters["@endereco_complemento"].Value = (endereco.endereco_complemento ?? "");
							cmInsert.Parameters["@bairro"].Value = (endereco.bairro ?? "");
							cmInsert.Parameters["@cidade"].Value = (endereco.cidade ?? "");
							cmInsert.Parameters["@uf"].Value = (endereco.uf ?? "");
							cmInsert.Parameters["@cep"].Value = (endereco.cep ?? "");
							cmInsert.Parameters["@address_id"].Value = endereco.address_id;
							cmInsert.Parameters["@parent_id"].Value = endereco.parent_id;
							cmInsert.Parameters["@customer_address_id"].Value = endereco.customer_address_id;
							cmInsert.Parameters["@quote_address_id"].Value = endereco.quote_address_id;
							cmInsert.Parameters["@region_id"].Value = endereco.region_id;
							cmInsert.Parameters["@address_type"].Value = (endereco.address_type ?? "");
							cmInsert.Parameters["@street"].Value = (endereco.street ?? "");
							cmInsert.Parameters["@city"].Value = (endereco.city ?? "");
							cmInsert.Parameters["@region"].Value = (endereco.region ?? "");
							cmInsert.Parameters["@postcode"].Value = (endereco.postcode ?? "");
							cmInsert.Parameters["@country_id"].Value = (endereco.country_id ?? "");
							cmInsert.Parameters["@firstname"].Value = (endereco.firstname ?? "");
							cmInsert.Parameters["@middlename"].Value = (endereco.middlename ?? "");
							cmInsert.Parameters["@lastname"].Value = (endereco.lastname ?? "");
							cmInsert.Parameters["@email"].Value = (endereco.email ?? "");
							cmInsert.Parameters["@telephone"].Value = (endereco.telephone ?? "");
							cmInsert.Parameters["@celular"].Value = (endereco.celular ?? "");
							cmInsert.Parameters["@fax"].Value = (endereco.fax ?? "");
							cmInsert.Parameters["@tipopessoa"].Value = (endereco.tipopessoa ?? "");
							cmInsert.Parameters["@rg"].Value = (endereco.rg ?? "");
							cmInsert.Parameters["@ie"].Value = (endereco.ie ?? "");
							cmInsert.Parameters["@cpfcnpj"].Value = (endereco.cpfcnpj ?? "");
							cmInsert.Parameters["@empresa"].Value = (endereco.empresa ?? "");
							cmInsert.Parameters["@nomefantasia"].Value = (endereco.nomefantasia ?? "");
							cmInsert.Parameters["@street_detail"].Value = (endereco.street_detail ?? "");
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								endereco.id = generatedId;
							}
							catch (Exception ex)
							{
								generatedId = 0;
								msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
								Global.gravaLogAtividade(httpRequestId, msg);
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion
						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados as informações do endereço do pedido Magento obtidos através da API após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ insertMagentoPedidoXmlDecodeItem ]
		public static bool insertMagentoPedidoXmlDecodeItem(Guid? httpRequestId, MagentoErpPedidoXmlDecodeItem produtoItem, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "insertMagentoPedidoXmlDecodeItem()";
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string msg;
			string strSql;
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_MAGENTO_API_PEDIDO_XML_DECODE_ITEM (" +
								"id_magento_api_pedido_xml, " +
								"sku, " +
								"qty_ordered, " +
								"product_id, " +
								"item_id, " +
								"order_id, " +
								"quote_item_id, " +
								"price, " +
								"base_price, " +
								"original_price, " +
								"base_original_price, " +
								"discount_percent, " +
								"discount_amount, " +
								"base_discount_amount, " +
								"name, " +
								"product_type, " +
								"has_children, " +
								"parent_item_id," +
								"weight," +
								"is_virtual," +
								"free_shipping," +
								"is_qty_decimal," +
								"no_discount," +
								"qty_canceled," +
								"qty_invoiced," +
								"qty_refunded," +
								"qty_shipped," +
								"tax_percent," +
								"tax_amount," +
								"base_tax_amount," +
								"tax_invoiced," +
								"base_tax_invoiced," +
								"discount_invoiced," +
								"base_discount_invoiced," +
								"amount_refunded," +
								"base_amount_refunded," +
								"row_total," +
								"base_row_total," +
								"row_invoiced," +
								"base_row_invoiced," +
								"row_weight," +
								"price_incl_tax," +
								"base_price_incl_tax," +
								"row_total_incl_tax," +
								"base_row_total_incl_tax," +
								"mktp_datasource_special_price,"+
								"mktp_datasource_shipping_cost,"+
								"mktp_datasource_original_price"+
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@id_magento_api_pedido_xml, " +
								"@sku, " +
								"@qty_ordered, " +
								"@product_id, " +
								"@item_id, " +
								"@order_id, " +
								"@quote_item_id, " +
								"@price, " +
								"@base_price, " +
								"@original_price, " +
								"@base_original_price, " +
								"@discount_percent, " +
								"@discount_amount, " +
								"@base_discount_amount, " +
								"@name, " +
								"@product_type, " +
								"@has_children, " +
								"@parent_item_id," +
								"@weight," +
								"@is_virtual," +
								"@free_shipping," +
								"@is_qty_decimal," +
								"@no_discount," +
								"@qty_canceled," +
								"@qty_invoiced," +
								"@qty_refunded," +
								"@qty_shipped," +
								"@tax_percent," +
								"@tax_amount," +
								"@base_tax_amount," +
								"@tax_invoiced," +
								"@base_tax_invoiced," +
								"@discount_invoiced," +
								"@base_discount_invoiced," +
								"@amount_refunded," +
								"@base_amount_refunded," +
								"@row_total," +
								"@base_row_total," +
								"@row_invoiced," +
								"@base_row_invoiced," +
								"@row_weight," +
								"@price_incl_tax," +
								"@base_price_incl_tax," +
								"@row_total_incl_tax," +
								"@base_row_total_incl_tax," +
								"@mktp_datasource_special_price," +
								"@mktp_datasource_shipping_cost," +
								"@mktp_datasource_original_price" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@id_magento_api_pedido_xml", SqlDbType.Int);
					cmInsert.Parameters.Add("@sku", SqlDbType.VarChar, 8);
					cmInsert.Parameters.Add(new SqlParameter("@qty_ordered", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@product_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@item_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@order_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@quote_item_id", SqlDbType.Int);
					cmInsert.Parameters.Add(new SqlParameter("@price", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_price", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@original_price", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_original_price", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@discount_percent", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@discount_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_discount_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@name", SqlDbType.VarChar, 400);
					cmInsert.Parameters.Add("@product_type", SqlDbType.VarChar, 30);
					cmInsert.Parameters.Add("@has_children", SqlDbType.VarChar, 10);
					cmInsert.Parameters.Add("@parent_item_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@weight", SqlDbType.Real);
					cmInsert.Parameters.Add("@is_virtual", SqlDbType.SmallInt);
					cmInsert.Parameters.Add("@free_shipping", SqlDbType.SmallInt);
					cmInsert.Parameters.Add("@is_qty_decimal", SqlDbType.SmallInt);
					cmInsert.Parameters.Add("@no_discount", SqlDbType.SmallInt);
					cmInsert.Parameters.Add(new SqlParameter("@qty_canceled", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@qty_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@qty_refunded", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@qty_shipped", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@tax_percent", SqlDbType.Real);
					cmInsert.Parameters.Add(new SqlParameter("@tax_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_tax_amount", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@tax_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_tax_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@discount_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_discount_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@amount_refunded", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_amount_refunded", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@row_total", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_row_total", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@row_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_row_invoiced", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@row_weight", SqlDbType.Real);
					cmInsert.Parameters.Add(new SqlParameter("@price_incl_tax", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_price_incl_tax", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@row_total_incl_tax", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@base_row_total_incl_tax", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@mktp_datasource_special_price", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@mktp_datasource_shipping_cost", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add(new SqlParameter("@mktp_datasource_original_price", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@id_magento_api_pedido_xml"].Value = produtoItem.id_magento_api_pedido_xml;
							cmInsert.Parameters["@sku"].Value = (produtoItem.sku ?? "");
							cmInsert.Parameters["@qty_ordered"].Value = produtoItem.qty_ordered;
							cmInsert.Parameters["@product_id"].Value = produtoItem.product_id;
							cmInsert.Parameters["@item_id"].Value = produtoItem.item_id;
							cmInsert.Parameters["@order_id"].Value = produtoItem.order_id;
							cmInsert.Parameters["@quote_item_id"].Value = produtoItem.quote_item_id;
							cmInsert.Parameters["@price"].Value = produtoItem.price;
							cmInsert.Parameters["@base_price"].Value = produtoItem.base_price;
							cmInsert.Parameters["@original_price"].Value = produtoItem.original_price;
							cmInsert.Parameters["@base_original_price"].Value = produtoItem.base_original_price;
							cmInsert.Parameters["@discount_percent"].Value = produtoItem.discount_percent;
							cmInsert.Parameters["@discount_amount"].Value = produtoItem.discount_amount;
							cmInsert.Parameters["@base_discount_amount"].Value = produtoItem.base_discount_amount;
							cmInsert.Parameters["@name"].Value = (produtoItem.name ?? "");
							cmInsert.Parameters["@product_type"].Value = (produtoItem.product_type ?? "");
							cmInsert.Parameters["@has_children"].Value = (produtoItem.has_children ?? "");
							cmInsert.Parameters["@parent_item_id"].Value = produtoItem.parent_item_id;
							cmInsert.Parameters["@weight"].Value = produtoItem.weight;
							cmInsert.Parameters["@is_virtual"].Value = produtoItem.is_virtual;
							cmInsert.Parameters["@free_shipping"].Value = produtoItem.free_shipping;
							cmInsert.Parameters["@is_qty_decimal"].Value = produtoItem.is_qty_decimal;
							cmInsert.Parameters["@no_discount"].Value = produtoItem.no_discount;
							cmInsert.Parameters["@qty_canceled"].Value = produtoItem.qty_canceled;
							cmInsert.Parameters["@qty_invoiced"].Value = produtoItem.qty_invoiced;
							cmInsert.Parameters["@qty_refunded"].Value = produtoItem.qty_refunded;
							cmInsert.Parameters["@qty_shipped"].Value = produtoItem.qty_shipped;
							cmInsert.Parameters["@tax_percent"].Value = produtoItem.tax_percent;
							cmInsert.Parameters["@tax_amount"].Value = produtoItem.tax_amount;
							cmInsert.Parameters["@base_tax_amount"].Value = produtoItem.base_tax_amount;
							cmInsert.Parameters["@tax_invoiced"].Value = produtoItem.tax_invoiced;
							cmInsert.Parameters["@base_tax_invoiced"].Value = produtoItem.base_tax_invoiced;
							cmInsert.Parameters["@discount_invoiced"].Value = produtoItem.discount_invoiced;
							cmInsert.Parameters["@base_discount_invoiced"].Value = produtoItem.base_discount_invoiced;
							cmInsert.Parameters["@amount_refunded"].Value = produtoItem.amount_refunded;
							cmInsert.Parameters["@base_amount_refunded"].Value = produtoItem.base_amount_refunded;
							cmInsert.Parameters["@row_total"].Value = produtoItem.row_total;
							cmInsert.Parameters["@base_row_total"].Value = produtoItem.base_row_total;
							cmInsert.Parameters["@row_invoiced"].Value = produtoItem.row_invoiced;
							cmInsert.Parameters["@base_row_invoiced"].Value = produtoItem.base_row_invoiced;
							cmInsert.Parameters["@row_weight"].Value = produtoItem.row_weight;
							cmInsert.Parameters["@price_incl_tax"].Value = produtoItem.price_incl_tax;
							cmInsert.Parameters["@base_price_incl_tax"].Value = produtoItem.base_price_incl_tax;
							cmInsert.Parameters["@row_total_incl_tax"].Value = produtoItem.row_total_incl_tax;
							cmInsert.Parameters["@base_row_total_incl_tax"].Value = produtoItem.base_row_total_incl_tax;
							cmInsert.Parameters["@mktp_datasource_special_price"].Value = produtoItem.mktp_datasource_special_price;
							cmInsert.Parameters["@mktp_datasource_shipping_cost"].Value = produtoItem.mktp_datasource_shipping_cost;
							cmInsert.Parameters["@mktp_datasource_original_price"].Value = produtoItem.mktp_datasource_original_price;
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								produtoItem.id = generatedId;
							}
							catch (Exception ex)
							{
								generatedId = 0;
								msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
								Global.gravaLogAtividade(httpRequestId, msg);
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion
						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados as informações do item do pedido Magento obtidos através da API após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ insertMagentoPedidoXmlDecodeStatusHistory ]
		public static bool insertMagentoPedidoXmlDecodeStatusHistory(Guid? httpRequestId, MagentoErpPedidoXmlDecodeStatusHistory statusHistory, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "insertMagentoPedidoXmlDecodeStatusHistory()";
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string msg;
			string strSql;
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_MAGENTO_API_PEDIDO_XML_DECODE_STATUS_HISTORY (" +
								"id_magento_api_pedido_xml, " +
								"parent_id, " +
								"is_customer_notified, " +
								"is_visible_on_front, " +
								"comment, " +
								"status, " +
								"created_at, " +
								"entity_name, " +
								"store_id" +
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@id_magento_api_pedido_xml, " +
								"@parent_id, " +
								"@is_customer_notified, " +
								"@is_visible_on_front, " +
								"@comment, " +
								"@status, " +
								"@created_at, " +
								"@entity_name, " +
								"@store_id" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@id_magento_api_pedido_xml", SqlDbType.Int);
					cmInsert.Parameters.Add("@parent_id", SqlDbType.Int);
					cmInsert.Parameters.Add("@is_customer_notified", SqlDbType.TinyInt);
					cmInsert.Parameters.Add("@is_visible_on_front", SqlDbType.TinyInt);
					cmInsert.Parameters.Add("@comment", SqlDbType.VarChar, -1); // varchar(max)
					cmInsert.Parameters.Add("@status", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@created_at", SqlDbType.VarChar, 19);
					cmInsert.Parameters.Add("@entity_name", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@store_id", SqlDbType.Int);
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@id_magento_api_pedido_xml"].Value = statusHistory.id_magento_api_pedido_xml;
							cmInsert.Parameters["@parent_id"].Value = statusHistory.parent_id;
							cmInsert.Parameters["@is_customer_notified"].Value = statusHistory.is_customer_notified;
							cmInsert.Parameters["@is_visible_on_front"].Value = statusHistory.is_visible_on_front;
							cmInsert.Parameters["@comment"].Value = (statusHistory.comment ?? "");
							cmInsert.Parameters["@status"].Value = (statusHistory.status ?? "");
							cmInsert.Parameters["@created_at"].Value = (statusHistory.created_at ?? "");
							cmInsert.Parameters["@entity_name"].Value = (statusHistory.entity_name ?? "");
							cmInsert.Parameters["@store_id"].Value = statusHistory.store_id;
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								statusHistory.id = generatedId;
							}
							catch (Exception ex)
							{
								generatedId = 0;
								msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
								Global.gravaLogAtividade(httpRequestId, msg);
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion
						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados as informações do status history do pedido Magento obtidos através da API após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ insertMagentoPedidoXmlDecodeSkyhubMktpPayment ]
		public static bool insertMagentoPedidoXmlDecodeSkyhubMktpPayment(Guid? httpRequestId, MagentoErpPedidoXmlDecodeSkyhubMktpPayment payment, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "insertMagentoPedidoXmlDecodeSkyhubMktpPayment()";
			bool blnSucesso = false;
			int generatedId;
			int intQtdeTentativas = 0;
			string msg;
			string strSql;
			StringBuilder sbLog = new StringBuilder("");
			SqlConnection cn;
			SqlCommand cmInsert;
			#endregion

			msg_erro = "";
			try
			{
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ cmInsert ]
					strSql = "INSERT INTO t_MAGENTO_API_PEDIDO_XML_DECODE_SKYHUB_MKTP_PAYMENT (" +
								"id_magento_api_pedido_xml, " +
								"value, " +
								"type, " +
								"transaction_date, " +
								"status, " +
								"parcels, " +
								"method, " +
								"description, " +
								"card_issuer, " +
								"autorization_id, " +
								"sefaz_type_integration, " +
								"sefaz_payment_indicator, " +
								"sefaz_name_payment, " +
								"sefaz_name_card_issuer, " +
								"sefaz_id_payment, " +
								"sefaz_id_card_issuer" +
							")" +
							" OUTPUT INSERTED.id" +
							" VALUES " +
							"(" +
								"@id_magento_api_pedido_xml, " +
								"@value, " +
								"@type, " +
								"@transaction_date, " +
								"@status, " +
								"@parcels, " +
								"@method, " +
								"@description, " +
								"@card_issuer, " +
								"@autorization_id, " +
								"@sefaz_type_integration, " +
								"@sefaz_payment_indicator, " +
								"@sefaz_name_payment, " +
								"@sefaz_name_card_issuer, " +
								"@sefaz_id_payment, " +
								"@sefaz_id_card_issuer" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@id_magento_api_pedido_xml", SqlDbType.Int);
					cmInsert.Parameters.Add(new SqlParameter("@value", SqlDbType.Decimal) { Precision = 18, Scale = 4 });
					cmInsert.Parameters.Add("@type", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@transaction_date", SqlDbType.VarChar, 40);
					cmInsert.Parameters.Add("@status", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@parcels", SqlDbType.Int);
					cmInsert.Parameters.Add("@method", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@description", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@card_issuer", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@autorization_id", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@sefaz_type_integration", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@sefaz_payment_indicator", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@sefaz_name_payment", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@sefaz_name_card_issuer", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@sefaz_id_payment", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@sefaz_id_card_issuer", SqlDbType.VarChar, 80);
					cmInsert.Prepare();
					#endregion

					try
					{
						#region [ Laço de tentativas de inserção no banco de dados ]
						do
						{
							intQtdeTentativas++;
							msg_erro = "";

							#region [ Preenche o valor dos parâmetros ]
							cmInsert.Parameters["@id_magento_api_pedido_xml"].Value = payment.id_magento_api_pedido_xml;
							cmInsert.Parameters["@value"].Value = (object)payment.value ?? DBNull.Value;
							cmInsert.Parameters["@type"].Value = (object)payment.type ?? DBNull.Value;
							cmInsert.Parameters["@transaction_date"].Value = (object)payment.transaction_date ?? DBNull.Value;
							cmInsert.Parameters["@status"].Value = (object)payment.status ?? DBNull.Value;
							cmInsert.Parameters["@parcels"].Value = (object)payment.parcels ?? DBNull.Value;
							cmInsert.Parameters["@method"].Value = (object)payment.method ?? DBNull.Value;
							cmInsert.Parameters["@description"].Value = (object)payment.description ?? DBNull.Value;
							cmInsert.Parameters["@card_issuer"].Value = (object)payment.card_issuer ?? DBNull.Value;
							cmInsert.Parameters["@autorization_id"].Value = (object)payment.autorization_id ?? DBNull.Value;
							cmInsert.Parameters["@sefaz_type_integration"].Value = (object)payment.sefaz_type_integration ?? DBNull.Value;
							cmInsert.Parameters["@sefaz_payment_indicator"].Value = (object)payment.sefaz_payment_indicator ?? DBNull.Value;
							cmInsert.Parameters["@sefaz_name_payment"].Value = (object)payment.sefaz_name_payment ?? DBNull.Value;
							cmInsert.Parameters["@sefaz_name_card_issuer"].Value = (object)payment.sefaz_name_card_issuer ?? DBNull.Value;
							cmInsert.Parameters["@sefaz_id_payment"].Value = (object)payment.sefaz_id_payment ?? DBNull.Value;
							cmInsert.Parameters["@sefaz_id_card_issuer"].Value = (object)payment.sefaz_id_card_issuer ?? DBNull.Value;
							#endregion

							#region [ Monta texto para o log em arquivo ]
							// Se houver conteúdo de alguma tentativa anterior, descarta
							sbLog = new StringBuilder("");
							foreach (SqlParameter item in cmInsert.Parameters)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								generatedId = (int)cmInsert.ExecuteScalar();
								payment.id = generatedId;
							}
							catch (Exception ex)
							{
								generatedId = 0;
								msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
								Global.gravaLogAtividade(httpRequestId, msg);
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (generatedId > 0)
							{
								blnSucesso = true;
							}
							else
							{
								Thread.Sleep(100);
							}
							#endregion
						} while ((!blnSucesso) && (intQtdeTentativas < 5));
						#endregion

						#region [ Processamento final de sucesso ou falha ]
						if (blnSucesso)
						{
							return true;
						}
						else
						{
							msg_erro = "Falha ao gravar no banco de dados de pagamento do pedido Magento informados pelo marketplace obtidos através da API após " + intQtdeTentativas.ToString() + " tentativas!!";
							return false;
						}
						#endregion
					}
					catch (Exception ex)
					{
						msg_erro = ex.Message;
						return false;
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ getMagentoPedidoXmlByTicket ]
		public static MagentoErpPedidoXml getMagentoPedidoXmlByTicket(string numeroPedidoMagento, string operationControlTicket, int api_versao, out string msg_erro)
		{
			#region [ Declarações ]
			MagentoErpPedidoXml pedidoXml;
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((numeroPedidoMagento ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o número do pedido Magento!";
					return null;
				}

				if ((operationControlTicket ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o ticket de controle da operação!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_MAGENTO_API_PEDIDO_XML" +
							" WHERE" +
								" (operationControlTicket = '" + operationControlTicket + "')" +
								" AND (pedido_magento = '" + numeroPedidoMagento + "')" +
								" AND (magento_api_versao = " + api_versao.ToString() + ")";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado o registro do pedido Magento " + numeroPedidoMagento + " com o ticket de controle da operação " + operationControlTicket;
						return null;
					}

					pedidoXml = magentoPedidoXmlLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return pedidoXml;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getMagentoPedidoXmlById ]
		public static MagentoErpPedidoXml getMagentoPedidoXmlById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			MagentoErpPedidoXml pedidoXml;
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if (id == 0)
				{
					msg_erro = "Não foi informado o ID do registro para recuperar os dados do pedido Magento armazenados no BD!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_MAGENTO_API_PEDIDO_XML" +
							" WHERE" +
								" (id = " + id.ToString() + ")";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado o registro do pedido Magento (id = " + id.ToString() + ")";
						return null;
					}

					pedidoXml = magentoPedidoXmlLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return pedidoXml;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion
	}
}