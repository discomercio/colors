using System;
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
	public class OrcamentistaIndicadorDAO
	{
		#region [ orcamentistaIndicadorCompletoLoadFromDataRow ]
		public static OrcamentistaIndicadorCompleto orcamentistaIndicadorCompletoLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			OrcamentistaIndicadorCompleto orcamentistaIndicador = new OrcamentistaIndicadorCompleto();
			#endregion

			orcamentistaIndicador.apelido = BD.readToString(rowDados["apelido"]);
			orcamentistaIndicador.cnpj_cpf = BD.readToString(rowDados["cnpj_cpf"]);
			orcamentistaIndicador.tipo = BD.readToString(rowDados["tipo"]);
			orcamentistaIndicador.ie_rg = BD.readToString(rowDados["ie_rg"]);
			orcamentistaIndicador.razao_social_nome = BD.readToString(rowDados["razao_social_nome"]);
			orcamentistaIndicador.endereco = BD.readToString(rowDados["endereco"]);
			orcamentistaIndicador.endereco_numero = BD.readToString(rowDados["endereco_numero"]);
			orcamentistaIndicador.endereco_complemento = BD.readToString(rowDados["endereco_complemento"]);
			orcamentistaIndicador.bairro = BD.readToString(rowDados["bairro"]);
			orcamentistaIndicador.cidade = BD.readToString(rowDados["cidade"]);
			orcamentistaIndicador.uf = BD.readToString(rowDados["uf"]);
			orcamentistaIndicador.cep = BD.readToString(rowDados["cep"]);
			orcamentistaIndicador.ddd = BD.readToString(rowDados["ddd"]);
			orcamentistaIndicador.telefone = BD.readToString(rowDados["telefone"]);
			orcamentistaIndicador.fax = BD.readToString(rowDados["fax"]);
			orcamentistaIndicador.ddd_cel = BD.readToString(rowDados["ddd_cel"]);
			orcamentistaIndicador.tel_cel = BD.readToString(rowDados["tel_cel"]);
			orcamentistaIndicador.contato = BD.readToString(rowDados["contato"]);
			orcamentistaIndicador.banco = BD.readToString(rowDados["banco"]);
			orcamentistaIndicador.agencia = BD.readToString(rowDados["agencia"]);
			orcamentistaIndicador.conta = BD.readToString(rowDados["conta"]);
			orcamentistaIndicador.favorecido = BD.readToString(rowDados["favorecido"]);
			orcamentistaIndicador.loja = BD.readToString(rowDados["loja"]);
			orcamentistaIndicador.vendedor = BD.readToString(rowDados["vendedor"]);
			orcamentistaIndicador.hab_acesso_sistema = (int)BD.readToInt(rowDados["hab_acesso_sistema"]);
			orcamentistaIndicador.status = BD.readToString(rowDados["status"]);
			orcamentistaIndicador.senha = BD.readToString(rowDados["senha"]);
			orcamentistaIndicador.datastamp = BD.readToString(rowDados["datastamp"]);
			orcamentistaIndicador.dt_ult_alteracao_senha = BD.readToDateTime(rowDados["dt_ult_alteracao_senha"]);
			orcamentistaIndicador.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			orcamentistaIndicador.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			orcamentistaIndicador.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			orcamentistaIndicador.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);
			orcamentistaIndicador.dt_ult_acesso = BD.readToDateTime(rowDados["dt_ult_acesso"]);
			orcamentistaIndicador.desempenho_nota = BD.readToString(rowDados["desempenho_nota"]);
			orcamentistaIndicador.desempenho_nota_data = BD.readToDateTime(rowDados["desempenho_nota_data"]);
			orcamentistaIndicador.desempenho_nota_usuario = BD.readToString(rowDados["desempenho_nota_usuario"]);
			orcamentistaIndicador.perc_desagio_RA = BD.readToSingle(rowDados["perc_desagio_RA"]);
			orcamentistaIndicador.vl_limite_mensal = BD.readToDecimal(rowDados["vl_limite_mensal"]);
			orcamentistaIndicador.email = BD.readToString(rowDados["email"]);
			orcamentistaIndicador.captador = BD.readToString(rowDados["captador"]);
			orcamentistaIndicador.checado_status = (int)BD.readToInt(rowDados["checado_status"]);
			orcamentistaIndicador.checado_data = BD.readToDateTime(rowDados["checado_data"]);
			orcamentistaIndicador.checado_usuario = BD.readToString(rowDados["checado_usuario"]);
			orcamentistaIndicador.obs = BD.readToString(rowDados["obs"]);
			orcamentistaIndicador.vl_meta = BD.readToDecimal(rowDados["vl_meta"]);
			orcamentistaIndicador.UsuarioUltAtualizVlMeta = BD.readToString(rowDados["UsuarioUltAtualizVlMeta"]);
			orcamentistaIndicador.DtHrUltAtualizVlMeta = BD.readToDateTime(rowDados["DtHrUltAtualizVlMeta"]);
			orcamentistaIndicador.permite_RA_status = (int)BD.readToInt(rowDados["permite_RA_status"]);
			orcamentistaIndicador.permite_RA_usuario = BD.readToString(rowDados["permite_RA_usuario"]);
			orcamentistaIndicador.permite_RA_data_hora = BD.readToDateTime(rowDados["permite_RA_data_hora"]);
			orcamentistaIndicador.forma_como_conheceu_codigo = BD.readToString(rowDados["forma_como_conheceu_codigo"]);
			orcamentistaIndicador.forma_como_conheceu_usuario = BD.readToString(rowDados["forma_como_conheceu_usuario"]);
			orcamentistaIndicador.forma_como_conheceu_data = BD.readToDateTime(rowDados["forma_como_conheceu_data"]);
			orcamentistaIndicador.forma_como_conheceu_data_hora = BD.readToDateTime(rowDados["forma_como_conheceu_data_hora"]);
			orcamentistaIndicador.forma_como_conheceu_codigo_anterior = BD.readToString(rowDados["forma_como_conheceu_codigo_anterior"]);
			orcamentistaIndicador.nome_fantasia = BD.readToString(rowDados["nome_fantasia"]);
			orcamentistaIndicador.tipo_estabelecimento = (int)BD.readToInt(rowDados["tipo_estabelecimento"]);
			orcamentistaIndicador.nextel = BD.readToString(rowDados["nextel"]);
			orcamentistaIndicador.email2 = BD.readToString(rowDados["email2"]);
			orcamentistaIndicador.email3 = BD.readToString(rowDados["email3"]);
			orcamentistaIndicador.razao_social_nome_iniciais_em_maiusculas = BD.readToString(rowDados["razao_social_nome_iniciais_em_maiusculas"]);
			orcamentistaIndicador.st_reg_copiado_automaticamente = BD.readToByte(rowDados["st_reg_copiado_automaticamente"]);
			orcamentistaIndicador.dt_hr_reg_atualizado_automaticamente = BD.readToDateTime(rowDados["dt_hr_reg_atualizado_automaticamente"]);
			orcamentistaIndicador.etq_endereco = BD.readToString(rowDados["etq_endereco"]);
			orcamentistaIndicador.etq_endereco_numero = BD.readToString(rowDados["etq_endereco_numero"]);
			orcamentistaIndicador.etq_endereco_complemento = BD.readToString(rowDados["etq_endereco_complemento"]);
			orcamentistaIndicador.etq_bairro = BD.readToString(rowDados["etq_bairro"]);
			orcamentistaIndicador.etq_cidade = BD.readToString(rowDados["etq_cidade"]);
			orcamentistaIndicador.etq_uf = BD.readToString(rowDados["etq_uf"]);
			orcamentistaIndicador.etq_cep = BD.readToString(rowDados["etq_cep"]);
			orcamentistaIndicador.etq_email = BD.readToString(rowDados["etq_email"]);
			orcamentistaIndicador.etq_ddd_1 = BD.readToString(rowDados["etq_ddd_1"]);
			orcamentistaIndicador.etq_tel_1 = BD.readToString(rowDados["etq_tel_1"]);
			orcamentistaIndicador.etq_ddd_2 = BD.readToString(rowDados["etq_ddd_2"]);
			orcamentistaIndicador.etq_tel_2 = BD.readToString(rowDados["etq_tel_2"]);
			orcamentistaIndicador.favorecido_cnpj_cpf = BD.readToString(rowDados["favorecido_cnpj_cpf"]);
			orcamentistaIndicador.agencia_dv = BD.readToString(rowDados["agencia_dv"]);
			orcamentistaIndicador.conta_operacao = BD.readToString(rowDados["conta_operacao"]);
			orcamentistaIndicador.conta_dv = BD.readToString(rowDados["conta_dv"]);
			orcamentistaIndicador.tipo_conta = BD.readToString(rowDados["tipo_conta"]);
			orcamentistaIndicador.vendedor_dt_ult_atualizacao = BD.readToDateTime(rowDados["vendedor_dt_ult_atualizacao"]);
			orcamentistaIndicador.vendedor_dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["vendedor_dt_hr_ult_atualizacao"]);
			orcamentistaIndicador.vendedor_usuario_ult_atualizacao = BD.readToString(rowDados["vendedor_usuario_ult_atualizacao"]);

			return orcamentistaIndicador;
		}
		#endregion

		#region [ orcamentistaIndicadorBasicoLoadFromDataRow ]
		public static OrcamentistaIndicadorBasico orcamentistaIndicadorBasicoLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			OrcamentistaIndicadorBasico orcamentistaIndicador = new OrcamentistaIndicadorBasico();
			#endregion

			orcamentistaIndicador.apelido = BD.readToString(rowDados["apelido"]);
			orcamentistaIndicador.cnpj_cpf = BD.readToString(rowDados["cnpj_cpf"]);
			orcamentistaIndicador.tipo = BD.readToString(rowDados["tipo"]);
			orcamentistaIndicador.ie_rg = BD.readToString(rowDados["ie_rg"]);
			orcamentistaIndicador.razao_social_nome = BD.readToString(rowDados["razao_social_nome"]);
			orcamentistaIndicador.endereco = BD.readToString(rowDados["endereco"]);
			orcamentistaIndicador.endereco_numero = BD.readToString(rowDados["endereco_numero"]);
			orcamentistaIndicador.endereco_complemento = BD.readToString(rowDados["endereco_complemento"]);
			orcamentistaIndicador.bairro = BD.readToString(rowDados["bairro"]);
			orcamentistaIndicador.cidade = BD.readToString(rowDados["cidade"]);
			orcamentistaIndicador.uf = BD.readToString(rowDados["uf"]);
			orcamentistaIndicador.cep = BD.readToString(rowDados["cep"]);
			orcamentistaIndicador.ddd = BD.readToString(rowDados["ddd"]);
			orcamentistaIndicador.telefone = BD.readToString(rowDados["telefone"]);
			orcamentistaIndicador.fax = BD.readToString(rowDados["fax"]);
			orcamentistaIndicador.ddd_cel = BD.readToString(rowDados["ddd_cel"]);
			orcamentistaIndicador.tel_cel = BD.readToString(rowDados["tel_cel"]);
			orcamentistaIndicador.contato = BD.readToString(rowDados["contato"]);
			orcamentistaIndicador.loja = BD.readToString(rowDados["loja"]);
			orcamentistaIndicador.vendedor = BD.readToString(rowDados["vendedor"]);
			orcamentistaIndicador.status = BD.readToString(rowDados["status"]);
			orcamentistaIndicador.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			orcamentistaIndicador.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			orcamentistaIndicador.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			orcamentistaIndicador.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);
			orcamentistaIndicador.email = BD.readToString(rowDados["email"]);
			orcamentistaIndicador.captador = BD.readToString(rowDados["captador"]);
			orcamentistaIndicador.permite_RA_status = (int)BD.readToInt(rowDados["permite_RA_status"]);
			orcamentistaIndicador.nome_fantasia = BD.readToString(rowDados["nome_fantasia"]);
			orcamentistaIndicador.tipo_estabelecimento = (int)BD.readToInt(rowDados["tipo_estabelecimento"]);
			orcamentistaIndicador.nextel = BD.readToString(rowDados["nextel"]);
			orcamentistaIndicador.email2 = BD.readToString(rowDados["email2"]);
			orcamentistaIndicador.email3 = BD.readToString(rowDados["email3"]);
			orcamentistaIndicador.razao_social_nome_iniciais_em_maiusculas = BD.readToString(rowDados["razao_social_nome_iniciais_em_maiusculas"]);
			orcamentistaIndicador.etq_endereco = BD.readToString(rowDados["etq_endereco"]);
			orcamentistaIndicador.etq_endereco_numero = BD.readToString(rowDados["etq_endereco_numero"]);
			orcamentistaIndicador.etq_endereco_complemento = BD.readToString(rowDados["etq_endereco_complemento"]);
			orcamentistaIndicador.etq_bairro = BD.readToString(rowDados["etq_bairro"]);
			orcamentistaIndicador.etq_cidade = BD.readToString(rowDados["etq_cidade"]);
			orcamentistaIndicador.etq_uf = BD.readToString(rowDados["etq_uf"]);
			orcamentistaIndicador.etq_cep = BD.readToString(rowDados["etq_cep"]);
			orcamentistaIndicador.etq_email = BD.readToString(rowDados["etq_email"]);
			orcamentistaIndicador.etq_ddd_1 = BD.readToString(rowDados["etq_ddd_1"]);
			orcamentistaIndicador.etq_tel_1 = BD.readToString(rowDados["etq_tel_1"]);
			orcamentistaIndicador.etq_ddd_2 = BD.readToString(rowDados["etq_ddd_2"]);
			orcamentistaIndicador.etq_tel_2 = BD.readToString(rowDados["etq_tel_2"]);

			return orcamentistaIndicador;
		}
		#endregion

		#region [ orcamentistaIndicadorResumoPesquisaLoadFromDataRow ]
		public static OrcamentistaIndicadorResumoPesquisa orcamentistaIndicadorResumoPesquisaLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador = new OrcamentistaIndicadorResumoPesquisa();
			#endregion

			orcamentistaIndicador.apelido = BD.readToString(rowDados["apelido"]);
			orcamentistaIndicador.cnpj_cpf = BD.readToString(rowDados["cnpj_cpf"]);
			orcamentistaIndicador.razao_social_nome = BD.readToString(rowDados["razao_social_nome"]);
			orcamentistaIndicador.permite_RA_status = (int)BD.readToInt(rowDados["permite_RA_status"]);
			orcamentistaIndicador.razao_social_nome_iniciais_em_maiusculas = BD.readToString(rowDados["razao_social_nome_iniciais_em_maiusculas"]);
			orcamentistaIndicador.cidade = BD.readToString(rowDados["cidade"]);
			orcamentistaIndicador.uf = BD.readToString(rowDados["uf"]);
			orcamentistaIndicador.loja = BD.readToString(rowDados["loja"]);
			orcamentistaIndicador.vendedor = BD.readToString(rowDados["vendedor"]);
			orcamentistaIndicador.captador = BD.readToString(rowDados["captador"]);
			orcamentistaIndicador.status = BD.readToString(rowDados["status"]);

			return orcamentistaIndicador;
		}
		#endregion

		#region [ getOrcamentistaIndicadorCompletoByApelido ]
		public static OrcamentistaIndicadorCompleto getOrcamentistaIndicadorCompletoByApelido(string apelido, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorCompleto orcamentistaIndicador;
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((apelido ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o identificador do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (apelido = @apelido)";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@apelido", SqlDbType.VarChar, 20);
					cmSelect.Prepare();
					cmSelect.Parameters["@apelido"].Value = (apelido ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o identificador: " + apelido;
						return null;
					}

					orcamentistaIndicador = orcamentistaIndicadorCompletoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return orcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorBasicoByApelido ]
		public static OrcamentistaIndicadorBasico getOrcamentistaIndicadorBasicoByApelido(string apelido, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorBasico orcamentistaIndicador;
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((apelido ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o identificador do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (apelido = @apelido)";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@apelido", SqlDbType.VarChar, 20);
					cmSelect.Prepare();
					cmSelect.Parameters["@apelido"].Value = (apelido ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o identificador: " + apelido;
						return null;
					}

					orcamentistaIndicador = orcamentistaIndicadorBasicoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return orcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorResumoPesquisaByApelido ]
		public static OrcamentistaIndicadorResumoPesquisa getOrcamentistaIndicadorResumoPesquisaByApelido(string apelido, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador;
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((apelido ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o identificador do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (apelido = @apelido)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@apelido", SqlDbType.VarChar, 20);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@apelido"].Value = (apelido ?? "");
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o identificador: " + apelido;
						return null;
					}

					orcamentistaIndicador = orcamentistaIndicadorResumoPesquisaLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return orcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorResumoPesquisaByApelidoParcial ]
		public static List<OrcamentistaIndicadorResumoPesquisa> getOrcamentistaIndicadorResumoPesquisaByApelidoParcial(string apelidoParcial, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador;
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador = new List<OrcamentistaIndicadorResumoPesquisa>();
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((apelidoParcial ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o identificador do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (apelido LIKE @apelidoParcial)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";

					strSql += " ORDER BY" +
								" apelido";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@apelidoParcial", SqlDbType.VarChar, 21);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@apelidoParcial"].Value = (apelidoParcial ?? "") + BD.CARACTER_CURINGA_TODOS;
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						return listaOrcamentistaIndicador;
					}

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						orcamentistaIndicador = orcamentistaIndicadorResumoPesquisaLoadFromDataRow(dtbResultado.Rows[i]);
						listaOrcamentistaIndicador.Add(orcamentistaIndicador);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaOrcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorCompletoPesquisaByCnpjCpf ]
		public static List<OrcamentistaIndicadorCompleto> getOrcamentistaIndicadorCompletoPesquisaByCnpjCpf(string cnpjCpf, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorCompleto orcamentistaIndicador;
			List<OrcamentistaIndicadorCompleto> listaOrcamentistaIndicador = new List<OrcamentistaIndicadorCompleto>();
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((cnpjCpf ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o CNPJ/CPF do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (cnpj_cpf = @cnpj_cpf)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";

					strSql += " ORDER BY" +
								" dt_cadastro";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@cnpj_cpf"].Value = Global.digitos((cnpjCpf ?? ""));
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o CNPJ/CPF: " + Global.formataCnpjCpf(cnpjCpf);
						return null;
					}

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						orcamentistaIndicador = orcamentistaIndicadorCompletoLoadFromDataRow(dtbResultado.Rows[i]);
						listaOrcamentistaIndicador.Add(orcamentistaIndicador);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaOrcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorBasicoPesquisaByCnpjCpf ]
		public static List<OrcamentistaIndicadorBasico> getOrcamentistaIndicadorBasicoPesquisaByCnpjCpf(string cnpjCpf, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorBasico orcamentistaIndicador;
			List<OrcamentistaIndicadorBasico> listaOrcamentistaIndicador = new List<OrcamentistaIndicadorBasico>();
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((cnpjCpf ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o CNPJ/CPF do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (cnpj_cpf = @cnpj_cpf)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";

					strSql += " ORDER BY" +
								" dt_cadastro";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@cnpj_cpf"].Value = Global.digitos((cnpjCpf ?? ""));
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o CNPJ/CPF: " + Global.formataCnpjCpf(cnpjCpf);
						return null;
					}

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						orcamentistaIndicador = orcamentistaIndicadorBasicoLoadFromDataRow(dtbResultado.Rows[i]);
						listaOrcamentistaIndicador.Add(orcamentistaIndicador);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaOrcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorResumoPesquisaByCnpjCpf ]
		public static List<OrcamentistaIndicadorResumoPesquisa> getOrcamentistaIndicadorResumoPesquisaByCnpjCpf(string cnpjCpf, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador;
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador = new List<OrcamentistaIndicadorResumoPesquisa>();
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((cnpjCpf ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o CNPJ/CPF do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (cnpj_cpf = @cnpj_cpf)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";

					strSql += " ORDER BY" +
								" dt_cadastro";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@cnpj_cpf"].Value = Global.digitos((cnpjCpf ?? ""));
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o CNPJ/CPF: " + Global.formataCnpjCpf(cnpjCpf);
						return null;
					}

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						orcamentistaIndicador = orcamentistaIndicadorResumoPesquisaLoadFromDataRow(dtbResultado.Rows[i]);
						listaOrcamentistaIndicador.Add(orcamentistaIndicador);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaOrcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorResumoPesquisaByCnpjCpfParcial ]
		public static List<OrcamentistaIndicadorResumoPesquisa> getOrcamentistaIndicadorResumoPesquisaByCnpjCpfParcial(string cnpjCpfParcial, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador;
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador = new List<OrcamentistaIndicadorResumoPesquisa>();
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((cnpjCpfParcial ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o CNPJ/CPF parcial do orçamentista/indicador!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (cnpj_cpf LIKE @cnpjCpfParcial)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";

					strSql += " ORDER BY" +
								" dt_cadastro";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@cnpjCpfParcial", SqlDbType.VarChar, 16);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@cnpjCpfParcial"].Value = Global.digitos((cnpjCpfParcial ?? "")) + BD.CARACTER_CURINGA_TODOS;
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Não foi localizado nenhum orçamentista/indicador com o CNPJ/CPF parcial: " + Global.digitos(cnpjCpfParcial);
						return null;
					}

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						orcamentistaIndicador = orcamentistaIndicadorResumoPesquisaLoadFromDataRow(dtbResultado.Rows[i]);
						listaOrcamentistaIndicador.Add(orcamentistaIndicador);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaOrcamentistaIndicador;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getOrcamentistaIndicadorResumoPesquisaByNomeParcial ]
		public static List<OrcamentistaIndicadorResumoPesquisa> getOrcamentistaIndicadorResumoPesquisaByNomeParcial(string nomeParcial, string lojaParamDinamico, string lojaParamEstatico, string vendedorParamDinamico, string vendedorParamEstatico, string statusParamDinamico, string statusParamEstatico, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			OrcamentistaIndicadorResumoPesquisa orcamentistaIndicador;
			List<OrcamentistaIndicadorResumoPesquisa> listaOrcamentistaIndicador = new List<OrcamentistaIndicadorResumoPesquisa>();
			SqlConnection cn;
			SqlCommand cmSelect;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				if ((nomeParcial ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o nome parcial a ser pesquisado!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_ORCAMENTISTA_E_INDICADOR" +
							" WHERE" +
								" (razao_social_nome LIKE @nomeParcial)";

					if ((lojaParamDinamico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamDinamico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamDinamico))))";
					if ((lojaParamEstatico ?? "").Trim().Length > 0) strSql += " AND ((loja = @lojaParamEstatico) OR (vendedor IN (SELECT DISTINCT usuario FROM t_USUARIO_X_LOJA WHERE (loja = @lojaParamEstatico))))";
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamDinamico)";
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) strSql += " AND (vendedor = @vendedorParamEstatico)";
					if ((statusParamDinamico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamDinamico)";
					if ((statusParamEstatico ?? "").Trim().Length > 0) strSql += " AND (status = @statusParamEstatico)";

					strSql += " ORDER BY" +
								" apelido";
					#endregion

					#region [ Prepara objeto Command ]
					cmSelect = new SqlCommand();
					cmSelect.Connection = cn;
					cmSelect.CommandText = strSql;
					cmSelect.Parameters.Add("@nomeParcial", SqlDbType.VarChar, 62);
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamDinamico", SqlDbType.VarChar, 3);
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@lojaParamEstatico", SqlDbType.VarChar, 3);
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamDinamico", SqlDbType.VarChar, 10);
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@vendedorParamEstatico", SqlDbType.VarChar, 10);
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamDinamico", SqlDbType.VarChar, 1);
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters.Add("@statusParamEstatico", SqlDbType.VarChar, 1);
					cmSelect.Prepare();
					cmSelect.Parameters["@nomeParcial"].Value = BD.CARACTER_CURINGA_TODOS + (nomeParcial ?? "") + BD.CARACTER_CURINGA_TODOS;
					if ((lojaParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamDinamico"].Value = (lojaParamDinamico ?? "");
					if ((lojaParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@lojaParamEstatico"].Value = (lojaParamEstatico ?? "");
					if ((vendedorParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamDinamico"].Value = (vendedorParamDinamico ?? "");
					if ((vendedorParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@vendedorParamEstatico"].Value = (vendedorParamEstatico ?? "");
					if ((statusParamDinamico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamDinamico"].Value = (statusParamDinamico ?? "");
					if ((statusParamEstatico ?? "").Trim().Length > 0) cmSelect.Parameters["@statusParamEstatico"].Value = (statusParamEstatico ?? "");
					#endregion

					#region [ Executa a consulta ]
					daDataAdapter.SelectCommand = cmSelect;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						return listaOrcamentistaIndicador;
					}

					for (int i = 0; i < dtbResultado.Rows.Count; i++)
					{
						orcamentistaIndicador = orcamentistaIndicadorResumoPesquisaLoadFromDataRow(dtbResultado.Rows[i]);
						listaOrcamentistaIndicador.Add(orcamentistaIndicador);
					}
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return listaOrcamentistaIndicador;
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