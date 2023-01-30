#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace ADM2
{
	public class GeralDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
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
		public GeralDAO(ref BancoDados bd)
		{
			_bd = bd;
		}
		#endregion

		#region [ codigoDescricaoLoadFromDataRow ]
		public CodigoDescricao codigoDescricaoLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			CodigoDescricao codigoDescricao = new CodigoDescricao();
			#endregion

			codigoDescricao.grupo = BD.readToString(rowDados["grupo"]);
			codigoDescricao.codigo = BD.readToString(rowDados["codigo"]);
			codigoDescricao.ordenacao = BD.readToInt(rowDados["ordenacao"]);
			codigoDescricao.st_inativo = BD.readToByte(rowDados["st_inativo"]);
			codigoDescricao.descricao = BD.readToString(rowDados["descricao"]);
			codigoDescricao.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			codigoDescricao.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			codigoDescricao.dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["dt_hr_ult_atualizacao"]);
			codigoDescricao.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);
			codigoDescricao.st_possui_sub_codigo = BD.readToByte(rowDados["st_possui_sub_codigo"]);
			codigoDescricao.st_eh_sub_codigo = BD.readToByte(rowDados["st_eh_sub_codigo"]);
			codigoDescricao.grupo_pai = BD.readToString(rowDados["grupo_pai"]);
			codigoDescricao.codigo_pai = BD.readToString(rowDados["codigo_pai"]);
			codigoDescricao.lojas_habilitadas = BD.readToString(rowDados["lojas_habilitadas"]);
			codigoDescricao.parametro_1_campo_flag = BD.readToByte(rowDados["parametro_1_campo_flag"]);
			codigoDescricao.parametro_2_campo_flag = BD.readToByte(rowDados["parametro_2_campo_flag"]);
			codigoDescricao.parametro_3_campo_flag = BD.readToByte(rowDados["parametro_3_campo_flag"]);
			codigoDescricao.parametro_4_campo_flag = BD.readToByte(rowDados["parametro_4_campo_flag"]);
			codigoDescricao.parametro_5_campo_flag = BD.readToByte(rowDados["parametro_5_campo_flag"]);
			codigoDescricao.parametro_campo_inteiro = BD.readToInt(rowDados["parametro_campo_inteiro"]);
			codigoDescricao.parametro_campo_monetario = BD.readToDecimal(rowDados["parametro_campo_monetario"]);
			codigoDescricao.parametro_campo_real = BD.readToSingle(rowDados["parametro_campo_real"]);
			codigoDescricao.parametro_campo_data = BD.readToDateTime(rowDados["parametro_campo_data"]);
			codigoDescricao.parametro_campo_texto = BD.readToString(rowDados["parametro_campo_texto"]);
			codigoDescricao.parametro_2_campo_texto = BD.readToString(rowDados["parametro_2_campo_texto"]);
			codigoDescricao.parametro_3_campo_texto = BD.readToString(rowDados["parametro_3_campo_texto"]);
			codigoDescricao.parametro_4_campo_texto = BD.readToString(rowDados["parametro_4_campo_texto"]);
			codigoDescricao.descricao_parametro = BD.readToString(rowDados["descricao_parametro"]);

			return codigoDescricao;
		}
		#endregion

		#region [ getCodigoDescricao ]
		public CodigoDescricao getCodigoDescricao(string grupo, string codigo, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			CodigoDescricao codigoDescricao;
			#endregion

			msg_erro = "";
			try
			{
				if ((grupo ?? "").Trim().Length == 0)
				{
					msg_erro = "Identificação do grupo não foi informado!";
					return null;
				}

				if ((codigo ?? "").Trim().Length == 0)
				{
					msg_erro = "Identificação do código não foi informado!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_CODIGO_DESCRICAO" +
						" WHERE" +
							" (grupo = '" + grupo + "')" +
							" AND (codigo = '" + codigo + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0)
				{
					msg_erro = "Código não encontrado!";
					return null;
				}

				codigoDescricao = codigoDescricaoLoadFromDataRow(dtbResultado.Rows[0]);

				return codigoDescricao;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getCodigoDescricaoByGrupo ]
		public List<CodigoDescricao> getCodigoDescricaoByGrupo(string grupo, Global.eFiltroFlagStInativo st_inativo, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			string strWhere;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			CodigoDescricao codigoDescricao;
			List<CodigoDescricao> listaCodigoDescricao = new List<CodigoDescricao>();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta cláusula Where ]
				strWhere = "";
				if ((grupo ?? "").Trim().Length > 0)
				{
					if (strWhere.Length > 0) strWhere += " AND";
					strWhere += " (grupo = '" + grupo + "')";
				}
				if (st_inativo != Global.eFiltroFlagStInativo.FLAG_IGNORADO)
				{
					if (strWhere.Length > 0) strWhere += " AND";
					strWhere += " (st_inativo = " + st_inativo.ToString() + ")";
				}

				if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;
				#endregion

				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_CODIGO_DESCRICAO" +
						strWhere +
						" ORDER BY" +
							" grupo," +
							" ordenacao";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					codigoDescricao = codigoDescricaoLoadFromDataRow(dtbResultado.Rows[i]);
					listaCodigoDescricao.Add(codigoDescricao);
				}

				return listaCodigoDescricao;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ NFeEmitenteLoadFromDataRow ]
		public NFeEmitente NFeEmitenteLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			NFeEmitente emitente = new NFeEmitente();
			#endregion

			emitente.id = BD.readToInt(rowDados["id"]);
			emitente.id_boleto_cedente = BD.readToInt(rowDados["id_boleto_cedente"]);
			emitente.st_ativo = BD.readToByte(rowDados["st_ativo"]);
			emitente.apelido = BD.readToString(rowDados["apelido"]);
			emitente.cnpj = BD.readToString(rowDados["cnpj"]);
			emitente.razao_social = BD.readToString(rowDados["razao_social"]);
			emitente.endereco = BD.readToString(rowDados["endereco"]);
			emitente.endereco_numero = BD.readToString(rowDados["endereco_numero"]);
			emitente.endereco_complemento = BD.readToString(rowDados["endereco_complemento"]);
			emitente.bairro = BD.readToString(rowDados["bairro"]);
			emitente.cidade = BD.readToString(rowDados["cidade"]);
			emitente.uf = BD.readToString(rowDados["uf"]);
			emitente.cep = BD.readToString(rowDados["cep"]);
			emitente.NFe_st_emitente_padrao = BD.readToByte(rowDados["NFe_st_emitente_padrao"]);
			emitente.NFe_serie_NF = BD.readToInt(rowDados["NFe_serie_NF"]);
			emitente.NFe_numero_NF = BD.readToInt(rowDados["NFe_numero_NF"]);
			emitente.NFe_T1_servidor_BD = BD.readToString(rowDados["NFe_T1_servidor_BD"]);
			emitente.NFe_T1_nome_BD = BD.readToString(rowDados["NFe_T1_nome_BD"]);
			emitente.NFe_T1_usuario_BD = BD.readToString(rowDados["NFe_T1_usuario_BD"]);
			emitente.NFe_T1_senha_BD = BD.readToString(rowDados["NFe_T1_senha_BD"]);
			emitente.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			emitente.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			emitente.usuario_cadastro = BD.readToString(rowDados["usuario_cadastro"]);
			emitente.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			emitente.dt_hr_ult_atualizacao = BD.readToDateTime(rowDados["dt_hr_ult_atualizacao"]);
			emitente.usuario_ult_atualizacao = BD.readToString(rowDados["usuario_ult_atualizacao"]);

			return emitente;
		}
		#endregion

		#region [ getNFeEmitenteById ]
		public NFeEmitente getNFeEmitenteById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			NFeEmitente emitente;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"e.*, " +
							"n.NFe_Numero_NF, " +
							"n.NFe_Serie_NF " +
						" FROM t_NFe_EMITENTE e" +
						" INNER JOIN t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" +
						" WHERE" +
							" (e.id = " + id.ToString() + ")";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

				emitente = NFeEmitenteLoadFromDataRow(dtbResultado.Rows[0]);

				return emitente;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getNFeEmitenteByCnpj ]
		public List<NFeEmitente> getNFeEmitenteByCnpj(string cnpj, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			NFeEmitente emitente;
			List<NFeEmitente> listaEmitente = new List<NFeEmitente>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistência ]
				if ((cnpj ?? "").Trim().Length == 0)
				{
					msg_erro = "CNPJ não informado!";
					return listaEmitente;
				}
				#endregion

				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"e.*, " +
							"n.NFe_Numero_NF, " +
							"n.NFe_Serie_NF " +
						" FROM t_NFe_EMITENTE e" +
						" INNER JOIN t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" +
						" WHERE" +
							" (e.cnpj = '" + Global.digitos(cnpj) + "')" +
						" ORDER BY" +
							" e.id";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return listaEmitente;

				foreach (DataRow row in dtbResultado.Rows)
				{
					emitente = NFeEmitenteLoadFromDataRow(row);
					listaEmitente.Add(emitente);
				}

				return listaEmitente;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getAllNFeEmitenteByCnpj ]
		public List<NFeEmitente> getAllNFeEmitenteByCnpj(out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			NFeEmitente emitente;
			List<NFeEmitente> listaEmitente = new List<NFeEmitente>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"e.*, " +
							"n.NFe_Numero_NF, " +
							"n.NFe_Serie_NF " +
						" FROM t_NFe_EMITENTE e" +
						" INNER JOIN t_NFe_EMITENTE_NUMERACAO n ON e.cnpj = n.cnpj" +
						" ORDER BY" +
							" e.id";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return listaEmitente;

				foreach (DataRow row in dtbResultado.Rows)
				{
					emitente = NFeEmitenteLoadFromDataRow(row);
					listaEmitente.Add(emitente);
				}

				return listaEmitente;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ TransportadoraLoadFromDataRow ]
		public Transportadora TransportadoraLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			Transportadora transportadora = new Transportadora();
			#endregion

			transportadora.id = BD.readToString(rowDados["id"]);
			transportadora.cnpj = BD.readToString(rowDados["cnpj"]);
			transportadora.ie = BD.readToString(rowDados["ie"]);
			transportadora.nome = BD.readToString(rowDados["nome"]);
			transportadora.razao_social = BD.readToString(rowDados["razao_social"]);
			transportadora.endereco = BD.readToString(rowDados["endereco"]);
			transportadora.bairro = BD.readToString(rowDados["bairro"]);
			transportadora.cidade = BD.readToString(rowDados["cidade"]);
			transportadora.uf = BD.readToString(rowDados["uf"]);
			transportadora.cep = BD.readToString(rowDados["cep"]);
			transportadora.ddd = BD.readToString(rowDados["ddd"]);
			transportadora.telefone = BD.readToString(rowDados["telefone"]);
			transportadora.fax = BD.readToString(rowDados["fax"]);
			transportadora.contato = BD.readToString(rowDados["contato"]);
			transportadora.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			transportadora.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			transportadora.endereco_numero = BD.readToString(rowDados["endereco_numero"]);
			transportadora.endereco_complemento = BD.readToString(rowDados["endereco_complemento"]);
			transportadora.email = BD.readToString(rowDados["email"]);
			transportadora.email2 = BD.readToString(rowDados["email2"]);

			return transportadora;
		}
		#endregion

		#region [ getTransportadoraById ]
		public Transportadora getTransportadoraById(string id, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			Transportadora transportadora;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"*" +
						" FROM t_TRANSPORTADORA" +
						" WHERE" +
							" (id = '" + (id ?? "").Trim() + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

				transportadora = TransportadoraLoadFromDataRow(dtbResultado.Rows[0]);

				return transportadora;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getAllTransportadora ]
		public List<Transportadora> getAllTransportadora(out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			List<Transportadora> listaTransportadora = new List<Transportadora>();
			Transportadora transportadora;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = _bd.criaSqlCommand();
				daDataAdapter = _bd.criaSqlDataAdapter();
				#endregion

				#region [ Monta SQL ]
				strSql = "SELECT " +
							"*" +
						" FROM t_TRANSPORTADORA" +
						" ORDER BY" +
							" id";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return listaTransportadora;

				foreach (DataRow row in dtbResultado.Rows)
				{
					transportadora = TransportadoraLoadFromDataRow(row);
					listaTransportadora.Add(transportadora);
				}

				return listaTransportadora;
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
