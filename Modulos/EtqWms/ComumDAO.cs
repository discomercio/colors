#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
#endregion

namespace EtqWms
{
	#region [ ComumDAO ]
	class ComumDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta;
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
		static ComumDAO()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta ]
			strSql = "UPDATE t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO SET " +
						"etiqueta_impressao_status = 1, " +
						"etiqueta_impressao_qtde_impressoes = etiqueta_impressao_qtde_impressoes + 1, " +
						"etiqueta_impressao_primeira_vez_data = CASE etiqueta_impressao_status WHEN 0 THEN " + Global.sqlMontaGetdateSomenteData() + " ELSE etiqueta_impressao_primeira_vez_data END, " +
						"etiqueta_impressao_primeira_vez_data_hora = CASE etiqueta_impressao_status WHEN 0 THEN getdate() ELSE etiqueta_impressao_primeira_vez_data_hora END, " +
						"etiqueta_impressao_primeira_vez_usuario = CASE etiqueta_impressao_status WHEN 0 THEN @usuario ELSE etiqueta_impressao_primeira_vez_usuario END, " +
						"etiqueta_impressao_ultima_vez_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"etiqueta_impressao_ultima_vez_data_hora = getdate(), " +
						"etiqueta_impressao_ultima_vez_usuario = @usuario" +
					" WHERE (id = @id)";
			cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta = BD.criaSqlCommand();
			cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta.CommandText = strSql;
			cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta.Prepare();
			#endregion
		}
		#endregion

		#region [ atualizaWmsEtqN1SepZonaRelEmissaoCompleta ]
		public static bool atualizaWmsEtqN1SepZonaRelEmissaoCompleta(int id, string usuario, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "atualizaWmsEtqN1SepZonaRelEmissaoCompleta()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta.Parameters["@id"].Value = id;
				cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta.Parameters["@usuario"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWmsEtqN1SepZonaRelEmissaoCompleta);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString());
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
					strMsgErro = "Falha ao tentar atualizar o registro do relatório NSU=" + id.ToString() + " na tabela t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ getWmsEtqN1SepZonaRel ]
		public static WmsEtqN1SepZonaRel getWmsEtqN1SepZonaRel(int id)
		{
			#region [ Declarações ]
			String strSql;
			WmsEtqN1SepZonaRel wmsEtqN1SepZonaRel = new WmsEtqN1SepZonaRel();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id == 0) throw new Exception("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém dados ]
			strSql = "SELECT " +
						"*" +
					" FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO" +
					" WHERE" +
						" (id = " + id.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) throw new Exception("Registro id=" + id.ToString() + " não localizado na tabela t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO!!");

			rowResultado = dtbResultado.Rows[0];

			wmsEtqN1SepZonaRel.id = BD.readToInt(rowResultado["id"]);
			wmsEtqN1SepZonaRel.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
			wmsEtqN1SepZonaRel.dt_hr_cadastro = BD.readToDateTime(rowResultado["dt_hr_cadastro"]);
			wmsEtqN1SepZonaRel.dt_emissao = BD.readToDateTime(rowResultado["dt_emissao"]);
			wmsEtqN1SepZonaRel.dt_hr_emissao = BD.readToDateTime(rowResultado["dt_hr_emissao"]);
			wmsEtqN1SepZonaRel.usuario = BD.readToString(rowResultado["usuario"]);
			wmsEtqN1SepZonaRel.filtro_dt_inicio = BD.readToString(rowResultado["filtro_dt_inicio"]);
			wmsEtqN1SepZonaRel.filtro_dt_termino = BD.readToString(rowResultado["filtro_dt_termino"]);
			wmsEtqN1SepZonaRel.filtro_NFe_emitida = BD.readToString(rowResultado["filtro_NFe_emitida"]);
			wmsEtqN1SepZonaRel.filtro_transportadora = BD.readToString(rowResultado["filtro_transportadora"]);
			wmsEtqN1SepZonaRel.filtro_qtde_max_pedidos = BD.readToString(rowResultado["filtro_qtde_max_pedidos"]);
			wmsEtqN1SepZonaRel.filtro_qtde_disponivel_pedidos = BD.readToString(rowResultado["filtro_qtde_disponivel_pedidos"]);
			wmsEtqN1SepZonaRel.lista_zonas_cadastradas = BD.readToString(rowResultado["lista_zonas_cadastradas"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_status = BD.readToByte(rowResultado["etiqueta_impressao_status"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_qtde_impressoes = BD.readToInt(rowResultado["etiqueta_impressao_qtde_impressoes"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_primeira_vez_data = BD.readToDateTime(rowResultado["etiqueta_impressao_primeira_vez_data"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_primeira_vez_data_hora = BD.readToDateTime(rowResultado["etiqueta_impressao_primeira_vez_data_hora"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_primeira_vez_usuario = BD.readToString(rowResultado["etiqueta_impressao_primeira_vez_usuario"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_ultima_vez_data = BD.readToDateTime(rowResultado["etiqueta_impressao_ultima_vez_data"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_ultima_vez_data_hora = BD.readToDateTime(rowResultado["etiqueta_impressao_ultima_vez_data_hora"]);
			wmsEtqN1SepZonaRel.etiqueta_impressao_ultima_vez_usuario = BD.readToString(rowResultado["etiqueta_impressao_ultima_vez_usuario"]);
			#endregion

			return wmsEtqN1SepZonaRel;
		}
		#endregion
	}
	#endregion
}
