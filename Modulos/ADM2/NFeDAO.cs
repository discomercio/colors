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
	public class NFeDAO
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
		public NFeDAO(ref BancoDados bd)
		{
			_bd = bd;
		}
		#endregion

		#region [ Métodos ]

		#region [ getNFeImagemByNF ]
		/// <summary>
		/// Localiza e retorna os dados do registro de t_NFe_IMAGEM através do número da NF
		/// </summary>
		/// <param name="cnpjEmitente">CNPJ do emitente da NF</param>
		/// <param name="serieNF">Número da série da NF</param>
		/// <param name="numeroNF">Número da NF</param>
		/// <returns></returns>
		public List<NFeImagem> getNFeImagemByNF(string cnpjEmitente, int serieNF, int numeroNF)
		{
			#region [ Declarações ]
			string strSql;
			String strListaIdNfeEmitente = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			NFeImagem nfeImagem;
			List<NFeImagem> listaNFeImagem = new List<NFeImagem>();
			#endregion

			#region [ Consistências ]
			if (cnpjEmitente == null) throw new Exception("CNPJ do emitente da NF a ser pesquisada não foi fornecido!");
			if (cnpjEmitente.Length == 0) throw new Exception("CNPJ do emitente da NF a ser pesquisada não foi informado!");
			if (serieNF <= 0) throw new Exception("Nº da série da NF a ser pesquisada não foi informado!");
			if (numeroNF <= 0) throw new Exception("Nº da NF a ser pesquisada não foi informado!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = _bd.criaSqlCommand();
			daDataAdapter = _bd.criaSqlDataAdapter();
			#endregion

			#region [ Identifica o emitente da NF ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						"*" +
					" FROM t_NFe_EMITENTE" +
					" WHERE" +
						" (cnpj = '" + Global.digitos(cnpjEmitente) + "')" +
					" ORDER BY" +
						" id";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Trata o resultado ]
			if (dtbResultado.Rows.Count == 0) throw new Exception("O CNPJ " + Global.formataCnpjCpf(cnpjEmitente) + " NÃO foi localizado como emitente de NFe no sistema!");

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				if (strListaIdNfeEmitente.Length > 0) strListaIdNfeEmitente += ", ";
				strListaIdNfeEmitente += BD.readToInt(rowResultado["id"]).ToString();
			}
			#endregion

			#endregion

			#region [ Pesquisa registro em t_NFe_IMAGEM ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" *" +
					" FROM t_NFe_IMAGEM" +
					" WHERE" +
						" (st_anulado = 0)" +
						" AND (codigo_retorno_NFe_T1 = 1)" +
						" AND (id_nfe_emitente IN (" + strListaIdNfeEmitente + "))" +
						" AND (NFe_Serie_NF = " + serieNF.ToString() + ")" +
						" AND (NFe_numero_NF = " + numeroNF.ToString() + ")" +
					" ORDER BY" +
						" id DESC";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Trata o resultado ]
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];

				#region [ Carrega os dados ]
				nfeImagem = new NFeImagem();

				nfeImagem.id = BD.readToInt(rowResultado["id"]);
				nfeImagem.id_nfe_emitente = BD.readToInt(rowResultado["id_nfe_emitente"]);
				nfeImagem.NFe_serie_NF = BD.readToInt(rowResultado["NFe_serie_NF"]);
				nfeImagem.NFe_numero_NF = BD.readToInt(rowResultado["NFe_numero_NF"]);
				nfeImagem.data = BD.readToDateTime(rowResultado["data"]);
				nfeImagem.data_hora = BD.readToDateTime(rowResultado["data_hora"]);
				nfeImagem.usuario = BD.readToString(rowResultado["usuario"]);
				nfeImagem.pedido = BD.readToString(rowResultado["pedido"]);
				nfeImagem.operacional__email = BD.readToString(rowResultado["operacional__email"]);
				nfeImagem.ide__natOp = BD.readToString(rowResultado["ide__natOp"]);
				nfeImagem.ide__indPag = BD.readToString(rowResultado["ide__indPag"]);
				nfeImagem.ide__serie = BD.readToString(rowResultado["ide__serie"]);
				nfeImagem.ide__nNF = BD.readToString(rowResultado["ide__nNF"]);
				nfeImagem.ide__dEmi = BD.readToString(rowResultado["ide__dEmi"]);
				nfeImagem.ide__dSaiEnt = BD.readToString(rowResultado["ide__dSaiEnt"]);
				nfeImagem.ide__tpNF = BD.readToString(rowResultado["ide__tpNF"]);
				nfeImagem.ide__cMunFG = BD.readToString(rowResultado["ide__cMunFG"]);
				nfeImagem.ide__tpAmb = BD.readToString(rowResultado["ide__tpAmb"]);
				nfeImagem.ide__finNFe = BD.readToString(rowResultado["ide__finNFe"]);
				nfeImagem.ide__IEST = BD.readToString(rowResultado["ide__IEST"]);
				nfeImagem.dest__CNPJ = BD.readToString(rowResultado["dest__CNPJ"]);
				nfeImagem.dest__CPF = BD.readToString(rowResultado["dest__CPF"]);
				nfeImagem.dest__xNome = BD.readToString(rowResultado["dest__xNome"]);
				nfeImagem.dest__xLgr = BD.readToString(rowResultado["dest__xLgr"]);
				nfeImagem.dest__nro = BD.readToString(rowResultado["dest__nro"]);
				nfeImagem.dest__xCpl = BD.readToString(rowResultado["dest__xCpl"]);
				nfeImagem.dest__xBairro = BD.readToString(rowResultado["dest__xBairro"]);
				nfeImagem.dest__cMun = BD.readToString(rowResultado["dest__cMun"]);
				nfeImagem.dest__xMun = BD.readToString(rowResultado["dest__xMun"]);
				nfeImagem.dest__UF = BD.readToString(rowResultado["dest__UF"]);
				nfeImagem.dest__CEP = BD.readToString(rowResultado["dest__CEP"]);
				nfeImagem.dest__cPais = BD.readToString(rowResultado["dest__cPais"]);
				nfeImagem.dest__xPais = BD.readToString(rowResultado["dest__xPais"]);
				nfeImagem.dest__fone = BD.readToString(rowResultado["dest__fone"]);
				nfeImagem.dest__IE = BD.readToString(rowResultado["dest__IE"]);
				nfeImagem.dest__ISUF = BD.readToString(rowResultado["dest__ISUF"]);
				nfeImagem.entrega__CNPJ = BD.readToString(rowResultado["entrega__CNPJ"]);
				nfeImagem.entrega__xLgr = BD.readToString(rowResultado["entrega__xLgr"]);
				nfeImagem.entrega__nro = BD.readToString(rowResultado["entrega__nro"]);
				nfeImagem.entrega__xCpl = BD.readToString(rowResultado["entrega__xCpl"]);
				nfeImagem.entrega__xBairro = BD.readToString(rowResultado["entrega__xBairro"]);
				nfeImagem.entrega__cMun = BD.readToString(rowResultado["entrega__cMun"]);
				nfeImagem.entrega__xMun = BD.readToString(rowResultado["entrega__xMun"]);
				nfeImagem.entrega__UF = BD.readToString(rowResultado["entrega__UF"]);
				nfeImagem.total__vBC = BD.readToString(rowResultado["total__vBC"]);
				nfeImagem.total__vICMS = BD.readToString(rowResultado["total__vICMS"]);
				nfeImagem.total__vBCST = BD.readToString(rowResultado["total__vBCST"]);
				nfeImagem.total__vST = BD.readToString(rowResultado["total__vST"]);
				nfeImagem.total__vProd = BD.readToString(rowResultado["total__vProd"]);
				nfeImagem.total__vFrete = BD.readToString(rowResultado["total__vFrete"]);
				nfeImagem.total__vSeg = BD.readToString(rowResultado["total__vSeg"]);
				nfeImagem.total__vDesc = BD.readToString(rowResultado["total__vDesc"]);
				nfeImagem.total__vII = BD.readToString(rowResultado["total__vII"]);
				nfeImagem.total__vIPI = BD.readToString(rowResultado["total__vIPI"]);
				nfeImagem.total__vPIS = BD.readToString(rowResultado["total__vPIS"]);
				nfeImagem.total__vCOFINS = BD.readToString(rowResultado["total__vCOFINS"]);
				nfeImagem.total__vOutro = BD.readToString(rowResultado["total__vOutro"]);
				nfeImagem.total__vNF = BD.readToString(rowResultado["total__vNF"]);
				nfeImagem.transp__modFrete = BD.readToString(rowResultado["transp__modFrete"]);
				nfeImagem.transporta__CNPJ = BD.readToString(rowResultado["transporta__CNPJ"]);
				nfeImagem.transporta__CPF = BD.readToString(rowResultado["transporta__CPF"]);
				nfeImagem.transporta__xNome = BD.readToString(rowResultado["transporta__xNome"]);
				nfeImagem.transporta__IE = BD.readToString(rowResultado["transporta__IE"]);
				nfeImagem.transporta__xEnder = BD.readToString(rowResultado["transporta__xEnder"]);
				nfeImagem.transporta__xMun = BD.readToString(rowResultado["transporta__xMun"]);
				nfeImagem.transporta__UF = BD.readToString(rowResultado["transporta__UF"]);
				nfeImagem.vol__qVol = BD.readToString(rowResultado["vol__qVol"]);
				nfeImagem.vol__esp = BD.readToString(rowResultado["vol__esp"]);
				nfeImagem.vol__marca = BD.readToString(rowResultado["vol__marca"]);
				nfeImagem.vol__nVol = BD.readToString(rowResultado["vol__nVol"]);
				nfeImagem.vol__pesoL = BD.readToString(rowResultado["vol__pesoL"]);
				nfeImagem.vol__pesoB = BD.readToString(rowResultado["vol__pesoB"]);
				nfeImagem.vol_nLacre = BD.readToString(rowResultado["vol_nLacre"]);
				nfeImagem.infAdic__infAdFisco = BD.readToString(rowResultado["infAdic__infAdFisco"]);
				nfeImagem.infAdic__infCpl = BD.readToString(rowResultado["infAdic__infCpl"]);
				nfeImagem.codigo_retorno_NFe_T1 = BD.readToString(rowResultado["codigo_retorno_NFe_T1"]);
				nfeImagem.msg_retorno_NFe_T1 = BD.readToString(rowResultado["msg_retorno_NFe_T1"]);
				nfeImagem.st_anulado = BD.readToByte(rowResultado["st_anulado"]);
				nfeImagem.dt_anulado = BD.readToDateTime(rowResultado["dt_anulado"]);
				nfeImagem.dt_hr_anulado = BD.readToDateTime(rowResultado["dt_hr_anulado"]);
				nfeImagem.usuario_anulado = BD.readToString(rowResultado["usuario_anulado"]);
				nfeImagem.versao_layout_NFe = BD.readToString(rowResultado["versao_layout_NFe"]);
				nfeImagem.entrega__CPF = BD.readToString(rowResultado["entrega__CPF"]);
				nfeImagem.total__vTotTrib = BD.readToString(rowResultado["total__vTotTrib"]);
				nfeImagem.ide__dEmiUTC = BD.readToString(rowResultado["ide__dEmiUTC"]);
				nfeImagem.ide__idDest = BD.readToString(rowResultado["ide__idDest"]);
				nfeImagem.ide__indFinal = BD.readToString(rowResultado["ide__indFinal"]);
				nfeImagem.ide__indPres = BD.readToString(rowResultado["ide__indPres"]);
				nfeImagem.dest__idEstrangeiro = BD.readToString(rowResultado["dest__idEstrangeiro"]);
				nfeImagem.dest__indIEDest = BD.readToString(rowResultado["dest__indIEDest"]);
				nfeImagem.total__vICMSDeson = BD.readToString(rowResultado["total__vICMSDeson"]);
				nfeImagem.dest__email = BD.readToString(rowResultado["dest__email"]);
				nfeImagem.total__vFCPUFDest = BD.readToString(rowResultado["total__vFCPUFDest"]);
				nfeImagem.total__vICMSUFDest = BD.readToString(rowResultado["total__vICMSUFDest"]);
				nfeImagem.total__vICMSUFRemet = BD.readToString(rowResultado["total__vICMSUFRemet"]);
				nfeImagem.total__vFCP = BD.readToString(rowResultado["total__vFCP"]);
				nfeImagem.total__vFCPST = BD.readToString(rowResultado["total__vFCPST"]);
				nfeImagem.total__vFCPSTRet = BD.readToString(rowResultado["total__vFCPSTRet"]);
				nfeImagem.total__vIPIDevol = BD.readToString(rowResultado["total__vIPIDevol"]);

				listaNFeImagem.Add(nfeImagem);
				#endregion
			}
			#endregion

			#endregion

			return listaNFeImagem;
		}
		#endregion

		#endregion
	}
}
