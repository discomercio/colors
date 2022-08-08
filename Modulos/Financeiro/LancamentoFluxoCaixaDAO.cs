#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace Financeiro
{
	class LancamentoFluxoCaixaDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmLancamentoInsert;
		private static SqlCommand cmLancamentoInsertDevidoBoletoOcorrencia02;
		private static SqlCommand cmLancamentoUpdate;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia02;
		private static SqlCommand cmLancamentoDelete;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia06;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia09;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia10;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia12;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia13;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia14;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia15;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia16;
		private static SqlCommand cmLancamentoInsertDevidoBoletoOcorrencia17;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia22;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia23;
		private static SqlCommand cmLancamentoUpdateDevidoBoletoOcorrencia34;
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
		static LancamentoFluxoCaixaDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmLancamentoInsert ]
			strSql = "INSERT INTO t_FIN_FLUXO_CAIXA (" +
						"id, " +
						"id_conta_corrente, " +
						"id_plano_contas_empresa, " +
						"id_plano_contas_grupo, " +
						"id_plano_contas_conta, " +
						"natureza, " +
						"dt_competencia, " +
						"valor, " +
						"descricao, " +
						"cnpj_cpf, " +
						"numero_NF, " +
						"tipo_cadastro, " +
						"editado_manual, " +
						"dt_cadastro, " +
						"dt_hr_cadastro, " +
						"usuario_cadastro, " +
						"dt_ult_atualizacao, " +
						"dt_hr_ult_atualizacao, " +
						"usuario_ult_atualizacao, " +
						"dt_mes_competencia" +
					") VALUES (" +
						"@id, " +
						"@id_conta_corrente, " +
						"@id_plano_contas_empresa, " +
						"@id_plano_contas_grupo, " +
						"@id_plano_contas_conta, " +
						"@natureza, " +
						"@dt_competencia, " +
						"@valor, " +
						"@descricao, " +
						"@cnpj_cpf, " +
						"@numero_NF, " +
						"@tipo_cadastro, " +
						"@editado_manual, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_cadastro, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_ult_atualizacao, " +
						"@dt_mes_competencia" +
					")";
			cmLancamentoInsert = BD.criaSqlCommand();
			cmLancamentoInsert.CommandText = strSql;
			cmLancamentoInsert.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoInsert.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
			cmLancamentoInsert.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
			cmLancamentoInsert.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
			cmLancamentoInsert.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
			cmLancamentoInsert.Parameters.Add("@natureza", SqlDbType.Char, 1);
			cmLancamentoInsert.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoInsert.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoInsert.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
			cmLancamentoInsert.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
			cmLancamentoInsert.Parameters.Add("@numero_NF", SqlDbType.Int);
			cmLancamentoInsert.Parameters.Add("@tipo_cadastro", SqlDbType.Char, 1);
			cmLancamentoInsert.Parameters.Add("@editado_manual", SqlDbType.Char, 1);
			cmLancamentoInsert.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmLancamentoInsert.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
            cmLancamentoInsert.Parameters.Add("@dt_mes_competencia", SqlDbType.VarChar, 10);
			cmLancamentoInsert.Prepare();
			#endregion

			#region [ cmLancamentoInsertDevidoBoletoOcorrencia02 ]
			strSql = "INSERT INTO t_FIN_FLUXO_CAIXA (" +
						"id, " +
						"id_conta_corrente, " +
						"id_plano_contas_empresa, " +
						"id_plano_contas_grupo, " +
						"id_plano_contas_conta, " +
						"natureza, " +
						"dt_competencia, " +
						"valor, " +
						"descricao, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"ctrl_pagto_status, " +
						"id_boleto_cedente, " +
						"id_cliente, " +
						"cnpj_cpf, " +
						"tipo_cadastro, " +
						"editado_manual, " +
						"dt_cadastro, " +
						"dt_hr_cadastro, " +
						"usuario_cadastro, " +
						"dt_ult_atualizacao, " +
						"dt_hr_ult_atualizacao, " +
						"usuario_ult_atualizacao, " +
						"st_confirmacao_pendente" +
					") VALUES (" +
						"@id, " +
						"@id_conta_corrente, " +
						"@id_plano_contas_empresa, " +
						"@id_plano_contas_grupo, " +
						"@id_plano_contas_conta, " +
						"@natureza, " +
						"@dt_competencia, " +
						"@valor, " +
						"@descricao, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						((byte)Global.Cte.FIN.eCtrlPagtoStatus.CADASTRADO_INICIAL).ToString() + ", " +
						"@id_boleto_cedente, " +
						"@id_cliente, " +
						"@cnpj_cpf, " +
						"@tipo_cadastro, " +
						"@editado_manual, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_cadastro, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_ult_atualizacao, " +
						"@st_confirmacao_pendente" +
					")";
			cmLancamentoInsertDevidoBoletoOcorrencia02 = BD.criaSqlCommand();
			cmLancamentoInsertDevidoBoletoOcorrencia02.CommandText = strSql;
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@natureza", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@tipo_cadastro", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@editado_manual", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters.Add("@st_confirmacao_pendente", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia02.Prepare();
			#endregion

			#region [ cmLancamentoUpdate ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"id_conta_corrente = @id_conta_corrente, " +
						"id_plano_contas_empresa = @id_plano_contas_empresa, " +
						"id_plano_contas_grupo = @id_plano_contas_grupo, " +
						"id_plano_contas_conta = @id_plano_contas_conta, " +
						"natureza = @natureza, " +
						"st_sem_efeito = @st_sem_efeito, " +
						"st_confirmacao_pendente = @st_confirmacao_pendente, " +
						"dt_competencia = @dt_competencia, " +
						"dt_mes_competencia = @dt_mes_competencia, " +
						"valor = @valor, " +
						"descricao = @descricao, " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"cnpj_cpf = @cnpj_cpf, " +
						"numero_NF = @numero_NF, " +
						"editado_manual = @editado_manual, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE (id = @id)";
			cmLancamentoUpdate = BD.criaSqlCommand();
			cmLancamentoUpdate.CommandText = strSql;
			cmLancamentoUpdate.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdate.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
			cmLancamentoUpdate.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
			cmLancamentoUpdate.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
			cmLancamentoUpdate.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
			cmLancamentoUpdate.Parameters.Add("@natureza", SqlDbType.Char, 1);
			cmLancamentoUpdate.Parameters.Add("@st_sem_efeito", SqlDbType.TinyInt);
			cmLancamentoUpdate.Parameters.Add("@st_confirmacao_pendente", SqlDbType.TinyInt);
			cmLancamentoUpdate.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
            cmLancamentoUpdate.Parameters.Add("@dt_mes_competencia", SqlDbType.VarChar, 10);
			cmLancamentoUpdate.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoUpdate.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
			cmLancamentoUpdate.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
			cmLancamentoUpdate.Parameters.Add("@numero_NF", SqlDbType.Int);
			cmLancamentoUpdate.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdate.Parameters.Add("@editado_manual", SqlDbType.Char, 1);
			cmLancamentoUpdate.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdate.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia02 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"st_sem_efeito = @st_sem_efeito, " +
						"dt_competencia = @dt_competencia, " +
						"valor = @valor, " +
						"descricao = @descricao, " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE (id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia02 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia02.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@st_sem_efeito", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia02.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia06 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"st_sem_efeito = @st_sem_efeito, " +
						"dt_competencia = @dt_competencia, " +
						"valor = @valor, " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"st_confirmacao_pendente = @st_confirmacao_pendente, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE (id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia06 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia06.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@st_sem_efeito", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@st_confirmacao_pendente", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia06.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia09 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"st_boleto_baixado = @st_boleto_baixado, " +
						"dt_ocorrencia_banco_boleto_baixado = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_baixado") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia09 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia09.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@st_boleto_baixado", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@dt_ocorrencia_banco_boleto_baixado", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia09.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia10 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"st_boleto_baixado = @st_boleto_baixado, " +
						"dt_ocorrencia_banco_boleto_baixado = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_baixado") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia10 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia10.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@st_boleto_baixado", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@dt_ocorrencia_banco_boleto_baixado", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia10.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia12 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"valor = @valor, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id) " +
						"AND (@valor >= 0)";
			cmLancamentoUpdateDevidoBoletoOcorrencia12 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia12.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia12.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia13 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"valor = @valor, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id) " +
						"AND (@valor >= 0) " +
						"AND (ctrl_pagto_status <> " + ((byte)Global.Cte.FIN.eCtrlPagtoStatus.PAGO).ToString() + ")";
			cmLancamentoUpdateDevidoBoletoOcorrencia13 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia13.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia13.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia14 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"dt_competencia = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_competencia") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia14 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia14.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia14.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia15 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"st_sem_efeito = @st_sem_efeito, " +
						"dt_competencia = @dt_competencia, " +
						"valor = @valor, " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"st_confirmacao_pendente = @st_confirmacao_pendente, " +
						"st_boleto_ocorrencia_15 = @st_boleto_ocorrencia_15, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_15 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_15") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE (id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia15 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia15.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@st_sem_efeito", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@st_confirmacao_pendente", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@st_boleto_ocorrencia_15", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_15", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia15.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia16 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"st_sem_efeito = @st_sem_efeito, " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"st_boleto_pago_cheque = @st_boleto_pago_cheque, " +
						"dt_ocorrencia_banco_boleto_pago_cheque = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_pago_cheque") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia16 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia16.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@st_sem_efeito", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@st_boleto_pago_cheque", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@dt_ocorrencia_banco_boleto_pago_cheque", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia16.Prepare();
			#endregion

			#region [ cmLancamentoInsertDevidoBoletoOcorrencia17 ]
			strSql = "INSERT INTO t_FIN_FLUXO_CAIXA (" +
						"id, " +
						"id_conta_corrente, " +
						"id_plano_contas_empresa, " +
						"id_plano_contas_grupo, " +
						"id_plano_contas_conta, " +
						"natureza, " +
						"dt_competencia, " +
						"valor, " +
						"descricao, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"ctrl_pagto_status, " +
						"id_boleto_cedente, " +
						"id_cliente, " +
						"cnpj_cpf, " +
						"tipo_cadastro, " +
						"editado_manual, " +
						"dt_cadastro, " +
						"dt_hr_cadastro, " +
						"usuario_cadastro, " +
						"dt_ult_atualizacao, " +
						"dt_hr_ult_atualizacao, " +
						"usuario_ult_atualizacao, " +
						"st_confirmacao_pendente, " +
						"st_boleto_ocorrencia_17, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_17" +
					") VALUES (" +
						"@id, " +
						"@id_conta_corrente, " +
						"@id_plano_contas_empresa, " +
						"@id_plano_contas_grupo, " +
						"@id_plano_contas_conta, " +
						"@natureza, " +
						"@dt_competencia, " +
						"@valor, " +
						"@descricao, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						"@ctrl_pagto_status, " +
						"@id_boleto_cedente, " +
						"@id_cliente, " +
						"@cnpj_cpf, " +
						"@tipo_cadastro, " +
						"@editado_manual, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_cadastro, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_ult_atualizacao, " +
						"@st_confirmacao_pendente, " +
						"@st_boleto_ocorrencia_17, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_17") +
					")";
			cmLancamentoInsertDevidoBoletoOcorrencia17 = BD.criaSqlCommand();
			cmLancamentoInsertDevidoBoletoOcorrencia17.CommandText = strSql;
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@natureza", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@tipo_cadastro", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@editado_manual", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@st_confirmacao_pendente", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@st_boleto_ocorrencia_17", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_17", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoOcorrencia17.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia22 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"ctrl_pagto_status = @ctrl_pagto_status, " +
						"st_boleto_pago_cheque = @st_boleto_pago_cheque, " +
						"dt_ocorrencia_banco_boleto_pago_cheque = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_pago_cheque") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia22 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia22.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@st_boleto_pago_cheque", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@dt_ocorrencia_banco_boleto_pago_cheque", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia22.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia23 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"st_boleto_ocorrencia_23 = @st_boleto_ocorrencia_23, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_23 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_23") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia23 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia23.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@st_boleto_ocorrencia_23", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_23", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia23.Prepare();
			#endregion

			#region [ cmLancamentoUpdateDevidoBoletoOcorrencia34 ]
			strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						"st_boleto_ocorrencia_34 = @st_boleto_ocorrencia_34, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_34 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_34") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE " +
						"(id = @id)";
			cmLancamentoUpdateDevidoBoletoOcorrencia34 = BD.criaSqlCommand();
			cmLancamentoUpdateDevidoBoletoOcorrencia34.CommandText = strSql;
			cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@st_boleto_ocorrencia_34", SqlDbType.TinyInt);
			cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_34", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoUpdateDevidoBoletoOcorrencia34.Prepare();
			#endregion

			#region [ cmLancamentoDelete ]
			strSql = "DELETE FROM t_FIN_FLUXO_CAIXA WHERE (id = @id)";
			cmLancamentoDelete = BD.criaSqlCommand();
			cmLancamentoDelete.CommandText = strSql;
			cmLancamentoDelete.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoDelete.Prepare();
			#endregion
		}
		#endregion

		#region [ obtemRegistroLancamentoBD ]
		/// <summary>
		/// Retorna o DataTable contendo o registro do lançamento de fluxo de caixa especificado pelo parâmetro
		/// </summary>
		/// <param name="id">
		/// Parâmetro especificando o registro que deve ser obtido
		/// </param>
		/// <returns>
		/// Retorna o DataTable contendo o registro especificado
		/// </returns>
		private static DsDataSource.DtbFinFluxoCaixaDataTable obtemRegistroLancamentoBD(int id)
		{
			String strSql;
			String strWhere = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa = new DsDataSource.DtbFinFluxoCaixaDataTable();

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();

			strWhere = " (id = " + id.ToString() + ")";
			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_FLUXO_CAIXA" +
					strWhere;
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinFluxoCaixa);

			return dtbFinFluxoCaixa;
		}
		#endregion

		#region [ obtemRegistroLancamentoByCtrlPagtoIdParcela ]
		/// <summary>
		/// Retorna o DataTable contendo o registro do lançamento de fluxo de caixa, sendo que a
		/// pesquisa é feita através do campo ctrl_pagto_id_parcela + ctrl_pagto_modulo, ou seja,
		/// está sendo localizado o lançamento de fluxo de caixa associado a um dos boletos, cheques
		/// ou parcelas do Visa.
		/// </summary>
		/// <param name="ctrlPagtoIdParcela">
		/// Nº identificação do registro no módulo de controle que gerou o lançamento no fluxo de caixa (módulos
		/// de controle: boleto, cheque, Visa)
		/// </param>
		/// <param name="ctrlPagtoModulo">
		/// Código de identificação do módulo de controle
		///		1 = Boleto
		///		2 = Cheque
		///		3 = Visa
		/// </param>
		/// <returns>
		/// Retorna o DataTable contendo o registro especificado
		/// </returns>
		public static DsDataSource.DtbFinFluxoCaixaDataTable obtemRegistroLancamentoByCtrlPagtoIdParcela(int ctrlPagtoIdParcela, byte ctrPagtoModulo)
		{
			String strSql;
			String strWhere = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa = new DsDataSource.DtbFinFluxoCaixaDataTable();

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();

			strWhere = " (ctrl_pagto_id_parcela = " + ctrlPagtoIdParcela.ToString() + ")" +
					   " AND (ctrl_pagto_modulo = " + ctrPagtoModulo.ToString() + ")";
			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_FLUXO_CAIXA" +
					strWhere;
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinFluxoCaixa);

			return dtbFinFluxoCaixa;
		}
		#endregion

		#region [ getLancamentoFluxoCaixa ]
		/// <summary>
		/// Retorna um objeto representando o lançamento do fluxo de caixa contendo os dados lidos do BD
		/// </summary>
		/// <param name="id">
		/// Identificação do registro do fluxo de caixa
		/// </param>
		/// <returns>
		/// Retorna um objeto LancamentoFluxoCaixa com os dados do lançamento
		/// </returns>
		public static LancamentoFluxoCaixa getLancamentoFluxoCaixa(int id)
		{
			LancamentoFluxoCaixa lancamento = new LancamentoFluxoCaixa();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixa;

			if (id == 0) throw new FinanceiroException("O identificador do registro não foi informado!!");

			dtbFinFluxoCaixa = obtemRegistroLancamentoBD(id);

			if (dtbFinFluxoCaixa.Rows.Count == 0) throw new FinanceiroException("Registro id=" + id.ToString() + " não localizado na tabela de fluxo de caixa!!");

			rowFinFluxoCaixa = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixa.Rows[0];

			lancamento.id = rowFinFluxoCaixa.id;
			lancamento.id_conta_corrente = rowFinFluxoCaixa.id_conta_corrente;
			lancamento.id_plano_contas_empresa = rowFinFluxoCaixa.id_plano_contas_empresa;
			lancamento.id_plano_contas_grupo = rowFinFluxoCaixa.id_plano_contas_grupo;
			lancamento.id_plano_contas_conta = rowFinFluxoCaixa.id_plano_contas_conta;
			lancamento.natureza = rowFinFluxoCaixa.natureza;
			lancamento.st_sem_efeito = rowFinFluxoCaixa.st_sem_efeito;
			lancamento.dt_competencia = rowFinFluxoCaixa.dt_competencia;
            lancamento.dt_mes_competencia = (rowFinFluxoCaixa.Isdt_mes_competenciaNull() ? DateTime.MinValue : rowFinFluxoCaixa.dt_mes_competencia);
			lancamento.valor = rowFinFluxoCaixa.valor;
			lancamento.descricao = rowFinFluxoCaixa.descricao;
			lancamento.ctrl_pagto_id_parcela = rowFinFluxoCaixa.ctrl_pagto_id_parcela;
			lancamento.ctrl_pagto_modulo = rowFinFluxoCaixa.ctrl_pagto_modulo;
			lancamento.ctrl_pagto_status = rowFinFluxoCaixa.ctrl_pagto_status;
			lancamento.id_cliente = (rowFinFluxoCaixa.Isid_clienteNull() ? "" : rowFinFluxoCaixa.id_cliente);
			lancamento.cnpj_cpf = (rowFinFluxoCaixa.Iscnpj_cpfNull() ? "" : rowFinFluxoCaixa.cnpj_cpf);
			lancamento.numero_NF = (rowFinFluxoCaixa.Isnumero_NFNull() ? 0 : rowFinFluxoCaixa.numero_NF);
			lancamento.tipo_cadastro = rowFinFluxoCaixa.tipo_cadastro;
			lancamento.editado_manual = rowFinFluxoCaixa.editado_manual;
			lancamento.dt_cadastro = rowFinFluxoCaixa.dt_cadastro;
			lancamento.dt_hr_cadastro = rowFinFluxoCaixa.dt_hr_cadastro;
			lancamento.usuario_cadastro = rowFinFluxoCaixa.usuario_cadastro;
			lancamento.dt_ult_atualizacao = rowFinFluxoCaixa.dt_ult_atualizacao;
			lancamento.dt_hr_ult_atualizacao = rowFinFluxoCaixa.dt_hr_ult_atualizacao;
			lancamento.usuario_ult_atualizacao = rowFinFluxoCaixa.usuario_ult_atualizacao;
			lancamento.st_confirmacao_pendente = rowFinFluxoCaixa.st_confirmacao_pendente;
			lancamento.st_boleto_pago_cheque = rowFinFluxoCaixa.st_boleto_pago_cheque;
			lancamento.dt_ocorrencia_banco_boleto_pago_cheque = (!rowFinFluxoCaixa.Isdt_ocorrencia_banco_boleto_pago_chequeNull() ? (DateTime)rowFinFluxoCaixa.dt_ocorrencia_banco_boleto_pago_cheque : DateTime.MinValue);
			lancamento.st_boleto_ocorrencia_17 = rowFinFluxoCaixa.st_boleto_ocorrencia_17;
			lancamento.dt_ocorrencia_banco_boleto_ocorrencia_17 = (!rowFinFluxoCaixa.Isdt_ocorrencia_banco_boleto_ocorrencia_17Null() ? (DateTime)rowFinFluxoCaixa.dt_ocorrencia_banco_boleto_ocorrencia_17 : DateTime.MinValue);
			lancamento.st_boleto_ocorrencia_15 = rowFinFluxoCaixa.st_boleto_ocorrencia_15;
			lancamento.dt_ocorrencia_banco_boleto_ocorrencia_15 = (!rowFinFluxoCaixa.Isdt_ocorrencia_banco_boleto_ocorrencia_15Null() ? (DateTime)rowFinFluxoCaixa.dt_ocorrencia_banco_boleto_ocorrencia_15 : DateTime.MinValue);
			lancamento.st_boleto_ocorrencia_23 = rowFinFluxoCaixa.st_boleto_ocorrencia_23;
			lancamento.dt_ocorrencia_banco_boleto_ocorrencia_23 = (!rowFinFluxoCaixa.Isdt_ocorrencia_banco_boleto_ocorrencia_23Null() ? (DateTime)rowFinFluxoCaixa.dt_ocorrencia_banco_boleto_ocorrencia_23 : DateTime.MinValue);
			lancamento.st_boleto_ocorrencia_34 = rowFinFluxoCaixa.st_boleto_ocorrencia_34;
			lancamento.dt_ocorrencia_banco_boleto_ocorrencia_34 = (!rowFinFluxoCaixa.Isdt_ocorrencia_banco_boleto_ocorrencia_34Null() ? (DateTime)rowFinFluxoCaixa.dt_ocorrencia_banco_boleto_ocorrencia_34 : DateTime.MinValue);
			lancamento.st_boleto_baixado = rowFinFluxoCaixa.st_boleto_baixado;
			lancamento.dt_ocorrencia_banco_boleto_baixado = (!rowFinFluxoCaixa.Isdt_ocorrencia_banco_boleto_baixadoNull() ? (DateTime)rowFinFluxoCaixa.dt_ocorrencia_banco_boleto_baixado : DateTime.MinValue);
			lancamento.id_boleto_cedente = rowFinFluxoCaixa.id_boleto_cedente;

			return lancamento;
		}
		#endregion

		#region [ insere ]
		/// <summary>
		/// Grava o novo lançamento de fluxo de caixa no banco de dados
		/// </summary>
		/// <param name="usuario">
		/// Usuário logado que está realizando a operação para registrar no log
		/// </param>
		/// <param name="lancamento">
		/// Objeto representando o lançamento do fluxo de caixa contendo os dados a gravar
		/// </param>
		/// <param name="strDescricaoLog">
		/// Parâmetro que retornará os dados para gravar no log
		/// </param>
		/// <param name="strMsgErro">
		/// Parâmetro que retornará a mensagem de erro em caso de exception
		/// </param>
		/// <returns>
		/// true: gravação efetuada com sucesso
		/// false: falha na gravação
		/// </returns>
		public static bool insere(String usuario,
								  LancamentoFluxoCaixa lancamento,
								  ref String strDescricaoLog,
								  ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intQtdeTentativas = 0;
			int intNsu = 0;
			int intRetorno;
			String strOperacao = "Gravação de lançamento no fluxo de caixa (" + Global.retornaDescricaoFluxoCaixaNatureza(lancamento.natureza) + ")";
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_FLUXO_CAIXA, ref intNsu, ref strMsgErro);

					#region [ Se gerou o NSU, tenta gravar o registro ]
					if (blnGerouNsu)
					{
                        #region [ Preenche o valor dos parâmetros ]
                        cmLancamentoInsert.Parameters["@id"].Value = intNsu;
						cmLancamentoInsert.Parameters["@id_conta_corrente"].Value = lancamento.id_conta_corrente;
						cmLancamentoInsert.Parameters["@id_plano_contas_empresa"].Value = lancamento.id_plano_contas_empresa;
						cmLancamentoInsert.Parameters["@id_plano_contas_grupo"].Value = lancamento.id_plano_contas_grupo;
						cmLancamentoInsert.Parameters["@id_plano_contas_conta"].Value = lancamento.id_plano_contas_conta;
						cmLancamentoInsert.Parameters["@natureza"].Value = lancamento.natureza;
						cmLancamentoInsert.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(lancamento.dt_competencia);
						cmLancamentoInsert.Parameters["@valor"].Value = lancamento.valor;
						cmLancamentoInsert.Parameters["@descricao"].Value = lancamento.descricao;
						cmLancamentoInsert.Parameters["@cnpj_cpf"].Value = Global.digitos(lancamento.cnpj_cpf);
						cmLancamentoInsert.Parameters["@numero_NF"].Value = lancamento.numero_NF;
						cmLancamentoInsert.Parameters["@tipo_cadastro"].Value = lancamento.tipo_cadastro;
						cmLancamentoInsert.Parameters["@editado_manual"].Value = Global.Cte.FIN.EditadoManual.NAO;
						cmLancamentoInsert.Parameters["@usuario_cadastro"].Value = usuario;
						cmLancamentoInsert.Parameters["@usuario_ult_atualizacao"].Value = usuario;
                        cmLancamentoInsert.Parameters["@dt_mes_competencia"].Value = lancamento.natureza.ToString().ToUpper().Equals("D") ? Global.formataDataYyyyMmDdComSeparador(lancamento.dt_mes_competencia) : (object)DBNull.Value;
						#endregion

						#region [ Monta texto para o log em arquivo ]
						// Se houver conteúdo de alguma tentativa anterior, descarta
						sbLog = new StringBuilder("");
						foreach (SqlParameter item in cmLancamentoInsert.Parameters)
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
						#endregion

						#region [ Tenta inserir o registro ]
						try
						{
							intRetorno = BD.executaNonQuery(ref cmLancamentoInsert);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
						}
						#endregion

						#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
						if (intRetorno == 1)
						{
							lancamento.id = intNsu;
							strDescricaoLog = sbLog.ToString();
							Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
							blnSucesso = true;
						}
						else
						{
							Thread.Sleep(100);
						}
						#endregion
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gravar no banco de dados o lançamento de fluxo de caixa após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ altera ]
		/// <summary>
		/// Altera os dados de lançamento do fluxo de caixa já gravado
		/// </summary>
		/// <param name="usuario">
		/// Usuário logado que está realizando a operação para registrar no log
		/// </param>
		/// <param name="lancamento">
		/// Objeto representando o lançamento do fluxo de caixa contendo os dados a gravar
		/// </param>
		/// <param name="strDescricaoLog">
		/// Parâmetro que retornará os dados para gravar no log
		/// </param>
		/// Parâmetro que retornará a mensagem de erro em caso de exception
		/// <param name="strMsgErro">
		/// </param>
		/// <returns>
		/// true: alteração efetuada com sucesso
		/// false: falha na alteração
		/// </returns>
		public static bool altera(String usuario,
								  LancamentoFluxoCaixa lancamento,
								  ref String strDescricaoLog,
								  ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal = null;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado = null;
			String strOperacao = "Alteração de lançamento no fluxo de caixa (" + Global.retornaDescricaoFluxoCaixaNatureza(lancamento.natureza) + ")";
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				#region [ Laço de tentativas de alteração no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";

					#region [ Obtém dados originais para montar mensagem do log ]
					dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoBD(lancamento.id);
					if (dtbFinFluxoCaixaOriginal.Rows.Count > 0) rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];
					#endregion

					#region [ Preenche o valor dos parâmetros ]
					cmLancamentoUpdate.Parameters["@id"].Value = lancamento.id;
					cmLancamentoUpdate.Parameters["@id_conta_corrente"].Value = lancamento.id_conta_corrente;
					cmLancamentoUpdate.Parameters["@id_plano_contas_empresa"].Value = lancamento.id_plano_contas_empresa;
					cmLancamentoUpdate.Parameters["@id_plano_contas_grupo"].Value = lancamento.id_plano_contas_grupo;
					cmLancamentoUpdate.Parameters["@id_plano_contas_conta"].Value = lancamento.id_plano_contas_conta;
					cmLancamentoUpdate.Parameters["@natureza"].Value = lancamento.natureza;
					cmLancamentoUpdate.Parameters["@st_sem_efeito"].Value = lancamento.st_sem_efeito;
					cmLancamentoUpdate.Parameters["@st_confirmacao_pendente"].Value = lancamento.st_confirmacao_pendente;
					cmLancamentoUpdate.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(lancamento.dt_competencia);
                    cmLancamentoUpdate.Parameters["@dt_mes_competencia"].Value = lancamento.natureza.ToString().ToUpper().Equals("D") ? Global.formataDataYyyyMmDdComSeparador(lancamento.dt_mes_competencia) : (object)DBNull.Value;
                    cmLancamentoUpdate.Parameters["@valor"].Value = lancamento.valor;
					cmLancamentoUpdate.Parameters["@descricao"].Value = lancamento.descricao;
					cmLancamentoUpdate.Parameters["@ctrl_pagto_status"].Value = lancamento.ctrl_pagto_status;
					cmLancamentoUpdate.Parameters["@cnpj_cpf"].Value = Global.digitos(lancamento.cnpj_cpf);
					cmLancamentoUpdate.Parameters["@numero_NF"].Value = lancamento.numero_NF;
					cmLancamentoUpdate.Parameters["@editado_manual"].Value = Global.Cte.FIN.EditadoManual.SIM;
					cmLancamentoUpdate.Parameters["@usuario_ult_atualizacao"].Value = usuario;
					#endregion

					#region [ Tenta alterar o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmLancamentoUpdate);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de alteração ]
					if (intRetorno == 1)
					{
						#region [ Obtém dados alterados para montar mensagem do log ]
						dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(lancamento.id);
						if (dtbFinFluxoCaixaEditado.Rows.Count > 0) rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
						#endregion

						#region [ Monta mensagem para o log ]
						if ((rowFinFluxoCaixaOriginal != null) && (rowFinFluxoCaixaEditado != null))
						{
							foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
							{
								if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
								{
									if (sbLog.Length > 0) sbLog.Append("; ");
									sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
								}
							}
						}
						#endregion

						strDescricaoLog = sbLog.ToString();
						Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
						blnSucesso = true;
					}
					else
					{
						Thread.Sleep(100);
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_UPDATE_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar alterar no banco de dados o lançamento de fluxo de caixa após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ alteraPorEdicaoEmLote ]
		/// <summary>
		/// Altera os dados de lançamento do fluxo de caixa já gravado, mas apenas 
		/// os campos editáveis pelo painel de edição em lote.
		/// </summary>
		/// <param name="usuario">
		/// Usuário logado que está realizando a operação para registrar no log
		/// </param>
		/// <param name="intIdLancto">
		/// Nº identificação do registro do lançamento
		/// </param>
		/// <param name="byteStSemEfeito">
		/// Valor do campo st_sem_efeito
		/// </param>
		/// <param name="byteStConfirmacaoPendente">
		/// Valor do campo st_confirmacao_pendente
		/// </param>
		/// <param name="byteCtrlPagtoStatus">
		/// Valor do campo ctrl_pagto_status
		/// </param>
		/// <param name="dtCompetencia">
		/// Valor do campo dt_competencia
		/// </param>
		/// <param name="strDescricaoLancamento">
		/// Valor do campo descricao
		/// </param>
		/// <param name="strDescricaoLog">
		/// Parâmetro que retornará os dados para gravar no log
		/// </param>
		/// <param name="strMsgErro">
		/// Parâmetro que retornará a mensagem de erro em caso de exception
		/// </param>
		/// <returns>
		/// true: alteração efetuada com sucesso
		/// false: falha na alteração
		/// </returns>
		public static bool alteraPorEdicaoEmLote(String usuario,
									int intIdLancto,
									byte byteStSemEfeito,
									byte byteStConfirmacaoPendente,
									byte byteCtrlPagtoStatus,
									DateTime dtCompetencia,
                                    DateTime dtComp2,
									byte byteContaCorrente,
									byte bytePlanoContasEmpresa,
									int intPlanoContasGrupo,
									int intPlanoContasConta,
									String strDescricaoLancamento,
									ref String strDescricaoLog,
									ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			String strSql;
			String strClausulaSet = "";
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal = null;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado = null;
			String strOperacao = "Alteração de lançamento no fluxo de caixa por edição em lote";
			StringBuilder sbLog = new StringBuilder("");
			SqlCommand cmUpdate;
			#endregion

			try
			{
				#region [ Monta comando SQL de acordo c/ os parâmetros a serem alterados ]

				#region [ Cria o command ]
				cmUpdate = BD.criaSqlCommand();
				#endregion

				#region [ Campo: st_sem_efeito ]
				if (byteStSemEfeito != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "st_sem_efeito = @st_sem_efeito";
					cmUpdate.Parameters.Add("@st_sem_efeito", SqlDbType.TinyInt);
					cmUpdate.Parameters["@st_sem_efeito"].Value = byteStSemEfeito;
				}
				#endregion

				#region [ Campo: st_confirmacao_pendente ]
				if (byteStConfirmacaoPendente != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "st_confirmacao_pendente = @st_confirmacao_pendente";
					cmUpdate.Parameters.Add("@st_confirmacao_pendente", SqlDbType.TinyInt);
					cmUpdate.Parameters["@st_confirmacao_pendente"].Value = byteStConfirmacaoPendente;
				}
				#endregion

				#region [ Campo: ctrl_pagto_status ]
				if (byteCtrlPagtoStatus != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "ctrl_pagto_status = @ctrl_pagto_status";
					cmUpdate.Parameters.Add("@ctrl_pagto_status", SqlDbType.TinyInt);
					cmUpdate.Parameters["@ctrl_pagto_status"].Value = byteCtrlPagtoStatus;
				}
				#endregion

				#region [ Campo: dt_competencia ]
				if (dtCompetencia != DateTime.MinValue)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "dt_competencia = @dt_competencia";
					cmUpdate.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
					cmUpdate.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dtCompetencia);
				}
                #endregion

                #region [ Campo: dt_mes_competencia ]
                if (dtComp2 != DateTime.MinValue)
                {
                    if (strClausulaSet.Length > 0) strClausulaSet += ", ";
                    strClausulaSet += "dt_mes_competencia = @dt_mes_competencia";
                    cmUpdate.Parameters.Add("@dt_mes_competencia", SqlDbType.VarChar, 10);
                    cmUpdate.Parameters["@dt_mes_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dtComp2);
                }
				#endregion

				#region [ Campo: id_conta_corrente ]
				if (byteContaCorrente != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "id_conta_corrente = @id_conta_corrente";
					cmUpdate.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
					cmUpdate.Parameters["@id_conta_corrente"].Value = byteContaCorrente;
				}
				#endregion

				#region [ Campo: id_plano_contas_empresa ]
				if (bytePlanoContasEmpresa != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "id_plano_contas_empresa = @id_plano_contas_empresa";
					cmUpdate.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
					cmUpdate.Parameters["@id_plano_contas_empresa"].Value = bytePlanoContasEmpresa;
				}
				#endregion

				#region [ Campo: id_plano_contas_conta ]
				if (intPlanoContasConta != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "id_plano_contas_conta = @id_plano_contas_conta";
					cmUpdate.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
					cmUpdate.Parameters["@id_plano_contas_conta"].Value = intPlanoContasConta;
				}
				#endregion

				#region [ Campo: id_plano_contas_grupo ]
				if (intPlanoContasGrupo != Global.Cte.Etc.FLAG_NAO_SETADO)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "id_plano_contas_grupo = @id_plano_contas_grupo";
					cmUpdate.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
					cmUpdate.Parameters["@id_plano_contas_grupo"].Value = intPlanoContasGrupo;
				}
				#endregion

				#region [ Campo: descricao ]
				if (strDescricaoLancamento.Length > 0)
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ", ";
					strClausulaSet += "descricao = @descricao";
					cmUpdate.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
					cmUpdate.Parameters["@descricao"].Value = strDescricaoLancamento;
				}
				#endregion

				#region [ Há alterações? ]
				if (strClausulaSet.Length == 0) throw new Exception("Não há alterações para efetuar no lançamento do fluxo de caixa Id=" + intIdLancto.ToString() + "!!");
				#endregion

				#region [ Demais campos que informam sobre alterações no registro ]
				if (strClausulaSet.Length > 0) strClausulaSet += ", ";
				strClausulaSet += "editado_manual = @editado_manual, " +
								  "dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
								  "dt_hr_ult_atualizacao = getdate(), " +
								  "usuario_ult_atualizacao = @usuario_ult_atualizacao ";

				cmUpdate.Parameters.Add("@editado_manual", SqlDbType.Char, 1);
				cmUpdate.Parameters["@editado_manual"].Value = Global.Cte.FIN.EditadoManual.SIM;

				cmUpdate.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
				cmUpdate.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Campo: id ]
				cmUpdate.Parameters.Add("@id", SqlDbType.Int);
				cmUpdate.Parameters["@id"].Value = intIdLancto;
				#endregion

				#region [ Monta o SQL ]
				strSql = "UPDATE t_FIN_FLUXO_CAIXA SET " +
						 strClausulaSet +
						 " WHERE" +
							  " (id = @id)";
				cmUpdate.CommandText = strSql;
				#endregion

				#endregion

				#region [ Obtém dados originais para montar mensagem do log ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoBD(intIdLancto);
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 0) rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];
				#endregion

				#region [ Laço de tentativas de alteração no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";

					#region [ Tenta alterar o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmUpdate);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de alteração ]
					if (intRetorno == 1)
					{
						#region [ Obtém dados alterados para montar mensagem do log ]
						dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(intIdLancto);
						if (dtbFinFluxoCaixaEditado.Rows.Count > 0) rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
						#endregion

						#region [ Monta mensagem para o log ]
						if ((rowFinFluxoCaixaOriginal != null) && (rowFinFluxoCaixaEditado != null))
						{
							foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
							{
								if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
								{
									if (sbLog.Length > 0) sbLog.Append("; ");
									sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
								}
							}
						}
						#endregion

						strDescricaoLog = sbLog.ToString();
						Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
						blnSucesso = true;
					}
					else
					{
						Thread.Sleep(100);
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_UPDATE_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar alterar no banco de dados o lançamento de fluxo de caixa após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ exclui ]
		/// <summary>
		/// Apaga do banco de dados o registro do lançamento do fluxo de caixa
		/// </summary>
		/// <param name="usuario">
		/// Usuário logado que está realizando a operação para registrar no log
		/// </param>
		/// <param name="intIdSelecionado">
		/// Nº identificação do registro a ser excluído
		/// </param>
		/// <param name="strDescricaoLog">
		/// Parâmetro que retornará os dados para gravar no log
		/// </param>
		/// Parâmetro que retornará a mensagem de erro em caso de exception
		/// <param name="strMsgErro">
		/// </param>
		/// <returns>
		/// true: exclusão efetuada com sucesso
		/// false: falha na exclusão
		/// </returns>
		public static bool exclui(String usuario,
								  int intIdSelecionado,
								  ref String strDescricaoLog,
								  ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intQtdeTentativas = 0;
			int intRetorno;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal = null;
			String strOperacao = "Exclusão de lançamento no fluxo de caixa";
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				#region [ Laço de tentativas de exclusão no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";

					#region [ Obtém dados originais para montar mensagem do log ]
					dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoBD(intIdSelecionado);
					if (dtbFinFluxoCaixaOriginal.Rows.Count > 0)
					{
						rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];
						// Complementa o título da operação adicionando a natureza do lançamento (crédito/débito)
						strOperacao += " (" + Global.retornaDescricaoFluxoCaixaNatureza(rowFinFluxoCaixaOriginal.natureza) + ")";
					}
					#endregion

					#region [ Preenche o valor dos parâmetros ]
					cmLancamentoDelete.Parameters["@id"].Value = intIdSelecionado;
					#endregion

					#region [ Tenta excluir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmLancamentoDelete);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de exclusão ]
					if (intRetorno == 1)
					{
						#region [ Monta mensagem para o log ]
						if (rowFinFluxoCaixaOriginal != null)
						{
							foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
							{
								if (sbLog.Length > 0) sbLog.Append("; ");
								sbLog.Append(coluna.ColumnName + "=" + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString());
							}
						}
						#endregion

						strDescricaoLog = sbLog.ToString();
						Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
						blnSucesso = true;
					}
					else
					{
						Thread.Sleep(100);
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_DELETE_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar excluir do banco de dados o lançamento de fluxo de caixa após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ insereLancamentoDevidoBoletoOcorrencia02 ]
		/// <summary>
		/// Gera automaticamente o lançamento no fluxo de caixa com a previsão do valor a receber.
		/// Caso o registro já tenha sido criado anteriormente, atualiza os dados.
		/// </summary>
		/// <param name="usuario">
		/// Usuário logado que está realizando a operação para registrar no log
		/// </param>
		/// <param name="idBoletoItem">
		/// Id do registro desta parcela do boleto (referente ao Id da tabela t_FIN_BOLETO_ITEM).
		/// </param>
		/// <param name="dataCompetencia">
		/// Data de vencimento da parcela.
		/// </param>
		/// <param name="valorTitulo">
		/// Valor da parcela.
		/// </param>
		/// <param name="descricao">
		/// Descrição da parcela.
		/// </param>
		/// <param name="id_fluxo_caixa">
		/// Retorna o identificador do registro do lançamento do fluxo de caixa criado/atualizado.
		/// </param>
		/// <param name="tipoAtualizacaoEfetuada">
		/// Retorna o tipo de atualização efetuada: insert ou update
		/// </param>
		/// <param name="strMsgErro">
		/// Mensagem de erro no caso de ocorrer erro.
		/// </param>
		/// <returns>
		/// true: rotina executada com sucesso.
		/// false: falha na execução.
		/// </returns>
		public static bool insereLancamentoDevidoBoletoOcorrencia02(String usuario,
														  int idBoletoItem,
														  DateTime dataCompetencia,
														  decimal valorTitulo,
														  String descricao,
														  out int id_fluxo_caixa,
														  out Global.eTipoAtualizacaoEfetuada tipoAtualizacaoEfetuada,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Gravação de lançamento no fluxo de caixa devido a boleto com ocorrência 02 (entrada confirmada)";
			BoletoCedente boletoCedente;
			bool blnExisteLancamento;
			bool blnGerouNsu;
			int intNsuLancamento = 0;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			FinLog finLog = new FinLog();
			BoletoPlanoContasDestino boletoPlanoContasDestino;
			#endregion

			tipoAtualizacaoEfetuada = Global.eTipoAtualizacaoEfetuada.NENHUMA_ALTERACAO_REALIZADA;
			id_fluxo_caixa = 0;
			strMsgErro = "";
			try
			{
				#region [ Obtém os dados do plano de contas em que o lançamento deve ser cadastrado ]
				boletoPlanoContasDestino = BoletoDAO.obtemBoletoPlanoContasDestinoByIdBoletoItem(idBoletoItem);
				#endregion

				#region [ Pesquisa BD p/ verificar se já existe registro no fluxo de caixa associado a este boleto ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				blnExisteLancamento = true;
				if (dtbFinFluxoCaixaOriginal == null)
					blnExisteLancamento = false;
				else if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
					blnExisteLancamento = false;

				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				if (blnExisteLancamento)
				{
					rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

					#region [ Já existe um registro no fluxo de caixa associado a este boleto, então atualiza-o ]
					tipoAtualizacaoEfetuada = Global.eTipoAtualizacaoEfetuada.ALTERADO_REGISTRO_JA_EXISTENTE;
					id_fluxo_caixa = rowFinFluxoCaixaOriginal.id;

					#region [ Preenche o valor dos parâmetros ]
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@st_sem_efeito"].Value = Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO;
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dataCompetencia);
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@valor"].Value = valorTitulo;
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@descricao"].Value = Texto.leftStr(descricao, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.CADASTRADO_INICIAL;
					cmLancamentoUpdateDevidoBoletoOcorrencia02.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
					#endregion

					#region [ Tenta atualizar o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia02);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + "\nAlteração sobre registro já existente!!\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de atualização ]
					if (intRetorno != 1)
					{
						strMsgErro = "Falha ao tentar atualizar o registro associado do lançamento do fluxo de caixa devido a boleto com ocorrência 02 (entrada confirmada)!!";
						return false;
					}
					#endregion

					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 02 (entrada confirmada) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_02;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

					#endregion
				}
				else
				{
					#region [ Não há nenhum registro no fluxo de caixa associado a este boleto, então cria um novo ]
					tipoAtualizacaoEfetuada = Global.eTipoAtualizacaoEfetuada.INCLUSAO_NOVO_REGISTRO;

					#region [ Gera NSU ]
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_FLUXO_CAIXA, ref intNsuLancamento, ref strMsgErro);
					if (!blnGerouNsu)
					{
						strMsgErro = "Falha ao gerar o NSU para o lançamento no fluxo de caixa!!\n" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Atualiza parâmetro de saída c/ Id do lançamento ]
					id_fluxo_caixa = intNsuLancamento;
					#endregion

					#region [ Obtém dados do registro principal do boleto ]
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal == null)
					{
						strMsgErro = "Falha ao obter os dados do registro principal do boleto!!";
						return false;
					}
					#endregion

					#region [ Obtém os dados do cedente ]
					boletoCedente = BoletoCedenteDAO.getBoletoCedente(rowBoletoPrincipal.id_boleto_cedente);
					if (boletoCedente == null)
					{
						strMsgErro = "Falha ao obter os dados do cedente!!";
						return false;
					}
					#endregion

					#region [ Preenche o valor dos parâmetros ]
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id"].Value = intNsuLancamento;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id_conta_corrente"].Value = boletoCedente.id_conta_corrente;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id_plano_contas_empresa"].Value = boletoPlanoContasDestino.id_plano_contas_empresa;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id_plano_contas_grupo"].Value = boletoPlanoContasDestino.id_plano_contas_grupo;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id_plano_contas_conta"].Value = boletoPlanoContasDestino.id_plano_contas_conta;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@natureza"].Value = Global.Cte.FIN.Natureza.CREDITO;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dataCompetencia);
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@valor"].Value = valorTitulo;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@descricao"].Value = Texto.leftStr(descricao, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id_boleto_cedente"].Value = boletoCedente.id;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@id_cliente"].Value = rowBoletoPrincipal.id_cliente;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@cnpj_cpf"].Value = Global.digitos(rowBoletoPrincipal.num_inscricao_sacado);
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@tipo_cadastro"].Value = Global.Cte.FIN.TipoCadastro.SISTEMA;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@editado_manual"].Value = Global.Cte.FIN.EditadoManual.NAO;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@usuario_cadastro"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
					cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters["@st_confirmacao_pendente"].Value = Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO;
					#endregion

					#region [ Monta texto p/ o log ]
					foreach (SqlParameter item in cmLancamentoInsertDevidoBoletoOcorrencia02.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmLancamentoInsertDevidoBoletoOcorrencia02);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
					}
					#endregion

					#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
					if (intRetorno != 1)
					{
						strMsgErro = "Falha ao tentar gerar o registro automático do lançamento do fluxo de caixa devido a boleto com ocorrência 02 (entrada confirmada)!!";
						return false;
					}
					#endregion

					#region [ Registra log com a inclusão dos dados ]
					strDescricaoLog = "Inserção do registro em t_FIN_FLUXO_CAIXA.id=" + intNsuLancamento.ToString() + " devido à ocorrência 02 (entrada confirmada) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
					Global.gravaLogAtividade(strDescricaoLog);
					finLog.usuario = usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_OCORRENCIA_02;
					finLog.natureza = Global.Cte.FIN.Natureza.CREDITO;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
					finLog.id_registro_origem = intNsuLancamento;
					finLog.id_conta_corrente = boletoCedente.id_conta_corrente;
					finLog.id_plano_contas_empresa = boletoPlanoContasDestino.id_plano_contas_empresa;
					finLog.id_plano_contas_grupo = boletoPlanoContasDestino.id_plano_contas_grupo;
					finLog.id_plano_contas_conta = boletoPlanoContasDestino.id_plano_contas_conta;
					finLog.id_boleto_cedente = boletoCedente.id;
					finLog.id_cliente = rowBoletoPrincipal.id_cliente;
					finLog.cnpj_cpf = Global.digitos(rowBoletoPrincipal.num_inscricao_sacado);
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					#endregion

					#endregion
				}

				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaLancamentoDevidoBoletoOcorrencia06 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 06 (liquidação normal).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataCredito">Data em que o valor pago foi creditado pelo banco</param>
		/// <param name="valorPago">Valor pago</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia06(String usuario,
														  int idBoletoItem,
														  DateTime dataCredito,
														  decimal valorPago,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 06 (liquidação normal)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@st_sem_efeito"].Value = Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dataCredito);
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@valor"].Value = valorPago;
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.PAGO;
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@st_confirmacao_pendente"].Value = Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia06.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia06);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 06 (liquidação normal) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_06;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 06 (liquidação normal)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia09 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 09 (baixado automaticamente via arquivo).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia09(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 09 (baixado automaticamente via arquivo)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_BAIXADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters["@st_boleto_baixado"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters["@dt_ocorrencia_banco_boleto_baixado"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmLancamentoUpdateDevidoBoletoOcorrencia09.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia09);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 09 (baixado automaticamente via arquivo) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_09;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia10 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 10 (baixado conforme instruções da agência).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia10(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 10 (baixado conforme instruções da agência)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_BAIXADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters["@st_boleto_baixado"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters["@dt_ocorrencia_banco_boleto_baixado"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmLancamentoUpdateDevidoBoletoOcorrencia10.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia10);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 10 (baixado conforme instruções da agência) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_10;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 10 (baixado conforme instruções da agência)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia12 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 12 (abatimento concedido).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="novoValorTitulo">Novo valor do título</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia12(String usuario,
														  int idBoletoItem,
														  decimal novoValorTitulo,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 12 (abatimento concedido)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia12.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia12.Parameters["@valor"].Value = novoValorTitulo;
				cmLancamentoUpdateDevidoBoletoOcorrencia12.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia12);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 12 (abatimento concedido) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_12;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 12 (abatimento concedido)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia13 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 13 (abatimento cancelado).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="novoValorTitulo">Novo valor do título</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia13(String usuario,
														  int idBoletoItem,
														  decimal novoValorTitulo,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 13 (abatimento cancelado)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia13.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia13.Parameters["@valor"].Value = novoValorTitulo;
				cmLancamentoUpdateDevidoBoletoOcorrencia13.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia13);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 13 (abatimento cancelado) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_13;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 13 (abatimento cancelado)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia14 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 14 (vencimento alterado).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dtNovoVencto">Nova data de vencimento</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia14(String usuario,
														  int idBoletoItem,
														  DateTime dtNovoVencto,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 14 (vencimento alterado)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia14.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia14.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dtNovoVencto);
				cmLancamentoUpdateDevidoBoletoOcorrencia14.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia14);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 14 (vencimento alterado) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_14;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 14 (vencimento alterado)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia15 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 15 (liquidação em cartório).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataCredito">Data em que o valor pago foi creditado pelo banco</param>
		/// <param name="valorPago">Valor pago</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia15(String usuario,
														  int idBoletoItem,
														  DateTime dataCredito,
														  decimal valorPago,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 15 (liquidação em cartório)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@st_sem_efeito"].Value = Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dataCredito);
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@valor"].Value = valorPago;
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.PAGO;
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@st_confirmacao_pendente"].Value = Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@st_boleto_ocorrencia_15"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_15"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmLancamentoUpdateDevidoBoletoOcorrencia15.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia15);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 15 (liquidação em cartório) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_15;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 15 (liquidação em cartório)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia16 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 16 (título pago em cheque).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia16(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 16 (título pago em cheque)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters["@st_sem_efeito"].Value = Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_PAGO_CHEQUE_VINCULADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters["@st_boleto_pago_cheque"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters["@dt_ocorrencia_banco_boleto_pago_cheque"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmLancamentoUpdateDevidoBoletoOcorrencia16.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia16);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 16 (título pago em cheque) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_16;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 16 (título pago em cheque)!!";
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

		#region [ insereLancamentoDevidoBoletoOcorrencia17 ]
		/// <summary>
		/// Gera automaticamente o lançamento no fluxo de caixa com um valor já quitado devido
		/// a um boleto com ocorrência 17 (liquidação após baixa ou título não registrado).
		/// </summary>
		/// <param name="usuario">
		/// Usuário logado que está realizando a operação para registrar no log
		/// </param>
		/// <param name="idBoletoItem">
		/// Id do registro desta parcela do boleto (referente ao Id da tabela t_FIN_BOLETO_ITEM).
		/// Pode ser zero no caso de se tratar de um boleto não cadastrado no sistema.
		/// </param>
		/// <param name="dataCompetencia">
		/// Data de competência do lançamento no fluxo de caixa.
		/// </param>
		/// <param name="valorTitulo">
		/// Valor do lançamento no fluxo de caixa.
		/// </param>
		/// <param name="descricao">
		/// Descrição da parcela.
		/// </param>
		/// <param name="dataOcorrenciaBanco">
		/// Data da ocorrência no banco.
		/// </param>
		/// <param name="id_fluxo_caixa">
		/// Retorna o nº identificação do registro gerado no fluxo de caixa.
		/// </param>
		/// <param name="strMsgErro">
		/// Mensagem de erro no caso de ocorrer erro.
		/// </param>
		/// <returns>
		/// true: rotina executada com sucesso.
		/// false: falha na execução.
		/// </returns>
		public static bool insereLancamentoDevidoBoletoOcorrencia17(String usuario,
														  int idBoletoItem,
														  int id_boleto_cedente,
														  DateTime dataCompetencia,
														  decimal valorTitulo,
														  String descricao,
														  DateTime dataOcorrenciaBanco,
														  out int id_fluxo_caixa,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Gravação de lançamento no fluxo de caixa devido a boleto com ocorrência 17 (liquidação após baixa ou título não registrado)";
			BoletoCedente boletoCedente;
			bool blnGerouNsu;
			int intNsuLancamento = 0;
			int intRetorno;
			String strIdCliente = "";
			String strCnpjCpf = "";
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			FinLog finLog = new FinLog();
			BoletoPlanoContasDestino boletoPlanoContasDestino;
			#endregion

			id_fluxo_caixa = 0;
			strMsgErro = "";
			try
			{
				if (idBoletoItem > 0)
				{
					#region [ Obtém os dados do plano de contas em que o lançamento deve ser cadastrado ]
					boletoPlanoContasDestino = BoletoDAO.obtemBoletoPlanoContasDestinoByIdBoletoItem(idBoletoItem);
					#endregion

					#region [ Obtém dados do registro principal do boleto ]
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal == null)
					{
						strMsgErro = "Falha ao obter os dados do registro principal do boleto!!";
						return false;
					}

					strIdCliente = rowBoletoPrincipal.id_cliente;
					strCnpjCpf = rowBoletoPrincipal.num_inscricao_sacado;
					#endregion

					#region [ Obtém os dados do cedente ]
					boletoCedente = BoletoCedenteDAO.getBoletoCedente(rowBoletoPrincipal.id_boleto_cedente);
					if (boletoCedente == null)
					{
						strMsgErro = "Falha ao obter os dados do cedente!!";
						return false;
					}
					#endregion
				}
				else
				{
					#region [ Obtém os dados do cedente ]
					boletoCedente = BoletoCedenteDAO.getBoletoCedente(id_boleto_cedente);
					if (boletoCedente == null)
					{
						strMsgErro = "Falha ao obter os dados do cedente!!";
						return false;
					}
					#endregion

					#region [ Obtém os dados do plano de contas em que o lançamento deve ser cadastrado ]
					boletoPlanoContasDestino = BoletoDAO.obtemBoletoPlanoContasDestinoByNumLoja((int)Global.converteInteiro(boletoCedente.loja_default_boleto_plano_contas));
					if (boletoPlanoContasDestino == null)
					{
						strMsgErro = "Não foi localizado o plano de contas de destino para a loja default (" + boletoCedente.loja_default_boleto_plano_contas + ") do cedente " + boletoCedente.apelido + "!!";
						return false;
					}
					#endregion
				}

				#region [ Insere novo registro ]

				#region [ Gera NSU ]
				blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_FLUXO_CAIXA, ref intNsuLancamento, ref strMsgErro);
				if (!blnGerouNsu)
				{
					strMsgErro = "Falha ao gerar o NSU para o lançamento no fluxo de caixa!!\n" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza parâmetro de saída c/ Id do lançamento ]
				id_fluxo_caixa = intNsuLancamento;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id"].Value = intNsuLancamento;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id_conta_corrente"].Value = boletoCedente.id_conta_corrente;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id_plano_contas_empresa"].Value = boletoPlanoContasDestino.id_plano_contas_empresa;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id_plano_contas_grupo"].Value = boletoPlanoContasDestino.id_plano_contas_grupo;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id_plano_contas_conta"].Value = boletoPlanoContasDestino.id_plano_contas_conta;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@natureza"].Value = Global.Cte.FIN.Natureza.CREDITO;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(dataCompetencia);
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@valor"].Value = valorTitulo;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@descricao"].Value = Texto.leftStr(descricao, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.PAGO;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id_boleto_cedente"].Value = boletoCedente.id;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@id_cliente"].Value = strIdCliente;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@cnpj_cpf"].Value = Global.digitos(strCnpjCpf);
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@tipo_cadastro"].Value = Global.Cte.FIN.TipoCadastro.SISTEMA;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@editado_manual"].Value = Global.Cte.FIN.EditadoManual.NAO;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@usuario_cadastro"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@st_confirmacao_pendente"].Value = Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@st_boleto_ocorrencia_17"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_17"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				#endregion

				#region [ Monta texto p/ o log ]
				foreach (SqlParameter item in cmLancamentoInsertDevidoBoletoOcorrencia17.Parameters)
				{
					if (sbLog.Length > 0) sbLog.Append("; ");
					sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
				}
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoInsertDevidoBoletoOcorrencia17);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
				if (intRetorno != 1)
				{
					strMsgErro = "Falha ao tentar gerar o registro automático do lançamento do fluxo de caixa devido a boleto com ocorrência 17 (liquidação após baixa ou título não registrado)!!";
					return false;
				}
				#endregion

				#region [ Registra log com a inclusão dos dados ]
				strDescricaoLog = "Inserção do registro em t_FIN_FLUXO_CAIXA.id=" + intNsuLancamento.ToString() + " devido à ocorrência 17 (liquidação após baixa ou título não registrado) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
				Global.gravaLogAtividade(strDescricaoLog);
				finLog.usuario = usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_OCORRENCIA_17;
				finLog.natureza = Global.Cte.FIN.Natureza.CREDITO;
				finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
				finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
				finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
				finLog.id_registro_origem = intNsuLancamento;
				finLog.id_conta_corrente = boletoCedente.id_conta_corrente;
				finLog.id_plano_contas_empresa = boletoPlanoContasDestino.id_plano_contas_empresa;
				finLog.id_plano_contas_grupo = boletoPlanoContasDestino.id_plano_contas_grupo;
				finLog.id_plano_contas_conta = boletoPlanoContasDestino.id_plano_contas_conta;
				finLog.id_boleto_cedente = boletoCedente.id;
				finLog.id_cliente = strIdCliente;
				finLog.cnpj_cpf = Global.digitos(strCnpjCpf);
				finLog.descricao = strDescricaoLog;
				FinLogDAO.insere(usuario, finLog, ref strMsgErro);
				#endregion

				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaLancamentoDevidoBoletoOcorrencia22 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 22 (título com pagamento cancelado).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia22(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 22 (título com pagamento cancelado)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters["@ctrl_pagto_status"].Value = (byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_COM_PAGAMENTO_CANCELADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters["@st_boleto_pago_cheque"].Value = Global.Cte.FIN.StCampoFlag.FLAG_DESLIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters["@dt_ocorrencia_banco_boleto_pago_cheque"].Value = "";
				cmLancamentoUpdateDevidoBoletoOcorrencia22.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia22);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 22 (título com pagamento cancelado) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_22;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 22 (título com pagamento cancelado)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia23 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 23 (entrada do título em cartório).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia23(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 23 (entrada do título em cartório)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters["@st_boleto_ocorrencia_23"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_23"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmLancamentoUpdateDevidoBoletoOcorrencia23.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia23);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 23 (entrada do título em cartório) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_23;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 23 (entrada do título em cartório)!!";
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

		#region [ atualizaLancamentoDevidoBoletoOcorrencia34 ]
		/// <summary>
		/// Atualiza o lançamento do fluxo de caixa em decorrência do boleto retornado
		/// com ocorrência 34 (retirado de cartório e manutenção carteira).
		/// </summary>
		/// <param name="usuario">Usuário responsável pelo processamento do arquivo de retorno</param>
		/// <param name="idBoletoItem">Identificação do registro do boleto na tabela t_FIN_BOLETO_ITEM</param>
		/// <param name="dataOcorrenciaBanco">Data da ocorrência no banco</param>
		/// <param name="strMsgErro">Mensagem de erro no caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na atualização do lançamento do fluxo de caixa
		/// false: falha na atualização do lançamento do fluxo de caixa
		/// </returns>
		public static bool atualizaLancamentoDevidoBoletoOcorrencia34(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro do lançamento no fluxo de caixa devido a boleto com ocorrência 34 (retirado de cartório e manutenção carteira)";
			bool blnSucesso = false;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			String strDescricaoLog;
			FinLog finLog = new FinLog();
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaOriginal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixaEditado;
			DsDataSource.DtbFinFluxoCaixaRow rowFinFluxoCaixaEditado;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém registro do lançamento do fluxo de caixa associado ]
				dtbFinFluxoCaixaOriginal = obtemRegistroLancamentoByCtrlPagtoIdParcela(idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				#endregion

				#region [ Conseguiu localizar o lançamento do fluxo de caixa? ]
				if (dtbFinFluxoCaixaOriginal.Rows.Count == 0)
				{
					strMsgErro = "Não foi localizado o registro do lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				if (dtbFinFluxoCaixaOriginal.Rows.Count > 1)
				{
					strMsgErro = "Há mais de 1 registro de lançamento do fluxo de caixa associado ao módulo de controle de pagamentos (" + Global.retornaDescricaoCtrlPagtoModulo(Global.Cte.FIN.CtrlPagtoModulo.BOLETO) + ") registro id=" + idBoletoItem.ToString();
					return false;
				}
				#endregion

				rowFinFluxoCaixaOriginal = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaOriginal.Rows[0];

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters["@id"].Value = rowFinFluxoCaixaOriginal.id;
				cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters["@st_boleto_ocorrencia_34"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_34"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmLancamentoUpdateDevidoBoletoOcorrencia34.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.FIN.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoUpdateDevidoBoletoOcorrencia34);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					#region [ Registra log com a alteração dos dados ]
					dtbFinFluxoCaixaEditado = obtemRegistroLancamentoBD(rowFinFluxoCaixaOriginal.id);
					rowFinFluxoCaixaEditado = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixaEditado.Rows[0];
					foreach (DataColumn coluna in rowFinFluxoCaixaOriginal.Table.Columns)
					{
						if (!rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString().Equals(rowFinFluxoCaixaEditado[coluna.ColumnName].ToString()))
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(coluna.ColumnName + ": " + rowFinFluxoCaixaOriginal[coluna.ColumnName].ToString() + " => " + rowFinFluxoCaixaEditado[coluna.ColumnName].ToString());
						}
					}
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Atualização do registro em t_FIN_FLUXO_CAIXA.id=" + rowFinFluxoCaixaOriginal.id.ToString() + " devido à ocorrência 34 (retirado de cartório e manutenção carteira) do boleto t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_34;
						finLog.natureza = rowFinFluxoCaixaOriginal.natureza;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.FLUXO_CAIXA;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
						finLog.id_registro_origem = rowFinFluxoCaixaOriginal.id;
						finLog.id_conta_corrente = rowFinFluxoCaixaOriginal.id_conta_corrente;
						finLog.id_plano_contas_empresa = rowFinFluxoCaixaOriginal.id_plano_contas_empresa;
						finLog.id_plano_contas_grupo = rowFinFluxoCaixaOriginal.id_plano_contas_grupo;
						finLog.id_plano_contas_conta = rowFinFluxoCaixaOriginal.id_plano_contas_conta;
						finLog.id_boleto_cedente = rowFinFluxoCaixaOriginal.id_boleto_cedente;
						finLog.id_cliente = rowFinFluxoCaixaOriginal.id_cliente;
						finLog.cnpj_cpf = rowFinFluxoCaixaOriginal.cnpj_cpf;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
					#endregion

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
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!";
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

		#endregion
	}
}
