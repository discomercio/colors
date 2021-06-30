#region [ using ]
using System;
using System.Text;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Reflection;
using System.Configuration;
using System.Collections.Generic;
#endregion

namespace FinanceiroService
{
	public class Global
	{
		/*
		 * Esta classe possui um construtor estático que está implementado logo em seguida à seção de declarações
		*/

		#region [ Constantes ]
		public static class Cte
		{
			#region[ Versão do Aplicativo ]
			public static class Aplicativo
			{
				public const string NOME_OWNER = "Artven";
				public const string NOME_SISTEMA = "Financeiro Service";
				public static readonly string ID_SISTEMA_EVENTLOG = GetConfigurationValue("ServiceName");
				public const string VERSAO_NUMERO = "1.38";
				public const string VERSAO_DATA = "29.JUN.2021";
				public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
				public const string M_ID = NOME_SISTEMA + "  -  " + VERSAO;
				public const string M_DESCRICAO = "Serviço do Windows para execução automática de rotinas financeiras";
				public static string IDENTIFICADOR_AMBIENTE_OWNER = "";
				public static readonly string AMBIENTE_EXECUCAO = GetConfigurationValue("AmbienteExecucao");
			}
			#endregion

			#region[ Comentário sobre as versões ]
			/*================================================================================================
			 * v 1.00 - 10.09.2010 - por HHO
			 *		Início.
			 *		Este serviço do Windows realiza diversas rotinas automáticas.
			 *		A versão inicial contém apenas as rotinas de limpeza automática dos dados antigos
			 *		das tabelas de log.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.01 - 28.08.2013 - por HHO
			 *		Implementação de rotina para cancelamento automático de pedidos.
			 *		Regras para o cancelamento:
			 *			1) Pendente Cartão de Crédito: após 7 dias corridos consecutivos.
			 *			2) Crédito Ok Aguardando Depósito: após 9 dias corridos consecutivos.
			 *			3) Pendente Vendas: após 14 dias corridos consecutivos.
			 *		Lembrando que se um pedido tiver uma transição no status da análise de crédito, a contagem
			 *		deve ser reiniciada.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.02 - 24.09.2013 - por HHO
			 *		Alteração da rotina de cancelamento automático de pedidos p/ não cancelar pedidos que
			 *		estejam com status de pagamento 'Pago' ou 'Pago Parcial'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.03 - 14.07.2014 - por HHO
			 *		Alteração da rotina de cancelamento automático de pedidos p/ cancelar pedidos que estejam
			 *		com o campo transportadora preenchido, desde que a transportadora tenha sido selecionada
			 *		automaticamente pelo sistema com base no CEP.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.04 - 12.11.2014 - por HHO
			 *		Alteração dos prazos de cancelamento automático de pedidos:
			 *			CREDITO_OK_AGUARDANDO_DEPOSITO: 9 dias -> 7 dias
			 *			PENDENTE_VENDAS: 14 dias -> 10 dias
			 * -----------------------------------------------------------------------------------------------
			 * v 1.05 - 12.01.2015 - por HHO
			 *		Acerto dos parâmetros de conexão ao banco de dados para operar com o novo servidor.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.06 - 01.11.2015 - por HHO
			 *		Alteração do cancelamento automático de pedidos para gravar o código do motivo do cance-
			 *		lamento e o texto explicativo.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.07 - 30.05.2016 - por HHO
			 *		Implementação de ajustes para permitir a instalação de múltiplas instâncias (multi-
			 *		empresas).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.08 - 06.07.2016 - por HHO
			 *		Implementação do tratamento para envio das transações de análise antifraude para a
			 *		Clearsale.
			 *		Além disso, implementação de rotina para periodicamente consultar a Clearsale para obter
			 *		o resultado da análise antifraude. Em caso de aprovação, é realizada a captura da
			 *		transação de pagamento e em caso de reprovação, o cancelamento.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.09 - 10.07.2016 - por HHO
			 *		Ajustes nas rotinas que gravam logs (event viewer, arquivo log atividade e BD) para
			 *		que as informações sejam gravadas em situações com informações úteis, evitando um volume
			 *		muito grande de dados registrados, principalmente em rotinas que são executadas a cada
			 *		segundo.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10 - 11.07.2016 - por HHO
			 *		Ajustes nas rotinas que gravam log em arquivo (log atividade) para reduzir as mensagens
			 *		desnecessárias.
			 *		Alteração na rotina que envia pedidos para a Clearsale para obter o email a partir dos
			 *		dados armazenados em t_PAGTO_GW_PAG_PAYMENT ao invés do cadastro do cliente a fim de
			 *		evitar a situação em que o cadastro do cliente é editado p/ ficar sem endereço de email
			 *		e o pedido ser enviado p/ a Clearsale sem email.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11 - 12.07.2016 - por HHO
			 *		Alteração na rotina que processa o resultado da Clearsale para obter os comentários do
			 *		analista e gravar no bloco de notas.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.12 - 13.07.2016 - por HHO
			 *		Alteração na rotina que faz a requisição SetOrderAsReturned() na Clearsale para usar
			 *		o método enviaRequisicaoComRetry().
			 *		Transferência dos parâmetros da Braspag e Clearsale (Entity Code, URL's dos web services,
			 *		etc) declaradas em constantes para parâmetros dentro do arquivo de configuração. Esta
			 *		alteração tem o objetivo de reduzir os riscos de se causar danos por usar uma versão
			 *		compilada para 'release' em ambiente de homologação. O mesmo vale p/ a versão compilada
			 *		para 'debug' ser usada em ambiente de produção.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.13 - 13.07.2016 - por HHO
			 *		Correção de bug na rotina que atualiza periodicamente os parâmetros armazenados no banco
			 *		de dados fazendo nova leitura.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14 - 14.07.2016 - por HHO
			 *		Alteração nas requisições para a Braspag para usar o método enviaRequisicaoComRetry().
			 * -----------------------------------------------------------------------------------------------
			 * v 1.15 - 15.07.2016 - por HHO
			 *		Implementação do período de inatividade para o processamento de envio de email de alerta
			 *		sobre pedido novo aguardando tratamento da análise de crédito.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16 - 19.07.2016 - por HHO
			 *		Ajuste na consulta SQL da rotina de cancelamento automático de pedidos para tratar a
			 *		situação em que a soma dos pagamentos em cartão do pedido retorne NULL.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - 25.07.2016 - por HHO
			 *		Ajuste na rotina que envia mensagem de alerta sobre transações pendentes próximas do
			 *		cancelamento automático:
			 *			1) Tratamento para retirar a parte da hora na data de autorização.
			 *			2) Se for 6ªf, inclui na mensagem as transações que expiram no domingo.
			 *		Implementação da rotina que processa a captura automática de transação pendente devido
			 *		ao prazo final de cancelamento automático.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.18 - 16.09.2016 - por HHO
			 *		Implementação de tratamento para os dados recebidos pelo mecanismo de 2º post da Braspag.
			 *		O tratamento é específico para os boletos do e-commerce (Bradesco SPS), sendo realizado
			 *		o registro automático do pagamento no pedido e alteração do status da análise de crédito.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.19 - 10.11.2016 - por HHO
			 *		Alteração da rotina de cancelamento automático de pedidos para ignorar os pedidos das
			 *		lojas definidas no parâmetro 'CancelamentoAutomaticoPedidosLojasIgnoradas' do arquivo
			 *		de configuração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.20 - 18.01.2017 - por HHO
			 *		Implementação de tratamento para o auto-split de pedidos.
			 *		Para tanto, as rotinas de controle de estoque foram ajustadas para administrar o estoque
			 *		de acordo com a empresa definida no pedido e em t_ESTOQUE.id_nfe_emitente
			 *		Além disso, ao enviar os dados dos itens do pedido para a Clearsale, todos os itens
			 *		da família de pedidos são consolidados para serem enviados como se tratasse de um único
			 *		pedido e usando o nº do pedido-base.
			 *		O cancelamento automático de pedidos também foi ajustado p/ que todos os pedidos da
			 *		família de pedidos possam ser cancelados, já que antes a rotina tinha como premissa
			 *		que somente pedidos com análise de crédito ok poderiam estar splitados.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.21 - 15.05.2017 - por HHO
			 *		Correção de um bug em um cast para short que causava um exception:
			 *			pedidoItem.qtde = BD.readToShort(rowResultado["total_qtde"]);
			 *		Foi necessário fazer antes um cast para int para somente depois fazer o cast p/ short:
			 *			pedidoItem.qtde = (short)BD.readToInt(rowResultado["total_qtde"]);
			 * -----------------------------------------------------------------------------------------------
			 * v 1.22- 17.05.2017 - por HHO
			 *		Ajustes na rotina Clearsale.enviaNovasTransacoes() para tratar a possibilidade de ocorrer
			 *		exception em PedidoDAO.getPedidoConsolidadoFamilia().
			 *		Sem esse tratamento, quando ocorre um exception no referido método, o processamento de
			 *		envio de pedidos para a Clearsale é interrompido, ou seja, os pedidos posteriores ao
			 *		pedido problemático ficam pendentes indefinidamente.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.23 - 22.09.2017 - por HHO
			 *		Ajustes nas rotinas de envio de requisições para a Braspag e Clearsale para aumentar a
			 *		quantidade de tentativas de retry de 5 para 10 vezes e aumentar o intervalo entre as
			 *		tentativas de 1s para 5s. O ajuste foi realizado devido à frequência em que estão
			 *		ocorrendo falhas de timeout nas requisições, principalmente com a Braspag nas rotinas
			 *		que são executadas automaticamente no início do dia.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.24 - 21.10.2017 - por HHO
			 *		Ajustes no tratamento do fluxo do estorno devido ao início das operações com a Getnet,
			 *		já que esta opera de modo diferente que a Cielo. A Cielo processa o estorno de modo
			 *		online, retornando o resultado imediatamente. A Getnet primeiro responde que recebeu a
			 *		requisição e realiza o processamento posteriormente, podendo demorar entre D+1 e D+2 para
			 *		concluir o estorno. Portanto, no caso da Getnet, é necessário realizar consultas periódi-
			 *		cas posteriormente para verificar o resultado da requisição de estorno. Em um dos testes
			 *		em ambiente de produção, porém, o estorno foi processado no mesmo dia, aproximadamente
			 *		3h depois da requisição.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.25 - 24.10.2017 - por HHO
			 *		Implementação das seguintes correções:
			 *		1) Processamento dos estornos pendentes: quando o estorno ainda continua pendente, o
			 *			valor de retorno da rotina foi alterado para 'true' a fim de não ser tratado como
			 *			falha de processamento.
			 *		2) Processamento da captura da transação: quando ocorre falha na requisição de captura
			 *			(ex: timeout), foi inibido o processamento do pagamento no pedido, pois senão quando
			 *			a transação for capturada automaticamente no último dia do prazo de transações
			 *			pendentes de captura, ocorrerá duplicidade de registros de pagamento em
			 *			t_PEDIDO_PAGAMENTO e no histório de pagamento em t_FIN_PEDIDO_HIST_PAGTO (o registro
			 *			original criado na transação de autorização e outro criado na data da captura
			 *			automática).
			 *			Além disso, foi adicionado o envio de uma mensagem informativa alertando sobre a
			 *			falha na captura.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.26 - 12.02.2018 - por HHO
			 *		Ajuste para tornar o módulo compatível com o TLS v1.2 devido à desativação programada
			 *		da Braspag das versões do TLS v1.0 e v1.1
			 *		A compatibilização foi realizada através da alteração da versão do .NET Framework da
			 *		versão 4.5 para 4.6, pois na versão 4.5 o TLS 1.2 é suportado, mas necessita que o uso
			 *		dessa versão seja explicitamente declarado. Já na versão 4.6, o TLS 1.2 é suportado
			 *		por padrão.
			 *		Alterado o tempo de timeout das requisições Braspag e Clearsale de 100.000 ms (default)
			 *		para 180.000 ms (3 minutos). Alterado também a quantidade de tentativas de retry
			 *		automático, de 10 para 5.
			 *		Ajustes para excluir automaticamente os dados antigos da tabela t_MAGENTO_API_PEDIDO_XML,
			 *		que armazena as informações obtidas via API do Magento na operação de cadastramento
			 *		semi-automático de pedidos do e-commerce.
			 *		Ajustes para excluir automaticamente os dados antigos da tabela
			 *		t_ESTOQUE_VENDA_SALDO_DIARIO, que armazena as informações do custo médio de aquisição
			 *		dos produtos e que são usadas no cálculo do valor estimado da margem.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.27 - 22.05.2018 - por HHO
			 *		Inclusão de tratamento para o status 'APP' (Aprovação por Política) para o resultado
			 *		da Clearsale.
			 *		Implementação de tratamento para limpar os campos de session token que permanecem gravados
			 *		quando o usuário não encerra a sessão corretamente (tabela t_USUARIO, campos
			 *		SessionTokenModuloCentral e SessionTokenModuloLoja).
			 *		Implementação de rotina de limpeza dos arquivos salvos no servidor através da WebAPI
			 *		(UploadFile) para excluir automaticamente os seguintes tipos de arquivo:
			 *			1) Arquivos temporários (st_temporary_file = 1)
			 *			2) Arquivos com confirmação pendente (st_confirmation_required = 1 e
			 *				st_confirmation_ok = 0)
			 *			3) Arquivos com solicitação de exclusão agendada (st_delete_file = 1 e
			 *				dt_delete_file_scheduled_date < Now)
			 *		Além de excluir o arquivo do disco, apaga também do banco de dados, caso o conteúdo
			 *		esteja salvo em t_UPLOAD_FILE.file_content e t_UPLOAD_FILE.file_content_text
			 * -----------------------------------------------------------------------------------------------
			 * v 1.28 - 25.07.2018 - por HHO
			 *		Ajustes no processamento dos dados do Webhook da Braspag para tratar o boleto registrado
			 *		Bradesco (código de meio de pagamento: 585).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.29 - 28.02.2019 - por HHO
			 *		Ajustes para gravar a data em que o status de pagamento do pedido foi alterado.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.30 - 10.06.2019 - por HHO
			 *		Ajustes para enviar o número do cartão para a Clearsale somente se ele não estiver
			 *		mascarado para proteger os dados.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.31 - 16.07.2019 - por HHO
			 *		Implementação de tratamento para o novo meio de pagamento 'cartão (maquineta)'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.32 - 24.09.2019 - por HHO
             *      Ajustes para realização da limpeza automática da tabela t_CTRL_RELATORIO_USUARIO_X_PEDIDO
			 * -----------------------------------------------------------------------------------------------
			 * v 1.33 - 14.08.2020 - por HHO
			 *      Ajustes para tratar a memorização do endereço de cobrança no pedido, pois, a partir de
			 *      agora, ao invés de obter os dados do endereço no cadastro do cliente (t_CLIENTE), deve-se
			 *      usar os dados que estão gravados no próprio pedido. O tratamento que já ocorria com o
			 *      endereço de entrega deve passar a ser feito p/ o endereço de cobrança/cadastro.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.34 - 10.04.2021 - por HHO
			 *      Ajustes na lógica do processamento dos produtos vendidos sem presença no estoque para
			 *      definir critérios de prioridade para os pedidos.
			 *      Desenvolvimento de rotina para executar o processamento dos produtos vendidos sem presença
			 *      no estoque para que seja executada sob demanda a partir da sinalização feita através de
			 *      flag definida em parâmetro.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.35 - 18.04.2021 - por HHO
			 *      Ajustes na solicitação de execução sob demanda da rotina de processamento de produtos
			 *      vendidos sem presença no estoque para aceitar no parâmetro  opção que processa para todos
			 *      os códigos de id_nfe_emitente.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.36 - 19.04.2021 - por HHO
			 *      Ajustes na solicitação de execução sob demanda da rotina de processamento de produtos
			 *      vendidos sem presença no estoque para resetar o parâmetro logo após a sua leitura a fim
			 *      de minimizar o risco de problemas por acesso concorrente.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.37 - 21.05.2021 - por HHO
			 *      Ajustes em BraspagDAO.registraPagamentoNoPedido() para não alterar o status da análise de
			 *      crédito quando este estiver nos seguintes status:
			 *      CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO, CREDITO_OK_AGUARDANDO_DEPOSITO,
			 *      CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV e PENDENTE_PAGTO_ANTECIPADO_BOLETO
			 *      Esses status exigem que seja feita uma confirmação manual do pagamento, sendo que nesses
			 *      casos pode ter sido registrado antecipadamente o valor de uma parcela no pedido apenas
			 *      para indicar que um boleto foi emitido.
			 *      Implementação de tratamento para acesso concorrente nas operações com o banco de dados
			 *      que podem causar um problema grave. O tratamento se baseia em obter previamente o lock
			 *      exclusivo do(s) registro(s) através de um update que realiza o flip de um campo bit.
			 *      A ativação do tratamento de acesso concorrente é feita através do novo parâmetro no
			 *      arquivo de configuração: TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO
			 * -----------------------------------------------------------------------------------------------
			 * v 1.38 - 29.06.2021 - por HHO
			 *      Implementação de tratamento dos dados recebidos pelo post de notificação da Braspag, já
			 *      que o antigo serviço de 2º post será desativado em 01/07/2021. O novo post de notificação
			 *      informa dados diferentes que o post anterior, sendo que as principais são:
			 *          1) PaymentId: passa a informar o GUID referente ao Braspag Transaction ID e não mais
			 *             o NumPedido (nº do pedido informado em OrderId na requisição de pagamento).
			 *          2) CODPAGAMENTO: deixa de ser informado o código do meio de pagamento, portanto, agora
			 *             é necessário realizar um processamento usando o PaymentId p/ descobrir se o paga-
			 *             mento se refere a um boleto ou cartão.
			 *      Enquanto antes os dados eram armazenados nas tabelas t_BRASPAG_WEBHOOK e
			 *      t_BRASPAG_WEBHOOK_COMPLEMENTAR, o tratamento para o novo post de notificação armazena em
			 *      t_BRASPAG_WEBHOOK_V2 e t_BRASPAG_WEBHOOK_V2_COMPLEMENTAR
			 * -----------------------------------------------------------------------------------------------
			 * v 1.39 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.40 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.41 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.42 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.43 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.44 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.45 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.46 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.47 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.48 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.49 - XX.XX.20XX - por XXX
			 *      
			 * -----------------------------------------------------------------------------------------------
			 * v 1.50 - XX.XX.20XX - por XXX
			 *      
			 * ===============================================================================================
			 */
			#endregion

			#region [ FIN ]
			public static class FIN
			{
				#region [ ID_T_PARAMETRO ]
				public static class ID_T_PARAMETRO
				{
					public const string DT_HR_ULT_PROC_CLIENTES_EM_ATRASO = "FinSvc_DtHrUltProcClientesEmAtraso";
					public const string DT_HR_ULT_MANUTENCAO_ARQ_LOG_ATIVIDADE = "FinSvc_DtHrUltManutencaoArqLogAtividade";
					public const string DT_HR_ULT_MANUTENCAO_BD_LOG_ANTIGO = "FinSvc_DtHrUltManutencaoBdLogAntigo";
					public const string DT_HR_ULT_CANCELAMENTO_AUTOMATICO_PEDIDOS = "FinSvc_DtHrUltCancelamentoAutomaticoPedidos";
					public const string DT_HR_ULT_PROCESSAMENTO_BP_CS_ANTIFRAUDE_CLEARSALE = "FinSvc_DtHrUltProcessamentoBpCsAntifraudeClearsale";
					public const string FLAG_HABILITACAO_CANCELAMENTO_AUTOMATICO_PEDIDOS = "FinSvc_FlagHabilitacao_CancelamentoAutomaticoPedidos";
					public const string FLAG_HABILITACAO_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE = "FinSvc_FlagHabilitacao_ProcProdutosVendidosSemPresencaEstoque";
					public const string FLAG_EXECUCAO_SOLICITADA_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE = "FinSvc_FlagExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque";
					public const string CONSULTA_EXECUCAO_SOLICITADA_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE_EM_SEG = "FinSvc_ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg";
					public const string DT_HR_ULT_CONSULTA_EXECUCAO_SOLICITADA_PROC_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE = "FinSvc_DtHrUltConsultaExecucaoSolicitadaProcProdutosVendidosSemPresencaEstoque";
					public const string FLAG_HABILITACAO_BP_CS_ANTIFRAUDE_CLEARSALE = "FinSvc_BP_CS_Clearsale_FlagHabilitacao";
					public const string BP_CS_CLEARSALE_MAX_TENTATIVAS_TX_ANTIFRAUDE = "FinSvc_BP_CS_Clearsale_MaxTentativasTX";
					public const string BP_CS_CLEARSALE_TEMPO_MIN_ENTRE_TENTATIVAS_EM_SEG = "FinSvc_BP_CS_Clearsale_TempoMinEntreTentativasEmSeg";
					public const string BP_CS_CLEARSALE_TEMPO_ENTRE_PROCESSAMENTO_EM_SEG = "FinSvc_BP_CS_Clearsale_TempoEntreProcessamentoEmSeg";
					public const string BP_CS_CLEARSALE_TEMPO_MAX_CLIENTE_TOTALIZAR_PAGTO_EM_SEG = "FinSvc_BP_CS_Clearsale_TempoMaxClienteTotalizarPagtoEmSeg";
					public const string BP_CS_CLEARSALE_MAX_QTDE_FALHAS_CONSECUTIVAS_METODO_GETRETURNANALYSIS = "FinSvc_BP_CS_Clearsale_MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis";
					public const string FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE = "FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao";
					public const string BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE_HORARIO_INICIO = "FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio";
					public const string BP_CS_PROCESSAMENTO_PERIODO_INATIVIDADE_HORARIO_TERMINO = "FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino";
					public const string FLAG_HABILITACAO_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE = "FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao";
					public const string PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE_HORARIO_INICIO = "FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio";
					public const string PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO_PERIODO_INATIVIDADE_HORARIO_TERMINO = "FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino";
					public const string FLAG_HABILITACAO_BP_CS_BRASPAG_ATUALIZA_STATUS_TRANSACOES_PENDENTES = "FinSvc_BP_CS_Braspag_FlagHabilitacao_ProcAtualizaStatusTransacoesPendentes";
					public const string DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_ATUALIZA_STATUS_TRANSACOES_PENDENTES = "FinSvc_BP_CS_Braspag_DtHrUltProcAtualizaStatusTransacoesPendentes";
					public const string BP_CS_BRASPAG_PROCESSAMENTO_ATUALIZA_STATUS_TRANSACOES_PENDENTES_HORARIO = "FinSvc_BP_CS_Braspag_ProcessamentoAtualizaStatusTransacoesPendentes_Horario";
					public const string FLAG_HABILITACAO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = "FinSvc_BP_CS_Braspag_FlagHabilitacao_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto";
					public const string BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO_HORARIO = "FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario";
					public const string DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = "FinSvc_BP_CS_Braspag_DtHrUltProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto";
					public const string BP_CS_BRASPAG_DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = "FinSvc_BP_CS_Braspag_DestinatarioMsgAlertaTransacoesPendentesProxCancelAuto";
					public const string FLAG_HABILITACAO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = "FinSvc_BP_CS_Braspag_FlagHabilitacao_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto";
					public const string BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO_HORARIO = "FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario";
					public const string DT_HR_ULT_PROCESSAMENTO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = "FinSvc_BP_CS_Braspag_DtHrUltProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto";
					public const string BP_CS_BRASPAG_DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = "FinSvc_BP_CS_Braspag_DestinatarioMsgAlertaCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto";
					public const string FLAG_HABILITACAO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = "FinSvc_FlagHabilitacao_ProcEnviarEmailAlertaPedidoNovoAnaliseCredito";
					public const string DT_HR_ULT_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = "FinSvc_DtHrUltProcEnviarEmailAlertaPedidoNovoAnaliseCredito";
					public const string DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = "FinSvc_DestinatarioMsgAlertaPedidoNovoAnaliseCredito";
					public const string DT_HR_ULT_PROCESSAMENTO_WEBHOOK_BRASPAG = "FinSvc_DtHrUltProcWebhookBraspag";
					public const string DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG = "FinSvc_DestinatarioMsgAlertaWebhookBraspag";
					public const string FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG = "FinSvc_FlagHabilitacao_ProcWebhookBraspag";
					public const string FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE = "FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao";
					public const string PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE_HORARIO_INICIO = "FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio";
					public const string PROCESSAMENTO_WEBHOOK_BRASPAG_PERIODO_INATIVIDADE_HORARIO_TERMINO = "FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino";
					public const string DT_HR_ULT_PROCESSAMENTO_WEBHOOK_BRASPAG_V2 = "FinSvc_DtHrUltProcWebhookBraspagV2";
					public const string DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG_V2 = "FinSvc_DestinatarioMsgAlertaWebhookBraspagV2";
					public const string FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG_V2 = "FinSvc_FlagHabilitacao_ProcWebhookBraspagV2";
					public const string FLAG_HABILITACAO_PROCESSAMENTO_WEBHOOK_BRASPAG_V2_PERIODO_INATIVIDADE = "FinSvc_ProcessamentoWebhookBraspagV2_PeriodoInatividade_FlagHabilitacao";
					public const string PROCESSAMENTO_WEBHOOK_BRASPAG_V2_PERIODO_INATIVIDADE_HORARIO_INICIO = "FinSvc_ProcessamentoWebhookBraspagV2_PeriodoInatividade_HorarioInicio";
					public const string PROCESSAMENTO_WEBHOOK_BRASPAG_V2_PERIODO_INATIVIDADE_HORARIO_TERMINO = "FinSvc_ProcessamentoWebhookBraspagV2_PeriodoInatividade_HorarioTermino";
					public const string FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES = "FinSvc_BP_CS_Processamento_EstornosPendentes_FlagHabilitacao";
					public const string FLAG_HABILITACAO_BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE = "FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao";
					public const string BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE_HORARIO_INICIO = "FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio";
					public const string BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_PERIODO_INATIVIDADE_HORARIO_TERMINO = "FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino";
					public const string BP_CS_PROCESSAMENTO_ESTORNOS_PENDENTES_TEMPO_ENTRE_PROCESSAMENTO_EM_SEG = "FinSvc_BP_CS_Processamento_EstornosPendentes_TempoEntreProcessamentoEmSeg";
					public const string DT_HR_ULT_PROCESSAMENTO_ESTORNOS_PENDENTES = "FinSvc_DtHrUltProcEstornosPendentes";
					public const string DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS = "FinSvc_DestinatarioMsgAlertaEstornosPendentesAbortados";
					public const string ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS = "FinSvc_EstornosPendentes_PrazoMaximoVerificacaoEmDias";
					public const string FLAG_HABILITACAO_LIMPEZA_SESSION_TOKEN = "FinSvc_FlagHabilitacao_LimpezaSessionToken";
					public const string LIMPEZA_SESSION_TOKEN_HORARIO = "FinSvc_LimpezaSessionToken_Horario";
					public const string DT_HR_ULT_LIMPEZA_SESSION_TOKEN = "FinSvc_DtHrUltLimpezaSessionToken";
					public const string FLAG_HABILITACAO_UPLOAD_FILE_MANUTENCAO_ARQUIVOS = "FinSvc_FlagHabilitacao_UploadFile_ManutencaoArquivos";
					public const string UPLOAD_FILE_MANUTENCAO_ARQUIVOS_HORARIO = "FinSvc_UploadFile_ManutencaoArquivos_Horario";
					public const string DT_HR_ULT_UPLOAD_FILE_MANUTENCAO_ARQUIVOS = "FinSvc_DtHrUltUploadFileManutencaoArquivos";
					public const string ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS = "Flag_Pedido_MemorizacaoCompletaEnderecos";
				}
				#endregion

				#region [ CodBoletoArqRetornoStProcessamento ]
				public static class CodBoletoArqRetornoStProcessamento
				{
					public const short EM_PROCESSAMENTO = 1;
					public const short SUCESSO = 2;
					public const short FALHA = 3;
				}
				#endregion

				#region [ CtrlPagtoModulo ]
				public class CtrlPagtoModulo
				{
					public const byte BOLETO = 1;
					public const byte CHEQUE = 2;
					public const byte VISA = 3;
					public const byte BRASPAG_CARTAO = 4;
					public const byte BRASPAG_CLEARSALE = 5;
					public const byte BRASPAG_WEBHOOK = 6;
					public const byte BRASPAG_WEBHOOK_V2 = 7;
					public const byte PAGTO_COMISSAO_INDICADOR = 11;
				}
				#endregion

				#region [ NSU ]
				public class NSU
				{
					public const String T_FIN_FLUXO_CAIXA = "t_FIN_FLUXO_CAIXA";
					public const string T_PAGTO_GW_AF = "T_PAGTO_GW_AF";
					public const string T_PAGTO_GW_AF_ITEM = "T_PAGTO_GW_AF_ITEM";
					public const string T_PAGTO_GW_AF_PAYMENT = "T_PAGTO_GW_AF_PAYMENT";
					public const string T_PAGTO_GW_AF_PHONE = "T_PAGTO_GW_AF_PHONE";
					public const string T_PAGTO_GW_AF_SESSIONID = "T_PAGTO_GW_AF_SESSIONID";
					public const string T_PAGTO_GW_AF_XML = "T_PAGTO_GW_AF_XML";
					public const string T_PAGTO_GW_PAG = "T_PAGTO_GW_PAG";
					public const string T_PAGTO_GW_PAG_ERROR = "T_PAGTO_GW_PAG_ERROR";
					public const string T_PAGTO_GW_PAG_OP_COMPLEMENTAR = "T_PAGTO_GW_PAG_OP_COMPLEMENTAR";
					public const string T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML = "T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML";
					public const string T_PAGTO_GW_PAG_PAYMENT = "T_PAGTO_GW_PAG_PAYMENT";
					public const string T_PAGTO_GW_PAG_XML = "T_PAGTO_GW_PAG_XML";
					public const string T_PAGTO_GW_EMAIL_CTRL = "T_PAGTO_GW_EMAIL_CTRL";
					public const string T_FINSVC_LOG = "T_FINSVC_LOG";
					public const string T_PAGTO_GW_AF_OP_COMPLEMENTAR = "T_PAGTO_GW_AF_OP_COMPLEMENTAR";
					public const string T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML = "T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML";
					public const string T_BRASPAG_WEBHOOK = "T_BRASPAG_WEBHOOK";
					public const string T_BRASPAG_WEBHOOK_COMPLEMENTAR = "T_BRASPAG_WEBHOOK_COMPLEMENTAR";
					public const string T_BRASPAG_WEBHOOK_V2 = "T_BRASPAG_WEBHOOK_V2";
					public const string T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR = "T_BRASPAG_WEBHOOK_V2_COMPLEMENTAR";
				}
				#endregion

				#region [ Natureza ]
				public class Natureza
				{
					public const char CREDITO = 'C';
					public const char DEBITO = 'D';
				}
				#endregion

				#region [ TipoCadastro ]
				public class TipoCadastro
				{
					public const char MANUAL = 'M';
					public const char SISTEMA = 'S';
				}
				#endregion

				#region [ EditadoManual ]
				public class EditadoManual
				{
					public const char NAO = 'N';
					public const char SIM = 'S';
				}
				#endregion

				#region [ Modulo ]
				public class Modulo
				{
					public const String FLUXO_CAIXA = "FLC";
					public const String BOLETO = "BOL";
					public const String CHEQUE = "CHQ";
					public const String VISA = "VIS";
					public const String HIST_PAGTO_PEDIDO = "HPP";
					public const String SERASA_RECIPROCIDADE = "SER";
					public const String FINANCEIRO_SERVICE = "FSV";
				}
				#endregion

				#region [ Códigos de tabelas de origem ]
				public class TabelaOrigem
				{
					public const byte T_FIN_FLUXO_CAIXA = 1;
					public const byte T_FIN_NF_PARCELA_PAGTO = 2;
					public const byte T_FIN_BOLETO = 3;
					public const byte T_FIN_BOLETO_CEDENTE = 4;
					public const byte T_FIN_BOLETO_OCORRENCIA = 5;
					public const byte T_FIN_BOLETO_ITEM = 6;
					public const byte T_SERASA_CLIENTE = 7;
					public const byte T_SERASA_TITULO_MOVIMENTO = 8;
				}
				#endregion

				#region [ TamanhoCampo ]
				public class TamanhoCampo
				{
					public const int CONTA_CORRENTE_ID = 1;
					public const int CONTA_CORRENTE_CONTA = 12;
					public const int PLANO_CONTAS_EMPRESA = 1;
					public const int PLANO_CONTAS_GRUPO = 2;
					public const int PLANO_CONTAS_CONTA = 4;
					public const int FLUXO_CAIXA_DESCRICAO = 40;
					public const int FIN_LOG_DESCRICAO = 7500;  // Para prevenir erro: "exceeds the maximum number of bytes per row (8060)"
					public const int COMENTARIO_OCORRENCIA_TRATADA = 240;
				}
				#endregion

				#region [ LogOperacao - Códigos de operação para o log ]
				public class LogOperacao
				{
					public const String FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_ECOMMERCE = "FCInsBolEC";
				}
				#endregion
			}
			#endregion

			#region [ TipoPessoa ]
			public static class TipoPessoa
			{
				public const string PJ = "PJ";
				public const string PF = "PF";
			}
			#endregion

			#region [ FluxoXml ]
			public sealed class FluxoXml
			{
				private readonly String value;

				public static readonly FluxoXml TX = new FluxoXml("TX");
				public static readonly FluxoXml RX = new FluxoXml("RX");

				private FluxoXml(string value)
				{
					this.value = value;
				}

				public string GetValue()
				{
					return value;
				}

				public override string ToString()
				{
					return value;
				}
			}
			#endregion

			#region [ Braspag ]
			public static class Braspag
			{
				#region [ Endereço Web Service ]
				public static readonly string WS_ENDERECO_PAGADOR_TRANSACTION = GetConfigurationValue("BraspagWsEnderecoPagadorTransaction");
				public static readonly string WS_ENDERECO_PAGADOR_QUERY = GetConfigurationValue("BraspagWsEnderecoPagadorQuery");
				#endregion

				#region [ Parâmetros ]
				// The Timeout applies to the entire request and response, not individually to the GetRequestStream and GetResponse method calls
				public static readonly int REQUEST_TIMEOUT_EM_MS = 3 * 60 * 1000;
				#endregion

				#region [ Pagador ]
				public static class Pagador
				{
					public const string Version = "1.0";
					public const int PRAZO_CAPTURA_EM_DIAS_CORRIDOS = 5;

					#region [ GlobalStatus ]
					public sealed class GlobalStatus
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly GlobalStatus INDEFINIDA = new GlobalStatus("G00", "Indefinida");
						public static readonly GlobalStatus CAPTURADA = new GlobalStatus("G01", "Capturada");
						public static readonly GlobalStatus AUTORIZADA = new GlobalStatus("G02", "Autorizada");
						public static readonly GlobalStatus NAO_AUTORIZADA = new GlobalStatus("G03", "Não Autorizada");
						public static readonly GlobalStatus CAPTURA_CANCELADA = new GlobalStatus("G04", "Captura Cancelada");
						public static readonly GlobalStatus ESTORNADA = new GlobalStatus("G05", "Estornada");
						public static readonly GlobalStatus AGUARDANDO_RESPOSTA = new GlobalStatus("G06", "Aguardando Resposta");
						public static readonly GlobalStatus ERRO_DESQUALIFICANTE = new GlobalStatus("G07", "Erro Desqualificante");
						public static readonly GlobalStatus ESTORNO_PENDENTE = new GlobalStatus("G13", "Estorno Pendente");

						private GlobalStatus(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public override string ToString()
						{
							return value;
						}

						public string GetValue()
						{
							return value;
						}

						public string GetDescription()
						{
							return description;
						}

						public static string GetDescription(string codigo)
						{
							if (codigo == null) return "";

							Type type = typeof(GlobalStatus);
							foreach (var p in type.GetFields())
							{
								var v = p.GetValue(null);
								if (codigo.Equals(((GlobalStatus)v).GetValue())) return ((GlobalStatus)v).GetDescription();
							}

							if (codigo.Trim().Length > 0) return "Código Desconhecido: " + codigo;
							return "";
						}
					}
					#endregion

					#region [ PaymentDataResponseStatus ]
					public sealed class PaymentDataResponseStatus
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly PaymentDataResponseStatus CAPTURADA = new PaymentDataResponseStatus("0", "Capturada");
						public static readonly PaymentDataResponseStatus AUTORIZADA = new PaymentDataResponseStatus("1", "Autorizada");
						public static readonly PaymentDataResponseStatus NAO_AUTORIZADA = new PaymentDataResponseStatus("2", "Não Autorizada");
						public static readonly PaymentDataResponseStatus ERRO_DESQUALIFICANTE = new PaymentDataResponseStatus("3", "Erro Desqualificante");
						public static readonly PaymentDataResponseStatus AGUARDANDO_RESPOSTA = new PaymentDataResponseStatus("4", "Aguardando Resposta");

						private PaymentDataResponseStatus(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public override string ToString()
						{
							return value;
						}

						public string GetValue()
						{
							return value;
						}

						public string GetDescription()
						{
							return description;
						}
					}
					#endregion

					#region [ CaptureCreditCardTransactionResponseStatus ]
					public sealed class CaptureCreditCardTransactionResponseStatus
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly CaptureCreditCardTransactionResponseStatus CAPTURE_CONFIRMED = new CaptureCreditCardTransactionResponseStatus("0", "Capture Confirmed");
						public static readonly CaptureCreditCardTransactionResponseStatus CAPTURE_DENIED = new CaptureCreditCardTransactionResponseStatus("2", "Capture Denied");

						private CaptureCreditCardTransactionResponseStatus(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public override string ToString()
						{
							return value;
						}

						public string GetValue()
						{
							return value;
						}

						public string GetDescription()
						{
							return description;
						}
					}
					#endregion

					#region [ VoidCreditCardTransactionResponseStatus ]
					public sealed class VoidCreditCardTransactionResponseStatus
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly VoidCreditCardTransactionResponseStatus VOID_CONFIRMED = new VoidCreditCardTransactionResponseStatus("0", "Void Confirmed");
						public static readonly VoidCreditCardTransactionResponseStatus VOID_DENIED = new VoidCreditCardTransactionResponseStatus("1", "Void Denied");
						public static readonly VoidCreditCardTransactionResponseStatus INVALID_TRANSACTION = new VoidCreditCardTransactionResponseStatus("2", "Invalid Transaction");

						private VoidCreditCardTransactionResponseStatus(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public override string ToString()
						{
							return value;
						}

						public string GetValue()
						{
							return value;
						}

						public string GetDescription()
						{
							return description;
						}
					}
					#endregion

					#region [ RefundCreditCardTransactionResponseStatus ]
					public sealed class RefundCreditCardTransactionResponseStatus
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly RefundCreditCardTransactionResponseStatus REFUND_CONFIRMED = new RefundCreditCardTransactionResponseStatus("0", "Refund Confirmed");
						public static readonly RefundCreditCardTransactionResponseStatus REFUND_DENIED = new RefundCreditCardTransactionResponseStatus("1", "Refund Denied");
						public static readonly RefundCreditCardTransactionResponseStatus INVALID_TRANSACTION = new RefundCreditCardTransactionResponseStatus("2", "Invalid Transaction");
						public static readonly RefundCreditCardTransactionResponseStatus REFUND_ACCEPTED = new RefundCreditCardTransactionResponseStatus("3", "Refund Accepted");

						private RefundCreditCardTransactionResponseStatus(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public override string ToString()
						{
							return value;
						}

						public string GetValue()
						{
							return value;
						}

						public string GetDescription()
						{
							return description;
						}
					}
					#endregion

					#region [ GetTransactionDataResponseStatus ]
					public sealed class GetTransactionDataResponseStatus
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly GetTransactionDataResponseStatus INDEFINIDA = new GetTransactionDataResponseStatus("0", "Indefinida");
						public static readonly GetTransactionDataResponseStatus CAPTURADA = new GetTransactionDataResponseStatus("1", "Capturada");
						public static readonly GetTransactionDataResponseStatus AUTORIZADA = new GetTransactionDataResponseStatus("2", "Autorizada");
						public static readonly GetTransactionDataResponseStatus NAO_AUTORIZADA = new GetTransactionDataResponseStatus("3", "Não Autorizada");
						public static readonly GetTransactionDataResponseStatus CAPTURA_CANCELADA = new GetTransactionDataResponseStatus("4", "Captura Cancelada");
						public static readonly GetTransactionDataResponseStatus ESTORNADA = new GetTransactionDataResponseStatus("5", "Estornada");
						public static readonly GetTransactionDataResponseStatus AGUARDANDO_RESPOSTA = new GetTransactionDataResponseStatus("6", "Aguardando Resposta");
						public static readonly GetTransactionDataResponseStatus ERRO_DESQUALIFICANTE = new GetTransactionDataResponseStatus("7", "Erro Desqualificante");

						private GetTransactionDataResponseStatus(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public override string ToString()
						{
							return value;
						}

						public string GetValue()
						{
							return value;
						}

						public string GetDescription()
						{
							return description;
						}
					}
					#endregion

					#region [ OperacaoFinalizacao ]
					public sealed class OperacaoFinalizacao
					{
						// Type safe enum pattern
						private readonly String value;

						public static readonly OperacaoFinalizacao CAPTURA = new OperacaoFinalizacao("CAP");
						public static readonly OperacaoFinalizacao CANCELAMENTO = new OperacaoFinalizacao("CAN");

						private OperacaoFinalizacao(string value)
						{
							this.value = value;
						}

						public string GetValue()
						{
							return value;
						}

						public override string ToString()
						{
							return value;
						}
					}
					#endregion

					#region [ OperacaoRegistraPagtoPedido ]
					public sealed class OperacaoRegistraPagtoPedido
					{
						// Type safe enum pattern
						private readonly String value;
						private readonly String description;

						public static readonly OperacaoRegistraPagtoPedido CAPTURA = new OperacaoRegistraPagtoPedido("CAP", "Captura");
						public static readonly OperacaoRegistraPagtoPedido AUTORIZACAO = new OperacaoRegistraPagtoPedido("AUT", "Autorização");
						public static readonly OperacaoRegistraPagtoPedido CANCELAMENTO = new OperacaoRegistraPagtoPedido("CAN", "Cancelamento");
						public static readonly OperacaoRegistraPagtoPedido ESTORNO = new OperacaoRegistraPagtoPedido("EST", "Estorno");

						private OperacaoRegistraPagtoPedido(string value, string description)
						{
							this.value = value;
							this.description = description;
						}

						public string GetDescription()
						{
							return description;
						}

						public string GetValue()
						{
							return value;
						}

						public override string ToString()
						{
							return value;
						}
					}
					#endregion

					#region [ Transacao ]
					public sealed class Transacao
					{
						// Type safe enum pattern
						private readonly string methodName;
						private readonly string codOpLog;
						private readonly string enderecoWebService;
						private readonly string soapAction;

						public static readonly Transacao AuthorizeTransaction = new Transacao("AuthorizeTransaction", "Auth", WS_ENDERECO_PAGADOR_TRANSACTION, "https://www.pagador.com.br/webservice/pagador/AuthorizeTransaction");
						public static readonly Transacao CaptureCreditCardTransaction = new Transacao("CaptureCreditCardTransaction", "Capt_CCT", WS_ENDERECO_PAGADOR_TRANSACTION, "https://www.pagador.com.br/webservice/pagador/CaptureCreditCardTransaction");
						public static readonly Transacao VoidCreditCardTransaction = new Transacao("VoidCreditCardTransaction", "Void_CCT", WS_ENDERECO_PAGADOR_TRANSACTION, "https://www.pagador.com.br/webservice/pagador/VoidCreditCardTransaction");
						public static readonly Transacao RefundCreditCardTransaction = new Transacao("RefundCreditCardTransaction", "Refund_CCT", WS_ENDERECO_PAGADOR_TRANSACTION, "https://www.pagador.com.br/webservice/pagador/RefundCreditCardTransaction");
						public static readonly Transacao GetTransactionData = new Transacao("GetTransactionData", "Get_TD", WS_ENDERECO_PAGADOR_QUERY, "https://www.pagador.com.br/query/pagadorquery/GetTransactionData");
						public static readonly Transacao GetOrderIdData = new Transacao("GetOrderIdData", "Get_OID", WS_ENDERECO_PAGADOR_QUERY, "https://www.pagador.com.br/query/pagadorquery/GetOrderIdData");
						public static readonly Transacao GetOrderData = new Transacao("GetOrderData", "Get_OD", WS_ENDERECO_PAGADOR_QUERY, "https://www.pagador.com.br/query/pagadorquery/GetOrderData");
						public static readonly Transacao GetBoletoData = new Transacao("GetBoletoData", "Get_BOLD", WS_ENDERECO_PAGADOR_QUERY, "https://www.pagador.com.br/query/pagadorquery/GetBoletoData");

						private Transacao(string methodName, string codOpLog, string enderecoWebService, string soapAction)
						{
							this.methodName = methodName;
							this.codOpLog = codOpLog;
							this.enderecoWebService = enderecoWebService;
							this.soapAction = soapAction;
						}

						public string GetMethodName()
						{
							return methodName;
						}

						public string GetEnderecoWebService()
						{
							return enderecoWebService;
						}

						public string GetSoapAction()
						{
							return soapAction;
						}

						public string GetCodOpLog()
						{
							return codOpLog;
						}

						public override string ToString()
						{
							return codOpLog;
						}
					}
					#endregion
				}
				#endregion

				#region [ Bandeira ]
				public sealed class Bandeira
				{
					// Type safe enum pattern
					private readonly String value;
					private readonly String description;

					public static readonly Bandeira VISA = new Bandeira("visa", "Visa");
					public static readonly Bandeira MASTERCARD = new Bandeira("mastercard", "Mastercard");
					public static readonly Bandeira AMEX = new Bandeira("amex", "Amex");
					public static readonly Bandeira ELO = new Bandeira("elo", "Elo");
					public static readonly Bandeira DINERS = new Bandeira("diners", "Diners");
					public static readonly Bandeira DISCOVER = new Bandeira("discover", "Discover");
					public static readonly Bandeira AURA = new Bandeira("aura", "Aura");
					public static readonly Bandeira JCB = new Bandeira("jcb", "JCB");
					public static readonly Bandeira CELULAR = new Bandeira("celular", "Celular");

					private Bandeira(string value, string description)
					{
						this.value = value;
						this.description = description;
					}

					public override string ToString()
					{
						return value;
					}

					public string GetValue()
					{
						return value;
					}

					public string GetDescription()
					{
						return description;
					}

					public static string GetDescription(string bandeira)
					{
						if (bandeira == null) return "";

						Type type = typeof(Bandeira);
						foreach (var p in type.GetFields())
						{
							var v = p.GetValue(null);
							if (bandeira.Equals(((Bandeira)v).GetValue())) return ((Bandeira)v).GetDescription();
						}

						if (bandeira.Trim().Length > 0) return "Bandeira Desconhecida: " + bandeira;
						return "";
					}
				}
				#endregion

				#region [ PaymentMethod ]
				public sealed class PaymentMethod
				{
					// Type safe enum pattern
					private readonly String value;
					private readonly String description;

					public static readonly PaymentMethod Boleto_Bradesco = new PaymentMethod("06", "Boleto Bradesco");
					public static readonly PaymentMethod Boleto_CaixaEconomicaFederal = new PaymentMethod("07", "Boleto Caixa Econômica Federal");
					public static readonly PaymentMethod Boleto_HSBC = new PaymentMethod("08", "Boleto HSBC");
					public static readonly PaymentMethod Boleto_BancoBrasil = new PaymentMethod("09", "Boleto Banco do Brasil");
					public static readonly PaymentMethod Boleto_Real_ABN_AMRO = new PaymentMethod("10", "Boleto Real ABN AMRO");
					public static readonly PaymentMethod Boleto_Citibank = new PaymentMethod("13", "Boleto Citibank");
					public static readonly PaymentMethod Boleto_Itau = new PaymentMethod("14", "Boleto Itaú");
					public static readonly PaymentMethod Cielo_Visa_Electron = new PaymentMethod("123", "Cielo Visa Electron");
					public static readonly PaymentMethod Boleto_Santander = new PaymentMethod("124", "Boleto Santander");
					public static readonly PaymentMethod Cielo_Visa = new PaymentMethod("500", "Cielo Visa");
					public static readonly PaymentMethod Cielo_Mastercard = new PaymentMethod("501", "Cielo Mastercard");
					public static readonly PaymentMethod Cielo_Amex = new PaymentMethod("502", "Cielo Amex");
					public static readonly PaymentMethod Cielo_Diners = new PaymentMethod("503", "Cielo Diners");
					public static readonly PaymentMethod Cielo_ELO = new PaymentMethod("504", "Cielo ELO");
					public static readonly PaymentMethod Banorte_Visa = new PaymentMethod("505", "Banorte Visa");
					public static readonly PaymentMethod Banorte_Mastercard = new PaymentMethod("506", "Banorte Mastercard");
					public static readonly PaymentMethod Banorte_Diners = new PaymentMethod("507", "Banorte Diners");
					public static readonly PaymentMethod Banorte_Amex = new PaymentMethod("508", "Banorte Amex");
					public static readonly PaymentMethod Redecard_Webservice_Visa = new PaymentMethod("509", "Redecard Webservice Visa");
					public static readonly PaymentMethod Redecard_Webservice_Mastercard = new PaymentMethod("510", "Redecard Webservice Mastercard");
					public static readonly PaymentMethod Redecard_Webservice_Diners = new PaymentMethod("511", "Redecard Webservice Diners");
					public static readonly PaymentMethod PagosOnLine_Visa = new PaymentMethod("512", "PagosOnLine Visa");
					public static readonly PaymentMethod PagosOnLine_Mastercard = new PaymentMethod("513", "PagosOnLine Mastercard");
					public static readonly PaymentMethod PagosOnLine_Amex = new PaymentMethod("514", "PagosOnLine Amex");
					public static readonly PaymentMethod PagosOnLine_Diners = new PaymentMethod("515", "PagosOnLine Diners");
					public static readonly PaymentMethod Banorte_Cargos_Automaticos_Visa = new PaymentMethod("520", "Banorte Cargos Automáticos Visa");
					public static readonly PaymentMethod Banorte_Cargos_Automaticos_Mastercard = new PaymentMethod("521", "Banorte Cargos Automáticos Mastercard");
					public static readonly PaymentMethod Amex_2P = new PaymentMethod("523", "Amex 2P");
					public static readonly PaymentMethod Sitef_Visa = new PaymentMethod("524", "Sitef Visa");
					public static readonly PaymentMethod SiTef_Mastercard = new PaymentMethod("525", "SiTef Mastercard");
					public static readonly PaymentMethod SiTef_Amex = new PaymentMethod("526", "SiTef Amex");
					public static readonly PaymentMethod SiTef_Diners = new PaymentMethod("527", "SiTef Diners");
					public static readonly PaymentMethod SiTef_Hipercard = new PaymentMethod("528", "SiTef Hipercard");
					public static readonly PaymentMethod SiTef_Leader = new PaymentMethod("529", "SiTef Leader");
					public static readonly PaymentMethod SiTef_Aura = new PaymentMethod("530", "SiTef Aura");
					public static readonly PaymentMethod SiTef_Santander_Visa = new PaymentMethod("531", "SiTef Santander Visa");
					public static readonly PaymentMethod SiTef_Santander_Mastercard = new PaymentMethod("532", "SiTef Santander Mastercard");
					public static readonly PaymentMethod OneBuy = new PaymentMethod("533", "OneBuy");
					public static readonly PaymentMethod Sub1_Visa = new PaymentMethod("535", "Sub1 Visa");
					public static readonly PaymentMethod Sub1_Mastercard = new PaymentMethod("536", "Sub1 Mastercard");
					public static readonly PaymentMethod Sub1_Amex = new PaymentMethod("537", "Sub1 Amex");
					public static readonly PaymentMethod Sub1_Diners = new PaymentMethod("538", "Sub1 Diners");
					public static readonly PaymentMethod Sub1_Naranja = new PaymentMethod("540", "Sub1 Naranja");
					public static readonly PaymentMethod Sub1_Nevada = new PaymentMethod("541", "Sub1 Nevada");
					public static readonly PaymentMethod Sub1_Cabal = new PaymentMethod("542", "Sub1 Cabal");
					public static readonly PaymentMethod Cielo_Discover = new PaymentMethod("543", "Cielo Discover");
					public static readonly PaymentMethod Cielo_JCB = new PaymentMethod("544", "Cielo JCB");
					public static readonly PaymentMethod Cielo_Aura = new PaymentMethod("545", "Cielo Aura");
					public static readonly PaymentMethod Redecard_Webservice_Hipercard = new PaymentMethod("548", "Redecard Webservice Hipercard");
					public static readonly PaymentMethod CredSystem = new PaymentMethod("550", "CredSystem");
					public static readonly PaymentMethod Boleto_Caixa_SIGCB = new PaymentMethod("551", "Boleto Caixa SIGCB");
					public static readonly PaymentMethod Cielo_Mastercard_Debito = new PaymentMethod("552", "Cielo Mastercard Débito");
					public static readonly PaymentMethod Credibanco_Visa = new PaymentMethod("559", "Credibanco Visa");
					public static readonly PaymentMethod Credibanco_Mastercard = new PaymentMethod("560", "Credibanco Mastercard");
					public static readonly PaymentMethod Credibanco_Credential = new PaymentMethod("561", "Credibanco Credential");
					public static readonly PaymentMethod Credibanco_Diners = new PaymentMethod("562", "Credibanco Diners");
					public static readonly PaymentMethod Credibanco_Amex = new PaymentMethod("563", "Credibanco Amex");
					public static readonly PaymentMethod DM_Card = new PaymentMethod("564", "DM Card");
					public static readonly PaymentMethod Credz = new PaymentMethod("565", "Credz");
					public static readonly PaymentMethod Transferencia_Eletronica_Bradesco_SPS = new PaymentMethod("567", "Transferência Eletrônica Bradesco SPS");
					public static readonly PaymentMethod Boleto_Bradesco_SPS = new PaymentMethod("568", "Boleto Bradesco SPS");
					public static readonly PaymentMethod Boleto_Registrado_Bradesco = new PaymentMethod("585", "Boleto Registrado Bradesco");
					public static readonly PaymentMethod SafetyPay_Express = new PaymentMethod("569", "SafetyPay Express");
					public static readonly PaymentMethod EPay = new PaymentMethod("570", "EPay");
					public static readonly PaymentMethod Banorte_V2_Visa = new PaymentMethod("572", "Banorte V2 Visa");
					public static readonly PaymentMethod Banorte_V2_Mastercard = new PaymentMethod("573", "Banorte V2 Mastercard");
					public static readonly PaymentMethod Banorte_V2_Cargos_Automaticos_Visa = new PaymentMethod("574", "Banorte V2 Cargos Automáticos Visa");
					public static readonly PaymentMethod Banorte_V2_Cargos_Automaticos_Mastercard = new PaymentMethod("575", "Banorte V2 Cargos Automáticos Mastercard");
					public static readonly PaymentMethod Sub1_Discover = new PaymentMethod("576", "Sub1 Discover");
					public static readonly PaymentMethod Banese_Card = new PaymentMethod("577", "Banese Card");
					public static readonly PaymentMethod E_Rede = new PaymentMethod("578", "E-Rede");
					public static readonly PaymentMethod E_Rede_Debito = new PaymentMethod("579", "E-Rede Débito");

					private PaymentMethod(string value, string description)
					{
						this.value = value;
						this.description = description;
					}

					public override string ToString()
					{
						return value;
					}

					public string GetValue()
					{
						return value;
					}

					public string GetDescription()
					{
						return description;
					}

					public static string GetDescription(string codigoPaymentMethod)
					{
						if (codigoPaymentMethod == null) return "";

						Type type = typeof(PaymentMethod);
						foreach (var p in type.GetFields())
						{
							var v = p.GetValue(null);
							if (codigoPaymentMethod.Equals(((PaymentMethod)v).GetValue())) return ((PaymentMethod)v).GetDescription();
						}

						if (codigoPaymentMethod.Trim().Length > 0) return "PaymentMethod desconhecido: " + codigoPaymentMethod;
						return "";
					}
				}
				#endregion

				#region [ Webhook ]
				public static class Webhook
				{
					#region [ EmailEnviadoStatus ]
					public static class EmailEnviadoStatus
					{
						public const byte Inicial = 0;
						public const byte EnviadoComSucesso = 1;
						public const byte ErroERP = 3;
						public const byte TransacaoJaProcessadaAnteriormente = 4;
						public const byte EmpresaInvalida = 6;
						public const byte OperacaoForaEscopo = 7;
						public const byte DadosPreImplantacao = 8;
						public const byte ExcedeuMaxTentativasQueryDadosComplementares = 9;
					}
					#endregion

					#region [ ProcessamentoErpStatus ]
					public static class ProcessamentoErpStatus
					{
						public const int Inicial = 0;
						public const int ProcessadoComSucesso = 1;
						public const int ErroERP = 3;
						public const int TransacaoJaProcessadaAnteriormente = 4;
						public const int EmpresaInvalida = 6;
						public const int OperacaoForaEscopo = 7;
						public const int DadosPreImplantacao = 8;
						public const int ExcedeuMaxTentativasQueryDadosComplementares = 9;
					}
					#endregion

					#region [ BraspagDadosComplementaresQueryStatus ]
					public static class BraspagDadosComplementaresQueryStatus
					{
						public const byte Inicial = 0;
						public const byte ProcessadoComSucesso = 1;
						public const byte ErroERP = 3;
						public const byte TransacaoJaProcessadaAnteriormente = 4;
						public const byte FalhaConsultaBraspag = 5;
						public const byte EmpresaInvalida = 6;
						public const byte ExcedeuMaxTentativasQueryDadosComplementares = 9;
					}
					#endregion
				}
				#endregion

				#region [ WebhookV2 ]
				public static class WebhookV2
				{
					#region [ NotificacaoProcessadoStatus ]
					public static class NotificacaoProcessadoStatus
					{
						public const byte Inicial = 0;
						public const byte PaymentMethodIdentificado = 1;
						public const byte NaoProcessado = 2;
						public const byte Sucesso = 3;
						public const byte Falha = 4;
						public const byte TransacaoJaProcessadaAnteriormente = 5;
						public const byte PagamentoJaRegistrado = 6;
					}
					#endregion

					#region [ EmailEnviadoStatus ]
					public static class EmailEnviadoStatus
					{
						public const byte Inicial = 0;
						public const byte EnviadoComSucesso = 1;
						public const byte ErroERP = 3;
						public const byte TransacaoJaProcessadaAnteriormente = 4;
						public const byte EmpresaInvalida = 6;
						public const byte OperacaoForaEscopo = 7;
						public const byte DadosPreImplantacao = 8;
						public const byte ExcedeuMaxTentativasQueryDadosComplementares = 9;
					}
					#endregion

					#region [ ProcessamentoErpStatus ]
					public static class ProcessamentoErpStatus
					{
						public const int Inicial = 0;
						public const int ProcessadoComSucesso = 1;
						public const int ErroERP = 3;
						public const int TransacaoJaProcessadaAnteriormente = 4;
						public const int EmpresaInvalida = 6;
						public const int OperacaoForaEscopo = 7;
						public const int DadosPreImplantacao = 8;
						public const int ExcedeuMaxTentativasQueryDadosComplementares = 9;
					}
					#endregion

					#region [ BraspagDadosComplementaresQueryStatus ]
					public static class BraspagDadosComplementaresQueryStatus
					{
						public const byte Inicial = 0;
						public const byte ProcessadoComSucesso = 1;
						public const byte ErroERP = 3;
						public const byte TransacaoJaProcessadaAnteriormente = 4;
						public const byte FalhaConsultaBraspag = 5;
						public const byte EmpresaInvalida = 6;
						public const byte ExcedeuMaxTentativasQueryDadosComplementares = 9;
					}
					#endregion
				}
				#endregion
			}
			#endregion

			#region [ Clearsale ]
			public static class Clearsale
			{
				#region [ Endereço Web Service ]
				public static readonly string CS_ENTITY_CODE = GetConfigurationValue("ClearsaleEntityCode");
				public static readonly string WS_CS_ENDERECO_SERVICE = GetConfigurationValue("ClearsaleWsEnderecoService");
				public static readonly string WS_CS_ENDERECO_EXTENDED_SERVICE = GetConfigurationValue("ClearsaleWsEnderecoExtendedService");
				#endregion

				#region [ Parâmetros ]
				// The Timeout applies to the entire request and response, not individually to the GetRequestStream and GetResponse method calls
				public static readonly int REQUEST_TIMEOUT_EM_MS = 3 * 60 * 1000;
				#endregion

				#region [ Transacao ]
				public sealed class Transacao
				{
					// Type safe enum pattern
					private readonly string methodName;
					private readonly string codOpLog;
					private readonly string enderecoWebService;
					private readonly string soapAction;

					public static readonly Transacao SendOrders = new Transacao("SendOrders", "SendOrders", WS_CS_ENDERECO_SERVICE, "http://www.clearsale.com.br/integration/SendOrders");
					public static readonly Transacao GetReturnAnalysis = new Transacao("GetReturnAnalysis", "GetReturnAnalysis", WS_CS_ENDERECO_SERVICE, "http://www.clearsale.com.br/integration/GetReturnAnalysis");
					public static readonly Transacao SetOrderAsReturned = new Transacao("SetOrderAsReturned", "SetOrderAsReturned", WS_CS_ENDERECO_SERVICE, "http://www.clearsale.com.br/integration/SetOrderAsReturned");
					public static readonly Transacao GetAnalystComments = new Transacao("GetAnalystComments", "GetAnalystComments", WS_CS_ENDERECO_SERVICE, "http://www.clearsale.com.br/integration/GetAnalystComments");

					private Transacao(string methodName, string codOpLog, string enderecoWebService, string soapAction)
					{
						this.methodName = methodName;
						this.codOpLog = codOpLog;
						this.enderecoWebService = enderecoWebService;
						this.soapAction = soapAction;
					}

					public string GetMethodName()
					{
						return methodName;
					}

					public string GetEnderecoWebService()
					{
						return enderecoWebService;
					}

					public string GetSoapAction()
					{
						return soapAction;
					}

					public string GetCodOpLog()
					{
						return codOpLog;
					}

					public override string ToString()
					{
						return codOpLog;
					}
				}
				#endregion

				#region [ Email ]
				public static class Email
				{
					public static readonly string REMETENTE_MSG_ALERTA_SISTEMA = GetConfigurationValue("RemetenteMsgAlertaSistema");
					public static readonly string DESTINATARIO_MSG_ALERTA_SISTEMA = GetConfigurationValue("DestinatarioMsgAlertaSistema");
					public static string DESTINATARIO_MSG_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = ""; // Obtém o parâmetro no BD (FinSvc_BP_CS_Braspag_DestinatarioMsgAlertaTransacoesPendentesProxCancelAuto)
					public static string DESTINATARIO_MSG_ALERTA_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = ""; // Obtém o parâmetro no BD (FinSvc_BP_CS_Braspag_DestinatarioMsgAlertaCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto)
				}
				#endregion

				#region [ StatusAF ]
				public sealed class StatusAF
				{
					// Type safe enum pattern
					private readonly String value;
					private readonly String description;

					public static readonly StatusAF APROVACAO_AUTOMATICA = new StatusAF("APA", "Aprovação Automática");
					public static readonly StatusAF APROVACAO_MANUAL = new StatusAF("APM", "Aprovação Manual");
					public static readonly StatusAF APROVACAO_POR_POLITICA = new StatusAF("APP", "Aprovação por Política");
					public static readonly StatusAF REPROVADO_SEM_SUSPEITA = new StatusAF("RPM", "Reprovado Sem Suspeita");
					public static readonly StatusAF ANALISE_MANUAL = new StatusAF("AMA", "Análise Manual");
					public static readonly StatusAF ERRO = new StatusAF("ERR", "Erro");
					public static readonly StatusAF NOVO = new StatusAF("NVO", "Novo");
					public static readonly StatusAF SUSPENSAO_MANUAL = new StatusAF("SUS", "Suspensão Manual");
					public static readonly StatusAF CANCELADO_PELO_CLIENTE = new StatusAF("CAN", "Cancelado pelo Cliente");
					public static readonly StatusAF FRAUDE_CONFIRMADA = new StatusAF("FRD", "Fraude Confirmada");
					public static readonly StatusAF REPROVACAO_AUTOMATICA = new StatusAF("RPA", "Reprovação Automática");
					public static readonly StatusAF REPROVACAO_POR_POLITICA = new StatusAF("RPP", "Reprovação por Política");

					private StatusAF(string value, string description)
					{
						this.value = value;
						this.description = description;
					}

					public string GetValue()
					{
						return value;
					}

					public override string ToString()
					{
						return value;
					}

					public string GetDescription()
					{
						return description;
					}

					public static string GetDescription(string codigo)
					{
						if (codigo == null) return "";

						Type type = typeof(StatusAF);
						foreach (var p in type.GetFields())
						{
							var v = p.GetValue(null);
							if (codigo.Equals(((StatusAF)v).GetValue())) return ((StatusAF)v).GetDescription();
						}

						if (codigo.Trim().Length > 0) return "Código Desconhecido: " + codigo;
						return "";
					}
				}
				#endregion

				#region [ TransactionStatus_StatusCode ]
				public sealed class TransactionStatus_StatusCode
				{
					// Type safe enum pattern
					private readonly String value;
					private readonly String description;

					public static readonly TransactionStatus_StatusCode OK = new TransactionStatus_StatusCode("OK", "OK");

					private TransactionStatus_StatusCode(string value, string description)
					{
						this.value = value;
						this.description = description;
					}

					public override string ToString()
					{
						return value;
					}

					public string GetValue()
					{
						return value;
					}

					public string GetDescription()
					{
						return description;
					}
				}
				#endregion

				#region [ PackageStatus_StatusCode ]
				public sealed class PackageStatus_StatusCode
				{
					// Type safe enum pattern
					private readonly String value;
					private readonly String description;

					public static readonly PackageStatus_StatusCode TRANSACAO_CONCLUIDA = new PackageStatus_StatusCode("00", "Transação Concluída");
					public static readonly PackageStatus_StatusCode USUARIO_INEXISTENTE = new PackageStatus_StatusCode("01", "Usuário Inexistente");
					public static readonly PackageStatus_StatusCode ERRO_VALIDACAO_XML = new PackageStatus_StatusCode("02", "Erro na Validação do XML");
					public static readonly PackageStatus_StatusCode ERRO_TRANSFORMAR_XML = new PackageStatus_StatusCode("03", "Erro ao Transformar XML");
					public static readonly PackageStatus_StatusCode ERRO_INESPERADO = new PackageStatus_StatusCode("04", "Erro Inesperado");
					public static readonly PackageStatus_StatusCode PEDIDO_JA_ENVIADO = new PackageStatus_StatusCode("05", "Pedido Já Enviado ou Não Está em Reanálise");
					public static readonly PackageStatus_StatusCode ERRO_PLUGIN_ENTRADA = new PackageStatus_StatusCode("06", "Erro no Plugin de Entrada");
					public static readonly PackageStatus_StatusCode ERRO_PLUGIN_SAIDA = new PackageStatus_StatusCode("07", "Erro no Plugin de Saída");

					private PackageStatus_StatusCode(string value, string description)
					{
						this.value = value;
						this.description = description;
					}

					public override string ToString()
					{
						return value;
					}

					public string GetValue()
					{
						return value;
					}

					public string GetDescription()
					{
						return description;
					}
				}
				#endregion

				#region [ T_PAGTO_GW_AF_PHONE_IdBlocoXml ]
				public sealed class T_PAGTO_GW_AF_PHONE_IdBlocoXml
				{
					// Type safe enum pattern
					private readonly String value;

					public static readonly T_PAGTO_GW_AF_PHONE_IdBlocoXml Order_BillingData_Phones = new T_PAGTO_GW_AF_PHONE_IdBlocoXml("Order/BillingData/Phones");
					public static readonly T_PAGTO_GW_AF_PHONE_IdBlocoXml Order_ShippingData_Phones = new T_PAGTO_GW_AF_PHONE_IdBlocoXml("Order/ShippingData/Phones");

					private T_PAGTO_GW_AF_PHONE_IdBlocoXml(string value)
					{
						this.value = value;
					}

					public string GetValue()
					{
						return value;
					}

					public override string ToString()
					{
						return value;
					}
				}
				#endregion
			}
			#endregion

			#region [ EmailCtrl ]
			public static class EmailCtrl
			{
				// Type safe enum pattern
				#region [ TipoDestinatario ]
				public sealed class TipoDestinatario
				{
					private readonly String value;

					public static readonly TipoDestinatario CLIENTE = new TipoDestinatario("C");
					public static readonly TipoDestinatario ADMINISTRADOR_SISTEMA = new TipoDestinatario("A");

					private TipoDestinatario(string value)
					{
						this.value = value;
					}

					public string GetValue()
					{
						return value;
					}

					public override string ToString()
					{
						return value;
					}
				}
				#endregion

				#region [ Modulo ]
				public sealed class Modulo
				{
					private readonly String value;

					public static readonly Modulo BRASPAG = new Modulo("BP");
					public static readonly Modulo CLEARSALE = new Modulo("CS");

					private Modulo(string value)
					{
						this.value = value;
					}

					public string GetValue()
					{
						return value;
					}

					public override string ToString()
					{
						return value;
					}
				}
				#endregion

				#region [ TipoMsg ]
				public sealed class TipoMsg
				{
					private readonly String value;

					public static readonly TipoMsg ALERTA = new TipoMsg("A");
					public static readonly TipoMsg INFORMATIVA = new TipoMsg("I");
					public static readonly TipoMsg FALHA = new TipoMsg("F");

					private TipoMsg(string value)
					{
						this.value = value;
					}

					public string GetValue()
					{
						return value;
					}

					public override string ToString()
					{
						return value;
					}
				}
				#endregion

				#region [ CodigoMsg ]
				public sealed class CodigoMsg
				{
					private readonly String value;

					public static readonly CodigoMsg CLEARSALE_EXCEDEU_LIMITE_TENTATIVAS_TX = new CodigoMsg("CS9001");

					private CodigoMsg(string value)
					{
						this.value = value;
					}

					public string GetValue()
					{
						return value;
					}

					public override string ToString()
					{
						return value;
					}
				}
				#endregion
			}
			#endregion

			#region [ Log ]
			public static class LogAtividade
			{
				public static string PathLogAtividade = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\LOG_ATIVIDADE";
				public const int CorteArqLogEmDias = 365;
				public const string ExtensaoArqLog = "LOG";
			}
			#endregion

			#region [ LogBd ]
			public static class LogBd
			{
				#region [ Usuario ]
				public static class Usuario
				{
					public const string ID_USUARIO_SISTEMA = "SISTEMA";
					public const string ID_USUARIO_LOG = "FINSVC";
				}
				#endregion

				#region [ Operacao ]
				public static class Operacao
				{
					public const string OP_LOG_RECONEXAO_BD = "Reconexao-BD";
					public const string OP_LOG_ELIMINA_LOG_ANTIGO = "APAGA LOG ANTIGO";
					public const string OP_LOG_EXECUTA_LIMPEZA_TABELA = "LIMPEZA TABELA";
					public const string OP_LOG_FINANCEIROSERVICE_INICIADO = "FINSVC INICIADO";
					public const string OP_LOG_FINANCEIROSERVICE_ENCERRADO = "FINSVC ENCERRADO";
					public const string OP_LOG_CANCELAMENTO_AUTOMATICO_PEDIDO = "PED CANCEL AUTO";
					public const string OP_LOG_FINANCEIROSERVICE_PROCESSAMENTO_PRODUTOS_VENDIDOS_SEM_PRESENCA_ESTOQUE = "FINSVC PROC PROD SP";
					public const string OP_LOG_PROCESSAMENTO_BP_CS_ANTIFRAUDE_CLEARSALE = "BP_CS_SVC_ANTIFRAUDE";
					public const string OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_CLEARSALE = "PedPagtoContabBpCs";
					public const string OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_ATUALIZACAO_STATUS_TR_PENDENTES = "BP_CS_PrcBpUpdTrPend";
					public const string OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_ENVIAR_EMAIL_ALERTA_TRANSACOES_PENDENTES_PROX_CANCEL_AUTO = "BP_CS_MsgPrxCancAuto";
					public const string OP_LOG_PROCESSAMENTO_BP_CS_BRASPAG_CAPTURA_TRANSACAO_PENDENTE_DEVIDO_PRAZO_FINAL_CANCEL_AUTO = "BP_CS_CapTrxCancAuto";
					public const string OP_LOG_PROCESSAMENTO_ENVIAR_EMAIL_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = "MsgPedNovoAnCred";
					public const string OP_LOG_PROCESSAMENTO_WEBHOOK_BRASPAG = "WebhookBraspag";
					public const string OP_LOG_PROCESSAMENTO_WEBHOOK_BRASPAG_V2 = "WebhookBraspagV2";
					public const string OP_LOG_PROCESSAMENTO_ESTORNOS_PENDENTES = "FINSVC_ESTORNOS_PEND";
					public const string OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_WEBHOOK = "PedPagtoContabBpWH";
					public const string OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_WEBHOOK_V2 = "PedPagtoContabBpWHV2";
					public const string OP_LOG_FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_ECOMMERCE = "FCInsBolEC";
					public const string OP_LOG_LIMPEZA_SESSION_TOKEN = "LimpezaSessionToken";
					public const string OP_LOG_MANUTENCAO_ARQUIVOS_UPLOAD_FILE = "ManutArqUploadFile";

					public const string OP_LOG_ESTOQUE_ENTRADA = "ESTOQUE INCLUSÃO";
					public const string OP_LOG_ESTOQUE_REMOVE = "ESTOQUE EXCLUSÃO";
					public const string OP_LOG_ESTOQUE_ALTERACAO = "ESTOQUE EDIÇÃO";
					public const string OP_LOG_ESTOQUE_CONVERSAO_KIT = "ESTOQUE CONVERTE KIT";
					public const string OP_LOG_ESTOQUE_PROCESSA_SP = "ESTOQUE PROCESSA SP";
					public const string OP_LOG_ESTOQUE_TRANSFERENCIA = "ESTOQUE TRANSFERE";
					public const string OP_LOG_ESTOQUE_TRANSF_PEDIDO = "ESTOQUE TRANSF PED";
				}
				#endregion
			}
			#endregion

			#region [ ManutencaoLogBd ]
			public static class ManutencaoLogBd
			{
				#region [ Corte ]
				public static class Corte
				{
					public const int T_FIN_LOG__CORTE_EM_DIAS = 12 * 31;
					public const int T_LOG__CORTE_EM_DIAS = 12 * 31;
					public const int T_SESSAO_ABANDONADA__CORTE_EM_DIAS = 6 * 31;
					public const int T_SESSAO_HISTORICO__CORTE_EM_DIAS = 6 * 31;
					public const int T_SESSAO_RESTAURADA__CORTE_EM_DIAS = 6 * 31;
					public const int T_ESTOQUE_LOG__CORTE_EM_DIAS = 5 * 366;
					public const int T_ESTOQUE_SALDO_DIARIO__CORTE_EM_DIAS = T_ESTOQUE_LOG__CORTE_EM_DIAS + 1;
					public const int T_ESTOQUE_VENDA_SALDO_DIARIO__CORTE_EM_DIAS = 5 * 366;
					public const int T_FINSVC_LOG__CORTE_EM_DIAS = 12 * 31;
					public const int T_EMAILSNDSVC_LOG__CORTE_EM_DIAS = 12 * 31;
					public const int T_EMAILSNDSVC_LOG_ERRO__CORTE_EM_DIAS = 12 * 31;
					public const int T_MAGENTO_API_PEDIDO_XML__INFO_UTILIZADA__CORTE_EM_DIAS = 36 * 31;
					public const int T_MAGENTO_API_PEDIDO_XML__INFO_DESCARTADA__CORTE_EM_DIAS = 1 * 31;
                    public const int T_CTRL_RELATORIO_USUARIO_X_PEDIDO__CORTE_EM_DIAS = 1 * 31;
                }
				#endregion
			}
			#endregion

			#region[ Data/Hora ]
			public static class DataHora
			{
				public const string FmtDia = "dd";
				public const string FmtDiaAbreviado = "ddd";
				public const string FmtDiaExtenso = "dddd";
				public const string FmtMes = "MM";
				public const string FmtMesAbreviado = "MMM";
				public const string FmtMesExtenso = "MMMM";
				public const string FmtAno = "yyyy";
				public const string FmtAnoCom2Digitos = "yy";
				public const string FmtHora = "HH";
				public const string FmtHora12 = "hh";
				public const string FmtMin = "mm";
				public const string FmtSeg = "ss";
				public const string FmtMiliSeg = "fff";
				public const string FmtAmPm = "tt";
				public const string FmtYYYYMMDD = FmtAno + FmtMes + FmtDia;
				public const string FmtHHMMSS = FmtHora + FmtMin + FmtSeg;
				public const string FmtHhMmComSeparador = FmtHora + ":" + FmtMin;
				public const string FmtHhMmSsComSeparador = FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtDdMmYyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAnoCom2Digitos;
				public const string FmtDdMmYyyyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno;
				public const string FmtDdMmYyyyHhMmComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin;
				public const string FmtDdMmYyyyHhMmSsComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
				public const string FmtYyyyMmDdComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia;
				public const string FmtYyyyMmDdHhMmSsComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
			}
			#endregion

			#region [ BlocoNotasPedidoNivelAcesso ]
			public class BlocoNotasPedidoNivelAcesso
			{
				// NÍVEL DE ACESSO DO BLOCO DE NOTAS DO PEDIDO
				public const int COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO = 0;
				public const int COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__ILIMITADO = -1;
				public const int COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO = 10;
				public const int COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO = 20;
				public const int COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__SIGILOSO = 30;
			}
			#endregion

			#region [ Etc ]
			public class Etc
			{
				public const int TAM_MIN_NUM_PEDIDO = 6;    // SOMENTE PARTE NUMÉRICA DO NÚMERO DO PEDIDO
				public const int TAM_MIN_ID_PEDIDO = 7; // PARTE NUMÉRICA DO NÚMERO DO PEDIDO + LETRA REFERENTE AO ANO
				public const char COD_SEPARADOR_FILHOTE = '-';
				public const int COD_NEGATIVO_UM = -1;
				public const int TAM_MAX_NSU = 12;
				public const decimal MAX_VALOR_MARGEM_ERRO_PAGAMENTO = 0.5m;
				public const string SIMBOLO_MONETARIO = "R$";
			}
			#endregion

			#region [ Nsu ]
			public class Nsu
			{
				public const string NSU_ID_ESTOQUE_MOVTO = "ESTOQUE_MOVTO";
				public const string NSU_PEDIDO_PAGAMENTO = "PEDIDO_PAGAMENTO";
				public const string T_PEDIDO_BLOCO_NOTAS = "T_PEDIDO_BLOCO_NOTAS";
				public const string T_FIN_PEDIDO_HIST_PAGTO = "T_FIN_PEDIDO_HIST_PAGTO";
			}
			#endregion

			#region [ PrazoCancelAutoPedidoEmDias ]
			// Type safe enum pattern
			public sealed class PrazoCancelAutoPedidoEmDias
			{
				private readonly String name;
				private readonly int value;

				public static readonly PrazoCancelAutoPedidoEmDias PENDENTE_CARTAO_CREDITO = new PrazoCancelAutoPedidoEmDias("PENDENTE_CARTAO_CREDITO", 7);
				public static readonly PrazoCancelAutoPedidoEmDias CREDITO_OK_AGUARDANDO_DEPOSITO = new PrazoCancelAutoPedidoEmDias("CREDITO_OK_AGUARDANDO_DEPOSITO", 7);
				public static readonly PrazoCancelAutoPedidoEmDias PENDENTE_VENDAS = new PrazoCancelAutoPedidoEmDias("PENDENTE_VENDAS", 10);

				private PrazoCancelAutoPedidoEmDias(string name, int value)
				{
					this.name = name;
					this.value = value;
				}

				public int GetValue()
				{
					return value;
				}

				public override string ToString()
				{
					return value.ToString();
				}

				public string GetName()
				{
					return name;
				}
			}
			#endregion

			#region [ PedidoCanceladoCodigoMotivo ]
			// Type safe enum pattern
			public sealed class PedidoCanceladoCodigoMotivo
			{
				private readonly String value;

				public static readonly PedidoCanceladoCodigoMotivo CANCELAMENTO_AUTOMATICO = new PedidoCanceladoCodigoMotivo("001");

				private PedidoCanceladoCodigoMotivo(string value)
				{
					this.value = value;
				}

				public string GetValue()
				{
					return value;
				}

				public override string ToString()
				{
					return value;
				}
			}
			#endregion

			#region [ Tipos de Estoque ]
			public class TipoEstoque
			{
				public const String ID_ESTOQUE_VENDA = "VDA";
				public const String ID_ESTOQUE_VENDIDO = "VDO";
				public const String ID_ESTOQUE_SEM_PRESENCA = "SPE";
				public const String ID_ESTOQUE_KIT = "KIT";
				public const String ID_ESTOQUE_SHOW_ROOM = "SHR";
				public const String ID_ESTOQUE_DANIFICADOS = "DAN";
				public const String ID_ESTOQUE_DEVOLUCAO = "DEV";
				public const String ID_ESTOQUE_ROUBO = "ROU";
				public const String ID_ESTOQUE_ENTREGUE = "ETG";
			}
			#endregion

			#region [ Operação de Movimentação do Estoque ]
			public class OperacaoMovimentoEstoque
			{
				// OPERAÇÕES (MOVIMENTOS) DO ESTOQUE
				public const String OP_ESTOQUE_ENTRADA = "CAD";
				public const String OP_ESTOQUE_VENDA = "VDA";
				public const String OP_ESTOQUE_CONVERSAO_KIT = "KIT";
				public const String OP_ESTOQUE_TRANSFERENCIA = "TRF";
				public const String OP_ESTOQUE_ENTREGA = "ETG";
				public const String OP_ESTOQUE_DEVOLUCAO = "DEV";
			}
			#endregion

			#region [ OperacaoLogEstoque ]
			public class OperacaoLogEstoque
			{
				// OPERAÇÕES NO LOG DE MOVIMENTAÇÃO DO ESTOQUE (T_ESTOQUE_LOG)
				public const String OP_ESTOQUE_LOG_ENTRADA = "CAD";
				public const String OP_ESTOQUE_LOG_ENTRADA_VIA_KIT = "CKT";
				public const String OP_ESTOQUE_LOG_VENDA = "VDA";
				public const String OP_ESTOQUE_LOG_CONVERSAO_KIT = "KIT";
				public const String OP_ESTOQUE_LOG_TRANSFERENCIA = "TRF";
				public const String OP_ESTOQUE_LOG_ENTREGA = "ETG";
				public const String OP_ESTOQUE_LOG_DEVOLUCAO = "DEV";
				public const String OP_ESTOQUE_LOG_ESTORNO = "EST";
				public const String OP_ESTOQUE_LOG_CANCELA_SEM_PRESENCA = "XSP";
				public const String OP_ESTOQUE_LOG_SPLIT = "SPL";
				public const String OP_ESTOQUE_LOG_VENDA_SEM_PRESENCA = "VSP";
				public const String OP_ESTOQUE_LOG_REMOVE_ENTRADA_ESTOQUE = "XEE";
				public const String OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_NOVO_ITEM = "EEN";
				public const String OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_INCREMENTA = "EEI";
				public const String OP_ESTOQUE_LOG_ENTRADA_ESTOQUE_ALTERA_DECREMENTA = "EED";
				public const String OP_ESTOQUE_LOG_PRODUTO_VENDIDO_SEM_PRESENCA_SAIDA = "SPS";
				public const String OP_ESTOQUE_LOG_TRANSF_PRODUTO_VENDIDO_ENTRE_PEDIDOS = "TVP";
			}
			#endregion

			#region [ Códigos para formas de pagamento do pedido ]
			public class TipoParcelamentoPedido
			{
				public const short COD_FORMA_PAGTO_A_VISTA = 1;
				public const short COD_FORMA_PAGTO_PARCELADO_CARTAO = 2;
				public const short COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA = 3;
				public const short COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA = 4;
				public const short COD_FORMA_PAGTO_PARCELA_UNICA = 5;
				public const short COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA = 6;
			}
			#endregion

			#region [ Códigos para opções de forma de pagamento do pedido ]
			public class CodFormaPagtoPedido
			{
				public const short ID_FORMA_PAGTO_DINHEIRO = 1;
				public const short ID_FORMA_PAGTO_DEPOSITO = 2;
				public const short ID_FORMA_PAGTO_CHEQUE = 3;
				public const short ID_FORMA_PAGTO_BOLETO = 4;
				public const short ID_FORMA_PAGTO_CARTAO = 5;
				public const short ID_FORMA_PAGTO_BOLETO_AV = 6;
				public const short ID_FORMA_PAGTO_CARTAO_MAQUINETA = 7;
			}
			#endregion

			#region [ Status de Pagamento do Pedido ]
			public class StPagtoPedido
			{
				public const String ST_PAGTO_PAGO = "S";
				public const String ST_PAGTO_NAO_PAGO = "N";
				public const String ST_PAGTO_PARCIAL = "P";
			}
			#endregion

			#region [ Tipo de pagamento (t_PEDIDO_PAGAMENTO) ]
			public class PedidoPagtoTipoOperacao
			{
				public const string QUITACAO = "Q";
				public const string PARCIAL = "P";
				public const string VISANET = "V";
				public const string CIELO = "C";
				public const string BRASPAG = "B";
				public const string GW_BRASPAG_CLEARSALE = "G";
				public const string BRASPAG_WEBHOOK = "W";
				public const string BRASPAG_WEBHOOK_V2 = "W";
			}
			#endregion

			#region [ Status de Entrega do Pedido ]
			public class StEntregaPedido
			{
				public const String ST_ENTREGA_ESPERAR = "ESP";
				public const String ST_ENTREGA_SPLIT_POSSIVEL = "SPL";
				public const String ST_ENTREGA_SEPARAR = "SEP";
				public const String ST_ENTREGA_A_ENTREGAR = "AET";
				public const String ST_ENTREGA_ENTREGUE = "ETG";
				public const String ST_ENTREGA_CANCELADO = "CAN";
			}
			#endregion

			#region [ Status de Pedido Recebido ]
			public class StPedidoRecebido
			{
				public const short COD_ST_PEDIDO_RECEBIDO_NAO = 0;
				public const short COD_ST_PEDIDO_RECEBIDO_SIM = 1;
				public const short COD_ST_PEDIDO_RECEBIDO_NAO_DEFINIDO = 10;
			}
			#endregion

			#region [ Código de Status de Cliente Contribuinte de ICMS ]
			public class StClienteContribuinteIcmsStatus
			{
				public const byte CONTRIBUINTE_ICMS_INICIAL = 0;
				public const byte CONTRIBUINTE_ICMS_NAO = 1;
				public const byte CONTRIBUINTE_ICMS_SIM = 2;
				public const byte CONTRIBUINTE_ICMS_ISENTO = 3;
			}
			#endregion

			#region [ Código de Status de Cliente Proudtor Rural ]
			public class StClienteProdutorRural
			{
				public const byte PRODUTOR_RURAL_INICIAL = 0;
				public const byte PRODUTOR_RURAL_NAO = 1;
				public const byte PRODUTOR_RURAL_SIM = 2;
			}
			#endregion

			#region [ T_PEDIDO__ANALISE_CREDITO_STATUS ]
			public class T_PEDIDO__ANALISE_CREDITO_STATUS
			{
				public const int ST_INICIAL = 0;
				public const int CREDITO_PENDENTE = 1;
				public const int CREDITO_OK = 2;
				public const int PENDENTE_ENDERECO = 6;
				public const int CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO = 7;
				public const int PENDENTE_VENDAS = 8;
				public const int CREDITO_OK_AGUARDANDO_DEPOSITO = 9;
				public const int NAO_ANALISADO = 10; // PEDIDOS ANTIGOS QUE JÁ ESTAVAM NA BASE
				public const int CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV = 11;
				public const int PENDENTE_PAGTO_ANTECIPADO_BOLETO = 12;
			}
			#endregion

			#region [ T_PEDIDO__ENTREGA_IMEDIATA_STATUS ]
			public class T_PEDIDO__ENTREGA_IMEDIATA_STATUS
			{
				public const int ETG_IMEDIATA_ST_INICIAL = 0;
				public const int ETG_IMEDIATA_NAO = 1;
				public const int ETG_IMEDIATA_SIM = 2;
				public const int ETG_IMEDIATA_NAO_DEFINIDO = 10;
			}
			#endregion

			#region [ T_FIN_PEDIDO_HIST_PAGTO__status ]
			public class T_FIN_PEDIDO_HIST_PAGTO__status
			{
				public const int PREVISAO = 1;
				public const int QUITADO = 2;
				public const int CANCELADO = 3;
			}
			#endregion
		}
		#endregion

		#region [ Parametros ]
		public static class Parametros
		{
			#region [ Geral ]
			public static class Geral
			{
				public static string DESTINATARIO_PADRAO_MSG_ALERTA_SISTEMA = "adm_finsvc@bonshop.com.br";
				public static bool ExecutarCancelamentoAutomaticoPedidos = false;
				public static bool ProcessamentoProdutosVendidosSemPresencaEstoque_FlagHabilitacao = false;
				public static TimeSpan HorarioCancelamentoAutomaticoPedidos = new TimeSpan(1, 20, 0); // Hours, Minutes, Seconds
				public static TimeSpan HorarioManutencaoArqLogAtividade = new TimeSpan(1, 20, 0); // Hours, Minutes, Seconds
				public static TimeSpan HorarioManutencaoBdLogAntigo = new TimeSpan(1, 20, 0); // Hours, Minutes, Seconds
				public static bool FinSvc_BP_CS_Processamento_PeriodoInatividade_FlagHabilitacao = true;
				public static TimeSpan FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioInicio = new TimeSpan(21, 0, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_BP_CS_Processamento_PeriodoInatividade_HorarioTermino = new TimeSpan(7, 0, 0); // Hours, Minutes, Seconds
				public static bool FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_FlagHabilitacao = true;
				public static TimeSpan FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioInicio = new TimeSpan(21, 0, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_ProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito_PeriodoInatividade_HorarioTermino = new TimeSpan(7, 0, 0); // Hours, Minutes, Seconds
				public static bool ExecutarProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito = false;
				public static int TempoEntreProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCreditoEmSeg = 5 * 60;
				public static string DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO = ""; // Obtém o parâmetro no BD (FinSvc_BP_CS_Braspag_DestinatarioMsgAlertaTransacoesPendentesProxCancelAuto)
				public static bool FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_FlagHabilitacao = true;
				public static TimeSpan FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioInicio = new TimeSpan(20, 0, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_ProcessamentoWebhookBraspag_PeriodoInatividade_HorarioTermino = new TimeSpan(7, 0, 0); // Hours, Minutes, Seconds
				public static bool ExecutarProcessamentoWebhookBraspag = false;
				public static int TempoEntreProcessamentoWebhookBraspagEmSeg = 10 * 60;
				public static string DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG = ""; // Obtém o parâmetro no BD (FinSvc_DestinatarioMsgAlertaWebhookBraspag)
				public static int FinSvc_ProcessamentoWebhookBraspag_MaxTentativasQueryDadosComplementares = 10;
				public static bool FinSvc_ProcessamentoWebhookBraspagV2_PeriodoInatividade_FlagHabilitacao = true;
				public static TimeSpan FinSvc_ProcessamentoWebhookBraspagV2_PeriodoInatividade_HorarioInicio = new TimeSpan(20, 0, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_ProcessamentoWebhookBraspagV2_PeriodoInatividade_HorarioTermino = new TimeSpan(7, 0, 0); // Hours, Minutes, Seconds
				public static bool ExecutarProcessamentoWebhookBraspagV2 = false;
				public static int TempoEntreProcessamentoWebhookBraspagV2EmSeg = 10 * 60;
				public static string DESTINATARIO_MSG_ALERTA_WEBHOOK_BRASPAG_V2 = ""; // Obtém o parâmetro no BD (FinSvc_DestinatarioMsgAlertaWebhookBraspagV2)
				public static int FinSvc_ProcessamentoWebhookBraspagV2_MaxTentativasQueryDadosComplementares = 10;
				public static List<int> CancelamentoAutomaticoPedidosLojasIgnoradas;
				public static bool FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_FlagHabilitacao = true;
				public static TimeSpan FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioInicio = new TimeSpan(21, 0, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_BP_CS_Processamento_EstornosPendentes_PeriodoInatividade_HorarioTermino = new TimeSpan(8, 0, 0); // Hours, Minutes, Seconds
				public static string DESTINATARIO_MSG_ALERTA_ESTORNOS_PENDENTES_ABORTADOS = ""; // Obtém o parâmetro no BD (FinSvc_DestinatarioMsgAlertaEstornosPendentesAbortados)
				public static int ESTORNOS_PENDENTES_PRAZO_MAXIMO_VERIFICACAO_EM_DIAS = 2;
				public static bool UploadFile_ManutencaoArquivos_FlagHabilitacao = false;
				public static TimeSpan UploadFile_ManutencaoArquivos_Horario = new TimeSpan(1, 20, 0); // Hours, Minutes, Seconds
				public static bool SessionToken_Limpeza_FlagHabilitacao = false;
				public static TimeSpan SessionToken_Limpeza_Horario = new TimeSpan(1, 20, 0); // Hours, Minutes, Seconds
				public static int ConsultaExecucaoSolicitada_ProcProdutosVendidosSemPresencaEstoque_TempoEntreProcEmSeg = 60;
				public static bool TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = false;
			}
			#endregion

			#region [ Braspag ]
			public static class Braspag
			{
				public static bool ExecutarProcessamentoBpCsBraspagAtualizaStatusTransacoesPendentes = false;
				public static bool ExecutarProcessamentoBpCsEstornosPendentes = false;
				public static bool ExecutarProcessamentoBpCsBraspagEnviarEmailAlertaTransacoesPendentesProxCancelAuto = false;
				public static bool ExecutarProcessamentoBpCsBraspagCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto = false;
				public static TimeSpan FinSvc_BP_CS_BraspagProcessamentoAtualizaStatusTransacoesPendentes_Horario = new TimeSpan(6, 10, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_BP_CS_Braspag_ProcEnviarEmailAlertaTransacoesPendentesProxCancelAuto_Horario = new TimeSpan(6, 10, 0); // Hours, Minutes, Seconds
				public static TimeSpan FinSvc_BP_CS_Braspag_ProcCapturaTransacaoPendenteDevidoPrazoFinalCancelAuto_Horario = new TimeSpan(21, 30, 0); // Hours, Minutes, Seconds
				public static int TempoEntreProcessamentoEstornosPendentesEmSeg = 60 * 60;

				#region [ Webhook ]

				#region [ WebhookBraspagMerchantId ]
				public class WebhookBraspagMerchantId
				{
					public string Empresa { get; set; }
					public string MerchantId { get; set; }
				}
				#endregion

				#region [ WebhookBraspagPlanoContasBoletoEC ]
				public class WebhookBraspagPlanoContasBoletoEC
				{
					public string Empresa { get; set; }
					public byte id_conta_corrente { get; set; }
					public byte id_plano_contas_empresa { get; set; }
					public short id_plano_contas_grupo { get; set; }
					public int id_plano_contas_conta { get; set; }
				}
				#endregion

				public static List<WebhookBraspagMerchantId> webhookBraspagMerchantIdList;
				public static List<WebhookBraspagPlanoContasBoletoEC> webhookBraspagPlanoContasBoletoECList;
				#endregion

				#region [ Webhook V2 ]

				#region [ WebhookBraspagV2MerchantId ]
				public class WebhookBraspagV2MerchantId
				{
					public string Empresa { get; set; }
					public string MerchantId { get; set; }
				}
				#endregion

				#region [ WebhookBraspagV2PlanoContasBoletoEC ]
				public class WebhookBraspagV2PlanoContasBoletoEC
				{
					public string Empresa { get; set; }
					public byte id_conta_corrente { get; set; }
					public byte id_plano_contas_empresa { get; set; }
					public short id_plano_contas_grupo { get; set; }
					public int id_plano_contas_conta { get; set; }
				}
				#endregion

				public static List<WebhookBraspagV2MerchantId> webhookBraspagV2MerchantIdList;
				public static List<WebhookBraspagV2PlanoContasBoletoEC> webhookBraspagV2PlanoContasBoletoECList;
				#endregion
			}
			#endregion

			#region [ Clearsale ]
			public static class Clearsale
			{
				public static int MaxTentativasEnvioTransacao = 10;
				public static int TempoMinEntreTentativasEmSeg = 60 * 60;
				public static int TempoEntreProcessamentoEmSeg = 5 * 60;
				public static int TempoMaxClienteTotalizarPagtoEmSeg = 2 * 60 * 60;
				public static int MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis = 10;
				public static bool ExecutarProcessamentoBpCsAntifraudeClearsale = false;
			}
			#endregion
		}
		#endregion

		#region [ enum: eOpcaoFiltroStAtivo ]
		public enum eOpcaoFiltroStAtivo : byte
		{
			SELECIONAR_TODOS = 0,
			SELECIONAR_SOMENTE_ATIVOS = 1,
			SELECIONAR_SOMENTE_INATIVOS = 2
		}
		#endregion

		#region[ ReaderWriterLock ]
		public static ReaderWriterLock rwlArqLogAtividade = new ReaderWriterLock();
		#endregion

		#region [ Construtor Estático ]
		static Global()
		{
			#region [ Declarações ]
			int index;
			int loja;
			string strValue;
			string strValueAux;
			string strLoja;
			string[] vEmpresa;
			string[] vParametros;
			string[] vParametrosAux;
			string[] vLojasIgnoradas;
			Parametros.Braspag.WebhookBraspagMerchantId webhookBraspagMerchantId;
			Parametros.Braspag.WebhookBraspagPlanoContasBoletoEC webhookBraspagPlanoContasBoletoEC;
			Parametros.Braspag.WebhookBraspagV2MerchantId webhookBraspagV2MerchantId;
			Parametros.Braspag.WebhookBraspagV2PlanoContasBoletoEC webhookBraspagV2PlanoContasBoletoEC;
			#endregion

			#region [ Carrega dados do Braspag MerchantId para serem usados no processamento dos dados recebidos pelo Webhook ]
			Parametros.Braspag.webhookBraspagMerchantIdList = new List<Parametros.Braspag.WebhookBraspagMerchantId>();
			strValue = GetConfigurationValue("WebhookBraspagMerchantIdList");
			if ((strValue ?? "").Length > 0)
			{
				vEmpresa = strValue.Split('|');
				for (int i = 0; i < vEmpresa.Length; i++)
				{
					vParametros = vEmpresa[i].Split('=');
					if (vParametros.Length >= 2)
					{
						webhookBraspagMerchantId = new Parametros.Braspag.WebhookBraspagMerchantId();
						webhookBraspagMerchantId.Empresa = vParametros[0];
						webhookBraspagMerchantId.MerchantId = vParametros[1];
						Parametros.Braspag.webhookBraspagMerchantIdList.Add(webhookBraspagMerchantId);
					}
				}
			}
			#endregion

			#region [ Carrega dados do plano de contas para gravação de lançamentos no fluxo de caixa dos boletos de e-commerce ]
			Parametros.Braspag.webhookBraspagPlanoContasBoletoECList = new List<Parametros.Braspag.WebhookBraspagPlanoContasBoletoEC>();
			strValue = GetConfigurationValue("WebhookBraspagPlanoContasBoletoEC");
			if ((strValue ?? "").Length > 0)
			{
				vEmpresa = strValue.Split('|');
				for (int i = 0; i < vEmpresa.Length; i++)
				{
					// Formato: IdentificadorEmpresa=id_conta_corrente;id_plano_contas_empresa;id_plano_contas_conta
					// Obs: id_plano_contas_grupo deve ser obtido a partir do cadastro da conta (plano de contas) no banco de dados,
					// mas isso não pode ser feito neste momento porque a conexão com o BD ainda não está estabelecida e nem as
					// classes e objetos necessários ainda não estão devidamente inicializados.
					vParametros = vEmpresa[i].Split('=');
					if (vParametros.Length >= 2)
					{
						webhookBraspagPlanoContasBoletoEC = new Parametros.Braspag.WebhookBraspagPlanoContasBoletoEC();
						webhookBraspagPlanoContasBoletoEC.Empresa = vParametros[0];
						strValueAux = vParametros[1];
						vParametrosAux = strValueAux.Split(';');
						if (vParametrosAux.Length >= 3)
						{
							index = 0;
							webhookBraspagPlanoContasBoletoEC.id_conta_corrente = (byte)converteInteiro(vParametrosAux[index]);
							index++;
							webhookBraspagPlanoContasBoletoEC.id_plano_contas_empresa = (byte)converteInteiro(vParametrosAux[index]);
							index++;
							webhookBraspagPlanoContasBoletoEC.id_plano_contas_conta = (int)converteInteiro(vParametrosAux[index]);

							Parametros.Braspag.webhookBraspagPlanoContasBoletoECList.Add(webhookBraspagPlanoContasBoletoEC);
						}
					}
				}
			}
			#endregion

			#region [ Carrega dados do Braspag MerchantId para serem usados no processamento dos dados recebidos pelo Webhook V2 ]
			Parametros.Braspag.webhookBraspagV2MerchantIdList = new List<Parametros.Braspag.WebhookBraspagV2MerchantId>();
			strValue = GetConfigurationValue("WebhookBraspagV2MerchantIdList");
			if ((strValue ?? "").Length > 0)
			{
				vEmpresa = strValue.Split('|');
				for (int i = 0; i < vEmpresa.Length; i++)
				{
					vParametros = vEmpresa[i].Split('=');
					if (vParametros.Length >= 2)
					{
						webhookBraspagV2MerchantId = new Parametros.Braspag.WebhookBraspagV2MerchantId();
						webhookBraspagV2MerchantId.Empresa = vParametros[0];
						webhookBraspagV2MerchantId.MerchantId = vParametros[1];
						Parametros.Braspag.webhookBraspagV2MerchantIdList.Add(webhookBraspagV2MerchantId);
					}
				}
			}
			#endregion

			#region [ Carrega dados do plano de contas para gravação de lançamentos no fluxo de caixa dos boletos de e-commerce (Webhook V2) ]
			Parametros.Braspag.webhookBraspagV2PlanoContasBoletoECList = new List<Parametros.Braspag.WebhookBraspagV2PlanoContasBoletoEC>();
			strValue = GetConfigurationValue("WebhookBraspagV2PlanoContasBoletoEC");
			if ((strValue ?? "").Length > 0)
			{
				vEmpresa = strValue.Split('|');
				for (int i = 0; i < vEmpresa.Length; i++)
				{
					// Formato: IdentificadorEmpresa=id_conta_corrente;id_plano_contas_empresa;id_plano_contas_conta
					// Obs: id_plano_contas_grupo deve ser obtido a partir do cadastro da conta (plano de contas) no banco de dados,
					// mas isso não pode ser feito neste momento porque a conexão com o BD ainda não está estabelecida e nem as
					// classes e objetos necessários ainda não estão devidamente inicializados.
					vParametros = vEmpresa[i].Split('=');
					if (vParametros.Length >= 2)
					{
						webhookBraspagV2PlanoContasBoletoEC = new Parametros.Braspag.WebhookBraspagV2PlanoContasBoletoEC();
						webhookBraspagV2PlanoContasBoletoEC.Empresa = vParametros[0];
						strValueAux = vParametros[1];
						vParametrosAux = strValueAux.Split(';');
						if (vParametrosAux.Length >= 3)
						{
							index = 0;
							webhookBraspagV2PlanoContasBoletoEC.id_conta_corrente = (byte)converteInteiro(vParametrosAux[index]);
							index++;
							webhookBraspagV2PlanoContasBoletoEC.id_plano_contas_empresa = (byte)converteInteiro(vParametrosAux[index]);
							index++;
							webhookBraspagV2PlanoContasBoletoEC.id_plano_contas_conta = (int)converteInteiro(vParametrosAux[index]);

							Parametros.Braspag.webhookBraspagV2PlanoContasBoletoECList.Add(webhookBraspagV2PlanoContasBoletoEC);
						}
					}
				}
			}
			#endregion

			#region [ Carrega a lista de lojas que devem ser ignoradas no cancelamento automático de pedidos ]
			Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas = new List<int>();
			strValue = GetConfigurationValue("CancelamentoAutomaticoPedidosLojasIgnoradas");
			if ((strValue ?? "").Length > 0)
			{
				strValue = strValue.Replace(';', ',');
				strValue = strValue.Replace('|', ',');
				vLojasIgnoradas = strValue.Split(',');
				for (int i = 0; i < vLojasIgnoradas.Length; i++)
				{
					strLoja = vLojasIgnoradas[i] ?? "";
					if (strLoja.Length > 0)
					{
						loja = (int)converteInteiro(strLoja);
						if (loja > 0) Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas.Add(loja);
					}
				}
			}
			#endregion

			#region [ Configuração do parâmetro que define o tratamento para evitar acesso concorrente nas operações com o BD ]
			strValue = GetConfigurationValue("TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO");
			if ((strValue ?? "").Length > 0)
			{
				if (strValue.ToUpper().Equals("TRUE")
					|| strValue.ToUpper().Equals("1")
					|| strValue.ToUpper().Equals("YES")
					|| strValue.ToUpper().Equals("Y")
					|| strValue.ToUpper().Equals("SIM")
					|| strValue.ToUpper().Equals("S"))
				{
					Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = true;
				}
				else
				{
					Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO = false;
				}
			}
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ arredondaParaMonetario ]
		public static decimal arredondaParaMonetario(decimal numero)
		{
			return converteNumeroDecimal(formataMoeda(numero));
		}
		#endregion

		#region[ barraInvertidaAdd ]
		public static string barraInvertidaAdd(string path)
		{
			if (path == null) return "";
			string strResp = path.TrimEnd();
			if (strResp.Length == 0) return "";
			if (strResp[strResp.Length - 1] == (char)92) return strResp;
			return strResp + (char)92;
		}
		#endregion

		#region [ calculaTimeSpanDias ]
		/// <summary>
		/// Calcula a quantidade de dias.
		/// Exemplo de uso:
		///		calculaDateTimeDias(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade de dias.
		/// </returns>
		public static int calculaTimeSpanDias(TimeSpan ts)
		{
			return ts.Days;
		}
		#endregion

		#region [ calculaTimeSpanHoras ]
		/// <summary>
		/// Calcula a quantidade de horas.
		/// Exemplo de uso:
		///		calculaDateTimeHoras(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade de horas.
		/// </returns>
		public static int calculaTimeSpanHoras(TimeSpan ts)
		{
			return ts.Hours + (24 * ts.Days);
		}
		#endregion

		#region [ calculaTimeSpanMiliSegundos ]
		/// <summary>
		/// Calcula a quantidade de milisegundos.
		/// Exemplo de uso:
		///		calculaDateTimeMiliSegundos(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade milisegundos.
		/// </returns>
		public static long calculaTimeSpanMiliSegundos(TimeSpan ts)
		{
			return (long)ts.Milliseconds + 1000L * ((long)ts.Seconds + (60L * ((long)ts.Minutes + (60L * ((long)ts.Hours + (24L * (long)ts.Days))))));
		}
		#endregion

		#region [ calculaTimeSpanMinutos ]
		/// <summary>
		/// Calcula a quantidade de minutos.
		/// Exemplo de uso:
		///		calculaDateTimeMinutos(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade minutos.
		/// </returns>
		public static int calculaTimeSpanMinutos(TimeSpan ts)
		{
			return ts.Minutes + (60 * (ts.Hours + (24 * ts.Days)));
		}
		#endregion

		#region [ calculaTimeSpanSegundos ]
		/// <summary>
		/// Calcula a quantidade de segundos.
		/// Exemplo de uso:
		///		calculaDateTimeSegundos(dtDataFinal - dtDataInicial);
		/// </summary>
		/// <param name="ts">
		/// O parâmetro do tipo TimeSpan pode ser passado através de:
		///		1) Uma variável declarada como TimeSpan
		///		2) Através do resultado da operação "dtDataFinal - dtDataInicial", já que o parâmetro de
		///		   retorno das operações de adição/subtração entre dois operandos do tipo DateTime é um tipo TimeSpan
		/// </param>
		/// <returns>
		/// Retorna a quantidade segundos.
		/// </returns>
		public static long calculaTimeSpanSegundos(TimeSpan ts)
		{
			return (long)ts.Seconds + (60L * ((long)ts.Minutes + (60L * ((long)ts.Hours + (24L * (long)ts.Days)))));
		}
		#endregion

		#region [ converteDateTimeFromISO8601 ]
		/// <summary>
		/// Converte para DateTime uma string representando data e hora no formato ISO 8601 em um dos seguintes formatos:
		///		Unspecified: 2009-06-15T13:45:30.0000000
		///		UTC: 2009-06-15T13:45:30.0000000Z
		///		Local: 2009-06-15T13:45:30.0000000-07:00
		/// </summary>
		/// <param name="dataHoraISO8601">Parâmetro string representando data e hora no formato ISO 8601</param>
		/// <returns></returns>
		public static DateTime converteDateTimeFromISO8601(string dataHoraISO8601)
		{
			DateTime dtDataHoraResp;
			if (DateTime.TryParse(dataHoraISO8601, null, System.Globalization.DateTimeStyles.RoundtripKind, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ converteDdMmYyyyHhMmSsParaDateTime ]
		/// <summary>
		/// Converte o texto que representa uma data/hora para DateTime
		/// </summary>
		/// <param name="strDdMmYyyyHhMmSs">
		/// Texto representando uma data/hora, com ou sem separadores, sendo que a parte da hora é opcional.
		/// </param>
		/// <returns>
		/// Retorna a data/hora como DateTime, se não for possível fazer a conversão, retorna DateTime.MinValue
		/// </returns>
		public static DateTime converteDdMmYyyyHhMmSsParaDateTime(string strDdMmYyyyHhMmSs)
		{
			#region [ Declarações ]
			char c;
			string strDia = "";
			string strMes = "";
			string strAno = "";
			string strHora = "";
			string strMinuto = "";
			string strSegundo = "";
			string strFormato;
			string strDataHoraAConverter;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			#endregion

			#region [ Dia ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strDia += c;
				if (strDia.Length == 2) break;
			}
			while (strDia.Length < 2) strDia = '0' + strDia;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Mês ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMes += c;
				if (strMes.Length == 2) break;
			}
			while (strMes.Length < 2) strMes = '0' + strMes;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Ano ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strAno += c;
				if (strAno.Length == 4) break;
			}
			if (strAno.Length == 2)
			{
				if (converteInteiro(strAno) >= 80)
					strAno = "19" + strAno;
				else
					strAno = "20" + strAno;
			}
			#endregion

			#region [ Remove separador(es) entre a data e hora, se houver ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				if (!isDigit(strDdMmYyyyHhMmSs[0]))
					strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				else
					break;
			}
			#endregion

			#region [ Hora ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strHora += c;
				if (strHora.Length == 2) break;
			}
			while (strHora.Length < 2) strHora = '0' + strHora;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Minuto ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMinuto += c;
				if (strMinuto.Length == 2) break;
			}
			while (strMinuto.Length < 2) strMinuto = '0' + strMinuto;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strDdMmYyyyHhMmSs.Length > 0) && (!isDigit(strDdMmYyyyHhMmSs[0]))) strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
			#endregion

			#region [ Segundo ]
			while (strDdMmYyyyHhMmSs.Length > 0)
			{
				c = strDdMmYyyyHhMmSs[0];
				strDdMmYyyyHhMmSs = strDdMmYyyyHhMmSs.Substring(1, strDdMmYyyyHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strSegundo += c;
				if (strSegundo.Length == 2) break;
			}
			while (strSegundo.Length < 2) strSegundo = '0' + strSegundo;
			#endregion

			#region [ Monta máscara ]
			strFormato = Cte.DataHora.FmtDia +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtAno +
						 ' ' +
						 Cte.DataHora.FmtHora +
						 Cte.DataHora.FmtMin +
						 Cte.DataHora.FmtSeg;
			#endregion

			#region [ Monta data/hora normalizada ]
			strDataHoraAConverter = strDia +
									strMes +
									strAno +
									' ' +
									strHora +
									strMinuto +
									strSegundo;
			#endregion

			if (DateTime.TryParseExact(strDataHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region[ converteYyyyMmDdHhMmSsParaDateTime ]
		/// <summary>
		/// Converte o texto que representa uma data/hora para DateTime
		/// </summary>
		/// <param name="strYyyyMmDdHhMmSs">
		/// Texto representando uma data/hora, com ou sem separadores, sendo que a parte da hora é opcional.
		/// </param>
		/// <returns>
		/// Retorna a data/hora como DateTime, se não for possível fazer a conversão, retorna DateTime.MinValue
		/// </returns>
		public static DateTime converteYyyyMmDdHhMmSsParaDateTime(string strYyyyMmDdHhMmSs)
		{
			#region [ Declarações ]
			char c;
			string strDia = "";
			string strMes = "";
			string strAno = "";
			string strHora = "";
			string strMinuto = "";
			string strSegundo = "";
			string strFormato;
			string strDataHoraAConverter;
			DateTime dtDataHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			#endregion

			#region [ Ano ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strAno += c;
				if (strAno.Length == 4) break;
			}
			if (strAno.Length == 2)
			{
				if (converteInteiro(strAno) >= 80)
					strAno = "19" + strAno;
				else
					strAno = "20" + strAno;
			}
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Mês ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMes += c;
				if (strMes.Length == 2) break;
			}
			while (strMes.Length < 2) strMes = '0' + strMes;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Dia ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strDia += c;
				if (strDia.Length == 2) break;
			}
			while (strDia.Length < 2) strDia = '0' + strDia;
			#endregion

			#region [ Remove separador(es) entre a data e hora, se houver ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				if (!isDigit(strYyyyMmDdHhMmSs[0]))
					strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				else
					break;
			}
			#endregion

			#region [ Hora ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strHora += c;
				if (strHora.Length == 2) break;
			}
			while (strHora.Length < 2) strHora = '0' + strHora;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Minuto ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMinuto += c;
				if (strMinuto.Length == 2) break;
			}
			while (strMinuto.Length < 2) strMinuto = '0' + strMinuto;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strYyyyMmDdHhMmSs.Length > 0) && (!isDigit(strYyyyMmDdHhMmSs[0]))) strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
			#endregion

			#region [ Segundo ]
			while (strYyyyMmDdHhMmSs.Length > 0)
			{
				c = strYyyyMmDdHhMmSs[0];
				strYyyyMmDdHhMmSs = strYyyyMmDdHhMmSs.Substring(1, strYyyyMmDdHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strSegundo += c;
				if (strSegundo.Length == 2) break;
			}
			while (strSegundo.Length < 2) strSegundo = '0' + strSegundo;
			#endregion

			#region [ Monta máscara ]
			strFormato = Cte.DataHora.FmtAno +
						 Cte.DataHora.FmtMes +
						 Cte.DataHora.FmtDia +
						 ' ' +
						 Cte.DataHora.FmtHora +
						 Cte.DataHora.FmtMin +
						 Cte.DataHora.FmtSeg;
			#endregion

			#region [ Monta data/hora normalizada ]
			strDataHoraAConverter = strAno +
									strMes +
									strDia +
									' ' +
									strHora +
									strMinuto +
									strSegundo;
			#endregion

			if (DateTime.TryParseExact(strDataHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			return DateTime.MinValue;
		}
		#endregion

		#region [ converteMmDdYyyyHhMmSsAmPmParaDateTime ]
		public static DateTime converteMmDdYyyyHhMmSsAmPmParaDateTime(string strMmDdYyyyHhMmSsAmPm)
		{
			#region [ Declarações ]
			string[] vDT;
			string[] vData;
			string[] vHora;
			string strBlocoData = "";
			string strBlocoHora = "";
			string strAmPm = "";
			string strDia = "";
			string strMes = "";
			string strAno = "";
			string strHora = "";
			string strMin = "";
			string strSeg = "";
			bool blnTemData = false;
			bool blnTemHora = false;
			string strDataHoraAConverter;
			string strFormato;
			DateTime dtDataHoraResp;
			// É necessário usar uma "cultura" que possua a designação AM/PM, caso contrário, o DateTime.TryParseExact() não irá funcionar
			CultureInfo myCultureInfo = new CultureInfo("en-US");
			#endregion

			if (strMmDdYyyyHhMmSsAmPm == null) return DateTime.MinValue;
			if (strMmDdYyyyHhMmSsAmPm.Trim().Length == 0) return DateTime.MinValue;

			vDT = strMmDdYyyyHhMmSsAmPm.Split(' ');
			if (vDT.Length >= 1) strBlocoData = vDT[0].Trim();
			if (vDT.Length >= 2) strBlocoHora = vDT[1].Trim();
			if (vDT.Length >= 3) strAmPm = vDT[2].Trim();

			if (strBlocoData.Length > 0)
			{
				vData = strBlocoData.Split('/');
				if (vData.Length >= 1) { strMes = vData[0].Trim(); blnTemData = true; }
				if (vData.Length >= 2) { strDia = vData[1].Trim(); blnTemData = true; }
				if (vData.Length >= 3) { strAno = vData[2].Trim(); blnTemData = true; }
			}

			if (strBlocoHora.Length > 0)
			{
				vHora = strBlocoHora.Split(':');
				if (vHora.Length >= 1) { strHora = vHora[0].Trim(); blnTemHora = true; }
				if (vHora.Length >= 2) { strMin = vHora[1].Trim(); blnTemHora = true; }
				if (vHora.Length >= 3) { strSeg = vHora[2].Trim(); blnTemHora = true; }

				if (strHora.Length == 1) strHora = '0' + strHora;
				if (strMin.Length == 1) strMin = '0' + strMin;
				if (strSeg.Length == 1) strSeg = '0' + strSeg;
			}

			if (blnTemData && blnTemHora)
			{
				strDataHoraAConverter = strAno +
										strMes +
										strDia +
										' ' +
										strHora +
										strMin +
										strSeg;
				strFormato = Cte.DataHora.FmtAno +
							 Cte.DataHora.FmtMes +
							 Cte.DataHora.FmtDia +
							 ' ' +
							 Cte.DataHora.FmtHora12 +
							 Cte.DataHora.FmtMin +
							 Cte.DataHora.FmtSeg;
				if (strAmPm.Length > 0)
				{
					strDataHoraAConverter += ' ' + strAmPm;
					strFormato += ' ' + Cte.DataHora.FmtAmPm;
				}
				if (DateTime.TryParseExact(strDataHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			}
			else if (blnTemData)
			{
				strDataHoraAConverter = strAno +
										strMes +
										strDia;
				strFormato = Cte.DataHora.FmtAno +
							 Cte.DataHora.FmtMes +
							 Cte.DataHora.FmtDia;
				if (DateTime.TryParseExact(strDataHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
			}

			return DateTime.MinValue;
		}
		#endregion

		#region[ converteHhMmParaTimeSpan ]
		/// <summary>
		/// Converte o texto que representa um horário HH:mm no formato 24h para TimeSpan
		/// </summary>
		/// <param name="strHhMm">
		/// Texto representando um horário HH:mm no formato 24h, com ou sem separadores
		/// </param>
		/// <returns>
		/// Retorna o horário como TimeSpan, se não for possível fazer a conversão, retorna TimeSpan.MinValue (-10675199.02:48:05.4775808)
		/// </returns>
		public static TimeSpan converteHhMmParaTimeSpan(string strHhMm)
		{
			#region [ Declarações ]
			char c;
			string strHora = "";
			string strMinuto = "";
			string strFormato;
			string strHoraAConverter;
			DateTime dtHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			#endregion

			#region [ Consistência ]
			if (strHhMm == null) return TimeSpan.MinValue;
			if (digitos(strHhMm).Length == 0) return TimeSpan.MinValue;
			#endregion

			#region [ Hora ]
			while (strHhMm.Length > 0)
			{
				c = strHhMm[0];
				strHhMm = strHhMm.Substring(1, strHhMm.Length - 1);
				if (!isDigit(c)) break;
				strHora += c;
				if (strHora.Length == 2) break;
			}
			while (strHora.Length < 2) strHora = '0' + strHora;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strHhMm.Length > 0) && (!isDigit(strHhMm[0]))) strHhMm = strHhMm.Substring(1, strHhMm.Length - 1);
			#endregion

			#region [ Minuto ]
			while (strHhMm.Length > 0)
			{
				c = strHhMm[0];
				strHhMm = strHhMm.Substring(1, strHhMm.Length - 1);
				if (!isDigit(c)) break;
				strMinuto += c;
				if (strMinuto.Length == 2) break;
			}
			while (strMinuto.Length < 2) strMinuto = '0' + strMinuto;
			#endregion

			#region [ Monta máscara ]
			strFormato = Cte.DataHora.FmtHora +
						":" +
						 Cte.DataHora.FmtMin;
			#endregion

			#region [ Monta data/hora normalizada ]
			strHoraAConverter = strHora +
								":" +
								strMinuto;
			#endregion

			if (DateTime.TryParseExact(strHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtHoraResp)) return dtHoraResp.TimeOfDay;
			return TimeSpan.MinValue;
		}
		#endregion

		#region[ converteHhMmSsParaTimeSpan ]
		/// <summary>
		/// Converte o texto que representa um horário no formato 24h para TimeSpan
		/// </summary>
		/// <param name="strHhMmSs">
		/// Texto representando um horário no formato 24h, com ou sem separadores
		/// </param>
		/// <returns>
		/// Retorna o horário como TimeSpan, se não for possível fazer a conversão, retorna TimeSpan.MinValue (-10675199.02:48:05.4775808)
		/// </returns>
		public static TimeSpan converteHhMmSsParaTimeSpan(string strHhMmSs)
		{
			#region [ Declarações ]
			char c;
			string strHora = "";
			string strMinuto = "";
			string strSegundo = "";
			string strFormato;
			string strHoraAConverter;
			DateTime dtHoraResp;
			CultureInfo myCultureInfo = new CultureInfo("pt-BR");
			#endregion

			#region [ Consistência ]
			if (strHhMmSs == null) return TimeSpan.MinValue;
			if (digitos(strHhMmSs).Length == 0) return TimeSpan.MinValue;
			#endregion

			#region [ Hora ]
			while (strHhMmSs.Length > 0)
			{
				c = strHhMmSs[0];
				strHhMmSs = strHhMmSs.Substring(1, strHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strHora += c;
				if (strHora.Length == 2) break;
			}
			while (strHora.Length < 2) strHora = '0' + strHora;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strHhMmSs.Length > 0) && (!isDigit(strHhMmSs[0]))) strHhMmSs = strHhMmSs.Substring(1, strHhMmSs.Length - 1);
			#endregion

			#region [ Minuto ]
			while (strHhMmSs.Length > 0)
			{
				c = strHhMmSs[0];
				strHhMmSs = strHhMmSs.Substring(1, strHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strMinuto += c;
				if (strMinuto.Length == 2) break;
			}
			while (strMinuto.Length < 2) strMinuto = '0' + strMinuto;
			#endregion

			#region [ Remove separador, se houver ]
			if ((strHhMmSs.Length > 0) && (!isDigit(strHhMmSs[0]))) strHhMmSs = strHhMmSs.Substring(1, strHhMmSs.Length - 1);
			#endregion

			#region [ Segundo ]
			while (strHhMmSs.Length > 0)
			{
				c = strHhMmSs[0];
				strHhMmSs = strHhMmSs.Substring(1, strHhMmSs.Length - 1);
				if (!isDigit(c)) break;
				strSegundo += c;
				if (strSegundo.Length == 2) break;
			}
			while (strSegundo.Length < 2) strSegundo = '0' + strSegundo;
			#endregion

			#region [ Monta máscara ]
			strFormato = Cte.DataHora.FmtHora +
						":" +
						 Cte.DataHora.FmtMin +
						 ":" +
						 Cte.DataHora.FmtSeg;
			#endregion

			#region [ Monta data/hora normalizada ]
			strHoraAConverter = strHora +
								":" +
								strMinuto +
								":" +
								strSegundo;
			#endregion

			if (DateTime.TryParseExact(strHoraAConverter, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtHoraResp)) return dtHoraResp.TimeOfDay;
			return TimeSpan.MinValue;
		}
		#endregion

		#region[ digitos ]
		public static string digitos(string texto)
		{
			StringBuilder d = new StringBuilder();
			if (texto == null) return "";
			for (int i = 0; i < texto.Length; i++)
			{
				if ((texto[i] >= '0') && (texto[i] <= '9')) d.Append(texto[i]);
			}
			return d.ToString();
		}
		#endregion

		#region[ converteInteiro ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo inteiro
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="valor">
		/// Texto representando um número inteiro
		/// </param>
		/// <returns>
		/// Retorna um número do tipo inteiro
		/// </returns>
		public static Int64 converteInteiro(string valor)
		{
			Int64 intResultado = 0;

			if (valor == null) return 0;

			string strValor = valor.Trim();
			if (strValor.Length == 0) return 0;

			try
			{
				intResultado = Int64.Parse(strValor);
			}
			catch (Exception)
			{
				intResultado = 0;
			}

			return intResultado;
		}
		#endregion

		#region[ converteInteiro ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo inteiro
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="valor">
		/// Texto representando um número inteiro
		/// </param>
		/// <param name="valorDefault">
		/// Valor que será retornado no caso da conversão falhar
		/// </param>
		/// <returns>
		/// Retorna um número do tipo inteiro
		/// </returns>
		public static Int64 converteInteiro(string valor, Int64 valorDefault)
		{
			Int64 intResultado = 0;

			if (valor == null) return valorDefault;

			string strValor = valor.Trim();
			if (strValor.Length == 0) return valorDefault;

			try
			{
				intResultado = Int64.Parse(strValor);
			}
			catch (Exception)
			{
				intResultado = valorDefault;
			}

			return intResultado;
		}
		#endregion

		#region [ converteNumeroDecimal ]
		/// <summary>
		/// Converte o número representado pelo texto do parâmetro em um número do tipo decimal
		/// Se não conseguir realizar a conversão, será retornado zero
		/// </summary>
		/// <param name="numero">
		/// Texto representando um número decimal
		/// </param>
		/// <returns>
		/// Retorna um número do tipo decimal
		/// </returns>
		public static decimal converteNumeroDecimal(String numero)
		{
			#region [ Declarações ]
			int i;
			char c_separador_decimal;
			String s_numero_aux;
			String s_inteiro = "";
			String s_centavos = "";
			int intSinal = 1;
			decimal decFracionario;
			decimal decInteiro;
			decimal decResultado;
			#endregion

			if (numero == null) return 0;
			if (numero.Trim().Length == 0) return 0;

			numero = numero.Trim();

			if (numero.IndexOf('-') != -1) intSinal = -1;

			c_separador_decimal = retornaSeparadorDecimal(numero);

			#region [ Separa parte inteira e os centavos ]
			s_numero_aux = numero.Replace(c_separador_decimal, 'V');
			String[] v = s_numero_aux.Split('V');
			for (i = 0; i < v.Length; i++)
			{
				if (v[i] == null) v[i] = "";
			}
			// Falha ao determinar o separador de decimal, então calcula como se não houvesse centavos
			if (v.Length > 2)
			{
				s_inteiro = digitos(numero);
			}
			else
			{
				if (v.Length >= 1) s_inteiro = digitos(v[0]);
				if (v.Length >= 2) s_centavos = digitos(v[1]);
			}
			if (s_inteiro.Length == 0) s_inteiro = "0";
			s_centavos = s_centavos.PadRight(2, '0');
			#endregion

			decInteiro = (decimal)converteInteiro(s_inteiro);
			decFracionario = (decimal)converteInteiro(s_centavos) / (decimal)Math.Pow(10, s_centavos.Length);
			decResultado = intSinal * (decInteiro + decFracionario);
			return decResultado;
		}
		#endregion

		#region [ obtemDescricaoAnaliseCredito ]
		public static string obtemDescricaoAnaliseCredito(int statusAnaliseCredito)
		{
			string strResp = "";

			if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.ST_INICIAL)
			{
				strResp = "Aguardando Análise Inicial";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_PENDENTE)
			{
				strResp = "Pendente";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
			{
				strResp = "Pendente Vendas";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_ENDERECO)
			{
				strResp = "Pendente Endereço";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
			{
				strResp = "Crédito OK";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK_AGUARDANDO_DEPOSITO)
			{
				strResp = "Crédito OK (aguardando depósito)";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)
			{
				strResp = "Crédito OK (depósito aguardando desbloqueio)";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.NAO_ANALISADO)
			{
				strResp = "Pedido Sem Análise de Crédito";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)
			{
				strResp = "Crédito OK (aguardando pagto boleto AV)";
			}
			else if (statusAnaliseCredito == Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_PAGTO_ANTECIPADO_BOLETO)
			{
				strResp = "Pendente - Pagto Antecipado Boleto";
			}
			else
			{
				strResp = "Código Desconhecido: " + statusAnaliseCredito.ToString();
			}

			return strResp;
		}
		#endregion

		#region [ executaManutencaoArqLogAtividade ]
		/// <summary>
		/// Apaga os arquivos de log de atividade antigos
		/// </summary>
		public static bool executaManutencaoArqLogAtividade(out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "executaManutencaoArqLogAtividade()";
			String strMsg;
			DateTime dtCorte = DateTime.Now.AddDays(-Global.Cte.LogAtividade.CorteArqLogEmDias);
			string strDataCorte = dtCorte.ToString(Global.Cte.DataHora.FmtYYYYMMDD);
			string[] ListaArqLog;
			string strNomeArq;
			int i;
			int intQtdeApagada = 0;
			DateTime dtHrInicio = DateTime.Now;
			#endregion

			strMsgErro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				#region[ Apaga arquivos de log de atividade antigos ]
				ListaArqLog = Directory.GetFiles(Global.Cte.LogAtividade.PathLogAtividade, "*." + Global.Cte.LogAtividade.ExtensaoArqLog, SearchOption.TopDirectoryOnly);
				for (i = 0; i < ListaArqLog.Length; i++)
				{
					strNomeArq = Global.extractFileName(ListaArqLog[i]);
					strNomeArq = strNomeArq.Substring(0, strDataCorte.Length);
					if (string.Compare(strNomeArq, strDataCorte) < 0)
					{
						File.Delete(ListaArqLog[i]);
						intQtdeApagada++;
					}
				}
				#endregion

				strMsg = "Rotina " + NOME_DESTA_ROTINA + " concluída com sucesso: " + intQtdeApagada.ToString() + " arquivos excluídos (duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio) + ")";
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				return false;
			}
		}
		#endregion

		#region[ extractFileName ]
		public static string extractFileName(string fileName)
		{
			string strResp = "";
			for (int i = (fileName.Length - 1); i >= 0; i--)
			{
				if (fileName[i] == (char)92) return strResp;
				if (fileName[i] == (char)47) return strResp;
				if (fileName[i] == (char)58) return strResp;
				strResp = fileName[i] + strResp;
			}
			return strResp;
		}
		#endregion

		#region [ filtraAcentuacao ]
		public static String filtraAcentuacao(String texto)
		{
			#region [ Declarações ]
			String strResp;
			#endregion

			if (texto == null) return texto;
			if (texto.Length == 0) return texto;

			strResp = texto.ToString();
			if (strResp.IndexOf('á') != -1) strResp = strResp.Replace('á', 'a');
			if (strResp.IndexOf('à') != -1) strResp = strResp.Replace('à', 'a');
			if (strResp.IndexOf('ã') != -1) strResp = strResp.Replace('ã', 'a');
			if (strResp.IndexOf('â') != -1) strResp = strResp.Replace('â', 'a');
			if (strResp.IndexOf('ä') != -1) strResp = strResp.Replace('ä', 'a');
			if (strResp.IndexOf('é') != -1) strResp = strResp.Replace('é', 'e');
			if (strResp.IndexOf('è') != -1) strResp = strResp.Replace('è', 'e');
			if (strResp.IndexOf('ê') != -1) strResp = strResp.Replace('ê', 'e');
			if (strResp.IndexOf('ë') != -1) strResp = strResp.Replace('ë', 'e');
			if (strResp.IndexOf('í') != -1) strResp = strResp.Replace('í', 'i');
			if (strResp.IndexOf('ì') != -1) strResp = strResp.Replace('ì', 'i');
			if (strResp.IndexOf('î') != -1) strResp = strResp.Replace('î', 'i');
			if (strResp.IndexOf('ï') != -1) strResp = strResp.Replace('ï', 'i');
			if (strResp.IndexOf('ó') != -1) strResp = strResp.Replace('ó', 'o');
			if (strResp.IndexOf('ò') != -1) strResp = strResp.Replace('ò', 'o');
			if (strResp.IndexOf('õ') != -1) strResp = strResp.Replace('õ', 'o');
			if (strResp.IndexOf('ô') != -1) strResp = strResp.Replace('ô', 'o');
			if (strResp.IndexOf('ö') != -1) strResp = strResp.Replace('ö', 'o');
			if (strResp.IndexOf('ú') != -1) strResp = strResp.Replace('ú', 'u');
			if (strResp.IndexOf('ù') != -1) strResp = strResp.Replace('ù', 'u');
			if (strResp.IndexOf('û') != -1) strResp = strResp.Replace('û', 'u');
			if (strResp.IndexOf('ü') != -1) strResp = strResp.Replace('ü', 'u');
			if (strResp.IndexOf('ç') != -1) strResp = strResp.Replace('ç', 'c');
			if (strResp.IndexOf('ñ') != -1) strResp = strResp.Replace('ñ', 'n');
			if (strResp.IndexOf('ÿ') != -1) strResp = strResp.Replace('ÿ', 'y');

			if (strResp.IndexOf('Á') != -1) strResp = strResp.Replace('Á', 'A');
			if (strResp.IndexOf('À') != -1) strResp = strResp.Replace('À', 'A');
			if (strResp.IndexOf('Ã') != -1) strResp = strResp.Replace('Ã', 'A');
			if (strResp.IndexOf('Â') != -1) strResp = strResp.Replace('Â', 'A');
			if (strResp.IndexOf('Ä') != -1) strResp = strResp.Replace('Ä', 'A');
			if (strResp.IndexOf('É') != -1) strResp = strResp.Replace('É', 'E');
			if (strResp.IndexOf('È') != -1) strResp = strResp.Replace('È', 'E');
			if (strResp.IndexOf('Ê') != -1) strResp = strResp.Replace('Ê', 'E');
			if (strResp.IndexOf('Ë') != -1) strResp = strResp.Replace('Ë', 'E');
			if (strResp.IndexOf('Í') != -1) strResp = strResp.Replace('Í', 'I');
			if (strResp.IndexOf('Ì') != -1) strResp = strResp.Replace('Ì', 'I');
			if (strResp.IndexOf('Î') != -1) strResp = strResp.Replace('Î', 'I');
			if (strResp.IndexOf('Ï') != -1) strResp = strResp.Replace('Ï', 'I');
			if (strResp.IndexOf('Ó') != -1) strResp = strResp.Replace('Ó', 'O');
			if (strResp.IndexOf('Ò') != -1) strResp = strResp.Replace('Ò', 'O');
			if (strResp.IndexOf('Õ') != -1) strResp = strResp.Replace('Õ', 'O');
			if (strResp.IndexOf('Ô') != -1) strResp = strResp.Replace('Ô', 'O');
			if (strResp.IndexOf('Ö') != -1) strResp = strResp.Replace('Ö', 'O');
			if (strResp.IndexOf('Ú') != -1) strResp = strResp.Replace('Ú', 'U');
			if (strResp.IndexOf('Ù') != -1) strResp = strResp.Replace('Ù', 'U');
			if (strResp.IndexOf('Û') != -1) strResp = strResp.Replace('Û', 'U');
			if (strResp.IndexOf('Ü') != -1) strResp = strResp.Replace('Ü', 'U');
			if (strResp.IndexOf('Ç') != -1) strResp = strResp.Replace('Ç', 'C');
			if (strResp.IndexOf('Ñ') != -1) strResp = strResp.Replace('Ñ', 'N');

			return strResp;
		}
		#endregion

		#region [ filtraAmpersand ]
		public static string filtraAmpersand(string texto)
		{
			#region [ Declarações ]
			string strResp;
			#endregion

			if (texto == null) return texto;

			strResp = texto.Trim();

			if (strResp.Contains("&")) strResp = strResp.Replace("&", " e ").Trim();

			while (strResp.Contains("  "))
			{
				strResp = strResp.Replace("  ", " ").Trim();
			}

			return strResp;
		}
		#endregion

		#region [ formataDataDdMmYyyyComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YYYY
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YYYY
		/// </returns>
		public static String formataDataDdMmYyyyComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador);
		}

		public static string formataDataDdMmYyyyComSeparador(DateTime? data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return ((DateTime)data).ToString(Global.Cte.DataHora.FmtDdMmYyyyComSeparador);
		}
		#endregion

		#region [ formataDataDdMmYyyyHhMmComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YYYY HH:MM
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YYYY HH:MM
		/// </returns>
		public static String formataDataDdMmYyyyHhMmComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmComSeparador);
		}
		#endregion

		#region [ formataDataDdMmYyyyHhMmSsComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YYYY HH:MM:SS
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato DD/MM/YYYY HH:MM
		/// </returns>
		public static String formataDataDdMmYyyyHhMmSsComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmSsComSeparador);
		}
		#endregion

		#region [ formataDataHoraYyyyMmDdTHhMmSs ]
		/// <summary>
		/// A partir de uma data/hora do tipo DateTime, formata um texto com a representação da data no formato yyyy-mm-ddThh:mm:ss
		/// </summary>
		/// <param name="data">
		/// Data/hora em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato yyyy-mm-ddThh:mm:ss
		/// </returns>
		public static String formataDataHoraYyyyMmDdTHhMmSs(DateTime dataHora)
		{
			if (dataHora == null) return "";
			if (dataHora == DateTime.MinValue) return "";
			return dataHora.ToString(Global.Cte.DataHora.FmtYyyyMmDdComSeparador) + "T" + dataHora.ToString(Global.Cte.DataHora.FmtHhMmSsComSeparador);
		}
		#endregion

		#region [ formataDataYyyyMmDdComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato YYYY-MM-DD
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato YYYY-MM-DD
		/// </returns>
		public static String formataDataYyyyMmDdComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtYyyyMmDdComSeparador);
		}
		#endregion

		#region [ formataDataYyyyMmDdHhMmSsComSeparador ]
		/// <summary>
		/// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato YYYY-MM-DD HH:MM:SS
		/// </summary>
		/// <param name="data">
		/// Data em parâmetro do tipo DateTime
		/// </param>
		/// <returns>
		/// Retorna a data representada em um texto no formato YYYY-MM-DD HH:MM:SS
		/// </returns>
		public static String formataDataYyyyMmDdHhMmSsComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtYyyyMmDdHhMmSsComSeparador);
		}

		public static String formataDataYyyyMmDdHhMmSsComSeparador(DateTime? data)
		{
			if (data == null) return "";
			return formataDataYyyyMmDdHhMmSsComSeparador((DateTime)data);
		}
		#endregion

		#region[ formataDuracaoHMS ]
		public static string formataDuracaoHMS(TimeSpan ts)
		{
			StringBuilder sb = new StringBuilder();
			// Dias
			if (ts.Days > 0) sb.Append(ts.Days.ToString() + 'd');
			// Horas
			if (sb.ToString().Length == 0)
			{
				if (ts.Hours > 0) sb.Append(ts.Hours.ToString() + 'h');
			}
			else
			{
				sb.Append(ts.Hours.ToString().PadLeft(2, '0') + 'h');
			}
			// Minutos
			if (sb.ToString().Length == 0)
			{
				if (ts.Minutes > 0) sb.Append(ts.Minutes.ToString() + 'm');
			}
			else
			{
				sb.Append(ts.Minutes.ToString().PadLeft(2, '0') + 'm');
			}
			// Segundos
			if (sb.ToString().Length == 0)
			{
				sb.Append(ts.Seconds.ToString() + 's');
			}
			else
			{
				sb.Append(ts.Seconds.ToString().PadLeft(2, '0') + 's');
			}
			return sb.ToString();
		}
		#endregion

		#region [ formataHoraHhMmSsComSeparador ]
		public static String formataHoraHhMmSsComSeparador(DateTime data)
		{
			if (data == null) return "";
			if (data == DateTime.MinValue) return "";
			return data.ToString(Global.Cte.DataHora.FmtHhMmSsComSeparador);
		}
		#endregion

		#region [ formata_hhnnss_para_hh_nn ]
		public static string formata_hhnnss_para_hh_nn(string hhnnss)
		{
			#region [ Declarações ]
			string s_hhnnss;
			string s_resp = "";
			string s_hora = "";
			string s_min = "";
			#endregion

			s_hhnnss = (hhnnss ?? "");
			s_hhnnss = digitos(s_hhnnss);
			if ((s_hhnnss.Length == 4) || (s_hhnnss.Length == 6))
			{
				s_hora = s_hhnnss.Substring(0, 2);
				s_min = s_hhnnss.Substring(2, 2);
				s_resp = s_hora + ":" + s_min;
				return s_resp;
			}
			else
			{
				return hhnnss;
			}
		}
		#endregion

		#region [ formata_hhnnss_para_hh_nn_ss ]
		public static string formata_hhnnss_para_hh_nn_ss(string hhnnss)
		{
			#region [ Declarações ]
			string s_hhnnss;
			string s_resp = "";
			string s_hora = "";
			string s_min = "";
			string s_seg = "";
			#endregion

			s_hhnnss = (hhnnss ?? "");
			s_hhnnss = digitos(s_hhnnss);
			if (s_hhnnss.Length == 6)
			{
				s_hora = s_hhnnss.Substring(0, 2);
				s_min = s_hhnnss.Substring(2, 2);
				s_seg = s_hhnnss.Substring(4, 2);
				s_resp = s_hora + ":" + s_min + ":" + s_seg;
				return s_resp;
			}
			else if (s_hhnnss.Length == 4)
			{
				s_hora = s_hhnnss.Substring(0, 2);
				s_min = s_hhnnss.Substring(2, 2);
				s_resp = s_hora + ":" + s_min;
				return s_resp;
			}
			else
			{
				return hhnnss;
			}
		}
		#endregion

		#region [ formataInteiro ]
		public static String formataInteiro(int numero)
		{
			String strResp = "";
			String strNumero;
			int intPonto = 0;

			strNumero = digitos(numero.ToString());
			for (int i = strNumero.Length - 1; i >= 0; i--)
			{
				intPonto++;
				strResp = strNumero[i] + strResp;
				if ((intPonto % 3 == 0) && (i != 0)) strResp = '.' + strResp;
			}
			return strResp;
		}
		#endregion

		#region [ formataMoeda ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato monetário
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um valor monetário
		/// </param>
		/// <returns>
		/// Retorna um texto com formato monetário
		/// </returns>
		public static String formataMoeda(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.00");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataMoedaClearsale ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato monetário no padrão usado pela Clearsale: Número (20,4)
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um valor monetário
		/// </param>
		/// <returns>
		/// Retorna um texto com formato monetário no padrão usado pela Clearsale: Número (20,4)
		/// </returns>
		public static String formataMoedaClearsale(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.0000");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 5), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(".", "");
				strValorFormatado = strValorFormatado.Replace(",", "");
				strValorFormatado = strValorFormatado.Replace("V", ".");
			}
			else if (strSeparadorDecimal.Equals(","))
			{
				strValorFormatado = strValorFormatado.Replace(",", "V");
				strValorFormatado = strValorFormatado.Replace(",", "");
				strValorFormatado = strValorFormatado.Replace(".", "");
				strValorFormatado = strValorFormatado.Replace("V", ".");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataPercentual ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato de percentual
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um percentual
		/// </param>
		/// <returns>
		/// Retorna um texto com formato de percentual
		/// </returns>
		public static String formataPercentual(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.00");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataPercentualCom1Decimal ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato de percentual
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um percentual
		/// </param>
		/// <returns>
		/// Retorna um texto com formato de percentual
		/// </returns>
		public static String formataPercentualCom1Decimal(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.0");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 2), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataPercentualCom2Decimais ]
		/// <summary>
		/// Formata o campo do tipo numérico em um texto com formato de percentual
		/// </summary>
		/// <param name="valor">
		/// Valor numérico representando um percentual
		/// </param>
		/// <returns>
		/// Retorna um texto com formato de percentual
		/// </returns>
		public static String formataPercentualCom2Decimais(decimal valor)
		{
			String strValorFormatado;
			String strSeparadorDecimal;
			strValorFormatado = valor.ToString("###,###,##0.00");
			// Verifica se o separador decimal é vírgula ou ponto
			strSeparadorDecimal = Texto.leftStr(Texto.rightStr(strValorFormatado, 3), 1);
			if (strSeparadorDecimal.Equals("."))
			{
				strValorFormatado = strValorFormatado.Replace(".", "V");
				strValorFormatado = strValorFormatado.Replace(",", ".");
				strValorFormatado = strValorFormatado.Replace("V", ",");
			}
			return strValorFormatado;
		}
		#endregion

		#region [ formataTimeSpanHorario ]
		public static string formataTimeSpanHorario(TimeSpan ts, string responseIfMinValue)
		{
			if (ts == TimeSpan.MinValue) return responseIfMinValue;
			return ts.ToString();
		}
		#endregion

		#region [ formaPagtoPedidoDescricao ]
		/// <summary>
		/// Retorna a descrição da forma de pagamento do pedido (dinheiro, depósito, cheque, boleto, cartão)
		/// </summary>
		/// <param name="codigo">
		/// Código da forma de pagamento do pedido
		/// </param>
		/// <returns>
		/// Retorna a descrição da forma de pagamento do pedido
		/// </returns>
		public static String formaPagtoPedidoDescricao(short codigo)
		{
			String strResp = "";

			if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_DINHEIRO)
				strResp = "Dinheiro";
			else if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_DEPOSITO)
				strResp = "Depósito";
			else if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CHEQUE)
				strResp = "Cheque";
			else if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				strResp = "Boleto";
			else if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				strResp = "Cartão (internet)";
			else if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO_MAQUINETA)
				strResp = "Cartão (maquineta)";
			else if (codigo == Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV)
				strResp = "Boleto AV";

			return strResp;
		}
		#endregion

		#region [ GetConfigurationValue ]
		public static string GetConfigurationValue(string key)
		{
			Assembly service = Assembly.GetAssembly(typeof(FinanceiroProjectInstaller));
			Configuration config = ConfigurationManager.OpenExeConfiguration(service.Location);
			if (config.AppSettings.Settings[key] != null)
			{
				return config.AppSettings.Settings[key].Value;
			}
			else
			{
				throw new IndexOutOfRangeException("Settings collection does not contain the requested key:" + key);
			}
		}
		#endregion

		#region[ gravaEventLog ]
		public static void gravaEventLog(string strSource, string strMessage, EventLogEntryType eTipoMensagem)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "gravaEventLog()";
			#endregion

			if (strMessage.Length > 32000)
			{
				// Tamanho máximo do log do Event Viewer é de 32766 caracteres !!
				strMessage = strMessage.Substring(0, 32000) + " (truncado ...)";
			}

			try
			{
				gravaLogAtividade(strMessage);
				System.Diagnostics.EventLog.WriteEntry(strSource, strMessage, eTipoMensagem);
			}
			catch (Exception ex)
			{
				gravaLogAtividade(NOME_DESTA_ROTINA + ": " + ex.ToString());
			}
		}
		#endregion

		#region[ gravaEventLog (overload) ]
		public static void gravaEventLog(string strMessage, EventLogEntryType eTipoMensagem)
		{
			gravaEventLog(Cte.Aplicativo.ID_SISTEMA_EVENTLOG, strMessage, eTipoMensagem);
		}
		#endregion

		#region[ gravaLogAtividade ]
		/// <summary>
		/// Grava a informação do parâmetro no arquivo de log, junto com a data/hora
		/// Se o parâmetro for 'null', será gravada uma linha em branco no arquivo
		/// Se o parâmetro uma string vazia, será gravada uma linha apenas com a data/hora
		/// </summary>
		/// <param name="mensagem"></param>
		public static void gravaLogAtividade(string mensagem)
		{
			string linha;
			DateTime dataHora = DateTime.Now;
			const string FmtHHMMSS = Cte.DataHora.FmtHora + ":" + Cte.DataHora.FmtMin + ":" + Cte.DataHora.FmtSeg + "." + Cte.DataHora.FmtMiliSeg;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			const string FmtYYYYMMDD = Cte.DataHora.FmtAno + Cte.DataHora.FmtMes + Cte.DataHora.FmtDia;
			string strArqLog = Global.barraInvertidaAdd(Global.Cte.LogAtividade.PathLogAtividade) +
							   DateTime.Now.ToString(FmtYYYYMMDD) +
							   "." +
							   Global.Cte.LogAtividade.ExtensaoArqLog;
			if (mensagem == null)
				linha = "";
			else
				linha = dataHora.ToString(FmtHHMMSS) + ": " + mensagem;

			try
			{
				rwlArqLogAtividade.AcquireWriterLock(60 * 1000);
				try
				{
					using (StreamWriter sw = new StreamWriter(strArqLog, true, encode))
					{
						sw.WriteLine(linha);
						sw.Flush();
						sw.Close();
					}
				}
				finally
				{
					rwlArqLogAtividade.ReleaseWriterLock();
				}
			}
			catch (Exception)
			{
				// Nop
			}
		}
		#endregion

		#region [ isCnpjOk ]
		/// <summary>
		/// Indica se o CNPJ está ok, ou seja, se os dígitos verificadores conferem
		/// </summary>
		/// <param name="cnpj">
		/// CNPJ a testar
		/// </param>
		/// <returns>
		/// true: CNPJ válido
		/// false: CNPJ inválido
		/// </returns>
		public static bool isCnpjOk(String cnpj)
		{
			String s_cnpj;
			String p1 = "543298765432";
			String p2 = "6543298765432";
			bool tudo_igual;
			int i;
			int d;

			if (cnpj == null) return false;

			s_cnpj = digitos(cnpj);
			if (s_cnpj.Length != 14) return false;

			// Dígitos são todos iguais?
			tudo_igual = true;
			for (i = 0; i < (s_cnpj.Length - 1); i++)
			{
				if (!s_cnpj.Substring(i, 1).Equals(s_cnpj.Substring(i + 1, 1)))
				{
					tudo_igual = false;
					break;
				}
			}
			if (tudo_igual) return false;

			// Verifica o primeiro check digit
			d = 0;
			for (i = 0; i < 12; i++)
			{
				d = d + int.Parse(p1.Substring(i, 1)) * int.Parse(s_cnpj.Substring(i, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cnpj.Substring(12, 1))) return false;

			// Verifica o segundo check digit
			d = 0;
			for (i = 0; i < 13; i++)
			{
				d = d + int.Parse(p2.Substring(i, 1)) * int.Parse(s_cnpj.Substring(i, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cnpj.Substring(13, 1))) return false;

			// Ok
			return true;
		}
		#endregion

		#region [ isCnpjCpfOk ]
		/// <summary>
		/// Indica se o CNPJ/CPF está ok, ou seja, se os dígitos verificadores conferem
		/// </summary>
		/// <param name="cnpj_cpf">
		/// CNPJ/CPF a testar
		/// </param>
		/// <returns>
		/// true: CNPJ/CPF válido
		/// false: CNPJ/CPF inválido
		/// </returns>
		public static bool isCnpjCpfOk(String cnpj_cpf)
		{
			String s;
			if (cnpj_cpf == null) return false;
			s = digitos(cnpj_cpf);
			if (s.Length == 11)
			{
				return isCpfOk(s);
			}
			else if (s.Length == 14)
			{
				return isCnpjOk(s);
			}
			return false;
		}
		#endregion

		#region [ isCpfOk ]
		/// <summary>
		/// Indica se o CPF está ok, ou seja, se os dígitos verificadores conferem
		/// </summary>
		/// <param name="cpf">
		/// CPF a testar
		/// </param>
		/// <returns>
		/// true: CPF válido
		/// false: CPF inválido
		/// </returns>
		public static bool isCpfOk(String cpf)
		{
			int i;
			int d;
			bool tudo_igual;
			String s_cpf;

			if (cpf == null) return false;

			s_cpf = digitos(cpf);
			if (s_cpf.Length != 11) return false;

			// Dígitos todos iguais?
			tudo_igual = true;
			for (i = 0; i < (s_cpf.Length - 1); i++)
			{
				if (!s_cpf.Substring(i, 1).Equals(s_cpf.Substring(i + 1, 1)))
				{
					tudo_igual = false;
					break;
				}
			}
			if (tudo_igual) return false;

			// Verifica o primeiro check digit
			d = 0;
			for (i = 1; i <= 9; i++)
			{
				d = d + (11 - i) * int.Parse(s_cpf.Substring(i - 1, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cpf.Substring(9, 1))) return false;

			// Verifica o segundo check digit
			d = 0;
			for (i = 2; i <= 10; i++)
			{
				d = d + (12 - i) * int.Parse(s_cpf.Substring(i - 1, 1));
			}
			d = 11 - (d % 11);
			if (d > 9) d = 0;
			if (d != int.Parse(s_cpf.Substring(10, 1))) return false;

			// Ok
			return true;
		}
		#endregion

		#region[ isDigit ]
		public static bool isDigit(char c)
		{
			if ((c >= '0') && (c <= '9')) return true;
			return false;
		}
		#endregion

		#region [ isHorarioDentroIntervalo ]
		public static bool isHorarioDentroIntervalo(TimeSpan horarioRef, TimeSpan horarioInicio, TimeSpan horarioTermino)
		{
			// Assegura que há apenas a informação do horário, sem nenhum dia
			horarioRef = new TimeSpan(horarioRef.Hours, horarioRef.Minutes, horarioRef.Seconds);
			horarioInicio = new TimeSpan(horarioInicio.Hours, horarioInicio.Minutes, horarioInicio.Seconds);
			horarioTermino = new TimeSpan(horarioTermino.Hours, horarioTermino.Minutes, horarioTermino.Seconds);

			// Verifica se o horário de término está no mesmo dia que o horário de início ou se virou o dia
			if (horarioTermino < horarioInicio)
			{
				if (horarioRef <= horarioTermino)
				{
					// A data de início está 1 dia antes do horário de referência, mas como o dia do TimeSpan está zerado, a data de referência é que será deslocada 1 dia p/ o futuro
					horarioRef = horarioRef.Add(new TimeSpan(1, 0, 0, 0)); // Days, Hours, Minutes, Seconds
				}

				horarioTermino = horarioTermino.Add(new TimeSpan(1, 0, 0, 0)); // Days, Hours, Minutes, Seconds
			}

			if ((horarioRef >= horarioInicio) && (horarioRef < horarioTermino)) return true;
			return false;
		}
		#endregion

		#region [ isLetra ]
		public static bool isLetra(char c)
		{
			return ((Char.ToUpper(c) >= 'A') && (Char.ToUpper(c) <= 'Z'));
		}
		#endregion

		#region [ isLetra ]
		public static bool isLetra(String c)
		{
			if (c == null) return false;
			if (c.Trim().Length == 0) return false;

			for (int i = 0; i < c.Length; i++)
			{
				if (!isLetra(c[i])) return false;
			}
			return true;
		}
		#endregion

		#region [ isPedidoFilhote ]
		/// <summary>
		/// Analisa se o número do pedido é de um pedido-base ou de um pedido-filhote
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido a ser analisado
		/// </param>
		/// <returns>
		/// true: trata-se de um número de pedido-filhote
		/// false: trata-se de um número de pedido-base
		/// </returns>
		public static bool isPedidoFilhote(String numeroPedido)
		{
			if (numeroPedido == null) return false;
			numeroPedido = numeroPedido.Trim();
			numeroPedido = normalizaNumeroPedido(numeroPedido);
			if (numeroPedido.IndexOf(Cte.Etc.COD_SEPARADOR_FILHOTE) > -1) return true;
			return false;
		}
		#endregion

		#region [ logEstoqueMontaDecremento ]
		public static string logEstoqueMontaDecremento(int quantidade, string fabricante, string produto)
		{
			#region [ Declarações ]
			String s;
			#endregion

			s = " -" + quantidade.ToString() + "x" + produto;
			if (fabricante != null)
			{
				if (fabricante.Trim().Length > 0)
				{
					s += "(" + fabricante + ")";
				}
			}

			return s;
		}
		#endregion

		#region [ logEstoqueMontaIncremento ]
		public static string logEstoqueMontaIncremento(int quantidade, string fabricante, string produto)
		{
			#region [ Declarações ]
			String s;
			#endregion

			s = " +" + quantidade.ToString() + "x" + produto;
			if (fabricante != null)
			{
				if (fabricante.Trim().Length > 0)
				{
					s += "(" + fabricante + ")";
				}
			}

			return s;
		}
		#endregion

		#region [ logProdutoMonta ]
		public static string logProdutoMonta(int quantidade, string fabricante, string produto)
		{
			#region [ Declarações ]
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			sbLog.Append(" " + quantidade.ToString() + "x" + produto);
			if (fabricante != null)
			{
				if (fabricante.Trim().Length > 0)
				{
					sbLog.Append("(" + fabricante + ")");
				}
			}

			return sbLog.ToString();
		}
		#endregion

		#region [ montaIdInstanciaServicoEmailSubject ]
		public static string montaIdInstanciaServicoEmailSubject()
		{
			string strResp = "FINSVC";
			if (Cte.Aplicativo.IDENTIFICADOR_AMBIENTE_OWNER.Length > 0)
			{
				strResp += "/" + Cte.Aplicativo.IDENTIFICADOR_AMBIENTE_OWNER;
			}
			else
			{
				strResp += "/" + Cte.Aplicativo.ID_SISTEMA_EVENTLOG;
			}

			if ((Cte.Aplicativo.AMBIENTE_EXECUCAO ?? "").Length > 0)
			{
				strResp += "/" + Cte.Aplicativo.AMBIENTE_EXECUCAO;
			}

			return "[" + strResp + "]";
		}
		#endregion

		#region [ normalizaNumeroPedido ]
		public static String normalizaNumeroPedido(String pedido)
		{
			String id_pedido;
			String s = "";
			String s_ano = "";
			String s_num = "";
			String s_filhote = "";
			char c;

			if (pedido == null) return "";
			id_pedido = pedido.Trim().ToUpper();
			if (id_pedido.Length == 0) return "";

			for (int i = 0; i < id_pedido.Length; i++)
			{
				if (isDigit(id_pedido[i]))
					s_num += id_pedido[i];
				else
					break;
			}
			if (s_num.Length == 0) return "";

			for (int i = 0; i < id_pedido.Length; i++)
			{
				c = id_pedido[i];
				if (isLetra(c))
				{
					if (s_ano.Length == 0)
					{
						s_ano = c.ToString();
					}
					else
					{
						if (s_filhote.Length == 0) s_filhote = c.ToString();
					}
				}
			}
			if (s_ano.Length == 0) return "";
			s_num = s_num.PadLeft(Cte.Etc.TAM_MIN_NUM_PEDIDO, '0');
			s = s_num + s_ano;
			if (s_filhote.Length > 0) s += Cte.Etc.COD_SEPARADOR_FILHOTE + s_filhote;
			return s;
		}
		#endregion

		#region [ retornaNumeroPedidoBase ]
		/// <summary>
		/// No caso do número do pedido ser de um pedido-filhote, retorna apenas a parte do número
		/// correspondente ao pedido-base.
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido a ser analisado.
		/// </param>
		/// <returns>
		/// Retorna apenas a parte do número que identifica o pedido-base.
		/// </returns>
		public static String retornaNumeroPedidoBase(String numeroPedido)
		{
			if (numeroPedido == null) return "";
			numeroPedido = numeroPedido.Trim();
			if (numeroPedido.Length == 0) return "";
			numeroPedido = normalizaNumeroPedido(numeroPedido);
			if (numeroPedido.IndexOf(Cte.Etc.COD_SEPARADOR_FILHOTE) == -1) return numeroPedido;
			return numeroPedido.Substring(0, numeroPedido.IndexOf(Cte.Etc.COD_SEPARADOR_FILHOTE));
		}
		#endregion

		#region [ retornaSeparadorDecimal ]
		/// <summary>
		/// Analisa o texto do parâmetro que representa um valor monetário para determinar se o separador decimal é ponto ou vírgula
		/// </summary>
		/// <param name="numero">
		/// Texto representando um valor monetário
		/// </param>
		/// <returns>
		/// Retorna o caracter usado para representação do separador decimal (de centavos)
		/// </returns>
		private static char retornaSeparadorDecimal(String valorMonetario)
		{
			int i;
			int n_ponto = 0;
			int n_virgula = 0;
			int n_digitos_finais = 0;
			int n_digitos_iniciais = 0;
			char c;
			String s_numero;
			char c_ult_sep = '\0';
			char c_separador_decimal;

			if (valorMonetario == null) return ',';
			if (valorMonetario.Trim().Length == 0) return ',';

			s_numero = valorMonetario.Trim();
			for (i = s_numero.Length - 1; i >= 0; i--)
			{
				c = s_numero[i];
				if (c == '.')
				{
					n_ponto++;
					if (c_ult_sep == '\0') c_ult_sep = c;
				}
				else if (c == ',')
				{
					n_virgula++;
					if (c_ult_sep == '\0') c_ult_sep = c;
				}
				if (isDigit(c) && (n_ponto == 0) && (n_virgula == 0)) n_digitos_finais++;
				if (isDigit(c) && ((n_ponto > 0) || (n_virgula > 0))) n_digitos_iniciais++;
			}

			// Default
			c_separador_decimal = ',';
			if (c_ult_sep == '.')
			{
				if ((n_ponto == 1) && (n_virgula == 0) && (n_digitos_iniciais <= 3) && (n_digitos_finais == 3))
				{
					// NOP: Considera 123.456 como cento e vinte e três mil e quatrocentos e cinquenta e seis
				}
				else if (n_ponto == 1)
				{
					c_separador_decimal = '.';
				}
			}
			else if (c_ult_sep == ',')
			{
				if ((n_virgula > 1) && (n_ponto == 0)) c_separador_decimal = '.';
			}
			return c_separador_decimal;
		}
		#endregion

		#region[ sqlMontaDateTimeParaSqlDateTime ]
		public static string sqlMontaDateTimeParaSqlDateTime(DateTime dtReferencia)
		{
			string strDataHora;
			string strSql;

			if (dtReferencia == null) return "NULL";
			if (dtReferencia == DateTime.MinValue) return "NULL";

			strDataHora = dtReferencia.ToString(Cte.DataHora.FmtAno) +
						  "-" +
						  dtReferencia.ToString(Cte.DataHora.FmtMes) +
						  "-" +
						  dtReferencia.ToString(Cte.DataHora.FmtDia) +
						  " " +
						  dtReferencia.ToString(Cte.DataHora.FmtHora) +
						  ":" +
						  dtReferencia.ToString(Cte.DataHora.FmtMin) +
						  ":" +
						  dtReferencia.ToString(Cte.DataHora.FmtSeg);
			strSql = "Convert(datetime, '" + strDataHora + "', 120)";
			return strSql;
		}
		#endregion

		#region[ sqlMontaDateTimeParaSqlDateTimeSomenteData ]
		public static string sqlMontaDateTimeParaSqlDateTimeSomenteData(DateTime dtReferencia)
		{
			string strData;
			string strSql;
			strData = dtReferencia.ToString(Cte.DataHora.FmtAno) +
					  "-" +
					  dtReferencia.ToString(Cte.DataHora.FmtMes) +
					  "-" +
					  dtReferencia.ToString(Cte.DataHora.FmtDia);
			strSql = "Convert(datetime, '" + strData + "', 120)";
			return strSql;
		}
		#endregion

		#region[ sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30 14:27:01
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <param name="strAlias">
		/// Informa o nome do Alias, caso seja informado uma string vazia, então será usado o nome do próprio campo.
		/// </param>
		/// <returns></returns>
		public static string sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador(string strNomeCampo, string strAlias)
		{
			string strResposta;
			if ((strAlias.Trim().Length == 0) && (strNomeCampo.IndexOf('(') == -1)) strAlias = strNomeCampo;
			strResposta = "Coalesce(Convert(varchar(19), " + strNomeCampo + ", 121), '')";
			if (strAlias.Length > 0) strResposta += " AS " + strAlias;
			return strResposta;
		}
		#endregion

		#region[ sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador ]
		/// <summary>
		/// Monta a expressão SQL para retornar um campo do tipo datetime como
		/// texto varchar no formato: 2009-01-30 14:27:01
		/// </summary>
		/// <param name="strNomeCampo">
		/// Informa o nome do campo do banco de dados que deve ser do tipo datetime
		/// </param>
		/// <returns></returns>
		public static string sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador(string strNomeCampo)
		{
			return sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador(strNomeCampo, "");
		}
		#endregion

		#region [ sqlMontaGetdateSomenteData ]
		/// <summary>
		/// Monta uma expressão para obter a data do Sql Server com data apenas, sem a hora
		/// </summary>
		/// <returns>
		/// Retorna uma expressão para obter a data do Sql Server com data apenas, sem a hora
		/// </returns>
		public static string sqlMontaGetdateSomenteData()
		{
			string strResposta;
			strResposta = "Convert(datetime, Convert(varchar(10), getdate(), 121), 121)";
			return strResposta;
		}
		#endregion

		#region [ sqlMontaGetdateSomenteHora ]
		/// <summary>
		/// Monta uma expressão para obter a hora do Sql Server com a hora no formato hh:mm:ss, sem a parte da data
		/// </summary>
		/// <returns>
		/// Retorna uma expressão para obter a hora do Sql Server com a hora no formato hh:mm:ss, sem a parte da data
		/// </returns>
		public static string sqlMontaGetdateSomenteHora()
		{
			string strResposta;
			strResposta = "Convert(varchar(8), getdate(), 108)";
			return strResposta;
		}
		#endregion

		#region [ obtemXmlNodeFirstChildValue ]
		public static string obtemXmlNodeFirstChildValue(XmlNode xmlNode)
		{
			if (xmlNode == null) return null;
			if (xmlNode.ChildNodes.Count == 0) return null;
			return xmlNode.FirstChild.Value;
		}
		#endregion

		#region [ obtemXmlChildNodeValue ]
		public static string obtemXmlChildNodeValue(XmlNode xmlNode, string xmlNodeName)
		{
			return obtemXmlChildNodeValue(xmlNode, xmlNodeName, "");
		}

		public static string obtemXmlChildNodeValue(XmlNode xmlNode, string nodeName, string valorDefault)
		{
			string strResp;

			try
			{
				if (xmlNode == null) return valorDefault;
				if (xmlNode.ChildNodes == null) return valorDefault;
				if (xmlNode.ChildNodes.Count == 0) return valorDefault;

				strResp = xmlNode[nodeName].InnerText;
			}
			catch (Exception)
			{
				return valorDefault;
			}

			return strResp;
		}
		#endregion

		#region [ serializaObjectToXml ]
		public static string serializaObjectToXml(object obj)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Global.serializaObjectToXml()";
			XmlSerializer xmlWriter;
			StringWriter stringWriter = new System.IO.StringWriter();
			#endregion

			if (obj == null) return "";

			try
			{
				xmlWriter = new XmlSerializer(obj.GetType());
				xmlWriter.Serialize(stringWriter, obj);
				return stringWriter.ToString();
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception\n" + ex.ToString());
				return "";
			}
		}
		#endregion

		#region [ sqlMontaCaseWhenParametroStringVaziaComoNull ]
		/// <summary>
		/// Para parâmetros de objetos SqlCommand que são usados para datas expressas como
		/// string no formato YYYY-MM-DD, monta uma expressão CASE WHEN para gravar NULL
		/// quando o valor do parâmetro for uma string vazia.
		/// Lembrando que o SQL Server grava automaticamente a data de 1900-01-01 quando
		/// converte uma string vazia para um campo datetime.
		/// </summary>
		/// <param name="nomeParametroDoCommand">Nome do parâmetro (ex: @dtVencto)</param>
		/// <returns>Retorna um texto contendo uma expressão CASE WHEN, ex: CASE WHEN @dt_vencto='' THEN NULL ELSE @dt_vencto END</returns>
		public static String sqlMontaCaseWhenParametroStringVaziaComoNull(String nomeParametroDoCommand)
		{
			String strResp;
			strResp = "CASE WHEN " + nomeParametroDoCommand + " = '' THEN NULL ELSE " + nomeParametroDoCommand + " END";
			return strResp;
		}
		#endregion

		#region [ sqlMontaCaseWhenParametroStringVaziaComoProprioCampo ]
		/// <summary>
		/// Para parâmetros de objetos SqlCommand que são usados para datas expressas como
		/// string no formato YYYY-MM-DD, monta uma expressão CASE WHEN para gravar o valor
		/// do próprio campo quando o valor do parâmetro for uma string vazia, ou seja, não altera o valor.
		/// Lembrando que o SQL Server grava automaticamente a data de 1900-01-01 quando
		/// converte uma string vazia para um campo datetime.
		/// </summary>
		/// <param name="nomeParametroDoCommand">Nome do parâmetro (ex: @dtVencto)</param>
		/// <returns>Retorna um texto contendo uma expressão CASE WHEN, ex: CASE WHEN @dt_vencto='' THEN dt_vencto ELSE @dt_vencto END</returns>
		public static String sqlMontaCaseWhenParametroStringVaziaComoProprioCampo(String nomeParametroDoCommand, String nomeCampo)
		{
			String strResp;
			strResp = "CASE WHEN " + nomeParametroDoCommand + " = '' THEN " + nomeCampo + " ELSE " + nomeParametroDoCommand + " END";
			return strResp;
		}
		#endregion

		#region [ stPagtoPedidoDescricao ]
		/// <summary>
		/// Obtém a descrição do status de pagamento do pedido
		/// </summary>
		/// <param name="status">
		/// Código do status de pagamento do pedido
		/// </param>
		/// <returns>
		/// Retorna uma descrição do status de pagamento do pedido
		/// </returns>
		public static String stPagtoPedidoDescricao(String status)
		{
			String strResp = "";

			if (status == null) return "";
			status = status.Trim();
			if (status.Length == 0) return "";

			if (status.Equals(Cte.StPagtoPedido.ST_PAGTO_PAGO))
				strResp = "Pago";
			else if (status.Equals(Cte.StPagtoPedido.ST_PAGTO_NAO_PAGO))
				strResp = "Não-Pago";
			else if (status.Equals(Cte.StPagtoPedido.ST_PAGTO_PARCIAL))
				strResp = "Pago Parcial";

			return strResp;
		}
		#endregion

		#endregion
	}
}
