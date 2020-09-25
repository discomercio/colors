#region[ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.Win32;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Security.Cryptography;
#endregion

namespace Financeiro
{
    class Global
    {
        #region [ Constantes ]
        public class Cte
        {
            #region[ Versão do Aplicativo ]
            public class Aplicativo
            {
                public const string NOME_OWNER = "Artven";
                public const string NOME_SISTEMA = "Financeiro";
                public const string VERSAO_NUMERO = "1.37";
                public const string VERSAO_DATA = "28.AGO.2020";
                public const string VERSAO = VERSAO_NUMERO + " - " + VERSAO_DATA;
                public const string M_ID = NOME_SISTEMA + "  -  " + VERSAO;
                public const string M_DESCRICAO = "Módulo para execução de rotinas financeiras";
            }
            #endregion

            #region[ Comentário sobre as versões ]
            /*================================================================================================
			 * v 1.00 - 14.09.2009 - por HHO
			 *        Início.
			 *        Este programa realiza diversas rotinas financeiras.
			 *        Esta versão inicial foi colocada em produção mesmo não estando totalmente completa,
			 *        apenas os recursos iniciais mais básicos e essenciais estavam concluídos.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.08 - 29.10.2009 - por HHO
			 *		  Implementação do cadastramento de boletos avulsos (vinculados a pedidos).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.09 - 03.11.2009 - por HHO
			 *		  Implementação de consistências durante o cadastramento de boletos: solicita confirmação
			 *		  se já houver boleto emitido ou se o status de pagamento estiver "pago".
			 * -----------------------------------------------------------------------------------------------
			 * v 1.10 - 10.11.2009 - por HHO
			 *		  1) Criação de um cache local para lista com os nomes de clientes usado para o auto com-
			 *		  plete de campos do formulário. O cache é feito em arquivo xml, sendo que os dados são
			 *		  atualizados na primeira execução do dia.
			 *		  2) Na carga do arquivo de retorno, quando o form perde o foco ele entra em um estado em
			 *		  que fica "congelado" e não atualiza mais os campos (progresso, relógio, etc). Entretanto,
			 *		  o processamento continua ocorrendo. Para contornar esse problema, foi introduzido uma
			 *		  chamada ao Application.DoEvents() no laço do processamento.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.11 - 04.01.2010 - por HHO
			 *        Implementação de aprimoramentos nas operações de fluxo de caixa, com operações de edição
			 *        em lotes.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.12 - 09.01.2010 - por HHO
			 *		  Implementação do novo relatório "Relatório Analítico de Movimentos".
			 * -----------------------------------------------------------------------------------------------
			 * v 1.13 - 06.02.2010 - por HHO
			 *		  Implementação de alteração nas operações de cadastrar boletos. Quando houver instrução
			 *		  de protesto, apenas algumas parcelas serão geradas com a instrução de protesto devido
			 *		  ao custo cobrado pelos cartórios.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14 - 22.02.2010 - por HHO
			 *		  Disponibilização da consulta e impressão da carteira em atraso.
			 *		  Implementação de tratamento na carga do arquivo de retorno para a ocorrência 22 (título
			 *		  com pagamento cancelado).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14(B) - 25.02.2010 - por HHO
			 *		  Disponibilização de consulta ao fluxo de caixa dentro do módulo de cobrança.
			 *		  Aumento do número de linhas do grid nos painéis de cadastramento de lançamentos de 
			 *		  débito e crédito em lote (de 30 para 80).
			 *		  Foi mantida o mesmo número de versão do aplicativo por não haver nenhuma alteração no
			 *		  banco de dados e porque as alterações não são críticas.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.14(C) - 04.03.2010 - por HHO
			 *		  Alteração no processamento do arquivo de retorno (ocorrência 28) para tratar a situação
			 *		  em que não é informado o número de controle do participante. Neste caso, a identificação
			 *		  do boleto será feita pelo nosso número.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.15 - 11.03.2010 - por HHO
			 *		  Como preparação para a emissão da NFe, todos os cadastros que contêm endereços foram 
			 *		  alterados para separar as informações "número" e "complemento" do endereço.
			 *		  Como consequência disso, este módulo foi adaptado p/ montar o endereço considerando 
			 *		  a possibilidade do endereço já estar no formato novo ou não.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16 - 24.03.2010 - por HHO
			 *		  Inclusão do parâmetro "encoding" durante a leitura de dados do arquivo de retorno.
			 *		  Inclusão de filtragem de caracteres acentuados durante a gravação do arquivo de remessa.
			 *		  Inclusão da cláusula "COLLATE Latin1_General_CI_AI" em consultas SQL que usam o operador
			 *		  "LIKE" em campos que podem conter acentos.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16 - 08.07.2010 - por HHO
			 *	      Alteração na carga do arquivo de retorno para tratar a situação em que a ocorrência 28
			 *	      (débito de tarifas/custas) não se refere a um boleto cadastrado no sistema. Neste caso,
			 *	      apenas as alterações na tabela t_FIN_BOLETO_ITEM foram ignoradas.
			 *	      OBS: esta versão foi gerada a partir da última versão que está em produção, pois há
			 *	      uma versão em desenvolvimento pendente há vários meses (recursos p/ cobrança de clientes
			 *	      em atraso). Por esta razão, a data desta versão não estará coerente com a data do backup
			 *	      do projeto, já que o backup foi colocado na sequência cronológica imediatamente a seguir
			 *	      da versão usada como base desta alteração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16 - 05.08.2010 - por HHO
			 *		  Alteração das rotinas que enviam boletos por e-mail para gerarem uma imagem jpeg para
			 *		  ser anexada. A técnica usada foi baseada em utilizar o mesmo código html já gerado e
			 *		  carregá-lo em um componente WebBrowser. Em seguida, é capturada a imagem contida dentro
			 *		  desse componente e é gerada a imagem jpeg.
			 *		  Além disso, com isso se resolve o problema causado pelas diferenças de exibição do 
			 *		  código html em função do programa utilizado p/ sua visualização.
			 *	      OBS: esta versão foi gerada a partir da última versão que está em produção, pois há
			 *	      uma versão em desenvolvimento pendente há vários meses (recursos p/ cobrança de clientes
			 *	      em atraso). Por esta razão, a data desta versão não estará coerente com a data do backup
			 *	      do projeto, já que o backup foi colocado na sequência cronológica imediatamente a seguir
			 *	      da versão usada como base desta alteração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(D) - 16.09.2010 - por HHO
			 *        Retirada da rotina que realiza a limpeza da tabela t_FIN_LOG porque foi criado um
			 *        serviço para ser instalado no servidor que será responsável por executar tarefas dessa
			 *        natureza automaticamente.
			 *	      OBS: esta versão foi gerada a partir da última versão que está em produção, pois há
			 *	      uma versão em desenvolvimento pendente há vários meses (recursos p/ cobrança de clientes
			 *	      em atraso). Por esta razão, a data desta versão não estará coerente com a data do backup
			 *	      do projeto, já que o backup foi colocado na sequência cronológica imediatamente a seguir
			 *	      da versão usada como base desta alteração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(E) - 04.01.2011 - por HHO
			 *        Inclusão de um filtro na consulta da carteira em atraso para permitir a seleção somente
			 *        de clientes que estejam com a 1ª parcela do boleto atrasada. O objetivo é facilitar a
			 *        identificação dos clientes que não receberam os boletos enviados pelo banco.
			 *	      OBS: esta versão foi gerada a partir da última versão que está em produção, pois há
			 *	      uma versão em desenvolvimento pendente há vários meses (recursos p/ cobrança de clientes
			 *	      em atraso). Por esta razão, a data desta versão não estará coerente com a data do backup
			 *	      do projeto, já que o backup foi colocado na sequência cronológica imediatamente a seguir
			 *	      da versão usada como base desta alteração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(E) - 01.02.2011 - por HHO
			 *		  Ajuste emergencial da rotina de carga do arquivo de retorno de boletos devido à presença
			 *		  de registros de ocorrência 16 (título pago em cheque) sem o conteúdo do campo "nº con-
			 *		  trole do participante". Foi feita uma adaptação p/ tentar identificar o registro do
			 *		  boleto através do campo "nosso número".
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(F) - 02.03.2011 - por HHO
			 *		  Alteração na consulta da carteira em atraso p/ incluir as seguintes colunas nos dados
			 *		  de resposta:
			 *			1) Nº parcela mais antiga em atraso
			 *			2) Vendedor
			 *			3) Parceiro
			 *			4) UF
			 *		  Além disso, foi criada uma função para gerar o relatório em planilha MS-Excel.
			 *		  Para deixar a geração da planilha o mais flexível possível com relação às diferentes
			 *		  versões do Excel, foi adotada a técnica do "late binding". Inicialmente, a implementação
			 *		  foi feita usando o "early binding" por facilitar a codificação (disponibiliza o auto
			 *		  complete) e ter melhor desempenho, mas para funcionar com várias versões seria neces-
			 *		  sário que o Interop do Excel referenciado no projeto fosse o da versão mais antiga que 
			 *		  se pretende usar. Além disso, ao instalar uma versão mais nova do Excel na máquina do
			 *		  projeto, o Interop provavelmente seria atualizado p/ a versão mais recente. Devido a
			 *		  isso, foi feita a opção de se usar o "late binding", que dispensa o uso do Interop e se
			 *		  mostra como melhor solução p/ o uso em ambientes com várias versões do Excel.
			 *	      OBS: esta versão foi gerada a partir da última versão que está em produção, pois há
			 *	      uma versão em desenvolvimento pendente há vários meses (recursos p/ cobrança de clientes
			 *	      em atraso). Por esta razão, a data desta versão não estará coerente com a data do backup
			 *	      do projeto, já que o backup foi colocado na sequência cronológica imediatamente a seguir
			 *	      da versão usada como base desta alteração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(F) - 08.04.2011 - por HHO
			 *		  Alteração na consulta da carteira em atraso p/ incluir o valor do RA bruto ao gerar
			 *		  o relatório em planilha MS-Excel.
			 *	      OBS: esta versão foi gerada a partir da última versão que está em produção, pois há
			 *	      uma versão em desenvolvimento pendente há vários meses (recursos p/ cobrança de clientes
			 *	      em atraso). Por esta razão, a data desta versão não estará coerente com a data do backup
			 *	      do projeto, já que o backup foi colocado na sequência cronológica imediatamente a seguir
			 *	      da versão usada como base desta alteração.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(G) - 25.07.2011 - por HHO
			 *		  Alterações no cadastramento de boletos, geração do arquivo de remessa e carga do arquivo
			 *		  de retorno para passar a tratar a emissão de boletos pela OLD03, lembrando que até o mo-
			 *		  mento eram emitidos boletos somente pela OLD01.
			 *		  A partir de agora, o módulo Financeiro está apto a emitir boletos para múltiplas empre-
			 *		  sas cedentes.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.16(H) - 28.07.2011 - por HHO
			 *		  Alterações nos seguintes painéis p/ melhorar a consulta e identificação do cedente do
			 *		  boleto:
			 *		  1) Painel de consulta de boletos.
			 *		  2) Painel de consulta de ocorrências (boletos).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - 15.02.2012 - por HHO
			 *		  Implementação dos relatórios:
			 *			1) Fluxo de Caixa: Relatório Sintético de Movimentos (Rateio)
			 *			2) Fluxo de Caixa: Relatório Analítico de Movimentos (Rateio)
			 *			
			 *		  Implementação das alterações:
			 *			1) No cadastramento de lançamentos em lote foi disponibilizada a opção de selecionar
			 *			   a quantidade de lançamentos que se pode cadastrar em uma única operação. A moti-
			 *			   vação para essa alteração é devido à conferência do valor total do lote de lança-
			 *			   mentos.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - 05.04.2012 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Alteração na planilha Excel gerada no painel "Cobrança: Administração da Carteira em
			 *		  Atraso" (FCobrancaAdministracao) para incluir uma nova coluna informando o email do
			 *		  parceiro.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - 17.04.2012 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Alteração na planilha Excel gerada no painel "Cobrança: Administração da Carteira em
			 *		  Atraso" (FCobrancaAdministracao) para incluir a informação da data em que o pedido
			 *		  teve o status da análise de crédito alterado p/ "Crédito OK".
			 * -----------------------------------------------------------------------------------------------
			 * v 1.17 - 25.04.2012 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Alteração na planilha Excel gerada no painel "Cobrança: Administração da Carteira em
			 *		  Atraso" (FCobrancaAdministracao) para separar em uma coluna própria a informação da
			 *		  data em que o pedido teve o status da análise de crédito alterado p/ "Crédito OK".
			 *		  No caso de haver mais do que um pedido com pagamentos em atraso para o mesmo cliente,
			 *		  será exibida a data mais antiga.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.18 - 17.05.2012 - por HHO
			 *		  Alteração no tratamento do arquivo de retorno para a ocorrência 16 (título pago em 
			 *		  cheque), mais especificamente na rotina trataOcorrencia16TituloPagoEmCheque().
			 *		  Foi constatado que há situações em que um boleto já liquidado ou baixado é pago
			 *		  novamente através de cheque. Neste caso, se o cheque for devidamente compensado, a con-
			 *		  firmação do pagamento será informada através de uma ocorrência 17 (liquidação após baixa
			 *		  ou título não registrado). E no caso de uma ocorrência 17, a regra é sempre criar um
			 *		  novo lançamento no fluxo de caixa. Por este motivo, o lançamento no fluxo de caixa
			 *		  que já estava com o status 'Pago' ou 'Boleto baixado' estava sendo alterado para o
			 *		  status 'Boleto pago com cheque (vinculado)' e assim permanecia de forma definitiva.
			 *		  Portanto, no caso do boleto já estar liquidado ou baixado e houver uma ocorrência 16,
			 *		  o lançamento existente no fluxo de caixa e o registro no histórico de pagamentos dos
			 *		  pedidos não devem ser atualizados, já que quando a ocorrência 17 for recebida em decor-
			 *		  rência da compensação do cheque, serão criados novos registros no fluxo de caixa e
			 *		  no histórico de pagamentos dos pedidos.
			 *		  A título informativo, o boleto pode ser liquidado por uma ocorrência 06 (liquidação
			 *		  normal) ou uma ocorrência 15 (liquidação em cartório). E o boleto pode ser baixado
			 *		  por uma ocorrência 09 (baixado automat. via arquivo) ou uma ocorrência 10 (baixado
			 *		  conforme instruções da agência).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.18 - 18.09.2012 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Implementação de função para gerar planilha Excel com os dados da consulta do fluxo de
			 *		  caixa.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.19 - 23.10.2012 - por HHO
			 *		  As funções implementadas diretamente no SQL Server estavam sob o esquema cujo nome era
			 *		  o mesmo do usuário do banco de dados. Como algumas dessas funções começaram a ser usadas
			 *		  em 'computed columns', surgiu uma dificuldade ao copiar o BD de produção em ambiente de
			 *		  homologação, já que os nomes dos usuários de BD são diferentes.
			 *		  Para uniformizar o uso dessas funções nos diferentes ambientes, a solução usada foi
			 *		  recriar as funções sob o esquema 'dbo'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.20 - 04.02.2013 - por HHO
			 *		  Alterações em FCobrancaAdministracao para aumentar de 3 p/ 4 dígitos o tamanho máximo
			 *		  dos campos que definem a faixa da quantidade de dias em atraso. Alteração da formatação
			 *		  da quantidade de dias em atraso p/ exibir o separador de milhar no grid de resultados,
			 *		  na impressão e na planilha Excel.
			 *		  Análise das consultas SQL realizadas pelo painel FCobrancaAdministracao para reduzir o
			 *		  tempo de consulta. Reestruturação de alguns índices no BD.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.21 - 02.04.2013 - por HHO
			 *		  Inclusão do nº porta para o servidor SMTP usado para o envio de boletos por email.
			 *		  Foi criado o campo no banco de dados e no painel de configuração dos parâmetros de
			 *		  email.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.21 - 26.09.2013 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  No painel de cadastramento de lançamentos de débito em lote (FFluxoDebitoLote.cs), 
			 *		  foi adicionada a coluna 'Plano de Contas Empresa' no grid para permitir cadastrar
			 *		  lançamentos de diferentes empresas na mesma operação. Essa necessidade surgiu devido
			 *		  aos lançamentos da folha de pagamento que misturam funcionários de diferentes empresas.
			 *		  O lançamento em uma única operação possibilita a checagem do valor total antes de
			 *		  confirmar a gravação dos dados.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.21 - 06.11.2013 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Nos relatórios "Relatório Sintético de Movimentos" e "Relatório Analítico de Movimentos"
			 *		  foram incluídos os filtros do tipo checkbox "CPF" e "CNPJ".
			 *		  Quando esses filtros estão assinalados, a consulta considera apenas os registros do
			 *		  fluxo de caixa que tenham informação armazenada no campo "cnpj_cpf".
			 *		  Se ambos estiverem assinalados, são considerados os registros que possuam a informação
			 *		  de CPF ou CNPJ.
			 *		  É importante ressaltar que há registros que estão com o campo "cnpj_cpf" vazio.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.22 - 29.01.2014 - por HHO
			 *		  Inclusão de botão para selecionar todos os boletos na operação de enviá-los por email.
			 *		  Otimização da consulta da carteira em atraso recuperando apenas as parcelas dos clientes
			 *		  selecionados (antes estavam sendo recuperadas todas as parcelas de todos os clientes).
			 *		  Implementação de tratamento para a situação em que a conexão com o BD é perdida p/ ten-
			 *		  tar fazer a reconexão automaticamente.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.23 - 28.03.2014 - por HHO
			 *		  Correção de bug identificado no Relatório Sintético de Movimentos em que somente um
			 *		  lançamento estava sendo contabilizado quando na realidade havia dois.
			 *		  Exemplo:
			 *			Compet        C/C        Plano de Conta              C/D    Efeito    Valor       Descrição
			 *			15/03/2014    28096-8    0300 - ALUGUEL TURMALINA    D      Válido    6.000,00    ref.03-01-15-1/2
			 *			30/03/2014    28096-8    0300 - ALUGUEL TURMALINA    D      Válido    6.000,00    ref.03-15-31-2/2
			 *		  O problema na consulta SQL estava na cláusula UNION, pois como os valores dos campos
			 *		  usados no SELECT eram iguais, os registros "duplicados" estavam sendo descartados inde-
			 *		  vidamente. A construção correta que foi implementada foi substituir a cláusula "UNION"
			 *		  por "UNION ALL".
			 *		  A correção foi realizada na rotina: FFluxoRelatorioMovimentoSintetico.montaSqlConsulta()
			 *		  Também foram corrigidas as seguintes rotinas por possuírem o mesmo problema:
			 *		      FFluxoRelatorioMovimentoRateioSintetico.montaSqlConsulta()
			 *		      FFluxoRelatorioCtaCorrente.calculaSaldoCtaCorrente()
			 * -----------------------------------------------------------------------------------------------
			 * v 1.23 - 14.04.2014 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Implementação de tratamento nas rotinas FBoletoConsulta.trataBotaoBoletoEmail() e
			 *		  FBoletoHtml.trataBotaoEnviaBoletoPorEmail() p/ prevenir o exception
			 *		  System.NullReferenceException ao verificar as condições:
			 *				wb.Document.Body.ScrollRectangle.Height < 600
			 *				!wb.Document.GetElementById("c_codigo_barras_loaded").OuterHtml.ToUpper().Contains("VALUE=S")
			 * -----------------------------------------------------------------------------------------------
			 * v 1.24 - 17.04.2014 - por HHO
			 *		  Alteração da carga do arquivo de retorno p/ detectar pagamentos com divergência de
			 *		  valor e registrar na tabela de ocorrências. A lógica implementada anteriormente não
			 *		  estava funcionando porque o banco atribuia como desconto o valor da diferença, fazendo
			 *		  c/ que o Financeiro interpretasse o valor pago como sendo exatamente o valor devido (no
			 *		  caso de pagamentos c/ valor inferior).
			 * -----------------------------------------------------------------------------------------------
			 * v 1.25 - 27.06.2014 - por HHO
			 *		  Alteração da carga do arquivo de retorno p/ gerar dados p/ o módulo responsável pela
			 *		  integração c/ o sistema de reciprocidade da Serasa.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.25 - 14.08.2014 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Alteração no painel de consulta do fluxo de caixa para adicionar um filtro que faça
			 *		  a restrição por clientes que estejam c/ o flag de negativado no SPC ativado.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.25 - 03.09.2014 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Correção da rotina FMain.reinicializaObjetosEstaticosUnitsDAO(), pois não estava
			 *		  chamando a rotina SerasaDAO.inicializaObjetosEstaticos(), o que causava erro ao
			 *		  tentar manipular as tabelas do módulo Serasa depois que uma reconexão automática ao BD
			 *		  tivesse ocorrido.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.26 - 06.10.2014 - por HHO
			 *		  Ajustes para tratar o novo meio de pagamento "Boleto AV" (código 6), ou seja,
			 *		  boleto à vista. Esta opção de pagamento tem o mesmo desconto concedido p/ dinheiro.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.26 - 03.12.2014 - por HHO
			 *		  Mantido o número da versão anterior.
			 *		  Ajuste na rotina de gravação das alterações de um lançamento de fluxo de caixa p/
			 *		  retirar a pergunta de confirmação.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.27 - 23.01.2015 - por HHO
			 *		  Ajuste na URL que gera o código de barras do boleto de 'bin' para 'e_x_e_c' devido a
			 *		  restrições de segurança no IIS 7.5 relativas ao uso do path 'bin'. Seria possível libe-
			 *		  rar essa restrição no novo servidor, mas optou-se em privilegiar a segurança.
			 *		  Essa modificação foi realizada dentro do escopo do processo de migração do servidor
			 *		  (Windows Server 2003 p/ Windows Server 2008 R2).
			 *		  Alteração nos dados de conexão com o BD, já que o SQL Server no novo servidor não
			 *		  está mais utilizando a porta padrão por questões de segurança.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.28 - 10.03.2016 - por HHO
			 *		  Alterações para tratar a nova estrutura de seleção do cedente ao cadastrar boletos.
			 *		  Ao invés do cedente ser definido através da loja do pedido, o sistema passa a definir
			 *		  a empresa responsável pela emissão da NFe já no cadastramento do pedido, sendo que
			 *		  cada emitente de NFe possui um campo indicando a empresa usada para a emissão
			 *		  de boletos.
			 *		  Importante lembrar que um pedido-filhote pode definir uma empresa emitente de NFe
			 *		  diferente da do pedido-pai, consequentemente, pode significar que os boletos podem
			 *		  ser emitidos por uma empresa diferente no pedido-filhote em relação ao pedido-pai.
			 *		  Esta premissa será totalmente válida a partir do momento em que a forma de
			 *		  pagamento for devidamente "rateada" entre o pedido-pai e o pedido-filhote durante
			 *		  o cadastramento do pedido ao implementar o "auto-split" de acordo com a disponibilidade
			 *		  dos produtos em cada CD.
			 *		  Alterações no processamento do arquivo de retorno ao gerar os dados do Serasa Recipro-
			 *		  cidade para tratar a possibilidade de haver mais do que uma empresa participante.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.29 - 16.05.2016 - por HHO
			 *		  Implementação de ajustes para alterar a cor de fundo dos painéis de acordo com o
			 *		  ambiente acessado.
			 *		  A cor inicialmente é obtida a partir do arquivo de configuração e, após realizar a
			 *		  conexão com o banco de dados, a cor é obtida através do campo 'cor_fundo_padrao' da
			 *		  tabela t_VERSAO. Caso a cor definida no banco de dados seja diferente da do arquivo,
			 *		  o parâmetro do arquivo é atualizado para respeitar a cor especificada no BD.
			 *		  Reversão do tratamento criado nas tabelas do Serasa Reciprocidade p/ permitir o
			 *		  cenário de várias empresas fazendo troca de dados com o Serasa. Devido à falta de
			 *		  tempo no cronograma e pela perspectiva de que somente a futura empresa (DIS Cobrança)
			 *		  irá fazer a troca de dados, optou-se por abortar esse ajuste, principalmente no
			 *		  lado do módulo Reciprocidade, que ainda estava pendente de ser ajustado.
			 *		  Ressaltando que a versão 1.28 não chegou a ser usada em produção.
			 *		  A única alteração mantida foi o campo t_FIN_BOLETO_CEDENTE.st_participante_serasa_reciprocidade
			 *		  para permitir que o módulo Financeiro possa determinar se deve ou não gerar os dados
			 *		  p/ o Serasa durante a carga do arquivo de retorno de boletos.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.29 - 16.12.2016 - por TRR
			 *		  Mantido o número da versão anterior.
			 *		  Ajustes na geração da imagem do boleto:
			 *			1) Incluir CNPJ/CPF do cliente.
			 *			2) Revisão do texto contido no email com orientações para impressão do boleto.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.30 - 10.01.2017 - por TRR
			 *		  Implementação do novo campo no fluxo de caixa 'dt_mes_competencia', cujo nome de
			 *		  exibição na tela é 'Comp2'. Este novo campo armazena somente o mês e ano, ou seja, o
			 *		  dia é sempre fixo em '01'. O seu uso destina-se ao controle contábil de despesas, razão
			 *		  pela qual somente os lançamentos de débito possuem o campo preenchido. Além disso,
			 *		  o mês/ano dos campos 'dt_competencia' e 'dt_mes_competencia' podem ser diferentes, já
			 *		  que o primeiro se refere à data em que o lançamento ocorreu no fluxo de caixa e o
			 *		  segundo se refere à data em que a despesa foi gerada.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.31 - 14.05.2017 - por HHO
			 *		  Ajustes para tratar a geração de boletos com base no cedente definido no cadastro da
			 *		  empresa emitente da NFe.
			 *		  Lembrando que:
			 *		      1) Cada pedido especifica obrigatoriamente uma empresa para a emissão da NFe.
			 *		      2) O cedente pode ser referenciado por mais de uma empresa emitente de NFe, pois
			 *		         as filiais podem ser configuradas para emitir boletos através da matriz.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.31(B) - 13.11.2017 - por HHO
			 *		  Ajustes para incluir filtro por data de alteração no painel de consulta de lançamentos.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.31(C) - 24.11.2017 - por HHO
			 *		  Implementação do Relatório Sintético Comparativo de Movimentos.
			 *		  Neste primeiro momento, o relatório está preparado para operar apenas com a opção de
			 *		  saída 'comparativo entre períodos'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.31(D) - 25.11.2017 - por HHO
			 *		  Relatório Sintético Comparativo de Movimentos: conclusão do desenvolvimento do relatório
			 *		  com a implementação da opção de saída 'comparativo mensal'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.32 - 04.03.2018 - por HHO
			 *		  Ajustes nas rotinas que pesquisam dados de CEP para acessar o novo banco de dados
			 *		  que contém dados atualizados, mas possui uma estrutura diferente.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.33 - 12.03.2018 - por HHO
			 *		  Ajustes no tratamento dos boletos em atraso para alterar a data de referência, passando
			 *		  a usar o campo 'data da gravação do arquivo' ao invés da 'data do crédito'. Lembrando
			 *		  que esses campos são informados no arquivo de retorno.
			 *		  Esta alteração visa solucionar a anomalia de quantidade negativa de dias em atraso.
			 *		  Essa anomalia passou a ocorrer quando o banco passou a informar uma data de crédito
			 *		  futura no arquivo de retorno.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.33(B) - 12.01.2019 - por HHO
			 *		  Ajustes no painel de edição de lançamentos em lote para incluir os campos conta
			 *		  corrente, empresa e plano de conta.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.34 - 16.07.2019 - por HHO
			 *		  Implementação do campo 'numero_NF' no cadastro de lançamentos do fluxo de caixa, com
			 *		  ajustes nos painéis de cadastramento, consulta e edição.
			 *		  Implementação de tratamento para o novo meio de pagamento 'cartão (maquineta)'.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.34(B) - 22.07.2019 - por HHO
			 *		  Implementação de ajustes relacionados ao campo 'numero_NF' do fluxo de caixa: exibição
			 *		  do campo em nova coluna no resultado da consulta dos lançamentos na tela, impressão e
			 *		  na planilha Excel.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.35 - 11.11.2019 - por LHGX
			 *		  Reestruturação da tabela t_NFE_EMITENTE, prevendo a existência da nova tabela
			 *		  t_NFE_EMITENTE_NUMERACAO.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.36 - 14.11.2019 - por LHGX
			 *		  Correção de bug (JOIN com t_NFE_EMITENTE_NUMERACAO)
			 * -----------------------------------------------------------------------------------------------
			 * v 1.37 - 28.08.2020 - por HHO
			 *		  Ajustes para tratar a memorização do endereço de cobrança no pedido, pois, a partir de
			 *		  agora, ao invés de obter os dados do endereço no cadastro do cliente (t_CLIENTE), deve-se
			 *		  usar os dados que estão gravados no próprio pedido. O tratamento que já ocorria com o
			 *		  endereço de entrega deve passar a ser feito p/ o endereço de cobrança/cadastro.
			 *		  Implementação de tratamento na carga do arquivo de retorno de boletos para ignorar o
			 *		  envio de boleto AV para o Serasa.
			 * -----------------------------------------------------------------------------------------------
			 * v 1.38 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.39 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.40 - XX.XX.20XX - por XXX
			 *		  
			 * -----------------------------------------------------------------------------------------------
			 * v 1.XX - XX.XX.20XX - por XXX
			 *		  Implementação de recursos para cobrança de clientes em atraso.
			 * ===============================================================================================
			 */
            #endregion

            #region [ Etc ]
            public class Etc
            {
                public const String SIMBOLO_MONETARIO = "R$";
                public const byte FLAG_NAO_SETADO = 255;
                public const int TAM_MIN_LOJA = 2;
                public const int TAM_MIN_NUM_PEDIDO = 6;    // SOMENTE PARTE NUMÉRICA DO NÚMERO DO PEDIDO
                public const int TAM_MIN_ID_PEDIDO = 7; // PARTE NUMÉRICA DO NÚMERO DO PEDIDO + LETRA REFERENTE AO ANO
                public const char COD_SEPARADOR_FILHOTE = '-';
                public const int MAX_TAM_BOLETO_CAMPO_ENDERECO = 40;
                public const int MAX_TAM_BOLETO_CAMPO_NOME_SACADO = 40;
                public const String ID_PF = "PF";
                public const String ID_PJ = "PJ";
                public const int TAMANHO_CPF = 11;
                public const int TAMANHO_CNPJ = 14;
                public const int TAMANHO_RAIZ_CNPJ = 8;
                public const String PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE = "TFBI";
                public const String SQL_COLLATE_CASE_ACCENT = " COLLATE Latin1_General_CI_AI";
            }
            #endregion

            #region [ Log ]
            public class LogAtividade
            {
                public static string PathLogAtividade = Application.StartupPath + "\\LOG_ATIVIDADE";
                public const int CorteArqLogEmDias = 365;
                public const string ExtensaoArqLog = "LOG";
            }
            #endregion

            #region [ Imagens ]
            public class Imagens
            {
                public static string PathImagens = Application.StartupPath + "\\Imagens";
                public const string ArqLogoBradesco = "drd_pb_g.gif";
            }
            #endregion

            #region[ Data/Hora ]
            public class DataHora
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
                public const string FmtMin = "mm";
                public const string FmtSeg = "ss";
                public const string FmtMiliSeg = "fff";
                public const string FmtYYYYMMDD = FmtAno + FmtMes + FmtDia;
                public const string FmtHHMMSS = FmtHora + FmtMin + FmtSeg;
                public const string FmtDdMmYyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAnoCom2Digitos;
                public const string FmtDdMmYyyyComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno;
                public const string FmtDdMmYyyyHhMmComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin;
                public const string FmtDdMmYyyyHhMmSsComSeparador = FmtDia + "/" + FmtMes + "/" + FmtAno + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
                public const string FmtYyyyMmDdComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia;
                public const string FmtYyyyMmDdHhMmSsComSeparador = FmtAno + "-" + FmtMes + "-" + FmtDia + " " + FmtHora + ":" + FmtMin + ":" + FmtSeg;
            }
            #endregion

            #region [ Classe FIN ]
            public class FIN
            {
                public const String ID_USUARIO_SISTEMA = "SISTEMA";
                public static bool FLAG_PERIODO_OBRIGATORIO_FILTRO_CONSULTA = false;
                public const String PREFIXO_NUMERO_DOCUMENTO_BOLETO_AVULSO = "A";
                public const String NOME_ARQ_CACHE_LISTA_NOME_CLIENTE_AUTO_COMPLETE = "FIN-CacheListaNomeClienteAutoComplete.xml";

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

                #region [ Natureza ]
                public class Natureza
                {
                    public const char CREDITO = 'C';
                    public const char DEBITO = 'D';
                }
                #endregion

                #region [ CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado ]
                public class CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado
                {
                    public const byte APENAS_ATRASADOS = 1;
                    public const byte IGNORAR_ATRASADOS = 2;
                }
                #endregion

                #region [ CodOpcaoCobrancaAdmSituacao ]
                public class CodOpcaoCobrancaAdmSituacao
                {
                    public const byte EM_ATRASO_NAO_ALOCADO = 1;
                    public const byte EM_ATRASO_JA_ALOCADO = 2;
                }
                #endregion

                #region [ StAtivo ]
                public class StAtivo
                {
                    public const byte INATIVO = 0;
                    public const byte ATIVO = 1;
                }
                #endregion

                #region [ StSistema ]
                public class StSistema
                {
                    public const byte NAO = 0;
                    public const byte SIM = 1;
                }
                #endregion

                #region [ StCampoFlag ]
                public class StCampoFlag
                {
                    public const byte FLAG_DESLIGADO = 0;
                    public const byte FLAG_LIGADO = 1;
                }
                #endregion

                #region [ StSemEfeito ]
                public class StSemEfeito
                {
                    public const byte FLAG_DESLIGADO = 0;
                    public const byte FLAG_LIGADO = 1;
                }
                #endregion

                #region [ StConfirmacaoPendente ]
                public class StConfirmacaoPendente
                {
                    public const byte FLAG_DESLIGADO = 0;
                    public const byte FLAG_LIGADO = 1;
                }
                #endregion

                #region [ ST_T_FIN_NF_PARCELA_PAGTO ]
                public class ST_T_FIN_NF_PARCELA_PAGTO
                {
                    public const byte INICIAL = 0;
                    public const byte CANCELADO = 1;
                    public const byte TRATADO = 2;
                }
                #endregion

                #region [ ST_T_FIN_PEDIDO_HIST_PAGTO ]
                public class ST_T_FIN_PEDIDO_HIST_PAGTO
                {
                    public const byte PREVISAO = 1;
                    public const byte QUITADO = 2;
                    public const byte CANCELADO = 3;
                }
                #endregion

                #region [ T_PEDIDO__BOLETO_CONFECCIONADO_STATUS ]
                public class T_PEDIDO__BOLETO_CONFECCIONADO_STATUS
                {
                    public const byte NAO = 0;
                    public const byte SIM = 1;
                    public const byte NAO_DEFINIDO = 10;
                }
                #endregion

                #region [ T_PEDIDO__GARANTIA_INDICADOR_STATUS ]
                public class T_PEDIDO__GARANTIA_INDICADOR_STATUS
                {
                    public const byte NAO = 0;
                    public const byte SIM = 1;
                    public const byte NAO_DEFINIDO = 10;
                }
                #endregion

                #region [ T_PEDIDO__ANALISE_CREDITO_STATUS ]
                public class T_PEDIDO__ANALISE_CREDITO_STATUS
                {
                    public const int ST_INICIAL = 0;
                    public const int CREDITO_PENDENTE = 1;
                    public const int CREDITO_OK = 2;
                    public const int PENDENTE_VENDAS = 8;
                    public const int CREDITO_OK_AGUARDANDO_DEPOSITO = 9;
                    public const int NAO_ANALISADO = 10; // PEDIDOS ANTIGOS QUE JÁ ESTAVAM NA BASE
                }
                #endregion

                #region [ CtrlPagtoModulo ]
                public class CtrlPagtoModulo
                {
                    public const byte BOLETO = 1;
                    public const byte CHEQUE = 2;
                    public const byte VISA = 3;
                    public const byte BRASPAG_CARTAO = 4;
                }
                #endregion

                #region [ CtrlPagtoStatus ]
                public enum eCtrlPagtoStatus
                {
                    // IMPORTANTE: NUNCA usar o valor reservado FLAG_NAO_SETADO = 255
                    CONTROLE_MANUAL = 0,
                    CADASTRADO_INICIAL = 1,
                    BOLETO_BAIXADO = 3,
                    CHEQUE_DEVOLVIDO = 4,
                    VISA_CANCELADO = 5,
                    BOLETO_PAGO_CHEQUE_VINCULADO = 6,
                    BOLETO_COM_PAGAMENTO_CANCELADO = 7,
                    PAGO = 10
                }
                #endregion

                #region [ FormaPagto ]
                public class FormaPagto
                {
                    public const byte ID_FORMA_PAGTO_DINHEIRO = 1;
                    public const byte ID_FORMA_PAGTO_DEPOSITO = 2;
                    public const byte ID_FORMA_PAGTO_CHEQUE = 3;
                    public const byte ID_FORMA_PAGTO_BOLETO = 4;
                    public const byte ID_FORMA_PAGTO_CARTAO = 5;
                    public const byte ID_FORMA_PAGTO_BOLETO_AV = 6;
					public const byte ID_FORMA_PAGTO_CARTAO_MAQUINETA = 7;
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

                #region [ LogOperacao - Códigos de operação para o log ]
                public class LogOperacao
                {
                    // Texto com 12 posições
                    public const String LOGON = "Logon";
                    public const String LOGOFF = "Logoff";
                    public const String RECONEXAO_BD = "Reconexao-BD";
                    public const String FLUXO_CAIXA_CREDITO_INSERE = "FluxoCredIns";
                    public const String FLUXO_CAIXA_CREDITO_LOTE_INSERE = "FluxCrdLtIns";
                    public const String FLUXO_CAIXA_DEBITO_INSERE = "FluxoDebIns";
                    public const String FLUXO_CAIXA_DEBITO_LOTE_INSERE = "FluxDebLtIns";
                    public const String FLUXO_CAIXA_EDITA = "FluxoEdita";
                    public const String FLUXO_CAIXA_EDITA_LOTE = "FluxoEditLot";
                    public const String FLUXO_CAIXA_EXCLUI = "FluxoExclui";
                    public const String BOLETO_PRE_CADASTRADO_ANULA = "BolPreAnula";
                    public const String BOLETO_CADASTRA = "BolCadastra";
                    public const String BOLETO_AVULSO_CADASTRA = "BolAvulsoCad";
                    public const String BOLETO_CANCELA_MANUAL = "BolCancelMan";
                    public const String BOLETO_DESFEITO = "BolDesfeito";
                    public const String BOLETO_GERA_ARQ_REMESSA = "BolGeraArqRm";
                    public const String BOLETO_GERA_INDICE_DIARIO_ARQ_REMESSA = "BolIdxArqRm";
                    public const String BOLETO_GERA_NSU_ARQ_REMESSA = "BolNsuArqRm";
                    public const String BOLETO_CARREGA_ARQ_RETORNO = "BolCarArqRet";
                    public const String BOLETO_OCORRENCIAS_TRATA_CEP_IRREGULAR = "BolOcTrCepIr";
                    public const String BOLETO_OCORRENCIAS_TRATA_VALA_COMUM = "BolOcTrValaC";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_02 = "FCEdBolOc02";
                    public const String FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_OCORRENCIA_02 = "FCInsBolOc02";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_06 = "FCEdBolOc06";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_09 = "FCEdBolOc09";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_10 = "FCEdBolOc10";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_12 = "FCEdBolOc12";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_13 = "FCEdBolOc13";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_14 = "FCEdBolOc14";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_15 = "FCEdBolOc15";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_16 = "FCEdBolOc16";
                    public const String FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_OCORRENCIA_17 = "FCInsBolOc17";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_22 = "FCEdBolOc22";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_23 = "FCEdBolOc23";
                    public const String FLUXO_CAIXA_EDITA_DEVIDO_BOLETO_OCORRENCIA_34 = "FCEdBolOc34";
                    public const String SERASA_CLIENTE_INSERE = "SerCliIns";
                    public const String SERASA_TITULO_MOVIMENTO_INSERE = "SerTitMovIns";
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

                #region [ CodBoletoTipoVinculo ]
                public class CodBoletoTipoVinculo
                {
                    public const byte BOLETO_COM_PEDIDO_EMISSAO_AUTOMATICA = 1;
                    public const byte BOLETO_AVULSO_SEM_PEDIDO = 2;
                    public const byte BOLETO_AVULSO_COM_PEDIDO = 3;
                }
                #endregion

                #region [ CodBoletoStatus ]
                public class CodBoletoStatus
                {
                    public const short INICIAL = 0;
                    public const short CANCELADO_MANUAL = 1;
                    public const short ENVIADO_REMESSA_BANCO = 3;
                }
                #endregion

                #region [ CodBoletoItemStatus ]
                public class CodBoletoItemStatus
                {
                    public const short INICIAL = 0;
                    public const short CANCELADO_MANUAL = 1;
                    public const short BOLETO_BAIXADO = 2;
                    public const short ENVIADO_REMESSA_BANCO = 3;
                    public const short ENTRADA_CONFIRMADA = 4;
                    public const short BOLETO_PAGO = 5;
                    public const short BOLETO_REJEITADO_CEP_IRREGULAR = 6;
                    public const short VALA_COMUM = 7;
                    public const short BOLETO_PAGO_EM_CHEQUE = 8;
                    public const short BOLETO_PAGO_EM_OCORRENCIA_17 = 9;
                    public const short BOLETO_PAGO_EM_OCORRENCIA_15 = 10;
                    public const short BOLETO_COM_PAGAMENTO_CANCELADO = 11;
                }
                #endregion

                #region [ CodBoletoItemTipoVencto ]
                public class CodBoletoItemTipoVencto
                {
                    public const byte VENCTO_DEFINIDO = 1;
                    public const byte VENCTO_A_VISTA = 2;
                    public const byte CONTRA_APRESENTACAO = 3;
                    public const byte VER_INSTRUCOES = 4;
                    public const byte ALTERADO_PARA_A_VISTA = 5;
                }
                #endregion

                #region [ BoletoBradesco ]
                public class BoletoBradesco
                {
                    #region [ CodTipoSacado ]
                    public class CodTipoSacado
                    {
                        public const String CPF = "01";
                        public const String CNPJ = "02";
                        public const String PIS_PASEP = "03";
                        public const String NAO_TEM = "98";
                        public const String OUTROS = "99";
                    }
                    #endregion
                }
                #endregion

                #region [ CodBoletoArqRemessaStGeracao ]
                public class CodBoletoArqRemessaStGeracao
                {
                    public const short SUCESSO = 1;
                    public const short FALHA = 2;
                }
                #endregion

                #region [ CodBoletoArqRetornoStProcessamento ]
                public class CodBoletoArqRetornoStProcessamento
                {
                    public const short EM_PROCESSAMENTO = 1;
                    public const short SUCESSO = 2;
                    public const short FALHA = 3;
                }
                #endregion

                #region [ CodBoletoOcorrenciaStOcorrenciaTratada ]
                public class CodBoletoOcorrenciaStOcorrenciaTratada
                {
                    public const byte NAO_TRATADA = 0;
                    public const byte JA_TRATADA = 1;
                }
                #endregion

                #region [ NSU ]
                public class NSU
                {
                    public const String T_FIN_FLUXO_CAIXA = "t_FIN_FLUXO_CAIXA";
                    public const String T_FIN_NF_PARCELA_PAGTO = "t_FIN_NF_PARCELA_PAGTO";
                    public const String T_FIN_BOLETO = "t_FIN_BOLETO";
                    public const String T_FIN_BOLETO_ITEM = "t_FIN_BOLETO_ITEM";
                    public const String T_FIN_BOLETO_ARQ_REMESSA = "t_FIN_BOLETO_ARQ_REMESSA";
                    public const String T_FIN_BOLETO_ARQ_RETORNO = "t_FIN_BOLETO_ARQ_RETORNO";
                    public const String T_FIN_PEDIDO_HIST_PAGTO = "t_FIN_PEDIDO_HIST_PAGTO";
                    public const String T_FIN_BOLETO_MOVIMENTO = "t_FIN_BOLETO_MOVIMENTO";
                    public const String T_FIN_BOLETO_OCORRENCIA = "t_FIN_BOLETO_OCORRENCIA";
                    public const String NSU_BOLETO_AVULSO_NUMERO_DOCUMENTO = "NSU_BOLETO_AVULSO_NUMERO_DOCUMENTO";
                    public const String T_SERASA_CLIENTE = "t_SERASA_CLIENTE";
                    public const String T_SERASA_TITULO_MOVIMENTO = "t_SERASA_TITULO_MOVIMENTO";
                }
                #endregion

                #region [ ID_T_PARAMETRO ]
                public static class ID_T_PARAMETRO
                {
                    public const string SERASA_RECIPROCIDADE_CNPJ_IGNORADOS = "SerasaReciprocidadeCnpjIgnorados";
                    public const string ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS = "Flag_Pedido_MemorizacaoCompletaEnderecos";
                }
                #endregion
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

            #region [ Status de Pagamento do Pedido ]
            public class StPagtoPedido
            {
                public const String ST_PAGTO_PAGO = "S";
                public const String ST_PAGTO_NAO_PAGO = "N";
                public const String ST_PAGTO_PARCIAL = "P";
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

            #region [ Marketplace ]
            public class Marketplace
            {
                public const byte COD_PLANILHA_PAGAMENTO_B2W = 2;
                public const byte COD_PLANILHA_PAGAMENTO_CNOVA = 3;
                public const byte COD_PLANILHA_PAGAMENTO_WALMART = 4;
            } 
            #endregion
        }
        #endregion

        #region [ Atributos ]
        public static DateTime dtHrInicioRefRelogioServidor;
        public static DateTime dtHrInicioRefRelogioLocal;
        public static int contadorLancamentoDebitoInserido = 0;
        public static int contadorLancamentoDebitoLoteInserido = 0;
        public static int contadorLancamentoCreditoInserido = 0;
        public static int contadorLancamentoCreditoLoteInserido = 0;
        public static string PATH_BOLETO_ARQUIVO_REMESSA = Application.StartupPath + "\\BOLETOS\\ARQUIVO_REMESSA";
        public static Color BackColorPainelPadrao = SystemColors.Control;
        #endregion

        #region [ Classe Acesso ]
        public class Acesso
        {
            #region [ Constantes ]
            public const String OP_CEN_FIN_APP_FINANCEIRO_ACESSO_AO_MODULO = "21400";
            public const String OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_CONSULTAR = "21500";
            public const String OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_DEBITO = "21600";
            public const String OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_LANCTO_CREDITO = "21700";
            public const String OP_CEN_FIN_APP_FINANCEIRO_FLUXO_CAIXA_EDITAR_LANCTO = "21800";
            public const String OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR = "22000";
            public const String OP_CEN_FIN_APP_COBRANCA_ACESSO_AO_MODULO = "22200";
            public const String OP_CEN_FIN_APP_COBRANCA_ADMINISTRACAO_CARTEIRA_EM_ATRASO = "22300";
            public const String OP_CEN_FIN_APP_COBRANCA_COBRANCA_CARTEIRA_EM_ATRASO = "22400";
            public const String OP_CEN_ACESSO_TODAS_LOJAS = "10100";
            #endregion

            #region [ Atributos ]
            public static List<String> listaOperacoesPermitidas = new List<String>();
            #endregion

            #region [ Métodos ]

            #region [ operacaoPermitida ]
            /// <summary>
            /// Indica se a operação especificada no parâmetro consta na lista de operações permitidas do usuário
            /// </summary>
            /// <param name="idOperacao">
            /// Operação a ser pesquisada na lista de operações permitidas
            /// </param>
            /// <returns>
            /// true: a operação pesquisada consta na lista de operações permitidas
            /// false: a operação pesquisada não consta na lista de operações permitidas
            /// </returns>
            public static bool operacaoPermitida(String idOperacao)
            {
                if (idOperacao == null) return false;
                if (idOperacao.Trim().Length == 0) return false;

                for (int i = 0; i < listaOperacoesPermitidas.Count; i++)
                {
                    if (listaOperacoesPermitidas[i].ToString().Equals(idOperacao)) return true;
                }
                // Operação não consta da lista de operações permitidas
                return false;
            }
            #endregion

            #endregion
        }
        #endregion

        #region [ Classe Usuario ]
        public class Usuario
        {
            #region [ Atributos ]
            public static String usuario = "";
            public static String senhaDigitada = "";
            public static String senhaCriptografada = "";
            public static String senhaDescriptografada = "";
            public static String nome = "";
            public static bool cadastrado = false;
            public static bool bloqueado = false;
            public static bool senhaExpirada = false;
            public static String fin_email_remetente;
            public static String fin_display_name_remetente;
            public static String fin_servidor_smtp_endereco;
            public static int fin_servidor_smtp_porta;
            public static String fin_usuario_smtp;
            public static String fin_senha_smtp;
            #endregion

            #region [ Defaults ]
            public class Defaults
            {
                public static byte contaCorrente = 0;
                public static byte planoContasEmpresa = 0;
                public static int planoContasContaCredito = 0;
                public static int planoContasContaDebito = 0;
                public static String pathBoletoArquivoRetornoRelatorios = "";
                public static String pathBoletoArquivoRemessaRelatorio = "";
                public static String relatorioArqRetornoTipoSaida = "";
                public static String relatorioArqRemessaTipoSaida = "";
                public static String relatorioMovimentoChkIncluirAtrasados = "";
                public static String dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete = "";
                public static int fluxoCreditoLoteQtdeLancamentos = 0;
                public static int fluxoDebitoLoteQtdeLancamentos = 0;
                public class FBoletoArqRemessa
                {
                    public static String pathBoletoArquivoRemessa = "";
                    public static int boletoCedente = 0;
                }
                public class FBoletoArqRetorno
                {
                    public static String pathBoletoArquivoRetorno = "";
                }
            }
            #endregion
        }
        #endregion

        #region [ RegistryApp ]
        public class RegistryApp
        {
            public const string REGISTRY_BASE_PATH = "Software\\" + Cte.Aplicativo.NOME_OWNER + "\\" + Cte.Aplicativo.NOME_SISTEMA;

            #region [ Chaves ]
            public class Chaves
            {
                public static String left = "Left";
                public static String top = "Top";
                public static String usuario = "Usuario";
                public static String contaCorrente = "contaCorrente";
                public static String planoContasEmpresa = "planoContasEmpresa";
                public static String planoContasContaCredito = "planoContasContaCredito";
                public static String planoContasContaDebito = "planoContasContaDebito";
                public static String pathBoletoArquivoRetornoRelatorios = "pathBoletoArquivoRetornoRelatorios";
                public static String pathBoletoArquivoRemessaRelatorio = "pathBoletoArquivoRemessaRelatorio";
                public static String relatorioArqRetornoTipoSaida = "relatorioArqRetornoTipoSaida";
                public static String relatorioArqRemessaTipoSaida = "relatorioArqRemessaTipoSaida";
                public static String relatorioMovimentoChkIncluirAtrasados = "relatorioMovimentoChkIncluirAtrasados";
                public static String dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete = "dataYyyMmDdUltArqCacheListaNomeClienteAutoComplete";
                public static String fluxoCreditoLoteQtdeLancamentos = "fluxoCreditoLoteQtdeLancamentos";
                public static String fluxoDebitoLoteQtdeLancamentos = "fluxoDebitoLoteQtdeLancamentos";
                public class FBoletoArqRemessa
                {
                    public static String pathBoletoArquivoRemessa = "pathBoletoArquivoRemessa";
                    public static String boletoCedente = "FBoletoArqRemessa_BoletoCedente";
                }
                public class FBoletoArqRetorno
                {
                    public static String pathBoletoArquivoRetorno = "pathBoletoArquivoRetorno";
                }
            }
            #endregion

            #region [ Métodos ]

            #region [ criaRegistryKey ]
            public static RegistryKey criaRegistryKey(String subKey)
            {
                RegistryKey regKey = Registry.CurrentUser;
                regKey = regKey.CreateSubKey(subKey);
                return regKey;
            }
            #endregion

            #endregion
        }
        #endregion

        #region[ ReaderWriterLock ]
        public static ReaderWriterLock rwlArqLogAtividade = new ReaderWriterLock();
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

        #region[ barraInvertidaDel ]
        public static string barraInvertidaDel(string path)
        {
            if (path == null) return "";
            string strResp = path.TrimEnd();
            while (true)
            {
                if (strResp.Length == 0) return "";
                if (strResp[strResp.Length - 1] != (char)92) return strResp;
                strResp = strResp.Substring(0, strResp.Length - 1).TrimEnd();
            }
        }
        #endregion

        #region [ calculaBoletoItemFlagInstrucaoProtesto ]
        /// <summary>
        /// No caso de uma série de boletos ser gerada com instrução de protesto, esta rotina calcula
        /// se a parcela do boleto indicada deve ser gerada com instrução de protesto ou não.
        /// Devido aos custos dos cartórios, foram definidas algumas regras de quais parcelas de uma série
        /// de boletos serão geradas c/ a instrução de protesto e quais não.
        /// </summary>
        /// <param name="numeroParcela">Número da parcela do boleto</param>
        /// <param name="totalParcelas">Quantidade total de parcelas da série de boletos</param>
        /// <returns>
        /// 0: sem instrução de protesto (valor que será usado p/ gravar no banco de dados)
        /// 1: com instrução de protesto (valor que será usado p/ gravar no banco de dados)
        /// </returns>
        public static byte calculaBoletoItemFlagInstrucaoProtesto(byte numeroParcela, byte totalParcelas)
        {
            #region [ Regra ]
            // Regra:
            //	1 boleto: protesta
            //	2 boletos: protesta o segundo
            //	3 boletos: protesta o segundo e o terceiro
            //	4 boletos: protesta o segundo e o quarto
            //	5 boletos: protesta o segundo e o quinto
            //	e assim por diante
            #endregion

            // A última parcela sempre tem instrução de protesto
            // Se houver apenas 1 parcela, deverá ter instrução de protesto
            if (numeroParcela == totalParcelas) return 1;

            // A segunda parcela sempre tem instrução de protesto
            if (numeroParcela == 2) return 1;

            return 0;
        }
        #endregion

        #region [ calculaDigitoVerificadorCodigoBarrasBradesco ]
        /// <summary>
        /// Calcula o dígito verificador para o código de barras.
        /// O cálculo é feito através do módulo 11, com base de cálculo igual a 9.
        /// </summary>
        /// <param name="numero">
        /// Números que compõem o código de barras.
        /// </param>
        /// <returns>
        /// Retorna o dígito verificador para o código de barras.
        /// </returns>
        public static String calculaDigitoVerificadorCodigoBarrasBradesco(String numero)
        {
            #region [ Declarações ]
            const int baseCalculo = 9;
            int intFator = 2;
            int intSoma = 0;
            int intNumeroAux;
            int intDV;
            String strNumero;
            #endregion

            if (numero == null) return "";
            if (numero.Trim().Length == 0) return "";

            strNumero = digitos(numero);

            for (int i = (strNumero.Length - 1); i >= 0; i--)
            {
                intNumeroAux = (int)converteInteiro(strNumero[i].ToString());
                intSoma += intNumeroAux * intFator;
                if (intFator == baseCalculo) intFator = 2; else intFator++;
            }

            intDV = 11 - (intSoma % 11);
            if ((intDV == 0) || (intDV == 1) || (intDV > 9)) intDV = 1;
            return intDV.ToString();
        }
        #endregion

        #region [ calculaDigitoVerificadorLinhaDigitavelBradesco ]
        /// <summary>
        /// Calcula o dígito verificador para ser utilizado na linha digitável.
        /// O cálculo é feito através do módulo 10.
        /// </summary>
        /// <param name="numero">
        /// Números para os quais se deseja calcular o dígito verificador.
        /// </param>
        /// <returns>
        /// Retorna o dígito verificador para ser utilizado na linha digitável.
        /// </returns>
        public static String calculaDigitoVerificadorLinhaDigitavelBradesco(String numero)
        {
            #region [ Declarações ]
            int intSoma = 0;
            int intParImpar = 2;
            int intNumAux1;
            int intNumAux2;
            int intDV;
            String strNumero;
            String strNumeroAux;
            #endregion

            if (numero == null) return "";
            if (numero.Trim().Length == 0) return "";

            strNumero = digitos(numero);
            for (int i = (strNumero.Length - 1); i >= 0; i--)
            {
                intNumAux1 = (int)converteInteiro(strNumero[i].ToString());
                intNumAux2 = intNumAux1 * (2 - (intParImpar % 2));
                if (intNumAux2 >= 10)
                {
                    strNumeroAux = intNumAux2.ToString();
                    intNumAux2 = (int)converteInteiro(strNumeroAux[0].ToString()) + (int)converteInteiro(strNumeroAux[1].ToString());
                }
                intSoma += intNumAux2;
                intParImpar++;
            }

            if ((intSoma % 10) == 0)
                intDV = 0;
            else
                intDV = 10 - (intSoma % 10);

            return intDV.ToString();
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
        public static int calculaTimeSpanMiliSegundos(TimeSpan ts)
        {
            return ts.Milliseconds + 1000 * (ts.Seconds + (60 * (ts.Minutes + (60 * (ts.Hours + (24 * ts.Days))))));
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
        public static int calculaTimeSpanSegundos(TimeSpan ts)
        {
            return ts.Seconds + (60 * (ts.Minutes + (60 * (ts.Hours + (24 * ts.Days)))));
        }
        #endregion

        #region [ converteColorFromHtml ]
        public static Color? converteColorFromHtml(string htmlColor)
        {
            #region [ Declarações ]
            Color cor;
            #endregion

            if (htmlColor == null) return null;
            if (htmlColor.Trim().Length == 0) return null;

            try
            {
                htmlColor = htmlColor.Trim();
                if (!htmlColor.StartsWith("#")) htmlColor = "#" + htmlColor;
                cor = ColorTranslator.FromHtml(htmlColor);
                return cor;
            }
            catch (Exception)
            {
                return null;
            }

        }
        #endregion

        #region[ converteDdMmYyParaDateTime ]
        /// <summary>
        /// Converte um texto no formato DDMMYY (ano c/ 2 dígitos) com ou sem separadores para o tipo DateTime.
        /// O pivotamento do ano é feito com base de ano 80.
        /// </summary>
        /// <param name="strDdMmYy">Texto representando uma data no formato DDMMYY (ano com 2 dígitos) com ou sem separadores</param>
        /// <returns>
        /// Retorna a data representada no tipo DateTime
        /// </returns>
        public static DateTime converteDdMmYyParaDateTime(string strDdMmYy)
        {
            DateTime dtDataHoraResp;
            CultureInfo myCultureInfo = new CultureInfo("pt-BR");
            string strDdMmYyyy;
            String strDdMm;
            String strYyyy;
            string strFormato;

            strDdMm = Texto.leftStr(digitos(strDdMmYy), 4);

            strYyyy = Texto.rightStr(digitos(strDdMmYy), 2);
            if (converteInteiro(strYyyy) >= 80) strYyyy = "19" + strYyyy; else strYyyy = "20" + strYyyy;

            strDdMmYyyy = strDdMm + strYyyy;

            strFormato = Cte.DataHora.FmtDia +
                         Cte.DataHora.FmtMes +
                         Cte.DataHora.FmtAno;
            if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
            return DateTime.MinValue;
        }
        #endregion

        #region[ converteDdMmYyyyParaDateTime ]
        public static DateTime converteDdMmYyyyParaDateTime(string strDdMmYyyy)
        {
            string strFormato;
            DateTime dtDataHoraResp;
            CultureInfo myCultureInfo = new CultureInfo("pt-BR");
            strFormato = Cte.DataHora.FmtDia +
                         Cte.DataHora.FmtMes +
                         Cte.DataHora.FmtAno;
            if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
            return DateTime.MinValue;
        }
        #endregion

        #region[ converteYyyyMmDdParaDateTime ]
        public static DateTime converteYyyyMmDdParaDateTime(string strYyyyMmDd)
        {
            string strYyyyMmDdAux;
            string strDdMmYyyy;
            string strFormato;
            DateTime dtDataHoraResp;
            CultureInfo myCultureInfo = new CultureInfo("pt-BR");
            strYyyyMmDdAux = digitos(strYyyyMmDd);
            if (strYyyyMmDdAux.Length == 0) return DateTime.MinValue;
            strDdMmYyyy = strYyyyMmDdAux.Substring(6, 2) + strYyyyMmDdAux.Substring(4, 2) + strYyyyMmDdAux.Substring(0, 4);
            strFormato = Cte.DataHora.FmtDia +
                         Cte.DataHora.FmtMes +
                         Cte.DataHora.FmtAno;
            if (DateTime.TryParseExact(digitos(strDdMmYyyy), strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp)) return dtDataHoraResp;
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

        #region [ decodificaCampoMonetario ]
        /// <summary>
        /// Converte para número decimal um campo monetário informado sem formatação e considerando 
        /// que os 2 últimos dígitos são referentes aos centavos.
        /// O sinal de negativo também é aceito e considerado na conversão.
        /// </summary>
        /// <param name="valor">
        /// Campo monetário a ser convertido.
        /// </param>
        /// <returns>
        /// Retorna o valor monetário convertido para número decimal.
        /// </returns>
        public static decimal decodificaCampoMonetario(String valor)
        {
            #region [ Declarações ]
            String strSinal = "";
            String strCentavos;
            String strValorInteiro;
            #endregion

            #region [ Consistência ]
            if (valor == null) return 0m;
            if (valor.Trim().Length == 0) return 0m;
            #endregion

            if (valor.IndexOf('-') != -1)
            {
                strSinal = "-";
                valor = valor.Replace("-", "");
            }
            valor = digitos(valor);
            valor = valor.PadLeft(3, '0');
            strCentavos = Texto.rightStr(valor, 2);
            strValorInteiro = Texto.leftStr(valor, valor.Length - 2);
            // Retira zeros à esquerda da parte inteira
            while (strValorInteiro.Length > 0)
            {
                if (strValorInteiro[0] == '0')
                    strValorInteiro = strValorInteiro.Substring(1);
                else
                    break;
            }
            if (strValorInteiro.Length == 0) strValorInteiro = "0";

            return Decimal.Parse(strSinal + strValorInteiro + strCentavos) / 100m;
        }
        #endregion

        #region [ decodificaIdentificacaoOcorrencia ]
        public static String decodificaIdentificacaoOcorrencia(String identificacaoOcorrencia)
        {
            #region [ Declarações ]
            String strResp = "";
            #endregion

            #region [ Consistência ]
            if (identificacaoOcorrencia == null) return "";
            if (identificacaoOcorrencia.Trim().Length == 0) return "";
            #endregion

            if (identificacaoOcorrencia.Equals("02"))
                strResp = "Entrada Confirmada";
            else if (identificacaoOcorrencia.Equals("03"))
                strResp = "Entrada Rejeitada";
            else if (identificacaoOcorrencia.Equals("06"))
                strResp = "Liquidação Normal";
            else if (identificacaoOcorrencia.Equals("09"))
                strResp = "Baixado Automat. via Arquivo";
            else if (identificacaoOcorrencia.Equals("10"))
                strResp = "Baixado conforme instruções da Agência";
            else if (identificacaoOcorrencia.Equals("11"))
                strResp = "Em Ser - Arquivo de Títulos pendentes";
            else if (identificacaoOcorrencia.Equals("12"))
                strResp = "Abatimento Concedido";
            else if (identificacaoOcorrencia.Equals("13"))
                strResp = "Abatimento Cancelado";
            else if (identificacaoOcorrencia.Equals("14"))
                strResp = "Vencimento Alterado";
            else if (identificacaoOcorrencia.Equals("15"))
                strResp = "Liquidação em Cartório";
            else if (identificacaoOcorrencia.Equals("16"))
                strResp = "Título Pago em Cheque – Vinculado";
            else if (identificacaoOcorrencia.Equals("17"))
                strResp = "Liquidação após baixa ou Título não registrado";
            else if (identificacaoOcorrencia.Equals("18"))
                strResp = "Acerto de Depositária";
            else if (identificacaoOcorrencia.Equals("19"))
                strResp = "Confirmação Receb. Inst. de Protesto";
            else if (identificacaoOcorrencia.Equals("20"))
                strResp = "Confirmação Recebimento Instrução Sustação de Protesto";
            else if (identificacaoOcorrencia.Equals("21"))
                strResp = "Acerto do Controle do Participante";
            else if (identificacaoOcorrencia.Equals("22"))
                strResp = "Título Com Pagamento Cancelado";
            else if (identificacaoOcorrencia.Equals("23"))
                strResp = "Entrada do Título em Cartório";
            else if (identificacaoOcorrencia.Equals("24"))
                strResp = "Entrada rejeitada por CEP Irregular";
            else if (identificacaoOcorrencia.Equals("27"))
                strResp = "Baixa Rejeitada";
            else if (identificacaoOcorrencia.Equals("28"))
                strResp = "Débito de tarifas/custas";
            else if (identificacaoOcorrencia.Equals("30"))
                strResp = "Alteração de Outros Dados Rejeitados";
            else if (identificacaoOcorrencia.Equals("32"))
                strResp = "Instrução Rejeitada";
            else if (identificacaoOcorrencia.Equals("33"))
                strResp = "Confirmação Pedido Alteração Outros Dados";
            else if (identificacaoOcorrencia.Equals("34"))
                strResp = "Retirado de Cartório e Manutenção Carteira";
            else if (identificacaoOcorrencia.Equals("35"))
                strResp = "Desagendamento do débito automático";
            else if (identificacaoOcorrencia.Equals("40"))
                strResp = "Estorno de pagamento";
            else if (identificacaoOcorrencia.Equals("55"))
                strResp = "Sustado judicial";
            else if (identificacaoOcorrencia.Equals("68"))
                strResp = "Acerto dos dados do rateio de Crédito";
            else if (identificacaoOcorrencia.Equals("69"))
                strResp = "Cancelamento dos dados do rateio";

            return strResp;
        }
        #endregion

        #region [ decodificaMotivoOcorrencia ]
        /// <summary>
        /// Retorna a lista com a descrição do motivo da ocorrência.
        /// Cada identificação de ocorrência possui seu próprio conjunto de motivos.
        /// No arquivo de retorno, o campo "Motivos das Rejeições para os Códigos de Ocorrência" é um campo de 10 posições,
        /// sendo que cada motivo ocupa 2 posições, portanto, neste campo podem ser informados até 5 motivos.
        /// </summary>
        /// <param name="identificacaoOcorrencia">
        /// Código da identificação da ocorrência.
        /// </param>
        /// <param name="motivosOcorrencia">
        /// Campo de 10 posições que informa até 5 motivos de ocorrência.
        /// </param>
        /// <returns>
        /// Retorna uma lista com os códigos e as descrições da ocorrência.
        /// </returns>
        public static List<TipoDescricaoMotivoOcorrencia> decodificaMotivoOcorrencia(String identificacaoOcorrencia, String motivosOcorrencia)
        {
            #region [ Declarações ]
            List<TipoDescricaoMotivoOcorrencia> listaRespostaMotivos = new List<TipoDescricaoMotivoOcorrencia>();
            List<String> listaMotivos = new List<String>();
            TipoDescricaoMotivoOcorrencia descricaoMotivoOcorrencia;
            String strMotivoOcorrencia;
            String strResp;
            String strMotivoAux;
            bool blnIgnorar;
            #endregion

            #region [ Consistência ]
            if (identificacaoOcorrencia == null) return listaRespostaMotivos;
            if (identificacaoOcorrencia.Trim().Length == 0) return listaRespostaMotivos;

            if (motivosOcorrencia == null) return listaRespostaMotivos;
            if (motivosOcorrencia.Trim().Length == 0) return listaRespostaMotivos;
            #endregion

            #region [ Decompõe os vários motivos de ocorrência ]
            // Lembrando que as posições não utilizadas são preenchidas com zeros e que
            // em algumas ocorrências, o motivo "00" é um motivo válido.
            // Portanto, após o 1º motivo, ignora os motivos "00" subsequentes.
            while (motivosOcorrencia.Length >= 2)
            {
                strMotivoAux = Texto.leftStr(motivosOcorrencia, 2);
                blnIgnorar = false;
                if (strMotivoAux.Equals("00") && listaMotivos.Count > 0) blnIgnorar = true;
                if (!blnIgnorar) listaMotivos.Add(strMotivoAux);
                motivosOcorrencia = motivosOcorrencia.Substring(2);
            }
            #endregion

            for (int i = 0; i < listaMotivos.Count; i++)
            {
                strResp = "";

                strMotivoOcorrencia = listaMotivos[i];

                #region [ Decodifica motivos da ocorrência ]
                if (identificacaoOcorrencia.Equals("02"))
                {
                    #region [ Decodifica motivos ocorrência 02 ]
                    if (strMotivoOcorrencia.Equals("00"))
                        strResp = "Ocorrência aceita";
                    else if (strMotivoOcorrencia.Equals("01"))
                        strResp = "Código do Banco inválido";
                    else if (strMotivoOcorrencia.Equals("04"))
                        strResp = "Código do movimento não permitido para a carteira";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Características da cobrança incompatíveis";
                    else if (strMotivoOcorrencia.Equals("17"))
                        strResp = "Data de vencimento anterior a data de emissão";
                    else if (strMotivoOcorrencia.Equals("21"))
                        strResp = "Espécie do Título inválido";
                    else if (strMotivoOcorrencia.Equals("24"))
                        strResp = "Data da emissão inválida";
                    else if (strMotivoOcorrencia.Equals("27"))
                        strResp = "Valor/taxa de juros mora inválido";
                    else if (strMotivoOcorrencia.Equals("38"))
                        strResp = "Prazo para protesto inválido";
                    else if (strMotivoOcorrencia.Equals("39"))
                        strResp = "Pedido para protesto não permitido para título";
                    else if (strMotivoOcorrencia.Equals("43"))
                        strResp = "Prazo para baixa e devolução inválido";
                    else if (strMotivoOcorrencia.Equals("45"))
                        strResp = "Nome do Sacado inválido";
                    else if (strMotivoOcorrencia.Equals("46"))
                        strResp = "Tipo/num. de inscrição do Sacado inválidos";
                    else if (strMotivoOcorrencia.Equals("47"))
                        strResp = "Endereço do Sacado não informado";
                    else if (strMotivoOcorrencia.Equals("48"))
                        strResp = "CEP inválido";
                    else if (strMotivoOcorrencia.Equals("50"))
                        strResp = "CEP referente a Banco correspondente";
                    else if (strMotivoOcorrencia.Equals("53"))
                        strResp = "Nº de inscrição do Sacador/avalista inválidos (CPF/CNPJ)";
                    else if (strMotivoOcorrencia.Equals("54"))
                        strResp = "Sacador/avalista não informado";
                    else if (strMotivoOcorrencia.Equals("67"))
                        strResp = "Débito automático agendado";
                    else if (strMotivoOcorrencia.Equals("68"))
                        strResp = "Débito não agendado - erro nos dados de remessa";
                    else if (strMotivoOcorrencia.Equals("69"))
                        strResp = "Débito não agendado - Sacado não consta no cadastro de autorizante";
                    else if (strMotivoOcorrencia.Equals("70"))
                        strResp = "Débito não agendado - Cedente não autorizado pelo Sacado";
                    else if (strMotivoOcorrencia.Equals("71"))
                        strResp = "Débito não agendado - Cedente não participa da modalidade de déb.automático";
                    else if (strMotivoOcorrencia.Equals("72"))
                        strResp = "Débito não agendado - Código de moeda diferente de R$";
                    else if (strMotivoOcorrencia.Equals("73"))
                        strResp = "Débito não agendado - Data de vencimento inválida";
                    else if (strMotivoOcorrencia.Equals("75"))
                        strResp = "Débito não agendado - Tipo do número de inscrição do sacado debitado inválido";
                    else if (strMotivoOcorrencia.Equals("86"))
                        strResp = "Seu número do documento inválido";
                    else if (strMotivoOcorrencia.Equals("89"))
                        strResp = "Email Sacado não enviado - título com débito automático";
                    else if (strMotivoOcorrencia.Equals("90"))
                        strResp = "Email sacado não enviado - título de cobrança sem registro";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("03"))
                {
                    #region [ Decodifica motivos ocorrência 03 ]
                    if (strMotivoOcorrencia.Equals("02"))
                        strResp = "Código do registro detalhe inválido";
                    else if (strMotivoOcorrencia.Equals("03"))
                        strResp = "Código da ocorrência inválida";
                    else if (strMotivoOcorrencia.Equals("04"))
                        strResp = "Código de ocorrência não permitida para a carteira";
                    else if (strMotivoOcorrencia.Equals("05"))
                        strResp = "Código de ocorrência não numérico";
                    else if (strMotivoOcorrencia.Equals("07"))
                        strResp = "Agência/conta/Digito - Inválido";
                    else if (strMotivoOcorrencia.Equals("08"))
                        strResp = "Nosso número inválido";
                    else if (strMotivoOcorrencia.Equals("09"))
                        strResp = "Nosso número duplicado";
                    else if (strMotivoOcorrencia.Equals("10"))
                        strResp = "Carteira inválida";
                    else if (strMotivoOcorrencia.Equals("13"))
                        strResp = "Identificação da emissão do bloqueto inválida";
                    else if (strMotivoOcorrencia.Equals("16"))
                        strResp = "Data de vencimento inválida";
                    else if (strMotivoOcorrencia.Equals("18"))
                        strResp = "Vencimento fora do prazo de operação";
                    else if (strMotivoOcorrencia.Equals("20"))
                        strResp = "Valor do Título inválido";
                    else if (strMotivoOcorrencia.Equals("21"))
                        strResp = "Espécie do Título inválida";
                    else if (strMotivoOcorrencia.Equals("22"))
                        strResp = "Espécie não permitida para a carteira";
                    else if (strMotivoOcorrencia.Equals("24"))
                        strResp = "Data de emissão inválida";
                    else if (strMotivoOcorrencia.Equals("28"))
                        strResp = "Código do desconto inválido";
                    else if (strMotivoOcorrencia.Equals("38"))
                        strResp = "Prazo para protesto inválido";
                    else if (strMotivoOcorrencia.Equals("44"))
                        strResp = "Agência Cedente não prevista";
                    else if (strMotivoOcorrencia.Equals("45"))
                        strResp = "Nome do sacado não informado";
                    else if (strMotivoOcorrencia.Equals("46"))
                        strResp = "Tipo/número de inscrição do sacado inválidos";
                    else if (strMotivoOcorrencia.Equals("47"))
                        strResp = "Endereço do sacado não informado";
                    else if (strMotivoOcorrencia.Equals("48"))
                        strResp = "CEP inválido";
                    else if (strMotivoOcorrencia.Equals("50"))
                        strResp = "CEP irregular - Banco Correspondente";
                    else if (strMotivoOcorrencia.Equals("63"))
                        strResp = "Entrada para Título já cadastrado";
                    else if (strMotivoOcorrencia.Equals("65"))
                        strResp = "Limite excedido";
                    else if (strMotivoOcorrencia.Equals("66"))
                        strResp = "Número autorização inexistente";
                    else if (strMotivoOcorrencia.Equals("68"))
                        strResp = "Débito não agendado - erro nos dados de remessa";
                    else if (strMotivoOcorrencia.Equals("69"))
                        strResp = "Débito não agendado - Sacado não consta no cadastro de autorizante";
                    else if (strMotivoOcorrencia.Equals("70"))
                        strResp = "Débito não agendado - Cedente não autorizado pelo Sacado";
                    else if (strMotivoOcorrencia.Equals("71"))
                        strResp = "Débito não agendado - Cedente não participa do débito Automático";
                    else if (strMotivoOcorrencia.Equals("72"))
                        strResp = "Débito não agendado - Código de moeda diferente de R$";
                    else if (strMotivoOcorrencia.Equals("73"))
                        strResp = "Débito não agendado - Data de vencimento inválida";
                    else if (strMotivoOcorrencia.Equals("74"))
                        strResp = "Débito não agendado - Conforme seu pedido, Título não registrado";
                    else if (strMotivoOcorrencia.Equals("75"))
                        strResp = "Débito não agendado – Tipo de número de inscrição do debitado inválido";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("06"))
                {
                    #region [ Decodifica motivos ocorrência 06 ]
                    if (strMotivoOcorrencia.Equals("00"))
                        strResp = "Título pago com dinheiro";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Título pago com cheque";
                    else if (strMotivoOcorrencia.Equals("42"))
                        strResp = "Rateio não efetuado, cód. Calculo 2 (VLR. Registro) e v";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("09"))
                {
                    #region [ Decodifica motivos ocorrência 09 ]
                    if (strMotivoOcorrencia.Equals("00"))
                        strResp = "Ocorrência Aceita";
                    else if (strMotivoOcorrencia.Equals("10"))
                        strResp = "Baixa Comandada pelo cliente";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("10"))
                {
                    #region [ Decodifica motivos ocorrência 10 ]
                    if (strMotivoOcorrencia.Equals("00"))
                        strResp = "Baixado Conforme Instruções da Agência";
                    else if (strMotivoOcorrencia.Equals("14"))
                        strResp = "Título Protestado";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Título excluído";
                    else if (strMotivoOcorrencia.Equals("16"))
                        strResp = "Título Baixado pelo Banco por decurso Prazo";
                    else if (strMotivoOcorrencia.Equals("20"))
                        strResp = "Título Baixado e transferido para desconto";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("15"))
                {
                    #region [ Decodifica motivos ocorrência 15 ]
                    if (strMotivoOcorrencia.Equals("00"))
                        strResp = "Título pago com dinheiro";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Título pago com cheque";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("17"))
                {
                    #region [ Decodifica motivos ocorrência 17 ]
                    if (strMotivoOcorrencia.Equals("00"))
                        strResp = "Título pago com dinheiro";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Título pago com cheque";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("24"))
                {
                    #region [ Decodifica motivos ocorrência 24 ]
                    if (strMotivoOcorrencia.Equals("48"))
                        strResp = "CEP inválido";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("27"))
                {
                    #region [ Decodifica motivos ocorrência 27 ]
                    if (strMotivoOcorrencia.Equals("04"))
                        strResp = "Código de ocorrência não permitido para a carteira";
                    else if (strMotivoOcorrencia.Equals("07"))
                        strResp = "Agência/Conta/dígito inválidos";
                    else if (strMotivoOcorrencia.Equals("08"))
                        strResp = "Nosso número inválido";
                    else if (strMotivoOcorrencia.Equals("10"))
                        strResp = "Carteira inválida";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Carteira/Agência/Conta/nosso número inválidos";
                    else if (strMotivoOcorrencia.Equals("40"))
                        strResp = "Título com ordem de protesto emitido";
                    else if (strMotivoOcorrencia.Equals("42"))
                        strResp = "Código para baixa/devolução via Telebradesco inválido";
                    else if (strMotivoOcorrencia.Equals("60"))
                        strResp = "Movimento para Título não cadastrado";
                    else if (strMotivoOcorrencia.Equals("77"))
                        strResp = "Transferência para desconto não permitido para a carteira";
                    else if (strMotivoOcorrencia.Equals("85"))
                        strResp = "Título com pagamento vinculado";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("28"))
                {
                    #region [ Decodifica motivos ocorrência 28 ]
                    if (strMotivoOcorrencia.Equals("02"))
                        strResp = "Tarifa de permanência título cadastrado";
                    else if (strMotivoOcorrencia.Equals("03"))
                        strResp = "Tarifa de sustação";
                    else if (strMotivoOcorrencia.Equals("04"))
                        strResp = "Tarifa de protesto";
                    else if (strMotivoOcorrencia.Equals("05"))
                        strResp = "Tarifa de outras instruções";
                    else if (strMotivoOcorrencia.Equals("06"))
                        strResp = "Tarifa de outras ocorrências";
                    else if (strMotivoOcorrencia.Equals("08"))
                        strResp = "Custas de protesto";
                    else if (strMotivoOcorrencia.Equals("12"))
                        strResp = "Tarifa de registro";
                    else if (strMotivoOcorrencia.Equals("13"))
                        strResp = "Tarifa título pago no Bradesco";
                    else if (strMotivoOcorrencia.Equals("14"))
                        strResp = "Tarifa título pago compensação";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Tarifa título baixado não pago";
                    else if (strMotivoOcorrencia.Equals("16"))
                        strResp = "Tarifa alteração de vencimento";
                    else if (strMotivoOcorrencia.Equals("17"))
                        strResp = "Tarifa concessão abatimento";
                    else if (strMotivoOcorrencia.Equals("18"))
                        strResp = "Tarifa cancelamento de abatimento";
                    else if (strMotivoOcorrencia.Equals("19"))
                        strResp = "Tarifa concessão desconto";
                    else if (strMotivoOcorrencia.Equals("20"))
                        strResp = "Tarifa cancelamento desconto";
                    else if (strMotivoOcorrencia.Equals("21"))
                        strResp = "Tarifa título pago cics";
                    else if (strMotivoOcorrencia.Equals("22"))
                        strResp = "Tarifa título pago Internet";
                    else if (strMotivoOcorrencia.Equals("23"))
                        strResp = "Tarifa título pago term. gerencial serviços";
                    else if (strMotivoOcorrencia.Equals("24"))
                        strResp = "Tarifa título pago Pag-Contas";
                    else if (strMotivoOcorrencia.Equals("25"))
                        strResp = "Tarifa título pago Fone Fácil";
                    else if (strMotivoOcorrencia.Equals("26"))
                        strResp = "Tarifa título pago Déb. Postagem";
                    else if (strMotivoOcorrencia.Equals("27"))
                        strResp = "Tarifa impressão de títulos pendentes";
                    else if (strMotivoOcorrencia.Equals("28"))
                        strResp = "Tarifa título pago BDN";
                    else if (strMotivoOcorrencia.Equals("29"))
                        strResp = "Tarifa título pago Term. Multi Função";
                    else if (strMotivoOcorrencia.Equals("30"))
                        strResp = "Impressão de títulos baixados";
                    else if (strMotivoOcorrencia.Equals("31"))
                        strResp = "Impressão de títulos pagos";
                    else if (strMotivoOcorrencia.Equals("32"))
                        strResp = "Tarifa título pago Pagfor";
                    else if (strMotivoOcorrencia.Equals("33"))
                        strResp = "Tarifa reg/pgto - guichê caixa";
                    else if (strMotivoOcorrencia.Equals("34"))
                        strResp = "Tarifa título pago retaguarda";
                    else if (strMotivoOcorrencia.Equals("35"))
                        strResp = "Tarifa título pago Subcentro";
                    else if (strMotivoOcorrencia.Equals("36"))
                        strResp = "Tarifa título pago Cartão de Crédito";
                    else if (strMotivoOcorrencia.Equals("37"))
                        strResp = "Tarifa título pago Comp Eletrônica";
                    else if (strMotivoOcorrencia.Equals("38"))
                        strResp = "Tarifa título Baix. Pg. Cartório";
                    else if (strMotivoOcorrencia.Equals("39"))
                        strResp = "Tarifa título baixado acerto BCO";
                    else if (strMotivoOcorrencia.Equals("40"))
                        strResp = "Baixa registro em duplicidade";
                    else if (strMotivoOcorrencia.Equals("41"))
                        strResp = "Tarifa título baixado decurso prazo";
                    else if (strMotivoOcorrencia.Equals("42"))
                        strResp = "Tarifa título baixado Judicialmente";
                    else if (strMotivoOcorrencia.Equals("43"))
                        strResp = "Tarifa título baixado via remessa";
                    else if (strMotivoOcorrencia.Equals("44"))
                        strResp = "Tarifa título baixado rastreamento";
                    else if (strMotivoOcorrencia.Equals("45"))
                        strResp = "Tarifa título baixado conf. Pedido";
                    else if (strMotivoOcorrencia.Equals("46"))
                        strResp = "Tarifa título baixado protestado";
                    else if (strMotivoOcorrencia.Equals("47"))
                        strResp = "Tarifa título baixado p/ devolução";
                    else if (strMotivoOcorrencia.Equals("48"))
                        strResp = "Tarifa título baixado franco pagto";
                    else if (strMotivoOcorrencia.Equals("49"))
                        strResp = "Tarifa título baixado SUST/RET/CARTÓRIO";
                    else if (strMotivoOcorrencia.Equals("50"))
                        strResp = "Tarifa título baixado SUS/SEM/REM/CARTÓRIO";
                    else if (strMotivoOcorrencia.Equals("51"))
                        strResp = "Tarifa título transferido desconto";
                    else if (strMotivoOcorrencia.Equals("52"))
                        strResp = "Cobrado baixa manual";
                    else if (strMotivoOcorrencia.Equals("53"))
                        strResp = "Baixa por acerto cliente";
                    else if (strMotivoOcorrencia.Equals("54"))
                        strResp = "Tarifa baixa por contabilidade";
                    else if (strMotivoOcorrencia.Equals("55"))
                        strResp = "BIFAX";
                    else if (strMotivoOcorrencia.Equals("56"))
                        strResp = "Consulta informações via internet";
                    else if (strMotivoOcorrencia.Equals("57"))
                        strResp = "Arquivo retorno via internet";
                    else if (strMotivoOcorrencia.Equals("58"))
                        strResp = "Tarifa emissão Papeleta";
                    else if (strMotivoOcorrencia.Equals("59"))
                        strResp = "Tarifa fornec papeleta semi preenchida";
                    else if (strMotivoOcorrencia.Equals("60"))
                        strResp = "Acondicionador de papeletas (RPB)S";
                    else if (strMotivoOcorrencia.Equals("61"))
                        strResp = "Acond. de papeletas (RPB)s PERSONAL";
                    else if (strMotivoOcorrencia.Equals("62"))
                        strResp = "Papeleta formulário branco";
                    else if (strMotivoOcorrencia.Equals("63"))
                        strResp = "Formulário A4 serrilhado";
                    else if (strMotivoOcorrencia.Equals("64"))
                        strResp = "Fornecimento de softwares transmiss";
                    else if (strMotivoOcorrencia.Equals("65"))
                        strResp = "Fornecimento de softwares consulta";
                    else if (strMotivoOcorrencia.Equals("66"))
                        strResp = "Fornecimento Micro Completo";
                    else if (strMotivoOcorrencia.Equals("67"))
                        strResp = "Fornecimento MODEM";
                    else if (strMotivoOcorrencia.Equals("68"))
                        strResp = "Fornecimento de máquina FAX";
                    else if (strMotivoOcorrencia.Equals("69"))
                        strResp = "Fornecimento de máquinas óticas";
                    else if (strMotivoOcorrencia.Equals("70"))
                        strResp = "Fornecimento de Impressoras";
                    else if (strMotivoOcorrencia.Equals("71"))
                        strResp = "Reativação de título";
                    else if (strMotivoOcorrencia.Equals("72"))
                        strResp = "Alteração de produto negociado";
                    else if (strMotivoOcorrencia.Equals("73"))
                        strResp = "Tarifa emissão de contra recibo";
                    else if (strMotivoOcorrencia.Equals("74"))
                        strResp = "Tarifa emissão 2ª via papeleta";
                    else if (strMotivoOcorrencia.Equals("75"))
                        strResp = "Tarifa regravação arquivo retorno";
                    else if (strMotivoOcorrencia.Equals("76"))
                        strResp = "Arq. Títulos a vencer mensal";
                    else if (strMotivoOcorrencia.Equals("77"))
                        strResp = "Listagem auxiliar de crédito";
                    else if (strMotivoOcorrencia.Equals("78"))
                        strResp = "Tarifa cadastro cartela instrução permanente";
                    else if (strMotivoOcorrencia.Equals("79"))
                        strResp = "Canalização de Crédito";
                    else if (strMotivoOcorrencia.Equals("80"))
                        strResp = "Cadastro de Mensagem Fixa";
                    else if (strMotivoOcorrencia.Equals("81"))
                        strResp = "Tarifa reapresentação automática título";
                    else if (strMotivoOcorrencia.Equals("82"))
                        strResp = "Tarifa registro título déb. Automático";
                    else if (strMotivoOcorrencia.Equals("83"))
                        strResp = "Tarifa Rateio de Crédito";
                    else if (strMotivoOcorrencia.Equals("84"))
                        strResp = "Emissão papeleta sem valor";
                    else if (strMotivoOcorrencia.Equals("85"))
                        strResp = "Sem uso";
                    else if (strMotivoOcorrencia.Equals("86"))
                        strResp = "Cadastro de reembolso de diferença";
                    else if (strMotivoOcorrencia.Equals("87"))
                        strResp = "Relatório fluxo de pagto";
                    else if (strMotivoOcorrencia.Equals("88"))
                        strResp = "Emissão Extrato mov. Carteira";
                    else if (strMotivoOcorrencia.Equals("89"))
                        strResp = "Mensagem campo local de pagto";
                    else if (strMotivoOcorrencia.Equals("90"))
                        strResp = "Cadastro Concessionária serv. Publ.";
                    else if (strMotivoOcorrencia.Equals("91"))
                        strResp = "Classif. Extrato Conta Corrente";
                    else if (strMotivoOcorrencia.Equals("92"))
                        strResp = "Contabilidade especial";
                    else if (strMotivoOcorrencia.Equals("93"))
                        strResp = "Realimentação pagto";
                    else if (strMotivoOcorrencia.Equals("94"))
                        strResp = "Repasse de Créditos";
                    else if (strMotivoOcorrencia.Equals("95"))
                        strResp = "Tarifa reg. pagto Banco Postal";
                    else if (strMotivoOcorrencia.Equals("96"))
                        strResp = "Tarifa reg. Pagto outras mídias";
                    else if (strMotivoOcorrencia.Equals("97"))
                        strResp = "Tarifa Reg/Pagto - Net Empresa";
                    else if (strMotivoOcorrencia.Equals("98"))
                        strResp = "Tarifa título pago vencido";
                    else if (strMotivoOcorrencia.Equals("99"))
                        strResp = "TR Tít. Baixado por decurso prazo";
                    else if (strMotivoOcorrencia.Equals("100"))
                        strResp = "Arquivo Retorno Antecipado";
                    else if (strMotivoOcorrencia.Equals("101"))
                        strResp = "Arq retorno Hora/Hora";
                    else if (strMotivoOcorrencia.Equals("102"))
                        strResp = "TR. Agendamento Déb Aut";
                    else if (strMotivoOcorrencia.Equals("103"))
                        strResp = "TR. Tentativa cons Déb Aut";
                    else if (strMotivoOcorrencia.Equals("104"))
                        strResp = "TR Crédito on-line";
                    else if (strMotivoOcorrencia.Equals("105"))
                        strResp = "TR. Agendamento rat. Crédito";
                    else if (strMotivoOcorrencia.Equals("106"))
                        strResp = "TR Emissão aviso rateio";
                    else if (strMotivoOcorrencia.Equals("107"))
                        strResp = "Extrato de protesto";
                    else if (strMotivoOcorrencia.Equals("110"))
                        strResp = "Tarifa reg/pagto Bradesco Expresso";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("30"))
                {
                    #region [ Decodifica motivos ocorrência 30 ]
                    if (strMotivoOcorrencia.Equals("01"))
                        strResp = "Código do Banco inválido";
                    else if (strMotivoOcorrencia.Equals("04"))
                        strResp = "Código de ocorrência não permitido para a carteira";
                    else if (strMotivoOcorrencia.Equals("05"))
                        strResp = "Código da ocorrência não numérico";
                    else if (strMotivoOcorrencia.Equals("08"))
                        strResp = "Nosso número inválido";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Característica da cobrança incompatível";
                    else if (strMotivoOcorrencia.Equals("16"))
                        strResp = "Data de vencimento inválido";
                    else if (strMotivoOcorrencia.Equals("17"))
                        strResp = "Data de vencimento anterior a data de emissão";
                    else if (strMotivoOcorrencia.Equals("18"))
                        strResp = "Vencimento fora do prazo de operação";
                    else if (strMotivoOcorrencia.Equals("24"))
                        strResp = "Data de emissão Inválida";
                    else if (strMotivoOcorrencia.Equals("26"))
                        strResp = "Código de juros de mora inválido";
                    else if (strMotivoOcorrencia.Equals("27"))
                        strResp = "Valor/taxa de juros de mora inválido";
                    else if (strMotivoOcorrencia.Equals("28"))
                        strResp = "Código de desconto inválido";
                    else if (strMotivoOcorrencia.Equals("29"))
                        strResp = "Valor do desconto maior/igual ao valor do Título";
                    else if (strMotivoOcorrencia.Equals("30"))
                        strResp = "Desconto a conceder não confere";
                    else if (strMotivoOcorrencia.Equals("31"))
                        strResp = "Concessão de desconto já existente ( Desconto anterior )";
                    else if (strMotivoOcorrencia.Equals("32"))
                        strResp = "Valor do IOF inválido";
                    else if (strMotivoOcorrencia.Equals("33"))
                        strResp = "Valor do abatimento inválido";
                    else if (strMotivoOcorrencia.Equals("34"))
                        strResp = "Valor do abatimento maior/igual ao valor do Título";
                    else if (strMotivoOcorrencia.Equals("38"))
                        strResp = "Prazo para protesto inválido";
                    else if (strMotivoOcorrencia.Equals("39"))
                        strResp = "Pedido de protesto não permitido para o Título";
                    else if (strMotivoOcorrencia.Equals("40"))
                        strResp = "Título com ordem de protesto emitido";
                    else if (strMotivoOcorrencia.Equals("42"))
                        strResp = "Código para baixa/devolução inválido";
                    else if (strMotivoOcorrencia.Equals("46"))
                        strResp = "Tipo/número de inscrição do sacado inválidos";
                    else if (strMotivoOcorrencia.Equals("48"))
                        strResp = "CEP Inválido";
                    else if (strMotivoOcorrencia.Equals("53"))
                        strResp = "Tipo/Número de inscrição do sacador/avalista inválidos";
                    else if (strMotivoOcorrencia.Equals("54"))
                        strResp = "Sacador/avalista não informado";
                    else if (strMotivoOcorrencia.Equals("57"))
                        strResp = "Código da multa inválido";
                    else if (strMotivoOcorrencia.Equals("58"))
                        strResp = "Data da multa inválida";
                    else if (strMotivoOcorrencia.Equals("60"))
                        strResp = "Movimento para Título não cadastrado";
                    else if (strMotivoOcorrencia.Equals("79"))
                        strResp = "Data de Juros de mora Inválida";
                    else if (strMotivoOcorrencia.Equals("80"))
                        strResp = "Data do desconto inválida";
                    else if (strMotivoOcorrencia.Equals("85"))
                        strResp = "Título com Pagamento Vinculado";
                    else if (strMotivoOcorrencia.Equals("88"))
                        strResp = "E-mail Sacado não lido no prazo 5 dias";
                    else if (strMotivoOcorrencia.Equals("91"))
                        strResp = "E-mail sacado não recebido";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("32"))
                {
                    #region [ Decodifica motivos ocorrência 32 ]
                    if (strMotivoOcorrencia.Equals("01"))
                        strResp = "Código do Banco inválido";
                    else if (strMotivoOcorrencia.Equals("02"))
                        strResp = "Código do registro detalhe inválido";
                    else if (strMotivoOcorrencia.Equals("04"))
                        strResp = "Código de ocorrência não permitido para a carteira";
                    else if (strMotivoOcorrencia.Equals("05"))
                        strResp = "Código de ocorrência não numérico";
                    else if (strMotivoOcorrencia.Equals("07"))
                        strResp = "Agência/Conta/dígito inválidos";
                    else if (strMotivoOcorrencia.Equals("08"))
                        strResp = "Nosso número inválido";
                    else if (strMotivoOcorrencia.Equals("10"))
                        strResp = "Carteira inválida";
                    else if (strMotivoOcorrencia.Equals("15"))
                        strResp = "Características da cobrança incompatíveis";
                    else if (strMotivoOcorrencia.Equals("16"))
                        strResp = "Data de vencimento inválida";
                    else if (strMotivoOcorrencia.Equals("17"))
                        strResp = "Data de vencimento anterior a data de emissão";
                    else if (strMotivoOcorrencia.Equals("18"))
                        strResp = "Vencimento fora do prazo de operação";
                    else if (strMotivoOcorrencia.Equals("20"))
                        strResp = "Valor do título inválido";
                    else if (strMotivoOcorrencia.Equals("21"))
                        strResp = "Espécie do Título inválida";
                    else if (strMotivoOcorrencia.Equals("22"))
                        strResp = "Espécie não permitida para a carteira";
                    else if (strMotivoOcorrencia.Equals("24"))
                        strResp = "Data de emissão inválida";
                    else if (strMotivoOcorrencia.Equals("28"))
                        strResp = "Código de desconto via Telebradesco inválido";
                    else if (strMotivoOcorrencia.Equals("29"))
                        strResp = "Valor do desconto maior/igual ao valor do Título";
                    else if (strMotivoOcorrencia.Equals("30"))
                        strResp = "Desconto a conceder não confere";
                    else if (strMotivoOcorrencia.Equals("31"))
                        strResp = "Concessão de desconto - Já existe desconto anterior";
                    else if (strMotivoOcorrencia.Equals("33"))
                        strResp = "Valor do abatimento inválido";
                    else if (strMotivoOcorrencia.Equals("34"))
                        strResp = "Valor do abatimento maior/igual ao valor do Título";
                    else if (strMotivoOcorrencia.Equals("36"))
                        strResp = "Concessão abatimento - Já existe abatimento anterior";
                    else if (strMotivoOcorrencia.Equals("38"))
                        strResp = "Prazo para protesto inválido";
                    else if (strMotivoOcorrencia.Equals("39"))
                        strResp = "Pedido de protesto não permitido para o Título";
                    else if (strMotivoOcorrencia.Equals("40"))
                        strResp = "Título com ordem de protesto emitido";
                    else if (strMotivoOcorrencia.Equals("41"))
                        strResp = "Pedido cancelamento/sustação para Título sem instrução de protesto";
                    else if (strMotivoOcorrencia.Equals("42"))
                        strResp = "Código para baixa/devolução inválido";
                    else if (strMotivoOcorrencia.Equals("45"))
                        strResp = "Nome do Sacado não informado";
                    else if (strMotivoOcorrencia.Equals("46"))
                        strResp = "Tipo/número de inscrição do Sacado inválidos";
                    else if (strMotivoOcorrencia.Equals("47"))
                        strResp = "Endereço do Sacado não informado";
                    else if (strMotivoOcorrencia.Equals("48"))
                        strResp = "CEP Inválido";
                    else if (strMotivoOcorrencia.Equals("50"))
                        strResp = "CEP referente a um Banco correspondente";
                    else if (strMotivoOcorrencia.Equals("53"))
                        strResp = "Tipo de inscrição do sacador avalista inválidos";
                    else if (strMotivoOcorrencia.Equals("60"))
                        strResp = "Movimento para Título não cadastrado";
                    else if (strMotivoOcorrencia.Equals("85"))
                        strResp = "Título com pagamento vinculado";
                    else if (strMotivoOcorrencia.Equals("86"))
                        strResp = "Seu número inválido";
                    else if (strMotivoOcorrencia.Equals("94"))
                        strResp = "Título Penhorado - Instrução Não Liberada pela Agência";
                    #endregion
                }
                else if (identificacaoOcorrencia.Equals("35"))
                {
                    #region [ Decodifica motivos ocorrência 35 ]
                    if (strMotivoOcorrencia.Equals("81"))
                        strResp = "Tentativas esgotadas, baixado";
                    else if (strMotivoOcorrencia.Equals("82"))
                        strResp = "Tentativas esgotadas, pendente";
                    else if (strMotivoOcorrencia.Equals("83"))
                        strResp = "Cancelado pelo Sacado e Mantido Pendente, conforme negociação";
                    else if (strMotivoOcorrencia.Equals("84"))
                        strResp = "Cancelado pelo sacado e baixado, conforme negociação";
                    #endregion
                }
                #endregion

                if (strResp.Length > 0)
                {
                    descricaoMotivoOcorrencia = new TipoDescricaoMotivoOcorrencia(identificacaoOcorrencia, strMotivoOcorrencia, decodificaIdentificacaoOcorrencia(identificacaoOcorrencia), strResp);
                    listaRespostaMotivos.Add(descricaoMotivoOcorrencia);
                }
            }

            return listaRespostaMotivos;
        }
        #endregion

        #region [ decodificaMotivoOcorrencia19 ]
        public static String decodificaMotivoOcorrencia19(String motivoOcorrencia19)
        {
            String strResp = "";

            if (motivoOcorrencia19 == null) return "";
            if (motivoOcorrencia19.Trim().Length == 0) return "";

            #region [ Decodifica motivos ocorrência 19 ]
            if (motivoOcorrencia19.Equals("A"))
                strResp = "Aceito";
            else if (motivoOcorrencia19.Equals("D"))
                strResp = "Desprezado";
            #endregion

            return strResp;
        }
        #endregion

        #region [ decodificaBoletoNumeroControleParticipante ]
        /// <summary>
        /// Processa o campo "número de controle do participante" analisando se está no
        /// formato aguardado e tentando obter a identificação do registro do boleto no BD
        /// que está codificada neste campo.
        /// </summary>
        /// <param name="campoNumeroControleParticipante">
        /// Conteúdo do campo "número de controle do participante" que consta no registro
        /// do arquivo de retorno de boletos.
        /// </param>
        /// <returns>
        /// Retorna o número de identificação do registro do boleto no BD:
        ///		Zero: falha na decodificação
        ///		Maior que zero: número de identificação do registro do boleto no BD
        /// </returns>
        public static int decodificaBoletoNumeroControleParticipante(string campoNumeroControleParticipante, ref string strMsgErro)
        {
            #region [ Declarações ]
            int idBoletoItem;
            String[] vId;
            String strIdBoletoItem;
            #endregion

            strMsgErro = "";

            if (campoNumeroControleParticipante == null)
            {
                strMsgErro = "O campo que identifica o registro do boleto está com conteúdo nulo!!";
                return 0;
            }
            if (campoNumeroControleParticipante.Trim().Length == 0)
            {
                strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
                return 0;
            }
            if (campoNumeroControleParticipante.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
            {
                strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
                return 0;
            }

            #region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
            vId = campoNumeroControleParticipante.Split('=');
            strIdBoletoItem = vId[1];
            #endregion

            #region [ Consiste o valor do campo com o Id ]
            if (strIdBoletoItem == null)
            {
                strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
                return 0;
            }

            if (strIdBoletoItem.Trim().Length == 0)
            {
                strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
                return 0;
            }

            idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

            if (idBoletoItem <= 0)
            {
                strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
                return 0;
            }
            #endregion

            return idBoletoItem;
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

        #region [ excel_converte_numeracao_digito_para_letra ]
        public static string excel_converte_numeracao_digito_para_letra(int numeracao_digito)
        {
            #region [ Declarações ]
            const int TOTAL_LETRAS_ALFABETO = 26;
            string strResp;
            int intQuoc;
            int intResto;
            #endregion

            strResp = "";
            if (numeracao_digito <= 0) return "";
            intQuoc = (int)(numeracao_digito - 1) / TOTAL_LETRAS_ALFABETO;
            intResto = numeracao_digito - (intQuoc * TOTAL_LETRAS_ALFABETO);
            if (intQuoc > TOTAL_LETRAS_ALFABETO) return "";
            if (intQuoc > 0) strResp = ((char)(65 - 1 + intQuoc)).ToString();
            strResp += ((char)(65 - 1 + intResto)).ToString();
            return strResp;
        }
        #endregion

        #region [ executaManutencaoArqLogAtividade ]
        /// <summary>
        /// Apaga os arquivos de log de atividade antigos
        /// </summary>
        public static void executaManutencaoArqLogAtividade()
        {
            #region [ Declarações ]
            DateTime dtCorte = DateTime.Now.AddDays(-Global.Cte.LogAtividade.CorteArqLogEmDias);
            string strDataCorte = dtCorte.ToString(Global.Cte.DataHora.FmtYYYYMMDD);
            string[] ListaArqLog;
            string strNomeArq;
            int i;
            #endregion

            #region[ Apaga arquivos de log de atividade antigos ]
            ListaArqLog = Directory.GetFiles(Global.Cte.LogAtividade.PathLogAtividade, "*." + Global.Cte.LogAtividade.ExtensaoArqLog, SearchOption.TopDirectoryOnly);
            for (i = 0; i < ListaArqLog.Length; i++)
            {
                strNomeArq = Global.extractFileName(ListaArqLog[i]);
                strNomeArq = strNomeArq.Substring(0, strDataCorte.Length);
                if (string.Compare(strNomeArq, strDataCorte) < 0) File.Delete(ListaArqLog[i]);
            }
            #endregion
        }
        #endregion

        #region [ existeMotivoOcorrencia ]
        public static bool existeMotivoOcorrencia(String motivosOcorrencia, String motivoOcorrenciaAProcurar)
        {
            #region [ Declarações ]
            String strMotivoAux;
            #endregion

            #region [ Consistência ]
            if (motivosOcorrencia == null) return false;
            if (motivosOcorrencia.Trim().Length == 0) return false;

            if (motivoOcorrenciaAProcurar == null) return false;
            if (motivoOcorrenciaAProcurar.Trim().Length == 0) return false;
            #endregion

            while (motivosOcorrencia.Length >= 2)
            {
                strMotivoAux = Texto.leftStr(motivosOcorrencia, 2);
                if (strMotivoAux.Equals(motivoOcorrenciaAProcurar)) return true;
                motivosOcorrencia = motivosOcorrencia.Substring(2);
            }

            return false;
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

        #region [ filtraDigitacaoCep ]
        /// <summary>
        /// Filtra os caracteres durante a digitação de CEP
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoCep(char c)
        {
            if (!(isDigit(c) || (c == '-') || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoCnpjCpf ]
        /// <summary>
        /// Filtra os caracteres durante a digitação de CNPJ/CPF
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoCnpjCpf(char c)
        {
            if (!(isDigit(c) || (c == '.') || (c == '-') || (c == '/') || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoData ]
        /// <summary>
        /// Filtra os caracteres durante a digitação da data
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoData(char c)
        {
            // Deixa passar somente dígitos, o caracter separador de data e o backspace,
            // caso contrário, retorna o caracter nulo.
            if (!(((c >= '0') && (c <= '9')) || (c == '/') || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoEmail ]
        /// <summary>
        /// Filtra os caracteres durante a digitação de endereço de e-mail, aceitando também os
        /// seguintes caracteres separadores quando é digitada uma lista de e-mails: espaço em branco,
        /// vírgula e ponto e vírgula
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoEmail(char c)
        {
            if (!(isDigit(c) || isLetra(c) || (c == '@') || (c == '.') || (c == '_') || (c == '-') || (c == ' ') || (c == ',') || (c == ';') || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoMoeda ]
        /// <summary>
        /// Filtra os caracteres durante a digitação de valor monetário
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoMoeda(char c)
        {
            // Deixa passar somente dígitos, o sinal negativo, os caracteres separadores de milhar e 
            // decimal e o backspace, caso contrário, retorna o caracter nulo.
            if (!(((c >= '0') && (c <= '9')) || (c == '.') || (c == ',') || (c == '-') || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoNumeroInteiro ]
        public static char filtraDigitacaoNumeroInteiro(char c)
        {
            // Deixa passar somente dígitos e o backspace, caso contrário, retorna o caracter nulo.
            if (!(((c >= '0') && (c <= '9')) || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoNumeroPedido ]
        public static char filtraDigitacaoNumeroPedido(char c)
        {
            char letra;
            if (c == '\b') return c;
            letra = Char.ToUpper(c);
            if ((!isDigit(letra)) && (!isLetra(letra)) && (letra != Cte.Etc.COD_SEPARADOR_FILHOTE)) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoPercentual ]
        /// <summary>
        /// Filtra os caracteres durante a digitação de número percentual
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoPercentual(char c)
        {
            // Deixa passar somente dígitos, o caracter separador de decimal e o backspace, caso contrário,
            // retorna o caracter nulo.
            if (!(((c >= '0') && (c <= '9')) || (c == '.') || (c == ',') || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoSomenteLetras ]
        public static char filtraDigitacaoSomenteLetras(char c)
        {
            // Deixa passar somente letras e o backspace, caso contrário, retorna o caracter nulo.
            if (!(((c >= 'a') && (c <= 'z')) || ((c >= 'A') && (c <= 'Z')) || (c == '\b'))) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraDigitacaoTexto ]
        /// <summary>
        /// Filtra os caracteres durante a digitação de campo texto livre
        /// </summary>
        /// <param name="c">
        /// Caracter digitado
        /// </param>
        /// <returns>
        /// Retorna o próprio caracter se ele for válido ou, caso contrário, o caracter nulo
        /// </returns>
        public static char filtraDigitacaoTexto(char c)
        {
            // Filtra os caracteres Ascii 34 e 39 (aspas duplas e aspas simples, respectivamente)
            if ((c == '\x0022') || (c == '\x0027') || (c == '|')) c = '\0';
            return c;
        }
        #endregion

        #region [ filtraTexto ]
        /// <summary>
        /// Filtra caracteres inválidos para um campo texto livre. Ex: aspas simples, aspas duplas, etc.
        /// </summary>
        /// <param name="texto">
        /// Conteúdo de um campo do tipo texto livre.
        /// </param>
        /// <returns>
        /// Retorna o texto sem conter nenhum caracter inválido para um campo do tipo texto livre.
        /// </returns>
        public static String filtraTexto(String texto)
        {
            StringBuilder sb = new StringBuilder("");
            for (int i = 0; i < texto.Length; i++)
            {
                if ((texto[i] != '\x0022') &&
                    (texto[i] != '\x0027'))
                {
                    sb.Append(texto[i]);
                }
            }
            return sb.ToString();
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

        #region [ formataBoletoNossoNumero ]
        public static String formataBoletoNossoNumero(String nossoNumeroSemDigito, String digitoNossoNumero)
        {
            String strSeparador = "-";

            if (nossoNumeroSemDigito == null) return "";
            if (nossoNumeroSemDigito.Trim().Length == 0) return "";

            if (digitoNossoNumero == null) digitoNossoNumero = "";
            if (digitoNossoNumero.Trim().Length == 0) strSeparador = "";

            return nossoNumeroSemDigito.Trim() + strSeparador + digitoNossoNumero.Trim();
        }
        #endregion

        #region [ formataCep ]
        public static String formataCep(String cep)
        {
            String strCep;
            if (cep == null) return "";
            strCep = digitos(cep);
            if (strCep.Length != 8) return cep;
            strCep = strCep.Substring(0, 5) + '-' + strCep.Substring(5, 3);
            return strCep;
        }
        #endregion

        #region [ formataCnpjCpf ]
        /// <summary>
        /// Formata os dígitos de CNPJ/CPF informados aplicando a máscara de formatação
        /// </summary>
        /// <param name="cnpj_cpf">
        /// Dígitos do CNPJ/CPF
        /// </param>
        /// <returns>
        /// Retorna o CNPJ/CPF formatado
        /// </returns>
        public static String formataCnpjCpf(String cnpj_cpf)
        {
            String s;
            String s_aux;
            String s_resp;

            if (cnpj_cpf == null) return "";

            s = digitos(cnpj_cpf);

            #region [ Verifica se é um CNPJ mesmo ou se é um CPF c/ zeros p/ normalizar à esquerda ]
            if (s.Length == 14)
            {
                if (!isCnpjOk(s))
                {
                    if (Texto.leftStr(s, 3).Equals("000"))
                    {
                        s_aux = Texto.rightStr(s, 11);
                        if (isCpfOk(s_aux)) s = s_aux;
                    }
                }
            }
            #endregion

            // CPF
            if (s.Length == 11)
            {
                s_resp = s.Substring(0, 3) + '.' + s.Substring(3, 3) + '.' + s.Substring(6, 3) + '/' + s.Substring(9, 2);
            }
            // CNPJ
            else if (s.Length == 14)
            {
                s_resp = s.Substring(0, 2) + '.' + s.Substring(2, 3) + '.' + s.Substring(5, 3) + '/' + s.Substring(8, 4) + '-' + s.Substring(12, 2);
            }
            // Desconhecido
            else
            {
                s_resp = cnpj_cpf;
            }
            return s_resp;
        }
        #endregion

        #region [ formataDataCampoArquivoDdMmYyParaDDMMYYYYComSeparador ]
        /// <summary>
        /// A partir de uma data vindo de um arquivo no formato DDMMYY, tenta normalizar e retornar uma data no formato DD/MM/YYYY
        /// </summary>
        /// <param name="data">
        /// Texto com a data a ser normalizada
        /// Formatos aceitos: DDMMYY
        /// O valor 000000 indica que o campo está vazio e, neste caso, retorna uma String vazia
        /// </param>
        /// <returns>
        /// Retorna a data no formato DD/MM/YYYY caso a data informada esteja em um formato válido, caso contrário, retorna o próprio valor do parâmetro
        /// </returns>
        public static String formataDataCampoArquivoDdMmYyParaDDMMYYYYComSeparador(String data)
        {
            String strDia;
            String strMes;
            String strAno;

            if (data == null) return "";
            if (data.Trim().Length == 0) return "";
            if (data.Equals("000000")) return "";

            if (data.IndexOf('/') == -1)
            {
                // A data foi digitada sem os separadores
                data = digitos(data);
                // Neste caso, aceita somente se tiver sido digitada no formado DDMM ou DDMMYY ou DDMMYYYY
                if ((data.Length != 4) && (data.Length != 6) && (data.Length != 8)) return data;
                strDia = data.Substring(0, 2);
                strMes = data.Substring(2, 2);
                if (data.Length > 4)
                    strAno = data.Substring(4, data.Length - 4);
                else
                    strAno = DateTime.Now.ToString(Cte.DataHora.FmtAno);
            }
            else
            {
                String[] v = data.Split('/');
                // É necessário que a data tenha vindo separada em 2 ou 3 partes: dia/mês ou dia/mês/ano
                if ((v.Length != 2) && (v.Length != 3)) return data;
                for (int i = 0; i < v.Length; i++)
                {
                    if (v[i] == null) return data;
                    v[i] = digitos(v[i]);
                    if (v[i].Trim().Length == 0) return data;
                }
                strDia = v[0].PadLeft(2, '0');
                strMes = v[1].PadLeft(2, '0');
                if (v.Length > 2)
                    strAno = v[2];
                else
                    strAno = DateTime.Now.ToString(Cte.DataHora.FmtAno);
            }

            if (strAno.Length == 3)
            {
                if (converteInteiro(strAno) >= 900) strAno = "1" + strAno; else strAno = "2" + strAno;
            }
            else if (strAno.Length == 2)
            {
                if (converteInteiro(strAno) >= 80) strAno = "19" + strAno; else strAno = "20" + strAno;
            }
            else if (strAno.Length == 1)
            {
                strAno = DateTime.Now.Year.ToString().Substring(0, 3) + strAno;
            }
            else if (strAno.Length != 4) return data;

            return strDia + "/" + strMes + "/" + strAno;
        }
        #endregion

        #region [ formataDataDdMmYyComSeparador ]
        /// <summary>
        /// A partir de uma data do tipo DateTime, formata um texto com a representação da data no formato DD/MM/YY
        /// </summary>
        /// <param name="data">
        /// Data em parâmetro do tipo DateTime
        /// </param>
        /// <returns>
        /// Retorna a data representada em um texto no formato DD/MM/YY
        /// </returns>
        public static String formataDataDdMmYyComSeparador(DateTime data)
        {
            if (data == null) return "";
            if (data == DateTime.MinValue) return "";
            return data.ToString(Global.Cte.DataHora.FmtDdMmYyComSeparador);
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

        #region [ formataDataDigitadaParaDDMMYYYYComSeparador ]
        /// <summary>
        /// A partir de uma data digitada pelo usuário, tenta normalizar e retornar uma data no formato DD/MM/YYYY
        /// </summary>
        /// <param name="data">
        /// Texto com a data digitada pelo usuário a ser normalizado
        /// Formatos aceitos: DDMMYY, DDMMYYYY, DD/MM/YY, DD/MM/YYYY
        /// </param>
        /// <returns>
        /// Retorna a data no formato DD/MM/YYYY caso a data informada esteja em um formato válido, caso contrário, retorna o próprio texto informado
        /// </returns>
        public static String formataDataDigitadaParaDDMMYYYYComSeparador(String data)
        {
            String strDia;
            String strMes;
            String strAno;

            if (data == null) return "";
            if (data.Trim().Length == 0) return "";

            if (data.IndexOf('/') == -1)
            {
                // A data foi digitada sem os separadores
                data = digitos(data);
                // Neste caso, aceita somente se tiver sido digitada no formado DDMM ou DDMMYY ou DDMMYYYY
                if ((data.Length != 4) && (data.Length != 6) && (data.Length != 8)) return data;
                strDia = data.Substring(0, 2);
                strMes = data.Substring(2, 2);
                if (data.Length > 4)
                    strAno = data.Substring(4, data.Length - 4);
                else
                    strAno = DateTime.Now.ToString(Cte.DataHora.FmtAno);
            }
            else
            {
                String[] v = data.Split('/');
                // É necessário que a data tenha vindo separada em 2 ou 3 partes: dia/mês ou dia/mês/ano
                if ((v.Length != 2) && (v.Length != 3)) return data;
                for (int i = 0; i < v.Length; i++)
                {
                    if (v[i] == null) return data;
                    v[i] = digitos(v[i]);
                    if (v[i].Trim().Length == 0) return data;
                }
                strDia = v[0].PadLeft(2, '0');
                strMes = v[1].PadLeft(2, '0');
                if (v.Length > 2)
                    strAno = v[2];
                else
                    strAno = DateTime.Now.ToString(Cte.DataHora.FmtAno);
            }

            if (strAno.Length == 3)
            {
                if (converteInteiro(strAno) >= 900) strAno = "1" + strAno; else strAno = "2" + strAno;
            }
            else if (strAno.Length == 2)
            {
                if (converteInteiro(strAno) >= 80) strAno = "19" + strAno; else strAno = "20" + strAno;
            }
            else if (strAno.Length == 1)
            {
                strAno = DateTime.Now.Year.ToString().Substring(0, 3) + strAno;
            }
            else if (strAno.Length != 4) return data;

            return strDia + "/" + strMes + "/" + strAno;
        }
        #endregion

        #region [ formataDataDigitadaParaMMYYYYComSeparador ]
        /// <summary>
        /// A partir de uma data digitada pelo usuário, tenta normalizar e retornar uma data no formato MM/YYYY
        /// </summary>
        /// <param name="data">
        /// Texto com a data digitada pelo usuário a ser normalizado
        /// Formatos aceitos: MMYY, MMYYYY, MM/YY, MM/YYYY
        /// </param>
        /// <returns>
        /// Retorna a data no formato MM/YYYY caso a data informada esteja em um formato válido, caso contrário, retorna o próprio texto informado
        /// </returns>
        public static String formataDataDigitadaParaMMYYYYComSeparador(String data)
        {
            String strMes;
            String strAno;

            if (data == null) return "";
            if (data.Trim().Length == 0) return "";

            if (data.IndexOf('/') == -1)
            {
                // A data foi digitada sem os separadores
                data = digitos(data);
                // Neste caso, aceita somente se tiver sido digitada no formado MM ou MMYY ou MMYYYY
                if ((data.Length != 4) && (data.Length != 6)) return data;
                strMes = data.Substring(0, 2);
                strAno = data.Substring(2, data.Length - 2);
            }
            else
            {
                String[] v = data.Split('/');
                // É necessário que a data tenha vindo separada em 2 partes: mês/ano
                if ((v.Length != 2)) return data;
                for (int i = 0; i < v.Length; i++)
                {
                    if (v[i] == null) return data;
                    v[i] = digitos(v[i]);
                    if (v[i].Trim().Length == 0) return data;
                }
                strMes = v[0].PadLeft(2, '0');
                strAno = v[1];

            }

            if (strAno.Length == 3)
            {
                if (converteInteiro(strAno) >= 900) strAno = "1" + strAno; else strAno = "2" + strAno;
            }
            else if (strAno.Length == 2)
            {
                if (converteInteiro(strAno) >= 80) strAno = "19" + strAno; else strAno = "20" + strAno;
            }
            else if (strAno.Length == 1)
            {
                strAno = DateTime.Now.Year.ToString().Substring(0, 3) + strAno;
            }
            else if (strAno.Length != 4) return data;

            return strMes + "/" + strAno;
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

        #region [ formataEndereco ]
        public static String formataEndereco(String endereco, String endereco_numero, String endereco_complemento, String bairro, String cidade, String uf, String cep)
        {
            #region [ Declarações ]
            String strEndereco = "";
            String strEnderecoNumero = "";
            String strEnderecoComplemento = "";
            String strBairro = "";
            String strCidade = "";
            String strUf = "";
            String strCep = "";
            String strResposta = "";
            #endregion

            if (endereco != null) strEndereco = endereco.Trim();
            if (endereco_numero != null) strEnderecoNumero = endereco_numero.Trim();
            if (endereco_complemento != null) strEnderecoComplemento = endereco_complemento.Trim();
            if (bairro != null) strBairro = bairro.Trim();
            if (cidade != null) strCidade = cidade.Trim();
            if (uf != null) strUf = uf.Trim();
            if (cep != null) strCep = cep.Trim();

            if (strEndereco.Length == 0) return "";

            strResposta = strEndereco;
            if (strEnderecoNumero.Length > 0) strResposta += ", " + strEnderecoNumero;
            if (strEnderecoComplemento.Length > 0) strResposta += " " + strEnderecoComplemento;
            if (strBairro.Length > 0) strResposta += " - " + strBairro;
            if (strCidade.Length > 0) strResposta += " - " + strCidade;
            if (strUf.Length > 0) strResposta += " - " + strUf;
            if (strCep.Length > 0) strResposta += " - " + formataCep(strCep);

            return strResposta;
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

        #region [ formataMoedaDigitada ]
        /// <summary>
        /// A partir de um valor digitado pelo usuário, tentar normalizar e retornar um valor monetário
        /// formatado com separador de milhar e de decimais
        /// </summary>
        /// <param name="numero">
        /// Texto com o valor monetário digitado a ser normalizado, positivo ou negativo
        /// </param>
        /// <returns>
        /// Retorna o valor formatado com separador de milhar e de decimais: 999.999,99
        /// </returns>
        public static String formataMoedaDigitada(String numero)
        {
            #region [ Declarações ]
            int i;
            int j;
            char c_separador_decimal;
            String s_numero_aux;
            String s_inteiro = "";
            String s_centavos = "";
            String s_valor_formatado;
            String s_sinal = "";
            #endregion

            if (numero == null) return "";
            if (numero.Trim().Length == 0) return "";

            numero = numero.Trim();

            if (numero.IndexOf('-') != -1) s_sinal = "-";

            c_separador_decimal = retornaSeparadorDecimal(numero);

            #region [ Formata o valor monetário ]
            s_numero_aux = numero.Replace(c_separador_decimal, 'V');
            String[] v = s_numero_aux.Split('V');
            for (i = 0; i < v.Length; i++)
            {
                if (v[i] == null) v[i] = "";
            }
            // Falha ao determinar o separador de decimal, então retorna o próprio valor informado
            if (v.Length > 2) return numero;

            if (v.Length >= 1) s_inteiro = digitos(v[0]);
            if (v.Length >= 2) s_centavos = digitos(v[1]);
            if (s_inteiro.Length == 0) s_inteiro = "0";
            s_centavos = Texto.leftStr(s_centavos, 2);
            s_centavos = s_centavos.PadRight(2, '0');

            // Coloca os separadores de milhar
            s_numero_aux = "";
            j = 0;
            for (i = s_inteiro.Length - 1; i >= 0; i--)
            {
                j++;
                s_numero_aux = s_inteiro[i] + s_numero_aux;
                if (((j % 3) == 0) && (i != s_inteiro.Length - 1) && (i != 0)) s_numero_aux = "." + s_numero_aux;
            }
            s_inteiro = s_numero_aux;

            s_valor_formatado = s_sinal + s_inteiro + "," + s_centavos;
            #endregion

            return s_valor_formatado;
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
        public static String formataPercentual(double valor)
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
        public static String formataPercentualCom1Decimal(double valor)
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
        public static String formataPercentualCom2Decimais(double valor)
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

        #region [ formataTelefone ]
        public static String formataTelefone(String telefone)
        {
            int i;
            String strTel = "";

            if (telefone != null) strTel = digitos(telefone);
            if ((strTel.Length == 0) || (strTel.Length > 8) || (!isTelefoneOk(strTel))) return strTel;

            i = strTel.Length - 4;
            strTel = strTel.Substring(0, i) + "-" + strTel.Substring(i);
            return strTel;
        }

        public static String formataTelefone(String ddd, String telefone)
        {
            String strDDD = "";
            String strTel;
            strTel = formataTelefone(telefone);
            if (ddd != null) strDDD = digitos(ddd);
            if ((strTel.Length > 0) && (strDDD.Length > 0)) strTel = "(" + strDDD + ") " + strTel;
            return strTel;
        }

        public static String formataTelefone(String ddd, String telefone, String ramal)
        {
            String strRamal = "";
            String strTel;
            strTel = formataTelefone(ddd, telefone);
            if (ramal != null) strRamal = digitos(ramal);
            if ((strTel.Length > 0) && (strRamal.Length > 0)) strTel += " R:" + strRamal;
            return strTel;
        }
        #endregion

        #region [ getBackColorFromAppConfig ]
        public static Color? getBackColorFromAppConfig()
        {
            #region[ Declarações ]
            string sBackColor;
            #endregion

            #region [ Define a cor de fundo de acordo com o ambiente acessado ]
            sBackColor = ConfigurationManager.AppSettings["backgroundColorPainel"];
            return converteColorFromHtml(sBackColor);
            #endregion
        }
        #endregion

        #region [ getChecksum ]
        public static string GetChecksum(string file)
        {
            using (FileStream stream = File.OpenRead(file))
            {
                SHA256Managed sha = new SHA256Managed();
                byte[] checksum = sha.ComputeHash(stream);
                return BitConverter.ToString(checksum).Replace("-", String.Empty);
            }
        }
        #endregion

        #region [ getVScrollBarWidth ]
        /// <summary>
        /// Dado um componente (ex: DataGridView) que contém um vertical scroll bar, retorna a largura do scroll bar
        /// </summary>
        /// <param name="control">
        /// Objeto que contém o scroll bar
        /// </param>
        /// <returns>
        /// Retorna a largura do scroll bar
        /// </returns>
        public static int getVScrollBarWidth(Control control)
        {
            foreach (Control c in control.Controls)
            {
                if (c.GetType().Equals(typeof(VScrollBar)))
                {
                    return c.Width;
                }
            }
            return 0;
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

        #region[ haOutraInstanciaEmExecucao ]
        public static bool haOutraInstanciaEmExecucao()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);

            //Loop through the running processes in with the same name 
            foreach (Process process in processes)
            {
                //Ignore the current process 
                if (process.Id != current.Id)
                {
                    return true;
                }
            }
            return false;
        }
        #endregion

        #region [ isAlfaNumerico ]
        public static bool isAlfaNumerico(char c)
        {
            if (isDigit(c) || isLetra(c)) return true;
            return false;
        }
        #endregion

        #region [ isCepOk ]
        public static bool isCepOk(String cep)
        {
            String strCep;
            if (cep == null) return false;
            strCep = digitos(cep);
            if ((strCep.Length == 5) || (strCep.Length == 8)) return true;
            return false;
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

        #region [ isDataMMYYYYOk ]
        /// <summary>
        /// Indica se a data representada pelo texto no formato MM/YYYY é uma data válida
        /// </summary>
        /// <param name="data">
        /// Texto representando uma data no formato MM/YYYY
        /// </param>
        /// <returns>
        /// true: data válida
        /// false: data inválida
        /// </returns>
        public static bool isDataMMYYYYOk(String data)
        {
            bool blnDataOk;
            string strFormato;
            DateTime dtDataHoraResp;
            CultureInfo myCultureInfo = new CultureInfo("pt-BR");
            data = formataDataDigitadaParaMMYYYYComSeparador(data);
            if (data.Length != 7) return false;
            strFormato = Cte.DataHora.FmtMes + "/" + Cte.DataHora.FmtAno;
            blnDataOk = DateTime.TryParseExact(data, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp);
            return blnDataOk;
        }
        #endregion

        #region [ isDataOk ]
        /// <summary>
        /// Indica se a data representada pelo texto no formato DD/MM/YYYY é uma data válida
        /// </summary>
        /// <param name="data">
        /// Texto representando uma data no formato DD/MM/YYYY
        /// </param>
        /// <returns>
        /// true: data válida
        /// false: data inválida
        /// </returns>
        public static bool isDataOk(String data)
        {
            bool blnDataOk;
            string strFormato;
            DateTime dtDataHoraResp;
            CultureInfo myCultureInfo = new CultureInfo("pt-BR");
            data = formataDataDigitadaParaDDMMYYYYComSeparador(data);
            if (data.Length != 10) return false;
            strFormato = Cte.DataHora.FmtDia +
                         "/" +
                         Cte.DataHora.FmtMes +
                         "/" +
                         Cte.DataHora.FmtAno;
            blnDataOk = DateTime.TryParseExact(data, strFormato, myCultureInfo, DateTimeStyles.NoCurrentDateDefault, out dtDataHoraResp);
            return blnDataOk;
        }
        #endregion

        #region[ isDigit ]
        public static bool isDigit(char c)
        {
            if ((c >= '0') && (c <= '9')) return true;
            return false;
        }
        #endregion

        #region [ isEmailOk ]
        /// <summary>
        /// Indica se o e-mail possui sintaxe válida. Se for uma lista de e-mails, testa cada um dos e-mails.
        /// </summary>
        /// <param name="email">
        /// Um ou mais e-mails que devem ser analisados. Os e-mails podem ser separados por espaço em branco,
        /// vírgula ou ponto e vírgula.
        /// </param>
        /// <param name="relacaoEmailInvalido">
        /// Informa os e-mails inválidos separados por espaço em branco.
        /// </param>
        /// <returns>
        /// true: todos os e-mails são válidos
        /// false: um ou mais e-mails inválidos
        /// </returns>
        public static bool isEmailOk(String email, ref String relacaoEmailInvalido)
        {
            string strRegExEmailValidacao = "^([0-9a-zA-Z]([-.\\w]*[0-9a-zA-Z][_]*)*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";
            bool blnSucesso;
            int intQtdeEmail = 0;
            String[] v;
            String strEmail;
            Regex rgex = new Regex(strRegExEmailValidacao);

            relacaoEmailInvalido = "";
            if (email == null) return false;
            if (email.Trim().Length == 0) return false;

            blnSucesso = true;
            strEmail = email.Trim();
            strEmail = strEmail.Replace(',', ' ');
            strEmail = strEmail.Replace(';', ' ');
            strEmail = strEmail.Replace("\n", " ");
            strEmail = strEmail.Replace("\r", " ");
            v = strEmail.Split(' ');
            for (int i = 0; i < v.Length; i++)
            {
                if (v[i].Trim().Length > 0)
                {
                    intQtdeEmail++;
                    if (!rgex.IsMatch(v[i].Trim()))
                    {
                        if (relacaoEmailInvalido.Length > 0) relacaoEmailInvalido += " ";
                        relacaoEmailInvalido += v[i];
                        blnSucesso = false;
                    }
                }
            }
            if (intQtdeEmail <= 0) return false;
            return blnSucesso;
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

        #region [ isNumeroPedido ]
        public static bool isNumeroPedido(String numeroPedido)
        {
            String strParteNumerica;
            if (numeroPedido == null) return false;
            if (numeroPedido.Trim().Length == 0) return false;

            strParteNumerica = digitos(Texto.leftStr(numeroPedido, Cte.Etc.TAM_MIN_NUM_PEDIDO));
            if (strParteNumerica.Length != Cte.Etc.TAM_MIN_NUM_PEDIDO) return false;
            if (!isLetra(numeroPedido.Substring(Cte.Etc.TAM_MIN_NUM_PEDIDO, 1))) return false;
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

        #region [ isTeclaEspecialCopiarValorPadrao ]
        /// <summary>
        /// Na edição do grid que cadastra/edita lançamentos de fluxo de caixa em lote,
        /// o preenchimento das células pode ser feito através da cópia dos dados dos
        /// campos que contêm os valores padrão. A cópia do valor padrão é acionada
        /// através de teclas ou combinações de teclas específicas quando a célula está
        /// selecionada.
        /// </summary>
        /// <param name="e">
        /// Objeto "KeyEventArgs" oriundo do evento KeyDown
        /// </param>
        /// <returns>
        /// True: foi pressionada a tecla ou combinação de teclas que aciona a cópia do valor padrão.
        /// False: não foi pressionada a tecla ou combinação de teclas que aciona a cópia do valor padrão.
        /// </returns>
        public static bool isTeclaEspecialCopiarValorPadrao(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space) return true;
            if (e.Shift && (e.KeyCode == Keys.Space)) return true;
            if (e.Shift && (e.KeyCode == Keys.Enter)) return true;
            if (e.Control && (e.KeyCode == Keys.Enter)) return true;
            return false;
        }
        #endregion

        #region [ isTelefoneOk ]
        public static bool isTelefoneOk(String telefone)
        {
            String strTelefone;
            if (telefone == null) return false;
            strTelefone = digitos(telefone);
            if (strTelefone.Length < 7) return false;
            return true;
        }
        #endregion

        #region [ isUfOk ]
        public static bool isUfOk(String uf)
        {
            String strListaUf = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO";
            String strUf;
            String[] v;
            if (uf == null) return false;
            strUf = uf.Trim().ToUpper();
            if (strUf.Length != 2) return false;
            v = strListaUf.Split(' ');
            for (int i = 0; i < v.Length; i++)
            {
                if (strUf.Equals(v[i].Trim()))
                {
                    return true;
                }
            }
            return false;
        }
        #endregion

        #region [ isVScrollBarVisible ]
        /// <summary>
        /// Indica se o Vertical Scroll Bar de um componente (ex: DataGridView) está visível
        /// </summary>
        /// <param name="control">
        /// Objeto que contém o scroll bar
        /// </param>
        /// <returns>
        /// true: o vertical scroll bar está visível
        /// false: o vertical scroll bar não está visível
        /// </returns>
        public static bool isVScrollBarVisible(Control control)
        {
            foreach (Control c in control.Controls)
            {
                if (c.GetType().Equals(typeof(VScrollBar))) return c.Visible;
            }
            return false;
        }
        #endregion

        #region [ montaDescricaoOcorrenciaBoleto ]
        public static String montaDescricaoOcorrenciaBoleto(String identificacaoOcorrencia, String motivosRejeicoes, String motivoOcorrencia19)
        {
            #region [ Declarações ]
            String strResposta;
            String strDescricaoMotivoOcorrencia = "";
            List<Global.TipoDescricaoMotivoOcorrencia> listaMotivoOcorrencia;
            #endregion

            if (identificacaoOcorrencia.Equals("19"))
            {
                strDescricaoMotivoOcorrencia = motivoOcorrencia19 + " - " + decodificaMotivoOcorrencia19(motivoOcorrencia19);
            }
            else
            {
                listaMotivoOcorrencia = decodificaMotivoOcorrencia(identificacaoOcorrencia, motivosRejeicoes);
                for (int j = 0; j < listaMotivoOcorrencia.Count; j++)
                {
                    if (listaMotivoOcorrencia[j].descricaoMotivoOcorrencia.Length > 0)
                    {
                        if (strDescricaoMotivoOcorrencia.Length > 0) strDescricaoMotivoOcorrencia += "\n";
                        strDescricaoMotivoOcorrencia += "(" + listaMotivoOcorrencia[j].motivoOcorrencia + " - " + listaMotivoOcorrencia[j].descricaoMotivoOcorrencia + ")";
                    }
                }
            }

            strResposta = identificacaoOcorrencia;
            if (strResposta.Length > 0) strResposta += " - " + decodificaIdentificacaoOcorrencia(identificacaoOcorrencia);
            if ((strResposta.Length > 0) && (strDescricaoMotivoOcorrencia.Length > 0)) strResposta += "\n";
            strResposta += strDescricaoMotivoOcorrencia;
            return strResposta;
        }
        #endregion

        #region [ montaLinhaDigitavelECodigoBarrasBradesco ]
        /// <summary>
        /// Calcula os dígitos do código de barras e da linha digitável.
        /// </summary>
        /// <param name="numeroBancoCedente">Identificação do banco</param>
        /// <param name="agenciaCedente">Agência cedente sem o dígito verificador</param>
        /// <param name="contaCorrenteCedente">Conta do cedente sem o dígito verificador</param>
        /// <param name="carteira">Código da carteira</param>
        /// <param name="nossoNumero">Nosso número sem o dígito verificador (11 posições)</param>
        /// <param name="dtVencto">Data do vencimento da parcela</param>
        /// <param name="valorParcela">Valor da parcela</param>
        /// <param name="codigoBarras">Retorna os dígitos do código de barras</param>
        /// <param name="linhaDigitavel">Retorna a linha digitável formatada</param>
        /// <returns></returns>
        public static bool montaLinhaDigitavelECodigoBarrasBradesco(
                                            String numeroBancoCedente,
                                            String agenciaCedente,
                                            String contaCorrenteCedente,
                                            String carteira,
                                            String nossoNumero,
                                            DateTime dtVencto,
                                            decimal valorParcela,
                                            ref String codigoBarras,
                                            ref String linhaDigitavel
                                            )
        {
            #region [ Declarações ]
            int intFatorVencto;
            String strFatorVencto;
            String strCampoLivre = "";
            String strBarra1 = "";
            String strBarra2 = "";
            String strBarra3 = "";
            String strBarra4 = "";
            String strValorParcela;
            String strBarraAux;
            String strCodigoBarrasDV;
            String strCodigoBarras;
            String strLinhaDigitavelCampo1 = "";
            String strLinhaDigitavelCampo2 = "";
            String strLinhaDigitavelCampo3 = "";
            String strLinhaDigitavelCampo4 = "";
            String strLinhaDigitavelCampo5 = "";
            String strLinhaDigitavel;
            String strLinhaDigitavelCampo1DV;
            String strLinhaDigitavelCampo2DV;
            String strLinhaDigitavelCampo3DV;
            #endregion

            #region [ Inicialização ]
            codigoBarras = "";
            linhaDigitavel = "";
            #endregion

            #region [ Consistências ]
            if (numeroBancoCedente == null) return false;
            if (numeroBancoCedente.Trim().Length == 0) return false;

            if (agenciaCedente == null) return false;
            if (agenciaCedente.Trim().Length == 0) return false;

            if (contaCorrenteCedente == null) return false;
            if (contaCorrenteCedente.Trim().Length == 0) return false;

            if (carteira == null) return false;
            if (carteira.Trim().Length == 0) return false;

            if (nossoNumero == null) return false;
            if (nossoNumero.Trim().Length == 0) return false;

            if (dtVencto == null) return false;
            if (dtVencto == DateTime.MinValue) return false;

            if (valorParcela <= 0) return false;
            #endregion

            #region [ Montagem do campo livre ]
            // As posições do campo livre ficam a critério de cada banco arrecadador, sendo que este
            // é o padrão do Bradesco
            strCampoLivre = Texto.rightStr(agenciaCedente, 4).PadLeft(4, '0');
            strCampoLivre += Texto.rightStr(carteira, 2).PadLeft(2, '0');
            strCampoLivre += Texto.rightStr(nossoNumero, 11).PadLeft(11, '0');
            strCampoLivre += Texto.rightStr(contaCorrenteCedente, 7).PadLeft(7, '0');
            strCampoLivre += '0'; // Fixo
            #endregion

            #region [ Fator de vencimento ]
            intFatorVencto = calculaTimeSpanDias(dtVencto - new DateTime(1997, 10, 7));
            strFatorVencto = intFatorVencto.ToString().PadLeft(4, '0');
            #endregion

            #region [ Valor da parcela ]
            strValorParcela = digitos(formataMoeda(valorParcela)).PadLeft(10, '0');
            #endregion

            #region [ Montagem do código de barras ]
            strBarra1 = Texto.rightStr(numeroBancoCedente, 3).PadLeft(3, '0');
            strBarra1 += '9'; // Real=9, Outras=0

            strBarra2 = strFatorVencto;
            strBarra3 = strValorParcela;
            strBarra4 = strCampoLivre;

            strBarraAux = strBarra1 + strBarra2 + strBarra3 + strBarra4;
            strCodigoBarrasDV = calculaDigitoVerificadorCodigoBarrasBradesco(strBarraAux);

            strCodigoBarras = strBarra1 + strCodigoBarrasDV + strBarra2 + strBarra3 + strBarra4;
            #endregion

            #region [ Montagem da linha digitável ]

            #region [ Campo 1 ]
            strLinhaDigitavelCampo1 = Texto.rightStr(numeroBancoCedente, 3).PadLeft(3, '0');
            strLinhaDigitavelCampo1 += '9'; // Real=9, Outras=0
            strLinhaDigitavelCampo1 += Texto.leftStr(strCampoLivre, 5);
            strLinhaDigitavelCampo1DV = calculaDigitoVerificadorLinhaDigitavelBradesco(strLinhaDigitavelCampo1);
            strLinhaDigitavelCampo1 += strLinhaDigitavelCampo1DV;
            strLinhaDigitavelCampo1 = strLinhaDigitavelCampo1.Insert(5, ".");
            #endregion

            #region [ Campo 2 ]
            strLinhaDigitavelCampo2 = strCampoLivre.Substring(5, 10);
            strLinhaDigitavelCampo2DV = calculaDigitoVerificadorLinhaDigitavelBradesco(strLinhaDigitavelCampo2);
            strLinhaDigitavelCampo2 += strLinhaDigitavelCampo2DV;
            strLinhaDigitavelCampo2 = strLinhaDigitavelCampo2.Insert(5, ".");
            #endregion

            #region [ Campo 3 ]
            strLinhaDigitavelCampo3 = strCampoLivre.Substring(15, 10);
            strLinhaDigitavelCampo3DV = calculaDigitoVerificadorLinhaDigitavelBradesco(strLinhaDigitavelCampo3);
            strLinhaDigitavelCampo3 += strLinhaDigitavelCampo3DV;
            strLinhaDigitavelCampo3 = strLinhaDigitavelCampo3.Insert(5, ".");
            #endregion

            #region [ Campo 4 ]
            strLinhaDigitavelCampo4 = strCodigoBarrasDV;
            #endregion

            #region [ Campo 5 ]
            strLinhaDigitavelCampo5 = strFatorVencto + strValorParcela;
            #endregion

            strLinhaDigitavel = strLinhaDigitavelCampo1 +
                                ' ' +
                                strLinhaDigitavelCampo2 +
                                ' ' +
                                strLinhaDigitavelCampo3 +
                                ' ' +
                                strLinhaDigitavelCampo4 +
                                ' ' +
                                strLinhaDigitavelCampo5;
            #endregion

            codigoBarras = strCodigoBarras;
            linhaDigitavel = strLinhaDigitavel;
            return true;
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

        #region [ obtemDataReferenciaLimitePagamentoEmAtraso ]
        public static DateTime obtemDataReferenciaLimitePagamentoEmAtraso()
        {
            #region [ Declarações ]
            DateTime dtGravacaoArquivoUltArqRetorno;
            DateTime dthrProcessamentoUltArqRetorno;
            String strNomeUltArqRetorno;
            String strUsuarioProcessamentoUltArqRetorno;
            #endregion

            if (!BoletoDAO.boletoArqRetornoObtemUltimaDtGravacaoArquivo(out dtGravacaoArquivoUltArqRetorno, out strNomeUltArqRetorno, out dthrProcessamentoUltArqRetorno, out strUsuarioProcessamentoUltArqRetorno))
            {
				dtGravacaoArquivoUltArqRetorno = DateTime.Today.AddDays(-1);
            }

            if (dtGravacaoArquivoUltArqRetorno == DateTime.MinValue) dtGravacaoArquivoUltArqRetorno = DateTime.Today.AddDays(-1);

            return dtGravacaoArquivoUltArqRetorno;
        }
        #endregion

        #region [ retiraZerosAEsquerda ]
        /// <summary>
        /// Retira os zeros não significativos à esquerda do número.
        /// Exemplos de retorno: "060" -> "60",  "0" -> "0",  "00" -> "0",  "000" -> "0",  "0,00" -> "0,00",  "0.00" -> "0.00",  "-0,50" -> "-0,50",  "-060,00" -> "-60,00",  "-060" -> "-60",  "+060" -> "+60"
        /// </summary>
        /// <param name="numero">Texto expressando um valor numérico inteiro, decimal ou monetário</param>
        /// <returns>Retorna o texto informado no parâmetro sem os zeros não significativos à esquerda do número, se houver algu.</returns>
        public static String retiraZerosAEsquerda(String numero)
        {
            #region [ Declarações ]
            StringBuilder sbResp = new StringBuilder("");
            char c;
            bool blnHaDados = false;
            #endregion

            if (numero == null) return null;
            if (numero.Length == 0) return "";

            for (int i = 0; i < numero.Length; i++)
            {
                c = numero[i];
                if (c == '0')
                {
                    if (blnHaDados)
                    {
                        sbResp.Append(c);
                    }
                    else if (i < (numero.Length - 1))
                    {
                        if (!isDigit(numero[i + 1]))
                        {
                            sbResp.Append(c);
                            blnHaDados = true;
                        }
                    }
                    else if (i == (numero.Length - 1))
                    {
                        //	Se o texto for "0", "00", "000" ... então retorna "0"
                        sbResp.Append(c);
                        blnHaDados = true;
                    }
                }
                else
                {
                    sbResp.Append(c);
                    if ((c != '+') && (c != '-') && (c != ' ')) blnHaDados = true;
                }
            }

            return sbResp.ToString();
        }
        #endregion

        #region [ retornaCorFluxoCaixaNatureza ]
        /// <summary>
        /// Retorna a cor de exibição para o código da natureza da operação no fluxo de caixa
        /// </summary>
        /// <param name="natureza">
        /// Código da natureza da operação no fluxo de caixa
        /// </param>
        /// <returns>
        /// Retorna a cor para o código da natureza da operação no fluxo de caixa
        /// </returns>
        public static Color retornaCorFluxoCaixaNatureza(char natureza)
        {
            Color color;
            switch (natureza)
            {
                case Cte.FIN.Natureza.CREDITO:
                    color = Color.Green;
                    break;
                case Cte.FIN.Natureza.DEBITO:
                    color = Color.Red;
                    break;
                default:
                    color = Color.Black;
                    break;
            }
            return color;
        }
        #endregion

        #region [ retornaDescricaoCtrlPagtoModulo ]
        public static String retornaDescricaoCtrlPagtoModulo(byte ctrlPagtoModulo)
        {
            String strResposta;
            switch (ctrlPagtoModulo)
            {
                case Cte.FIN.CtrlPagtoModulo.BOLETO:
                    strResposta = "Boleto";
                    break;
                case Cte.FIN.CtrlPagtoModulo.CHEQUE:
                    strResposta = "Cheque";
                    break;
                case Cte.FIN.CtrlPagtoModulo.VISA:
                    strResposta = "Cartão Visa";
                    break;
                default:
                    strResposta = "";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoFluxoCaixaNatureza ]
        /// <summary>
        /// Retorna a descrição para o código da natureza da operação no fluxo de caixa
        /// </summary>
        /// <param name="natureza">
        /// Código da natureza da operação no fluxo de caixa
        /// </param>
        /// <returns>
        /// Retorna um texto com a descrição para o código da natureza da operação no fluxo de caixa
        /// </returns>
        public static String retornaDescricaoFluxoCaixaNatureza(char natureza)
        {
            String strResposta;
            switch (natureza)
            {
                case Cte.FIN.Natureza.CREDITO:
                    strResposta = "Crédito";
                    break;
                case Cte.FIN.Natureza.DEBITO:
                    strResposta = "Débito";
                    break;
                default:
                    strResposta = "Natureza desconhecida";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoOpcaoFluxoCaixaPesquisaLancamentoAtrasado ]
        public static String retornaDescricaoOpcaoFluxoCaixaPesquisaLancamentoAtrasado(byte codigo)
        {
            String strResposta;
            switch (codigo)
            {
                case Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.APENAS_ATRASADOS:
                    strResposta = "Apenas Atrasados";
                    break;
                case Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.IGNORAR_ATRASADOS:
                    strResposta = "Ignorar Atrasados";
                    break;
                default:
                    strResposta = "Valor desconhecido";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoFluxoCaixaStConfirmacaoPendente ]
        /// <summary>
        /// Retorna a descrição para o valor do flag "confirmação pendente" da operação no fluxo de caixa
        /// </summary>
        /// <param name="stConfirmacaoPendente">
        /// Valor do flag "confirmação pendente" da operação no fluxo de caixa
        /// </param>
        /// <returns>
        /// Retorna um texto com a descrição para o valor do flag "confirmação pendente" da operação no fluxo de caixa
        /// </returns>
        public static String retornaDescricaoFluxoCaixaStConfirmacaoPendente(byte stConfirmacaoPendente)
        {
            String strResposta;
            switch (stConfirmacaoPendente)
            {
                case Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO:
                    strResposta = "Confirmado";
                    break;
                case Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO:
                    strResposta = "Pendente";
                    break;
                default:
                    strResposta = "Valor desconhecido";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoFluxoCaixaStSemEfeito ]
        /// <summary>
        /// Retorna a descrição para o valor do flag "sem efeito" da operação no fluxo de caixa
        /// </summary>
        /// <param name="stSemEfeito">
        /// Valor do flag "sem efeito" da operação no fluxo de caixa
        /// </param>
        /// <returns>
        /// Retorna um texto com a descrição para o valor do flag "sem efeito" da operação no fluxo de caixa
        /// </returns>
        public static String retornaDescricaoFluxoCaixaStSemEfeito(byte stSemEfeito)
        {
            String strResposta;
            switch (stSemEfeito)
            {
                case Cte.FIN.StSemEfeito.FLAG_DESLIGADO:
                    strResposta = "Válido";
                    break;
                case Cte.FIN.StSemEfeito.FLAG_LIGADO:
                    strResposta = "Cancelado";
                    break;
                default:
                    strResposta = "Valor desconhecido";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoFluxoCaixaCtrlPagtoStatus ]
        public static String retornaDescricaoFluxoCaixaCtrlPagtoStatus(Cte.FIN.eCtrlPagtoStatus item)
        {
            String strResposta;
            switch (item)
            {
                case Cte.FIN.eCtrlPagtoStatus.CONTROLE_MANUAL:
                    strResposta = "Controle Manual";
                    break;
                case Cte.FIN.eCtrlPagtoStatus.CADASTRADO_INICIAL:
                    strResposta = "Cadastrado (status inicial)";
                    break;
                case Cte.FIN.eCtrlPagtoStatus.BOLETO_BAIXADO:
                    strResposta = "Boleto Baixado";
                    break;
                case Cte.FIN.eCtrlPagtoStatus.CHEQUE_DEVOLVIDO:
                    strResposta = "Cheque Devolvido";
                    break;
                case Cte.FIN.eCtrlPagtoStatus.VISA_CANCELADO:
                    strResposta = "Visa Cancelado";
                    break;
                case Cte.FIN.eCtrlPagtoStatus.PAGO:
                    strResposta = "Pago (OK)";
                    break;
                case Cte.FIN.eCtrlPagtoStatus.BOLETO_PAGO_CHEQUE_VINCULADO:
                    strResposta = "Boleto pago com cheque (vinculado)";
                    break;
                default:
                    strResposta = "Opção Desconhecida: " + ((int)item).ToString();
                    break;
            }
            return strResposta;
        }
		#endregion

		#region [ retornaDescricaoMesAbreviado ]
		public static String retornaDescricaoMesAbreviado(int numero_mes)
		{
			#region [ Declarações ]
			String sMes;
			#endregion
			
			sMes = retornaDescricaoMesExtenso(numero_mes);
			sMes = Texto.leftStr(sMes, 3);
			return sMes;
		}
		#endregion

		#region [ retornaDescricaoMesExtenso ]
		public static String retornaDescricaoMesExtenso(int numero_mes)
		{
			#region [ Declarações ]
			String sResp;
			#endregion

			switch (numero_mes)
			{
				case 1:
					sResp = "Janeiro";
					break;
				case 2:
					sResp = "Fevereiro";
					break;
				case 3:
					sResp = "Março";
					break;
				case 4:
					sResp = "Abril";
					break;
				case 5:
					sResp = "Maio";
					break;
				case 6:
					sResp = "Junho";
					break;
				case 7:
					sResp = "Julho";
					break;
				case 8:
					sResp = "Agosto";
					break;
				case 9:
					sResp = "Setembro";
					break;
				case 10:
					sResp = "Outubro";
					break;
				case 11:
					sResp = "Novembro";
					break;
				case 12:
					sResp = "Dezembro";
					break;
				default:
					sResp = "";
					break;
			}

			return sResp;
		}
		#endregion

		#region [ retornaDescricaoOpcaoCobrancaAdmSituacao ]
		public static String retornaDescricaoOpcaoCobrancaAdmSituacao(byte codigo)
        {
            String strResposta;
            switch (codigo)
            {
                case Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_NAO_ALOCADO:
                    strResposta = "Em Atraso - Não Alocado";
                    break;
                case Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_JA_ALOCADO:
                    strResposta = "Em Atraso - Já Alocado";
                    break;
                default:
                    strResposta = "Valor desconhecido";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoOpcaoPedidoComGarantiaIndicador ]
        public static String retornaDescricaoOpcaoPedidoComGarantiaIndicador(byte codigo)
        {
            String strResposta;
            switch (codigo)
            {
                case Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO:
                    strResposta = "Sem Garantia";
                    break;
                case Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM:
                    strResposta = "Com Garantia";
                    break;
                case Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO_DEFINIDO:
                    strResposta = "Não Definido";
                    break;
                default:
                    strResposta = "Valor desconhecido";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoTipoCadastramento ]
        /// <summary>
        /// Retorna a descrição para o tipo de cadastramento do lançamento do fluxo de caixa
        /// </summary>
        /// <param name="tipo">
        /// Código do tipo de cadastramento do lançamento do fluxo de caixa
        /// </param>
        /// <returns>
        /// Retorna um texto com a descrição para o código do tipo de cadastramento do lançamento do fluxo de caixa
        /// </returns>
        public static String retornaDescricaoTipoCadastramento(char tipo)
        {
            String strResposta;
            switch (tipo)
            {
                case 'S':
                    strResposta = "Sistema";
                    break;
                case 'M':
                    strResposta = "Manual";
                    break;
                default:
                    strResposta = "Modo desconhecido";
                    break;
            }
            return strResposta;
        }
        #endregion

        #region [ retornaDescricaoTipoParcelamentoPedido ]
        /// <summary>
        /// Dado um objeto da classe Pedido, monta uma descrição para o tipo de parcelamento cadastrado
        /// </summary>
        /// <param name="pedido">
        /// Objeto da classe Pedido
        /// </param>
        /// <returns>
        /// Retorna um texto descrevendo a forma de pagamento para o tipo de parcelamento cadastrado
        /// </returns>
        public static String retornaDescricaoTipoParcelamentoPedido(Pedido pedido)
        {
            String strResp = "";

            if (pedido.tipo_parcelamento == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
            {
                strResp = "À Vista (" + formaPagtoPedidoDescricao(pedido.av_forma_pagto) + ")";
            }
            else if (pedido.tipo_parcelamento == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
            {
                strResp = "Parcela Única: " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pu_valor) + " (" + formaPagtoPedidoDescricao(pedido.pu_forma_pagto) + ") vencendo após " + pedido.pu_vencto_apos.ToString() + " dias";
            }
            else if (pedido.tipo_parcelamento == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO)
            {
                strResp = "Parcelado no Cartão (internet) em " + pedido.pc_qtde_parcelas.ToString() + " x " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pc_valor_parcela);
            }
			else if (pedido.tipo_parcelamento == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA)
			{
				strResp = "Parcelado no Cartão (maquineta) em " + pedido.pc_maquineta_qtde_parcelas.ToString() + " x " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pc_maquineta_valor_parcela);
			}
			else if (pedido.tipo_parcelamento == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
            {
                strResp = "Entrada: " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pce_entrada_valor) + " (" + formaPagtoPedidoDescricao(pedido.pce_forma_pagto_entrada) + ")" +
                          "\r\n" +
                          "Prestações: " + pedido.pce_prestacao_qtde.ToString() + " x " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pce_prestacao_valor) + " (" + formaPagtoPedidoDescricao(pedido.pce_forma_pagto_prestacao) + ") vencendo a cada " + pedido.pce_prestacao_periodo.ToString() + " dias";
            }
            else if (pedido.tipo_parcelamento == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
            {
                strResp = "1ª Prestação: " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pse_prim_prest_valor) + " (" + formaPagtoPedidoDescricao(pedido.pse_forma_pagto_prim_prest) + ") vencendo após " + pedido.pse_prim_prest_apos.ToString() + " dias" +
                          "\r\n" +
                          "Demais Prestações: " + pedido.pse_demais_prest_qtde.ToString() + " x " + Cte.Etc.SIMBOLO_MONETARIO + " " + formataMoeda(pedido.pse_demais_prest_valor) + " (" + formaPagtoPedidoDescricao(pedido.pse_forma_pagto_demais_prest) + ") vencendo a cada " + pedido.pse_demais_prest_periodo.ToString() + " dias";
            }

            return strResp;
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

        #region [ setBackColorToAppConfig ]
        public static bool setBackColorToAppConfig(string htmlColor)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["backgroundColorPainel"].Value = (htmlColor == null ? "" : htmlColor);
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #region [ sqlFormataDdMmYyyyParaSqlYyyyMmDd ]
        /// <summary>
        /// A partir de um texto representando uma data no formato DD/MM/YYYY, com ou sem separadores, formata para um texto representando a data no formato 'YYYY-MM-DD' que é entendido pelo SQL Server como uma data
        /// </summary>
        /// <param name="dataDdMmYyyy">
        /// Texto representando uma data no formato DD/MM/YYYY, com ou sem separadores
        /// </param>
        /// <returns>
        /// Retorna um texto representando a data no formato 'YYYY-MM-DD' que é entendido pelo SQL Server como uma data
        /// </returns>
        public static String sqlFormataDdMmYyyyParaSqlYyyyMmDd(String dataDdMmYyyy)
        {
            string strData;

            if (dataDdMmYyyy == null) return "NULL";
            if (dataDdMmYyyy.Trim().Length == 0) return "NULL";

            strData = digitos(dataDdMmYyyy);
            if (strData.Length != 8) return "NULL";
            strData = strData.Substring(4, 4) + "-" + strData.Substring(2, 2) + "-" + strData.Substring(0, 2);
            return strData;
        }
        #endregion

        #region [ sqlFormataDecimal ]
        /// <summary>
        /// Dado um número do tipo decimal, formata um texto representando esse número de forma adequada para usá-lo em uma expressão SQL
        /// </summary>
        /// <param name="valor">
        /// Número do tipo decimal que se deseja representar em um texto para ser usado em expressão SQL
        /// </param>
        /// <returns>
        /// Retorna um texto representando o número em um formato adequado para ser usado em expressão SQL
        /// </returns>
        public static String sqlFormataDecimal(decimal valor)
        {
            String strValorFormatado;
            String strSeparadorDecimal = "";
            decimal decNumeroAuxiliar = .5M;
            String strNumeroAuxiliar;

            strNumeroAuxiliar = decNumeroAuxiliar.ToString();

            if (strNumeroAuxiliar.IndexOf(".") > -1)
                strSeparadorDecimal = ".";
            else if (strNumeroAuxiliar.IndexOf(",") > -1)
                strSeparadorDecimal = ",";

            strValorFormatado = valor.ToString();
            if (strSeparadorDecimal.Length > 0)
            {
                strValorFormatado = strValorFormatado.Replace(strSeparadorDecimal, "V");
                strValorFormatado = strValorFormatado.Replace(".", "");
                strValorFormatado = strValorFormatado.Replace(",", "");
                strValorFormatado = strValorFormatado.Replace("V", ".");
            }
            return strValorFormatado;
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

        #region [ sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaAguardandoLiquidacao ]
        /// <summary>
        /// Monta a expressão lógica para ser inserida dentro de um CASE WHEN que resulta em 
        /// 'true' para lançamentos do fluxo de caixa que estejam aguardando liquidação do cheque.
        /// </summary>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <returns>Expressão lógica para ser inserida dentro da cláusula CASE de uma expressão CASE WHEN</returns>
        public static String sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaAguardandoLiquidacao(String nomeAliasTabelaFluxoCaixa)
        {
            String strResp;
            String strAliasTabela = "";

            if ((nomeAliasTabelaFluxoCaixa != null) && (nomeAliasTabelaFluxoCaixa.Length > 0)) strAliasTabela = nomeAliasTabelaFluxoCaixa;
            if ((strAliasTabela.Length > 0) && (!strAliasTabela.EndsWith("."))) strAliasTabela += ".";

            strResp = " (" + strAliasTabela + "st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO.ToString() + ")" +
                  " AND (" + strAliasTabela + "st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO.ToString() + ")" +
                  " AND (" + strAliasTabela + "ctrl_pagto_modulo = " + Global.Cte.FIN.CtrlPagtoModulo.BOLETO.ToString() + ")" +
                  " AND (" + strAliasTabela + "ctrl_pagto_status = " + ((byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_PAGO_CHEQUE_VINCULADO).ToString() + ")";
            return strResp;
        }
        #endregion

        #region [ sqlMontaCaseWhenParaFluxoCaixaAguardandoLiquidacao ]
        /// <summary>
        /// Monta uma expressão CASE WHEN para calcular se um lançamento do fluxo de caixa está aguardando liquidação.
        /// </summary>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <param name="nomeAliasCampoResultado">Nome do alias para o campo que vai conter o resultado do CASE WHEN</param>
        /// <returns>
        /// Expressão CASE WHEN que calcula se um lançamento do fluxo de caixa está aguardando liquidação:
        /// 0 = Não está aguardando liquidação
        /// 1 = Aguardando liquidação
        /// </returns>
        public static String sqlMontaCaseWhenParaFluxoCaixaAguardandoLiquidacao(String nomeAliasTabelaFluxoCaixa, String nomeAliasCampoResultado)
        {
            String strResp;
            String strAliasCampoResultado = "";

            if ((nomeAliasCampoResultado != null) && (nomeAliasCampoResultado.Length > 0)) strAliasCampoResultado = " AS " + nomeAliasCampoResultado;

            strResp = " (CASE WHEN (" + sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaAguardandoLiquidacao(nomeAliasTabelaFluxoCaixa) + ") THEN " + Global.Cte.FIN.StCampoFlag.FLAG_LIGADO.ToString() + " ELSE " + Global.Cte.FIN.StCampoFlag.FLAG_DESLIGADO.ToString() + " END)" + strAliasCampoResultado;
            return strResp;
        }
        #endregion

        #region [ sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaEmAtraso ]
        /// <summary>
        /// Monta a expressão lógica para ser inserida dentro de um CASE WHEN que resulta em 
        /// 'true' para lançamentos do fluxo de caixa que estejam em atraso.
        /// </summary>
        /// <param name="dtReferenciaAtraso">Data usada como referência para se considerar um pagamento como atrasado ou não.</param>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <returns>Expressão lógica para ser inserida dentro da cláusula CASE de uma expressão CASE WHEN</returns>
        public static String sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaEmAtraso(DateTime dtReferenciaAtraso, String nomeAliasTabelaFluxoCaixa)
        {
            String strResp;
            String strAliasTabela = "";

            if ((nomeAliasTabelaFluxoCaixa != null) && (nomeAliasTabelaFluxoCaixa.Length > 0)) strAliasTabela = nomeAliasTabelaFluxoCaixa;
            if ((strAliasTabela.Length > 0) && (!strAliasTabela.EndsWith("."))) strAliasTabela += ".";

            strResp = " (" + strAliasTabela + "st_sem_efeito = " + Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO.ToString() + ")" +
                      " AND (" + strAliasTabela + "st_confirmacao_pendente = " + Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO.ToString() + ")" +
                      " AND (" +
                                "(" +
                                    "(" + strAliasTabela + "ctrl_pagto_modulo <> " + Global.Cte.FIN.CtrlPagtoModulo.BOLETO.ToString() + ")" +
                                ")" +
                                " OR " +
                                "(" +
                                    "(" + strAliasTabela + "ctrl_pagto_modulo = " + Global.Cte.FIN.CtrlPagtoModulo.BOLETO.ToString() + ")" +
                                    " AND " +
                                    "(" + strAliasTabela + "ctrl_pagto_status <> " + ((byte)Global.Cte.FIN.eCtrlPagtoStatus.BOLETO_PAGO_CHEQUE_VINCULADO).ToString() + ")" +
                                ")" +
                            ")" +
                      " AND (Coalesce(Datediff(day, " + strAliasTabela + "dt_competencia, " + sqlMontaDateTimeParaSqlDateTime(dtReferenciaAtraso) + "), 0) >= 0)";
            return strResp;
        }
        #endregion

        #region [ sqlMontaCaseWhenParaFluxoCaixaEmAtraso ]
        /// <summary>
        /// Monta uma expressão CASE WHEN para calcular se um lançamento do fluxo de caixa está em atraso.
        /// </summary>
        /// <param name="dtReferenciaAtraso">Data usada como referência para se considerar um pagamento como atrasado ou não.</param>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <param name="nomeAliasCampoResultado">Nome do alias para o campo que vai conter o resultado do CASE WHEN</param>
        /// <returns>
        /// Expressão CASE WHEN que calcula se um lançamento do fluxo de caixa está em atraso:
        /// 0 = Não está em atraso
        /// 1 = Está em atraso
        /// </returns>
        public static String sqlMontaCaseWhenParaFluxoCaixaEmAtraso(DateTime dtReferenciaAtraso, String nomeAliasTabelaFluxoCaixa, String nomeAliasCampoResultado)
        {
            String strResp;
            String strAliasCampoResultado = "";

            if ((nomeAliasCampoResultado != null) && (nomeAliasCampoResultado.Length > 0)) strAliasCampoResultado = " AS " + nomeAliasCampoResultado;

            strResp = " (CASE WHEN (" + sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaEmAtraso(dtReferenciaAtraso, nomeAliasTabelaFluxoCaixa) + ") THEN " + Global.Cte.FIN.StCampoFlag.FLAG_LIGADO.ToString() + " ELSE " + Global.Cte.FIN.StCampoFlag.FLAG_DESLIGADO.ToString() + " END)" + strAliasCampoResultado;
            return strResp;
        }
        #endregion

        #region [ sqlMontaCaseWhenParaFluxoCaixaCalculaDiasEmAtraso ]
        /// <summary>
        /// Monta uma expressão CASE WHEN para calcular os dias em atraso de um lançamento do fluxo de caixa.
        /// Caso o lançamento não esteja em atraso, será retornado zero.
        /// A expressão é completa, verificando os campos que definem se uma parcela está em atraso.
        /// </summary>
        /// <param name="dtReferenciaAtraso">Data usada como referência para se considerar um pagamento como atrasado ou não.</param>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <param name="nomeAliasCampoResultado">Nome do alias para o campo que vai conter o resultado do CASE WHEN</param>
        /// <returns>
        /// Quantidade de dias em que o lançamento do fluxo de caixa está em atraso.
        /// Zero, caso o lançamento não esteja em atraso.
        /// </returns>
        public static String sqlMontaCaseWhenParaFluxoCaixaCalculaDiasEmAtraso(DateTime dtReferenciaAtraso, String nomeAliasTabelaFluxoCaixa, String nomeAliasCampoResultado)
        {
            String strResp;
            String strAliasTabela = "";
            String strAliasCampoResultado = "";

            if ((nomeAliasTabelaFluxoCaixa != null) && (nomeAliasTabelaFluxoCaixa.Length > 0)) strAliasTabela = nomeAliasTabelaFluxoCaixa;
            if ((strAliasTabela.Length > 0) && (!strAliasTabela.EndsWith("."))) strAliasTabela += ".";
            if ((nomeAliasCampoResultado != null) && (nomeAliasCampoResultado.Length > 0)) strAliasCampoResultado = " AS " + nomeAliasCampoResultado;

            strResp = " (CASE WHEN (" + sqlMontaCaseWhenExpressaoLogicaCaseParaFluxoCaixaEmAtraso(dtReferenciaAtraso, "tFC") + ") THEN Coalesce(Datediff(day, " + strAliasTabela + "dt_competencia, getdate()),0) ELSE 0 END)" + strAliasCampoResultado;
            return strResp;
        }
        #endregion

        #region [ sqlMontaExpressaoCalculaDiasEmAtraso ]
        /// <summary>
        /// Monta uma expressão em SQL para calcular os dias em atraso, mas somente p/ o cálculo de datas, 
        /// sem verificar os demais campos que definem se uma parcela está atrasada.
        /// </summary>
        /// <param name="dtReferenciaAtraso">Data usada como referência para se considerar um pagamento como atrasado ou não.</param>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <param name="nomeAliasCampoResultado">Nome do alias para o campo que vai conter o resultado do CASE WHEN</param>
        /// <returns>
        /// Quantidade de dias em que o lançamento do fluxo de caixa está em atraso.
        /// </returns>
        public static String sqlMontaExpressaoCalculaDiasEmAtraso(DateTime dtReferenciaAtraso, String nomeAliasTabelaFluxoCaixa, String nomeAliasCampoResultado)
        {
            String strResp;
            String strAliasTabela = "";
            String strAliasCampoResultado = "";

            if ((nomeAliasTabelaFluxoCaixa != null) && (nomeAliasTabelaFluxoCaixa.Length > 0)) strAliasTabela = nomeAliasTabelaFluxoCaixa;
            if ((strAliasTabela.Length > 0) && (!strAliasTabela.EndsWith("."))) strAliasTabela += ".";
            if ((nomeAliasCampoResultado != null) && (nomeAliasCampoResultado.Length > 0)) strAliasCampoResultado = " AS " + nomeAliasCampoResultado;

            strResp = " Coalesce(Datediff(day, " + strAliasTabela + "dt_competencia, getdate()),0)" + strAliasCampoResultado;
            return strResp;
        }
        #endregion

        #region [ sqlMontaRestricoesClausulaWhereParaBoletosEmAtraso ]
        /// <summary>
        /// Monta os campos que serão usados em uma cláusula 'Where' para restringir aos registros do
        /// fluxo de caixa referentes às parcelas de boletos que se encontram em atraso (excluindo os
        /// boletos com pagamento pendente, ou seja, pagos em cheque).
        /// Não inclui a palavra-chave 'WHERE' no retorno.
        /// </summary>
        /// <param name="nomeAliasTabelaFluxoCaixa">Nome do alias para a tabela t_FIN_FLUXO_CAIXA</param>
        /// <returns>
        /// String contendo a expressão SQL.
        /// </returns>
        public static String sqlMontaRestricoesClausulaWhereParaBoletosEmAtraso(String nomeAliasTabelaFluxoCaixa)
        {
            #region [ Declarações ]
            String strAux;
            String strWhere = "";
            String strAliasTabela = "";
            DateTime dtReferenciaLimitePagamentoEmAtraso;
            #endregion

            if ((nomeAliasTabelaFluxoCaixa != null) && (nomeAliasTabelaFluxoCaixa.Length > 0)) strAliasTabela = nomeAliasTabelaFluxoCaixa;
            if ((strAliasTabela.Length > 0) && (!strAliasTabela.EndsWith("."))) strAliasTabela += ".";

            // Restringe somente aos atrasados
            dtReferenciaLimitePagamentoEmAtraso = obtemDataReferenciaLimitePagamentoEmAtraso();
            strAux = " (" + strAliasTabela + "dt_competencia <= " + sqlMontaDateTimeParaSqlDateTime(dtReferenciaLimitePagamentoEmAtraso) + ")";
            if (strWhere.Length > 0) strWhere += " AND";
            strWhere += strAux;

            // Restringe somente a boletos e que não estejam aguardando liquidação
            strAux = " (" + strAliasTabela + "ctrl_pagto_modulo = " + Cte.FIN.CtrlPagtoModulo.BOLETO.ToString() + ")" +
                     " AND (" + strAliasTabela + "st_confirmacao_pendente = " + Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO.ToString() + ")" +
                     " AND (" + strAliasTabela + "st_sem_efeito = " + Cte.FIN.StSemEfeito.FLAG_DESLIGADO.ToString() + ")" +
                     " AND (" + strAliasTabela + "ctrl_pagto_status <> " + ((byte)Cte.FIN.eCtrlPagtoStatus.BOLETO_PAGO_CHEQUE_VINCULADO).ToString() + ")";
            if (strWhere.Length > 0) strWhere += " AND";
            strWhere += strAux;

            return strWhere;
        }
        #endregion

        #region [ sqlMontaDdMmYyyyParaSqlDateTime ]
        /// <summary>
        /// A partir de um texto representando uma data no formato DD/MM/YYYY, com ou sem separadores, monta uma expressão SQL para converter para o tipo de dados DataTime do SQL Server
        /// </summary>
        /// <param name="dataDdMmYyyy">
        /// Texto representando uma data no formato DD/MM/YYYY, com ou sem separadores
        /// </param>
        /// <returns>
        /// Retorna uma expressão SQL para converter para o tipo de dados DateTime do SQL Server
        /// </returns>
        public static string sqlMontaDdMmYyyyParaSqlDateTime(String dataDdMmYyyy)
        {
            string strData;

            if (dataDdMmYyyy == null) return "NULL";
            if (dataDdMmYyyy.Trim().Length == 0) return "NULL";

            strData = digitos(dataDdMmYyyy);
            if (strData.Length != 8) return "NULL";
            strData = strData.Substring(4, 4) + "-" + strData.Substring(2, 2) + "-" + strData.Substring(0, 2);
            return "Convert(datetime, '" + strData + "', 120)";
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

        #region[ sqlMontaDateTimeParaYyyyMmDdComSeparador ]
        /// <summary>
        /// Monta a expressão SQL para retornar um campo do tipo datetime como
        /// texto varchar no formato: 2009-01-30
        /// </summary>
        /// <param name="strNomeCampo">
        /// Informa o nome do campo do banco de dados que deve ser do tipo datetime
        /// </param>
        /// <param name="strAlias">
        /// Informa o nome do Alias, caso seja informado uma string vazia, então será usado o nome do próprio campo.
        /// </param>
        /// <returns></returns>
        public static string sqlMontaDateTimeParaYyyyMmDdComSeparador(string strNomeCampo, string strAlias)
        {
            string strResposta;
            if (strAlias.Trim().Length == 0) strAlias = strNomeCampo;
            strResposta = "Coalesce(Convert(varchar(19), " + strNomeCampo + ", 121), '')";
            if (strAlias.Length > 0) strResposta += " AS " + strAlias;
            return strResposta;
        }
        #endregion

        #region[ sqlMontaDateTimeParaYyyyMmDdComSeparador ]
        /// <summary>
        /// Monta a expressão SQL para retornar um campo do tipo datetime como
        /// texto varchar no formato: 2009-01-30
        /// </summary>
        /// <param name="strNomeCampo">
        /// Informa o nome do campo do banco de dados que deve ser do tipo datetime
        /// </param>
        /// <returns></returns>
        public static string sqlMontaDateTimeParaYyyyMmDdComSeparador(string strNomeCampo)
        {
            return sqlMontaDateTimeParaYyyyMmDdComSeparador(strNomeCampo, "");
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
            if (strAlias.Trim().Length == 0) strAlias = strNomeCampo;
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
            String strNomeCampo = "getdate()";
            strResposta = "Coalesce(Convert(varchar(10), " + strNomeCampo + ", 121), '')";
            return strResposta;
        }
        #endregion

        #region [ sqlMontaGetdateSomenteData ]
        /// <summary>
        /// Monta uma expressão para obter a data do Sql Server com data apenas, sem a hora
        /// </summary>
        /// <returns>
        /// Retorna uma expressão para obter a data do Sql Server com data apenas, sem a hora
        /// </returns>
        public static string sqlMontaGetdateSomenteData(string strAlias)
        {
            string strResposta;
            strResposta = sqlMontaGetdateSomenteData();
            if (strAlias.Length > 0) strResposta += " AS " + strAlias;
            return strResposta;
        }
        #endregion

        #region [ sqlMontaPadLeftCampoNumerico ]
        /// <summary>
        /// Monta uma expressão SQL (sintaxe do SQL Server) para realizar a função de PadLeft() em um campo do tipo numérico que será convertido para varchar
        /// </summary>
        /// <param name="nomeCampo">
        /// Nome do campo no banco de dados
        /// </param>
        /// <param name="preenchimento">
        /// Caracter para preenchimento no padding
        /// </param>
        /// <param name="tamanhoCampo">
        /// Tamanho que o texto deve ficar após execução do padding
        /// </param>
        /// <returns>
        /// Expressão SQL (sintaxe do SQL Server) para realizar a função PadLeft()
        /// </returns>
        public static String sqlMontaPadLeftCampoNumerico(String nomeCampo, char preenchimento, int tamanhoCampo)
        {
            String strResp;
            strResp = " Coalesce(Replicate('" + preenchimento + "'," + tamanhoCampo.ToString() + "-Len(Convert(varchar," + nomeCampo + "))), '') + Convert(varchar," + nomeCampo + ")";
            return strResp;
        }
        #endregion

        #region [ sqlMontaPadLeftCampoTexto ]
        /// <summary>
        /// Monta uma expressão SQL (sintaxe do SQL Server) para realizar a função de PadLeft() em um campo do tipo texto
        /// </summary>
        /// <param name="nomeCampo">
        /// Nome do campo no banco de dados
        /// </param>
        /// <param name="preenchimento">
        /// Caracter para preenchimento no padding
        /// </param>
        /// <param name="tamanhoCampo">
        /// Tamanho que o texto deve ficar após execução do padding
        /// </param>
        /// <returns>
        /// Expressão SQL (sintaxe do SQL Server) para realizar a função PadLeft()
        /// </returns>
        public static String sqlMontaPadLeftCampoTexto(String nomeCampo, char preenchimento, int tamanhoCampo)
        {
            String strResp;
            strResp = " Coalesce(Replicate('" + preenchimento + "'," + tamanhoCampo.ToString() + "-Len(" + nomeCampo + ")), '') + " + nomeCampo;
            return strResp;
        }
        #endregion

        #region [ stEntregaPedidoCor ]
        /// <summary>
        /// Obtém a cor de exibição do status de entrega do pedido
        /// </summary>
        /// <param name="status">
        /// Código do status de entrega do pedido
        /// </param>
        /// <returns>
        /// Retorna uma cor para exibição do status de entrega do pedido
        /// </returns>
        public static Color stEntregaPedidoCor(String status)
        {
            Color cor = Color.Black;

            if (status == null) return cor;
            status = status.Trim();
            if (status.Length == 0) return cor;

            if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ESPERAR))
                cor = Color.DeepPink;
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SPLIT_POSSIVEL))
                cor = Color.DarkOrange;
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SEPARAR))
                cor = Color.Maroon;
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR))
                cor = Color.Blue;
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
                cor = Color.Green;
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
                cor = Color.Red;

            return cor;
        }
        #endregion

        #region [ stEntregaPedidoCor ]
        /// <summary>
        /// Obtém a cor de exibição do status de entrega do pedido
        /// </summary>
        /// <param name="status">
        /// Código do status de entrega do pedido
        /// </param>
        /// <param name="qtdeItensDevolvidos">
        /// Quantidade de itens devolvidos que o pedido já teve
        /// </param>
        /// <returns>
        /// Retorna uma cor para exibição do status de entrega do pedido
        /// </returns>
        public static Color stEntregaPedidoCor(String status, int qtdeItensDevolvidos)
        {
            Color cor;

            cor = stEntregaPedidoCor(status);

            if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
            {
                if (qtdeItensDevolvidos > 0) cor = Color.Red;
            }
            return cor;
        }
        #endregion

        #region [ stEntregaPedidoDescricao ]
        /// <summary>
        /// Obtém a descrição do status de entrega do pedido
        /// </summary>
        /// <param name="status">
        /// Código do status de entrega do pedido
        /// </param>
        /// <returns>
        /// Retorna uma descrição do status de entrega do pedido
        /// </returns>
        public static String stEntregaPedidoDescricao(String status)
        {
            String strResp = "";

            if (status == null) return "";
            status = status.Trim();
            if (status.Length == 0) return "";

            if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ESPERAR))
                strResp = "Esperar Mercadoria";
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SPLIT_POSSIVEL))
                strResp = "Split Possível";
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_SEPARAR))
                strResp = "Separar Mercadoria";
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR))
                strResp = "A Entregar";
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
                strResp = "Entregue";
            else if (status.Equals(Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
                strResp = "Cancelado";
            else
                strResp = "Desconhecido (" + status + ")";

            return strResp;
        }
        #endregion

        #region [ stPagtoPedidoCor ]
        /// <summary>
        /// Obtém a cor de exibição do status de pagamento do pedido
        /// </summary>
        /// <param name="status">
        /// Código do status de pagamento do pedido
        /// </param>
        /// <returns>
        /// Retorna uma cor para exibição do status de pagamento do pedido
        /// </returns>
        public static Color stPagtoPedidoCor(String status)
        {
            Color cor = Color.Black;

            if (status == null) return cor;
            status = status.Trim();
            if (status.Length == 0) return cor;

            if (status.Equals(Cte.StPagtoPedido.ST_PAGTO_PAGO))
                cor = Color.Green;
            else if (status.Equals(Cte.StPagtoPedido.ST_PAGTO_NAO_PAGO))
                cor = Color.Red;
            else if (status.Equals(Cte.StPagtoPedido.ST_PAGTO_PARCIAL))
                cor = Color.DeepPink;

            return cor;
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

        #region [ textBoxPosicionaCursorNoFinal ]
        public static void textBoxPosicionaCursorNoFinal(object sender)
        {
            TextBox c;
            c = (System.Windows.Forms.TextBox)sender;
            c.SelectionLength = 0;
            if (c.Text.Length > 0) c.SelectionStart = c.Text.Length;
        }
        #endregion

        #region [ textBoxSelecionaConteudo ]
        public static void textBoxSelecionaConteudo(object sender)
        {
            ((System.Windows.Forms.TextBox)sender).Select(0, ((System.Windows.Forms.TextBox)sender).Text.Length);
        }
        #endregion

        #region [ tipoParcelamentoPedidoDescricao ]
        /// <summary>
        /// Retorna a descrição para o tipo de parcelamento do pedido (à vista, parcela única, parcelado no cartão, parcelado com entrada, parcelado sem entrada)
        /// </summary>
        /// <param name="codigo">
        /// Código do tipo de parcelamento do pedido
        /// </param>
        /// <returns>
        /// Retorna a descrição para o tipo de parcelamento do pedido
        /// </returns>
        public static String tipoParcelamentoPedidoDescricao(short codigo)
        {
            String strResp = "";

            if (codigo == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
                strResp = "À Vista";
            else if (codigo == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
                strResp = "Parcela Única";
            else if (codigo == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO)
                strResp = "Parcelado no Cartão (internet)";
			else if (codigo == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA)
				strResp = "Parcelado no Cartão (maquineta)";
			else if (codigo == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
                strResp = "Parcelado com Entrada";
            else if (codigo == Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
                strResp = "Parcelado sem Entrada";

            return strResp;
        }
        #endregion

        #region [ trataComboBoxKeyDown ]
        /// <summary>
        /// Trata o evento KeyDown de um campo ComboBox
        /// </summary>
        /// <param name="sender">
        /// O próprio parâmetro "sender" do evento "KeyDown"
        /// </param>
        /// <param name="e">
        /// O próprio parâmetro "e" do evento "KeyDown"
        /// </param>
        /// <param name="proximo">
        /// O próximo para o qual deve ser passado o foco no caso de teclar "Enter" no campo atual
        /// </param>
        public static void trataComboBoxKeyDown(object sender, KeyEventArgs e, Control proximo)
        {
            ComboBox cb = null;

            if (sender.GetType() == typeof(ComboBox)) cb = (ComboBox)sender;

            #region [ Enter ]
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (proximo != null) proximo.Focus();
                return;
            }
            #endregion

            #region [ Delete ]
            if (e.KeyCode == Keys.Delete)
            {
                e.SuppressKeyPress = true;
                if (cb != null)
                {
                    if (cb.DroppedDown) cb.DroppedDown = false;
                    cb.SelectedIndex = -1;
                }
                return;
            }
            #endregion
        }
        #endregion

        #region [ trataTextBoxKeyDown ]
        /// <summary>
        /// Trata o evento KeyDown de um campo TextBox
        /// </summary>
        /// <param name="sender">
        /// O próprio parâmetro "sender" do evento "KeyDown"
        /// </param>
        /// <param name="e">
        /// O próprio parâmetro "e" do evento "KeyDown"
        /// </param>
        /// <param name="proximo">
        /// O próximo para o qual deve ser passado o foco no caso de teclar "Enter" no campo atual
        /// </param>
        public static void trataTextBoxKeyDown(object sender, KeyEventArgs e, Control proximo)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                proximo.Focus();
                return;
            }
        }
        #endregion

        #endregion

        #region [ Assembly ]

        public class AssemblyInfo
        {
            #region [ Assembly Attribute Accessors ]

            #region [ AssemblyTitle ]
            public static string AssemblyTitle
            {
                get
                {
                    object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                    if (attributes.Length > 0)
                    {
                        AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                        if (titleAttribute.Title != "")
                        {
                            return titleAttribute.Title;
                        }
                    }
                    return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
                }
            }
            #endregion

            #region [ AssemblyVersion ]
            public static string AssemblyVersion
            {
                get
                {
                    return Assembly.GetExecutingAssembly().GetName().Version.ToString();
                }
            }
            #endregion

            #region [ AssemblyDescription ]
            public static string AssemblyDescription
            {
                get
                {
                    object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                    if (attributes.Length == 0)
                    {
                        return "";
                    }
                    return ((AssemblyDescriptionAttribute)attributes[0]).Description;
                }
            }
            #endregion

            #region [ AssemblyProduct ]
            public static string AssemblyProduct
            {
                get
                {
                    object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                    if (attributes.Length == 0)
                    {
                        return "";
                    }
                    return ((AssemblyProductAttribute)attributes[0]).Product;
                }
            }
            #endregion

            #region [ AssemblyCopyright ]
            public static string AssemblyCopyright
            {
                get
                {
                    object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                    if (attributes.Length == 0)
                    {
                        return "";
                    }
                    return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
                }
            }
            #endregion

            #region [ AssemblyCompany ]
            public static string AssemblyCompany
            {
                get
                {
                    object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                    if (attributes.Length == 0)
                    {
                        return "";
                    }
                    return ((AssemblyCompanyAttribute)attributes[0]).Company;
                }
            }
            #endregion

            #endregion
        }
		#endregion

		#region [ struct: OpcaoFluxoCaixaNatureza ]
		public struct OpcaoFluxoCaixaNatureza
        {
            private char _codigo;
            private String _descricao;
            public OpcaoFluxoCaixaNatureza(char codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public char codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoFluxoCaixaPesquisaLancamentoAtrasado ]
        public struct OpcaoFluxoCaixaPesquisaLancamentoAtrasado
        {
            private byte _codigo;
            private String _descricao;
            public OpcaoFluxoCaixaPesquisaLancamentoAtrasado(byte codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public byte codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoFluxoCaixaStConfirmacaoPendente ]
        public struct OpcaoFluxoCaixaStConfirmacaoPendente
        {
            private byte _codigo;
            private String _descricao;
            public OpcaoFluxoCaixaStConfirmacaoPendente(byte codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public byte codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoFluxoCaixaStSemEfeito ]
        public struct OpcaoFluxoCaixaStSemEfeito
        {
            private byte _codigo;
            private String _descricao;
            public OpcaoFluxoCaixaStSemEfeito(byte codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public byte codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoFluxoCaixaCtrlPagtoStatus ]
        public struct OpcaoFluxoCaixaCtrlPagtoStatus
        {
            private byte _codigo;
            private String _descricao;
            public OpcaoFluxoCaixaCtrlPagtoStatus(byte codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public byte codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoBoletoIdentificacaoOcorrencia ]
        public struct OpcaoBoletoIdentificacaoOcorrencia
        {
            private String _codigo;
            private String _descricao;
            public OpcaoBoletoIdentificacaoOcorrencia(String codigo, String descricao)
            {
                String strDescricao;
                if (codigo.Equals(Cte.Etc.FLAG_NAO_SETADO.ToString()))
                {
                    strDescricao = descricao;
                }
                else
                {
                    strDescricao = codigo + " - " + descricao;
                }
                this._codigo = codigo;
                this._descricao = strDescricao;
            }
            public String codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
		#endregion

		#region [ struct: OpcaoAno ]
		public struct OpcaoAno
		{
			private int _numero;
			private string _descricao;
			public OpcaoAno(int numero, string descricao)
			{
				this._numero = numero;
				this._descricao = descricao;
			}
			public int numero { get { return _numero; } }
			public string descricao { get { return _descricao; } }
		}
		#endregion

		#region [ struct: OpcaoMes ]
		public struct OpcaoMes
		{
			private int _numero;
			private String _nome;
			public OpcaoMes(int numero, String nome)
			{
				this._numero = numero;
				this._nome = nome;
			}
			public int numero { get { return _numero; } }
			public String nome { get { return _nome; } }
		}
		#endregion

		#region [ struct: TipoDescricaoMotivoOcorrencia ]
		public struct TipoDescricaoMotivoOcorrencia
        {
            private String _identificacaoOcorrencia;
            private String _motivoOcorrencia;
            private String _descricaoOcorrencia;
            private String _descricaoMotivoOcorrencia;

            public TipoDescricaoMotivoOcorrencia(String identificacaoOcorrencia, String motivoOcorrencia, String descricaoOcorrencia, String descricaoMotivoOcorrencia)
            {
                this._identificacaoOcorrencia = identificacaoOcorrencia;
                this._motivoOcorrencia = motivoOcorrencia;
                this._descricaoOcorrencia = descricaoOcorrencia;
                this._descricaoMotivoOcorrencia = descricaoMotivoOcorrencia;
            }
            public String identificacaoOcorrencia { get { return _identificacaoOcorrencia; } }
            public String motivoOcorrencia { get { return _motivoOcorrencia; } }
            public String descricaoOcorrencia { get { return _descricaoOcorrencia; } }
            public String descricaoMotivoOcorrencia { get { return _descricaoMotivoOcorrencia; } }
        }
        #endregion

        #region [ struct: OpcaoCobrancaAdmSituacao ]
        public struct OpcaoCobrancaAdmSituacao
        {
            private byte _codigo;
            private String _descricao;
            public OpcaoCobrancaAdmSituacao(byte codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public byte codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoPedidoComGarantiaIndicador ]
        public struct OpcaoPedidoComGarantiaIndicador
        {
            private byte _codigo;
            private String _descricao;
            public OpcaoPedidoComGarantiaIndicador(byte codigo, String descricao)
            {
                this._codigo = codigo;
                this._descricao = descricao;
            }
            public byte codigo { get { return _codigo; } }
            public String descricao { get { return _descricao; } }
        }
        #endregion

        #region [ struct: OpcaoPlanilhaPagamentosMarketplace ]
        public struct OpcaoPlanilhaPagamentosMarketplace
        {
            public byte codigo { get; }
            public string descricao { get; }
            public OpcaoPlanilhaPagamentosMarketplace(byte codigo, string descricao)
            {
                this.codigo = codigo;
                this.descricao = descricao;
            }
        }
        #endregion

        #region [ enum: eOpcaoIncluirItemTodos ]
        public enum eOpcaoIncluirItemTodos : byte
        {
            NAO_INCLUIR = 0,
            INCLUIR = 1,
            INCLUIR_ITEM_EM_BRANCO = 2
        }
        #endregion

        #region [ enum: eTipoAtualizacaoEfetuada ]
        public enum eTipoAtualizacaoEfetuada
        {
            NENHUMA_ALTERACAO_REALIZADA = 0,
            ALTERADO_REGISTRO_JA_EXISTENTE = 1,
            INCLUSAO_NOVO_REGISTRO = 2
        }
        #endregion

        #region [ montaOpcaoFluxoCaixaNatureza ]
        public static OpcaoFluxoCaixaNatureza[] montaOpcaoFluxoCaixaNatureza(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoFluxoCaixaNatureza[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoFluxoCaixaNatureza[] { new OpcaoFluxoCaixaNatureza(' ', "Todas"),
                                                        new OpcaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.CREDITO, Global.retornaDescricaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.CREDITO)),
                                                        new OpcaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.DEBITO, Global.retornaDescricaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.DEBITO))};
            }
            else
            {
                lista = new OpcaoFluxoCaixaNatureza[] { new OpcaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.CREDITO, Global.retornaDescricaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.CREDITO)),
                                                        new OpcaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.DEBITO, Global.retornaDescricaoFluxoCaixaNatureza(Global.Cte.FIN.Natureza.DEBITO))};
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoFluxoCaixaPesquisaLancamentoAtrasado ]
        public static OpcaoFluxoCaixaPesquisaLancamentoAtrasado[] montaOpcaoFluxoCaixaPesquisaLancamentoAtrasado(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoFluxoCaixaPesquisaLancamentoAtrasado[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoFluxoCaixaPesquisaLancamentoAtrasado[] { new OpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.Etc.FLAG_NAO_SETADO, "Todos"),
                                                                          new OpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.APENAS_ATRASADOS, Global.retornaDescricaoOpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.APENAS_ATRASADOS)),
                                                                          new OpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.IGNORAR_ATRASADOS, Global.retornaDescricaoOpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.IGNORAR_ATRASADOS))};
            }
            else
            {
                lista = new OpcaoFluxoCaixaPesquisaLancamentoAtrasado[] { new OpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.APENAS_ATRASADOS, Global.retornaDescricaoOpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.APENAS_ATRASADOS)),
                                                                          new OpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.IGNORAR_ATRASADOS, Global.retornaDescricaoOpcaoFluxoCaixaPesquisaLancamentoAtrasado(Cte.FIN.CodOpcaoFluxoCaixaPesquisaLancamentoAtrasado.IGNORAR_ATRASADOS))};
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoFluxoCaixaStConfirmacaoPendente ]
        public static OpcaoFluxoCaixaStConfirmacaoPendente[] montaOpcaoFluxoCaixaStConfirmacaoPendente(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoFluxoCaixaStConfirmacaoPendente[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoFluxoCaixaStConfirmacaoPendente[] { new OpcaoFluxoCaixaStConfirmacaoPendente(Cte.Etc.FLAG_NAO_SETADO, "Todos"),
                                                                     new OpcaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO, Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO)),
                                                                     new OpcaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO, Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO))};
            }
            else if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
            {
                lista = new OpcaoFluxoCaixaStConfirmacaoPendente[] { new OpcaoFluxoCaixaStConfirmacaoPendente(Cte.Etc.FLAG_NAO_SETADO, ""),
                                                                     new OpcaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO, Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO)),
                                                                     new OpcaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO, Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO))};
            }
            else
            {
                lista = new OpcaoFluxoCaixaStConfirmacaoPendente[] { new OpcaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO, Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_DESLIGADO)),
                                                                     new OpcaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO, Global.retornaDescricaoFluxoCaixaStConfirmacaoPendente(Global.Cte.FIN.StConfirmacaoPendente.FLAG_LIGADO))};
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoFluxoCaixaStSemEfeito ]
        public static OpcaoFluxoCaixaStSemEfeito[] montaOpcaoFluxoCaixaStSemEfeito(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoFluxoCaixaStSemEfeito[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoFluxoCaixaStSemEfeito[] { new OpcaoFluxoCaixaStSemEfeito(Cte.Etc.FLAG_NAO_SETADO, "Todos"),
                                                           new OpcaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO, Global.retornaDescricaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)),
                                                           new OpcaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_LIGADO, Global.retornaDescricaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_LIGADO))};
            }
            else if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
            {
                lista = new OpcaoFluxoCaixaStSemEfeito[] { new OpcaoFluxoCaixaStSemEfeito(Cte.Etc.FLAG_NAO_SETADO, ""),
                                                           new OpcaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO, Global.retornaDescricaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)),
                                                           new OpcaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_LIGADO, Global.retornaDescricaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_LIGADO))};
            }
            else
            {
                lista = new OpcaoFluxoCaixaStSemEfeito[] { new OpcaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO, Global.retornaDescricaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_DESLIGADO)),
                                                           new OpcaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_LIGADO, Global.retornaDescricaoFluxoCaixaStSemEfeito(Global.Cte.FIN.StSemEfeito.FLAG_LIGADO))};
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoFluxoCaixaCtrlPagtoStatus ]
        public static OpcaoFluxoCaixaCtrlPagtoStatus[] montaOpcaoFluxoCaixaCtrlPagtoStatus(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            int intIndice = -1;
            OpcaoFluxoCaixaCtrlPagtoStatus[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoFluxoCaixaCtrlPagtoStatus[Enum.GetValues(typeof(Cte.FIN.eCtrlPagtoStatus)).Length + 1];
                intIndice++;
                lista[intIndice] = new OpcaoFluxoCaixaCtrlPagtoStatus(Cte.Etc.FLAG_NAO_SETADO, "Todos");
            }
            else if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR_ITEM_EM_BRANCO)
            {
                lista = new OpcaoFluxoCaixaCtrlPagtoStatus[Enum.GetValues(typeof(Cte.FIN.eCtrlPagtoStatus)).Length + 1];
                intIndice++;
                lista[intIndice] = new OpcaoFluxoCaixaCtrlPagtoStatus(Cte.Etc.FLAG_NAO_SETADO, "");
            }
            else
            {
                lista = new OpcaoFluxoCaixaCtrlPagtoStatus[Enum.GetValues(typeof(Cte.FIN.eCtrlPagtoStatus)).Length];
            }

            foreach (Cte.FIN.eCtrlPagtoStatus item in Enum.GetValues(typeof(Cte.FIN.eCtrlPagtoStatus)))
            {
                intIndice++;
                lista[intIndice] = new OpcaoFluxoCaixaCtrlPagtoStatus((byte)item, Global.retornaDescricaoFluxoCaixaCtrlPagtoStatus(item));
            }

            return lista;
        }
        #endregion

        #region [ montaOpcaoBoletoIdentificacaoOcorrencia ]
        public static OpcaoBoletoIdentificacaoOcorrencia[] montaOpcaoBoletoIdentificacaoOcorrencia(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoBoletoIdentificacaoOcorrencia[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoBoletoIdentificacaoOcorrencia[]{
                                new OpcaoBoletoIdentificacaoOcorrencia(Cte.Etc.FLAG_NAO_SETADO.ToString(), "Todos"),
                                new OpcaoBoletoIdentificacaoOcorrencia("02", decodificaIdentificacaoOcorrencia("02")),
                                new OpcaoBoletoIdentificacaoOcorrencia("03", decodificaIdentificacaoOcorrencia("03")),
                                new OpcaoBoletoIdentificacaoOcorrencia("06", decodificaIdentificacaoOcorrencia("06")),
                                new OpcaoBoletoIdentificacaoOcorrencia("09", decodificaIdentificacaoOcorrencia("09")),
                                new OpcaoBoletoIdentificacaoOcorrencia("10", decodificaIdentificacaoOcorrencia("10")),
                                new OpcaoBoletoIdentificacaoOcorrencia("11", decodificaIdentificacaoOcorrencia("11")),
                                new OpcaoBoletoIdentificacaoOcorrencia("12", decodificaIdentificacaoOcorrencia("12")),
                                new OpcaoBoletoIdentificacaoOcorrencia("13", decodificaIdentificacaoOcorrencia("13")),
                                new OpcaoBoletoIdentificacaoOcorrencia("14", decodificaIdentificacaoOcorrencia("14")),
                                new OpcaoBoletoIdentificacaoOcorrencia("15", decodificaIdentificacaoOcorrencia("15")),
                                new OpcaoBoletoIdentificacaoOcorrencia("16", decodificaIdentificacaoOcorrencia("16")),
                                new OpcaoBoletoIdentificacaoOcorrencia("17", decodificaIdentificacaoOcorrencia("17")),
                                new OpcaoBoletoIdentificacaoOcorrencia("18", decodificaIdentificacaoOcorrencia("18")),
                                new OpcaoBoletoIdentificacaoOcorrencia("19", decodificaIdentificacaoOcorrencia("19")),
                                new OpcaoBoletoIdentificacaoOcorrencia("20", decodificaIdentificacaoOcorrencia("20")),
                                new OpcaoBoletoIdentificacaoOcorrencia("21", decodificaIdentificacaoOcorrencia("21")),
                                new OpcaoBoletoIdentificacaoOcorrencia("22", decodificaIdentificacaoOcorrencia("22")),
                                new OpcaoBoletoIdentificacaoOcorrencia("23", decodificaIdentificacaoOcorrencia("23")),
                                new OpcaoBoletoIdentificacaoOcorrencia("24", decodificaIdentificacaoOcorrencia("24")),
                                new OpcaoBoletoIdentificacaoOcorrencia("27", decodificaIdentificacaoOcorrencia("27")),
                                new OpcaoBoletoIdentificacaoOcorrencia("28", decodificaIdentificacaoOcorrencia("28")),
                                new OpcaoBoletoIdentificacaoOcorrencia("30", decodificaIdentificacaoOcorrencia("30")),
                                new OpcaoBoletoIdentificacaoOcorrencia("32", decodificaIdentificacaoOcorrencia("32")),
                                new OpcaoBoletoIdentificacaoOcorrencia("33", decodificaIdentificacaoOcorrencia("33")),
                                new OpcaoBoletoIdentificacaoOcorrencia("34", decodificaIdentificacaoOcorrencia("34")),
                                new OpcaoBoletoIdentificacaoOcorrencia("35", decodificaIdentificacaoOcorrencia("35")),
                                new OpcaoBoletoIdentificacaoOcorrencia("40", decodificaIdentificacaoOcorrencia("40")),
                                new OpcaoBoletoIdentificacaoOcorrencia("55", decodificaIdentificacaoOcorrencia("55")),
                                new OpcaoBoletoIdentificacaoOcorrencia("68", decodificaIdentificacaoOcorrencia("68")),
                                new OpcaoBoletoIdentificacaoOcorrencia("69", decodificaIdentificacaoOcorrencia("69"))
                                };
            }
            else
            {
                lista = new OpcaoBoletoIdentificacaoOcorrencia[]{
                                new OpcaoBoletoIdentificacaoOcorrencia("02", decodificaIdentificacaoOcorrencia("02")),
                                new OpcaoBoletoIdentificacaoOcorrencia("03", decodificaIdentificacaoOcorrencia("03")),
                                new OpcaoBoletoIdentificacaoOcorrencia("06", decodificaIdentificacaoOcorrencia("06")),
                                new OpcaoBoletoIdentificacaoOcorrencia("09", decodificaIdentificacaoOcorrencia("09")),
                                new OpcaoBoletoIdentificacaoOcorrencia("10", decodificaIdentificacaoOcorrencia("10")),
                                new OpcaoBoletoIdentificacaoOcorrencia("11", decodificaIdentificacaoOcorrencia("11")),
                                new OpcaoBoletoIdentificacaoOcorrencia("12", decodificaIdentificacaoOcorrencia("12")),
                                new OpcaoBoletoIdentificacaoOcorrencia("13", decodificaIdentificacaoOcorrencia("13")),
                                new OpcaoBoletoIdentificacaoOcorrencia("14", decodificaIdentificacaoOcorrencia("14")),
                                new OpcaoBoletoIdentificacaoOcorrencia("15", decodificaIdentificacaoOcorrencia("15")),
                                new OpcaoBoletoIdentificacaoOcorrencia("16", decodificaIdentificacaoOcorrencia("16")),
                                new OpcaoBoletoIdentificacaoOcorrencia("17", decodificaIdentificacaoOcorrencia("17")),
                                new OpcaoBoletoIdentificacaoOcorrencia("18", decodificaIdentificacaoOcorrencia("18")),
                                new OpcaoBoletoIdentificacaoOcorrencia("19", decodificaIdentificacaoOcorrencia("19")),
                                new OpcaoBoletoIdentificacaoOcorrencia("20", decodificaIdentificacaoOcorrencia("20")),
                                new OpcaoBoletoIdentificacaoOcorrencia("21", decodificaIdentificacaoOcorrencia("21")),
                                new OpcaoBoletoIdentificacaoOcorrencia("22", decodificaIdentificacaoOcorrencia("22")),
                                new OpcaoBoletoIdentificacaoOcorrencia("23", decodificaIdentificacaoOcorrencia("23")),
                                new OpcaoBoletoIdentificacaoOcorrencia("24", decodificaIdentificacaoOcorrencia("24")),
                                new OpcaoBoletoIdentificacaoOcorrencia("27", decodificaIdentificacaoOcorrencia("27")),
                                new OpcaoBoletoIdentificacaoOcorrencia("28", decodificaIdentificacaoOcorrencia("28")),
                                new OpcaoBoletoIdentificacaoOcorrencia("30", decodificaIdentificacaoOcorrencia("30")),
                                new OpcaoBoletoIdentificacaoOcorrencia("32", decodificaIdentificacaoOcorrencia("32")),
                                new OpcaoBoletoIdentificacaoOcorrencia("33", decodificaIdentificacaoOcorrencia("33")),
                                new OpcaoBoletoIdentificacaoOcorrencia("34", decodificaIdentificacaoOcorrencia("34")),
                                new OpcaoBoletoIdentificacaoOcorrencia("35", decodificaIdentificacaoOcorrencia("35")),
                                new OpcaoBoletoIdentificacaoOcorrencia("40", decodificaIdentificacaoOcorrencia("40")),
                                new OpcaoBoletoIdentificacaoOcorrencia("55", decodificaIdentificacaoOcorrencia("55")),
                                new OpcaoBoletoIdentificacaoOcorrencia("68", decodificaIdentificacaoOcorrencia("68")),
                                new OpcaoBoletoIdentificacaoOcorrencia("69", decodificaIdentificacaoOcorrencia("69"))
                                };
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoCobrancaAdmSituacao ]
        public static OpcaoCobrancaAdmSituacao[] montaOpcaoCobrancaAdmSituacao(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoCobrancaAdmSituacao[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoCobrancaAdmSituacao[] { new OpcaoCobrancaAdmSituacao(Cte.Etc.FLAG_NAO_SETADO, "Todos"),
                                                         new OpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_NAO_ALOCADO, Global.retornaDescricaoOpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_NAO_ALOCADO)),
                                                         new OpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_JA_ALOCADO, Global.retornaDescricaoOpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_JA_ALOCADO))};
            }
            else
            {
                lista = new OpcaoCobrancaAdmSituacao[] { new OpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_NAO_ALOCADO, Global.retornaDescricaoOpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_NAO_ALOCADO)),
                                                         new OpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_JA_ALOCADO, Global.retornaDescricaoOpcaoCobrancaAdmSituacao(Cte.FIN.CodOpcaoCobrancaAdmSituacao.EM_ATRASO_JA_ALOCADO))};
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoPedidoComGarantiaIndicador ]
        public static OpcaoPedidoComGarantiaIndicador[] montaOpcaoPedidoComGarantiaIndicador(eOpcaoIncluirItemTodos opcaoIncluir)
        {
            OpcaoPedidoComGarantiaIndicador[] lista;
            if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
            {
                lista = new OpcaoPedidoComGarantiaIndicador[] { new OpcaoPedidoComGarantiaIndicador(Cte.Etc.FLAG_NAO_SETADO, "Todos"),
                                                                new OpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM, Global.retornaDescricaoOpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM)),
                                                                new OpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO, Global.retornaDescricaoOpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO))};
            }
            else
            {
                lista = new OpcaoPedidoComGarantiaIndicador[] { new OpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM, Global.retornaDescricaoOpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.SIM)),
                                                                new OpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO, Global.retornaDescricaoOpcaoPedidoComGarantiaIndicador(Cte.FIN.T_PEDIDO__GARANTIA_INDICADOR_STATUS.NAO))};
            }
            return lista;
        }
        #endregion

        #region [ montaOpcaoPlanilhaPagamentosMarketplace ]
        public static OpcaoPlanilhaPagamentosMarketplace[] montaOpcaoPlanilhaPagamentosMarketplace()
        {
            OpcaoPlanilhaPagamentosMarketplace[] lista;
            lista = new OpcaoPlanilhaPagamentosMarketplace[]
            {
                new OpcaoPlanilhaPagamentosMarketplace(Cte.Marketplace.COD_PLANILHA_PAGAMENTO_B2W, "B2W")
            };
            return lista;
        }
		#endregion

		#region [ montaOpcaoMes ]
		public static OpcaoMes[] montaOpcaoMes(eOpcaoIncluirItemTodos opcaoIncluir)
		{
			OpcaoMes[] lista;
			if (opcaoIncluir == eOpcaoIncluirItemTodos.INCLUIR)
			{
				lista = new OpcaoMes[] { new OpcaoMes(' ', "Todos"),
										new OpcaoMes(1, "Janeiro"),
										new OpcaoMes(2, "Fevereiro"),
										new OpcaoMes(3, "Março"),
										new OpcaoMes(4, "Abril"),
										new OpcaoMes(5, "Maio"),
										new OpcaoMes(6, "Junho"),
										new OpcaoMes(7, "Julho"),
										new OpcaoMes(8, "Agosto"),
										new OpcaoMes(9, "Setembro"),
										new OpcaoMes(10, "Outubro"),
										new OpcaoMes(11, "Novembro"),
										new OpcaoMes(12,"Dezembro")};
			}
			else
			{
				lista = new OpcaoMes[] { new OpcaoMes(1, "Janeiro"),
										new OpcaoMes(2, "Fevereiro"),
										new OpcaoMes(3, "Março"),
										new OpcaoMes(4, "Abril"),
										new OpcaoMes(5, "Maio"),
										new OpcaoMes(6, "Junho"),
										new OpcaoMes(7, "Julho"),
										new OpcaoMes(8, "Agosto"),
										new OpcaoMes(9, "Setembro"),
										new OpcaoMes(10, "Outubro"),
										new OpcaoMes(11, "Novembro"),
										new OpcaoMes(12, "Dezembro")};
			}
			return lista;
		}
		#endregion
	}
}
