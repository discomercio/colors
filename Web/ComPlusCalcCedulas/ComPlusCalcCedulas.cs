#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
using System.Diagnostics;
using System.Threading;
#endregion

[assembly: ApplicationName("ComPlusCalcCedulas")]
[assembly: Description("Calculo de quantidade de cedulas")]
namespace ComPlusCalcCedulas
{
    [JustInTimeActivation(true)]
    [Transaction(TransactionOption.Required)]
    [GuidAttribute("452C94DB-E6FF-417B-B9A9-31120E6C999F")]
    public class ComPlusCalcCedulas : ServicedComponent
    {
        #region [ Construtor ]
        public ComPlusCalcCedulas()
        {
            // NOP
            /*
             * ATENÇÃO: lembre-se que o código executado no construtor é executado
             * ======== sempre que um novo objeto instanciar este serviço de componente.
             * No caso de inicializar variáveis globais, é fundamental ter em mente
             * que as variáveis do tipo 'static' são únicas, independentemente da instância.
             * Ou seja, uma instância interfere na outra no caso das variáveis globais do
             * tipo 'static', por isso é necessário muito cuidado com elas.
             */
        }
        #endregion

        #region[ Versao ]
        public string Versao()
        {
            return Global.Cte.Versao.strVersao;
        }
        #endregion

        #region[ DataHora ]
        public string DataHora()
        {
            return DateTime.Now.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmSsComSeparador);
        }
        #endregion

        #region [ CalculaCedulas ]
        /// <summary>
        /// Este metodo calcula quantas cedulas é utilizadas em um certo valor passado.
        /// Parametros das cedulas só serão calculados corretamente caso forem passados em ordem crescente.
        /// Primeiro valor da cedula e depois a quantidade do limitador 2#1|5#2
        /// O valor da cedula corresponde ao valor atrás do # e o limitador após, e para quebrar para outra cedula utiliza-se o caracter "|" 
        /// </summary>
        /// <param name="dadosCalculo">Recebe o Valor Para o Calculo em String</param>
        /// <param name="cedula"> Recebe as Cedulas que Serão utilizdas e o limitador delas em ordem crescente ex(2#2|5#2)</param>
        /// <param name="resultadoCalculo">Este parametro ira retornar o resultado em string</param>
        /// <returns>Retorna true ou false : se foi possivel efetuar a conta ou não</returns>
        public bool CalculaCedulas(string dadosCalculo, string cedula, out string resultadoCalculo)
        {
            bool b = true;
            resultadoCalculo = "";
            char[] determinadorChar = { ' ', '|', '#' };
            string[] RecebeCedula = cedula.Split(determinadorChar);
            int[] cedulas = new int[(RecebeCedula.Length / 2)];
            int[] limitadorCedulas = new int[(RecebeCedula.Length / 2)];
            int x = 0;
            int y = 0;
            int[] qtdeCedulas = new int[cedulas.Length];
            int[] qtdeCedulas2 = new int[cedulas.Length];
            int i = cedulas.Length - 1;
            int aux = Int32.Parse(dadosCalculo);
            int somaValorCedulas = 0;
            int aux2 = Int32.Parse(dadosCalculo);
            int cont = 0; 

            //Recebe o Valor De Quais Cedulas Serão utilizadas e Recebe o limitador delas
            for (cont = 0; cont < RecebeCedula.Length; cont++)
            {
                if (cont % 2 == 0)
                {
                    cedulas[x] = Convert.ToInt32(RecebeCedula[cont]);
                    x++;
                }
                else
                {
                    limitadorCedulas[y] = Convert.ToInt32(RecebeCedula[cont]);
                    if(limitadorCedulas[y] < 0)
                    {
                        limitadorCedulas[y] = 0;
                    }
                    y++;
                }
            }

            //Calculo Para verificar qual o valor total das cedulas
            for (int z = cedulas.Length - 1; z >= 0; z--)
            {
                somaValorCedulas = somaValorCedulas + cedulas[z] * cedulas[z];
            }

                //Calcula Quantas Notas Serão utilizadas Da maior para menor nota com o limitador.
                while (i >= 0)
                {
                    if (aux >= cedulas[i])
                    {
                        if ((qtdeCedulas[i] = aux / cedulas[i]) >= limitadorCedulas[i])
                        {                       
                            qtdeCedulas[i] = limitadorCedulas[i];
                            aux = aux - (qtdeCedulas[i] * cedulas[i]);
                            i--;
                        }
                        else
                        {
                            qtdeCedulas[i] = aux / cedulas[i];
                            if (aux % cedulas[i] == 0)
                            {
                                if (qtdeCedulas[i] <= limitadorCedulas[i])
                                {  
                                    aux = aux - (qtdeCedulas[i] * cedulas[i]);
                                    i = -1;
                                }
                                else
                                {
                                    qtdeCedulas[i] = limitadorCedulas[i];                               
                                    aux = aux - (qtdeCedulas[i] * cedulas[i]);
                                    i--;
                                }
                            }
                            else
                            {
                                aux = aux - (qtdeCedulas[i] * cedulas[i]);      
                                i--;
                            }
                        }
                    }
                    else
                    {
                        i--;
                    }
                }          
            // faz o calculo sem limitador para retornar quais cedulas necessitão para o calculo
           
            if(aux != 0)
            {
                b = false;
                cont = cedulas.Length - 1;
                while (cont >= 0)
                {
                    if (aux2 >= cedulas[cont])
                    {
                        qtdeCedulas2[cont] = aux2 / cedulas[cont];
                        if (aux2 % cedulas[cont] == 0)
                        {
                            aux2 = aux2 - (qtdeCedulas2[cont] * cedulas[cont]);
                            cont = -1;
                        }
                        else
                        {
                            aux2 = aux2 - (qtdeCedulas2[cont] * cedulas[cont]);
                            cont--;
                        }
                    }
                    else
                    {
                        cont--;
                    }
                }          
             }        

            //Atribui a quantidade de notas para a mensagem de retorno se o valor for verdadeiro,foi efetuado o calculo com sucesso e retorna as cedulas
            if ((somaValorCedulas > aux) && (aux == 0) && (b==true))
            {
                for (i = cedulas.Length - 1; i >= 0; i--)
                {
                    if (i == 0)
                    {                       
                        resultadoCalculo = resultadoCalculo + cedulas[i] + "|" + (qtdeCedulas[i]);
                   
                    }
                    else
                    {
                        resultadoCalculo = resultadoCalculo + cedulas[i] + "|" + (qtdeCedulas[i]) + "|";
                     
                    }
                }
            }
            // Retorna as cedulas para verificar quais estão faltando para o calculo 
            else
            {
                for (i = cedulas.Length - 1; i >= 0; i--)
                {
                    if (i == 0)
                    {                   
                        resultadoCalculo = resultadoCalculo + cedulas[i] + "|" + (qtdeCedulas2[i]);
                     
                    }
                    else
                    {
                        resultadoCalculo = resultadoCalculo + cedulas[i] + "|" + (qtdeCedulas2[i]) + "|";
                    
                    }
                }
            }            
            return b;
        }
        #endregion

        #region [ DigitoFinal]
        /// <summary>
        /// Está função verifica o ultimo digito do parametro e trata ele conforme as exceções que são:
        /// valor com final 0,2,4,5 continua com  este valor
        /// valor com final 1 passa a ser 0             ex:81 = 80
        /// valor com final 3 passa a ser 2             ex:83 = 82
        /// valor com final 6 ou 7 passa a ser 5        ex:87 = 85
        /// valor com final 8 ou 9 passa a ser 10       ex:89 = 90
        /// </summary>
        /// <param name="dadosCalculo">Recebe o valor para o calculo do ultimo digito em string</param>
        /// <returns>Retorna o valor atualizado pelo calculo</returns>
        public string DigitoFinal(string dadosCalculo)
        {
           
            decimal aux = 0;
            
            //Verifica se é numero.
            try
            {
               //aux = Convert.ToDouble(dadosCalculo);
               aux = Global.converteNumeroDecimal(dadosCalculo); 
               aux = Math.Floor(aux);
               dadosCalculo = Convert.ToString(aux);
               string ultimoDigito = Convert.ToString(dadosCalculo.Substring(dadosCalculo.Length - 1));
               //Verifica Qual o ultimo digito e trata eles.
               switch (ultimoDigito)
               {
                   case "0":
                       break;
                   case "1":
                       aux = aux - 1;
                       dadosCalculo = Convert.ToString(aux);
                       break;
                   case "2":
                       break;
                   case "3":
                       aux = aux - 1;
                       dadosCalculo = Convert.ToString(aux);
                       break;
                   case "4":
                       break;
                   case "5":
                       break;
                   case "6":
                       aux = aux - 1;
                       dadosCalculo = Convert.ToString(aux);
                       break;
                   case "7":
                       aux = aux - 2;
                       dadosCalculo = Convert.ToString(aux);
                       break;
                   case "8":
                       aux = aux + 2;
                       dadosCalculo = Convert.ToString(aux);
                       break;
                   case "9":
                       aux = aux + 1;
                       dadosCalculo = Convert.ToString(aux);
                       break;
               }
            }
            catch(Exception )
            {
                dadosCalculo = "Não é possivel efetuar o calculo com letras";
                
            }

            return dadosCalculo;
        }
        #endregion

    }
}
