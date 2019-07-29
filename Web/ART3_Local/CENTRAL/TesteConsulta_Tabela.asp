<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/funcoes.asp"    -->
<%

' ----------------------------------------- Teste ---------------------------------------'
dim o
dim strMsg
dim resultadoCalculo,resultadoDigito,QtdeCedulas, TotalCedula
dim dadosCalculo
set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )
dim x,i,cont
dim teste,aux(5),y,totalArredondado,totalComissao,dec
dim cedulas()
dim qtdeCedula()

%>

<html>

<head>
	<title>CENTRAL</title>


	</head>

<body>
    <form id="Tbl" name="Tbl" method="post">
    <div>   	
        <%
            

redim  preserve cedulas(0)
teste = 0

'- Instanciar os numeros para o calculo das cedulas 
for i=0 to 300
    x=0
    y=0
'- Vl Arredondado
    dadosCalculo = o.DigitoFinal(Cstr(i))
    totalComissao = totalComissao + Cint(i)
    totalArredondado=totalArredondado+ Cdbl(dadosCalculo)
'- Parametros para o calculo, dadosCalculo sendo o Valor, depois as string com as cedulas e o limitador, apos isto as strings de retorno
    dadosCalculo = o.CalculaCedulas(dadosCalculo,"2#5|5#2|10#2|20#2|50#4|100#3",resultadoCalculo)
    dadosCalculo = resultadocalculo
    if dadosCalculo = "Não Contem cedulas para este valor" then
            totalArredondado = totalArredondado - cint(o.DigitoFinal(Cstr(i)))
     end if 

    teste = Split( dadosCalculo,"|")
    

    for cont=0 to Ubound(teste)
       if cont mod 2 = 0 then
            redim  preserve cedulas(x)
             cedulas(x)=  cint(teste(cont))
             
                      x = x +1
        else
             redim preserve qtdeCedula(y)
   '          redim preserve aux(y)
                qtdeCedula(y) = cint(teste(cont))
               aux(y) = aux(y) + qtdeCedula(y) 
                    y = y+ 1                   
       end if
   next 
   if i=0 then
    response.write("<table border=1 bgcolor='white' width='900px'>"&_
"<tr>"&_
"<td align='right' bgcolor='yellow'>" & "Comissão" & "</td>" &"<td bgcolor='yellow' align='right'>"& "VL Arredondado" &"</td>"&"<td align='right' bgcolor='yellow'>" & "2,00" & "</td>" &"<td bgcolor='yellow' align='right'>"& "5,00" &"</td>"&_
"<td align='right' bgcolor='yellow'>" & "10,00" & "</td>" &"<td bgcolor='yellow' align='right'>"& "20,00" &"</td>"&"<td align='right' bgcolor='yellow'>" & "50,00" & "</td>" &"<td bgcolor='yellow' align='right'>"& "100,00" &"</td>"&_
"</tr>")
   else
            response.Write("<tr>"  &_
"<td  align='right'>" & formata_moeda(cstr(i)) & "</td>" &"<td bgcolor='#aaddbb' align='right'>"& formata_moeda(o.DigitoFinal(Cstr(i))) &"</td>"&"<td align='right'>" & qtdeCedula(5) & "</td>" &"<td bgcolor='#aaddbb' align='right'>"& qtdeCedula(4) &"</td>"&_
"<td align='right'>" & qtdeCedula(3) & "</td>" &"<td bgcolor='#aaddbb' align='right'>"& qtdeCedula(2) &"</td>"&"<td align='right'>" & qtdeCedula(1) & "</td>" &"<td bgcolor='#aaddbb' align='right'>"& qtdeCedula(0) &"</td>"&_
"</tr>")
 if i=300 then
    response.Write("<tr>"&"</tr>"& "<tr>"&"</tr>"& "<tr>"&"</tr>"& "<tr>"&"</tr>"&"<tr>"&"</tr>"&"<tr>"&"</tr>"&_
    
"<tr>"  &_
"<td align='right'  bgcolor='yellow'>" & "Total:  " &  formata_moeda(cstr(totalComissao)) & "</td>" &"<td bgcolor='yellow' align='right'>"& formata_moeda(cstr(totalArredondado)) &"</td>"&"<td align='right' bgcolor='yellow'>" & aux(5) & "</td>" &"<td bgcolor='yellow' align='right'>"& aux(4) &"</td>"&_
"<td align='right' bgcolor='yellow'>" & aux(3) & "</td>" &"<td bgcolor='yellow' align='right'>"& aux(2) &"</td>"&"<td align='right' bgcolor='yellow'>" & aux(1) & "</td>" &"<td bgcolor='yellow' align='right'>"& aux(0) &"</td>"&_
"</tr>")
    end if       
            end if
next
 response.write("</table>")
%>
    
    </div>
    </form>
</body>
</html>