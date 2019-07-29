<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>

<%

' ----------------------------------------- Teste ---------------------------------------'
dim o
dim strMsg
dim resultadoCalculo,resultadoDigito,QtdeCedulas,TotalCedula,Teste1,Teste2
dim dadosCalculo
set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )
dim x,i,cont


  
'dadosCalculo = o.DigitoFinal(Cstr(i))
'if o.DigitoFinal(dadosCalculo) <> "Não é possivel efetuar o calculo com letras" then
'dadosCalculo = o.CalculaCedulas(o.DigitoFinal(dadosCalculo),"2#5|5#5|10#5|20#5|50#5|100#1",resultadoCalculo)
'dadosCalculo = resultadoCalculo
'else
'dadosCalculo = o.DigitoFinal(dadosCalculo)
'end if


%>

<html>

<head>
	<title>CENTRAL</title>

<script type="text/javascript">
    
    var dadosCalculo= "00"
    function TesteCalculo() {
        alert(dadosCalculo);
    }
</script>
	</head>

<body>
    <form id="Teste" name="Teste" method="post">
    <div>   	
        <%
dim teste,aux(5),y,total,b,limitador(5)
dim cedulas()
dim qtdeCedula()
redim  preserve cedulas(0)
teste = 0


   limitador(0)= "0"
limitador(1)= "0"
limitador(2)= "0"
limitador(3)= "1"
limitador(4)= "1"
limitador(5)= "0"

for i=10 to 20
    if i=0 then

            end if
    x=0
    y=0
    dadosCalculo = o.DigitoFinal(Cstr(i))
    total=total+ Cdbl(dadosCalculo)
    dadosCalculo = o.CalculaCedulas(dadosCalculo,"2#"& limitador(5) &"|5#"&limitador(4)&"|10#"&limitador(3)&"|20#"&limitador(2)&"|50#"&limitador(1)&"|100#"&limitador(0)&"",resultadoCalculo,TotalCedula, Teste1)
    
    Teste2= Split(Teste1,"|")

    for cont=0 to Ubound(Teste2)
       
          '  redim  preserve cedulas(x)
             limitador(cont)=  cint(Teste2(cont))
        
    next

    b = dadosCalculo
    dadosCalculo = resultadocalculo
    if dadosCalculo = "Não Contem cedulas para este valor" then
            total = total - cint(o.DigitoFinal(Cstr(i)))
     end if 

   'Split(testeTotal,"|")
    teste = Split(dadosCalculo,"|")

    for cont=0 to Ubound(teste)
       if cont mod 2 = 0 then
            redim  preserve cedulas(x)
          '   cedulas(x)=  cint(teste(cont))
             
                      x = x +1
        else
             redim preserve qtdeCedula(y)
   '          redim preserve aux(y)
                qtdeCedula(y) = cint(teste(cont))
               aux(y) = aux(y) + qtdeCedula(y) 
                    y = y+ 1                   
       end if
   next 
    if i=20 then
        if b = false then
              response.write("<h5> Cedulas de 100: " & qtdeCedula(0) & " |Cedulas de 50= "& qtdeCedula(1) & " |Cedulas de 20= "& qtdeCedula(2) & " |Cedulas de 10= "& qtdeCedula(3) &_
            " |Cedulas de 5= "& qtdeCedula(4) & " |Cedulas de 2= "& qtdeCedula(5) & "&nbsp Total: "& o.DigitoFinal(Cstr(i))&"|  Valor Passado |" & i & "</h5>")
        end if
    response.write("<h5> Total de Cedulas utilizadas Cedulas de 100= " & aux(0) & " |Cedulas de 50= "& aux(1) & " |Cedulas de 20= "& aux(2) & " |Cedulas de 10= "& aux(3) &_
            " |Cedulas de 5= "& aux(4) & " |Cedulas de 2= "& aux(5) &" | Soma Do Valor das Notas "  &total &"</h5>"& "</br>")

    else
           if b = false then
           
             response.write("<h5> Cedulas de 100: "& qtdeCedula(0) & " |Cedulas de 50= "& qtdeCedula(1) & " |Cedulas de 20= "& qtdeCedula(2) & " |Cedulas de 10= "& qtdeCedula(3) &_
            " |Cedulas de 5= "& qtdeCedula(4) & " |Cedulas de 2= ")
             if Cint(qtdeCedula(5)) < 0 then
           response.Write("<font color='red'>"& qtdeCedula(5) & "<font color='black'>"& "&nbsp Total: "& o.DigitoFinal(Cstr(i))&"|  Valor Passado |" & i & "</h5>")
           response.write("<h5>"&Teste1&"</h5>")
            else
            response.Write( qtdeCedula(5) &  "&nbsp Total: "& o.DigitoFinal(Cstr(i))&"|  Valor Passado |" & i & "</h5>")
             response.write("<h5>"&Teste1&"</h5>")
            ' response.write("não contem cedulas para este valor")
           end if
            
            else
             '   if b = false then
          '   response.write("não contem cedulas para este valor" & "</br>")
               
    response.write("<h5> Cedulas de 100: " & qtdeCedula(0) & " |Cedulas de 50= "& qtdeCedula(1) & " |Cedulas de 20= "& qtdeCedula(2) & " |Cedulas de 10= "& qtdeCedula(3) &_
            " |Cedulas de 5= "& qtdeCedula(4) & " |Cedulas de 2= "& qtdeCedula(5) & "&nbsp Total: "& o.DigitoFinal(Cstr(i))&"|  Valor Passado |" & i & "</h5>")
             response.write("<h5>"&Teste1&"</h5>")
            end if       
    end if   
next
%>
    
    </div>
    </form>
</body>
</html>
