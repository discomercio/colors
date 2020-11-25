dim o
dim strIE, strUF

set o = createobject( "ComPlusWrapper_DllInscE32.ComPlusWrapper_DllInscE32" )

Wscript.Echo "ComPlusWrapper_DllInscE32 - Versão: " & o.Versao

strIE = "7076456170088"
strUF = "MG"
Wscript.Echo strIE & " (" & strUF & "): " & o.ConsisteInscricaoEstadual(strIE, strUF)
Wscript.Echo strIE & " (" & strUF & "): " & o.isInscricaoEstadualOk(strIE, strUF)


strIE = "ISENTO"
strUF = "MG"
Wscript.Echo strIE & " (" & strUF & "): " & o.ConsisteInscricaoEstadual(strIE, strUF)
Wscript.Echo strIE & " (" & strUF & "): " & o.isInscricaoEstadualOk(strIE, strUF)
