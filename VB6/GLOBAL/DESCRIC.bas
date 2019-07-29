Attribute VB_Name = "mod_DESCRICAO"
Option Explicit

Type TIPO_CODIGO_X_DESCRICAO
    codigo As String * 16
    descricao As String * 50
    End Type

Function DESCRICAO_retorna(ByRef vetor() As TIPO_CODIGO_X_DESCRICAO, ByVal codigo As String, ByRef cadastrado As Boolean) As String
' __________________________________________________________________________________________
'|
'|  * O VETOR DEVE ESTAR ORDENADO
'|  * PRIMEIRA COLUNA DEVE CONTER O CÓDIGO A SER PESQUISADO
'|  * SEGUNDA COLUNA DEVE CONTER A DESCRIÇÃO
'|
'|  O ALGORITMO DE BUSCA É A "BUSCA BINÁRIA" E ESSE
'|  ALGORITMO EXIGE QUE O VETOR ESTEJA ORDENADO
'|

Dim inf As Integer
Dim sup As Integer
Dim meio As Integer
Dim maior As Integer
Dim menor As Integer
Dim s As String
Dim resp As String

    On Error GoTo DR_ERRO
    
  ' ESTABELECE LIMITES DE COMPARAÇÃO INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function






 ' LAÇO DE COMPARAÇÃO
 ' ~~~~~~~~~~~~~~~~~~
    Do While sup >= inf
    
        meio = (sup + inf) \ 2
        s = Trim$(vetor(meio).codigo)
        
      ' VERIFICA SE OPERAÇÕES SÃO IGUAIS
      ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If (codigo > s) Then
            maior = True
            inf = meio + 1
        ElseIf (codigo < s) Then
            menor = True
            sup = meio - 1
        Else
            maior = False
            menor = False
            resp = Trim$(vetor(meio).descricao)
            End If
    
      ' OPERAÇÕES SÃO DIFERENTES, TESTA PRÓXIMO DIREITO
        If (maior Or menor) Then GoTo DR_PROX
    
          
      ' DIREITO FOI ENCONTRADO
        sup = -1
        
    
DR_PROX:
'=======
        Loop
    
    

    If sup = -1 Then
        cadastrado = True
        DESCRICAO_retorna = resp
    Else
        cadastrado = False
        DESCRICAO_retorna = ""
        End If


Exit Function







'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DR_ERRO:
'=======
    s = CStr(Err) & ": " & Error$(Err)
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function


End Function


Public Sub ORDENA_codigo_X_descricao(ByRef vetor() As TIPO_CODIGO_X_DESCRICAO, ByVal inf As Integer, ByVal sup As Integer)

    If inf > sup Then Exit Sub
    
    QuickSort_codigo_X_descricao vetor(), inf, sup

End Sub


Private Sub QuickSort_codigo_X_descricao(ByRef vetor() As TIPO_CODIGO_X_DESCRICAO, ByVal inf As Integer, ByVal sup As Integer)
' _________________________________________________________________________________________
'|
'|  ALGORITMO DE ORDENAÇÃO QUICKSORT
'|  OBS: ALGORITMO É RECURSIVO
'|

Dim i As Integer
Dim j As Integer
Dim ref As TIPO_CODIGO_X_DESCRICAO
Dim temp As TIPO_CODIGO_X_DESCRICAO


    On Error GoTo QCXD_ERRO




  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' LAÇO DE ORDENAÇÃO
    
    Do
        i = inf
        j = sup
        ref = vetor((inf + sup) \ 2)
        
        Do
            
            Do
                If ref.codigo > vetor(i).codigo Then i = i + 1 Else Exit Do
                Loop

            
            
            Do
                If ref.codigo < vetor(j).codigo Then j = j - 1 Else Exit Do
                Loop



            If i <= j Then
                temp = vetor(i)
                vetor(i) = vetor(j)
                vetor(j) = temp
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j

        
        
        If inf < j Then QuickSort_codigo_X_descricao vetor(), inf, j
        
        inf = i
        
        Loop Until i >= sup




Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
QCXD_ERRO:
'=========
    aviso_erro CStr(Err) & ": " & Error$(Err)
    Exit Sub




End Sub



