<%
'*****************************************************
'RETORNA TODOS OS CAMPOS RECEBIDOS PELA PÁGINA
'*****************************************************
Sub ShowRequest(ByVal Method) 
    'Method = POST / GET
    dIM item
    Select Case Ucase(Method)
        Case "POST"
            For Each item In Request.Form
               Response.Write "<br>" & item & ": " & Request.Form(item)
            Next
        Case "GET"
            For Each item In Request.QueryString
                 Response.Write "<br>" & item & ": " & Request.QueryString(item)
            Next
    End Select
    response.write "<BR><BR><BR>"
End Sub

'*****************************************************
' CLASSICO IIF DO .NET
'*****************************************************
Function IIf(ByVal Condicao, ByVal Valor1, ByVal Valor2) 
    If Condicao Then IIf = Valor1 Else IIf = Valor2
End Function
   
'*****************************************************
' RETORNA DATAS NO FORMATO DD/MM/YYYY
'*****************************************************   
Function RetDtBr(ByVal Data) 
    if(Data <> "") Then
		RetDtBr = right("0" & cstr(day(Data)),2) & "/" & right("0" & cstr(month(Data)),2) & "/" & cstr(year(Data)) 
    Else
		RetDtBr = ""
    End if
End Function

'*****************************************************
' REMOVE OS PONTOS SEPARADORES DE MILHARES E SUBSTITUI AS VIRGULAS DECIMAIS POR PONTO
'*****************************************************
Function float2sql(ByVal Num) 
    if Num = "" Then
        float2sql = 0
    else
        float2sql = REPLACE(REPLACE(Cstr(Num),".",""),",",".")
    end if
End Function
    
'*****************************************************
' RETORNA NOME DO MÊS EM PORTUGUÊS
'*****************************************************	
Function NomeMes(ByVal Num As Integer) As String
  Dim meses As Variant
  meses = Array("Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", _
"Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
  If Num >= 1 And Num <= 12 Then
     NomeMes = meses(Num - 1)
   Else
      NomeMes = "Inválido"
   End If
End Function

Function CircNum(ByVal x)
	CircNum = "&#" & iif(x=0,"9450;",iif(9311 + x > 12991,"12991;+",cstr(9311 + x) & ";"))
End Function

Function Destak(ByVal Str, ByVal StrDestak)
	If Not IsNull(Str) And Len(Str) > 0 And Not IsNull(StrDestak) And Len(StrDestak) > 0 Then
		Destak = Replace(LCase(Str), LCase(StrDestak), "<span style=""background-color: #FFFF00;font-size: inherit;"">" & Cstr(StrDestak) & "</span>")
	Else
		Destak = Str
	End If
End Function


