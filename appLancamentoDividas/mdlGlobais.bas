Attribute VB_Name = "mdlGlobais"
Public gIntIdDivida As Integer
Public gStrCPF As String
Public gStrValor As String
Public gStrDtVencimento As String

Public Function RemoveMascaraValor(valor As String)
    Dim strValor As String
    
    strValor = Replace(valor, "R$ ", "")
    strValor = Replace(valor, ".", "")
    
    RemoveMascaraValor = strValor
End Function

Public Function FormataValorJSON(valor As String)
    Dim strValor As String
    
    strValor = Replace(valor, "R$ ", "")
    strValor = Replace(valor, ".", "")
    strValor = Replace(valor, ",", ".")
    
    FormataValorJSON = strValor
End Function

Public Function FormataValor(valor As String, blnMoeda As Boolean)
    If blnMoeda Then
        FormataValor = Format(valor, "R$ #,##0.00")
    Else
        FormataValor = Format(valor, "#,##0.00")
    End If
End Function

Public Function FormataPorcentagem(valor As String)
    FormataPorcentagem = Format(valor, "##0.00")
End Function

Public Function FormataCPF(Cpf As String)
    If Len(Cpf) = 11 And IsNumeric(Cpf) Then
        FormataCPF = Left$(Cpf, 3) & "." & Mid$(Cpf, 4, 3) & "." & Mid$(Cpf, 7, 3) & "-" & Right$(Cpf, 2)
    Else
        FormataCPF = Cpf
    End If
End Function

Public Function ValidaCampoNumerico(KeyAscii As Integer)
    ValidaCampoNumerico = KeyAscii
    
    ' Permite dígitos de 0 a 9 (códigos ASCII 48 a 57)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        ' A tecla pressionada é um número, permite a entrada
    ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyTab Then
        ' Permite teclas de controle
    Else
        ' A tecla pressionada não é um número ou uma tecla permitida
        ' Define KeyAscii para 0 para cancelar a entrada do caractere
        ValidaCampoNumerico = 0
    End If
End Function


Public Function ValidaCampoMoeda(txt As TextBox, KeyAscii As Integer)
    ValidaCampoMoeda = KeyAscii
    
    ' Permite dígitos de 0 a 9 (códigos ASCII 48 a 57)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        ' A tecla pressionada é um número, permite a entrada
    ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyTab Then
        ' Permite teclas de controle
    ElseIf KeyAscii = Asc(",") Then
        ' Permite a vírgula
        ' Verifica se já existe uma vírgula no TextBox para evitar múltiplas vírgulas
        If InStr(txt.Text, ",") > 0 Then
            ValidaCampoMoeda = 0 ' Cancela a entrada se já existir uma vírgula
        End If
    Else
        ' A tecla pressionada não é um número ou uma tecla permitida
        ' Define KeyAscii para 0 para cancelar a entrada do caractere
        ValidaCampoMoeda = 0
    End If
End Function

Public Function PreencherCollectionComJSONManual(strRespostaJSON As String) As Collection
    Dim arrResponse As Collection
    Dim objResponde As Dictionary
    Dim i As Long
    Dim startObject As Long
    Dim endObject As Long
    Dim currentPos As Long

    On Error GoTo TrataErro:

    Set arrResponse = New Collection

    ' Remover os colchetes '[' e ']' externos
    Dim jsonWithoutBrackets As String
    If Left$(strRespostaJSON, 1) = "[" And Right$(strRespostaJSON, 1) = "]" Then
        jsonWithoutBrackets = Mid$(strRespostaJSON, 2, Len(strRespostaJSON) - 2)
    Else
        MsgBox "Formato JSON inválido (esperava array).", vbCritical
        Exit Function
    End If

    ' Dividir a string em objetos JSON (separados por '},{')
    Dim arrObjetos As Variant
    arrObjetos = Split(jsonWithoutBrackets, "},{")

    ' Tratar o primeiro e o último objeto separadamente para remover '{' e '}'
    If UBound(arrObjetos) >= 0 Then
        arrObjetos(0) = Mid$(arrObjetos(0), 2) ' Remove o '{' inicial
        arrObjetos(UBound(arrObjetos)) = Left$(arrObjetos(UBound(arrObjetos)), Len(arrObjetos(UBound(arrObjetos))) - 1) ' Remove o '}' final
    End If

    ' Iterar sobre os objetos JSON
    For i = LBound(arrObjetos) To UBound(arrObjetos)
        Set objResponde = ParseJsonObject(Trim$(arrObjetos(i)))
        If Not objResponde Is Nothing Then
            arrResponse.Add objResponde
        End If
    Next i
    
    Set PreencherCollectionComJSONManual = arrResponse
    
    Set arrResponse = Nothing
    Set objResponde = Nothing

    Exit Function

TrataErro:
    MsgBox "Ocorreu um erro ao processar a resposta JSON manualmente.", vbCritical
End Function

Public Function ParseJsonObject(strObject As String) As Dictionary
    Dim obj As New Dictionary
    Dim arrPairs As Variant
    Dim pair As Variant
    Dim arrKeyValue As Variant
    Dim key As String
    Dim value As Variant

    arrPairs = Split(strObject, ",")
    For Each pair In arrPairs
        arrKeyValue = Split(pair, ":")
        If UBound(arrKeyValue) = 1 Then
            key = Trim$(Replace(arrKeyValue(0), """", ""))
            value = Trim$(Replace(arrKeyValue(1), """", ""))
            ' Tentar converter para número se possível
            If IsNumeric(value) And key <> "cpf" Then
                If InStr(value, ".") > 0 Then
                    value = CDbl(Replace(value, ".", ","))
                Else
                    value = CLng(value)
                End If
            ' Tratar booleanos
            ElseIf LCase$(value) = "true" Then
                value = True
            ElseIf LCase$(value) = "false" Then
                value = False
            ' Tratar null (pode precisar de mais refinamento)
            ElseIf LCase$(value) = "null" Then
                value = Null
            End If
            obj.Add key, value
        End If
    Next pair
    Set ParseJsonObject = obj
End Function


