Attribute VB_Name = "mdlGlobais"
Public Function FormataValor(valor As String)
    FormataValor = Format(valor, "R$ #,##0.00")
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
    
    ' Permite d�gitos de 0 a 9 (c�digos ASCII 48 a 57)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        ' A tecla pressionada � um n�mero, permite a entrada
    ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyTab Then
        ' Permite teclas de controle
    Else
        ' A tecla pressionada n�o � um n�mero ou uma tecla permitida
        ' Define KeyAscii para 0 para cancelar a entrada do caractere
        ValidaCampoNumerico = 0
    End If
End Function


Public Function ValidaCampoMoeda(txt As TextBox, KeyAscii As Integer)
    ValidaCampoMoeda = KeyAscii
    
    ' Permite d�gitos de 0 a 9 (c�digos ASCII 48 a 57)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        ' A tecla pressionada � um n�mero, permite a entrada
    ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Or KeyAscii = vbKeyTab Then
        ' Permite teclas de controle
    ElseIf KeyAscii = Asc(",") Then
        ' Permite a v�rgula
        ' Verifica se j� existe uma v�rgula no TextBox para evitar m�ltiplas v�rgulas
        If InStr(txt.Text, ",") > 0 Then
            ValidaCampoMoeda = 0 ' Cancela a entrada se j� existir uma v�rgula
        End If
    Else
        ' A tecla pressionada n�o � um n�mero ou uma tecla permitida
        ' Define KeyAscii para 0 para cancelar a entrada do caractere
        ValidaCampoMoeda = 0
    End If
End Function

Public Function PreencherCollectionComJSONManual(strRespostaJSON As String) As Collection
    Dim arrDividas As Collection
    Dim objDivida As Dictionary
    Dim i As Long
    Dim startObject As Long
    Dim endObject As Long
    Dim currentPos As Long

    On Error GoTo TrataErro:

    Set arrDividas = New Collection

    ' Remover os colchetes '[' e ']' externos
    Dim jsonWithoutBrackets As String
    If Left$(strRespostaJSON, 1) = "[" And Right$(strRespostaJSON, 1) = "]" Then
        jsonWithoutBrackets = Mid$(strRespostaJSON, 2, Len(strRespostaJSON) - 2)
    Else
        MsgBox "Formato JSON inv�lido (esperava array).", vbCritical
        Exit Function
    End If

    ' Dividir a string em objetos JSON (separados por '},{')
    Dim arrObjetos As Variant
    arrObjetos = Split(jsonWithoutBrackets, "},{")

    ' Tratar o primeiro e o �ltimo objeto separadamente para remover '{' e '}'
    If UBound(arrObjetos) >= 0 Then
        arrObjetos(0) = Mid$(arrObjetos(0), 2) ' Remove o '{' inicial
        arrObjetos(UBound(arrObjetos)) = Left$(arrObjetos(UBound(arrObjetos)), Len(arrObjetos(UBound(arrObjetos))) - 1) ' Remove o '}' final
    End If

    ' Iterar sobre os objetos JSON
    For i = LBound(arrObjetos) To UBound(arrObjetos)
        Set objDivida = ParseJsonObject(Trim$(arrObjetos(i)))
        If Not objDivida Is Nothing Then
            arrDividas.Add objDivida
        End If
    Next i
    
    Set PreencherCollectionComJSONManual = arrDividas
    
    Set arrDividas = Nothing
    Set objDivida = Nothing

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
            ' Tentar converter para n�mero se poss�vel
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


