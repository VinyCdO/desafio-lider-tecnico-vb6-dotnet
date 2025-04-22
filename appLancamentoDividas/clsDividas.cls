VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDividas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CadastrarDivida(Cpf As String, valor As Double, DataVencimento As Date, ByRef MsgErro As String) As Boolean

    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command

    On Error GoTo TrataErro

    Set cnn = New ADODB.Connection
    cnn.ConnectionString = CONN_STRING
    cnn.Open

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdText

    cmd.CommandText = "CALL public.prcdividainsert(?,?,?)"

    cmd.Parameters.Append cmd.CreateParameter("cpf", adVarChar, adParamInput, 11, Cpf)
    cmd.Parameters.Append cmd.CreateParameter("valor", adDouble, adParamInput, , valor)
    cmd.Parameters.Append cmd.CreateParameter("vencimento", adDBDate, adParamInput, , DataVencimento)
    
    cmd.Execute
    
    CadastrarDivida = True
    
LimpaObjetos:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    If Not cnn Is Nothing Then
        If cnn.State = adStateOpen Then
            cnn.Close
        End If
        Set cnn = Nothing
    End If
    Exit Function
    
TrataErro:
    MsgErro = "Ocorreu um erro ao inserir a dívida: " & Err.Description
    CadastrarDivida = False
    GoTo LimpaObjetos
End Function

Public Function CalculaNegociacao(valor As Double, qtdeParcelas As Integer, taxaJuros As Double, ByRef MsgErro As String) As Double

    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    
    On Error GoTo TrataErro

    Set cnn = New ADODB.Connection
    cnn.ConnectionString = CONN_STRING
    cnn.Open

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnn
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "public.fnCalculaJurosComposto"

    cmd.Parameters.Append cmd.CreateParameter("valorInicial", adDouble, adParamInput, , valor)
    cmd.Parameters.Append cmd.CreateParameter("qtdParcelas", adInteger, adParamInput, , qtdeParcelas)
    cmd.Parameters.Append cmd.CreateParameter("taxaJuros", adDouble, adParamInput, , taxaJuros)
    
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        CalculaNegociacao = rs(0)
    Else
        CalculaNegociacao = -1
    End If
    
LimpaObjetos:
    If Not cmd Is Nothing Then
        Set cmd = Nothing
    End If
    If Not cnn Is Nothing Then
        If cnn.State = adStateOpen Then
            cnn.Close
        End If
        Set cnn = Nothing
    End If
    Exit Function
    
TrataErro:
    MsgErro = "Ocorreu um erro ao inserir a dívida: " & Err.Description
    CalculaNegociacao = -1
    GoTo LimpaObjetos
End Function

