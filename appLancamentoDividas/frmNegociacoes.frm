VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNegociacoes 
   Caption         =   "Negociações de Dívidas - Paschoalloto"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSimulacao 
      Caption         =   "Simulação"
      Height          =   2655
      Left            =   4920
      TabIndex        =   11
      Top             =   1080
      Width           =   3495
      Begin VB.TextBox txtJuros 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtQtdParcelas 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdSimular 
         Caption         =   "Simular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Taxa Juros (%):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parcelas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1125
      End
   End
   Begin VB.Frame frmDivida 
      Caption         =   "Dívida"
      Height          =   2655
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   4095
      Begin VB.TextBox txtCPF 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtValor 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtVencimento 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor Original:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1710
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1515
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdNegociacoes 
      Height          =   3015
      Left            =   480
      TabIndex        =   14
      Top             =   4200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5318
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   100
      AllowBigSelection=   0   'False
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Simluação de Negociações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "frmNegociacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimular_Click()
    If Not IsNumeric(txtQtdParcelas.Text) Then
        MsgBox "Informe a quantidade de parcelas para realizar a simulação.", vbInformation, "Gestão de Dívidas - Paschoalloto"
        Exit Sub
    End If
    
    If Not IsNumeric(txtJuros.Text) Then
        MsgBox "Informe a taxa de juros para realizar a simulação.", vbInformation, "Gestão de Dívidas - Paschoalloto"
        Exit Sub
    End If
    
    grdNegociacoes.Clear
    SimularNegociacao
End Sub

Private Sub Form_Load()
    txtCPF.Text = gStrCPF
    txtValor.Text = gStrValor
    txtVencimento.Text = gStrDtVencimento
        
    grdNegociacoes.Clear
    GetNegociacoes
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSimular_Click
        Exit Sub
    End If
    
    KeyAscii = ValidaCampoMoeda(txtJuros, KeyAscii)
End Sub

Private Sub txtJuros_LostFocus()
    txtJuros.Text = FormataPorcentagem(txtJuros.Text)
End Sub

Private Sub txtQtdParcelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSimular_Click
        Exit Sub
    End If
    
    KeyAscii = ValidaCampoNumerico(KeyAscii)
End Sub

Sub PreencherGridComJSON(httpResponseJSON As String)

    Dim arrNegociacoes As Collection
    Dim objNegociacao As Dictionary
    
    Set arrNegociacoes = PreencherCollectionComJSONManual(httpResponseJSON)
    
    If (arrNegociacoes.Count > 0) Then
        
        grdNegociacoes.Clear
        grdNegociacoes.FixedRows = 1
        grdNegociacoes.Cols = 4
        grdNegociacoes.Rows = arrNegociacoes.Count + 1
        
        grdNegociacoes.TextMatrix(0, 0) = "Data Negociação"
        grdNegociacoes.TextMatrix(0, 1) = "Qtde Parcelas"
        grdNegociacoes.TextMatrix(0, 2) = "Taxa Juros"
        grdNegociacoes.TextMatrix(0, 3) = "Valor Total"
    
        For i = 1 To arrNegociacoes.Count
            Set objNegociacao = arrNegociacoes.Item(i)
            grdNegociacoes.TextMatrix(i, 0) = FormatDateTime(objNegociacao("data_negociacao"), vbShortDate)
            grdNegociacoes.TextMatrix(i, 1) = objNegociacao("qtd_parcelas")
            grdNegociacoes.TextMatrix(i, 2) = FormataPorcentagem(objNegociacao("taxa_juros"))
            grdNegociacoes.TextMatrix(i, 3) = FormataValor(objNegociacao("valor_total"), False)
        Next i
        
        For i = 0 To grdNegociacoes.Cols - 1
            grdNegociacoes.ColWidth(i) = (grdNegociacoes.Width / grdNegociacoes.Cols - 20)
        Next i
        
    End If
    
End Sub

Sub SimularNegociacao()

    Dim clDivida As New clsDividas
    Dim valorNegociacao As Double
    Dim MsgErro As String
    
    valorNegociacao = clDivida.CalculaNegociacao(CDbl(RemoveMascaraValor(txtValor.Text)), CDbl(RemoveMascaraValor(txtQtdParcelas.Text)), CDbl(RemoveMascaraValor(txtJuros.Text)), MsgErro)
    
    Set clDivida = Nothing
    
    If valorNegociacao = -1 Then
        MsgBox "Houve uma falha ao realizar a simução da negociação: " & vbCrLf & vbCrLf & MsgErro, vbCritical, "Gestão de Dívidas - Paschoalloto"
    Else
        If MsgBox("Valor da Negociação: " & FormataValor(CStr(valorNegociacao), True) & vbCrLf & vbCrLf & "Clique SIM para salvar a simulação no histório e NÃO para fechar esta janela sem registrar a simulação.", vbYesNo + vbQuestion, "Gestão de Dívidas - Paschoalloto") = vbNo Then
            txtJuros.Text = ""
            txtQtdParcelas.Text = ""
            txtQtdParcelas.SetFocus
            Exit Sub
        Else
            PostNegociacao (valorNegociacao)
        End If
    End If

End Sub

Sub GetNegociacoes()
    Dim objHTTP As Object
    Dim strURL As String
        
    On Error GoTo TrataErro:
    
    strURL = ENDPOINT_API & "/Negociacoes?id_divida=" & gIntIdDivida

    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    objHTTP.Open "GET", strURL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send

    If objHTTP.Status = 200 Then
        PreencherGridComJSON (objHTTP.responseText)
    Else
        MsgBox "Erro na comunicação ao realizar pesquisa de Negociações. Status: " & objHTTP.Status & " - " & objHTTP.statusText, "Gestão de Dívidas - Paschoalloto"
    End If
        
TrataErro:
    If Not objHTTP Is Nothing Then
        Set objHTTP = Nothing
    End If
End Sub

Sub PostNegociacao(valorNegociacao As Double)

    Dim objHTTP As Object
    Dim strURL As String
    Dim strBody As String
        
    On Error GoTo TrataErro:
    
    strURL = ENDPOINT_API & "/Negociacoes"

    strBody = "{" & vbCrLf & _
              "  ""idDivida"": " & gIntIdDivida & "," & vbCrLf & _
              "  ""qtdParcelas"": " & txtQtdParcelas.Text & "," & vbCrLf & _
              "  ""taxaJuros"": " & FormataValorJSON(txtJuros.Text) & "," & vbCrLf & _
              "  ""valorTotal"": " & FormataValorJSON(CStr(valorNegociacao)) & "," & vbCrLf & _
              "  ""dataNegociacao"": """ & Format(Now, "yyyy-mm-dd") & """" & vbCrLf & _
              "}"
              
    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    objHTTP.Open "POST", strURL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send strBody

    If objHTTP.Status = 201 Then
        MsgBox "Simulação salva no histórico com sucesso.", vbInformation, "Gestão de Dívidas - Paschoalloto"
    Else
        MsgBox "Erro na comunicação ao realizar pesquisa de Negociações. Status: " & objHTTP.Status & " - " & objHTTP.statusText, vbCritical, "Gestão de Dívidas - Paschoalloto"
    End If
    
TrataErro:
    If Not objHTTP Is Nothing Then
        Set objHTTP = Nothing
    End If
End Sub
