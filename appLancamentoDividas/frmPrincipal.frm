VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPrincipal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestão de Dívidas - Paschoalloto"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDividas 
      Bindings        =   "frmPrincipal.frx":0000
      Height          =   3015
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   6375
      _ExtentX        =   11245
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
   Begin MSAdodcLib.Adodc adoDividas 
      Height          =   495
      Left            =   5400
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "postgres"
      Password        =   "admin"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "Lançar Dívidas"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtCPF 
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
      Left            =   1080
      MaxLength       =   11
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "Pesquisar"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Gestão de Dívidas"
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
      TabIndex        =   4
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCadastrar_Click()
    frmDividas.Show 1
End Sub

Private Sub cmdPesquisar_Click()
    grdDividas.Clear
    
    Cpf = Trim(txtCPF.Text)
    
    If Cpf = "" Or IsNumeric(Cpf) = False Or Len(Cpf) <> 11 Then
        MsgBox "Informe CPF válido para realizar a pesquisa.", vbInformation
        Exit Sub
    End If
    
    FazerRequisicaoGetDividas
            
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPesquisar_Click
    End If
    
    KeyAscii = ValidaCampoNumerico(KeyAscii)
End Sub

Sub FazerRequisicaoGetDividas()
    Dim objHTTP As Object
    Dim strURL As String
    
    On Error GoTo TrataErro:
    
    strURL = ENDPOINT_API_DIVIDAS & "?cpf=" & txtCPF.Text

    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    objHTTP.Open "GET", strURL, False
    objHTTP.send

    If objHTTP.Status = 200 Then
        PreencherGridComJSON (objHTTP.responseText)
    Else
        MsgBox "Erro na requisição GET. Status: " & objHTTP.Status & " - " & objHTTP.statusText
    End If
    
    Set objHTTP = Nothing
    
    Exit Sub
    
TrataErro:
    MsgBox "Ocorreu alguma falha ao tentar pesquisar as dívidas para o CPF informado, verifique e tente novamente. Erro: " & Err.Description, vbCritical
End Sub

Sub PreencherGridComJSON(httpResponseJSON As String)

    Dim arrDividas As Collection
    Dim objDivida As Dictionary
    
    Set arrDividas = PreencherCollectionComJSONManual(httpResponseJSON)
    
    If (arrDividas.Count > 0) Then
        
        grdDividas.Clear
        grdDividas.FixedRows = 1
        grdDividas.Cols = 4
        grdDividas.Rows = arrDividas.Count + 1
        grdDividas.ColWidth(0) = 0
        
        grdDividas.TextMatrix(0, 0) = "ID"
        grdDividas.TextMatrix(0, 1) = "CPF"
        grdDividas.TextMatrix(0, 2) = "Valor"
        grdDividas.TextMatrix(0, 3) = "Vencimento"
    
        For i = 1 To arrDividas.Count
            Set objDivida = arrDividas.Item(i)
            grdDividas.TextMatrix(i, 0) = objDivida("id")
            grdDividas.TextMatrix(i, 1) = FormataCPF(objDivida("cpf"))
            grdDividas.TextMatrix(i, 2) = FormataValor(objDivida("valor_original"))
            grdDividas.TextMatrix(i, 3) = FormatDateTime(objDivida("data_vencimento"), vbShortDate)
        Next i
        
        For i = 1 To grdDividas.Cols - 1
            grdDividas.ColWidth(i) = ((grdDividas.Width / (grdDividas.Cols - 1)) - 20)
        Next i
        
    End If
    
    
End Sub

