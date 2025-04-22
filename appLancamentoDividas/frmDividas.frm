VERSION 5.00
Begin VB.Form frmDividas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Dívidas - Paschoalloto"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVencimento 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   420
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtValor 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "00000000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   420
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
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
      Height          =   420
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdCadastrar 
      BackColor       =   &H8000000E&
      Caption         =   "Cadastrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaskColor       =   &H00FF0000&
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
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
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   1515
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
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1710
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
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Lançamento de Dívidas"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmDividas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastrar_Click()

    Dim StrCPF As String
    Dim dblValor As Double
    Dim strVencimento As String
        
    If Not ValidaForm() Then Exit Sub
        
    StrCPF = txtCPF.Text
    dblValor = CDbl(RemoveMascaraValor(txtValor.Text))
    strVencimento = txtVencimento.Text
    
    Call InserirDivida(StrCPF, dblValor, CDate(strVencimento))
    
End Sub

Sub InserirDivida(Cpf As String, valor As Double, DataVencimento As Date)
    
    Dim strMsgErro As String
    Dim clDividas As New clsDividas
    
    If clDividas.CadastrarDivida(Cpf, valor, DataVencimento, strMsgErro) Then
        MsgBox "Dívida inserida com sucesso!", vbInformation, "Gestão de Dívidas - Paschoalloto"
                    
        txtCPF.Text = ""
        txtValor.Text = ""
        txtVencimento.Text = ""
    Else
        MsgBox "Ocorreu um erro ao inserir a dívida: " & strMsgErro, vbCritical, "Gestão de Dívidas - Paschoalloto"
    End If
    
    Set clDividas = Nothing
        
End Sub

Public Function ValidaForm() As Boolean
    ValidaForm = True
    
    If Not IsNumeric(txtCPF.Text) Or Len(txtCPF.Text) <> 11 Then
        MsgBox "Informe um CPF válido para prosseguir.", vbExclamation, "Gestão de Dívidas - Paschoalloto"
        ValidaForm = False
        Exit Function
    End If
    
    If Not IsNumeric(txtValor.Text) Then
        MsgBox "Informe um valor de dívida válido para prosseguir.", vbExclamation, "Gestão de Dívidas - Paschoalloto"
        ValidaForm = False
        Exit Function
    End If
    
    If Not IsDate(txtVencimento.Text) Then
        MsgBox "Informe uma data válida para prosseguir.", vbExclamation, "Gestão de Dívidas - Paschoalloto"
        ValidaForm = False
        Exit Function
    End If
End Function

Private Sub txtCPF_GotFocus()
    txtCPF.SelStart = 0
    txtCPF.SelLength = Len(txtCPF.Text) + 1
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaCampoNumerico(KeyAscii)
End Sub

Private Sub txtValor_GotFocus()
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text) + 1
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaCampoMoeda(txtValor, KeyAscii)
End Sub

Private Sub txtValor_LostFocus()
    txtValor.Text = FormataValor(txtValor.Text, False)
End Sub

Private Sub txtVencimento_GotFocus()
    txtVencimento.SelStart = 0
    txtVencimento.SelLength = Len(txtVencimento.Text)
End Sub
