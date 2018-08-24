VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "CHAMALEONBUTTON.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "GLXPBUTTONZ.OCX"
Begin VB.Form usrEditar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizar/Editar"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   Icon            =   "usrEditar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin glxpbuttonz.UserButtonz btnCopiar 
      Height          =   525
      Left            =   4080
      TabIndex        =   18
      ToolTipText     =   "Copiar Email"
      Top             =   2480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "usrEditar.frx":08CA
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkEditar 
      Caption         =   "Habilitar Edição"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin ChamaleonButton.ChameleonBtn btnAlterar 
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   4725
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Alterar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "usrEditar.frx":11A4
      PICN            =   "usrEditar.frx":11C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn btnCancelar 
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   4725
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "usrEditar.frx":1A9A
      PICN            =   "usrEditar.frx":1AB6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3330
      Width           =   4575
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2565
      Width           =   3975
   End
   Begin VB.TextBox txtCargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1845
      Width           =   2175
   End
   Begin VB.TextBox txtOperadora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1125
      Width           =   2175
   End
   Begin VB.TextBox txtContato 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1845
      Width           =   2295
   End
   Begin VB.TextBox txtEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   4575
   End
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contato"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3045
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   22170
      Left            =   0
      Picture         =   "usrEditar.frx":2390
      Stretch         =   -1  'True
      Top             =   0
      Width           =   44310
   End
End
Attribute VB_Name = "usrEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAlterar_Click()
connectBD
rs.Open "Select * from TBTelefones Where Código = " & txtCodigo, db, 3, 3
Do Until rs.EOF
rs(1) = usrEditar.txtEmpresa.Text
rs(2) = usrEditar.txtContato.Text
rs(3) = usrEditar.txtCargo.Text
rs(4) = usrEditar.txtOperadora.Text
rs(5) = usrEditar.txtNumero.Text
rs(6) = usrEditar.txtEmail.Text
rs(7) = usrEditar.txtObs.Text

rs.Update
rs.MoveNext
Loop
fechaBD
MsgBox "Dados atualizados com sucesso !!!", vbInformation, "Informação"
preencherlistview
Unload Me

End Sub

Private Sub btnCopiar1_Click()
txtUsuario.SelStart = 0
txtUsuario.SelLength = Len(txtUsuario.Text)
txtUsuario.SetFocus
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub btnCopiar2_Click()
txtSenha.SelStart = 0
txtSenha.SelLength = Len(txtSenha.Text)
txtSenha.SetFocus
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub btnCopiar_Click()
txtEmail.SelStart = 0
txtEmail.SelLength = Len(txtEmail.Text)
txtEmail.SetFocus
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub chkEditar_Click()
If chkEditar.Value = 1 Then
txtEmpresa.Locked = False
txtNumero.Locked = False
txtOperadora.Locked = False
txtContato.Locked = False
txtObs.Locked = False
txtCargo.Locked = False
txtEmail.Locked = False
Me.btnAlterar.Visible = True
txtEmpresa.SetFocus
End If
If chkEditar.Value = 0 Then
txtEmpresa.Locked = True
txtNumero.Locked = True
txtOperadora.Locked = True
txtContato.Locked = True
txtCargo.Locked = True
txtObs.Locked = True
txtEmail.Locked = True
Me.btnAlterar.Visible = False
End If
End Sub

Private Sub Form_Resize()
Cancel = 1
End Sub



Private Sub txtSistema_Change()
Dim Pos As Integer
Pos = txtSistema.SelStart
txtSistema.Text = VBA.UCase(txtSistema.Text)
txtSistema.SelStart = Pos
End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub
Private Sub Form_Load()

With usrPrincipal.lstTelefones
    
    txtCodigo.Text = .SelectedItem
    
End With

End Sub




'PROCEDIMENTO PARA PREENCHER A LISTVIEW'
Private Sub preencherlistview()
Dim item As ListItem
usrPrincipal.lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & usrPrincipal.txtEmpresa.Text & "%' AND contato like '" & usrPrincipal.txtContato.Text & "%' AND cargo like '" & usrPrincipal.txtCargo.Text & "%' and operadora like '" & usrPrincipal.txtOperadora.Text & "%' order by CONTATO", db, 3, 3

Do Until rs.EOF
Set item = usrPrincipal.lstTelefones.ListItems.Add(, , rs!Código)
    item.SubItems(1) = "" & rs!Empresa
    item.SubItems(2) = "" & rs!contato
    item.SubItems(3) = "" & rs!cargo
    item.SubItems(4) = "" & rs!operadora
    item.SubItems(5) = "" & rs!numero
    item.SubItems(6) = "" & rs!email
        
    rs.MoveNext
    
Loop
fechaBD

'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW
CONTARITENS
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW NA STATUSBAR
CONTARITENSNALISTA
usrPrincipal.CONTADOR_SELECIONADOS.Caption = "0"
End Sub

'PROCEDIMENTO PARA BUSCAR DADOS  ATRAVÉS DO CÓDIGO'
Private Sub txtCodigo_Change()

connectBD
rs.Open "Select * from TBTelefones Where CÓDIGO = " & txtCodigo.Text, db, 3, 3
Do Until rs.EOF

txtEmpresa.Text = "" & rs(1)
txtContato.Text = "" & rs(2)
txtCargo.Text = "" & rs(3)
txtOperadora.Text = "" & rs(4)
txtNumero.Text = "" & rs(5)
txtEmail.Text = "" & rs(6)
txtObs.Text = "" & rs(7)

rs.MoveNext
Loop
fechaBD

End Sub

Private Sub txtCargo_Change()
Dim Pos As Integer
Pos = txtCargo.SelStart
txtCargo.Text = VBA.UCase(txtCargo.Text)
txtCargo.SelStart = Pos
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtEmail.SetFocus
End Sub

Private Sub txtContato_Change()
Dim Pos As Integer
Pos = txtContato.SelStart
txtContato.Text = VBA.UCase(txtContato.Text)
txtContato.SelStart = Pos
End Sub

Private Sub txtContato_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtCargo.SetFocus
End Sub



Private Sub txtEmail_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtObs.SetFocus
End Sub

Private Sub txtEmpresa_Change()
Dim Pos As Integer
Pos = txtEmpresa.SelStart
txtEmpresa.Text = VBA.UCase(txtEmpresa.Text)
txtEmpresa.SelStart = Pos
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtNumero.SetFocus
End Sub

Private Sub txtNumero_Change()
Dim Pos As Integer
Pos = txtNumero.SelStart
txtNumero.Text = VBA.UCase(txtNumero.Text)
txtNumero.SelStart = Pos
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtOperadora.SetFocus
End Sub

Private Sub txtObs_Change()
Dim Pos As Integer
Pos = txtObs.SelStart
txtObs.Text = VBA.UCase(txtObs.Text)
txtObs.SelStart = Pos
End Sub

Private Sub txtOperadora_Change()
Dim Pos As Integer
Pos = txtOperadora.SelStart
txtOperadora.Text = VBA.UCase(txtOperadora.Text)
txtOperadora.SelStart = Pos
End Sub

Private Sub txtOperadora_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtContato.SetFocus
End Sub

