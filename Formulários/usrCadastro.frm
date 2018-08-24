VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "CHAMALEONBUTTON.OCX"
Begin VB.Form usrCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   Icon            =   "usrCadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
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
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3285
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
      TabIndex        =   5
      Top             =   2520
      Width           =   4575
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
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin ChamaleonButton.ChameleonBtn btnCadastrar 
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Cadastrar"
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
      MICON           =   "usrCadastro.frx":08CA
      PICN            =   "usrCadastro.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   2
      Top             =   1080
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
      TabIndex        =   3
      Top             =   1800
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
      TabIndex        =   0
      Top             =   360
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
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin ChamaleonButton.ChameleonBtn btnCancelar 
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   4440
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
      MICON           =   "usrCadastro.frx":1978
      PICN            =   "usrCadastro.frx":1994
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   15
      Top             =   3000
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
      TabIndex        =   14
      Top             =   2240
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
      Top             =   1515
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
      TabIndex        =   12
      Top             =   795
      Width           =   1575
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
      TabIndex        =   11
      Top             =   1515
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
      TabIndex        =   10
      Top             =   75
      Width           =   975
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
      TabIndex        =   9
      Top             =   795
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   22170
      Left            =   0
      Picture         =   "usrCadastro.frx":226E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   44310
   End
End
Attribute VB_Name = "usrCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCadastrar_Click()

'VERIFICA SE OS CAMPOS FORAM PREENCHIDOS'
If txtEmpresa.Text = "" Then
MsgBox ("Diga qual o NOME ou NUMERO da empresa a que se refere o registro"), vbExclamation, "Campo Obrigatório"
txtEmpresa.SetFocus
Exit Sub
Else
End If
If txtNumero.Text = "" Then
MsgBox ("Diga qual o NUMERO de telefone ou celular da pessoa a que se refere o registro"), vbExclamation, "Campo Obrigatório"
txtNumero.SetFocus
Exit Sub
Else
End If
If txtOperadora.Text = "" Then
MsgBox ("Diga qual a OPERADORA de telefone ou celular da pessoa a que se refere o registro"), vbExclamation, "Campo Obrigatório"
txtOperadora.SetFocus
Exit Sub
Else
End If
If txtContato.Text = "" Then
MsgBox ("Diga qual a Nome da pessoa refere o registro"), vbExclamation, "Campo Obrigatório"
txtContato.SetFocus
Exit Sub
Else
End If
If txtCargo.Text = "" Then
MsgBox ("Diga qual o CARGO da pessoa refere o registro"), vbExclamation, "Campo Obrigatório"
txtCargo.SetFocus
Exit Sub
Else
End If


connectBD
rs.Open "SELECT *FROM TBTelefones", db, 3, 3
rs.AddNew
rs!Empresa = txtEmpresa.Text
rs!contato = txtContato.Text
rs!cargo = txtCargo.Text
rs!operadora = txtOperadora.Text
rs!numero = txtNumero.Text
rs!email = txtEmail.Text
rs!observaçoes = txtObs.Text
rs.Update
fechaBD
MsgBox "Cadastro Realizado com Sucesso !!!", vbInformation, "Informação"
'LIMPAR OS CAMPOS E DEIXA PRONTO PARA O PRÓXIMO CADASTRO'
txtEmpresa.Text = ""
txtContato.Text = ""
txtCargo.Text = ""
txtOperadora.Text = ""
txtNumero.Text = ""
txtEmail.Text = ""
txtObs.Text = ""
preencherlistview
txtEmpresa.SetFocus
End Sub
Private Sub btnCancelar_Click()
Unload Me
End Sub
Private Sub Form_Terminate()
preencherlistview
End Sub
Private Sub txtSistema_Change()
Dim Pos As Integer
Pos = txtSistema.SelStart
txtSistema.Text = VBA.UCase(txtSistema.Text)
txtSistema.SelStart = Pos
End Sub
'PROCEDIMENTO PARA PREENCHER A LISTVIEW'
Private Sub preencherlistview()
Dim item As ListItem
usrPrincipal.lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & txtEmpresa.Text & "%' AND contato like '" & txtContato.Text & "%' AND cargo like '" & txtCargo.Text & "%' and operadora like '" & txtOperadora.Text & "%' order by CONTATO", db, 3, 3

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
