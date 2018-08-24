VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "CHAMALEONBUTTON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form usrPrincipal 
   Caption         =   "Agenda Telefônica"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   Icon            =   "usrPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11310
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   9480
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   9480
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   7485
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   873
      SimpleText      =   "54545"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3951
            MinWidth        =   3951
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "10/07/2017"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "23:11"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema Criado por José Paulo De Oliveira Junior"
            TextSave        =   "Sistema Criado por José Paulo De Oliveira Junior"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstTelefones 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8493
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777088
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CÓDIGO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EMPRESA"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CONTATO"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CARGO"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "OPERADORA"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "NUMERO"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "EMAIL"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":1082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":2114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":29EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":32C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":3BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":4C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":5CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":6D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":7DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":86C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":9756
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":A7E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":D902
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":F56C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":16A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":17B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":18B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":2FBA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":46BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":4CE50
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPrincipal.frx":4D72A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1429
      ButtonWidth     =   1826
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NOVO"
            Object.ToolTipText     =   "Permite inserir o registro de uma nova senha"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ABRIR"
            Object.ToolTipText     =   "Visualizar registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXCLUIR"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ATUALIZAR"
            Object.ToolTipText     =   "Atualiza o programa com a base de dados"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "BACKUP"
            Object.ToolTipText     =   "Realiza backup da base de dados"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SOBRE"
            Object.ToolTipText     =   "Informações sobre o sistema"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SAIR"
            Object.ToolTipText     =   "Sair do Sistema"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn btnLimpar 
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Limpar"
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "usrPrincipal.frx":4E004
      PICN            =   "usrPrincipal.frx":4E020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label CONTADOR_REGISTROS 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label CONTADOR_SELECIONADOS 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   9480
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7080
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C00000&
      X1              =   10080
      X2              =   10080
      Y1              =   1080
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   10080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      X1              =   360
      X2              =   120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   1440
      X2              =   10080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   960
      Width           =   975
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   22170
      Left            =   -120
      Picture         =   "usrPrincipal.frx":4F0B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   44310
   End
End
Attribute VB_Name = "usrPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodCli As Integer

'PROCEDIMENTO PARA PREENCHER A LISTVIEW'
Private Sub preencherlistview()
Dim item As ListItem
lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & txtEmpresa.Text & "%' AND contato like '" & txtContato.Text & "%' AND cargo like '" & txtCargo.Text & "%' and operadora like '" & txtOperadora.Text & "%' order by CONTATO", db, 3, 3


Do Until rs.EOF
Set item = lstTelefones.ListItems.Add(, , rs!Código)
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
CONTADOR_SELECIONADOS.Caption = "0"

End Sub
'PROCEDIMENTO PARA LIMPAR OS CAMPOS DE PESQUISA'
Private Sub btnLimpar_Click()
txtEmpresa.Text = ""
txtContato.Text = ""
txtCargo.Text = ""
txtOperadora.Text = ""
preencherlistview
txtEmpresa.SetFocus
CONTADOR_SELECIONADOS.Caption = "0"
End Sub
'PROCEDIMENTO PARA EXECUTAR COMANDOS AO LIMPAR O FORM'
Private Sub Form_Load()
preencherlistview
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW
CONTARITENSNALISTA
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW NA STATUSBAR
CONTARITENSNALISTA
End Sub
'PROCEDIMENTO PARA REDIMENSIONAR AS DIMENSÕES DO FORM'
Private Sub Form_Resize()


'PROCEDIMENTO PARA REDIMENSIONAR FORMULARIO'

On Error Resume Next
Dim ALTURA As Integer
Dim LARGURA As Integer

ALTURA = usrPrincipal.Height
LARGURA = usrPrincipal.Width

lstTelefones.Height = ALTURA - 3750
lstTelefones.Width = LARGURA - 460

StatusBar1.Width = LARGURA
Image1.Height = ALTURA
Image1.Width = LARGURA
StatusBar1.Panels(4).Width = LARGURA - 460

End Sub

'PREOCEDIMENTO PARA ORDENAR COLUNAS AO CLICAR NELAS'
'Private Sub lstTelefones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
'With lstTelefones
'
'If (.Sorted) And (ColumnHeader.SubItemIndex = .SortKey) Then
'            If .SortOrder = lvwAscending Then
'                .SortOrder = lvwDescending
'            Else
'                .SortOrder = lvwAscending
'            End If
'        Else
'            .Sorted = True
'            .SortKey = ColumnHeader.SubItemIndex
'            .SortOrder = lvwAscending
'        End If
'        .Refresh
'    End With
'
'    If Not lstTelefones.SelectedItem Is Nothing Then
'        lstTelefones.SelectedItem.EnsureVisible
'    End If
'End Sub
'PROCEDIMENTO PARA ABRIR REGISTRO AO DAR DOIS CLIQUES NA LISTA'
Private Sub lstTelefones_DblClick()
If CONTADOR_REGISTROS.Caption = "0" Then
MsgBox "Nenhum registro encontrado", vbExclamation, "Informação"
End If
If CONTADOR_REGISTROS.Caption <> "0" Then
usrEditar.Show 1
End If
End Sub

Private Sub lstTelefones_ItemClick(ByVal item As MSComctlLib.ListItem)
If lstTelefones.ListItems.Count = 0 Then Exit Sub
CodCli = lstTelefones.SelectedItem
Text1.Text = lstTelefones.SelectedItem.SubItems(1)
Text2.Text = lstTelefones.SelectedItem.SubItems(2)
'Text3.Text = lstTelefones.SelectedItem.SubItems(3)
'Text4.Text = lstTelefones.SelectedItem.SubItems(4)

CONTADOR_SELECIONADOS.Caption = lstTelefones.SelectedItem.ListSubItems.Count


End Sub

'PROCEDIMENTO PARA ABRIR COMANDOS AO CLICAR NA BARRA SUPERIOR DO FORM'
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
'BOTÃO NOVO'
Case 2
usrCadastro.Show 1
'BOTÃO ABRIR'
Case 4
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW
CONTARITENS
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW NA STATUSBAR
CONTARITENSNALISTA
If CONTADOR_REGISTROS.Caption = "0" Then
MsgBox "Nenhum registro encontrado", vbExclamation, "Informação"
End If
If CONTADOR_REGISTROS.Caption <> "0" Then
usrEditar.Show 1
End If
'BOTÃO EXCLUIR'
Case 6
Excluir
'BOTÃO REFRESH
Case 8
preencherlistview
Case 10
MsgBox "Em Breve...", vbInformation, "Em Breve..."
'REALIZA CÓPIA DE BACKUP DA BASE DE DADOS'
'BackupCopy
Case 12
MsgBox "Versão 1.0.0.0", vbInformation, "Dados da Versão"
'BOTÃO DE SAIR
Case 14
Unload Me
End Select
End Sub

Private Sub txtCargo_Change()
'DIGITAR APENAS LETRAS MAIUSCULAS'
Dim Pos As Integer
Pos = txtCargo.SelStart
txtCargo.Text = VBA.UCase(txtCargo.Text)
txtCargo.SelStart = Pos
Dim item As ListItem
lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & txtEmpresa.Text & "%' AND contato like '" & txtContato.Text & "%' AND cargo like '" & txtCargo.Text & "%' and operadora like '" & txtOperadora.Text & "%' order by CONTATO", db, 3, 3

Do Until rs.EOF
Set item = lstTelefones.ListItems.Add(, , rs!Código)
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
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtOperadora.SetFocus
End Sub

Private Sub txtContato_Change()
Dim Pos As Integer
Pos = txtContato.SelStart
txtContato.Text = VBA.UCase(txtContato.Text)
txtContato.SelStart = Pos
Dim item As ListItem
lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & txtEmpresa.Text & "%' AND contato like '" & txtContato.Text & "%' AND cargo like '" & txtCargo.Text & "%' and operadora like '" & txtOperadora.Text & "%' order by CONTATO", db, 3, 3

Do Until rs.EOF
Set item = lstTelefones.ListItems.Add(, , rs!Código)
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
End Sub

Private Sub txtContato_KeyPress(KeyAscii As Integer)

'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtCargo.SetFocus
End Sub

'PROCEDIMENTO PARA FILTRAR PELO CAMPO EMPRESA'
Private Sub txtEmpresa_Change()
Dim Pos As Integer
Pos = txtEmpresa.SelStart
txtEmpresa.Text = VBA.UCase(txtEmpresa.Text)
txtEmpresa.SelStart = Pos
Dim item As ListItem
lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & txtEmpresa.Text & "%' AND contato like '" & txtContato.Text & "%' AND cargo like '" & txtCargo.Text & "%' and operadora like '" & txtOperadora.Text & "%' order by CONTATO", db, 3, 3

Do Until rs.EOF
Set item = lstTelefones.ListItems.Add(, , rs!Código)
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
End Sub

Private Sub txtEmpresa_Click()
CONTADOR_SELECIONADOS.Caption = "0"
End Sub

'PROCEDIMENTO PARA DIGITAR APENAS NUMEROS NO CAMPO EMPRESA'
Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then txtContato.SetFocus
End Sub

'PROCEDIMENTO PARA EXIBIR CONFIRMAÇÃO DE SAÍDA'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Confirma Saída ?", vbYesNo + vbDefaultButton2 + vbExclamation, "Agenda Telefônica") = vbNo Then Cancel = 1
End Sub
Private Sub Excluir()
If CONTADOR_SELECIONADOS.Caption = "0" Then
MsgBox "Clique novamente no item que deseja excluir", vbExclamation, "Informação"
Else
If MsgBox("Confirma a exclusão do registro referente a empresa (" & Text1.Text & ") e corresponde ao contato (" & Text2.Text & ") ?", vbYesNo + vbDefaultButton2 + vbExclamation, "Confirmação") = vbYes Then
connectBD
rs.Open "Select * from TBTelefones Where CÓDIGO = " & CodCli, db, 3, 3
rs.Delete
fechaBD
preencherlistview
End If
End If
End Sub
Private Sub txtSistema_Click()
CONTADOR_SELECIONADOS.Caption = "0"
End Sub

Private Sub txtOperadora_KeyPress(KeyAscii As Integer)
'MUDAR PARA A PRÓXIMA TEXTBOX AO CLICAR EM ENTER
If KeyAscii = 13 Then btnLimpar.SetFocus

Dim item As ListItem
lstTelefones.ListItems.Clear
connectBD
rs.Open "SELECT *FROM TBTelefones where Empresa like '" & txtEmpresa.Text & "%' AND contato like '" & txtContato.Text & "%' AND cargo like '" & txtCargo.Text & "%' and operadora like '" & txtOperadora.Text & "%' order by CONTATO", db, 3, 3

Do Until rs.EOF
Set item = lstTelefones.ListItems.Add(, , rs!Código)
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

End Sub
