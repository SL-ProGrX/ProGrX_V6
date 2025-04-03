VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_BeneficiosBancosX 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Beneficios - Bancos de Tramite Rápido"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   7935
      _Version        =   524288
      _ExtentX        =   13996
      _ExtentY        =   11245
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_BeneficiosBancosX.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   $"frmAF_BeneficiosBancosX.frx":0698
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8172
   End
End
Attribute VB_Name = "frmAF_BeneficiosBancosX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbInicializa()
Dim strSQL As String


vPaso = True

'Ingresa los Tes_Bancos nuevos
strSQL = "insert into afi_bene_Bancos_X(id_banco,cheque,transferencia) select id_banco,0,0 from Tes_Bancos" _
       & " where id_Banco not in (select id_Banco from afi_bene_Bancos_X)"
Call ConectionExecute(strSQL)

strSQL = "select X.id_banco,B.descripcion,X.cheque,X.transferencia" _
       & " from afi_bene_Bancos_X X inner join Tes_Bancos B on X.id_banco = B.id_Banco order by B.id_banco"
Call sbCargaGrid(vGrid, 4, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

End Sub


Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()

vModulo = 7
vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)

Call sbInicializa

End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = Col
   
Select Case Col
  Case 1, 2
     Exit Sub
  Case 3 'CK
     strSQL = "update afi_bene_Bancos_X set cheque = " & vGrid.Value
  Case 4 'TE
     strSQL = "update afi_bene_Bancos_X set transferencia = " & vGrid.Value
End Select
   
vGrid.Col = 1
strSQL = strSQL & " where id_Banco = " & vGrid.Text
Call ConectionExecute(strSQL)


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


