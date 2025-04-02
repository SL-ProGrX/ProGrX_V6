VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCR_Comisiones_Bancos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas Bancarias para Pago de Comisiones de Cr�dito"
   ClientHeight    =   6564
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6564
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7932
      _Version        =   524288
      _ExtentX        =   13991
      _ExtentY        =   9546
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
      SpreadDesigner  =   "frmCR_Comisiones_Bancos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   852
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8292
      _Version        =   1245187
      _ExtentX        =   14626
      _ExtentY        =   1503
      _StockProps     =   14
      Caption         =   "Seleccione las Cuentas Bancarias y sus tipos de desembolsos para Comisiones de Colocaci�n de Cr�ditos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmCR_Comisiones_Bancos"
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
strSQL = "insert into CRD_Bancos_Autorizados(id_banco,cheques,transferencias,registro_fecha,registro_usuario)" _
       & " select id_banco,0,0,dbo.MyGetdate(),'" & glogon.Usuario & "' from Tes_Bancos" _
       & " where id_Banco not in (select id_Banco from CRD_Bancos_Autorizados)"
Call ConectionExecute(strSQL)

strSQL = "select X.id_banco,B.descripcion,X.cheques,X.transferencias" _
       & " from CRD_Bancos_Autorizados X inner join Tes_Bancos B on X.id_banco = B.id_Banco order by B.id_banco"
Call sbCargaGrid(vGrid, 4, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

End Sub


Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

Call Formularios(Me)
Call RefrescaTags(Me)

Call sbInicializa

End Sub


Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.col = col
   
Select Case col
  Case 1, 2
     Exit Sub
  Case 3 'CK
     strSQL = "update CRD_Bancos_Autorizados set cheques = " & vGrid.Value
  Case 4 'TE
     strSQL = "update CRD_Bancos_Autorizados set transferencias = " & vGrid.Value
End Select
   
vGrid.col = 1
strSQL = strSQL & " where id_Banco = " & vGrid.Text
Call ConectionExecute(strSQL)


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



