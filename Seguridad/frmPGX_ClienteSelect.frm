VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmPGX_ClienteSelect 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cliente: Selección"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   12000
      _Version        =   1441793
      _ExtentX        =   21167
      _ExtentY        =   12091
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   6240
      Top             =   240
   End
   Begin XtremeSuiteControls.FlatEdit txtCriterio 
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12012
      _Version        =   1441793
      _ExtentX        =   21188
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Selecciones una Empresa:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPGX_ClienteSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vPortal_Admin As Boolean
Dim vPaso As Boolean

Private Sub Form_Load()

With lsw.ColumnHeaders
   .Clear
   .Add , , "Acrónimo", 5200
   .Add , , "Nombre", 6400
End With

gPortal.Empresa_Id = 0
gPortal.Empresa_Name = "No Identidicada"


vPortal_Admin = Sys_Portal_Admin_Valid(glogon.Usuario)

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

If lsw.ListItems.Count <= 0 Then Exit Sub

gPortal.Empresa_Id = Item.Tag
gPortal.Empresa_Name = Item.Text

Unload Me

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0

Call sblistaClientes

End Sub


Private Sub sblistaClientes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vPaso = True
lsw.ListItems.Clear
gPortal.Empresa_Id = 0
gPortal.Empresa_Name = "(No se ha seleccionado ningún cliente)"


txtCriterio.Text = fxSysCleanTxtInject(txtCriterio.Text)

strSQL = "exec spSEG_Admin_Client_Access_List '" & glogon.Usuario & "', '" & txtCriterio.Text & "', 30"

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Nombre_Corto)
      itmX.SubItems(1) = rs!Nombre_Largo
      itmX.Tag = rs!cod_Empresa
  rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub



Private Sub txtCriterio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sblistaClientes
End If
End Sub
