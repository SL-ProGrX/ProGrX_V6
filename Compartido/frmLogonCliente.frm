VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmLogonCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Organización:"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7812
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   12000
      _Version        =   1441792
      _ExtentX        =   21167
      _ExtentY        =   13779
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
      Left            =   5880
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtCriterio 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   12012
      _Version        =   1441792
      _ExtentX        =   21188
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione su Empresa:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4332
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   1092
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12012
      _Version        =   1441792
      _ExtentX        =   21188
      _ExtentY        =   1926
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLogonCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Load()

With lsw.ColumnHeaders
   .Clear
   .Add , , "Acrónimo", 5200
   .Add , , "Nombre", 6400
End With

gPortal.Empresa_Id = 0
gPortal.Empresa_Name = "No Identidicada"
End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

If lsw.ListItems.Count <= 0 Then Exit Sub

On Error GoTo vError

gPortal.Empresa_Id = Item.Tag
gPortal.Empresa_Name = Item.Text

Unload Me

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0

txtCriterio.SetFocus
Call sblistaClientes

End Sub


Private Sub sblistaClientes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vPaso = True
lsw.ListItems.Clear

gPortal.Empresa_Id = 0
gPortal.Empresa_Name = "(No se ha seleccionado ningún cliente)"

strSQL = "exec spPGX_Usuario_Access_List '" & glogon.Usuario & "','" & Trim(txtCriterio.Text) & "',''"
Call OpenRecordSet(rs, strSQL, 1)

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , Trim(rs!Nombre_Corto & ""))
      itmX.SubItems(1) = Trim(rs!Nombre_Largo & "")
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


