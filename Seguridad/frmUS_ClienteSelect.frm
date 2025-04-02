VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUS_ClienteSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cliente: Selección"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   6240
      Top             =   240
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   13361
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.TextBox txtCriterio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Text            =   "..."
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmUS_ClienteSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub lsw_ItemClick(ByVal Item As MSComctlLib.ListItem)
If vPaso Then Exit Sub

gPortal.Empresa_Id = Item.Key
gPortal.Empresa_Name = Item.Text

Unload Me

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0

Call sblistaClientes

End Sub


Private Sub sblistaClientes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

vPaso = True
lsw.ListItems.Clear
gPortal.Empresa_Id = 0
gPortal.Empresa_Name = "(No se ha seleccionado ningún cliente)"

strSQL = "select Top 30 Nombre_Largo, Nombre_Corto,Cod_Empresa" _
       & " from PGX_Clientes" _
       & " where Nombre_Largo like '%" & txtCriterio.Text & "%'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, rs!Cod_Empresa, rs!Nombre_Largo)
  rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub
