VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFNDSeleccionaOperadora 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selección de Operadora"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5640
      Top             =   120
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Operadora"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione la Operadora, que por defecto quiere afectar con sus movimientos...."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmFNDSeleccionaOperadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lsw_Click()
On Error GoTo vError

MDIMenu.StatusBar.Panels(8).Text = lsw.SelectedItem.SubItems(1)
Unload Me

Exit Sub
vError:
  MsgBox Err.Description
  Unload Me
End Sub

Private Sub Timer1_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Timer1.Interval = 0

On Error GoTo vError

Me.Icon = frmContenedor.Icon

strSQL = "Select Cod_Operadora,Descripcion from FND_OPERADORAS"
rs.Open strSQL, glogon.Conection
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!cod_operadora)
       itmX.SubItems(1) = Trim(rs!descripcion)
   rs.MoveNext
Loop
rs.Close

strSQL = "Select coalesce(count(*),0) as Operadoras from FND_OPERADORAS"
rs.Open strSQL, glogon.Conection
If rs.RecordCount = 1 Then
  MDIMenu.StatusBar.Panels(8).Text = lsw.ListItems.Item(1).SubItems(1)
  Unload Me
End If
rs.Close

Exit Sub
vError:
  MsgBox Err.Description
  Unload Me

End Sub
