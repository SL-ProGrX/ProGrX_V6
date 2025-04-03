VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPrea_Abandona_Motivos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Motivos de Abandono del Expediente"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   10821
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      Appearance      =   21
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   600
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Motivos de Abandono"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   4845
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPrea_Abandona_Motivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean



Private Sub Form_Load()
Me.Caption = "Expediente : " & gPreAnalisis.Expediente

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Código", 1200
lsw.ColumnHeaders.Add , , "Descripción", 3200
lsw.Checkboxes = True


End Sub


Private Sub sbLista()

On Error GoTo vError

vPaso = True

strSQL = "exec spCrdPreaListaMotivosSeleccion '" & gPreAnalisis.Expediente & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Id_Motivo)
      itmX.SubItems(1) = rs!Motivo
      
      If rs!Activo = 1 Then
         itmX.Checked = True
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub


On Error GoTo vError

strSQL = "exec spCrdPreaMotivosAbandono_Registro '" & gPreAnalisis.Expediente & "', " & Item.Text & ", " & IIf(Item.Checked, 1, 0) & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Timer1_Timer()
Timer1.Interval = 0
Timer1.Enabled = False

Call sbLista

End Sub
