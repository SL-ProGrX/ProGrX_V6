VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_SeguimientoCausas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Causas de Seguimiento"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4332
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7572
      _Version        =   1245185
      _ExtentX        =   13356
      _ExtentY        =   7641
      _StockProps     =   77
      BackColor       =   -2147483643
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
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   6840
      Top             =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Indique las causas por las cuales esta solicitud queda Pendiente o Denegada con las opciones siguientes"
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
      Height          =   612
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   5172
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmCR_SeguimientoCausas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Function fxChecked(vCausa As String, vTipo As String) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from operacion_gestion" _
       & " where cod_causas = '" & vCausa & "' and Tipo = '" & vTipo _
       & "' and id_solicitud = " & Operacion.Operacion
Call OpenRecordSet(rsX, strSQL, 0)
    fxChecked = IIf((rsX!Existe = 0), False, True)
rsX.Close

End Function



Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

With lsw.ColumnHeaders
  .Clear
  .Add , , "Causa", 1500, vbCenter
  .Add , , "Descripción", 5500
End With

End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub


If Item.Checked Then
    strSQL = "insert operacion_gestion(cod_causas,tipo,id_solicitud,codigo) values('" _
           & Item.Text & "','" & Operacion.EstadoSolicitud & "'," & Operacion.Operacion _
           & ",'" & Operacion.Codigo & "')"
Else
  strSQL = "delete operacion_gestion where cod_causas = '" & Item.Text & "' and tipo = '" _
         & Operacion.EstadoSolicitud & "' and id_solicitud = " & Operacion.Operacion
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

TimerX.Interval = 0

Me.MousePointer = vbHourglass

strSQL = "select * from operacion_causas where estado = 1 and tipo = '" & Operacion.EstadoSolicitud & "'"
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_causas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = fxChecked(rs!cod_causas, rs!Tipo)
     
     If itmX.Checked Then itmX.ForeColor = vbBlue
     
 rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub
