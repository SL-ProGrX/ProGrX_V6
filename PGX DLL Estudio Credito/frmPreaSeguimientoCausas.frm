VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaSeguimientoCausas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Causas de Seguimiento"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9195
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8895
      _Version        =   1572864
      _ExtentX        =   15684
      _ExtentY        =   8276
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
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   6840
      Top             =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique las causas por las cuales esta indicando que la solicitud queda Pendiente o Denegada con las opciones siguientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   1800
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
Attribute VB_Name = "frmPreaSeguimientoCausas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mCod_linea As String, vPaso As Boolean

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture


lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Código", 1200
lsw.ColumnHeaders.Add , , "Descripción", 3200
lsw.ColumnHeaders.Add , , "Fecha", 2800
lsw.ColumnHeaders.Add , , "Usuario", 2800

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "insert CRD_PREA_GESTION(cod_causas,tipo,cod_preanalisis,codigo,registro_fecha,registro_usuario) values('" _
           & Item.Text & "','" & gPreAnalisis.Estado & "','" & gPreAnalisis.Expediente _
           & "','" & mCod_linea & "',dbo.Mygetdate(), '" & glogon.Usuario & "')"
Else
  Call Bitacora("Elimina", "Causa SGT: " & Item.Text & ", Expediente: " & gPreAnalisis.Expediente)
    
  strSQL = "delete CRD_PREA_GESTION where cod_causas = '" & Item.Text & "' and tipo = '" _
         & gPreAnalisis.Estado & "' and cod_preanalisis = '" & gPreAnalisis.Expediente & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pTipo As String

On Error GoTo vError

TimerX.Interval = 0
TimerX.Enabled = False


Me.MousePointer = vbHourglass


pTipo = gPreAnalisis.Estado

lsw.ListItems.Clear

vPaso = True

strSQL = "select Cg.COD_CAUSAS, Cg.DESCRIPCION,case when isnull(Pa.Cod_Causas,'No Existe') = 'No Existe' then 0 else 1 end as 'Check'" _
       & " , Pa.Registro_Fecha, Pa.Registro_Usuario " _
       & " from OPERACION_CAUSAS Cg " _
       & "       left join CRD_PREA_GESTION Pa on Cg.COD_CAUSAS = Pa.COD_CAUSAS and Cg.TIPO = Pa.TIPO" _
       & "             and  Pa.COD_PREANALISIS = '" & gPreAnalisis.Expediente _
       & "' Where Cg.TIPO = '" & pTipo & "'" _
       & " order by isnull(Pa.REGISTRO_FECHA,getdate()) asc, Cg.Cod_Causas"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Causas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Registro_Fecha & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
     itmX.Checked = IIf((rs!Check = 1), True, False)
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

