VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCC_PlanillaBitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bitácora de Planilla"
   ClientHeight    =   5784
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10392
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   10392
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4212
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   10212
      _Version        =   1310720
      _ExtentX        =   18013
      _ExtentY        =   7429
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   6852
      _Version        =   1310720
      _ExtentX        =   12086
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   312
      Left            =   2400
      TabIndex        =   6
      Top             =   960
      Width           =   1692
      _Version        =   1310720
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitácora de Planillas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10452
   End
End
Attribute VB_Name = "frmCC_PlanillaBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vScroll As Boolean
Dim mFrecuencPago As String


Private Sub cboInstitucion_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear



strSQL = "select isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
Call OpenRecordSet(rs, strSQL)
    mFrecuencPago = rs!Frecuencia_ID
rs.Close

strSQL = "select * from prm_bitacora where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and proceso = " & txtProceso.Text & " Order by id_seq"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_seq)
     itmX.SubItems(1) = IIf(rs!Gestion = "R", "Recepción", "Envio")
     itmX.SubItems(2) = fxPlanillaTipoTransac(rs!Transaccion)
     itmX.SubItems(3) = rs!Documento
     itmX.SubItems(4) = rs!Usuario
     itmX.SubItems(5) = rs!fecha


 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 lsw.ListItems.Clear
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub FlatScrollBar_Change()
Dim vFecha As Currency

On Error GoTo vError

vFecha = txtProceso.Text


If vScroll Then
    
    If FlatScrollBar.Value = 1 Then
       vFecha = fxFechaProcesoSiguiente(vFecha)
    Else
       vFecha = fxFechaProcesoAnterior(vFecha)
    End If
    
    txtProceso.Text = vFecha
      
    Call cboInstitucion_Click
End If



vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


mFrecuencPago = "M"

With lsw.ColumnHeaders
   .Clear
   .Add , , "Id", 500, vbCenter
   .Add , , "Gestión", 1010
   .Add , , "Transacción", 3000
   .Add , , "Documento", 1740
   .Add , , "Usuario", 1200
   .Add , , "Fecha", 2240
End With

vPaso = True

txtProceso.Text = GLOBALES.glngFechaCR

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX from instituciones order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
    cboInstitucion.Text = rs!Descripcion
    mFrecuencPago = rs!Frecuencia_ID
rs.Close



vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

vPaso = False

Call cboInstitucion_Click

End Sub

