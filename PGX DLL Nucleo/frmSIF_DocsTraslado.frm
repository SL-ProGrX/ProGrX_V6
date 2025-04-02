VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmSIF_DocsTraslado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslado de Documentos "
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9225
   Icon            =   "frmSIF_DocsTraslado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6012
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   9012
      _Version        =   1441793
      _ExtentX        =   15896
      _ExtentY        =   10604
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Pendientes"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "cmdTrasladar"
      Item(0).Control(1)=   "Label1(2)"
      Item(0).Control(2)=   "lblEstatus"
      Item(0).Control(3)=   "chkDocumentos"
      Item(0).Control(4)=   "lswDocumentos"
      Item(1).Caption =   "Desbalanceados"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5412
         Left            =   -69760
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   8652
         _Version        =   524288
         _ExtentX        =   15261
         _ExtentY        =   9546
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   7
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmSIF_DocsTraslado.frx":6852
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdTrasladar 
         Height          =   612
         Left            =   7320
         TabIndex        =   12
         Top             =   5160
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Trasladar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmSIF_DocsTraslado.frx":6F58
      End
      Begin XtremeSuiteControls.ListView lswDocumentos 
         Height          =   4212
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   8772
         _Version        =   1441793
         _ExtentX        =   15473
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
         Checkboxes      =   -1  'True
         MultiSelect     =   -1  'True
         HideSelection   =   0   'False
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkDocumentos 
         Height          =   372
         Left            =   7800
         TabIndex        =   16
         Top             =   480
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Todos"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin VB.Label lblEstatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   792
         Left            =   240
         TabIndex        =   15
         Top             =   5160
         Width           =   6972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Pendientes de traslado...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   3372
      End
   End
   Begin XtremeSuiteControls.CheckBox chkBalanceados 
      Height          =   372
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Solo Asientos Balanceados"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   3
      Top             =   8208
      Width           =   9228
      _ExtentX        =   16272
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   612
      Left            =   6240
      TabIndex        =   4
      Top             =   1200
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSIF_DocsTraslado.frx":775D
   End
   Begin XtremeSuiteControls.PushButton btnReActivar 
      Height          =   612
      Left            =   7680
      TabIndex        =   5
      Top             =   1200
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Re Activar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSIF_DocsTraslado.frx":817B
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1080
      TabIndex        =   6
      Top             =   1200
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkReActivar 
      Height          =   372
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Re Activar Automáticamente"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Traslado de Asientos a Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   7572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   4
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   5
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIF_DocsTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public Function fxValidaPeriodoAsiento(vFecha As Date) As Boolean
'Dim strSQL As String, rsX As New ADODB.Recordset
'
'strSQL = "select * from CntX_periodos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
'        & " and estado = 'P' and cod_contabilidad = " & GLOBALES.gEnlace
'
'Call OpenRecordSet(rsX, strSQL)
'
'If rsX.EOF And rsX.BOF Then
' fxValidaPeriodoAsiento = False
'Else
' fxValidaPeriodoAsiento = True
'End If
'rsX.Close
'
'End Function

'Public Function fxUltimaLineaAsiento(pTipoAsiento As String, pNumAsiento As String, vFecha As Date) As Integer
'Dim strSQL As String, rsX As New ADODB.Recordset
'
'strSQL = "Select isnull(max(num_linea),0) as Linea from CntX_asientos_detalle" _
'       & " where num_asiento = '" & pNumAsiento & "' and Tipo_asiento = '" & pTipoAsiento & "'" _
'       & " and cod_contabilidad = " & GLOBALES.gEnlace
'
'Call OpenRecordSet(rsX, strSQL)
'    fxUltimaLineaAsiento = IIf(IsNull(rsX!Linea), 0, rsX!Linea)
'rsX.Close
'
'End Function
'
'Public Function fxVerificaExistenciaAsiento(pTipoAsiento As String, pNumAsiento As String, vFecha As Date) As Boolean
'Dim strSQL As String, rsX As New ADODB.Recordset
'
'strSQL = "Select num_asiento from CntX_Asientos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
'        & " and tipo_asiento = '" & pTipoAsiento & "' and num_asiento = '" & pNumAsiento _
'        & "' and cod_contabilidad = " & GLOBALES.gEnlace
'Call OpenRecordSet(rsX, strSQL)
'
'If rsX.EOF And rsX.BOF Then
' fxVerificaExistenciaAsiento = False
'Else
' fxVerificaExistenciaAsiento = True
'End If
'rsX.Close
'
'End Function

Private Sub btnBuscar_Click()

If chkReActivar.Value = vbChecked Then
    Call sbReActivar(1)
End If

Call sbBuscar

End Sub

Private Sub btnReActivar_Click()
Call sbReActivar(0)
End Sub


Private Sub chkDocumentos_Click()
Dim i As Integer

For i = 1 To lswDocumentos.ListItems.Count
  lswDocumentos.ListItems.Item(i).Checked = chkDocumentos.Value
Next i

End Sub

Private Sub cmdTrasladar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

With lswDocumentos.ListItems
  
  prgBar.Max = .Count

  For i = 1 To .Count
   If .Item(i).Checked Then
      strSQL = "select TIPO_DOCUMENTO,TIPO_ASIENTO,ASIENTO_MASCARA,ASIENTO_TRANSACCION,ASIENTO_MODULO" _
             & " from SIF_DOCUMENTOS where Tipo_Documento = '" & .Item(i).Tag & "'"
      Call OpenRecordSet(rs, strSQL)
      
      lblEstatus.Caption = "Procesando " & .Item(i).Text
      lblEstatus.Refresh
      
'      prgBar.Value = i
      
      If rs!Asiento_Transaccion = 1 Then
         Call sbAsientoIndividual(rs!Tipo_Documento, rs!Tipo_Asiento, rs!Asiento_Mascara)
      Else
         Call sbAsientoTipoDiario(rs!Tipo_Documento, rs!Tipo_Asiento)
      End If
      rs.Close
   End If
  Next i
End With

Call Bitacora("Aplica", "Asientos del Control de Documentos")

Me.MousePointer = vbDefault

lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

MsgBox "Se realizó el Traslado de Asientos a Contabilidad...!", vbInformation



Call sbBuscar


End Sub

Private Sub Form_Activate()
 vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


With lswDocumentos.ColumnHeaders
    .Clear
    .Add , , "Tipo Transacción", 3200
    .Add , , "Pendientes", 1400, vbCenter
    .Add , , "Bloqueados", 1400, vbCenter
End With

End Sub

Private Sub sbAsientoTipoDiario(pTipoDoc As String, pTipoAsiento As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumAsiento As String, intLinea As Long, vConcepto As String
Dim DH As String
Dim rsTmp As New ADODB.Recordset, vFecha As Date


On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spSys_Asientos_CtrlDoc_Traslado_Bloque_Diario '" & pTipoDoc & "', '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
        & " 00:00:00', '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59', '" & glogon.Usuario & "', " & chkBalanceados.Value
Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault
lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1


Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    MsgBox "Asiento ...: " & vNumAsiento & vbCrLf & vbCrLf & Err.Description, vbCritical

End Sub


Private Sub sbAsientoIndividual(pTipoDoc As String, pTipoAsiento As String, Optional pMascara As String = "")
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumAsiento As String

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spSys_Asientos_CtrlDoc_Traslado_Individual '" & pTipoDoc & "', '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
        & " 00:00:00', '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59', '" & glogon.Usuario & "', " & chkBalanceados.Value
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

Exit Sub

vError:
    lblEstatus.Caption = ""
    lblEstatus.Refresh
    prgBar.Value = 1
    Me.MousePointer = vbDefault
    
    MsgBox "Asiento ...: " & vNumAsiento & vbCrLf & vbCrLf & Err.Description, vbCritical

End Sub


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

lswDocumentos.ListItems.Clear

'Sacar los Documentos Pendientes de Inicio y Corte
strSQL = "exec spSys_Asientos_CtrlDoc_Busca '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' , '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59', " & chkBalanceados.Value
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswDocumentos.ListItems.Add(, , rs!Descripcion)
     itmX.SubItems(1) = rs!Pendientes & ""
     itmX.SubItems(2) = rs!Bloqueados & ""
     itmX.Tag = rs!Tipo_Documento
     itmX.Checked = chkDocumentos.Value
 rs.MoveNext
Loop
rs.Close


'Carga Transacciones con Asientos desbalanceados
strSQL = "exec spSys_Asientos_CtrlDoc_Desbalanceados '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' , '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call sbCargaGrid(vGrid, 7, strSQL, True)


Me.MousePointer = vbDefault

End Sub



Private Sub sbReActivar(Optional pAutomatico As Integer = 0)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


'Revisa Asientos Trasladados que no estan en la contabilidad
strSQL = "exec  spSys_Asiento_Revisa_Traslado '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call ConectionExecute(strSQL)

If glogon.error Then Exit Sub

Me.MousePointer = vbDefault

If pAutomatico = 0 Then
    MsgBox "Revisión de Documentos realizada satisfactoriamente!", vbInformation
    Call sbBuscar
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  Case "Reactivar"
    Call sbReActivar
End Select

End Sub

