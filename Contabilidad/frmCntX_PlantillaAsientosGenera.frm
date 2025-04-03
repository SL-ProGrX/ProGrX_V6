VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCntX_PlantillaAsientosGenera 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generador de Asientos de Plantillas..."
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18230
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGenerar 
         Height          =   495
         Left            =   8760
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Generar"
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
         Picture         =   "frmCntX_PlantillaAsientosGenera.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   9960
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl TabControl_Main 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18224
      _ExtentY        =   9123
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
      Item(0).Caption =   "Generación"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "chkTodos"
      Item(0).Control(2)=   "scTitulo"
      Item(1).Caption =   "Resultados"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "txt"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4455
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   10335
         _Version        =   1310723
         _ExtentX        =   18230
         _ExtentY        =   7858
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
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txt 
         Height          =   4815
         Left            =   -70000
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1310723
         _ExtentX        =   18230
         _ExtentY        =   8493
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   210
         Left            =   360
         TabIndex        =   4
         Top             =   460
         Width           =   210
         _Version        =   1310723
         _ExtentX        =   379
         _ExtentY        =   379
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   10335
         _Version        =   1310723
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione las Plantillas que desea Procesar:"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Generación de Asientos Fijos y Proyectados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "frmCntX_PlantillaAsientosGenera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


'Selecionar solo las plantillas que su inicio es menor igual al periodo actual
strSQL = "select * from CntX_Plantilla_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by cod_plantilla"
Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear
Do While Not rs.EOF
  Select Case True
     Case CLng(gCntX_Parametros.PeriodoAnio) = rs!anio_inicio
        If CLng(gCntX_Parametros.PeriodoMes) >= rs!mes_inicio Then
            Set itmX = lsw.ListItems.Add(, , rs!cod_plantilla)
                itmX.SubItems(1) = rs!Descripcion & ""
                itmX.SubItems(2) = rs!Consecutivo
        End If
     
     Case CLng(gCntX_Parametros.PeriodoAnio) > rs!anio_inicio
        Set itmX = lsw.ListItems.Add(, , rs!cod_plantilla)
            itmX.SubItems(1) = rs!Descripcion & ""
            itmX.SubItems(2) = rs!Consecutivo
     
   End Select
     
  rs.MoveNext
Loop
rs.Close

End Sub

Private Function fxCntX_PeriodosDiferencia(lngAnioInicio As Long, iMesInicio As Integer)
Dim lngAnio As Long, iMes As Integer
Dim i As Integer, vPaso As Boolean

lngAnio = gCntX_Parametros.PeriodoAnio
iMes = gCntX_Parametros.PeriodoMes

vPaso = True
i = 0

Do While vPaso
  If lngAnioInicio = lngAnio Then
     If iMesInicio = iMes Then
       vPaso = False
     Else
       i = i + 1
       If iMesInicio = 12 Then
           iMesInicio = 1
           lngAnioInicio = lngAnioInicio + 1
       Else
           iMesInicio = iMesInicio + 1
       End If
     End If
  
  Else
       
       i = i + 1
       If iMesInicio = 12 Then
           iMesInicio = 1
           lngAnioInicio = lngAnioInicio + 1
       Else
           iMesInicio = iMesInicio + 1
       End If
  
  End If
Loop

fxCntX_PeriodosDiferencia = i


End Function


Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdGenerar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCodPlantilla As Long, vDiferencia As Integer
Dim vConsecutivo As Long, rsTmp As New ADODB.Recordset
Dim curDebito As Currency, curCredito As Currency, x As Integer


Me.MousePointer = vbHourglass
On Error GoTo vError

If Not fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
  Me.MousePointer = vbDefault
  MsgBox "El Periodo Actual se Encuentra Cerrado o no se ha creado, verifique...", vbExclamation
  Exit Sub
End If

txt = ""

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
    DoEvents
    vCodPlantilla = lsw.ListItems.Item(i).Text
    txt = txt & vbCrLf & "PROCESANDO : " & vCodPlantilla & "-" & lsw.ListItems.Item(i).SubItems(1)
    
    strSQL = "select * from CntX_Plantilla_Asientos" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & vCodPlantilla
    Call OpenRecordSet(rs, strSQL, 0)
    
    'Consecutivos de Cntx_Asientos de esta Plantilla
    vConsecutivo = rs!Consecutivo + 1
    
    'Saca la Diferencia en CntX_Periodos para Proyectar los Incrementos
    'Desde el Periodo de Inicio hasta el Actual
    vDiferencia = fxCntX_PeriodosDiferencia(rs!anio_inicio, rs!mes_inicio)
    
    strSQL = "update CntX_Plantilla_Asientos set consecutivo = " & vConsecutivo _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_plantilla = " & vCodPlantilla
    Call ConectionExecute(strSQL, 0)
    
    'Crea Maestro del Asiento
    strSQL = "insert into Cntx_Asientos(cod_contabilidad,tipo_asiento,num_asiento,descripcion,fecha_asiento,balanceado,anio,mes" _
           & ",user_crea,modulo,notas) values(" & gCntX_Parametros.CodigoConta & ",'" & Trim(rs!Tipo_Asiento) & "','PT" _
           & Format(vCodPlantilla, "000") & "-" & Format(vConsecutivo, "000000") & "','" _
           & rs!asiento_descripcion & "','" & gCntX_Parametros.PeriodoAnio & "/" & Format(gCntX_Parametros.PeriodoMes, "00") _
           & "/01','S'," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes & ",'" & glogon.Usuario _
           & "',20,'GENERADO CON PLANTILLA Asientos FIJOS COD:" & Format(vCodPlantilla, "000") & "')"
    Call ConectionExecute(strSQL, 0)
     
    
    
    strSQL = "select * from CntX_Plantilla_detalle" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_plantilla = " & vCodPlantilla _
           & " order by num_linea"
    Call OpenRecordSet(rsTmp, strSQL, 0)
    Do While Not rsTmp.EOF
      'VOY POR AQUI
      curDebito = 0
      curCredito = 0
      
      If rsTmp!inc_tipo = "P" Then 'Incremento Porcentual
         
         If rsTmp!Debitos > 0 Then
            curDebito = rsTmp!Debitos
            For x = 1 To vDiferencia
                curDebito = curDebito + (curDebito * (rsTmp!inc_Valor / 100))
            Next x
         Else 'Creditos
            curCredito = rsTmp!Creditos
            For x = 1 To vDiferencia
                curCredito = curCredito + (curCredito * (rsTmp!inc_Valor / 100))
            Next x
         End If
      
      Else 'Incremento por Monto
         If rsTmp!Debitos > 0 Then
            curDebito = rsTmp!Debitos + (rsTmp!inc_Valor * vDiferencia)
         Else
            curCredito = rsTmp!Creditos + (rsTmp!inc_Valor * vDiferencia)
         End If
      End If 'Tipo Incremento
      
      strSQL = "insert Cntx_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,cod_cuenta," _
             & "Monto_Debito,Monto_credito,Documento,Detalle,num_linea,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
             & " values(" & gCntX_Parametros.CodigoConta _
             & ",'" & Trim(rs!Tipo_Asiento) & "','PT" & Format(vCodPlantilla, "000") & "-" _
             & Format(vConsecutivo, "000000") & "','" & Trim(rsTmp!cod_cuenta) & "'," _
             & curDebito & "," & curCredito & ",'" & Trim(rs!asiento_documento) & "','" _
             & Trim(rs!asiento_detalle) & "'," & rsTmp!Num_linea & ",'" & Trim(rsTmp!Cod_Unidad) & "','" _
             & Trim(rsTmp!cod_Divisa) & "'," & rsTmp!TC & ",'" & Trim(rsTmp!cod_centro_costo) & "')"
      Call ConectionExecute(strSQL, 0)
      
      rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    txt = txt & vbCrLf & "Asiento: " & Trim(rs!Tipo_Asiento) & "-PT" _
           & Format(vCodPlantilla, "000") & "-" & Format(vConsecutivo, "000000") _
           & " (Creado...)"
    
    rs.Close
  End If
  
Next i

Me.MousePointer = vbDefault
MsgBox "Plantillas de Asientos Fijos y Proyectados Generados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Plantilla", 1200
    .Add , , "Descripción", 5800
    .Add , , "Consecutivo", 1400, vbRightJustify
End With

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Call sbCargaLsw
End Sub
