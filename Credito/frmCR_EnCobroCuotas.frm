VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_EnCobroCuotas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desglose de Envio y Recepción de Cuotas al Cobro por Planillas"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11580
   Icon            =   "frmCR_EnCobroCuotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2772
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   11292
      _Version        =   1572864
      _ExtentX        =   19918
      _ExtentY        =   4890
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
      SelectedItem    =   1
      Item(0).Caption =   "Resumen Deductoras"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Bitácora"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswBitacora"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2412
         Left            =   -70000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   11292
         _Version        =   1572864
         _ExtentX        =   19918
         _ExtentY        =   4254
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
         MultiSelect     =   -1  'True
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   2412
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   11292
         _Version        =   1572864
         _ExtentX        =   19918
         _ExtentY        =   4254
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
         MultiSelect     =   -1  'True
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_EnCobroCuotas.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5880
      Top             =   120
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5076
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   8954
      _StockProps     =   64
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
      MaxCols         =   7
      MaxRows         =   499
      SpreadDesigner  =   "frmCR_EnCobroCuotas.frx":0413
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboDeductora 
      Height          =   312
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8916
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtMeses 
      Height          =   312
      Left            =   9360
      TabIndex        =   5
      Top             =   960
      Width           =   372
      _Version        =   1572864
      _ExtentX        =   656
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
      Text            =   "12"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Left            =   9840
      TabIndex        =   6
      Top             =   960
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Consultar"
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
      Appearance      =   17
      Picture         =   "frmCR_EnCobroCuotas.frx":1B4D
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   312
      Left            =   1080
      TabIndex        =   10
      Top             =   960
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   852
      _Version        =   1572864
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Periodo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMeses 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Meses a Extraer :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      TabIndex        =   3
      Top             =   960
      Width           =   852
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCR_EnCobroCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vCedula As String, vInstitucion As Integer, vPaso As Boolean
Dim mFrecuencPago As String

Private Sub btnConsulta_Click()
Call vGrid_SheetChanged(1, 1)
End Sub

Private Sub cboDeductora_Click()

If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = " select isnull(I.Frecuencia,'M') as 'Frecuencia_Id', isnull(E.PROCESO,0) as 'Proceso' , E.CORTE" _
       & " from INSTITUCIONES I" _
       & " left join vCrd_Deductora_Ultimo_Envio E on I.COD_INSTITUCION = E.COD_INSTITUCION" _
       & " where I.cod_institucion = " & cboDeductora.ItemData(cboDeductora.ListIndex)
Call OpenRecordSet(rs, strSQL)
    mFrecuencPago = rs!Frecuencia_ID
    txtProceso.Text = CStr(rs!Proceso)
rs.Close

If txtProceso.Text = "0" Then
    txtProceso.Text = GLOBALES.glngFechaCR
End If

Call btnConsulta_Click

Exit Sub

vError:

End Sub

Private Sub sbResumenDeductoras(pProceso As Currency)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

On Error GoTo vError

lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "Deductora", 1200
    .Add , , "Descripción", 3200
    .Add , , "Enviado", 1800
    .Add , , "Recibido", 1800
    .Add , , "Diferencia", 1800
End With

           
strSQL = "exec spPrm_Resumen_Persona '" & vCedula & "'," & pProceso
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!desc_corta)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = Format(rs!ENVIADO, "Standard")
     itmX.SubItems(3) = Format(rs!RECIBIDO, "Standard")
     itmX.SubItems(4) = Format(rs!RECIBIDO - rs!ENVIADO, "Standard")
      
 rs.MoveNext
Loop
rs.Close


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Function fxPlanillaTipoTransac(pTransaccion As String) As String
Dim vResultado As String

Select Case Trim(pTransaccion)
 Case "01"
   vResultado = "Cambia Fecha de Proceso"
 Case "02"
   vResultado = "Genera deducciones"
 Case "02.1"
   vResultado = "Construye Archivo de Deducciones"
 Case "02.2"
   vResultado = "Deducciones Modificadas Manualmente"
 Case "03"
   vResultado = "Carga deducciones"
 Case "04"
   vResultado = "Desglosa deducciones"
 Case "05"
   vResultado = "Aplica Ahorros"
 Case "06"
   vResultado = "Inconsistencias de Ahorros"
 Case "07"
   vResultado = "Devoluciones de Ahorros"
 Case "08"
   vResultado = "Aplica Abonos"
 Case "08.2"
   vResultado = "Aplica Abonos x Inconsistencia"
 Case "08.3"
   vResultado = "Crea Fondos x Clientes Activos"
 Case "08.4"
   vResultado = "Crea Fondos x Clientes Inactivos"
 Case "09"
   vResultado = "Reporte de Inconsistencias"
 Case "10"
   vResultado = "Actualiza Intereses Moratorios"
 Case "11"
   vResultado = "Actualiza Saldo del Mes"
 Case Else
   vResultado = "No.Identificado"
End Select
fxPlanillaTipoTransac = vResultado

End Function

Private Sub sbBitacora(pProceso As Currency)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

lswBitacora.ListItems.Clear

strSQL = "exec spPrm_Bitacora_Consulta " & cboDeductora.ItemData(cboDeductora.ListIndex) & ", " & pProceso
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lswBitacora.ListItems.Add(, , rs!Id_seq)
     itmX.SubItems(1) = rs!GestionDesc
     itmX.SubItems(2) = rs!TransaccionDesc
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
      
    Call sbResumenDeductoras(vFecha)
    Call sbBitacora(vFecha)
    
    Call btnConsulta_Click
    
    
      
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

On Error GoTo vError

imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

txtProceso.Text = GLOBALES.glngFechaCR

strSQL = "select cedula,nombre,cod_institucion from socios where cedula = '" _
       & GLOBALES.gCedulaActual & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   lbl.Caption = Trim(rs!Cedula) & " - " & Trim(rs!Nombre)
   lbl.Tag = Trim(rs!Cedula)
   vCedula = Trim(rs!Cedula)
   vInstitucion = rs!cod_institucion
End If
rs.Close

strSQL = "exec spAFI_Institucion_Vinculadas " & vInstitucion & ",3"
Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)


strSQL = "select isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & vInstitucion
Call OpenRecordSet(rs, strSQL)
    mFrecuencPago = rs!Frecuencia_ID
rs.Close


vPaso = False


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Deductora", 1200
    .Add , , "Descripción", 3200
    .Add , , "Enviado", 1800
    .Add , , "Recibido", 1800
    .Add , , "Diferencia", 1800
End With



With lswBitacora.ColumnHeaders
   .Clear
   .Add , , "Id", 500, vbCenter
   .Add , , "Gestión", 1010
   .Add , , "Transacción", 3000
   .Add , , "Documento", 1740
   .Add , , "Usuario", 1200
   .Add , , "Fecha", 2240
End With

Call cboDeductora_Click

Exit Sub

vError:


End Sub


Private Sub Timer1_Timer()

Timer1.Interval = 0

On Error GoTo vError

'Call btnConsulta_Click

Call sbResumenDeductoras(txtProceso.Text)
Call sbBitacora(txtProceso.Text)

Exit Sub

vError:


End Sub



Private Sub txtMeses_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call vGrid_SheetChanged(1, 5)
End If
End Sub


Private Sub txtProceso_Change()
On Error GoTo vError

Call sbResumenDeductoras(txtProceso.Text)
Call sbBitacora(txtProceso.Text)

Exit Sub

vError:

End Sub

Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vProceso As Currency
Dim curTotal As Currency, curTotal2 As Currency


On Error GoTo vError

Me.MousePointer = vbHourglass


vProceso = txtProceso.Text
curTotal = 0
curTotal2 = 0

If NewSheet = 5 Then
   lblMeses.Visible = True
   txtMeses.Visible = True
Else
   lblMeses.Visible = False
   txtMeses.Visible = False
End If


vInstitucion = cboDeductora.ItemData(cboDeductora.ListIndex)

With vGrid
    .Sheet = NewSheet
    .MaxRows = 0
    
    Select Case NewSheet
     Case 1 'Resumen
        strSQL = "exec spPrm_Compara_Persona " & vInstitucion & "," & vProceso & ",'" & vCedula & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         For i = 1 To 7
           .Col = i
           Select Case i
                Case 1 'Operación
                   .Text = CStr(rs!Operacion)
                Case 2 'Linea
                    .Text = CStr(rs!Linea)
                Case 3 'Envio
                    .Text = Format(rs!Envio, "Standard")
                Case 4 'Recepción
                    .Text = Format(rs!RECIBIDO, "Standard")
                Case 5 'Diferencia
                    .Text = Format(rs!RECIBIDO - rs!Envio, "Standard")
                Case 6 'Tipo de Caso
                    .Text = rs!TipoDesc
                Case 7 'Desc. Línea
                    .Text = CStr(rs!LineaDesc)
           End Select
         Next i
         curTotal = curTotal + rs!Envio
         curTotal2 = curTotal2 + rs!RECIBIDO
         
         rs.MoveNext
        Loop
        rs.Close
        
        .MaxRows = .MaxRows + 1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 3
        .Text = Format(curTotal, "Standard")
        .Col = 4
        .Text = Format(curTotal2, "Standard")
        .Col = 5
        .Text = Format(curTotal - curTotal2, "Standard")
        
     Case 2 'Detalle
     
     
     
        strSQL = "exec spPrm_Planilla_Compara_Detalle '" & vCedula & "'," & vProceso & "," & vInstitucion
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         For i = 1 To 15
           .Col = i
           Select Case i
                Case 1 'Operación
                   .Text = CStr(rs!Operacion)
                Case 2 'Linea
                    .Text = CStr(rs!Linea)
                Case 3 'Proceso
                    .Text = Format(rs!Proceso, "####-##")
                Case 4 'Int.Cor.
                    .Text = Format(rs!IntCor, "Standard")
                Case 5 'Int.Mor.
                    .Text = Format(rs!IntMor, "Standard")
                Case 6 'Cargos
                    .Text = Format(rs!Cargos, "Standard")
                Case 7 'Principal
                    .Text = Format(rs!Principal, "Standard")
                Case 8 'Total Abonad
                    .Text = Format(rs!Cargos + rs!IntCor + rs!IntMor + rs!Principal, "Standard")
                Case 9 'Total Enviado
                    .Text = Format(rs!ENVIADO, "Standard")
                Case 10 'Diferencia
                    .Text = Format((rs!Cargos + rs!IntCor + rs!IntMor + rs!Principal) - rs!ENVIADO, "Standard")
                Case 11 'Aplicación
                    .Text = Format(rs!fecha, "dd/mm/yyyy")
                Case 12 'Linea Descripción
                    .Text = CStr(rs!Descripcion)
                Case 13 'Destino
                    .Text = CStr(rs!Destino & "")
                Case 14 'Caso
                    .Text = CStr(rs!Caso)
                Case 15 'Tipo de Cuota
                    .Text = CStr(rs!Tipo)
           
           End Select
         Next i
         curTotal = curTotal + rs!ENVIADO
         curTotal2 = curTotal2 + rs!Cargos + rs!IntCor + rs!IntMor + rs!Principal
         
         rs.MoveNext
        Loop
        rs.Close
        
        .MaxRows = .MaxRows + 1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 8
        .Text = Format(curTotal2, "Standard")
        .Col = 9
        .Text = Format(curTotal, "Standard")
        .Col = 10
        .Text = Format(curTotal2 - curTotal, "Standard")
     
     
     
     
     Case 3 'Envío
        strSQL = "select C.id_solicitud,R.codigo,X.descripcion,C.cuota,C.morosidad,C.Cod_Deduccion " _
               & " from PRM_ENVIADO_DETALLE C inner join reg_creditos R on C.id_solicitud = R.id_solicitud" _
               & " inner join Catalogo X on R.codigo = X.codigo" _
               & " where C.fecpro = " & vProceso & " and C.cedula = '" & vCedula & "' and C.cod_Institucion = " & vInstitucion
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         curTotal = curTotal + rs!Cuota
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         For i = 1 To 6
           .Col = i
           Select Case i
                Case 1 'Operación
                   .Text = CStr(rs!ID_SOLICITUD)
                Case 2 'Línea
                    .Text = CStr(rs!Codigo)
                Case 3 'Descripcion
                    .Text = CStr(rs!Descripcion)
                Case 4 'Cuota
                    .Text = Format(rs!Cuota, "Standard")
                Case 5 'Tipo
                    .Text = IIf((rs!Morosidad = 0), "Ordinario", "Morosidad")
                Case 6 'Tipo
                    .Text = CStr(rs!cod_deduccion)
                    
           End Select
         Next i
         rs.MoveNext
        Loop
        rs.Close
                    
        .MaxRows = .MaxRows + 1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 3
        .Text = "Total Enviado ...:"
        .Col = 4
        .Text = Format(curTotal, "Standard")
     
     Case 4 'Recepción
     
        strSQL = "exec spPrm_Planilla_Recepcion_Aplicada " & vInstitucion & ", " & vProceso & ", '" & vCedula & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         curTotal = curTotal + rs!Abono
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         For i = 1 To 5
           .Col = i
           Select Case i
                Case 1 'Operación
                   .Text = CStr(rs!ID_SOLICITUD)
                Case 2 'Línea
                    .Text = Trim(CStr(rs!Codigo))
                Case 3 'Descripcion
                    .Text = CStr(rs!Descripcion)
                Case 4 'Cuota
                    .Text = Format(rs!Abono, "Standard")
                Case 5 'Tipo
                    .Text = IIf((rs!Tipo = "C"), "Ordinario", "Morosidad")
           End Select
         Next i
         rs.MoveNext
        Loop
        rs.Close
                    
        .MaxRows = .MaxRows + 1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 3
        .Text = "Total Planilla ...:"
        .Col = 4
        .Text = Format(curTotal, "Standard")
        
        
        
        
        strSQL = "exec spPrm_Planilla_Recepcion_Aplicada_Otros " & vInstitucion & ", " & vProceso & ", '" & vCedula & "', 'I'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 3
            .Text = "Total N.C...:"
            .Col = 4
            .Text = Format(rs!Monto, "Standard")
        
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 3
            .Text = "Total General ...:"
            .Col = 4
            .Text = Format(rs!Monto + curTotal, "Standard")
        Else
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 3
            .Text = "Total N.C...:"
            .Col = 4
            .Text = Format(0, "Standard")
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 3
            .Text = "Total General ...:"
            .Col = 4
            .Text = Format(curTotal, "Standard")
                
        End If
        rs.Close

        strSQL = "exec spPrm_Planilla_Recepcion_Aplicada_Otros " & vInstitucion & ", " & vProceso & ", '" & vCedula & "', 'C'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
            .MaxRows = .MaxRows + 1
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 3
            .Text = "Total Recaudado ...:"
            .Col = 4
            .Text = Format(rs!Monto, "Standard")
        End If
        rs.Close
 
     
     
     Case 5 'Historial
        strSQL = "exec spCrdHistoricoPlanilla '" & vCedula & "'," & CInt(txtMeses.Text)
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         For i = 1 To 5
           .Col = i
           Select Case i
                Case 1 'Proceso
                   .Text = Format(rs!Proceso, "####-##")
                Case 2 'Enviado
                    .Text = Format(rs!ENVIADO, "Standard")
                Case 3 'Recibido
                    .Text = Format(rs!RECIBIDO, "Standard")
                Case 4 'Diferencia
                    .Text = Format(rs!RECIBIDO - rs!ENVIADO, "Standard")
                Case 5 'Institucion
                    .Text = CStr(rs!Institucion)
           End Select
         Next i
         rs.MoveNext
        Loop
        rs.Close
    
    End Select

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
