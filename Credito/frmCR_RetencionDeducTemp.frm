VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_RetencionDeducTemp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retenciones Comparacion de Fecha Proceso vrs Planilla (Temporal)"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_RetencionDeducTemp.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Monto"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Monto"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   6855
   End
   Begin VB.ComboBox cboCliente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   6855
   End
   Begin VB.TextBox txtCasos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Monto"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   7125
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTemp.frx":02F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTemp.frx":6B55
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTemp.frx":D3B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   264
      Left            =   7800
      TabIndex        =   6
      Top             =   960
      Width           =   1404
      _ExtentX        =   2487
      _ExtentY        =   476
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buscar"
            Object.ToolTipText     =   "Buscar archivos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reporte"
            Object.ToolTipText     =   "Reporte de Deducciones"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "archivo"
            Object.ToolTipText     =   "Crear Archivo"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   8493
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_RetencionDeducTemp.frx":13C19
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Diferencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   9240
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   15
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   10320
      X2              =   0
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   9
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmCR_RetencionDeducTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto(0).Text = 0
    txtMonto(1).Text = 0
    txtMonto(2).Text = 0
    txtCasos.Text = 0
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

strSQL = "select codigo + ' - ' + descripcion as ItmX from catalogo where retencion = 'S' and activo = 1"
Call sbLlenaCbo(cboCliente, strSQL, False, False)

strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)


txtProceso.Text = 200803

vGrid.MaxCols = 5
vGrid.MaxRows = 0



End Sub

Private Function fxRevisaCedula(pCedula As String) As Long
Dim i As Long
Dim x As Long

'Revisa si existe una linea con la cedula, Regresa el numero de fila si existe caso contrario regresa 0
i = 0
pCedula = Trim(pCedula)

For x = 1 To vGrid.MaxRows
    vGrid.Row = x
    vGrid.col = 1
    
    If Trim(vGrid.Text) = pCedula Then
       i = x
       Exit For
    End If
   
Next x

fxRevisaCedula = i

End Function

Private Sub sbCargaDeducciones()
Dim strCadena As String, rs As New ADODB.Recordset, curMonto As Currency
Dim strSQL As String, pRow As Long, curMntAnterior As Currency
Dim x As Long, tmpActual As Currency, tmpAnterior As Currency

On Error GoTo vError


If Not IsNumeric(txtProceso.Text) Then
   MsgBox "Fecha de Proceso no es válida...", vbExclamation
   Exit Sub
End If


If cboCliente.ListCount <= 0 Then Exit Sub
If cboInstitucion.ListCount <= 0 Then Exit Sub


Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0

strSQL = "select S.cedula,S.nombre, isnull(sum(D.abono),0) as Monto" _
       & " from creditos_Dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " where D.tcon in('PLA','1') and D.NCon = '" & txtProceso.Text & "' and D.codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
       & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " group by S.cedula,S.nombre"
Call OpenRecordSet(rs, strSQL)

prgBar.Visible = True
prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1


With vGrid
    Do While Not rs.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .col = 1
            .Text = CStr(Trim(rs!Cedula))
            
            .col = 2
            .Text = CStr(Trim(rs!Nombre))
            
            .col = 3
            .Text = CCur(rs!Monto)
            
            .col = 4
            .Text = 0
            
            .col = 5
            .Text = 0
            
            curMonto = curMonto + rs!Monto
                        
            prgBar.Value = prgBar.Value + 1
            rs.MoveNext
    Loop
    rs.Close

End With


strSQL = "select S.cedula,S.nombre, isnull(sum(D.abAmortiza),0) as Monto" _
       & " from Morosidad D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " where D.tcon in('PLA','1')  and D.NCon = '" & txtProceso.Text & "' and D.codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
       & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and D.estado = 'C'" _
       & " group by S.cedula,S.nombre"
Call OpenRecordSet(rs, strSQL)

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1


With vGrid
    Do While Not rs.EOF
        pRow = fxRevisaCedula(rs!Cedula)
        If pRow = 0 Then
            'Nueva
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .col = 1
            .Text = CStr(Trim(rs!Cedula))
            
            .col = 2
            .Text = CStr(Trim(rs!Nombre))
            
            .col = 3
            .Text = CCur(rs!Monto)
            
            curMonto = curMonto + rs!Monto
        Else
            .Row = pRow
            .col = 3
            .Text = CCur(.Text) + CCur(rs!Monto)
            
            curMonto = curMonto + rs!Monto
                
        End If
                    
        .col = 4
        .Text = 0
        
        .col = 5
        .Text = 0
                    
        prgBar.Value = prgBar.Value + 1
        rs.MoveNext
    Loop
    rs.Close

End With



'Carga y Compra Proceso Anterior

strSQL = "select S.cedula,S.nombre, isnull(sum(D.abono),0) as Monto" _
       & " from creditos_Dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " where D.tcon in('PLA','1')  and D.FechaP = " & txtProceso.Text & " and D.codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
       & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " group by S.cedula,S.nombre"
Call OpenRecordSet(rs, strSQL)

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

curMntAnterior = 0

With vGrid
    Do While Not rs.EOF
        pRow = fxRevisaCedula(rs!Cedula)
        If pRow = 0 Then
            'Nueva
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .col = 1
            .Text = CStr(Trim(rs!Cedula))
            
            .col = 2
            .Text = CStr(Trim(rs!Nombre))
            
            .col = 3
            .Text = 0
            
            .col = 4
            .Text = CCur(rs!Monto)
            
            curMntAnterior = curMntAnterior + rs!Monto
        Else
            .Row = pRow
            .col = 4
            .Text = CCur(rs!Monto)
            
            curMntAnterior = curMntAnterior + rs!Monto
                
        End If
                    
        .col = 5
        .Text = 0
                    
        prgBar.Value = prgBar.Value + 1
        rs.MoveNext
    Loop
    rs.Close


'Calculando Diferencias
prgBar.Value = 1
prgBar.Max = .MaxRows

For x = 1 To .MaxRows
  .Row = x
  .col = 3
  tmpActual = CCur(.Text)
  .col = 4
  tmpAnterior = CCur(.Text)
  .col = 5
  .Text = tmpActual - tmpAnterior
  prgBar.Value = x
Next x

End With





'Totales
txtMonto(0).Text = Format(curMonto, "Standard")
txtMonto(1).Text = Format(curMntAnterior, "Standard")
txtMonto(2).Text = Format(CCur(txtMonto(0).Text) - CCur(txtMonto(1).Text), "Standard")

txtCasos.Text = vGrid.MaxRows


Me.MousePointer = vbDefault

prgBar.Visible = False

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto(0).Text = 0
    txtMonto(1).Text = 0
    txtMonto(2).Text = 0
    txtCasos.Text = 0
    prgBar.Visible = False
End Sub

Private Sub sbProcesar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long, vInstitucion As Long, vCodigo As String
Dim vFecha As Date, vPrimera As Long

Dim pCedula As String, pNombre As String, pMovimiento As String, pMonto As Currency

vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
vCodigo = fxCodigoCbo(cboCliente)

vFecha = fxFechaServidor
vPrimera = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)

With vGrid
    
    For lng = 1 To .MaxRows
    
       .Row = lng
       .col = 1
       pCedula = Trim(.Text)
       .col = 2
       pNombre = Trim(.Text)
       .col = 3
       pMonto = CCur(.Text)
       .col = 4
       pMovimiento = Mid(Trim(.Text), 1, 1)
       
       If .Text <> "Error" Then
       
       
       
       strSQL = "select isnull(count(*),0) as existe from socios where cedula = '" & pCedula & "'"
       Call OpenRecordSet(rs, strSQL)
       
       If rs!Existe = 0 Then
            strSQL = "insert socios(id_promotor,cedula,cod_institucion,cod_departamento,cod_seccion,cod_profesion" _
                   & ",cod_sector,FechaIngreso,EstadoActual,Nombre) values(" _
                   & "1,'" & pCedula & "'," & vInstitucion & ",'','',1,1,'" & Format(vFecha, "yyyy/mm/dd") & "','N','" & pNombre & "')"
            Call ConectionExecute(strSQL)
            
            strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte) values('" & pCedula & "',0,0)"
            Call ConectionExecute(strSQL)
       End If
       rs.Close
        
       Select Case pMovimiento
         Case "I", "C"
                strSQL = "select id_solicitud from reg_creditos" _
                       & " where cedula = '" & pCedula & "' and estado = 'A' and plazo = 999 and codigo = '" & vCodigo & "'"
                Call OpenRecordSet(rs, strSQL)
                If rs.EOF And rs.BOF Then
                        'Insertar la operacion
                        strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
                               & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
                               & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
                               & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
                               & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol) values('" & vCodigo & "',1,'" _
                               & pCedula & "'," & pMonto & "," & pMonto & ",0," & pMonto & ",0,0," & pMonto & "," & pMonto & ",0,0,999" _
                               & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "'," _
                               & "'" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
                               & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
                               & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','N'" _
                               & ",'N','OT','',0,1,0,'PROCESO AUTOMATICO : REMESAS','A'," & vPrimera _
                               & "," & GLOBALES.glngFechaCR & ",'F')"
                        Call ConectionExecute(strSQL)
                 Else
                        'Actualiza la cuota
                        strSQL = "update reg_creditos set montoapr = " & pMonto & ",saldo = " & pMonto & ",cuota = " & pMonto _
                               & ",fechasol = dbo.MyGetdate() where id_solicitud = " & rs!Id_Solicitud
                        Call ConectionExecute(strSQL)
                 End If
                 rs.Close
         Case "E"
                'Excluye la Operacion
                strSQL = "update reg_creditos set Estado = 'C', Saldo = 0,fechasol = dbo.MyGetdate()" _
                       & " where cedula = '" & pCedula & "' and estado = 'A' and plazo = 999 and codigo = '" & vCodigo & "'"
                Call ConectionExecute(strSQL)
       
       End Select
        
        
     End If 'Error
     
    
    Next lng


End With

Me.MousePointer = vbDefault

MsgBox "Deducciones Aplicadas Satisfactoriamente...", vbInformation

vGrid.MaxRows = 0

End Sub


Private Sub sbFormatoSIF()
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
Dim fnFile, vFechaProceso As Long


vFecha = fxFechaServidor
fnFile = FreeFile


If Not IsNumeric(txtProceso.Text) Then
  MsgBox "Fecha de Proceso no es válida..", vbExclamation
  Exit Sub
End If


vFechaProceso = txtProceso.Text

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\" & SIFGlobal.fxCodText(cboCliente.Text)
MkDir SIFGlobal.DirectorioDeResultados & "\" & SIFGlobal.fxCodText(cboCliente.Text) & "\" & vFechaProceso


vRuta = SIFGlobal.DirectorioDeResultados & "\" & SIFGlobal.fxCodText(cboCliente.Text) & "\" & vFechaProceso


vArchivo = vFechaProceso & "[Comparacion]_" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
          & " " & cboInstitucion.Text & " [SIF].txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError


Open vTempo For Output As #fnFile  ' Create file name.

For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.col = 1
 vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 15)
 vGrid.col = 2
 vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 50)
 vGrid.col = 3
 vCadena = vCadena & Format(vGrid.Text, "000000000.00")
 vGrid.col = 4
 vCadena = vCadena & Format(vGrid.Text, "000000000.00")
 vGrid.col = 5
 vCadena = vCadena & Format(vGrid.Text, "000000000.00")

 Print #fnFile, vCadena
Next i

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "buscar"
'    If CLng(txtProceso.Text) > 200805 Then
'       MsgBox "Este Opcion del sistema es temporal y fue creada para idenficar las diferencias " _
'              & " entre el proceso anterior de reporte de deducciones por fecha de proceso " _
'              & " vrs el actual que es por numero de comprobante de planilla" _
'              & " [POR TANTO SOLO UTILICE FECHAS DE PROCESO ENTRE 200712-200805]", vbExclamation
'       Exit Sub
'    End If
  
    Call sbCargaDeducciones
  Case "reporte"
    MsgBox "No hay registro para procesar", vbExclamation
  Case "archivo"
    If vGrid.MaxRows <= 0 Then
       MsgBox "No existen datos para procesar...!", vbExclamation
       Exit Sub
    End If
     Call sbFormatoSIF
    
End Select

End Sub




