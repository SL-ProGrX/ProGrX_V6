VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_RetencionDeducciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Deducciones Aplicadas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10260
   Icon            =   "frmCR_RetencionDeducciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   10260
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   10
      Top             =   7224
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2520
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
            Picture         =   "frmCR_RetencionDeducciones.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducciones.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducciones.frx":D0D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   264
      Left            =   8040
      TabIndex        =   6
      Top             =   1440
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
      Height          =   4452
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   8292
      _Version        =   524288
      _ExtentX        =   14626
      _ExtentY        =   7853
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_RetencionDeducciones.frx":13932
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      Appearance      =   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   312
      Left            =   2640
      TabIndex        =   12
      Top             =   240
      Width           =   6852
      _Version        =   1245185
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2640
      TabIndex        =   13
      Top             =   600
      Width           =   6852
      _Version        =   1245185
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboFormato 
      Height          =   312
      Left            =   2640
      TabIndex        =   14
      Top             =   1080
      Width           =   6852
      _Version        =   1245185
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2640
      TabIndex        =   15
      Top             =   1440
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   4800
      TabIndex        =   16
      Top             =   1440
      Width           =   1212
      _Version        =   1245185
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   6720
      TabIndex        =   17
      Top             =   1440
      Width           =   1212
      _Version        =   1245185
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   6720
      TabIndex        =   18
      Top             =   6840
      Width           =   1572
      _Version        =   1245185
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   312
      Left            =   8280
      TabIndex        =   19
      Top             =   6840
      Width           =   1212
      _Version        =   1245185
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Formato "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   8520
      TabIndex        =   4
      Top             =   6600
      Width           =   972
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   1440
      Width           =   732
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
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
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmCR_RetencionDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
End Sub


Private Sub cboTipo_Click()

lblFecha.Item(0).Visible = False
lblFecha.Item(1).Visible = False
lblFecha.Item(2).Visible = False

dtpCorte.Visible = False
dtpInicio.Visible = False

txtProceso.Visible = False

If Mid(cboTipo.Text, 1, 1) = "P" Then
    lblFecha.Item(2).Visible = True
    txtProceso.Visible = True
Else
    lblFecha.Item(0).Visible = True
    lblFecha.Item(1).Visible = True
    
    dtpCorte.Visible = True
    dtpInicio.Visible = True
End If

End Sub

Private Sub Form_Load()
Dim strSQL As String

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

strSQL = "select rtrim(codigo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from catalogo where retencion = 'S' and activo = 1" _
       & " order by descripcion"
Call sbCbo_Llena_New(cboCliente, strSQL, False, True)

strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1 order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


cboFormato.AddItem "01 - Formato de Salida : Sistema ProGrX"
cboFormato.AddItem "02 - Formato de Salida : INTEGRA"
cboFormato.AddItem "03 - Formato de Salida : CCSS"
cboFormato.AddItem "04 - Formato de Salida : SPA"


cboFormato.Text = "01 - Formato de Salida : Sistema ProGrX"

cboTipo.AddItem "Fechas"
cboTipo.AddItem "Proceso"
cboTipo.Text = "Proceso"

txtProceso.Text = GLOBALES.glngFechaCR

vGrid.MaxCols = 3
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
Dim strSQL As String, pRow As Long

On Error GoTo vError

If dtpInicio.Value > dtpCorte.Value And Mid(cboTipo.Text, 1, 1) = "F" Then
   MsgBox "Las Fechas no son válidas...", vbExclamation
   Exit Sub
End If

If cboCliente.ListCount <= 0 Then Exit Sub

If cboInstitucion.Text <> "TODOS" Then
    If cboInstitucion.ListCount <= 0 Then Exit Sub
End If

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0

'Sin Plan de Pagos
If GLOBALES.SysPlanPagos = 0 Then
    If Mid(cboTipo.Text, 1, 1) = "F" Then
            strSQL = "select S.cedula,S.nombre, isnull(sum(D.abono),0) as Monto" _
                   & " from creditos_Dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
                   & " inner join socios S on R.cedula = S.cedula" _
                   & " where D.tcon in('PLA','1')  and D.fechas between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and D.codigo = '" & cboCliente.ItemData(cboCliente.ListIndex)
    Else
            strSQL = "select S.cedula,S.nombre, isnull(sum(D.abono),0) as Monto" _
                   & " from creditos_Dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
                   & " inner join socios S on R.cedula = S.cedula" _
                   & " where D.tcon in('PLA','1') and D.NCon like '" & txtProceso.Text & "%' and D.codigo = '" & cboCliente.ItemData(cboCliente.ListIndex)
    End If
End If


'Con Plan de Pagos
If GLOBALES.SysPlanPagos = 1 Then
    If Mid(cboTipo.Text, 1, 1) = "F" Then
            strSQL = "select S.cedula,S.nombre, isnull(sum(D.MOV_MONTO),0) as Monto" _
                   & " from CRD_OPERACION_TRANSAC D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
                   & " inner join socios S on R.cedula = S.cedula" _
                   & " where D.TIPO_DOCUMENTO = 'PLA' and D.MOV_FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & " 00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
                   & " and D.codigo = '" & cboCliente.ItemData(cboCliente.ListIndex)
    Else
            strSQL = "select S.cedula,S.nombre, isnull(sum(D.MOV_MONTO),0) as Monto" _
                   & " from CRD_OPERACION_TRANSAC D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
                   & " inner join socios S on R.cedula = S.cedula" _
                   & " where D.TIPO_DOCUMENTO = 'PLA' and D.NUM_COMPROBANTE like '" & txtProceso.Text _
                   & "%' and D.codigo = '" & cboCliente.ItemData(cboCliente.ListIndex)
    End If
End If




If cboInstitucion.Text = "TODOS" Then
   strSQL = strSQL & "' group by S.cedula,S.nombre"
Else
   strSQL = strSQL & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
               & " group by S.cedula,S.nombre"
End If

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
            
            curMonto = curMonto + rs!Monto
                        
            prgBar.Value = prgBar.Value + 1
            rs.MoveNext
    Loop
    rs.Close

End With



'Sin Plan de Pagos
If GLOBALES.SysPlanPagos = 0 Then
        'Cargando Morosidad
        If Mid(cboTipo.Text, 1, 1) = "F" Then
                strSQL = "select S.cedula,S.nombre, isnull(sum(D.abAmortiza),0) as Monto" _
                       & " from Morosidad D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
                       & " inner join socios S on R.cedula = S.cedula" _
                       & " where D.tcon = '1' and D.Fecult between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and D.codigo = '" & fxCodigoCbo(cboCliente)
        Else
                strSQL = "select S.cedula,S.nombre, isnull(sum(D.abAmortiza),0) as Monto" _
                       & " from Morosidad D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
                       & " inner join socios S on R.cedula = S.cedula" _
                       & " where D.tcon = '1' and D.NCon like '" & txtProceso.Text & "%' and D.codigo = '" & fxCodigoCbo(cboCliente)
        End If
        
        If cboInstitucion.Text = "TODOS" Then
           strSQL = strSQL & "' and D.estado = 'C' group by S.cedula,S.nombre"
        
        Else
           strSQL = strSQL & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                       & " and D.estado = 'C' group by S.cedula,S.nombre"
        End If
        
        
        
        
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
                            
                prgBar.Value = prgBar.Value + 1
                rs.MoveNext
            Loop
            rs.Close
        
        End With

End If 'Morosidad sin Plan de Pagos

'Totales
txtMonto.Text = Format(curMonto, "Standard")
txtCasos.Text = Format(vGrid.MaxRows, "###,###,##0")


Me.MousePointer = vbDefault

prgBar.Visible = False

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
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

If Mid(cboTipo.Text, 1, 1) = "F" Then
    vFechaProceso = Year(dtpInicio.Value) & Format(Month(dtpInicio.Value), "00")
Else
    vFechaProceso = txtProceso.Text
End If



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex)
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso


vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso


If Mid(cboTipo.Text, 1, 1) = "F" Then
    vArchivo = "SIF-" & Year(vFecha) & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
              & "-" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & "." & cboInstitucion.Text & ".txt"

Else
    vArchivo = vFechaProceso & "_" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & " " & cboInstitucion.Text & " [ProGrX].txt"
End If


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

Private Sub sbFormatoIntegra()
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
Dim fnFile, vFechaProceso As Long


vFecha = fxFechaServidor
fnFile = FreeFile

If Not IsNumeric(txtProceso.Text) Then
  MsgBox "Fecha de Proceso no es válida..", vbExclamation
  Exit Sub
End If

If Mid(cboTipo.Text, 1, 1) = "F" Then
    vFechaProceso = Year(dtpInicio.Value) & Format(Month(dtpInicio.Value), "00")
Else
    vFechaProceso = txtProceso.Text
End If

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex)
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso


vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso


If Mid(cboTipo.Text, 1, 1) = "F" Then
    vArchivo = "INTEGRA-" & Year(vFecha) & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
              & "-" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & "." & cboInstitucion.Text & ".txt"

Else
    vArchivo = vFechaProceso & "_" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & " " & cboInstitucion.Text & " [INTEGRA].txt"
End If

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
 vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "I", "0", 10) & vbTab & "102000115" & vbTab
 vGrid.col = 3
 vCadena = vCadena & Round(CCur(vGrid.Text), 0) & vbTab & Format(dtpCorte.Value, "dd/mm/yyyy")

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



Private Sub sbFormatoCCSS()
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
Dim fnFile, vFechaProceso As Long


vFecha = fxFechaServidor
fnFile = FreeFile

If Not IsNumeric(txtProceso.Text) Then
  MsgBox "Fecha de Proceso no es válida..", vbExclamation
  Exit Sub
End If

If Mid(cboTipo.Text, 1, 1) = "F" Then
    vFechaProceso = Year(dtpInicio.Value) & Format(Month(dtpInicio.Value), "00")
Else
    vFechaProceso = txtProceso.Text
End If

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex)
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso


vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso

If Mid(cboTipo.Text, 1, 1) = "F" Then
    vArchivo = "CCSS-" & Year(vFecha) & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
              & "-" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & "." & cboInstitucion.Text & ".txt"

Else
    vArchivo = vFechaProceso & "_" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & " " & cboInstitucion.Text & " [CCSS].txt"
End If

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
 vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "I", "0", 11) & "3464025252      "
 vGrid.col = 3
 vCadena = vCadena & SIFGlobal.fxStringRelleno(Format((CCur(vGrid.Text) * 100), "################"), "I", "0", 13) & SIFGlobal.fxStringRelleno("", "D", "0", 47)
 vGrid.col = 2
 vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 31) & "."
 
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


Private Sub sbFormatoSPA()
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
Dim fnFile, vFechaProceso As Long


vFecha = fxFechaServidor
fnFile = FreeFile

If Not IsNumeric(txtProceso.Text) Then
  MsgBox "Fecha de Proceso no es válida..", vbExclamation
  Exit Sub
End If

If Mid(cboTipo.Text, 1, 1) = "F" Then
    vFechaProceso = Year(dtpInicio.Value) & Format(Month(dtpInicio.Value), "00")
Else
    vFechaProceso = txtProceso.Text
End If


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex)
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso


vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\" & vFechaProceso

If Mid(cboTipo.Text, 1, 1) = "F" Then
    vArchivo = "SPA-" & Year(vFecha) & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
              & "-" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & "." & cboInstitucion.Text & ".txt"

Else
    vArchivo = vFechaProceso & "_" & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
              & " " & cboInstitucion.Text & " [SPA].txt"
End If


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
 vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "I", "0", 10)
 vGrid.col = 2
 vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 28) & "04573500439"
 vGrid.col = 3
 vCadena = vCadena & SIFGlobal.fxStringRelleno(Format((CCur(vGrid.Text) * 100), "################"), "I", "0", 8) _
         & "Q01   000000000" & Format(dtpCorte.Value, "yyyymmdd") & "5535511"
 
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
    Call sbCargaDeducciones
  Case "reporte"
    MsgBox "No hay registro para procesar", vbExclamation
  Case "archivo"
    If vGrid.MaxRows <= 0 Then
       MsgBox "No existen datos para procesar...!", vbExclamation
       Exit Sub
    End If
    Select Case Mid(cboFormato, 1, 2)
       Case "01" 'SIF
          Call sbFormatoSIF
       Case "02" 'Integra
          Call sbFormatoIntegra
       Case "03" 'CCSS
          Call sbFormatoCCSS
       Case "04" 'SPA
          Call sbFormatoSPA
         
    End Select
    
End Select

End Sub


