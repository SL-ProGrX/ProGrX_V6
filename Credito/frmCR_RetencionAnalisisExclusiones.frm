VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_RetencionAnalisisExclusiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rebajos Realizados de Exclusiones Reportadas"
   ClientHeight    =   7452
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   10104
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_RetencionAnalisisExclusiones.frx":0000
   ScaleHeight     =   7452
   ScaleWidth      =   10104
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   840
      Width           =   6855
   End
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "Monto"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtCasos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Monto"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Monto"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboCliente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   9855
      _Version        =   524288
      _ExtentX        =   17383
      _ExtentY        =   7646
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_RetencionAnalisisExclusiones.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   7
      Top             =   7332
      Width           =   10104
      _ExtentX        =   17822
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageListX 
      Left            =   360
      Top             =   960
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionAnalisisExclusiones.frx":6F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionAnalisisExclusiones.frx":D780
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionAnalisisExclusiones.frx":13FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionAnalisisExclusiones.frx":1A844
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   528
      Left            =   9600
      TabIndex        =   17
      Top             =   840
      Width           =   456
      _ExtentX        =   804
      _ExtentY        =   931
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      ImageList       =   "ImageListX"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buscar"
            Object.ToolTipText     =   "Buscar archivos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cargar"
            Object.ToolTipText     =   "Cargar información"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   312
      Left            =   6960
      TabIndex        =   19
      Top             =   1680
      Width           =   2724
      _ExtentX        =   4805
      _ExtentY        =   550
      ButtonWidth     =   1778
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageListX"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "aplicar"
            Object.ToolTipText     =   "Aplicar Archivo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "cancelar operacion"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   1680
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   18
      Top             =   840
      Width           =   1335
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
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   15
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rebajos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aplicados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   13
      Top             =   6720
      Width           =   1575
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
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmCR_RetencionAnalisisExclusiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    
    If FlatScrollBar.Value = 1 Then
       txtProceso.Text = fxFechaProcesoSiguiente(txtProceso.Text)
    Else
       txtProceso.Text = fxFechaProcesoAnterior(txtProceso.Text)
    End If
    
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String


Me.Icon = Me.Picture
vGrid.AppearanceStyle = fxGridStyle

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True


strSQL = "select codigo + ' - ' + descripcion as ItmX from catalogo where retencion = 'S' and activo = 1"
Call sbLlenaCbo(cboCliente, strSQL, False, False)

strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)

txtArchivo.Text = ""

txtProceso.Text = GLOBALES.glngFechaCR

vGrid.MaxCols = 5
vGrid.MaxRows = 0
End Sub



Private Sub sbCargaListado()
Dim i As Integer

On Error GoTo vError

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

        DaoControl.Connect = "Excel 8.0;"
        DaoControl.DatabaseName = txtArchivo.Text
        DaoControl.RecordSource = "SIF$"
        DaoControl.Refresh
        
        With vGrid
        
            Do While Not DaoControl.Recordset.EOF
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .col = 1
                    .Text = Trim(CStr(DaoControl.Recordset!Cedula))
                    
                    .col = 2
                    .Text = Trim(CStr(DaoControl.Recordset!Nombre))
                    
                    
                    For i = 3 To .MaxCols
                      .col = i
                      .Text = "0.00"
                    Next i
                    
              DaoControl.Recordset.MoveNext
            Loop
        End With
        

Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
End Sub

Private Sub sbProcesar()
Dim strSQL As String, lng As Long
Dim vCedula As String

Dim vTotalRebajo As Currency, vTotalCliente As Currency

vTotalCliente = 0
vTotalRebajo = 0

With vGrid
   prgBar.Max = .MaxRows * (.MaxCols - 2)
   prgBar.Value = 1
    For lng = 1 To .MaxRows
    
       .Row = lng
       .col = 1
       vCedula = .Text
       
       .col = 3
       .Text = fxRebajos(1, vCedula)
        vTotalRebajo = vTotalRebajo + CCur(.Text)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       .col = 4
       .Text = fxRebajos(2, vCedula)
        vTotalCliente = vTotalCliente + CCur(.Text)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       
       .col = 5
       .Text = fxRebajos(3, vCedula)
        vTotalCliente = vTotalCliente + CCur(.Text)
       
       If prgBar.Value < prgBar.Max Then
        prgBar.Value = prgBar.Value + 1
        DoEvents
       End If
    
    Next lng
End With

txtMonto.Item(0).Text = Format(vTotalRebajo, "Standard")
txtMonto.Item(1).Text = Format(vTotalCliente, "Standard")
txtMonto.Item(2).Text = Format(vTotalRebajo - vTotalRebajo, "Standard")
txtCasos.Text = vGrid.MaxRows

MsgBox "Rebajos identificados satisfactoriamente...", vbInformation

End Sub


Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "aplicar"
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos cargados ...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
  
  Case "cancelar"
    vGrid.MaxRows = 0
    txtArchivo.Text = ""

End Select

End Sub


Private Function fxRebajos(pTipo As Byte, pCedula As String) As Double
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMonto As Currency

vMonto = 0

Select Case pTipo
 Case 1
 'Rebajo Total
 strSQL = "select isnull(sum(Monto),0) as Monto from prm_cargado" _
        & " Where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
        & " And fecha_Proceso = " & txtProceso.Text & " and cedula = '" & pCedula & "'"

Case 2
 'Rebajo Cuenta Cliente
 strSQL = "select isnull(sum(Abono),0) as Monto from prm_creditos" _
        & " Where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
        & " And fecha_Proceso = " & txtProceso.Text & " and Codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
        & "' and cedula = '" & pCedula & "'"


Case 3
 'Rebajo Fondo de Inconsistencia
 strSQL = "select isnull(sum(Monto),0) as Monto from prm_fondo" _
        & " Where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
        & " And Proceso = " & txtProceso.Text & " and cedula = '" & pCedula & "'"

End Select

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    vMonto = rs!Monto
End If
rs.Close


fxRebajos = vMonto

End Function


Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "buscar"
        
        txtArchivo.Text = ""
        
        With Cmd
                 .InitDir = "C:\"
                 .DialogTitle = "Localice Archivo de Listado [Microsoft EXCEL 97-2003]..."
                 .Filter = "*.xls"
                 .ShowOpen
                 
                 If .FileName = "" Then
                   MsgBox "Archivo no válido...", vbExclamation
                   Exit Sub
                 End If
                 
                 If UCase(Right(.FileName, 3)) <> "XLS" Then
                   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                   Exit Sub
                 End If
                
           
         txtArchivo.Text = .FileName
        
        End With

  Case "cargar"
    Call sbCargaListado
  
End Select

End Sub




