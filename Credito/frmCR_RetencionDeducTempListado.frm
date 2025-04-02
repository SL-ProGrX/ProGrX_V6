VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCR_RetencionDeducTempListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparación Fecha de Proceso (Listado)"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_RetencionDeducTempListado.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   13035
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   1
      Top             =   6564
      Width           =   13032
      _ExtentX        =   22992
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12735
      _Version        =   524288
      _ExtentX        =   22463
      _ExtentY        =   8705
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
      MaxCols         =   499
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_RetencionDeducTempListado.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   528
      Left            =   9120
      TabIndex        =   4
      Top             =   600
      Width           =   456
      _ExtentX        =   794
      _ExtentY        =   926
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
   Begin MSComctlLib.ImageList ImageListX 
      Left            =   360
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTempListado.frx":7063
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTempListado.frx":D8C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTempListado.frx":14127
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionDeducTempListado.frx":1A989
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   312
      Left            =   9960
      TabIndex        =   7
      Top             =   960
      Width           =   2724
      _ExtentX        =   4815
      _ExtentY        =   556
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
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
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
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   10320
      X2              =   0
      Y1              =   1440
      Y2              =   1440
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
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmCR_RetencionDeducTempListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

strSQL = "select codigo + ' - ' + descripcion as ItmX from catalogo where retencion = 'S' and activo = 1"
Call sbLlenaCbo(cboCliente, strSQL, False, False)

txtArchivo.Text = ""

vGrid.MaxCols = 8
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

With vGrid
   prgBar.Max = .MaxRows * (.MaxCols - 2)
   prgBar.Value = 1
    For lng = 1 To .MaxRows
    
       .Row = lng
       .col = 1
       vCedula = .Text
       
       .col = 3
       .Text = fxDiferenciaMes(200712, vCedula)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       .col = 4
       .Text = fxDiferenciaMes(200801, vCedula)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       .col = 5
       .Text = fxDiferenciaMes(200802, vCedula)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       .col = 6
       .Text = fxDiferenciaMes(200803, vCedula)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       .col = 7
       .Text = fxDiferenciaMes(200804, vCedula)
        prgBar.Value = prgBar.Value + 1
        DoEvents
       
       .col = 8
       .Text = fxDiferenciaMes(200805, vCedula)
       If prgBar.Value < prgBar.Max Then
        prgBar.Value = prgBar.Value + 1
        DoEvents
       End If
    
    Next lng


End With

MsgBox "Diferencias identificadas satisfactoriamente...", vbInformation

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


Private Function fxDiferenciaMes(pProceso As Long, pCedula As String) As Double
Dim strSQL As String, rs As New ADODB.Recordset
Dim tmpActual As Currency, tmpAnterior As Currency

tmpActual = 0
tmpAnterior = 0


strSQL = "select isnull(sum(D.abono),0) as Monto" _
       & " from creditos_Dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " where D.tcon in('PLA','1') and D.NCon = '" & pProceso & "' and D.codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
       & "' and S.cedula = '" & pCedula & "'"
      
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    tmpActual = rs!Monto
End If
rs.Close


strSQL = "select isnull(sum(D.abAmortiza),0) as Monto" _
       & " from Morosidad D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " where D.tcon in('PLA','1') and D.NCon ='" & pProceso & "' and D.codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
       & "' and D.estado = 'C' and S.cedula = '" & pCedula & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    tmpActual = tmpActual + rs!Monto
End If
rs.Close


'Carga y Compra Proceso Anterior

strSQL = "select isnull(sum(D.abono),0) as Monto" _
       & " from creditos_Dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " where D.tcon in('PLA','1') and D.FechaP = " & pProceso & " and D.codigo = '" & SIFGlobal.fxCodText(cboCliente.Text) _
       & "' and S.cedula = '" & pCedula & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    tmpAnterior = rs!Monto
End If
rs.Close

 
fxDiferenciaMes = tmpActual - tmpAnterior

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



