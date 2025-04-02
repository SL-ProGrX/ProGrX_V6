VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSIFCargaPlanilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auxiliar de Planillas : SIF.A"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmSIFCargaPlanilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSIFCargaPlanilla.frx":169B2
   ScaleHeight     =   5205
   ScaleWidth      =   6855
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.TextBox txtPago 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Text            =   "1"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtProceso 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
   End
   Begin VB.CheckBox chkCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Procesar Registros de Creditos"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkAportes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Procesar Registros de Aportes"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   3480
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   795
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   4695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFCargaPlanilla.frx":2D364
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFCargaPlanilla.frx":424D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBuscar 
      Height          =   570
      Left            =   6000
      TabIndex        =   14
      Top             =   1680
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Archivo"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbProcesar 
      Height          =   780
      Left            =   5760
      TabIndex        =   15
      Top             =   4320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1376
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Procesar"
            Key             =   "Procesar"
            Object.ToolTipText     =   "Procesar Archivo"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmSIFCargaPlanilla.frx":57648
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   960
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Picture         =   "frmSIFCargaPlanilla.frx":576D9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   6840
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Carga de Planillas de Archivos Dbase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Pago de la Planilla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   2
      Left            =   3000
      TabIndex        =   10
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   6720
      X2              =   0
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   6720
      X2              =   0
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmSIFCargaPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboInstitucion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select pr_fecha_corte from instituciones where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  txtProceso = Year(rs!pr_fecha_corte) & Format(Month(rs!pr_fecha_corte), "00")
End If
rs.Close

vError:

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_institucion,descripcion,pr_fecha_corte from instituciones"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 cboInstitucion.AddItem rs!descripcion
 cboInstitucion.ItemData(cboInstitucion.NewIndex) = rs!cod_institucion
 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim itmX As ListItem, vArchivo As String
Dim vPasa As Boolean


With frmContenedor.dlg
 .InitDir = "C:\"
 .ShowOpen
 
 If .FileName = "" Then
   MsgBox "Archivo no válido...", vbExclamation
   Exit Sub
 End If
 
 If UCase(Right(.FileName, 3)) <> "DBF" Then
   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
   Exit Sub
 End If

 txtArchivo = .FileName

End With

DaoControl.RecordSource = Dir(txtArchivo, vbArchive)
DaoControl.DatabaseName = Mid(txtArchivo, 1, Len(txtArchivo) - (Len(DaoControl.RecordSource) + 1))
DaoControl.Refresh

End Sub

Private Sub tlbProcesar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


lng = 1

strSQL = "delete prm_cargado where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and fecha_proceso = " & txtProceso & " and pago = " & txtPago
glogon.Conection.Execute strSQL

With DaoControl.Recordset

Do While Not .EOF

  lbl.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  lbl.Refresh
  
  If chkAportes.Value = vbChecked Then
    strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto)" _
           & " values(" & cboInstitucion.ItemData(cboInstitucion.ListIndex) & "," & txtPago _
           & "," & txtProceso & ",1,'" & Trim(!cedula) & "'," & CCur(!Aportes) & ")"
    glogon.Conection.Execute strSQL
  End If
  
  If chkCredito.Value = vbChecked Then
    strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto)" _
           & " values(" & cboInstitucion.ItemData(cboInstitucion.ListIndex) & "," & txtPago _
           & "," & txtProceso & ",3,'" & Trim(!cedula) & "'," & CCur(!abonos) & ")"
   If !abonos > 0 Then glogon.Conection.Execute strSQL
  End If
  
  lng = lng + 1
  .MoveNext
Loop

End With

Me.MousePointer = vbDefault

lbl.Caption = ""

MsgBox "Proceso finalizado Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub
