VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auxiliar de Planillas : SIF.A"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPrincipal.frx":169B2
   ScaleHeight     =   4545
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   855
      Left            =   5640
      Picture         =   "frmPrincipal.frx":2D364
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
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
      TabIndex        =   7
      Text            =   "1"
      Top             =   2280
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
      TabIndex        =   6
      Top             =   2280
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
      TabIndex        =   5
      Top             =   1920
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
      TabIndex        =   4
      Top             =   3000
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
      TabIndex        =   3
      Top             =   2760
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CommandButton cmdArchivo 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   6000
      Picture         =   "frmPrincipal.frx":424C6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Archivo de Planillas"
      Top             =   960
      Width           =   735
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
      Top             =   960
      Width           =   4695
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
      Caption         =   "Carga de Planillas al Sistema SIF.A"
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   2280
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
      TabIndex        =   11
      Top             =   2280
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
      TabIndex        =   10
      Top             =   1920
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
      TabIndex        =   9
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   6720
      X2              =   0
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   6720
      X2              =   0
      Y1              =   1800
      Y2              =   1800
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
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmPrincipal"
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

Private Sub cmdArchivo_Click()
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

Private Sub cmdProcesar_Click()
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

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
