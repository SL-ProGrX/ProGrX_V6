VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUS_Accesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accesos a la Base de Datos"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "&Archivo"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   480
      Width           =   855
   End
   Begin VB.CheckBox chkVersion 
      Caption         =   "Versión"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtVersion 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox cbo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtApp 
      Height          =   315
      Left            =   720
      TabIndex        =   10
      Top             =   480
      Width           =   3375
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ingreso"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Aplicación"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Versión"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox chkUsuario 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   115277827
      CurrentDate     =   37444
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy  hh:mm:ss"
      Format          =   115277827
      CurrentDate     =   37444
   End
   Begin VB.Label Label1 
      Caption         =   "App"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Tag             =   "a"
      Top             =   480
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8760
      X2              =   0
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Corte"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmUS_Accesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUsuario_Click()
If chkUsuario.Value Then
   txtUser.SetFocus
Else
   txtUser = ""
End If
End Sub

Private Sub chkVersion_Click()
If chkVersion.Value = vbChecked Then
  cbo.Enabled = True
  txtVersion.Enabled = True
Else
  cbo.Enabled = False
  txtVersion.Enabled = False
End If
End Sub

Private Sub cmdArchivo_Click()
Dim fn, strArchivo As String, bPaso As Boolean
Dim lng As Long, i As Integer, vCadena As String

On Error GoTo vError

If lsw.ListItems.Count = 0 Then Exit Sub

fn = FreeFile
 frmContenedor.CD.InitDir = "C:\"
 frmContenedor.CD.ShowSave
 
 Open frmContenedor.CD.FileName For Output As #fn
    Print #fn, "MODULO DE SEGURIDAD: CONTROL DE ACCESOS"
    Print #fn, "SERVIDOR: " & vbTab & glogon.Servidor
    Print #fn, "BASE DE DATOS: " & vbTab & glogon.BaseDatos & vbCrLf
    
    vCadena = "PARAMETROS: INICIO: " & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:01 CORTE: " _
            & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59"
    Print #fn, vCadena

    Print #fn, "USUARIOS: " & vbTab & IIf((chkUsuario.Value = vbChecked), txtUser, "TODOS")
    Print #fn, "APLICACION: " & vbTab & IIf((Trim(txtApp) <> ""), txtApp, "TODAS")
    Print #fn, "VERSION: " & vbTab & IIf((chkVersion.Value = vbChecked), cbo.Text & " " & txtVersion, "TODAS")
    Print #fn, ""
 
    vCadena = "FECHA" & vbTab & "USUARIO" & vbTab & "APLICACION" & vbTab & "VERSION" & vbCrLf
    Print #fn, vCadena
    
    For lng = 1 To lsw.ListItems.Count
      vCadena = lsw.ListItems.Item(lng) & vbTab
      For i = 1 To lsw.ColumnHeaders.Count - 1
        vCadena = vCadena & lsw.ListItems.Item(lng).SubItems(i) & vbTab
      Next i
      Print #fn, vCadena
    Next lng

 Close #fn
 MsgBox "Información Guardada en " & frmContenedor.CD.FileName, vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "select * from accesos where fecha between '" & Format(dtpInicio, "yyyy/mm/dd") _
       & " 00:00:01' and '" & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59'"

If txtUser <> "" Then strSQL = strSQL & " and nombre = '" & txtUser & "'"

If Trim(txtApp) <> "" Then strSQL = strSQL & " and aplicacion like '" & txtApp & "%'"

If chkVersion.Value = vbChecked Then
   strSQL = strSQL & " and version " & cbo.Text & "'" & txtVersion & "'"
End If

strSQL = strSQL & " order by fecha"

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

lsw.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Format(rs!Fecha, "yyyy/mm/dd hh:mm:ss"))
     itmX.SubItems(1) = rs!nombre
     itmX.SubItems(2) = rs!aplicacion
     itmX.SubItems(3) = rs!Version & ""
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
MsgBox "Consulta Finalizada...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
'vModulo = 13

Set Me.Icon = frmContenedor.Icon
dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cbo.AddItem "="
cbo.AddItem ">"
cbo.AddItem "<"

cbo.Text = "="

End Sub


Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
gBusquedas.Convertir = "N"
gBusquedas.Columna = "Nombre"
gBusquedas.Orden = "Nombre"
gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
gBusquedas.Filtro = ""
frmBusquedas.Show vbModal
txtUser = gBusquedas.Resultado
End Sub
