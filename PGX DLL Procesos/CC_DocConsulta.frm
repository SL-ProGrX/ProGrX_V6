VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCC_DocConsulta 
   Caption         =   "Consulta de Documentos - Afectaciones a Creditos"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   Icon            =   "CC_DocConsulta.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "CC_DocConsulta.frx":000C
   ScaleHeight     =   8910
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
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
      Left            =   9240
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Imprimir"
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
      Left            =   7800
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtNumDoc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cboTipo 
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
      ItemData        =   "CC_DocConsulta.frx":685E
      Left            =   2040
      List            =   "CC_DocConsulta.frx":6860
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7815
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   12135
      _Version        =   524288
      _ExtentX        =   21405
      _ExtentY        =   13785
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
      SpreadDesigner  =   "CC_DocConsulta.frx":6862
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCC_DocConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTcon As Integer, i As Integer
Dim curIntC As Currency, curIntM As Currency, curAmortiza As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError


curIntC = 0
curIntM = 0
curAmortiza = 0

Select Case cboTipo.Text
   Case "Recibo"
      vTcon = 2
   Case "Nota de Credito"
      vTcon = 7
   Case "Nota de Debito"
      vTcon = 8
   Case "Deposito"
      vTcon = 6
   Case "Planilla"
      vTcon = 1
End Select

vGrid.MaxRows = 0
vGrid.MaxCols = 9

strSQL = "select S.cedula,S.nombre,D.id_solicitud,D.codigo,D.intcp,D.amortiza,I.descripcion as Institucion" _
       & " from creditos_dt D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " Where D.tcon = '" & vTcon & "' And D.ncon = '" & txtNumDoc _
       & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 1
  vGrid.Text = "Ab.Ord./Ext"
  vGrid.Col = 2
  vGrid.Text = CStr(rs!Id_solicitud)
  vGrid.Col = 3
  vGrid.Text = CStr(rs!Codigo)
  vGrid.Col = 4
  vGrid.Text = Format(rs!intcp, "Standard")
  vGrid.Col = 5
  vGrid.Text = Format(0, "Standard")
  vGrid.Col = 6
  vGrid.Text = Format(rs!amortiza, "Standard")
  vGrid.Col = 7
  vGrid.Text = CStr(rs!Cedula)
  vGrid.Col = 8
  vGrid.Text = CStr(rs!Nombre)
  vGrid.Col = 9
  vGrid.Text = CStr(rs!Institucion)
  
  curAmortiza = curAmortiza + rs!amortiza
  curIntC = curIntC + rs!intcp
  
  rs.MoveNext
Loop
rs.Close


'MORATORIOS
strSQL = "select S.cedula,S.nombre,D.id_solicitud,D.codigo,D.abintc,D.abintm,D.abamortiza,I.descripcion as Institucion" _
       & " from morosidad D inner join reg_creditos R on D.id_solicitud = R.id_solicitud" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " Where D.tcon = '" & vTcon & "' And D.ncon = '" & txtNumDoc _
       & "' and D.estado in('C','N')"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 1
  vGrid.Text = "Ab.Mora"
  vGrid.Col = 2
  vGrid.Text = CStr(rs!Id_solicitud)
  vGrid.Col = 3
  vGrid.Text = CStr(rs!Codigo)
  vGrid.Col = 4
  vGrid.Text = Format(rs!abIntC, "Standard")
  vGrid.Col = 5
  vGrid.Text = Format(rs!abIntM, "Standard")
  vGrid.Col = 6
  vGrid.Text = Format(rs!abAmortiza, "Standard")
  vGrid.Col = 7
  vGrid.Text = CStr(rs!Cedula)
  vGrid.Col = 8
  vGrid.Text = CStr(rs!Nombre)
  vGrid.Col = 9
  vGrid.Text = CStr(rs!Institucion)
  
  curAmortiza = curAmortiza + rs!abAmortiza
  curIntC = curIntC + rs!abIntC
  curIntM = curIntM + rs!abIntM
  
  rs.MoveNext
Loop
rs.Close


  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  For i = 3 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = "________________"
  Next

  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 3
  vGrid.Text = "Totales :"
  vGrid.Col = 4
  vGrid.Text = Format(curIntC, "Standard")
  vGrid.Col = 5
  vGrid.Text = Format(curIntM, "Standard")
  vGrid.Col = 6
  vGrid.Text = Format(curAmortiza, "Standard")
  vGrid.Col = 7
  vGrid.Text = Format(curIntC + curIntM + curAmortiza, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdExcel_Click()
Dim x As String

x = vGrid.ExportToExcel("C:\SIF\Doc_" & cboTipo.Text & "_" & txtNumDoc.Text & ".xls", "Sheet 1", "C:\LOGFILE.TXT")
MsgBox "Archivo Generado Automaticamente en C:\SIF\Doc_" & cboTipo.Text & "_" & txtNumDoc.Text & ".xls", vbInformation

End Sub

Private Sub cmdReporte_Click()


Me.MousePointer = vbHourglass

vGrid.PrintFooter = "Afectación Documento : " & cboTipo.Text & " [ # " & txtNumDoc.Text & " ]   Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
vGrid.PrintHeader = "Documento : " & cboTipo.Text & " [ # " & txtNumDoc.Text & " ] "
If vGrid.MaxCols > 5 Then
    vGrid.PrintOrientation = PrintOrientationLandscape
Else
    vGrid.PrintOrientation = PrintOrientationPortrait
End If
vGrid.PrintSheet
  
Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()


vGrid.AppearanceStyle = fxGridStyle


cboTipo.AddItem "Recibo"
cboTipo.AddItem "Nota de Credito"
cboTipo.AddItem "Nota de Debito"
cboTipo.AddItem "Deposito"
cboTipo.AddItem "Planilla"

cboTipo.Text = "Recibo"

Me.Icon = Me.Picture

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 360
vGrid.Height = Me.Height - 1550

End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then cmdBuscar.SetFocus
End Sub
