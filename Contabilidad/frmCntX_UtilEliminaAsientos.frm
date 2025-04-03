VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCntX_UtilEliminaAsientos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elimina Asientos"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6270
   HelpContextID   =   1
   Icon            =   "frmCntX_UtilEliminaAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCntX_UtilEliminaAsientos.frx":000C
   ScaleHeight     =   3405
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtHasta 
      Appearance      =   0  'Flat
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
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtDesde 
      Appearance      =   0  'Flat
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
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin XtremeSuiteControls.PushButton cmdOk 
      Height          =   612
      Left            =   4680
      TabIndex        =   9
      Top             =   2520
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Elimina"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmCntX_UtilEliminaAsientos.frx":035C
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6345
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblEstatus 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image imgCalcular 
      Height          =   372
      Left            =   4560
      Picture         =   "frmCntX_UtilEliminaAsientos.frx":0CF1
      Stretch         =   -1  'True
      ToolTipText     =   "Cálcular Impacto"
      Top             =   1560
      Width           =   372
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -105
      X2              =   6240
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo Asiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblPeriodo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmCntX_UtilEliminaAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vFecha As Date, vDetalle As String, vTipo As String

Me.MousePointer = vbHourglass

lblEstatus.Visible = True
prgBar.Visible = True

vFecha = fxFechaServidor

vTipo = SIFGlobal.fxCodText(cboTipo.Text)

On Error GoTo vError

lblEstatus.Caption = "Cargando..."
lblEstatus.Refresh

strSQL = "select num_asiento,tipo_asiento,cod_contabilidad from Cntx_Asientos where anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " _
       & gCntX_Parametros.PeriodoMes & " and tipo_asiento = '" & vTipo & "'" _
       & " and num_asiento between '" & txtDesde & "' and '" & txtHasta & "'" _
       & " and fecha_aplicado is null and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and modulo = 20"

Call OpenRecordSet(rs, strSQL, 0)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
  lblEstatus.Caption = "Eliminando Asiento : " & rs!Tipo_Asiento & " - " & rs!Num_Asiento
  lblEstatus.Refresh
  
  strSQL = "delete Cntx_Asientos_detalle where num_asiento = '" & rs!Num_Asiento _
         & "' and tipo_asiento = '" & rs!Tipo_Asiento & "' and cod_contabilidad = " _
         & rs!COD_CONTABILIDAD
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete Cntx_Asientos where num_asiento = '" & rs!Num_Asiento _
         & "' and tipo_asiento = '" & rs!Tipo_Asiento & "' and cod_contabilidad = " _
         & rs!COD_CONTABILIDAD
  Call ConectionExecute(strSQL, 0)
  
  rs.MoveNext
  If prgBar.Max > prgBar.Value Then prgBar.Value = prgBar.Value + 1
Loop
rs.Close

vDetalle = "TIPO:" & vTipo & " D:" & txtDesde & " H:" & txtHasta & " AFECTA:" & prgBar.Max - 1

Call Bitacora("Elimina", vDetalle)

lblEstatus.Visible = False
prgBar.Visible = False
Me.MousePointer = vbDefault

MsgBox "Eliminación Finalizada...", vbInformation

Exit Sub

vError:
    lblEstatus.Visible = False
    prgBar.Visible = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 20

lblPeriodo.Caption = "PERIODO : " & gCntX_Parametros.PeriodoAnio & " - " & gCntX_Parametros.PeriodoMes

strSQL = "select rtrim(Tipo_Asiento) + ' - ' + rtrim(descripcion) as 'ItmX'" _
       & " from CntX_Tipos_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta

Call sbLlenaCbo(cboTipo, strSQL, False)


vError:


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub imgCalcular_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select isnull(count(*),0) as Total from Cntx_Asientos where anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " _
       & gCntX_Parametros.PeriodoMes & " and tipo_asiento = '" & SIFGlobal.fxCodText(cboTipo.Text) & "'" _
       & " and num_asiento between '" & txtDesde & "' and '" & txtHasta & "'" _
       & " and fecha_aplicado is null and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and modulo = 20"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)
   MsgBox "Total de Asientos a Eliminar : " & rs!Total, vbInformation
rs.Close

vError:

Me.MousePointer = vbDefault

End Sub



Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtHasta.SetFocus
End Sub
