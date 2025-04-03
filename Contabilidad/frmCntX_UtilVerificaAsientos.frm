VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCntX_UtilVerificaAsientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Verificación de Asientos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10965
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.CheckBox chkCorrige 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   7080
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Corregir Inconsistencias Menores Automáticamente?"
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
      Value           =   1
      Alignment       =   1
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   7875
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   615
      Left            =   9120
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_UtilVerificaAsientos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdVerificar 
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Verificar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_UtilVerificaAsientos.frx":07BC
   End
   Begin XtremeSuiteControls.CheckBox chkAsientoSinDetalle 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7320
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8488
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Elimina Asientos sin detalle contable?"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txt 
      Height          =   6015
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10695
      _Version        =   1441792
      _ExtentX        =   18865
      _ExtentY        =   10610
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Verificación de Asientos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_UtilVerificaAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
On Error GoTo vError

With Printer
 Printer.Print txt
 .NewPage
 .EndDoc
End With

Exit Sub

vError:

End Sub



Private Sub cmdVerificar_Click()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

prgBar.Value = 1
prgBar.Max = 8

txt = ">> Revisión de Integridad"
DoEvents



strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 1"
Call ConectionExecute(strSQL)



txt.Text = txt.Text & vbCrLf & vbCrLf & ">> Asientos Desbalanceados"
DoEvents

prgBar.Value = 2

strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 2"
Call ConectionExecute(strSQL)

prgBar.Value = 3

strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 3, " & chkCorrige.Value
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 txt.Text = txt.Text & vbCrLf & rs!Mensaje
 rs.MoveNext
Loop
rs.Close
    


txt.Text = txt.Text & vbCrLf & vbCrLf & ">> Afectación en Balance Diferente a la Fecha del Asiento"
DoEvents
prgBar.Value = 4

strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 4"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 txt.Text = txt.Text & vbCrLf & rs!Mensaje
 rs.MoveNext
Loop
rs.Close
    



'Asientos con Cuentas de Mayor recibiendo movimientos

txt.Text = txt.Text & vbCrLf & vbCrLf & ">> Asientos con Cuentas de Mayor recibiendo movimientos"
DoEvents

prgBar.Value = 5

strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 5"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 txt.Text = txt.Text & vbCrLf & rs!Mensaje
 rs.MoveNext
Loop
rs.Close



'Asientos sin Lineas de Detalle
txt.Text = txt.Text & vbCrLf & vbCrLf & ">> Asientos sin Lineas de Detalle"
DoEvents
prgBar.Value = 6

strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 6, " & chkAsientoSinDetalle.Value
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 txt.Text = txt.Text & vbCrLf & rs!Mensaje
 rs.MoveNext
Loop
rs.Close


'No Existe la Cuenta Madre en el Catalogo
txt.Text = txt.Text & vbCrLf & vbCrLf & ">> No Existe la Cuenta Madre en el Catalogo"
DoEvents

prgBar.Value = 7

strSQL = "exec spCntX_Asientos_Revision_Integral " & gCntX_Parametros.CodigoConta & ", " & gCntX_Parametros.PeriodoAnio _
       & ", " & gCntX_Parametros.PeriodoMes & ", '" & glogon.Usuario & "', 7"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 txt.Text = txt.Text & vbCrLf & rs!Mensaje
 rs.MoveNext
Loop
rs.Close

prgBar.Value = 8

txt = txt & vbCrLf & vbCrLf & "--> ASIENTOS VERIFICADOS! " _
    & vbCrLf & "PERIODO {MES-AÑO} : " & gCntX_Parametros.PeriodoMes & "-" & gCntX_Parametros.PeriodoAnio


Me.MousePointer = vbDefault
MsgBox "Verificación Finalizada...", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
vModulo = 20


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
