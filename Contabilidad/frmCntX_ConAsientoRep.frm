VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCntX_ConAsientoRep 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes de Asientos Consolidados"
   ClientHeight    =   2244
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   6408
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2244
   ScaleWidth      =   6408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   170000387
      CurrentDate     =   37304
   End
   Begin VB.ComboBox cboReporte 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "frmCntX_ConAsientoRep.frx":0000
      Left            =   1320
      List            =   "frmCntX_ConAsientoRep.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   170000387
      CurrentDate     =   37304
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   4800
      TabIndex        =   7
      Top             =   1560
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCntX_ConAsientoRep.frx":0035
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consolidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCntX_ConAsientoRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReporte_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xMascara As String, iCodEmpresa As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select E.*" _
       & " from CNTX_CONTABILIDADES E inner join CNTX_CONSOLIDA_DEFINICION C" _
       & " On E.COD_CONTABILIDAD = C.COD_CONTABILIDAD" _
       & " where C.cod_consolida = " & cbo.ItemData(cbo.ListIndex)
rs.Open strSQL, glogon.Conection, adOpenStatic
  xMascara = rs!Nivel1 & rs!Nivel2 & rs!Nivel3 & rs!Nivel4 & rs!Nivel5
  iCodEmpresa = rs!COD_CONTABILIDAD
rs.Close

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ProGrX: Contabilidad"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Titulo='" & UCase(Mid(cboReporte, 5, 60)) & "'"
 .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
 .Formulas(3) = "Mascara='" & xMascara & "'"
 .Formulas(4) = "SubTitulo='Desde " & Format(dtpInicio, "yyyy/mm/dd") & " Hasta " & Format(dtpCorte, "yyyy/mm/dd") _
              & " Consolidación: " & cbo.Text & "'"
 .Connect = glogon.ConectRPT
 
 .ReportFileName = App.Path & "\ConAsientos.rpt"

 .SelectionFormula = "{CON_ASIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex) _
       & " AND {CON_ASIENTOS.FECHA} in Date(" & Year(dtpInicio.Value) & "," _
       & Month(dtpInicio.Value) & "," & Day(dtpInicio.Value) _
       & ") to Date(" & Year(dtpCorte.Value) & "," & Month(dtpCorte.Value) & "," _
       & Day(dtpCorte.Value) & ") AND {CUENTAS.COD_CONTABILIDAD} = " & iCodEmpresa

 Select Case Mid(cboReporte, 1, 2)
   Case "01" 'Todos los Asientos
     .SelectionFormula = .SelectionFormula
     
   Case "02" 'Solo Asientos Aplicados
     .SelectionFormula = .SelectionFormula & " AND {CON_ASIENTOS.APLICADO} = 'S'"
     
   Case "03" 'Asientos Sin Aplicar
     .SelectionFormula = .SelectionFormula & " AND {CON_ASIENTOS.APLICADO} = 'N'"
  End Select
  
  .PrintReport
  
End With

vError:
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

 
 vPaso = False
 
 strSQL = "select * from CNTX_CONSOLIDA_DEFINICION"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 cbo.Clear
 
 Do While Not rs.EOF
   cbo.AddItem Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
   cbo.ItemData(cbo.NewIndex) = rs!COD_CONSOLIDA
   vPaso = True
   rs.MoveNext
 Loop
 
 If vPaso Then
   rs.MoveFirst
   cbo.Text = Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
 End If
 rs.Close

 cboReporte.Clear
 cboReporte.AddItem "01 - Asientos Cargados(Apl/S.Apl)"
 cboReporte.AddItem "02 - Asientos Aplicados"
 cboReporte.AddItem "03 - Asientos Sin Aplicar"

 cboReporte.Text = "01 - Asientos Cargados(Apl/S.Apl)"
 
 dtpInicio = fxFechaServidor
 dtpCorte = dtpInicio


End Sub
