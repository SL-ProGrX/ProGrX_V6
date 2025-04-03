VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_PeriodosDefinicion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Periodos"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6210
   HelpContextID   =   2005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   612
      Left            =   1560
      TabIndex        =   5
      Top             =   2520
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Aplicar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_PeriodosDefinicion.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_PeriodosDefinicion.frx":07D8
   End
   Begin XtremeSuiteControls.FlatEdit txtDesde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   852
      _Version        =   1310723
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHasta 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   2760
      TabIndex        =   8
      Top             =   1800
      Width           =   852
      _Version        =   1310723
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDesdeMes 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   612
      _Version        =   1310723
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHastaMes 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   10
      Top             =   1800
      Width           =   612
      _Version        =   1310723
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodos Contables"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   8532
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   852
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   852
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1440
      Width           =   852
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_PeriodosDefinicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicar_Click()
Dim strSQL As String, vAnio As Long, vMes As Integer

On Error GoTo vError

vAnio = txtDesde.Text
vMes = txtDesdeMes.Text

If vAnio >= CLng(txtHasta.Text) Then
  MsgBox "Los Periodos son con corte anual revise...!", vbExclamation
  Exit Sub
End If

strSQL = ""
Do While Not (vAnio = CLng(txtHasta.Text) And vMes = CLng(txtHastaMes.Text))
    
    strSQL = strSQL & Space(10) & "insert into CntX_Periodos(COD_CONTABILIDAD,anio,mes,estado,PERIODO_CORTE) values(" _
           & gCntX_Parametros.CodigoConta & "," & vAnio & "," & vMes & ",'P'" _
           & ",dbo.fxSys_FechaAnioMesToDatetime(" & vAnio & "," & vMes & "))"
    
    If vMes = 12 Then
       vMes = 1
       vAnio = vAnio + 1
    Else
       vMes = vMes + 1
    End If
    
    'Ultimo Mes
    If (vAnio = CLng(txtHasta.Text) And vMes = CLng(txtHastaMes.Text)) Then
        strSQL = strSQL & Space(10) & "insert into CntX_Periodos(COD_CONTABILIDAD,anio,mes,estado,PERIODO_CORTE) values(" _
               & gCntX_Parametros.CodigoConta & "," & vAnio & "," & vMes & ",'P'" _
               & ",dbo.fxSys_FechaAnioMesToDatetime(" & vAnio & "," & vMes & "))"
    End If
    
    If Len(strSQL) > 20000 Then
       Call ConectionExecute(strSQL, 0)
       strSQL = ""
    End If
    
Loop

'Fix Para nuevo campo de control
strSQL = strSQL & Space(10) & "UPDATE CNTX_PERIODOS SET PERIODO_CORTE = dbo.fxSys_FechaAnioMesToDatetime(anio,mes)" _
       & " WHERE PERIODO_CORTE IS NULL"
Call ConectionExecute(strSQL, 0)

MsgBox "Periodos creados satisfactoriamente!...", vbInformation

Call sbInicial

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cmdReporte_Click()
Call sbCntX_Reportes_Catalogos("Periodos")
End Sub

Private Sub sbInicial()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select dateadd(m,1, ISNULL(max(PERIODO_CORTE),GETDATE()) ) AS 'FECHA' FROM CNTX_PERIODOS" _
       & " WHERE COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL)


txtDesde.Text = Year(rs!fecha)
txtDesdeMes.Text = Month(rs!fecha)

txtHasta.Text = CLng(txtDesde.Text) + 1
If Month(rs!fecha) > 1 Then
    txtHastaMes.Text = CLng(txtDesdeMes.Text) - 1
Else
    txtHastaMes.Text = 12
    txtHasta.Text = txtDesde.Text
End If
rs.Close

Exit Sub

vError:
   
End Sub

Private Sub Form_Load()
vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

Call sbInicial
End Sub

Private Sub txtDesde_Change()
On Error GoTo vError
 
txtHasta.Text = CLng(txtDesde.Text) + 1
txtHastaMes.Text = CLng(txtDesdeMes.Text) - 1
 
vError:
End Sub

Private Sub txtDesdeMes_Change()
On Error GoTo vError

txtHasta.Text = CLng(txtDesde.Text) + 1
txtHastaMes.Text = CLng(txtDesdeMes.Text) - 1
vError:

End Sub

