VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmFNDReportesConciliacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Conciliación"
   ClientHeight    =   4500
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10128
   HelpContextID   =   7004
   Icon            =   "frmFNDReportesConciliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   10128
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   480
      Top             =   120
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8400
      TabIndex        =   0
      Top             =   480
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   6132
      _Version        =   1245187
      _ExtentX        =   10816
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkPlanes 
      Height          =   252
      Left            =   9000
      TabIndex        =   6
      Top             =   480
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
      ForeColor       =   16777215
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Width           =   6132
      _Version        =   1245187
      _ExtentX        =   10816
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   312
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboResultados 
      Height          =   312
      Left            =   5760
      TabIndex        =   14
      Top             =   2160
      Width           =   2532
      _Version        =   1245187
      _ExtentX        =   4466
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   5760
      TabIndex        =   15
      Top             =   2520
      Width           =   2532
      _Version        =   1245187
      _ExtentX        =   4466
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboRsm 
      Height          =   312
      Left            =   5760
      TabIndex        =   16
      Top             =   2880
      Width           =   2532
      _Version        =   1245187
      _ExtentX        =   4466
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   9972
      _Version        =   1245187
      _ExtentX        =   17590
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   6600
         TabIndex        =   18
         Top             =   240
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   14
         Picture         =   "frmFNDReportesConciliacion.frx":030A
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Entidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   9
      Left            =   960
      TabIndex        =   12
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   4680
      TabIndex        =   10
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   4680
      TabIndex        =   9
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   4680
      TabIndex        =   8
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDReportesConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean


Private Sub btnReporte_Click()
Call sbReporte
End Sub

Private Sub cboOperadora_Click()
txtCodigo_LostFocus
End Sub

Private Sub cboResultados_Click()
If cboResultados.Text = "Resumen Fondo" Then
  cboRsm.Enabled = True
Else
  cboRsm.Enabled = False
End If
End Sub

Private Sub sbReporte()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFondo As String

On Error GoTo vError

If cboPeriodos.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass

If chkPlanes.Value = xtpChecked Then
 vFondo = ""
Else
  vFondo = txtCodigo.Text
End If


With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Conciliación"
    
    .Connect = glogon.ConectRPT

    Select Case cboResultados.Text
       Case "Metodo Contable"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoCuentas.rpt")
            .Formulas(0) = "SUBTITULO='PERIODO: " & UCase(cboPeriodos.Text) & "'"
            .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
            .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
            .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
            .Formulas(4) = "MASCARA='3232'"

            strSQL = "select * from fnd_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)

            .SelectionFormula = "{FND_PER_CUENTAS.ANIO}=" & rs!Anio _
                              & " AND {FND_PER_CUENTAS.MES}=" & rs!Mes
            rs.Close
  
      
       
       Case "General Fondo"
       
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoDetalle.rpt")
            .Formulas(0) = "SUBTITULO='PERIODO: " & UCase(cboPeriodos.Text) & " / FILTRO : " & UCase(cboEstado.Text) & "'"
            .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
            .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
            .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
            

            
            strSQL = "select * from fnd_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
        
            .SelectionFormula = "{FND_PER_CERRADOS.ANIO}=" & rs!Anio _
                              & " AND {FND_PER_CERRADOS.MES}=" & rs!Mes
            
            rs.Close
       
       Case "Resumen Fondo"
            .Formulas(0) = "SUBTITULO='" & UCase(cboRsm.Text) & " / FILTRO : " & UCase(cboEstado.Text) & "'"
            .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
            .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
            .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
            
          Select Case cboRsm.Text
            Case "Periodo"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoResumen.rpt")
                .Formulas(0) = "SUBTITULO='" & UCase(cboRsm.Text) & " : " & UCase(cboPeriodos.Text) & " / FILTRO : " & UCase(cboEstado.Text) & "'"
                
                strSQL = "select * from fnd_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
                Call OpenRecordSet(rs, strSQL)
            
                    .SelectionFormula = "{FND_PER_CERRADOS.ANIO}=" & rs!Anio _
                                      & " AND {FND_PER_CERRADOS.MES}=" & rs!Mes
                rs.Close
           
            Case "Historico Anual vrs Mes"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoResumen.rpt")
            
            Case "Historico Anual vrs Fondo"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoResumenMes.rpt")

            
            Case "Historico Anual"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoResumenAnual.rpt")
                .SelectionFormula = "{FND_PER_CERRADOS.MES}= 12"
           
         End Select
       

       
       
    End Select
    
    If Mid(cboEstado.Text, 1, 1) <> "T" Then
       .SelectionFormula = "{FND_PER_CERRADOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
    End If
                
    If Len(vFondo) > 0 Then
       .SelectionFormula = .SelectionFormula & " AND {FND_PER_CERRADOS.COD_PLAN} = '" & vFondo & "'"
    End If
            
    If cboInstitucion.Text <> "TODOS" Then
        .SelectionFormula = .SelectionFormula & " AND {SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If

    .PrintReport
End With

vError:

Me.MousePointer = vbDefault

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_Plan
      txtCodigo_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset

vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True


cboResultados.Clear
cboResultados.AddItem "Metodo Contable"
cboResultados.AddItem "General Fondo"
cboResultados.AddItem "Resumen Fondo"

cboRsm.Clear
cboRsm.AddItem "Periodo"
cboRsm.AddItem "Historico Anual vrs Mes"
cboRsm.AddItem "Historico Anual vrs Fondo"
cboRsm.AddItem "Historico Anual"
cboRsm.Text = "Periodo"

cboResultados.Text = "Metodo Contable"

cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.AddItem "Activos"
cboEstado.AddItem "Liquidados"
cboEstado.AddItem "Bloqueados"
cboEstado.AddItem "Inactivos"
cboEstado.Text = "Todos"



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

'chkPlanes.BackColor = txtDescripcion.BackColor

Dim strSQL As String, rs As New ADODB.Recordset



strSQL = "select descripcion as 'itmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select cod_institucion as 'IdX', rtrim(descripcion) as 'ItmX' from instituciones where Activa = 1 order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

strSQL = "select * from fnd_per_historico order by anio desc,mes desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboPeriodos.AddItem rs!Anio & "-" & rs!Mes
 cboPeriodos.ItemData(cboPeriodos.ListCount - 1) = CStr(rs!id_per_historico)
 rs.MoveNext
Loop
If rs.RecordCount > 1 Then
  rs.MoveFirst
  cboPeriodos.Text = rs!Anio & "-" & rs!Mes
End If
rs.Close


Call cboOperadora_Click

End Sub




Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select Descripcion from Fnd_Planes where Cod_Operadora="
strSQL = strSQL & cboOperadora.ItemData(cboOperadora.ListIndex) & " And "
strSQL = strSQL & "Cod_Plan='" & Trim(txtCodigo) & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       txtDescripcion = Trim(!Descripcion)
    Else
       txtCodigo = ""
       txtDescripcion = ""
    End If
 .Close
End With

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   cboEstado.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub

