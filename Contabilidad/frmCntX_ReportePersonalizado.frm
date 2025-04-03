VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_RepPersonalizado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reporte Personalizado (Eliminado!)"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7110
   HelpContextID   =   26
   Icon            =   "frmCntX_ReportePersonalizado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPeriodico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Utilizar Saldos para Inventarios Periodicos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Width           =   3855
   End
   Begin VB.ComboBox cboErEspecial 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox txtSubTitulo 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtTitulo 
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
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   5055
   End
   Begin VB.ComboBox cboNivel 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   5055
   End
   Begin VB.ComboBox cboReporte 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   5055
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   4320
      TabIndex        =   11
      Top             =   3000
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   16777215
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
      Picture         =   "frmCntX_ReportePersonalizado.frx":000C
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Er.Especial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   7080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SubTitulo"
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCntX_RepPersonalizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vUltimoNivel As Integer

Private Sub cboNivel_Click()
 txtTitulo = cboReporte.Text & " [" & cboNivel.Text & "]"
End Sub

Private Function fxTituloErEspecial()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select * from er_especial where cod_er_especial = " _
      & cboErEspecial.ItemData(cboErEspecial.ListIndex) _
      & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  fxTituloErEspecial = Trim(rs!Titulo)
End If
rs.Close


Exit Function
vError:
 fxTituloErEspecial = ""
End Function

Private Sub cboReporte_Click()


If cboReporte.ItemData(cboReporte.ListIndex) = 4 Then
  cboErEspecial.Enabled = True
  txtTitulo = fxTituloErEspecial
Else
  cboErEspecial.Enabled = False
  txtTitulo = cboReporte.Text & " [" & cboNivel.Text & "]"
End If


End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String, strTitulo As String
Dim vPeriodoDesc As String

On Error GoTo vError

vPeriodoDesc = fxCntX_PeriodoDesc(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

If fxCntX_PeriodoVerifica(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes) Then
 strTitulo = txtTitulo & " PERIODO: " & vPeriodoDesc & " [PENDIENTE]"
Else
 strTitulo = txtTitulo & " PERIODO: " & vPeriodoDesc & " [CERRADO]"
End If

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ProGrX: Contabilidad"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & gCntX_Parametros.NombreEmpresa & "'"
 .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
 .Connect = glogon.ConectRPT
  
  
 Select Case cboReporte.ItemData(cboReporte.ListIndex)
    Case 1, 2 'Estado de Resultados y Balance General
        .Formulas(3) = "Estado='" & UCase(strTitulo) & "'"
        .Formulas(4) = "SubTitulo='" & txtSubTitulo & "'"
        If chkPeriodico.Value = vbChecked Then
          vPeriodico = 1
        Else
          vPeriodico = 0
        End If
        
        Call sbCntX_Estados(IIf((cboReporte.ItemData(cboReporte.ListIndex) = 1), "ER", "BG") _
                       , gCntX_Parametros.PeriodoMes, gCntX_Parametros.PeriodoAnio, cboNivel.ItemData(cboNivel.ListIndex))
        
        .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosPersonal.rpt")
        .Connect = glogon.ConectRPT
        .SelectionFormula = "{CNTX_REP_BALANCES_PERSONALIZADO.USUARIO}= '" & glogon.Usuario & "'"
    
    Case 3 'Balance de Comprobación
      Select Case cboNivel.ItemData(cboNivel.ListIndex)
        Case 1 'Resumen Original
            .Formulas(3) = "SubTitulo='" & txtSubTitulo & "'"
            strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                   & " AND {CntX_Cuentas.CUENTA_MADRE} ='' AND {vCntX_Mov_Cuentas_General.ANIO} = " _
                   & gCntX_Parametros.PeriodoAnio & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceComprobacionGeneral.rpt")
            .Formulas(4) = "GrupoCuenta = mid({CntX_Cuentas.COD_CUENTA},1," & gCntX_Parametros.Nivel1 & ")"
            .Formulas(5) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
            .SelectionFormula = strSQL
        
        Case 6, vUltimoNivel 'Detalle Original
            .Formulas(3) = "SubTitulo='" & txtSubTitulo & "'"
            strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                   & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
                   & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes
            
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceComprobacionGeneral.rpt")
            .Formulas(4) = "GrupoCuenta = mid({CntX_Cuentas.COD_CUENTA},1," & gCntX_Parametros.Nivel1 & ")"
            .Formulas(5) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
            .SelectionFormula = strSQL
  
        Case Else 'Reporte Personalizado
            .Formulas(3) = "Estado='" & UCase(strTitulo) & "'"
            .Formulas(4) = "SubTitulo='" & txtSubTitulo & "'"
            strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                   & " AND {vCntX_Mov_Cuentas_General.ANIO} = " & gCntX_Parametros.PeriodoAnio _
                   & " AND {vCntX_Mov_Cuentas_General.MES} = " & gCntX_Parametros.PeriodoMes _
                   & " AND {CntX_Cuentas.NIVEL} = " & cboNivel.ItemData(cboNivel.ListIndex)

            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_BalanceComprobacionGeneral.rpt")
            .Formulas(4) = "GrupoCuenta = mid({CntX_Cuentas.COD_CUENTA},1," & gCntX_Parametros.Nivel1 & ")"
            .Formulas(5) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
            .SelectionFormula = strSQL
      
      End Select
      
      
    Case 4 'Estado de Resultados Especial
        
        If chkPeriodico.Value = vbChecked Then
          vPeriodico = 1
        Else
          vPeriodico = 0
        End If
        
        .Formulas(3) = "Estado='" & UCase(strTitulo) & "'"
        .Formulas(4) = "SubTitulo='" & txtSubTitulo & "'"
        
        Call sbCntX_ER_Especial(cboErEspecial.ItemData(cboErEspecial.ListIndex), gCntX_Parametros.PeriodoMes, gCntX_Parametros.PeriodoAnio, cboNivel.ItemData(cboNivel.ListIndex))
        
        .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EstadoResultadosPersonal.rpt")
        .Connect = glogon.ConectRPT
        .SelectionFormula = "{CNTX_REP_BALANCES_PERSONALIZADO.USUARIO}= '" & glogon.Usuario & "'"
      
      
 End Select
 
.Action = 1
 
 vPeriodico = 0
 
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs  As New ADODB.Recordset, vPasa As Boolean
vModulo = 20


vPasa = False

cboReporte.Clear
cboReporte.AddItem "Estado de Resultados"
cboReporte.ItemData(cboReporte.NewIndex) = 1
'--
cboReporte.AddItem "Balance General"
cboReporte.ItemData(cboReporte.NewIndex) = 2
'--
cboReporte.AddItem "Balance de Comprobación"
cboReporte.ItemData(cboReporte.NewIndex) = 3
'--
cboReporte.AddItem "ER Especial"
cboReporte.ItemData(cboReporte.NewIndex) = 4

cboReporte.Text = "Estado de Resultados"
' cboReporte.Enabled = False

strSQL = "select * from CntX_Er_especial where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 vPasa = True
 cboErEspecial.AddItem rs!Descripcion & ""
 cboErEspecial.ItemData(cboErEspecial.NewIndex) = rs!cod_er_especial
 rs.MoveNext
Loop

If vPasa Then
  rs.MoveFirst
  cboErEspecial.Text = rs!Descripcion & ""
End If
rs.Close

Call cboReporte_Click

cboNivel.Clear
If gCntX_Parametros.Nivel1 > 0 Then
   cboNivel.AddItem "Nivel 1 (Resumen)"
   cboNivel.ItemData(cboNivel.NewIndex) = 1
   vUltimoNivel = 1
End If
If gCntX_Parametros.Nivel2 > 0 Then
   cboNivel.AddItem "Nivel 2"
   cboNivel.ItemData(cboNivel.NewIndex) = 2
   vUltimoNivel = 2
End If
If gCntX_Parametros.Nivel3 > 0 Then
   cboNivel.AddItem "Nivel 3"
   cboNivel.ItemData(cboNivel.NewIndex) = 3
   vUltimoNivel = 3
End If
If gCntX_Parametros.Nivel4 > 0 Then
   cboNivel.AddItem "Nivel 4"
   cboNivel.ItemData(cboNivel.NewIndex) = 4
   vUltimoNivel = 4
End If
If gCntX_Parametros.Nivel5 > 0 Then
   cboNivel.AddItem "Nivel 5"
   cboNivel.ItemData(cboNivel.NewIndex) = 5
   vUltimoNivel = 5
End If
cboNivel.AddItem "Todos los Niveles (Detalle)"
cboNivel.ItemData(cboNivel.NewIndex) = 6

cboNivel.Text = "Todos los Niveles (Detalle)"

End Sub
