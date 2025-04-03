VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCntX_AreasReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes para las Areas de Trabajo"
   ClientHeight    =   3768
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7548
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3768
   ScaleWidth      =   7548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCompara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Comparar con"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   144
      TabIndex        =   9
      Top             =   1800
      Width           =   7092
      Begin VB.TextBox txtMesCorte 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   13
         ToolTipText     =   "(F4) Mes del Periodo"
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtAnioCorte 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   12
         ToolTipText     =   "(F4) Año del Periodo"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtPeriodoCorte 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "(F4) Descripción del Periodo"
         Top             =   360
         Width           =   5112
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Periodo"
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
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtMes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      MaxLength       =   2
      TabIndex        =   8
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   1080
      Width           =   315
   End
   Begin VB.TextBox txtAnio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1296
      MaxLength       =   4
      TabIndex        =   7
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   1080
      Width           =   645
   End
   Begin VB.TextBox txtPeriodo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1956
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   1080
      Width           =   4635
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para consultar"
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtCodArea 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin XtremeSuiteControls.PushButton cmdGenera 
      Height          =   492
      Left            =   5760
      TabIndex        =   16
      Top             =   2880
      Width           =   1452
      _Version        =   1310720
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Generar"
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
      Appearance      =   14
      Picture         =   "frmCntX_AreasReportes.frx":0000
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      Height          =   372
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   5052
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   612
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
End
Attribute VB_Name = "frmCntX_AreasReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCompara_Click()
If chkCompara.Value = vbChecked Then
  fra.Enabled = True
Else
  fra.Enabled = False
End If
End Sub

Private Sub cmdGenera_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

'Hay que crear para todos los casos el Balance de Comprobacion del Area
'Pasos 1 - Crear Balance Comprobacion del Periodo Indicado

lbl.Caption = "Procesando Periodo Base..."
lbl.Refresh
Call sbCntX_Areas_Balance_Comprobacion(txtAnio, txtMes, txtCodArea)
'Pasos 2 - Crear Suplemento del Balance de Comprobacion con los Comprativos
If chkCompara.Value = vbChecked Then
    lbl.Caption = "Procesando Periodo Comparativo..."
    lbl.Refresh
    Call sbCntX_Areas_Balance_Compara(txtAnioCorte, txtMesCorte, txtCodArea)
End If
'Pasos 3 - Mayorizar el Balance de Comprobacion
lbl.Caption = "Mayorizando Resultados..."
lbl.Refresh
Call sbCntX_Areas_Mayorizar(txtCodArea)

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
 .Formulas(3) = "Mascara='" & gCntX_Parametros.MascaraCod & "'"
 .Formulas(4) = "Area='" & txtNombre & "'"
 .Formulas(5) = "Titulo='" & UCase(Mid(cbo, 6, 50)) & "'"
 
 If chkCompara.Value = vbChecked Then
   .Formulas(6) = "SubTitulo='" & txtPeriodo & " vrs " & txtPeriodoCorte & "'"
 Else
   .Formulas(6) = "SubTitulo='" & txtPeriodo & "'"
 End If
 
 .SelectionFormula = "{CntX_Acceso_Historico.USUARIO} = '" & glogon.Usuario _
             & "' AND {CntX_Acceso_Historico.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
             & " AND {CntX_Acceso_Historico.COD_AREA} = " & txtCodArea
 .Connect = glogon.ConectRPT
  

    Select Case Val(Mid(cbo, 1, 2))
      Case 1 'Balance de Comprobacion
           .ReportFileName = App.Path & "\" & "AreaBalanceComprobacion.rpt"
      
      Case 2 'Balance de Resultados
        'Llena Variables Globales
        Call sbCntX_Areas_Utilidad(txtCodArea)
        
        If chkCompara.Value = vbChecked Then
           .Formulas(7) = "UT_MES = " & vAreaUTMes
           .Formulas(8) = "UT_ACUMULADA = " & vAreaUTAcumulada
           .Formulas(9) = "UTC_MES = " & vAreaUTCMes
           .Formulas(10) = "UTC_ACUMULADA =" & vAreaUTCAcumulada
           .ReportFileName = App.Path & "\" & "AreaBalanceResultadosCMP.rpt"
        Else
           .Formulas(7) = "UT_MES = " & vAreaUTMes
           .Formulas(8) = "UT_ACUMULADA =" & vAreaUTAcumulada
           .ReportFileName = App.Path & "\" & "AreaBalanceResultados.rpt"
        End If
      
      Case 3 'Balance General
        'Llena Variables Globales
        Call sbCntX_Areas_Utilidad(txtCodArea)
        
        If chkCompara.Value = vbChecked Then
           .Formulas(7) = "UT_MES = " & vAreaUTMes
           .Formulas(8) = "UT_ACUMULADA = " & vAreaUTAcumulada
           .Formulas(9) = "UTC_MES = " & vAreaUTCMes
           .Formulas(10) = "UTC_ACUMULADA =" & vAreaUTCAcumulada
           .ReportFileName = App.Path & "\" & "AreaBalanceGeneralCMP.rpt"
        Else
           .Formulas(7) = "UT_MES = " & vAreaUTMes
           .Formulas(8) = "UT_ACUMULADA =" & vAreaUTAcumulada
           .ReportFileName = App.Path & "\" & "AreaBalanceGeneral.rpt"
        End If
      
      Case 4 'Balance de Situacion
           .ReportFileName = App.Path & "\" & "AreaBalanceSituacion.rpt"
      Case 5 'Costos/Gastos
        .SelectionFormula = .SelectionFormula & " AND {CntX_Tipos_Cuentas.CLASIFICACION} = 'G'"
        If chkCompara.Value = vbChecked Then
           .ReportFileName = App.Path & "\" & "AreaGastosIngresosCMP.rpt"
        Else
           .ReportFileName = App.Path & "\" & "AreaGastosIngresos.rpt"
        End If
      
      Case 6 'Ventas/Ingresos
        .SelectionFormula = .SelectionFormula & " AND {CntX_Tipos_Cuentas.CLASIFICACION} = 'I'"
        If chkCompara.Value = vbChecked Then
           .ReportFileName = App.Path & "\" & "AreaGastosIngresosCMP.rpt"
        Else
           .ReportFileName = App.Path & "\" & "AreaGastosIngresos.rpt"
        End If
    
    End Select

   .PrintReport

End With

vError:

lbl.Caption = ""
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
vModulo = 20

txtAnio = gCntX_Parametros.PeriodoAnio
txtAnioCorte = gCntX_Parametros.PeriodoAnio

txtMes = gCntX_Parametros.PeriodoMes
txtMesCorte = gCntX_Parametros.PeriodoMes

cbo.AddItem "01 - Balance de Comprobación"
cbo.AddItem "02 - Balance de Resultados"
cbo.AddItem "03 - Balance General"
cbo.AddItem "04 - Balance de Situación"
cbo.AddItem "05 - Costos/Gastos"
cbo.AddItem "06 - Ventas/Ingresos"

cbo.Text = "01 - Balance de Comprobación"

End Sub


Private Sub txtCodArea_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "cod_area"
    gBusquedas.Orden = "cod_area"
    gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
    gBusquedas.Consulta = "select cod_area,descripcion from CntX_Area_Definicion"
    frmBusquedas.Show vbModal
    
    txtCodArea.SetFocus
    txtCodArea = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
    txtNombre = IIf((gBusquedas.Resultado2 = ""), "", gBusquedas.Resultado2)
End If


If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus


End Sub

Private Sub txtCodArea_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select * from CntX_Area_Definicion where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_area = " & txtCodArea
Call OpenRecordSet(rs, strSQL, 0)
 txtNombre = rs!Descripcion & ""
rs.Close
Exit Sub

vError:

End Sub

Private Sub txtMes_Change()
On Error GoTo vError
  txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
  txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtMesCorte_Change()
On Error GoTo vError
  txtPeriodoCorte = fxCntX_PeriodoDesc(txtAnioCorte, txtMesCorte)
vError:
End Sub

Private Sub txtAnioCorte_Change()
On Error GoTo vError
  txtPeriodoCorte = fxCntX_PeriodoDesc(txtAnioCorte, txtMesCorte)
vError:
End Sub


