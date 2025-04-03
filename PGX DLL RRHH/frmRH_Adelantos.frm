VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Adelantos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Adelantos de Salarios"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16680
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   16680
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lswBancos 
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   5318
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      BackColor       =   16777215
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar prgBarX 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   16680
      _ExtentX        =   29422
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton btnBoletaBancos 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Adelantos.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnAdelantos 
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Consultar"
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
      Picture         =   "frmRH_Adelantos.frx":0707
   End
   Begin XtremeSuiteControls.PushButton btnAdelantos 
      Height          =   735
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmRH_Adelantos.frx":1125
   End
   Begin XtremeSuiteControls.PushButton btnAdelantos 
      Height          =   735
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   8400
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Pago"
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
      Picture         =   "frmRH_Adelantos.frx":192A
   End
   Begin XtremeSuiteControls.PushButton btnAdelantos 
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   9240
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Notificación Email"
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
      Picture         =   "frmRH_Adelantos.frx":1D8C
   End
   Begin XtremeSuiteControls.PushButton btnAdelantos 
      Height          =   735
      Index           =   5
      Left            =   2160
      TabIndex        =   8
      Top             =   9240
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Boletas Impresas"
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
      Picture         =   "frmRH_Adelantos.frx":25A9
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAdelantos 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Actualizar Base"
      Top             =   3960
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   1291
      _StockProps     =   79
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
      Picture         =   "frmRH_Adelantos.frx":2D65
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   8400
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Autorización"
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
      Picture         =   "frmRH_Adelantos.frx":3728
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   8295
      Left            =   4560
      TabIndex        =   14
      Top             =   1200
      Width           =   12135
      _Version        =   524288
      _ExtentX        =   21405
      _ExtentY        =   14631
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      SpreadDesigner  =   "frmRH_Adelantos.frx":3F06
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtPagoNo 
      Height          =   555
      Left            =   3600
      TabIndex        =   15
      Top             =   2880
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNominaNo 
      Height          =   555
      Left            =   960
      TabIndex        =   17
      Top             =   2880
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAdelantoId 
      Height          =   555
      Left            =   960
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnNuevo 
      Height          =   555
      Index           =   6
      Left            =   3120
      TabIndex        =   21
      ToolTipText     =   "Nuevo Adelanto"
      Top             =   2280
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      Picture         =   "frmRH_Adelantos.frx":4868
   End
   Begin XtremeSuiteControls.FlatEdit txtInicio 
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   3480
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCorte 
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   3480
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   855
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEstado 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Adelanto No.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina No.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago No.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Nómina"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Salidas por Banco.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   12
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adelantos de Salario"
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
      Height          =   600
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   16935
   End
End
Attribute VB_Name = "frmRH_Adelantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnAdelantos_Click(Index As Integer)
Dim i As Integer, pDetalle As String

On Error GoTo vError

pDetalle = "Adelanto de Salarios> " & txtAdelantoId.Text & " > Nómina: " & cboNomina.Text & "   No.: " & txtNominaNo.Text

Select Case Index
    Case 0 'Actualizar
        Call Bitacora("Actualiza", "Cálculos de " & pDetalle)
        Call sbActualiza
        
    Case 1 'Buscar
        Call sbBuscar
    
    Case 2 'Exportar
        Call sbExportar
        
    Case 3 'Pago
    
    If lblEstado.Tag = "P" Then
        MsgBox "Los Adelantos de Salarios ya fueron cancelados!", vbExclamation
        Exit Sub
    End If
    
    i = MsgBox("Esta seguro que desea PAGAR los Adelantos de Salarios Id: " & txtAdelantoId.Text & " ?", vbYesNo)
    If i = vbYes Then
    
        Me.MousePointer = vbHourglass
        
        strSQL = "exec spRH_Adelanto_Pago " & txtAdelantoId.Text & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Pago de " & pDetalle)
        
        Me.MousePointer = vbDefault
        MsgBox "Los Adelantos de Salarios han sido Reportados a Bancos para su pago!", vbInformation
        
        Call cboNomina_Click
    End If
    
    
    Case 4 'Notificacion Email
    
          Call Bitacora("Aplica", "Notificación Email de " & pDetalle)
    
          Call sbRH_Boleta_Adelanto_Email(txtAdelantoId.Text, "")
        
          Me.MousePointer = vbDefault
          MsgBox "Boletas de Aguinaldo, fueron enviadas por Email a los Empleados!", vbInformation
  
    
    Case 5 'Boleta Impresora
  
        Call Bitacora("Aplica", "Boletas de " & pDetalle)
        Call sbBoleta("")
    
    
End Select

Exit Sub

vError:
 Me.MousePointer = vbHourglass
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub btnAutorizacion_Click()
Dim i As Integer, pDetalle As String

On Error GoTo vError

strSQL = "exec spRH_Adelanto_Autoriza_Valida " & txtAdelantoId.Text & ",'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Pass = 0 Then
   MsgBox rs!Mensaje, vbExclamation
   Exit Sub
End If
rs.Close

    i = MsgBox("Esta seguro que desea Autorizar el Adelanto de Salarios No." & txtAdelantoId.Text & "  Nómina: " & cboNomina.Text & ", No.: " & txtNominaNo.Text & " ?", vbYesNo)
    If i = vbYes Then
        
        Me.MousePointer = vbHourglass
               
        
        pDetalle = "Adelanto Salario> " & txtAdelantoId.Text & " > Nómina: " & cboNomina.Text & "   No.: " & txtNominaNo.Text
               
               
        strSQL = "exec spRH_Adelanto_Autoriza " & txtAdelantoId.Text & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Autorización de " & pDetalle)
        
        Me.MousePointer = vbDefault
        MsgBox "El Adelanto ha sido Autorizado!", vbInformation
        
        Call cboNomina_Click
    
    End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBoletaBancos_Click()

Dim pTitulo As String

On Error GoTo vError

pTitulo = "Adelanto Id: " & txtAdelantoId.Text & "   Nómina: " & cboNomina.Text & "   Estado: " & lblEstado.Caption

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH, Adelanto de Salarios: Boleta de Control"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxSubTitulo = '" & pTitulo & "'"
    .Formulas(3) = "fxUsuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(4) = "fxFecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Adelanto_Boleta_Control.rpt")
    strSQL = "{vRH_Adelanto_Estado_Rsm.ADELANTO_ID} = " & txtAdelantoId.Text
    .SelectionFormula = strSQL
    .SubreportToChange = "sbBancosResumen"
    
    .StoredProcParam(0) = txtAdelantoId.Text

    .Action = 1
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAdelanto_Consulta()

On Error GoTo vError

strSQL = "exec spRH_Adelanto_Nomina_Id '" & cboNomina.ItemData(cboNomina.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtAdelantoId.Text = rs!ADELANTO_ID
    
    txtNominaNo.Text = rs!NOMINA_NUM
    txtPagoNo.Text = rs!NPAGO_MES
    
    Select Case rs!Estado
    Case "A"
        lblEstado.Caption = "Abierto"
    Case "D"
        lblEstado.Caption = "Descartado"
    Case "X"
        lblEstado.Caption = "Autorizada"
    Case "P"
        lblEstado.Caption = "Pagado"
    End Select
    
    lblEstado.Tag = rs!Estado
    
    txtInicio.Text = Format(rs!Fecha_Inicio, "yyyy-mm-dd")
    txtCorte.Text = Format(rs!Fecha_Corte, "yyyy-mm-dd")
Else
    txtAdelantoId.Text = 0
    txtPagoNo.Text = 0
    lblEstado.Caption = ""
    lblEstado.Tag = ""
    txtInicio.Text = ""
    txtCorte.Text = ""
    
    vGrid.MaxRows = 0
    lswBancos.ListItems.Clear
End If

rs.Close

If txtAdelantoId.Text <> "0" Then
    Call sbBuscar
End If

Exit Sub

vError:


End Sub


Private Sub btnNuevo_Click(Index As Integer)
If vPaso Or cboNomina.ListCount = 0 Then Exit Sub

On Error GoTo vError

strSQL = "exec spRH_Adelanto_Nuevo '" & cboNomina.ItemData(cboNomina.ListIndex) & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call cboNomina_Click

Exit Sub

vError:

End Sub

Private Sub cboNomina_Click()
If vPaso Or cboNomina.ListCount = 0 Then Exit Sub

strSQL = "exec spRH_Nomina_Actual '" & cboNomina.ItemData(cboNomina.ListIndex) & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

txtNominaNo.Text = rs!NOMINA_ID

rs.Close

Call sbAdelanto_Consulta

End Sub

Private Sub Form_Activate()
vModulo = 23
End Sub

Private Sub Form_Load()

vModulo = 23

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

'Nomina
vPaso = True
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX" _
           & " from RH_NOMINAS_CATALOGO Where I_Adelanto_Salario = 1"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)
vPaso = False



vGrid.MaxRows = 0

With lswBancos.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Casos", 800, vbRightJustify
    .Add , , "Monto", 1800, vbRightJustify
End With
lswBancos.BackColor = RGB(214, 234, 248)

lblEstado.Caption = ""

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbBoleta(Optional pEmpleado As String = "")

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del RRHH: Boleta de Adelanto de Salario"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("RH_Boleta_Adelanto.rpt")
    strSQL = "{vRH_Adelanto_Boleta.ADELANTO_ID} = " & txtAdelantoId.Text
                
    If pEmpleado <> "" Then
        strSQL = strSQL & " AND {vRH_Adelanto_Boleta.EMPLEADO_ID} = '" & pEmpleado & "'"
    End If
        
        
     .SelectionFormula = strSQL
    .PrintReport
End With

End Sub


Private Sub sbActualiza()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Adelanto_Control_Calcula " & txtAdelantoId.Text & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Call sbBuscar

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Adelanto_Resumen_Consulta " & txtAdelantoId.Text
Call sbCargaGrid_Local(vGrid, vGrid.MaxCols, strSQL)


btnAdelantos.Item(0).Enabled = False 'Actualiza
btnAdelantos.Item(1).Enabled = False 'Consulta
btnAdelantos.Item(2).Enabled = False 'Exporta
btnAdelantos.Item(3).Enabled = False 'Pago
btnAdelantos.Item(4).Enabled = False 'Email
btnAdelantos.Item(5).Enabled = False 'Boletas

btnAutorizacion.Enabled = False

Select Case lblEstado.Tag
    Case "A" 'Abierta
            
        If btnAdelantos.Item(0).Tag = "1" Then
            btnAdelantos.Item(0).Enabled = True 'Actualiza
        End If
        
        btnAdelantos.Item(1).Enabled = True 'Consulta
        btnAdelantos.Item(2).Enabled = True 'Exporta
    
        If btnAutorizacion.Tag = "1" Then
            btnAutorizacion.Enabled = True
        End If
    
    Case "X" 'Autorizada
        btnAdelantos.Item(1).Enabled = True 'Consulta
        btnAdelantos.Item(2).Enabled = True 'Exporta
        
        If btnAdelantos.Item(3).Tag = "1" Then
            btnAdelantos.Item(3).Enabled = True 'Paga
            btnAdelantos.Item(4).Enabled = True 'Boleta Email
            btnAdelantos.Item(5).Enabled = True 'Boleta Imprime
        End If
        
   
    Case "P" 'Pagada
        btnAdelantos.Item(1).Enabled = True 'Consulta
        btnAdelantos.Item(2).Enabled = True 'Exporta
        
        If btnAdelantos.Item(3).Tag = "1" Then
            btnAdelantos.Item(4).Enabled = True 'Boleta Email
            btnAdelantos.Item(5).Enabled = True 'Boleta Imprime
        End If


End Select



'Resumen de Salidas por Banco
Dim pCasos As Long, pMonto As Currency

pCasos = 0
pMonto = 0

strSQL = "exec spRH_Adelanto_Pago_Banco_Rsm " & txtAdelantoId.Text
Call OpenRecordSet(rs, strSQL)

With lswBancos.ListItems
    .Clear
  Do While Not rs.EOF
    Set itmX = .Add(, , rs!Descripcion)
        itmX.SubItems(1) = Format(rs!Casos, "###,###0")
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
    
        pCasos = pCasos + rs!Casos
        pMonto = pMonto + rs!Monto
    rs.MoveNext
  Loop
  rs.Close

Set itmX = .Add(, , "")
    itmX.SubItems(1) = "________________"
    itmX.SubItems(2) = "________________"
Set itmX = .Add(, , "TOTAL:")
    itmX.SubItems(1) = Format(pCasos, "###,###0")
    itmX.SubItems(2) = Format(pMonto, "Standard")

End With


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbCargaGrid_Local(pGrid As Object, pGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim i As Integer

On Error GoTo vErrorLoad

Call OpenRecordSet(rs, strSQL, 0)
  
pGrid.MaxRows = 0
Do While Not rs.EOF
  pGrid.MaxRows = pGrid.MaxRows + 1
  pGrid.Row = pGrid.MaxRows
  For i = 3 To pGrid.MaxCols
    
    pGrid.Col = i
    Select Case i
        Case 3
            pGrid.Text = rs!Empleado_ID
        Case 4
            pGrid.Text = rs!IDENTIFICACION
        Case 5
            pGrid.Text = rs!NOMBRE_COMPLETO
        Case 6
            pGrid.Text = Format(rs!Salario_Ordinario, "Standard")
        Case 7
            pGrid.Text = Format(rs!Porcentaje, "Standard")
        Case 8
            pGrid.Text = Format(rs!Monto_Adelanto, "Standard")
        Case 9
            pGrid.Text = rs!CUENTA_BANCARIA & ""
        Case 10
            pGrid.Text = rs!Tesoreria_NSolicitud & ""
        Case 11
            pGrid.Text = rs!Tesoreria_Fecha & ""
        Case 12
            pGrid.Text = rs!Verificacion & ""
    End Select

  Next i
  rs.MoveNext
Loop
rs.Close

Exit Sub

vErrorLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub



Private Sub sbExportar()
 Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols

    vHeaders.Headers(1) = "@"
    vHeaders.Headers(2) = "[Pr]"
    
    vHeaders.Headers(3) = "Empleado Id"
    vHeaders.Headers(4) = "Identificación"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Salario Ordinario"
    vHeaders.Headers(7) = "% Adelanto"
    vHeaders.Headers(8) = "Monto Adelanto"
    vHeaders.Headers(9) = "IBAN"
    vHeaders.Headers(10) = "Tesorería Id"
    vHeaders.Headers(11) = "Tesorería Fecha"
    vHeaders.Headers(12) = "Validación"

    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_RRHH_Adelanto_" & txtAdelantoId.Text & "_" & cboNomina.Text)
End Sub



Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
vGrid.Width = Me.Width - (vGrid.Left + 350)
vGrid.Height = Me.Height - (vGrid.Top + 700)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call cboNomina_Click


End Sub

Private Sub sbAjustes_Porcentaje(pAdelantoId As Long, pEmpleadoId As String, pValor As Currency, Optional pTipo As String = "P")

On Error GoTo vError

Dim i As Integer

Me.MousePointer = vbHourglass

strSQL = "EXEC spRH_Adelanto_Ajuste " & pAdelantoId & ", '" & pEmpleadoId & "', " & CCur(pValor) & ", '" & pTipo & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pEmpleadoId As String

On Error GoTo vError

vGrid.Row = Row
vGrid.Col = 3
pEmpleadoId = vGrid.Text

Select Case Col
    Case 1 'Email
        Call sbRH_Boleta_Adelanto_Email(txtAdelantoId.Text, pEmpleadoId)
        MsgBox "Correo Electrónico Enviado al Empleado: " & pEmpleadoId, vbInformation
    Case 2 'Boleta
        Call sbBoleta(pEmpleadoId)
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim pEmpleadoId As String, pValor As Currency
Dim pSalario As Currency

If lblEstado.Tag <> "A" Then Exit Sub

If Col = 7 Then
  vGrid.Row = Row
  vGrid.Col = 3
  pEmpleadoId = vGrid.Text
  vGrid.Col = 6
  pSalario = vGrid.Text
  vGrid.Col = Col
  pValor = vGrid.Text
  
  vGrid.Col = 8
  vGrid.Text = Format(pSalario * pValor / 100, "Standard")
  Call sbAjustes_Porcentaje(txtAdelantoId.Text, pEmpleadoId, pValor, "P")
  

End If

If Col = 8 Then
  vGrid.Row = Row
  vGrid.Col = 3
  pEmpleadoId = vGrid.Text
  vGrid.Col = 6
  pSalario = vGrid.Text
  vGrid.Col = Col
  pValor = vGrid.Text
  
  vGrid.Col = 7
  vGrid.Text = Format((pValor / pSalario) * 100, "Standard")
  
  Call sbAjustes_Porcentaje(txtAdelantoId.Text, pEmpleadoId, pValor, "M")
End If

End Sub

