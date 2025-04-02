VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_CR_LiquidacionMasiva 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Liquidación Masiva"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   14520
   Begin XtremeSuiteControls.PushButton btnCausaRefresh 
      Height          =   375
      Left            =   13200
      TabIndex        =   27
      Top             =   960
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_CR_LiquidacionMasiva.frx":0000
   End
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   14295
      _Version        =   1441793
      _ExtentX        =   25215
      _ExtentY        =   1296
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   0
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmAF_CR_LiquidacionMasiva.frx":0700
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.CheckBox chkS06 
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         ToolTipText     =   "Etiqueta de Revisión Masiva Automática"
         Top             =   360
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "S06 Masivo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   1
         Left            =   6360
         TabIndex        =   23
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Picture         =   "frmAF_CR_LiquidacionMasiva.frx":0E00
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   2
         Left            =   9840
         TabIndex        =   24
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Actualizar Abonos"
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
         Picture         =   "frmAF_CR_LiquidacionMasiva.frx":0F6A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   3
         Left            =   11520
         TabIndex        =   25
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Liquidar"
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
         Picture         =   "frmAF_CR_LiquidacionMasiva.frx":166A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   135
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   14895
         _Version        =   1441793
         _ExtentX        =   26273
         _ExtentY        =   238
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   1
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   14295
      _Version        =   524288
      _ExtentX        =   25215
      _ExtentY        =   10610
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
      MaxCols         =   13
      SpreadDesigner  =   "frmAF_CR_LiquidacionMasiva.frx":1D91
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1200
      TabIndex        =   8
      Top             =   1440
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   5160
      TabIndex        =   11
      Top             =   960
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCausa 
      Height          =   330
      Left            =   8400
      TabIndex        =   12
      Top             =   960
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   5160
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   8400
      TabIndex        =   14
      Top             =   1440
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEjecutivo 
      Height          =   330
      Left            =   1200
      TabIndex        =   18
      Top             =   1920
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   330
      Left            =   8400
      TabIndex        =   20
      Top             =   1920
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   19
      Top             =   1920
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inst/Empr."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ejecutivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   10
      Top             =   1440
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Nombre"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   9
      Top             =   1440
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cédula"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   6
      Top             =   960
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Causas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fechas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidación Masiva"
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
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   765
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmAF_CR_LiquidacionMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

'spAFI_Renuncia_Liquidacion_Pendiente(@Inicio datetime, @Corte datetime, @Tipo char(1) = Null, @Institucion int = Null, @Causa int = Null
'                                              ,  @Cedula varchar(20) = Null, @Nombre varchar(100) = Null, @Ejecutivo varchar(100) = Null, @Usuario varchar(30) = Null )

Dim pTipo As String, pInstitucion As String, pCausa As String

If cboTipo.Text = "TODAS" Then
  pTipo = "Null"
Else
  pTipo = "'" & Mid(cboTipo.Text, 1, 1) & "'"
End If

If cboInstitucion.Text = "TODOS" Then
    pInstitucion = "Null"
Else
    pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

If cboCausa.Text = "TODOS" Then
    pCausa = "Null"
Else
    pCausa = cboCausa.ItemData(cboCausa.ListIndex)
End If

strSQL = "exec spAFI_Renuncia_Liquidacion_Pendiente '" & Format(dtpInicio.Value, "yyyy-mm-dd") & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
       & " 23:59', " & pTipo & ", " & pInstitucion & ", " & pCausa & ", '" & Trim(txtCedula.Text) & "', '" & Trim(txtNombre.Text) _
       & "', '" & Trim(txtEjecutivo.Text) & "', '" & Trim(txtUsuario.Text) & "'"
Call OpenRecordSet(rs, strSQL)

With vGrid
  .MaxRows = 0
  Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Value = chkTodas.Value
     .Col = 2
     .Text = CStr(rs!cod_Renuncia)
     .Col = 3
     .Value = chkS06.Value
     .Col = 4
     .Text = Trim(rs!Cedula)
     .Col = 5
     .Text = Trim(rs!Nombre)
     .Col = 6
     .Text = Trim(rs!Tipo_Desc)
     .Col = 7
     .Text = Trim(rs!Causa_Desc)
     .Col = 8
     .Text = Trim(rs!Estado_Desc)
     .Col = 9
     .Text = rs!Resuelto_Fecha_Mask
     .Col = 10
     .Text = Trim(rs!Resuelto_User & "")
     .Col = 11
     .Text = rs!Registro_Fecha_Mask
     .Col = 12
     .Text = Trim(rs!Registro_User & "")
     .Col = 13
     .Text = Trim(rs!Promotor_Desc)
   rs.MoveNext
  Loop
  rs.Close
End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 13
    vHeaders.Headers(1) = "Liquida?"
    vHeaders.Headers(2) = "Renuncia Id"
    vHeaders.Headers(3) = "S06?"
    vHeaders.Headers(4) = "Cédula"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Tipo"
    vHeaders.Headers(7) = "Causa"
    vHeaders.Headers(8) = "Estado"
    
    vHeaders.Headers(9) = "Res.Fecha"
    vHeaders.Headers(10) = "Res.Usuario"
    
    vHeaders.Headers(11) = "Reg.Fecha"
    vHeaders.Headers(12) = "Reg.Usuario"
    vHeaders.Headers(13) = "Ejecutivo"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_RenunciaPend_Liquidar")

End Sub

Private Sub sbAbonoActualiza()

End Sub

Private Sub sbLiquidar()
Dim i As Long, pS06 As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

ProgressBarX.Visible = True

With vGrid
    ProgressBarX.Max = .MaxRows
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = 1 Then
           .Col = 3
           pS06 = .Value
           .Col = 2
            strSQL = strSQL & Space(10) & "exec spAFI_Renuncia_Liquidacion_Procesa " & .Text & ", '" & glogon.Usuario & "', " & pS06
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
        ProgressBarX.Value = i
        DoEvents
        Me.MousePointer = vbHourglass
    Next i

    'Lote Final
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If

End With

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

MsgBox "Los casos marcados fueron procesados satisfactoriamente!", vbInformation

Call btnAccion_Click(0)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  ProgressBarX.Visible = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
      Call sbBuscar
    Case 1 'Exportar
      Call sbExportar
    Case 2 'Actualiza Abonos
      Call sbAbonoActualiza
    Case 3 'Liquidar
      Call sbLiquidar
    Case Else
    
End Select

End Sub

Private Sub btnCausaRefresh_Click()

If vPaso Then Exit Sub
If cboTipo.Text = "TODAS" Then
    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
           & " from causas_renuncias WHERE ACTIVO = 1"
Else
    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
           & " from causas_renuncias WHERE ACTIVO = 1" _
           & " and Tipo_Apl in('A', '" & IIf((Mid(cboTipo.Text, 1, 1) = "A"), "I", "P") & "')"

End If

strSQL = strSQL & " and Id_Causa in(" _
       & " select ID_CAUSA" _
       & "  From AFI_CR_RENUNCIAS" _
       & "  Where registro_Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
       & "    and Tipo in('A', '" & IIf((Mid(cboTipo.Text, 1, 1) = "A"), "I", "P") & "')" _
       & " and Estado = 'P' and LIQ is null" _
       & "  group by ID_CAUSA)"

Call sbCbo_Llena_New(cboCausa, strSQL, True, True)


End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub
If cboTipo.Text = "TODAS" Then
    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
           & " from causas_renuncias WHERE ACTIVO = 1"
Else
    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
           & " from causas_renuncias WHERE ACTIVO = 1" _
           & " and Tipo_Apl in('A', '" & IIf((Mid(cboTipo.Text, 1, 1) = "A"), "I", "P") & "')"

End If

Call sbCbo_Llena_New(cboCausa, strSQL, True, True)

End Sub

Private Sub chkS06_Click()
Dim i As Long

With vGrid
    For i = 1 To .MaxRows
        .Row = i
        .Col = 3
        .Value = chkS06.Value
    Next i
End With

End Sub

Private Sub chkTodas_Click()
Dim i As Long

With vGrid
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        .Value = chkTodas.Value
    Next i
End With

End Sub

Private Sub Form_Load()
vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
    cboTipo.AddItem "TODAS"
    cboTipo.AddItem "Asociación"
    cboTipo.AddItem "Patronales"
    cboTipo.Text = "TODAS"
vPaso = False

'Causas
Call cboTipo_Click

'Instituciones
strSQL = "select cod_Institucion as 'IdX', Descripcion as 'ItmX'" _
       & " from Instituciones WHERE ACTIVA = 1"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -15, dtpCorte.Value)

chkTodas.Value = xtpUnchecked
chkS06.Value = xtpChecked

vGrid.MaxRows = 0

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

gbAccion.Width = Me.Width - 550
ProgressBarX.Width = gbAccion.Width

vGrid.Width = gbAccion.Width

vGrid.Height = Me.Height - (vGrid.Top + 650)

End Sub
