VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmFNDConMovCobro 
   Caption         =   "Movimientos de Operaciones de Cobro no Reportadas al Fondo"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16065
   Icon            =   "frmFNDConMovCobro.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   16065
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkOperacionesNoValidas 
      Height          =   360
      Left            =   8640
      TabIndex        =   29
      Top             =   900
      Width           =   3372
      _Version        =   1441793
      _ExtentX        =   5948
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Revisar Operaciones no Asociadas a Planes de Ahorros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkMarca 
      Height          =   216
      Left            =   600
      TabIndex        =   26
      Top             =   960
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   381
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Información de Conciliación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8520
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtPlDifCasos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtPlFndCasos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPlCrdCasos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtPlDifMnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtPlFndMnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtPlCrdMnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image imgInfo 
         Height          =   480
         Left            =   5400
         Picture         =   "frmFNDConMovCobro.frx":030A
         ToolTipText     =   "Cerrar Información"
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Diferencias"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   5760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Planilla Registrada a Contratos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Entrada Planilla (Retenciones)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   168
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   16068
      _ExtentX        =   28337
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   9120
      TabIndex        =   11
      Top             =   240
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDConMovCobro.frx":6B5C
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   4800
      TabIndex        =   14
      Top             =   480
      Width           =   3372
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton cmdArchivo 
      Height          =   492
      Left            =   10560
      TabIndex        =   15
      Top             =   240
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDConMovCobro.frx":757A
   End
   Begin XtremeSuiteControls.PushButton cmdInfo 
      Height          =   492
      Left            =   8640
      TabIndex        =   16
      ToolTipText     =   "Filtros Adicionales"
      Top             =   240
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   372
      Left            =   12480
      TabIndex        =   17
      Top             =   360
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "1000"
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   312
      Left            =   4800
      TabIndex        =   23
      Top             =   120
      Width           =   3372
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   582
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3612
      Left            =   0
      TabIndex        =   24
      Top             =   1440
      Width           =   14532
      _Version        =   524288
      _ExtentX        =   25633
      _ExtentY        =   6371
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDConMovCobro.frx":7C7F
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboAccion 
      Height          =   312
      Left            =   2280
      TabIndex        =   27
      Top             =   960
      Width           =   4572
      _Version        =   1441793
      _ExtentX        =   8070
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   312
      Left            =   6960
      TabIndex        =   30
      Top             =   960
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Accion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   1440
      TabIndex        =   28
      Top             =   960
      Width           =   1212
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   492
      Left            =   0
      TabIndex        =   25
      Top             =   840
      Width           =   14412
      _Version        =   1441793
      _ExtentX        =   25421
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   5
      Left            =   1560
      TabIndex        =   21
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   4
      Left            =   1560
      TabIndex        =   20
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   3960
      TabIndex        =   19
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   12480
      TabIndex        =   18
      Top             =   120
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmFNDConMovCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnAccion_Click()
Dim strSQL As String, i As Integer

Dim pOperacion As Long, pTipoDoc As String, pNumDoc As String, pMonto As Currency, pAccion As Integer

If cboAccion.ItemData(cboAccion.ListIndex) = "0" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

pAccion = cboAccion.ItemData(cboAccion.ListIndex)

With vGrid

    strSQL = ""
    For i = 1 To .MaxRows
        .Row = i
        .col = 2
        pOperacion = .Text
            
        .col = 6
        pMonto = CCur(.Text)
        
        .col = 11
        pTipoDoc = .Text
        
        .col = 12
        pNumDoc = .Text
            
        .col = 1
        If .Value = vbChecked Then
            strSQL = strSQL & Space(10) & "exec spFnd_AcreditaMovCbrPendiente " & pOperacion _
                   & ",'" & glogon.Usuario & "'," & pAccion & ",'" & pTipoDoc & "','" & pNumDoc & "'," & pMonto
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
        scMain.Caption = "Procesando caso " & i & " de " & .MaxRows
        DoEvents
    Next i

    scMain.Caption = "Finalizando..."
    DoEvents
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If

    scMain.Caption = ""


End With

Me.MousePointer = vbDefault
MsgBox "Casos Procesados Satisfactoriamente!", vbInformation

Call cmdBuscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboPlan_Click()

txtPlCrdMnt.Text = ""
txtPlCrdCasos.Text = ""

txtPlFndMnt.Text = ""
txtPlFndCasos.Text = ""

txtPlDifMnt.Text = ""
txtPlDifCasos.Text = ""

End Sub

Private Sub chkMarca_Click()
Dim i As Long

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 1
    vGrid.Value = chkMarca.Value
Next i

End Sub

Private Sub cmdArchivo_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = ""
    vHeaders.Headers(2) = "Operacion"
    vHeaders.Headers(3) = "Código"
    vHeaders.Headers(4) = "Cédula"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Monto"
    vHeaders.Headers(7) = "Cuenta"
    vHeaders.Headers(8) = "Detalle"
    vHeaders.Headers(9) = "Fecha"
    vHeaders.Headers(10) = "Deductora"
    vHeaders.Headers(11) = "Tipo Doc."
    vHeaders.Headers(12) = "Num. Doc."

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_FndMovCobradosNoAcreditados")

End Sub

Private Function fxEstadoContrato(pOperacion As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String

strSQL = "select cod_contrato,cod_plan,estado from fnd_contratos where operacion = " & pOperacion
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    vResultado = "No existe contrato..."
Else
    vResultado = "Pln.:" & Trim(rs!cod_Plan) & "..Cnt:" & rs!COD_CONTRATO & "..Est.:" & IIf(rs!Estado = "A", "Activo", "Cancelado")
End If
rs.Close


fxEstadoContrato = vResultado

End Function


Private Sub sbBuscarRetNoReg()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset
Dim curMonto As Currency

On Error GoTo vError


Me.MousePointer = vbHourglass


'Casos Normales, de contratos Asociados a Operaciones de Retencion
scMain.Caption = "Cargando Información..."
DoEvents


strSQL = "Select P.codigo,P.id_solicitud,P.Principal,P.fecha,P.Proceso,I.descripcion as InstitucionX" _
       & ",S.Nombre,C.CTANAMORT,C.CTAOAMORT,Cnt.cedula,Cnt.cod_plan,Cnt.cod_operadora,Cnt.cod_contrato" _
       & ",Cnt.cod_operadora,Cnt.cod_plan,Cnt.Estado, P.tcon,P.ncon" _
       & " from vCRDsReportesMov P" _
       & " inner join fnd_contratos Cnt on P.id_solicitud = Cnt.operacion" _
       & " inner join socios S on Cnt.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " inner join catalogo C on P.codigo = C.codigo" _
       & " where P.tcon in('1','PRM','PLA') and P.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & " and dbo.fxFnd_MovimientoExiste(Cnt.cod_operadora,Cnt.cod_Plan,Cnt.cod_Contrato,P.Tcon, P.Ncon) = 0"

If cboPlan.Text <> "TODOS" Then
   strSQL = strSQL & " and Cnt.cod_plan = '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
End If

       
Call OpenRecordSet(rs, strSQL)
curMonto = 0

vGrid.MaxRows = 0

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1

Do While Not rs.EOF
  
  scMain.Caption = "Procesando " & PrgBar.Value & " de " & PrgBar.Max
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 1
  vGrid.Value = chkMarca.Value
  
  vGrid.col = 2
  vGrid.Text = CStr(rs!Id_Solicitud)
  vGrid.col = 3
  vGrid.Text = CStr(rs!Codigo)
  vGrid.col = 4
  vGrid.Text = CStr(rs!Cedula)
  vGrid.col = 5
  vGrid.Text = CStr(rs!Nombre)
  vGrid.col = 6
  vGrid.Text = Format(rs!Principal, "Standard")
  vGrid.col = 7
  vGrid.Text = fxgCntCuentaFormato(True, rs!CtaNamort, 0)
           
  vGrid.col = 8
  vGrid.Text = IIf(rs!Estado = "A", "Activo", "Cancelado")
  vGrid.col = 9
  vGrid.Text = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  vGrid.col = 10
  vGrid.Text = rs!InstitucionX
  vGrid.col = 11
  vGrid.Text = CStr(rs!tcon)
  vGrid.col = 12
  vGrid.Text = CStr(rs!nCon)

       
  curMonto = curMonto + rs!Principal
  
  PrgBar.Value = PrgBar.Value + 1
  
  rs.MoveNext
Loop
rs.Close




'Procesando Retenciones que no tienen referencias a Contratos de Fondos
If chkOperacionesNoValidas.Value = vbChecked Then
    scMain.Caption = "Cargando Información..."

    strSQL = "Select P.codigo,P.id_solicitud,P.Principal,R.opex,P.fecha,P.Proceso,I.descripcion as InstitucionX" _
           & ",S.Nombre,C.CTANAMORT,C.CTAOAMORT,R.cedula,F.cod_plan,F.cod_operadora,P.tcon,P.ncon" _
           & " from vCRDsReportesMov P left Join reg_creditos R on P.id_solicitud = R.id_solicitud" _
           & " inner join socios S on R.cedula = S.cedula" _
           & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
           & " inner join catalogo C on R.codigo = C.codigo" _
           & " inner join fnd_planes F on P.codigo = F.codigo_ase" _
           & "  left join fnd_contratos Cnt on F.cod_Operadora = Cnt.Cod_Operadora and F.cod_plan = Cnt.Cod_Plan" _
           & "   and P.id_solicitud = Cnt.Operacion" _
           & " where P.tcon in('1','PRM','PLA') and P.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and Cnt.cod_contrato is null"

    If cboPlan.Text <> "TODOS" Then
       strSQL = strSQL & " and F.cod_plan = '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
    End If

    'strSQL = strSQL & " and P.id_solicitud not in(select operacion from fnd_contratos"
    'If cboPlan.Text <> "[TODOS]" Then
    '   strSQL = strSQL & " where cod_plan = '" & cboPlan.ItemData(cboPlan.ListIndex)  & "')"
    'Else
    '   strSQL = strSQL & ")"
    'End If

    Call OpenRecordSet(rs, strSQL)
    PrgBar.Max = rs.RecordCount + 1
    PrgBar.Value = 1

    Do While Not rs.EOF

        scMain.Caption = "Procesando " & PrgBar.Value & " de " & PrgBar.Max
        
        vGrid.MaxRows = vGrid.MaxRows + 1
        vGrid.Row = vGrid.MaxRows
        
        vGrid.col = 1
        vGrid.Value = chkMarca.Value
        
        vGrid.col = 2
        vGrid.Text = CStr(rs!Id_Solicitud)
        vGrid.col = 3
        vGrid.Text = CStr(rs!Codigo)
        vGrid.col = 4
        vGrid.Text = CStr(rs!Cedula)
        vGrid.col = 5
        vGrid.Text = CStr(rs!Nombre)
        vGrid.col = 6
        vGrid.Text = Format(rs!Principal, "Standard")
        vGrid.col = 7
        vGrid.Text = fxgCntCuentaFormato(True, rs!CtaNamort, 0)
                 
        vGrid.col = 8
        vGrid.Text = "No Existe Contrato Asociado"
        vGrid.col = 9
        vGrid.Text = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
        vGrid.col = 10
        vGrid.Text = rs!InstitucionX
        vGrid.col = 11
        vGrid.Text = CStr(rs!tcon)
        vGrid.col = 12
        vGrid.Text = CStr(rs!nCon)

        curMonto = curMonto + rs!Principal

        PrgBar.Value = PrgBar.Value + 1

      rs.MoveNext
    Loop
    rs.Close
End If 'chk


scMain.Caption = "Total: " & Format(curMonto, "Standard")
PrgBar.Value = 1

Me.MousePointer = vbDefault

MsgBox "Consulta Finalizada satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbBuscarRegNoRet()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim itmX As ListItem, rsTmp As New ADODB.Recordset
'Dim vMascara As String, curMonto As Currency
'
'On Error GoTo vError
'
''Cobros en Retenciones no registrados
'With lsw.ColumnHeaders
' .Clear
'
'
' .Add , , "Operadora", 1440
' .Add , , "Plan", 1440
' .Add , , "Contrato", 1440
' .Add , , "Identificación", 1440
' .Add , , "Nombre", 2440
' .Add , , "Monto", 1440, vbRightJustify
' .Add , , "Fecha", 1440
' .Add , , "Tipo", 1440
' .Add , , "Comprobante", 1440
' .Add , , "Institución", 3440
'End With
'
'
'Me.MousePointer = vbHourglass
'
'vMascara = "#############"
'
'
''Revisa Aportes vía planillas que no fueron procesados por medio de Retenciones
'scMain.Caption = "Cargando Información..."
'scMain.Refresh
'
'strSQL = "select Cnt.COD_OPERADORA, Cnt.COD_PLAN,Cnt.COD_CONTRATO, Cnt.CEDULA, Soc.NOMBRE, Det.MONTO, Det.FECHA,Det.TCON,Det.NCON" _
'       & ",Inst.DESCRIPCION as 'Institucion'" _
'       & " from FND_CONTRATOS Cnt inner join FND_CONTRATOS_DETALLE Det on Cnt.COD_OPERADORA = Det.COD_OPERADORA" _
'       & "   and Cnt.COD_PLAN = Det.COD_PLAN and Cnt.COD_CONTRATO = Det.COD_CONTRATO" _
'       & "   inner join SOCIOS Soc on Cnt.CEDULA = Soc.CEDULA" _
'       & "   inner join INSTITUCIONES Inst on Soc.COD_INSTITUCION = Inst.COD_INSTITUCION" _
'       & "   left join CREDITOS_DT Crd on Cnt.OPERACION = Crd.ID_SOLICITUD and Det.TCON = Crd.TCON and Det.NCON = Crd.NCON" _
'       & "  where Det.TCON in('1','PRM')  and Crd.TCON is null" _
'       & "   and Det.FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
'       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
'
'
'If cboPlan.Text <> "[TODOS]" Then
'   strSQL = strSQL & " and Det.cod_plan = '" & cboPlan.Text & "'"
'End If
'
'
'Call OpenRecordSet(rs, strSQL)
'curMonto = 0
'
'lsw.ListItems.Clear
'prgBar.Max = rs.RecordCount + 1
'prgBar.Value = 1
'
'Do While Not rs.EOF
'
'  scMain.Caption = "Procesando registro # " & prgBar.Value & " de " & prgBar.Max
'  scMain.Refresh
'
'    Set itmX = lsw.ListItems.Add(, , rs!Cod_Operadora)
'        itmX.SubItems(1) = rs!cod_Plan
'        itmX.SubItems(2) = rs!COD_CONTRATO
'        itmX.SubItems(3) = rs!Cedula
'        itmX.SubItems(4) = rs!Nombre & ""
'        itmX.SubItems(5) = Format(rs!Monto, "Standard")
'        itmX.SubItems(6) = Format(rs!fecha, "dd/mm/yyyy")
'        itmX.SubItems(7) = "Planilla"
'        itmX.SubItems(8) = rs!nCon
'        itmX.SubItems(9) = rs!Institucion
'
'       curMonto = curMonto + rs!Monto
'
'  prgBar.Value = prgBar.Value + 1
'  rs.MoveNext
'Loop
'rs.Close
'
'Set itmX = lsw.ListItems.Add(, , "")
'        itmX.SubItems(5) = "____"
'Set itmX = lsw.ListItems.Add(, , "TOTAL")
'        itmX.SubItems(5) = Format(curMonto, "Standard")
'        itmX.ForeColor = vbBlue
'
'
'Me.MousePointer = vbDefault
'
'MsgBox "Consulta Finalizada satisfactoriamente...", vbInformation
'
'Exit Sub
'
'vError:
' Me.MousePointer = vbDefault
' MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdBuscar_Click()

If cboTipo.Text = "Cobros en Retenciones no registrados" Then
    Call sbBuscarRetNoReg
Else
    Call sbBuscarRegNoRet
End If

End Sub

Private Sub cmdInfo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strSQL = "select isnull(sum(Principal),0) as Monto, count(*) as Casos" _
       & " from vCRDsReportesMov where tcon in('1','PLA') and fecha between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and codigo in(select codigo_ase" _
       & " from fnd_planes where cod_plan = '" & cboPlan.ItemData(cboPlan.ListIndex) & "')"
Call OpenRecordSet(rs, strSQL)
    txtPlCrdMnt.Text = Format(rs!Monto, "Standard")
    txtPlCrdCasos.Text = Format(rs!Casos, "###,###,##0")
rs.Close
       
       
strSQL = "select isnull(sum(monto),0) as Monto, count(*) as Casos" _
       & " from fnd_contratos_Detalle where tcon in('1','PLA') and fecha between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and cod_plan = '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
    txtPlFndMnt.Text = Format(rs!Monto, "Standard")
    txtPlFndCasos.Text = Format(rs!Casos, "###,###,##0")
rs.Close

txtPlDifMnt.Text = Format(CCur(txtPlCrdMnt.Text) - CCur(txtPlFndMnt.Text), "Standard")
txtPlDifCasos.Text = Format(CLng(txtPlCrdCasos.Text) - CLng(txtPlFndCasos.Text), "###,###,##0")

If fraInfo.Visible Then
    fraInfo.Visible = False
Else
    fraInfo.Visible = True
End If
Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
Dim strSQL As String

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)

cboTipo.AddItem "Cobros en Retenciones no registrados"
cboTipo.AddItem "Cobros registrados no cobrados"
cboTipo.Text = "Cobros en Retenciones no registrados"

cboAccion.AddItem "Ninguna"
 cboAccion.ItemData(cboAccion.ListCount - 1) = CStr(0)
cboAccion.AddItem "Activar, Acreditar y Bloquear Deducción"
 cboAccion.ItemData(cboAccion.ListCount - 1) = CStr(1)
cboAccion.AddItem "Activar y Acreditar"
 cboAccion.ItemData(cboAccion.ListCount - 1) = CStr(2)
cboAccion.Text = "Ninguna"

strSQL = "select cod_plan as 'IdX', Descripcion as  'ItmX' from fnd_planes"
Call sbCbo_Llena_New(cboPlan, strSQL, True, True)

vGrid.MaxRows = 0


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
scMain.Width = Me.Width

vGrid.Width = Me.Width - 200
vGrid.Height = Me.Height - (vGrid.top + 850)



End Sub

Private Sub imgInfo_Click()
fraInfo.Visible = False
End Sub
