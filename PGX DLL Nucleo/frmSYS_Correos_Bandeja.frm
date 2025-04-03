VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmSYS_Correos_Bandeja 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Correos"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   16170
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   4335
      _Version        =   1310723
      _ExtentX        =   7646
      _ExtentY        =   3413
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.RadioButton rbExport 
      Height          =   255
      Index           =   0
      Left            =   12960
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Bandeja"
      BackColor       =   16777215
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
      Appearance      =   17
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   10800
      TabIndex        =   0
      Top             =   1080
      Width           =   495
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   6
      Picture         =   "frmSYS_Correos_Bandeja.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmSYS_Correos_Bandeja.frx":0700
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   8040
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      Height          =   330
      Left            =   9360
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtAsunto 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   3855
      _Version        =   1310723
      _ExtentX        =   6800
      _ExtentY        =   556
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPara 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3855
      _Version        =   1310723
      _ExtentX        =   6800
      _ExtentY        =   556
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   14295
      _Version        =   524288
      _ExtentX        =   25215
      _ExtentY        =   4048
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
      MaxCols         =   11
      SpreadDesigner  =   "frmSYS_Correos_Bandeja.frx":0FD1
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.RadioButton rbExport 
      Height          =   255
      Index           =   1
      Left            =   14160
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Resumen"
      BackColor       =   16777215
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
      Appearance      =   17
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   4335
      _Version        =   1310723
      _ExtentX        =   7646
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Resumen:"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Asunto:"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   8
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Para:"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   8
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Correos: Bandeja de Salida"
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
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   4452
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   4
      Top             =   840
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fechas"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   8
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmSYS_Correos_Bandeja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "exec spSys_Mail_Consulta_General '" & Trim(txtPara.Text) & "','" & Trim(txtAsunto.Text) _
       & "','" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00','" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59','D'"
Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)

'Resumen
strSQL = "exec spSys_Mail_Consulta_General '" & Trim(txtPara.Text) & "','" & Trim(txtAsunto.Text) _
       & "','" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00','" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59','R'"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_SMTP)
      itmX.SubItems(1) = Format(rs!Correos & "", "###,###,##0")
      itmX.SubItems(2) = rs!EstadoDesc
      itmX.SubItems(3) = rs!Anio
      itmX.SubItems(4) = rs!MesId
      itmX.SubItems(5) = rs!Mes
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click(Index As Integer)

Dim vHeaders As vGridHeaders

On Error GoTo vError


Select Case True
    Case rbExport.Item(0).Value

            vHeaders.Columnas = 11
            vHeaders.Headers(1) = "Id Mail"
            vHeaders.Headers(2) = "Cuenta"
            vHeaders.Headers(3) = "Para"
            vHeaders.Headers(4) = "Asunto"
            vHeaders.Headers(5) = "Estado"
            vHeaders.Headers(6) = "Fecha"
            vHeaders.Headers(7) = "Fecha Envío"
            vHeaders.Headers(8) = "Usuario"
            vHeaders.Headers(9) = "Año"
            vHeaders.Headers(10) = "Mes Id"
            vHeaders.Headers(11) = "Mes"
         
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Correos_Bandeja_Salida")
    
    Case rbExport.Item(1).Value

         Call Excel_Exportar_Lsw(lsw)

End Select



Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)


vGrid.MaxRows = 0
vGrid.MaxCols = 11

With lsw.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 1800
    .Add , , "Cantidad", 1200, vbRightJustify
    .Add , , "Estado", 2800
    .Add , , "Año", 1000, vbCenter
    .Add , , "Mes Id", 1200, vbCenter
    .Add , , "Mes", 2100
    
End With


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 150
scTitulo.Width = vGrid.Width
lsw.Width = vGrid.Width


vGrid.Height = Me.Height - (vGrid.Top + scTitulo.Height + lsw.Height + 450)

scTitulo.Top = vGrid.Top + vGrid.Height + 100
lsw.Top = scTitulo.Top + scTitulo.Height + 100

End Sub
