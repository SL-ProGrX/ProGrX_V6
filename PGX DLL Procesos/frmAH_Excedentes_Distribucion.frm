VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAH_Excedentes_Distribucion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Distribución de Excedentes"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   7011
      _StockProps     =   77
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbMain 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   2566
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboCorte 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnEXT 
         Height          =   255
         Index           =   0
         Left            =   10440
         TabIndex        =   4
         Top             =   5160
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_Excedentes_Distribucion.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   330
         Left            =   5760
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   582
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   0
         Left            =   10320
         TabIndex        =   13
         Top             =   1080
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_Distribucion.frx":016A
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   1
         Left            =   10800
         TabIndex        =   14
         Top             =   1080
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_Distribucion.frx":088A
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   375
         Left            =   11400
         TabIndex        =   15
         Top             =   1080
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_Distribucion.frx":0E2E
      End
      Begin XtremeSuiteControls.ComboBox cboBase 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   330
         Left            =   5760
         TabIndex        =   19
         Top             =   720
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSample 
         Height          =   330
         Left            =   6960
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblPorcentaje 
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   720
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "[ % ] Carga"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Base de Aplicacion"
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
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto a distribuir"
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Corte"
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
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Distribución"
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
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7646
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   1245
      Visible         =   0   'False
      Width           =   12375
      _Version        =   1572864
      _ExtentX        =   21828
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Tabla de Distribución de Excedentes"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Periodo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
      Caption         =   "Distribución de Excedentes"
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
      Height          =   492
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   5172
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmAH_Excedentes_Distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnAccion_Click(Index As Integer)

On Error GoTo vError


If InStr(1, UCase(cboPeriodo.Text), "CERRADO") > 0 Then
        MsgBox "El periodo ya se encuentra cerrado, no pueden realizar cambios!", vbExclamation
        Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spExc_Montos_Distribucion_Tabla_Add " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & IIf((Index = 0), "A", "B") _
        & "', '" & glogon.Usuario & "', '" & cboCorte.Text & " 23:59', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', '" & cboBase.ItemData(cboBase.ListIndex) _
        & "', " & CCur(txtMonto.Text) & ", " & CCur(txtPorcentaje.Text)
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Call sbLista

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboBase_Click()
If vPaso Then Exit Sub

Call sbCalculo

End Sub

Private Sub cboCorte_Click()
If vPaso Then Exit Sub

Call sbCalculo

End Sub

Private Sub sbLista()

On Error GoTo vError

strSQL = "exec spExc_Mnt_Distribuir " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Corte)
     itmX.SubItems(1) = rs!Mes
     itmX.SubItems(2) = Format(rs!Monto_Proyectado, "Standard")
     itmX.SubItems(3) = Format(rs!Monto_Cargado, "Standard")
     itmX.SubItems(4) = Format(rs!Porc_Distribuido, "Standard")
     itmX.SubItems(5) = Format(rs!Monto_Real, "Standard")
     itmX.SubItems(6) = Format(rs!Diferencia, "Standard")
     itmX.SubItems(7) = Format(rs!Monto_Prorrateado, "Standard")
     itmX.SubItems(8) = rs!Base_Calculo_Desc
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboPeriodo_Click()
If vPaso Then Exit Sub
If cboPeriodo.ListCount = 0 Then Exit Sub


On Error GoTo vError

vPaso = True

strSQL = "exec spExc_Periodo_Meses " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
Call sbCbo_Llena_New(cboCorte, strSQL, False, True)

vPaso = False

Call sbLista

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCalculo()

On Error GoTo vError

txtMonto.Text = Format(0, "Standard")

If cboPeriodo.ListCount = 0 Then Exit Sub
If cboCorte.ListCount = 0 Then Exit Sub
If cboBase.ListCount = 0 Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spExc_Mnt_Distribuir_Calculo " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & cboCorte.ItemData(cboCorte.ListIndex) _
       & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', '" & cboBase.ItemData(cboBase.ListIndex) & "', " & CCur(txtPorcentaje.Text)

Call OpenRecordSet(rs, strSQL)

txtMonto.Text = Format(rs!Monto, "Standard")

rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboTipo_Click()
If vPaso Then Exit Sub

If cboTipo.ListCount = 0 Then Exit Sub

If cboTipo.ItemData(cboTipo.ListIndex) = "C" Then
    cboBase.Enabled = True
    txtPorcentaje.Locked = False
    txtPorcentaje.BackColor = vbWhite
    
    txtMonto.Locked = True
    txtMonto.BackColor = txtSample.BackColor
Else
    cboBase.Enabled = False
    cboBase.Text = cboTipo.Text
    
    txtPorcentaje.Locked = True
    txtPorcentaje.BackColor = txtSample.BackColor
    
    txtMonto.Locked = False
    txtMonto.BackColor = vbWhite
End If

Call sbCalculo

End Sub

Private Sub Form_Activate()
vModulo = 2
End Sub


Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
  .Clear
  .Add , , "Corte", 2100
  .Add , , "Mes", 1600
  .Add , , "Monto Proyectado", 2000, vbRightJustify
  .Add , , "Monto Cargado", 2000, vbRightJustify
  .Add , , "% Carga", 1200, vbRightJustify
  .Add , , "Monto Real", 2000, vbRightJustify
  .Add , , "Diferencia", 2000, vbRightJustify
  .Add , , "Monto Prorrateado", 2000, vbRightJustify
  .Add , , "Base Cálculo", 1500, vbCenter
End With


 With cboBase
    .Clear
    .AddItem "Real Contable"
    .ItemData(.ListCount - 1) = "R"
    .AddItem "Proyectado"
    .ItemData(.ListCount - 1) = "P"
    .AddItem "Prorrateado"
    .ItemData(.ListCount - 1) = "T"
    .Text = "Proyectado"
 End With
 
 
 strSQL = "select A.USUARIO,  A.ACTIVO, A.CARGA, A.REAL, A.PROYECTADO, A.PRORRATEADO" _
        & "  from EXC_APLICADORES A left join USUARIOS U on A.USUARIO = U.NOMBRE" _
        & " Where U.ESTADO = 'A' and A.Activo = 1"
 Call OpenRecordSet(rs, strSQL)
 
 With cboTipo
    .Clear
    
    If Not rs.BOF And Not rs.BOF Then
       
       If rs!Carga = 1 Then
        .AddItem "Cargado [%]"
        .ItemData(.ListCount - 1) = "C"
        .Text = "Cargado [%]"
       End If
       If rs!Real = 1 Then
        .AddItem "Real Contable"
        .ItemData(.ListCount - 1) = "R"
       End If
       If rs!Proyectado = 1 Then
        .AddItem "Proyectado"
        .ItemData(.ListCount - 1) = "P"
       End If
       If rs!Prorrateado = 1 Then
        .AddItem "Prorrateado"
        .ItemData(.ListCount - 1) = "T"
       End If
    
    End If
 End With
rs.Close

Call Formularios(Me)

btnAccion(1).Tag = btnAccion(0).Tag

Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select IdX, ItmX from vExc_Periodos order by Idx desc"
Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)

vPaso = False


Call cboPeriodo_Click


Me.MousePointer = vbDefault

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")

vError:

End Sub

Private Sub txtPorcentaje_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbCalculo
End Sub
