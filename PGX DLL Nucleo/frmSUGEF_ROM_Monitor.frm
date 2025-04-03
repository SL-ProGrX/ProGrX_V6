VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmSUGEF_ROM_Monitor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "SUGEF: Monitor Operaciones Múltiples (ROM)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20558
      _ExtentY        =   5530
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
      Appearance      =   21
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswDet 
      Height          =   2535
      Left            =   0
      TabIndex        =   4
      Top             =   6000
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20558
      _ExtentY        =   4471
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
      Appearance      =   21
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnAdministrador 
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Seg. Administracion"
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20558
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   960
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
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
      CustomFormat    =   "MMMM yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Consultar"
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
      Picture         =   "frmSUGEF_ROM_Monitor.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmSUGEF_ROM_Monitor.frx":0719
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   2
      Left            =   10080
      TabIndex        =   11
      Top             =   5520
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "ROM"
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
      Picture         =   "frmSUGEF_ROM_Monitor.frx":0883
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   3
      Left            =   11040
      TabIndex        =   12
      Top             =   5520
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmSUGEF_ROM_Monitor.frx":0F9C
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   330
      Left            =   3720
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
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
      Text            =   "10000"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTC 
      Height          =   330
      Left            =   6360
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
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
      Text            =   "1"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Cambio"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   13
      Top             =   1320
      Width           =   735
      _Version        =   1572864
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Monto $"
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
   Begin XtremeShortcutBar.ShortcutCaption scDetalle 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5520
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20558
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Detalle de Operaciones Encontradas:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption scCorte 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20558
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Casos Localizados:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   735
      _Version        =   1572864
      _ExtentX        =   1296
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monitor ROM"
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
      Height          =   612
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSUGEF_ROM_Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, vROM As Long

Private Sub sbExportar(pTipo As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True
If pTipo = 1 Then
    Call Excel_Exportar_Lsw(lsw, ProgressBarX)
Else
    Call Excel_Exportar_Lsw(lswDet, ProgressBarX)
End If

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbConsulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSUGEF_ROM_Consulta '" & Format(dtpCorte.Value, "yyyy-mm-dd") & "', " & CCur(txtMonto.Text) _
       & ", " & CCur(txtTC.Text) & ", 'R', ''"
Call OpenRecordSet(rs, strSQL)
       
lswDet.ListItems.Clear
scDetalle.Caption = "Detalle de Operaciones Encontradas:"
scDetalle.Tag = ""
      
With lsw.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Cedula)
            itmX.SubItems(1) = rs!Nombre
            itmX.SubItems(2) = rs!TipoId_Desc
            itmX.SubItems(3) = rs!Estado_Persona
            itmX.SubItems(4) = rs!ROM_ID & ""
            itmX.SubItems(5) = rs!Transacciones
            itmX.SubItems(6) = Format(rs!Monto_Col, "Standard")
            itmX.SubItems(7) = Format(rs!Monto_Dol, "Standard")
            itmX.SubItems(8) = Format(rs!Salario_Col, "Standard")
            itmX.SubItems(9) = Format(rs!Salario_Dol, "Standard")
            itmX.SubItems(10) = Format(rs!Fecha_Nac, "dd-MM-yyyy")
            itmX.SubItems(11) = rs!TelMovil
            itmX.SubItems(12) = rs!Email
        rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbConsulta_Detalle()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSUGEF_ROM_Consulta '" & Format(dtpCorte.Value, "yyyy-mm-dd") & "', " & CCur(txtMonto.Text) _
       & ", " & CCur(txtTC.Text) & ", 'D', '" & scDetalle.Tag & "'"
Call OpenRecordSet(rs, strSQL)
       
With lswDet.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!TDoc)
            itmX.SubItems(1) = rs!NDoc
            itmX.SubItems(2) = Format(rs!Monto_Col, "Standard")
            itmX.SubItems(3) = Format(rs!Monto_Dol, "Standard")
            itmX.SubItems(4) = rs!Forma_Pago_Desc
            itmX.SubItems(5) = rs!Descripcion
            itmX.SubItems(6) = rs!Origen_Recursos_Desc
            itmX.SubItems(7) = rs!Pagador_Desc
            itmX.SubItems(8) = rs!Fecha
            itmX.SubItems(9) = rs!Usuario
            
        rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub btnAccion_Click(Index As Integer)


Select Case Index
    Case 0 'Consultar
        Call sbConsulta
        
    Case 1 'Exportar Main
        Call sbExportar(1)
        
    Case 2 'Buscar ROM
        MsgBox "Acá se mostrará la captura para confeccion del ROM", vbInformation
    
    Case 3 'Exportar Detalle
        Call sbExportar(2)
        
End Select

End Sub

Private Sub dtpCorte_Change()
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "select dbo.fxSUGEF_Tipo_Cambio('" & Format(dtpCorte.Value, "yyyy-mm-dd") & "') as 'TC'"
Call OpenRecordSet(rs, strSQL)

txtTC.Text = Format(rs!TC, "Standard")


lsw.ListItems.Clear
lswDet.ListItems.Clear

Exit Sub

vError:


End Sub

Private Sub Form_Load()

vModulo = 10
 
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
vPaso = True

dtpCorte.Value = fxFechaServidor
 
With lsw.ColumnHeaders
  .Clear
  .Add , , "Cédula", 1500
  .Add , , "Nombre", 3000
  .Add , , "Tipo Id", 1400
  .Add , , "Estado", 2100, vbCenter
  
  .Add , , "ROM Id", 1400, vbCenter
  .Add , , "No. Transac.", 2100, vbCenter
  
  .Add , , "Monto Col", 2100, vbRightJustify
  .Add , , "Monto Dol", 2100, vbRightJustify
  .Add , , "Salario Col", 2100, vbRightJustify
  .Add , , "Salario Dol", 2100, vbRightJustify
  
  .Add , , "Fecha Nac.", 2100, vbCenter
  .Add , , "Tel. Móvil", 2100, vbCenter
  .Add , , "Email", 3100
  
End With



With lswDet.ColumnHeaders
  .Clear
  .Add , , "T.Doc.", 1200
  .Add , , "N.Doc.", 3000
  .Add , , "Monto", 1800, vbRightJustify
  .Add , , "Monto $", 1800, vbRightJustify
  .Add , , "Forma Pago", 1800, vbRightJustify
  
  .Add , , "Descripción", 3000
  .Add , , "Origen Recursos", 3000
  .Add , , "Pagador", 3000
  
  .Add , , "Fecha", 2500, vbCenter
  .Add , , "Usuario", 2500, vbCenter
  
End With

vPaso = False


Call Formularios(Me)

btnAccion(0).Tag = btnAdministrador.Tag
btnAccion(1).Tag = btnAdministrador.Tag
btnAccion(2).Tag = btnAdministrador.Tag
btnAccion(3).Tag = btnAdministrador.Tag


Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
ProgressBarX.Width = Me.Width

scCorte.Width = Me.Width
scDetalle.Width = Me.Width

lsw.Width = Me.Width - 250
lswDet.Width = lsw.Width

Dim pAlto As Long

pAlto = 2535

lsw.Height = Me.Height - (lsw.Top + 1100 + pAlto)

scDetalle.Top = lsw.Top + lsw.Height + 150
lswDet.Top = scDetalle.Top + scDetalle.Height + 20


btnAccion(2).Top = scDetalle.Top
btnAccion(3).Top = scDetalle.Top

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub


scDetalle.Caption = "Operaciones Encontradas para: " & Item.SubItems(1)
scDetalle.Caption = Item.Text

If IsNumeric(Item.SubItems(4)) Then
    vROM = Item.SubItems(4)
Else
    vROM = 0
End If

Call sbConsulta_Detalle

End Sub


Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

On Error GoTo vError

txtMonto.Text = Format(10000, "Standard")

Call dtpCorte_Change

Exit Sub

vError:

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
