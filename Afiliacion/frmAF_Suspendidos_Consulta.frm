VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Suspendidos_Consulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Afiliación: Consulta de Asociados Suspendidos"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
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
      Left            =   2400
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
   Begin XtremeSuiteControls.ComboBox cboEvento 
      Height          =   330
      Left            =   5040
      TabIndex        =   4
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   5040
      TabIndex        =   5
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
      Left            =   8280
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   11280
      TabIndex        =   11
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Picture         =   "frmAF_Suspendidos_Consulta.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   12600
      TabIndex        =   12
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Picture         =   "frmAF_Suspendidos_Consulta.frx":0700
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   8280
      TabIndex        =   13
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5895
      Left            =   -120
      TabIndex        =   14
      Top             =   2400
      Width           =   14655
      _Version        =   524288
      _ExtentX        =   25850
      _ExtentY        =   10398
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
      MaxCols         =   28
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Suspendidos_Consulta.frx":086A
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTodas 
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   1440
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   15
      Top             =   960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   14535
      _Version        =   1441793
      _ExtentX        =   25638
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Resultados de la busqueda:"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Evento"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Index           =   5
      Left            =   7320
      TabIndex        =   7
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Asociados con Gestión en su condición"
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
      TabIndex        =   1
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
Attribute VB_Name = "frmAF_Suspendidos_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 25
    vHeaders.Headers(1) = "Cédula"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "F.Ingreso"
    vHeaders.Headers(4) = "Estado"
    vHeaders.Headers(5) = "Email"
    vHeaders.Headers(6) = "Tel.Celular"
    vHeaders.Headers(7) = "Tel.Habitación"
    vHeaders.Headers(8) = "Tel.Trabajo"
    vHeaders.Headers(9) = "F.Ult.Apo.Obrero"
    vHeaders.Headers(10) = "F.Ult.Apo.Patronal"
    vHeaders.Headers(11) = "Dias Obrero"
    vHeaders.Headers(12) = "Dias Patronal"
    vHeaders.Headers(13) = "F.Activa"
    vHeaders.Headers(14) = "F.Suspende"
    vHeaders.Headers(15) = "Empresa/Inst."
    vHeaders.Headers(16) = "UP"
    vHeaders.Headers(17) = "UP Descripción"
    vHeaders.Headers(18) = "Promotor"
    vHeaders.Headers(19) = "Aporte Obrero"
    vHeaders.Headers(20) = "Capitalización"
    vHeaders.Headers(21) = "Aporte Patronal"
    vHeaders.Headers(22) = "Fondos Acumulado"
    vHeaders.Headers(23) = "Saldo Créditos"
    vHeaders.Headers(24) = "Morosidad"
    
    vHeaders.Headers(25) = "Provincia"
    vHeaders.Headers(26) = "Cantón"
    vHeaders.Headers(27) = "Distrito"
    vHeaders.Headers(28) = "Dirección"

    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Control_" & cboEvento.Text & "_" + cboTipo.Text)
 
End Sub

Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)

If chkTodas.Value = xtpChecked Then
    strSQL = "exec spPAT_AsociadosSinAportes_Gestion_Consulta '1900-01-01'" _
           & ", '2200-01-01 23:59:59'," & cboTipo.ItemData(cboTipo.ListIndex) _
           & ", '" & Mid(txtCedula.Text, 1, 20) & "', '" & Mid(txtNombre.Text, 1, 100) & "'"
Else
    strSQL = "exec spPAT_AsociadosSinAportes_Gestion_Consulta '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
           & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'," & cboTipo.ItemData(cboTipo.ListIndex) _
           & ", '" & Mid(txtCedula.Text, 1, 20) & "', '" & Mid(txtNombre.Text, 1, 100) & "'"
End If

Call OpenRecordSet(rs, strSQL)

With vGrid
  .MaxRows = 0
  Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = Trim(rs!Cedula)
     .Col = 2
     .Value = Trim(rs!Nombre)
     .Col = 3
     .Text = Format(rs!FechaIngreso, "yyyy-mm-dd")
     .Col = 4
     .Text = Trim(rs!Estado_Asociado)
     .Col = 5
     .Text = Trim(rs!Email)
     .Col = 6
     .Text = Trim(rs!Telefono_Celular & "")
     .Col = 7
     .Text = Trim(rs!Telefono_Habitacion & "")
     .Col = 8
     .Text = Trim(rs!Telefono_Trabajo)
     
     .Col = 9
     .Text = Format(rs!UltimoAporteObrero, "yyyy-mm-dd")
     .Col = 10
     .Text = Format(rs!UltimoAportePatronal, "yyyy-mm-dd")
     .Col = 11
     .Text = CStr(rs!Dias_Aporte_Obrero & "")
     .Col = 12
     .Text = CStr(rs!Dias_Aporte_Patronal & "")
     
     .Col = 13
     .Text = Format(rs!Fecha_Activo & "", "yyyy-mm-dd")
     .Col = 14
     .Text = Format(rs!Fecha_Suspendido & "", "yyyy-mm-dd")
     
     .Col = 15
     .Text = Trim(rs!Institucion & "")
     .Col = 16
     .Text = Trim(rs!UP & "")
     .Col = 17
     .Text = Trim(rs!UP_Desc & "")
     .Col = 18
     .Text = Trim(rs!Promotor_Desc)
     
     .Col = 19
     .Text = Format(rs!Aporte_Obrero, "Standard")
     .Col = 20
     .Text = Format(rs!Capitalización, "Standard")
     .Col = 21
     .Text = Format(rs!Aporte_Patronal, "Standard")
     .Col = 22
     .Text = Format(rs!Fondos_Acumulado, "Standard")
     .Col = 23
     .Text = Format(rs!Creditos_Saldo, "Standard")
     .Col = 24
     .Text = Format(rs!Morosidad, "Standard")
     
     
     .Col = 25
     .Text = Trim(rs!Provincia_Desc & "")
     .Col = 26
     .Text = Trim(rs!Canton_Desc & "")
     .Col = 27
     .Text = Trim(rs!Distrito_Desc & "")
     .Col = 28
     .Text = Trim(rs!Direccion & "")
     

   rs.MoveNext
  Loop
  rs.Close
End With

scMain.Caption = "Casos localizados: " & vGrid.MaxRows

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnAccion_Click(Index As Integer)
Select Case Index
    Case 0 'Buscar
        Call sbBuscar
        
    Case 1 'Exportar
        Call sbExportar
        
End Select
End Sub

Private Sub cboEvento_Click()

cboTipo.Clear
If cboEvento.Text = "Activar" Then
    cboTipo.AddItem "Activos"
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(1)
    cboTipo.Text = "Activos"

Else
    cboTipo.AddItem "Suspendidos"
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(0)
    cboTipo.AddItem "Condición Especial"
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(2)
    cboTipo.AddItem "Suspende + C.E."
    cboTipo.ItemData(cboTipo.ListCount - 1) = CStr(3)
    cboTipo.Text = "Suspende + C.E."
End If


End Sub

Private Sub chkTodas_Click()
If chkTodas.Value = xtpChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub Form_Load()
vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle
vGrid.MaxRows = 0


cboEvento.AddItem "Activar"
cboEvento.AddItem "Inactivar"
cboEvento.Text = "Activar"

Call cboEvento_Click


dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -60, dtpCorte.Value)


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 320
vGrid.Height = Me.Height - (vGrid.Top + 700)

scMain.Width = vGrid.Width

End Sub
