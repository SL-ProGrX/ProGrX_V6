VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmAH_ExcedentesAjuste 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Ajustes"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ExcedentesAjustes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
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
      Appearance      =   20
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9960
      Top             =   720
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   3120
      Width           =   6615
      _Version        =   1441793
      _ExtentX        =   11668
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAjusteId 
      Height          =   555
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   555
      Left            =   3840
      TabIndex        =   11
      Top             =   600
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6800
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   255
      Index           =   0
      Left            =   10800
      TabIndex        =   13
      ToolTipText     =   "Exportar Listado a Excel"
      Top             =   3120
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   444
      _ExtentY        =   444
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "ExcedentesAjustes.frx":030A
   End
   Begin XtremeSuiteControls.FlatEdit txtAjuste 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   315
      Left            =   960
      TabIndex        =   19
      Top             =   7560
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAjustePositivo 
      Height          =   315
      Left            =   3720
      TabIndex        =   20
      Top             =   7560
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAjusteNegativo 
      Height          =   315
      Left            =   7200
      TabIndex        =   21
      Top             =   7560
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   315
      Left            =   10320
      TabIndex        =   22
      ToolTipText     =   "Cantidad de Casos a Mostrar"
      Top             =   7560
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Text            =   "100"
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1680
      TabIndex        =   23
      ToolTipText     =   "Nuevo"
      Top             =   120
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "ExcedentesAjustes.frx":0BDB
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   24
      ToolTipText     =   "Editar"
      Top             =   120
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "ExcedentesAjustes.frx":120D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   3120
      TabIndex        =   25
      ToolTipText     =   "Eliminar"
      Top             =   120
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "ExcedentesAjustes.frx":1808
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3720
      TabIndex        =   26
      ToolTipText     =   "Guardar"
      Top             =   120
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "ExcedentesAjustes.frx":1DAC
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   4080
      TabIndex        =   27
      ToolTipText     =   "Deshacer"
      Top             =   120
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "ExcedentesAjustes.frx":24DD
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   4560
      TabIndex        =   28
      ToolTipText     =   "Reporte"
      Top             =   120
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "ExcedentesAjustes.frx":2BDD
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   330
      Left            =   5520
      TabIndex        =   30
      Top             =   1680
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   795
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   9495
      _Version        =   1441793
      _ExtentX        =   16748
      _ExtentY        =   1402
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   29
      Top             =   1680
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Periodo"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   18
      Top             =   7560
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Lineas:"
      BackColor       =   -2147483633
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   17
      Top             =   7560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ajuste Negativo:"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   16
      Top             =   7560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ajuste Positivo:"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   7560
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Casos:"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Listado de Ajustes Pendientes                 Filtros:"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Ajuste Id"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Detalle"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ajuste"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cédula"
      BackColor       =   -2147483633
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmAH_ExcedentesAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vEdita As Boolean, vPaso As Boolean

Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub


Private Sub btnBarra_Click(Index As Integer)

Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpia
        txtCedula.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
      If txtAjusteId.Text = "" Then
        MsgBox "Seleccione una Mejora o Retiro de la lista del activo para modificacion...", vbInformation
      Else
        vEdita = True
        txtAjuste.SetFocus
        Call sbBarra_Accion("Editar")
      End If
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If txtAjusteId.Text = "" Then
        Call sbLimpia
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      End If
    
    Case 5 'REPORTES
   
End Select

End Sub


Private Function fxValida() As Boolean

Dim vMensaje As String

vMensaje = ""
fxValida = True

If Trim(txtCedula) = "" Or Trim(txtDetalle) = "" Then
  vMensaje = vMensaje & "- No Cedula o Detalle no es válido!" & vbCrLf
End If

If Not IsNumeric(txtAjuste.Text) Then
  vMensaje = vMensaje & "- Monto del Ajuste no es valido!"" & vbCrLf"
End If

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

If Trim(txtAjusteId.Text) = "" Then Exit Sub

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then

    strSQL = "exec spExc_Ajustes_Del  " & txtAjusteId.Text & ", '" & glogon.Usuario & "'"
    
    Call OpenRecordSet(rs, strSQL)
    If rs!Aplicado = 1 Then
        txtAjuste.Text = rs!Ajuste_Id
        Call Bitacora("Borra", "Excedente Ajuste Id: " & txtAjusteId.Text & ", Cedula: " & txtCedula.Text & ", Ajuste: " & txtAjuste.Text)
    End If
     
     Call sbListado_Load
     
     Call sbBarra_Accion("nuevo")
     Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbLimpia()

txtAjusteId.Text = ""
txtEstado.Text = ""
txtCedula.Text = ""
txtNombre.Text = ""
txtAjuste.Text = "0"
txtDetalle.Text = ""
     
End Sub

Private Sub sbListado_Load()

On Error GoTo vError

Dim pFiltro As String

pFiltro = fxSysCleanTxtInject(txtFiltro.Text)
txtLineas.Text = fxSysCleanTxtInject(txtLineas.Text)

Dim pPositivos As Currency, pNegativos As Currency

If Not IsNumeric(txtLineas.Text) Then
    txtLineas.Text = "100"
End If

strSQL = "Select Top " & txtLineas.Text & " A.*, S.Nombre, P.ItmX as 'PeriodoDesc'" _
       & " from Exc_Ajustes A inner join Socios S on A.cedula = S.cedula" _
       & " inner join vExc_Periodos P on P.IdX = A.id_Periodo" _
       & " where A.Estado = 'P'" _
       & " and S.cedula like '%" & pFiltro & "%'" _
       & " or S.Nombre like '%" & pFiltro & "%'"
lsw.ListItems.Clear

vPaso = True
pPositivos = 0
pNegativos = 0

Call OpenRecordSet(rs, strSQL)
 
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Ajuste_Id)
       itmX.SubItems(1) = Trim(rs!Cedula)
       itmX.SubItems(2) = rs!Nombre
       itmX.SubItems(3) = Format(rs!Ajuste, "Standard")
       itmX.SubItems(4) = Trim(rs!Detalle)
       itmX.SubItems(5) = Trim(rs!PeriodoDesc)
       itmX.SubItems(6) = Trim(rs!Registro_Usuario & "")
       itmX.SubItems(7) = Trim(rs!Registro_Fecha & "")
   
       itmX.Tag = rs!Id_Periodo
       
    If rs!Ajuste < 0 Then
        pNegativos = pNegativos + Abs(rs!Ajuste)
    Else
        pPositivos = pPositivos + rs!Ajuste
    End If
       
   rs.MoveNext
Loop
rs.Close

vPaso = False

txtCasos.Text = Format(lsw.ListItems.Count, "###,##0")
txtAjustePositivo.Text = Format(pPositivos, "Standard")
txtAjusteNegativo.Text = Format(pNegativos, "Standard")

Call sbLimpia

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub sbCedula_Load()

strSQL = "Select Top 1 A.*,S.Nombre, P.ItmX as 'PeriodoDesc' " _
       & " from Exc_Ajustes A inner join Socios S on A.cedula = S.cedula" _
       & " inner join vExc_Periodos P on A.Id_Periodo = P.IdX" _
       & " where A.cedula = '" & Trim(txtCedula.Text) _
       & "' order by AJUSTE_ID desc"

Call OpenRecordSet(rs, strSQL)
 
 
If Not rs.EOF And Not rs.BOF Then
   vEdita = True
   
   txtAjusteId.Text = rs!Ajuste_Id
   
   Select Case rs!Estado
    Case "P"
       txtEstado.Text = "Pendiente"
    Case "C"
       txtEstado.Text = "Cancelado"
   End Select
   
   txtAjuste = Format(rs!Ajuste, "Standard")
   txtDetalle = Trim(rs!Detalle)
   
   Call sbCboAsignaDato(cboPeriodo, rs!PeriodoDesc, True, rs!Id_Periodo)
   
   Call sbBarra_Accion("EDITAR")

Else
   txtAjusteId.Text = ""
   txtEstado.Text = ""
   txtNombre.Text = fxNombre(txtCedula.Text)
End If

rs.Close

End Sub

Private Sub sbGuardar()
Dim pAjuste As Currency

On Error GoTo vError

txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtDetalle.Text = fxSysCleanTxtInject(txtDetalle.Text)

 pAjuste = CCur(txtAjuste.Text)


strSQL = "exec spExc_Ajustes_Add " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCedula.Text _
       & "', " & pAjuste & ", '" & txtDetalle.Text & "', '" & glogon.Usuario & "', 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Aplicado = 1 Then
    txtAjuste.Text = rs!Ajuste_Id
    Call Bitacora("Registra", "Excedente Ajuste Id: " & txtAjusteId.Text & ", Cedula: " & txtCedula.Text & ", Ajuste: " & txtAjuste.Text)
End If
 
 Call sbBarra_Accion("NUEVO")
 
 Call sbListado_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click(Index As Integer)
Call Excel_Exportar_Lsw(lsw, ProgressBarX)
End Sub

Private Sub Form_Load()

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Cédula", 1500, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Ajuste", 1500, vbRightJustify
    .Add , , "Detalle", 3500
    .Add , , "Periodo", 2500
    .Add , , "Reg.Usuario", 1800, vbCenter
    .Add , , "Reg.Fecha", 1800
End With

 Call sbBarra_Accion("nuevo")
 Call sbLimpia

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
If lsw.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError
   
txtAjusteId.Text = Item.Text
txtEstado.Text = "Pendiente"
txtCedula.Text = Item.SubItems(1)
txtNombre.Text = Item.SubItems(2)
   

txtAjuste.Text = Format(CCur(Item.SubItems(3)), "Standard")
txtDetalle.Text = Item.SubItems(4)

Call sbCboAsignaDato(cboPeriodo, Item.SubItems(5), True, Item.Tag)
     
Call sbBarra_Accion("ACTIVO")

txtAjuste.SetFocus

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


strSQL = "select IdX, ItmX from vExc_Periodos order by Idx desc"
Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)

Call sbListado_Load

End Sub

Private Sub txtAjuste_GotFocus()
On Error GoTo vError
  txtAjuste.Text = CCur(txtAjuste.Text)
vError:
End Sub

Private Sub txtAjuste_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtDetalle.SetFocus
End If
End Sub

Private Sub txtAjuste_LostFocus()
On Error GoTo vError
  
  txtAjuste.Text = Format(CCur(txtAjuste.Text), "Standard")

vError:

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
    
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    
    txtCedula.Text = Trim(gBusquedas.Resultado)
    
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
    
    If Trim(txtCedula.Text) <> "" Then
        txtNombre.SetFocus
    End If

End If

End Sub

Private Sub txtCedula_LostFocus()

If Trim(txtCedula) <> "" Then
   txtNombre.Text = fxNombre(Trim(txtCedula))
   If Trim(txtNombre.Text) = "" Then
      txtAjusteId.Text = ""
      txtEstado.Text = ""
      MsgBox "Cedula Incorrecta", vbInformation
      txtCedula.SetFocus
   Else
      Call sbCedula_Load
   End If
Else
   txtNombre.Text = ""
   txtAjusteId.Text = ""
   txtEstado.Text = ""
End If

End Sub


Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbListado_Load
End If
End Sub


Private Sub txtLineas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbListado_Load
End If
End Sub

