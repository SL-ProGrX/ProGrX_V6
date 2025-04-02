VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_ExcedentesCapInd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mant. Capitalizaciones Individuales"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11310
   Icon            =   "ExcedentesCapIndividual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3975
      Left            =   120
      TabIndex        =   16
      Top             =   2880
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
   Begin XtremeSuiteControls.ScrollBar ScrollBarX 
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   1800
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   0
      _StockProps     =   64
      Min             =   2022
      Max             =   2100
      UseVisualStyle  =   0   'False
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9960
      Top             =   720
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   556
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
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
   Begin XtremeSuiteControls.FlatEdit txtId 
      Height          =   555
      Left            =   1680
      TabIndex        =   2
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1680
      TabIndex        =   5
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
      Picture         =   "ExcedentesCapIndividual.frx":030A
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   6
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
      Picture         =   "ExcedentesCapIndividual.frx":093C
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   3120
      TabIndex        =   7
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
      Picture         =   "ExcedentesCapIndividual.frx":0F37
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3720
      TabIndex        =   8
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
      Picture         =   "ExcedentesCapIndividual.frx":14DB
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   4080
      TabIndex        =   9
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
      Picture         =   "ExcedentesCapIndividual.frx":1C0C
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   4560
      TabIndex        =   10
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
      Picture         =   "ExcedentesCapIndividual.frx":230C
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   315
      Left            =   3000
      TabIndex        =   15
      Top             =   1800
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   556
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
      Text            =   "%"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtVencimiento 
      Height          =   315
      Left            =   5040
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   315
      Left            =   3960
      TabIndex        =   17
      Top             =   2520
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   255
      Index           =   0
      Left            =   10800
      TabIndex        =   18
      ToolTipText     =   "Exportar Listado a Excel"
      Top             =   2520
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   444
      _ExtentY        =   444
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "ExcedentesCapIndividual.frx":2A13
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   19
      Top             =   2280
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
      TabIndex        =   20
      Top             =   6960
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   556
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
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   315
      Left            =   10320
      TabIndex        =   21
      ToolTipText     =   "Cantidad de Casos a Mostrar"
      Top             =   6960
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   556
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
      Text            =   "100"
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Listado de Capitalizaciones                 Filtros:"
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
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Casos:"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   22
      Top             =   6960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Lineas:"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Vencimiento"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Porcentaje"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cédula"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   " Id"
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
End
Attribute VB_Name = "frmAH_ExcedentesCapInd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vEdita As Boolean, vPaso As Boolean, vFecha As Date


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
      If txtId.Text = "" Then
        MsgBox "Seleccione una Mejora o Retiro de la lista del activo para modificacion...", vbInformation
      Else
        vEdita = True
        txtPorcentaje.SetFocus
        Call sbBarra_Accion("Editar")
      End If
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If txtId.Text = "" Then
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

If Trim(txtCedula) = "" Then
  vMensaje = vMensaje & "- El No. de Cédula no es válido!" & vbCrLf
End If

If Not IsNumeric(txtPorcentaje.Text) Then
  vMensaje = vMensaje & "- El Porcentaje no es valido!" & vbCrLf
End If

If IsNumeric(txtPorcentaje.Text) Then
    If CCur(txtPorcentaje.Text) < 0 Or CCur(txtPorcentaje.Text) > 100 Then
      vMensaje = vMensaje & "- El Porcentaje no es valido!" & vbCrLf
    End If
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

If Trim(txtId.Text) = "" Then Exit Sub

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
     strSQL = "Delete from EXC_CAP_INDIVIDUAL where EXC_CAP_IND = " & Trim(txtId.Text)
     Call ConectionExecute(strSQL)
     
     Call Bitacora("Borra", "Excedente Cap.Extra Id: " & txtId.Text & ", Cedula: " & txtCedula.Text & ", Porcentaje : " & txtPorcentaje.Text)
     
     Call sbListado_Load
     
     Call sbBarra_Accion("nuevo")
     Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbLimpia()

txtId.Text = ""
txtCedula.Text = ""
txtNombre.Text = ""
txtPorcentaje.Text = "0"
txtVencimiento.Text = Year(vFecha)

ScrollBarX.Value = Year(vFecha)
     
End Sub

Private Sub sbListado_Load()

On Error GoTo vError

Dim pFiltro As String

pFiltro = fxSysCleanTxtInject(txtFiltro.Text)
txtLineas.Text = fxSysCleanTxtInject(txtLineas.Text)

If Not IsNumeric(txtLineas.Text) Then
    txtLineas.Text = "100"
End If

strSQL = "Select Top " & txtLineas.Text & " A.*, S.Nombre" _
       & " from EXC_CAP_INDIVIDUAL A inner join Socios S on A.cedula = S.cedula" _
       & " where S.cedula like '%" & pFiltro & "%'" _
       & " or S.Nombre like '%" & pFiltro & "%'"
lsw.ListItems.Clear

vPaso = True

Call OpenRecordSet(rs, strSQL)
 
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!EXC_CAP_IND)
       itmX.SubItems(1) = Trim(rs!Cedula)
       itmX.SubItems(2) = rs!Nombre
       itmX.SubItems(3) = Format(rs!Porcentaje, "Standard")
       itmX.SubItems(4) = Year(Trim(rs!Vencimiento))
   rs.MoveNext
Loop
rs.Close

vPaso = False

txtCasos.Text = Format(lsw.ListItems.Count, "###,##0")

Call sbLimpia

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub sbCedula_Load()

strSQL = "Select Top 1 A.*,S.Nombre " _
       & " from EXC_CAP_INDIVIDUAL A inner join Socios S on A.cedula = S.cedula" _
       & " where A.cedula = '" & Trim(txtCedula.Text) _
       & "' order by EXC_CAP_IND desc"

Call OpenRecordSet(rs, strSQL)
 
 
If Not rs.EOF And Not rs.BOF Then
   vEdita = True
   
   txtId.Text = rs!EXC_CAP_IND
   
   txtPorcentaje.Text = Format(rs!Porcentaje, "Standard")
   txtVencimiento.Text = CStr(rs!Vencimiento)
   
   Call sbBarra_Accion("EDITAR")

Else
   txtId.Text = ""
   txtNombre.Text = fxNombre(txtCedula.Text)
End If

rs.Close

End Sub

Private Sub sbGuardar()
Dim pId As Long

On Error GoTo vError



txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtVencimiento.Text = fxSysCleanTxtInject(txtVencimiento.Text)

If Not vEdita Then
    
   strSQL = "select isnull(max(EXC_CAP_IND),0) + 1 as 'Id' from EXC_CAP_INDIVIDUAL"
   Call OpenRecordSet(rs, strSQL)
     txtId.Text = rs!Id
   rs.Close
    
   strSQL = "Insert into EXC_CAP_INDIVIDUAL(EXC_CAP_IND,Cedula,PORCENTAJE,VENCIMIENTO, REGISTRO_FECHA, REGISTRO_USUARIO)" _
          & "  Values(" & txtId.Text & ", '" & Trim(txtCedula) & "'," & CCur(txtPorcentaje.Text) & ", '" & Trim(txtVencimiento.Text) _
          & "-12-31', dbo.MyGetdate(), '" & glogon.Usuario & "')"

     Call Bitacora("Registra", "Excedente Cap.Extra Id: " & txtId.Text & ", Cedula: " & txtCedula.Text & ", Porcentaje: " & txtPorcentaje.Text & ", Vence: " & txtVencimiento.Text)

Else
    strSQL = "Update EXC_CAP_INDIVIDUAL set Cedula='" & Trim(txtCedula) & "', Porcentaje = " & CCur(txtPorcentaje.Text) _
           & ", Vencimiento = '" & Trim(txtVencimiento) & "-12-31', REGISTRO_USUARIO = '" & glogon.Usuario & "', Registro_FEcha = dbo.mygetDate()" _
           & " Where EXC_CAP_IND = " & Trim(txtId.Text)
   
     Call Bitacora("Modifica", "Excedente Cap.Extra Id: " & txtId.Text & ", Cedula: " & txtCedula.Text & ", Porcentaje: " & txtPorcentaje.Text & ", Vence: " & txtVencimiento.Text)
   
End If
 
 Call ConectionExecute(strSQL)
 
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
    .Add , , "Porcentaje", 1500, vbRightJustify
    .Add , , "Vencimiento", 1500, vbCenter

End With

vFecha = fxFechaServidor

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
   
txtId.Text = Item.Text

txtCedula.Text = Item.SubItems(1)
txtNombre.Text = Item.SubItems(2)
   
txtPorcentaje.Text = Item.SubItems(3)
txtVencimiento.Text = Item.SubItems(4)

Call sbBarra_Accion("ACTIVO")

txtPorcentaje.SetFocus

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub ScrollBarX_Change()
txtVencimiento.Text = ScrollBarX.Value
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbListado_Load

End Sub

Private Sub txtPorcentaje_GotFocus()
On Error GoTo vError
  txtPorcentaje.Text = CCur(txtPorcentaje.Text)
vError:
End Sub

Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtVencimiento.SetFocus
End If
End Sub

Private Sub txtPorcentaje_LostFocus()
On Error GoTo vError
  
  txtPorcentaje.Text = Format(CCur(txtPorcentaje.Text), "Standard")

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
      txtId.Text = ""
      MsgBox "Cedula Incorrecta", vbInformation
      txtCedula.SetFocus
   Else
      Call sbCedula_Load
   End If
Else
   txtNombre.Text = ""
   txtId.Text = ""
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

