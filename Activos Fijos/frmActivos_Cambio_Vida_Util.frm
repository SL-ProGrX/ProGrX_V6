VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmActivos_Cambio_Vida_Util 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Vida Util"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnNuevo 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   1440
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Nuevo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   10815
      _Version        =   1441793
      _ExtentX        =   19076
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   615
         Left            =   9240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmActivos_Cambio_Vida_Util.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   762
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   8175
      _Version        =   1441793
      _ExtentX        =   14420
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ComboBox cboVU 
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.FlatEdit txtVU 
      Height          =   330
      Left            =   2280
      TabIndex        =   10
      Top             =   3600
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1291
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1035
      Left            =   4440
      TabIndex        =   13
      Top             =   3600
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10398
      _ExtentY        =   1826
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
   Begin XtremeSuiteControls.FlatEdit txtActual 
      Height          =   330
      Left            =   2280
      TabIndex        =   16
      Top             =   2520
      Width           =   8175
      _Version        =   1441793
      _ExtentX        =   14420
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   315
      Index           =   2
      Left            =   600
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Dato actual:"
      ForeColor       =   0
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Notas"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Vida útil:"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   11
      Top             =   4080
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Método Depreciación:"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   6615
      _Version        =   1441793
      _ExtentX        =   11668
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cambio de Vida Util"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11055
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Descripción"
      ForeColor       =   0
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   435
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "No. Placa"
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmActivos_Cambio_Vida_Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnNuevo_Click()
Call sbLimpia
End Sub

Private Sub Form_Load()
 vModulo = 36

 Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
 
 Call sbActivos_MetodosDepreciacion(cbo)
 
' vPaso = True
'  strSQL = "select rtrim(tipo_activo) as 'Idx', rtrim(descripcion) as 'ItmX'" _
'       & " from Activos_tipo_activo order by tipo_activo"
'  Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
'
'  strSQL = "select rtrim(COD_LOCALIZA) as 'Idx', rtrim(descripcion) as 'ItmX'" _
'       & " from ACTIVOS_LOCALIZACIONES Where Activa = 1 order by descripcion"
'  Call sbCbo_Llena_New(cboLocaliza, strSQL, False, True)
'
' vPaso = False
 vPaso = True

   cboVU.AddItem "Años"
   cboVU.AddItem "Meses"
   cboVU.Text = "Años"
   
 vPaso = False
 

Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Sub sbLimpia()
  txtCodigo.Text = ""
  txtDescripcion.Text = ""
  txtNotas.Text = ""
  txtVU.Text = "1"
    
  txtCodigo.SetFocus
End Sub


Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "num_placa"
  gBusquedas.Orden = "num_placa"
  
  gBusquedas.Col1Name = "Id Placa"
  gBusquedas.Col2Name = "Id Alterna"
  gBusquedas.Col3Name = "Nombre"
  
  gBusquedas.Consulta = "select num_placa, Placa_Alterna, Nombre from Activos_Principal"
  
  gBusquedas.Filtro = " and Estado = 'A'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub





Private Sub sbConsulta(pNumPlaca As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select A.Num_Placa, A.Nombre, A.vida_util_en, A.vida_util, A.met_depreciacion, A.tipo_activo" _
       & ",T.descripcion as 'Tipo_Activo_Desc'" _
       & " from Activos_Principal A" _
       & " inner join Activos_tipo_activo T on A.tipo_activo = T.tipo_activo" _
       & " where A.num_placa = '" & pNumPlaca & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  vPaso = False
    
  txtCodigo = rs!num_placa
  txtDescripcion = rs!Nombre
     
  txtActual.Text = fxActivos_MetodoDepreciacion(rs!met_depreciacion) & ", VU: " _
        & rs!Vida_Util & " " & IIf(rs!vida_util_en = "A", "años", "meses")
   
  Call sbCboAsignaDato(cbo, fxActivos_MetodoDepreciacion(rs!met_depreciacion), True, rs!met_depreciacion)
  
  txtVU.Text = rs!Vida_Util
  
  If rs!vida_util_en = "A" Then
    cboVU.Text = "Años"
  Else
    cboVU.Text = "Meses"
  End If
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

