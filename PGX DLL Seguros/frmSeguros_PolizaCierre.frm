VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_PolizaCierre 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre de Póliza"
   ClientHeight    =   6105
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3015
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   8655
      _Version        =   1441792
      _ExtentX        =   15266
      _ExtentY        =   5318
      _StockProps     =   79
      Caption         =   "Datos de Cierre:"
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
      Begin XtremeSuiteControls.PushButton btnCierre 
         Height          =   615
         Left            =   6720
         TabIndex        =   8
         Top             =   2400
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Cerrar el Contrato"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmSeguros_PolizaCierre.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cboCausa 
         Height          =   330
         Left            =   1560
         TabIndex        =   17
         Top             =   1800
         Width           =   6975
         _Version        =   1441792
         _ExtentX        =   12303
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
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   480
         Width           =   1695
         _Version        =   1441792
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   855
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   7095
         _Version        =   1441792
         _ExtentX        =   12515
         _ExtentY        =   1508
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Causa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   360
         TabIndex        =   6
         Top             =   900
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3480
      TabIndex        =   10
      Top             =   1560
      Width           =   5415
      _Version        =   1441792
      _ExtentX        =   9551
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit txtTipoSeguroDesc 
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Top             =   2280
      Width           =   5415
      _Version        =   1441792
      _ExtentX        =   9551
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipoSeguroCod 
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit txtAseguradoraDesc 
      Height          =   315
      Left            =   3480
      TabIndex        =   14
      Top             =   1920
      Width           =   5415
      _Version        =   1441792
      _ExtentX        =   9551
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAseguradora 
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   435
      Left            =   5880
      TabIndex        =   16
      Top             =   1080
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
      _ExtentY        =   767
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierre de Seguro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   6852
   End
   Begin VB.Label lblPagador 
      BackStyle       =   0  'Transparent
      Caption         =   "Aseguradora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblContrato 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Seguro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Poliza"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10572
   End
End
Attribute VB_Name = "frmSeguros_PolizaCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCierre_Click()
Dim i As Integer

On Error GoTo vError


i = MsgBox("Esta seguro que desea >> Cerrar << esta Póliza", vbYesNo)
If i = vbYes Then
   Call sbPolizaCierra(txtAseguradora.Text, txtPoliza.Text, txtDocumento.Text, txtNotas.Text, cboCausa.ItemData(cboCausa.ListIndex))
End If

UnLoad Me
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
vModulo = 17

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtPoliza.Text = GLOBALES.gTag
txtAseguradora.Text = GLOBALES.gTag2

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "select rtrim(cod_causa) as 'IdX', rtrim(descripcion) as 'ItmX' from SEGUROS_CAUSAS_TIPOS where tipo = 'Cierre' and Activa = 1"
Call sbCbo_Llena_New(cboCausa, strSQL, False, True)


strSQL = "select Pol.*,Ts.Descripcion as 'TipoSeguroDesc', Per.Nombre, isnull(Pol.Estado,'P') as 'Estado'" _
       & ",Ase.nombre as 'AseguradoraDesc' " _
       & " from SEGUROS_REGISTRO Pol inner join SEGUROS_TIPOS_PRODUCTOS Ts on Pol.COD_PRODUCTO = Ts.COD_PRODUCTO and Pol.Cod_Aseguradora = Ts.Cod_Aseguradora" _
       & " inner join SEGUROS_ASEGURADORAS Ase on Pol.Cod_Aseguradora = Ase.Cod_Aseguradora" _
       & " inner join Socios Per on Pol.cedula = Per.cedula" _
       & " where Pol.num_poliza = '" & txtPoliza.Text & "' and Pol.cod_Aseguradora  = '" & txtAseguradora.Text & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  txtPoliza.Text = rs!Num_Poliza
  txtTipoSeguroCod.Text = rs!COD_PRODUCTO
  txtTipoSeguroDesc.Text = rs!TipoSeguroDesc
  
  txtAseguradora.Text = rs!cod_Aseguradora
  txtAseguradoraDesc.Text = rs!AseguradoraDesc

  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbConsulta
End Sub
