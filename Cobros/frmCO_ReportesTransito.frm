VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCO_ReportesTransito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Morosidad en Real/Transito"
   ClientHeight    =   6384
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10668
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6384
   ScaleWidth      =   10668
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4572
      Left            =   120
      TabIndex        =   27
      Top             =   1680
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8064
      _ExtentY        =   8064
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10080
      Top             =   720
   End
   Begin XtremeSuiteControls.GroupBox gbCuotas 
      Height          =   852
      Left            =   4920
      TabIndex        =   22
      Top             =   5160
      Width           =   3372
      _Version        =   1245187
      _ExtentX        =   5948
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Cuotas atrasadas:"
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
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.TextBox txtDesde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtHasta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   23
         Text            =   "80"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   1
         Left            =   720
         TabIndex        =   26
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   0
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   612
      End
   End
   Begin XtremeSuiteControls.ComboBox cboOficina 
      Height          =   312
      Left            =   5880
      TabIndex        =   1
      Top             =   2040
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboGarantia 
      Height          =   312
      Left            =   5880
      TabIndex        =   2
      Top             =   1680
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboCartera 
      Height          =   312
      Left            =   5880
      TabIndex        =   3
      Top             =   2400
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboDestino 
      Height          =   312
      Left            =   5880
      TabIndex        =   4
      Top             =   3600
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboRecurso 
      Height          =   312
      Left            =   5880
      TabIndex        =   5
      Top             =   3960
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   5880
      TabIndex        =   6
      Top             =   4320
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboDeductora 
      Height          =   312
      Left            =   5880
      TabIndex        =   7
      Top             =   4680
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.CheckBox chkLineas 
      Height          =   252
      Left            =   9480
      TabIndex        =   8
      Top             =   2880
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   5880
      TabIndex        =   9
      Top             =   3240
      Width           =   852
      _Version        =   1245187
      _ExtentX        =   1503
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   6720
      TabIndex        =   10
      Top             =   3240
      Width           =   3732
      _Version        =   1245187
      _ExtentX        =   6583
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkRepResumen 
      Height          =   252
      Left            =   9120
      TabIndex        =   20
      Top             =   5160
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Resumen"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   612
      Left            =   8880
      TabIndex        =   21
      Top             =   5520
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
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
      Appearance      =   14
      Picture         =   "frmCO_ReportesTransito.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   5880
      TabIndex        =   28
      Top             =   1320
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
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
   Begin XtremeShortcutBar.ShortcutCaption lblRepGen 
      Height          =   372
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8064
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   7
      Left            =   4920
      TabIndex        =   19
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Garantía"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   5
      Left            =   4920
      TabIndex        =   18
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   4
      Left            =   4920
      TabIndex        =   17
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cartera"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   3
      Left            =   4920
      TabIndex        =   16
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   2
      Left            =   4920
      TabIndex        =   15
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Deductora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   37
      Left            =   4920
      TabIndex        =   14
      Top             =   4680
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   13
      Left            =   4920
      TabIndex        =   13
      Top             =   3600
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recurso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   3960
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   4920
      TabIndex        =   11
      Top             =   4320
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Morosidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Index           =   11
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   4572
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCO_ReportesTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

chkLineas.Value = vbChecked

strSQL = "select rtrim(Garantia) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from crd_garantia_tipos order by descripcion"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, False)

strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by descripcion"
Call sbCbo_Llena_New(cboOficina, strSQL, True, False)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from afi_estados_persona order by descripcion"
Call sbCbo_Llena_New(cboEstado, strSQL, True, False)
'Item Adicional
    cboEstado.AddItem "Ex.Socios"
    cboEstado.ItemData(cboEstado.ListCount - 1) = "X"


'Instituciones
vPaso = True
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)
vPaso = False


 cboCartera.Clear
 
 strSQL = "select rtrim(cod_clasificacion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
        & " from CBR_CLASIFICACION_CARTERA order by cod_clasificacion"
 
 Call sbCbo_Llena_New(cboCartera, strSQL, False, True)
 
 cboCartera.AddItem "(Todas las Carteras)"
 cboCartera.Text = "(Todas las Carteras)"
 
 vPaso = False


Call chkLineas_Click
Call cboInstitucion_Click

Me.MousePointer = vbDefault

End Sub




Private Sub btnReporte_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprime Reportes Generales de Cobro
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Utiliza variables globales
'-------------------------------------------------------------------------------------------
Dim strSQL As String, vSubTitulo As String
Dim i As Byte

Me.MousePointer = vbHourglass

strSQL = ""

If cboEstado.Text <> "TODOS" Then
    Select Case cboEstado.ItemData(cboEstado.ListIndex)
     Case "X"  'Todos
       strSQL = "({SOCIOS.ESTADOACTUAL} = 'A' OR {SOCIOS.ESTADOACTUAL} = 'P')"
     Case Else
       strSQL = "{SOCIOS.ESTADOACTUAL} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
    End Select
End If

If cboOficina.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.COD_OFICINA_R} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

If cboInstitucion.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

If cboDeductora.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{VISTA_MOROSIDAD.COD_DEDUCTORA} = " & cboDeductora.ItemData(cboDeductora.ListIndex)
End If


If cboDestino.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.COD_DESTINO} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
End If

If cboRecurso.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.COD_GRUPO} = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
End If

If cboGarantia.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.GARANTIA} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
End If

vSubTitulo = "Cartera: " & cboCartera.Text _
        & " ¦ Estado: " & cboEstado.Text _
        & " ¦ Garantía: " & cboGarantia.Text _
        & " ¦ Destino: " & cboDestino.Text _
        & " ¦ Recurso: " & cboRecurso.Text _
        & " ¦ Oficina: " & cboOficina.Text _
        & " ¦ Institución: " & cboInstitucion.Text _
        & " ¦ Deductora: " & cboDeductora.Text


If chkLineas.Value = vbChecked Then
  vSubTitulo = vSubTitulo & " ¦ Línea : Todas"
Else
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.CODIGO} = '" & txtCodigo.Text & "'"
  vSubTitulo = vSubTitulo & " ¦ Línea : " & txtCodigo.Text
End If

If cboCartera.Text <> "(Todas las Carteras)" Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{CBR_CLASIFICACION_DETALLE.COD_CLASIFICACION} = '" & cboCartera.ItemData(cboCartera.ListIndex) & "'"
End If


vSubTitulo = Mid(vSubTitulo, 1, 250)

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro"
     
    .Connect = glogon.ConectRPT
     
  Select Case lblRepGen.Tag
   Case "GENDET" 'Reporte General
    
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumen.rpt")
            .Formulas(1) = "Titulo='Informe General de Morosidad: Resumen'"
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetallado.rpt")
            .Formulas(1) = "Titulo='Informe General de Morosidad: Detalle'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL

   
   Case "GENTRA" 'Reporte General + Cuota en Transito
    
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Listado_Transito_Resumen.rpt")
            .Formulas(1) = "Titulo='Informe de Morosidad + Cuota en Transito: Resumen'"
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Listado_Transito_Detallado.rpt")
            .Formulas(1) = "Titulo='Informe de Morosidad + Cuota en Transito: Detalle'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
   
   Case "GENIDL" 'Reporte General por Institucion y Departamento y Linea
    
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenDeptLinea.rpt")
            .Formulas(1) = "Titulo='Informe de Morosidad: CENTRO DE TRABAJO: Resumen'"
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoDeptLinea.rpt")
            .Formulas(1) = "Titulo='Informe de Morosidad: CENTRO DE TRABAJO: Detalle'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
   
   
   Case "GENIDR" 'Reporte General por Institucion y Departamento Resumen
    
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenDept.rpt")
            .Formulas(1) = "Titulo='Informe de Morosidad: CENTRO DE TRABAJO: Resumen'"
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoDept.rpt")
            .Formulas(1) = "Titulo='Informe de Morosidad: CENTRO DE TRABAJO: Detalle'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
   
   Case "GENAGD"    ' General - Detallado Agrupado"
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='Informe de Morosidad: Especial Agrupado: Resumen'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenAgr.rpt")
        Else
            .Formulas(1) = "Titulo='Informe de Morosidad: Especial Agrupado: Detalle'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoAgr.rpt")
        End If
            
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
   
   Case "ESPCON" 'Convenios
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoXConvenios.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & " ¦ Filtro: " & Mid(cboEstado.Text, 4, 30) & "'"
        .Formulas(3) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta & " And {VISTA_MOROSIDAD.CODIGO}='" & txtCodigo & "'"
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
  
   Case "MORCAR" 'Resumen Comparativo
        
        
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ComparativoRsm.rpt")
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Comparativo.rpt")
        End If
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        Select Case Mid(cboEstado.Text, 1, 1)
         Case "T" 'Todos
            .StoredProcParam(0) = "."
            .StoredProcParam(1) = "."
            .StoredProcParam(2) = "."
            .StoredProcParam(3) = "."
         Case "P", "A", "X" 'Opex
            .StoredProcParam(0) = "A"
            .StoredProcParam(1) = "A"
            .StoredProcParam(2) = "P"
            .StoredProcParam(3) = "P"
         Case Else
            .StoredProcParam(0) = cboEstado.ItemData(cboEstado.ListIndex)
            .StoredProcParam(1) = cboEstado.ItemData(cboEstado.ListIndex)
            .StoredProcParam(2) = cboEstado.ItemData(cboEstado.ListIndex)
            .StoredProcParam(3) = cboEstado.ItemData(cboEstado.ListIndex)
        End Select
         
       strSQL = ""
       If cboGarantia.Text <> "TODOS" Then
          strSQL = "{spCBRComparativo;1.garantia} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
       End If
       
       If chkLineas.Value = vbUnchecked Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{spCBRComparativo;1.codigo} = '" & txtCodigo.Text & "'"
       End If
    
       If cboCartera.Text <> "(Todas las Carteras)" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{spCBRComparativo;1.cod_clasificacion} = '" & cboCartera.ItemData(cboCartera.ListIndex) & "'"
       End If
     
     
       .SelectionFormula = strSQL
     
  
   Case "MORGAR" 'Reporte Mora x Garantia
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoXGarantia.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        If Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'" & strSQL
  
  
    Case "MORGAG"  'Mora x Garantía - Agrupado
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoXGarantiaAgr.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        If Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'" & strSQL
  
  
   Case "DETPRV" 'Provincia
        
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='Informe de Morosidad: Provincias: Resumen'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenProv.rpt")
            Me.MousePointer = vbDefault
            i = MsgBox("Desea Mostrar x Líneas de Crédito", vbYesNo)
            If i = vbYes Then
              .Formulas(5) = "fxResumen=0"
            Else
              .Formulas(5) = "fxResumen=1"
            End If
        
        Else
            .Formulas(1) = "Titulo='Informe de Morosidad: Provincias: Detalle'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoProv.rpt")
        End If
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
   

  Case "DETPVT" 'Detalle x Provincia Trabajo
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='Informe de Morosidad: Provincias Laboral: Resumen'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenProvTra.rpt")
            Me.MousePointer = vbDefault
            i = MsgBox("Desea Mostrar x Líneas de Crédito", vbYesNo)
            If i = vbYes Then
              .Formulas(5) = "fxResumen=0"
            Else
              .Formulas(5) = "fxResumen=1"
            End If
        
        Else
            .Formulas(1) = "Titulo='Informe de Morosidad: Provincias Laboral: Detalle'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoProvTra.rpt")
        End If
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
  
    
    Case "DETUND" 'Detalle x Unidad
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD x UNIDAD'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenUnidad.rpt")
            Me.MousePointer = vbDefault
            i = MsgBox("Desea Mostrar x Líneas de Crédito", vbYesNo)
            If i = vbYes Then
              .Formulas(5) = "fxResumen=0"
            Else
              .Formulas(5) = "fxResumen=1"
            End If
        
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD x UNIDAD'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoUnidad.rpt")
        End If
        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
    
    Case "LSTFRM" 'Listado x Formalización

            .Formulas(1) = "Titulo='LISTADO DE MORA. ANALISIS DE FORMALIZACION'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoFormalizacion.rpt")

        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
    Case "LSTCMT" 'Listado x Comité Resolutor
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD x COMITE'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoComiteRsm.rpt")
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD x COMITE'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoComiteDet.rpt")
        End If
        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
     
     Case "ANTLEGAL", "ANTSALDOS", "ANTSALDOSPA", "ANTPROACUM" 'Antiguedad Legal
     
        Select Case lblRepGen.Tag
           Case "ANTLEGAL"      'Antiguedad de Saldos + Mora Legal
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA LEGAL] [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadLegalRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA LEGAL] [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadLegalDet.rpt")
                End If
           
           Case "ANTSALDOS"     'Antiguedad de Saldos (Pura)
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosDet.rpt")
                End If
           
           Case "ANTSALDOSPA"   'Antiguedad de Saldos + (Producto Acumulado)
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosProdAcumRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosProdAcumDet.rpt")
                End If
           
           Case "ANTPROACUM"    'Antiguedad Producto Acumulado
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadProdAcumRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadProdAcumDet.rpt")
                End If
        
        End Select

        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        
        i = MsgBox("Desea Mostrar Operaciones al Día", vbYesNo)
        If i = vbNo Then
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "{vCBRAntiguedadSaldos.CAl Día} = 0"
        End If
        
        .SelectionFormula = strSQL
     
     
     
     Case "ANTFINAN" 'Antiguedad Financiera
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA FINANCIERA] [RESUMEN]'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadFinancieraRsm.rpt")
        Else
            .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA FINANCIERA] [DETALLE]'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadFinancieraDet.rpt")
        End If
        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
    
'        i = MsgBox("Desea Mostrar Operaciones al Día", vbYesNo)
'        If i = vbNo Then
'           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'           strSQL = strSQL & "{vCBRAntiguedadSaldos.CAl Día} = 0"
'        End If
  
  End Select
     
     

    .PrintReport

End With

Me.MousePointer = vbDefault
End Sub

Private Sub cboInstitucion_Click()
Dim strSQL As String

If vPaso Then Exit Sub

cboDeductora.Clear

If cboInstitucion.Text = "TODOS" Then
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboDeductora, strSQL, True, True)
Else
    strSQL = "exec spAFI_Institucion_Vinculadas " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",3"
    Call sbCbo_Llena_New(cboDeductora, strSQL, True, True)
End If
End Sub

Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_grupos order by descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select rtrim(cod_destino) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_destinos order by descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "' order by R.descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select (R.cod_destino) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "' order by R.Descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)

End If


End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then txtDescripcion.Text = fxDescribeCodigo(Trim(txtCodigo))
 Call chkLineas_Click
End Sub

Private Sub Form_Load()


Dim strSQL As String, rs As New ADODB.Recordset


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Informe:", 4040
End With
 
 
 
 With lsw.ListItems
   .Clear
   .Add , "GENDET", "Listado General"
   .Add , "GENTRA", "Listado General + Cuota en Transito"
   .Add , "GENAGD", "Listado General Agrupado"
   
   .Add , "GENIDR", "Listado por Institución/Departamento"
   .Add , "GENIDL", "Listado por Institución/Departamento/Línea"
   
   .Add , "ESPCON", "Especial Convenios"
   .Add , "MORCAR", "Comparativo - Resumen"
   
   'Garantia
   .Add , "MORGAR", "Mora x Garantía"
   .Add , "MORGAG", "Mora x Garantía Agrupado"
 
   'Antiguedad Saldos
   .Add , "ANTSALDOS", "Antiguedad de Saldos"
   .Add , "ANTSALDOSPA", "Antiguedad de Saldos + (Producto Acumulado)"
   .Add , "ANTPROACUM", "Antiguedad de Producto Acumulado"
 
 
   'Antiguedad Saldos (Efectos Moratorios y Cobrabilidad)
   .Add , "ANTLEGAL", "Antiguedad de Saldos (Legal)"
   .Add , "ANTFINAN", "Antiguedad de Saldos (Financiera)"
 
   'Nuevos
   .Add , "DETPRV", "Listado x Provincia"
   .Add , "DETPVT", "Listado x Provincia Trabajo"
   .Add , "DETUND", "Listado x Unidad"
   .Add , "LSTFRM", "Listado x Formalización"
   .Add , "LSTCMT", "Listado x Comité Resolutor"
 
 
 
 End With
 

 
End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

lblRepGen.Caption = Item.Text
lblRepGen.Tag = Item.Key

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub

