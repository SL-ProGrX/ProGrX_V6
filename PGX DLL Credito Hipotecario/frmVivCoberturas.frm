VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmVivCoberturas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobertura de las Garantías Hipotecarias"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1575
      Left            =   4800
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9551
      _ExtentY        =   2778
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
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   19
      Top             =   2280
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cobertura General"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
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
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   2280
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDisponible 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   2640
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHipotecaExterna 
      Height          =   315
      Left            =   2280
      TabIndex        =   15
      Top             =   3240
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHipotecaInterna 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   3600
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtHipotecaLiberada 
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   3960
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCobertura 
      Height          =   315
      Left            =   2280
      TabIndex        =   18
      Top             =   4560
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   2280
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cobertura Individual"
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
   End
   Begin XtremeSuiteControls.Label lblX 
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9551
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Seleccione una Finca/Hipoteca para calcular..."
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "(+/-) Cobertura"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "( + ) Hipotecas Liberadas"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "( - ) Hipotecas Internas"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "( - ) Hipotecas Externas"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disponible"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Total Avalúo"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cobertura de las Garantías"
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
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   9735
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operación"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identificación"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmVivCoberturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
optX.Item(0).Value = True

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

If gOperacion > 0 Then
   Call sbCarga(gOperacion)
End If

End Sub


Public Sub sbCarga(pOperacion As Long, Optional pInicial As Boolean = True)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

txtTotal.Text = ""
txtDisponible.Text = ""
txtHipotecaExterna.Text = ""
txtHipotecaInterna.Text = ""
txtHipotecaLiberada.Text = ""
txtCobertura.Text = ""


If pInicial Then
   strSQL = "select R.id_solicitud,R.cedula,S.nombre from reg_creditos R inner join Socios S" _
          & " on R.cedula = S.cedula where R.id_solicitud = " & pOperacion
   
   Call OpenRecordSet(rs, strSQL)
   If Not rs.BOF And Not rs.EOF Then
     txtOperacion.Text = rs!Id_solicitud
     txtCedula.Text = rs!cedula
     txtNombre.Text = rs!Nombre
   End If
   rs.Close
   
   'Carga Avaluos
    With lsw.ColumnHeaders
        .Clear
        .Add , , "No. Finca", 1800
        .Add , , "Monto Avalúo", 1800, vbRightJustify
    End With
   
    lsw.ListItems.Clear
    strSQL = "select NumeroFinca,isnull(valorTerreno,0) + isnull(ValorConstruccion,0) as Avaluo" _
           & " from viviendaGarantia where NumeroOperacion = " & pOperacion
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!NumeroFinca)
         itmX.SubItems(1) = Format(rs!Avaluo, "Standard")
     rs.MoveNext
    Loop
    rs.Close
    lblX.Tag = ""
End If


Select Case True
 Case optX.Item(0).Value 'General
        strSQL = "exec spCRDViviendaCoberturaTotal " & pOperacion
 Case optX.Item(1).Value 'Individualizada
        strSQL = "exec spCRDViviendaCoberturaIndividual " & pOperacion & ",'" & lblX.Tag & "'"
End Select

Call OpenRecordSet(rs, strSQL)
    txtTotal.Text = Format(rs!Avaluo, "Standard")
    txtDisponible.Text = Format(rs!disponible, "Standard")
    txtHipotecaExterna.Text = Format(rs!HipExterna, "Standard")
    txtHipotecaInterna.Text = Format(rs!HipInterna, "Standard")
    txtHipotecaLiberada.Text = Format(rs!HipLibera, "Standard")
    txtCobertura.Text = Format(rs!Cobertura, "Standard")
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If lsw.ListItems.Count <= 0 Then Exit Sub

lblX.Tag = Item.Text

Call sbCarga(gOperacion, False)

End Sub

Private Sub optX_Click(Index As Integer)

Select Case Index
 Case 0 'General
    lblX.Visible = False
    Call sbCarga(gOperacion, False)
    
 Case 1 'Individual
    lblX.Visible = True
    lblX.Tag = ""
    txtTotal.Text = "0.00"
    txtDisponible.Text = "0.00"
    txtHipotecaExterna.Text = "0.00"
    txtHipotecaInterna.Text = "0.00"
    txtHipotecaLiberada.Text = "0.00"
    txtCobertura.Text = "0.00"

End Select

lsw.Visible = lblX.Visible

End Sub
