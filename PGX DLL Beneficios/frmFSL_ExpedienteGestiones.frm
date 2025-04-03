VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFSL_ExpedienteGestiones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestiones"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11415
      _Version        =   1441793
      _ExtentX        =   20135
      _ExtentY        =   10186
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Histórico"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vgGestiones"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "txtNotas"
      Item(1).Control(1)=   "cboGestiones"
      Item(1).Control(2)=   "cmdAplicar"
      Item(1).Control(3)=   "label5(0)"
      Item(1).Control(4)=   "label5(1)"
      Begin FPSpreadADO.fpSpread vgGestiones 
         Height          =   5295
         Left            =   -69880
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   9340
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
         MaxCols         =   4
         SpreadDesigner  =   "frmFSL_ExpedienteGestiones.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   615
         Left            =   7920
         TabIndex        =   9
         Top             =   4560
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
         Picture         =   "frmFSL_ExpedienteGestiones.frx":0652
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   3090
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   8055
         _Version        =   1441793
         _ExtentX        =   14208
         _ExtentY        =   5450
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
      Begin XtremeSuiteControls.ComboBox cboGestiones 
         Height          =   330
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gestión"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   873
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   495
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   873
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   6975
      _Version        =   1441793
      _ExtentX        =   12303
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Expediente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   330
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   582
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
      Alignment       =   1
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFSL_ExpedienteGestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

strSQL = "exec spFSL_GestionRegistra " & txtExpediente.Text & ",'" & cboGestiones.ItemData(cboGestiones.ListIndex) & "','" _
        & txtNotas.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Gestión registrada satisfactoriamente!", vbInformation
Call sbInicializa

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

vModulo = 7

txtExpediente.Text = GLOBALES.gTag

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset


Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

strSQL = "select rtrim(cod_gestion) as 'IdX',  rtrim(DESCRIPCION) as 'ItmX' from FSL_TIPOS_GESTIONES WHERE ACTIVA = 1"
Call sbCbo_Llena_New(cboGestiones, strSQL, False, True)



strSQL = "select Soc.Nombre,Ex.*" _
       & " from FSL_Expedientes Ex inner join Socios Soc on Ex.cedula = Soc.Cedula" _
       & " Where Ex.Cod_Expediente = " & txtExpediente.Text & " order by registro_fecha desc"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Or Not rs.BOF Then
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  
  txtNotas.Text = ""
  
 txtEstado.Tag = rs!Estado
  Select Case rs!Estado
   Case "P" 'Pendiente
        txtEstado.Text = "PENDIENTE"
    Case "A" 'Aprobado
        txtEstado.Text = "APROBADO"
    Case "R" 'Rechazado
        txtEstado.Text = "RECHAZADO"
    Case "X" 'Aplicado
        txtEstado.Text = "APLICADO"
  End Select

End If
rs.Close


'Histórico
strSQL = "select Tg.Descripcion, Eg.*" _
       & " from FSL_EXPEDIENTE_GESTIONES Eg inner join FSL_TIPOS_GESTIONES Tg on Eg.COD_GESTION = Tg.COD_GESTION" _
       & " Where Eg.cod_Expediente = " & txtExpediente.Text
Call OpenRecordSet(rs, strSQL)
vgGestiones.MaxRows = 0
Do While Not rs.EOF
  vgGestiones.MaxRows = vgGestiones.MaxRows + 1
  vgGestiones.Row = vgGestiones.MaxRows
  
  vgGestiones.Col = 1
  vgGestiones.Text = rs!Descripcion
  vgGestiones.TextTip = TextTipFixed
  vgGestiones.TextTipDelay = 1000

  vgGestiones.CellNote = "Fecha : " & rs!registro_Fecha & vbCrLf & "Usuario : " & rs!Registro_Usuario
  vgGestiones.CellTag = CStr(rs!Linea)
    
  vgGestiones.Col = 2
  vgGestiones.Text = rs!notas
      
  vgGestiones.Col = 3
  vgGestiones.Text = rs!registro_Fecha
      
  vgGestiones.Col = 4
  vgGestiones.Text = rs!Registro_Usuario
 
  vgGestiones.RowHeight(vgGestiones.Row) = vgGestiones.MaxTextRowHeight(vgGestiones.Row)
  
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

