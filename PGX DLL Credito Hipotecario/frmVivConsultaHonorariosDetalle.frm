VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmVivConsultaHonorariosDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle de honorarios registrados"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lvwDetalle 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   10575
      _Version        =   1310723
      _ExtentX        =   18653
      _ExtentY        =   8916
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
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
      _Version        =   1310723
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
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
      _Version        =   1310723
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
      TabIndex        =   5
      Top             =   1320
      Width           =   6255
      _Version        =   1310723
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.Label lblTitulo 
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   8775
      _Version        =   1310723
      _ExtentX        =   15478
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Utilice este formulario para conusultar el detalle de honorarios registrados para la garantía seleccionada"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Nombre"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identificación"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operacion"
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
   Begin VB.Label lblTotalMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      TabIndex        =   0
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   15
      Left            =   5640
      TabIndex        =   1
      Top             =   5160
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmVivConsultaHonorariosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_IdGarantia As Long

Public Sub sbHonorariosLoad()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vItem As ListViewItem
Dim vMontoTotal As Double

On Error GoTo vError

lvwDetalle.ColumnHeaders.Clear
lvwDetalle.ListItems.Clear

lvwDetalle.ColumnHeaders.Add , , "Linea", 0
lvwDetalle.ColumnHeaders.Add , , "Código", 1300
lvwDetalle.ColumnHeaders.Add , , "Descripción", 3000
lvwDetalle.ColumnHeaders.Add , , "Monto", 1500, 1
lvwDetalle.ColumnHeaders.Add , , "Profesional", 3000
lvwDetalle.ColumnHeaders.Add , , "Usuario", 2000
lvwDetalle.ColumnHeaders.Add , , "Fecha Registro", 2000


strSQL = "SELECT ViviendaGarantia.NumeroOperacion,ViviendaContactos.Nombre as Contacto, HonorariosDT.Tipo, HonorariosDT.IdGarantia," & _
       " HonorariosDT.Linea,HonorariosDT.Codigo, TD.Descripcion," & _
       " HonorariosDT.Monto, HonorariosDT.Usuario, HonorariosDT.Fecha," & _
       " SOCIOS.NOMBRE AS NombreSocio,SOCIOS.CEDULA AS CedulaSocio" & _
       " FROM ViviendaDesembolsosPendientesDT AS HonorariosDT INNER JOIN" & _
             " ViviendaTiposDesembolsos AS TD ON HonorariosDT.Codigo = TD.Codigo INNER JOIN" & _
             " ViviendaGarantia ON HonorariosDT.IdGarantia = ViviendaGarantia.IdGarantia INNER JOIN" & _
             " ViviendaContactos ON HonorariosDT.IdContacto = ViviendaContactos.IdContacto INNER JOIN" & _
             " REG_CREDITOS ON ViviendaGarantia.NumeroOperacion = REG_CREDITOS.ID_SOLICITUD INNER JOIN" & _
             " SOCIOS ON REG_CREDITOS.CEDULA = SOCIOS.CEDULA" & _
        " Where ViviendaGarantia.IdGarantia = " & m_IdGarantia & " order by HonorariosDT.Fecha"

Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
    txtOperacion.Text = rs!NumeroOperacion
    txtCedula.Text = rs!CedulaSocio
    txtNombre.Text = rs!NombreSocio
 
    Do While Not rs.EOF
      
        Set vItem = lvwDetalle.ListItems.Add(, , rs!Linea)
            vItem.SubItems(1) = Trim(rs!codigo)
            vItem.SubItems(2) = Trim(rs!Descripcion)
            vItem.SubItems(3) = Format(rs!Monto, "Standard")
            vItem.SubItems(4) = Trim(rs!Contacto)
            vItem.SubItems(5) = rs!Usuario
            vItem.SubItems(6) = Format(rs!fecha, "dd-mm-yyyy hh:mm AMPM")
            vMontoTotal = vMontoTotal + rs!Monto
           rs.MoveNext
    Loop
    rs.Close
End If

lblTotalMonto.Caption = Format(vMontoTotal, "Standard")
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

m_IdGarantia = GLOBALES.gTag

Call sbHonorariosLoad

End Sub

