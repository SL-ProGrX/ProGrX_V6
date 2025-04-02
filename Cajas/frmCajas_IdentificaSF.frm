VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCajas_IdentificaSF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Identifica Depósitos en Tramite"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox fraIdentifica 
      Height          =   5055
      Left            =   1680
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   8415
      _Version        =   1441793
      _ExtentX        =   14843
      _ExtentY        =   8916
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtId_NSolicitud 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_NumDocId 
         Height          =   315
         Left            =   5880
         TabIndex        =   9
         Top             =   600
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Fecha 
         Height          =   315
         Left            =   5880
         TabIndex        =   10
         Top             =   2280
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Cedula 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   3360
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Nombre 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   3720
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
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
         BackColor       =   16777215
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Banco 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
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
         BackColor       =   16777215
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Descripcion 
         Height          =   795
         Left            =   1560
         TabIndex        =   14
         Top             =   1440
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
         _ExtentY        =   1402
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
         BackColor       =   16777215
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Monto 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   2280
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnIdentifica 
         Height          =   495
         Index           =   0
         Left            =   5640
         TabIndex        =   16
         Top             =   4320
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCajas_IdentificaSF.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnIdentifica 
         Height          =   495
         Index           =   1
         Left            =   6840
         TabIndex        =   17
         Top             =   4320
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCajas_IdentificaSF.frx":0727
         ImageAlignment  =   4
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Identificación del Propietario del Depósito"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   600
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Solicitud:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuenta:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descripción:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   23
         Top             =   600
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Documento:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Top             =   2280
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   20
         Top             =   2880
         Width           =   3615
         _Version        =   1441793
         _ExtentX        =   6376
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación del Caso:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   19
         Top             =   3360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación:"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre:"
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpId_Inicio 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpId_Corte 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.FlatEdit txtId_NumDoc 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin FPSpreadADO.fpSpread vGridId 
      Height          =   6375
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   11655
      _Version        =   524288
      _ExtentX        =   20558
      _ExtentY        =   11245
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
      MaxCols         =   495
      SpreadDesigner  =   "frmCajas_IdentificaSF.frx":0E3D
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   1320
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCajas_IdentificaSF.frx":19F5
      ImageAlignment  =   4
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Identifica Depósitos en Tramite"
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
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   29
      Top             =   240
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta .:"
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
      Index           =   16
      Left            =   3960
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Doc.:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha .:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmCajas_IdentificaSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnAccion_Click()

vPaso = True
    Call sbConsultaDPTramite
vPaso = False

End Sub

Private Sub btnIdentifica_Click(Index As Integer)

On Error GoTo vError

If Index = 1 Then
   fraIdentifica.Visible = False
   Exit Sub
End If

If txtId_Nombre.Text = "" Then
    MsgBox "No se ha especificado ningún Id de Cliente válido", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_Identifica_TES_Depositos " & txtId_Banco.Tag & ",'" & txtId_NumDocId.Text & "','" & txtId_Cedula.Text _
       & "','" & txtId_Nombre.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

vGridId.DeleteRows txtId_Cedula.Tag, 1
vGridId.MaxRows = vGridId.MaxRows - 1

fraIdentifica.Visible = False

MsgBox "caso identificado correctamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
vModulo = 5

'Carga las cuentas bancarias asiganadas a la forma de pago
vPaso = True


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboBanco.Clear

strSQL = "exec spCajas_DepositosCuentasBancariasAut 'DP'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboBanco.AddItem Trim(rs!Cta) & " - " & Trim(rs!DESCRIPCION & "")
 cboBanco.ItemData(cboBanco.ListCount - 1) = CStr(rs!Id_Banco)
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
    cboBanco.Text = Trim(rs!Cta) & " - " & Trim(rs!DESCRIPCION & "")
End If
rs.Close


vPaso = True
    
vGridId.MaxCols = 10
vGridId.MaxRows = 0

vPaso = False



dtpId_Corte.Value = fxFechaServidor
dtpId_Inicio.Value = DateAdd("d", -10, dtpId_Corte.Value)


fraIdentifica.Visible = False

Call RefrescaTags(Me)
Call Formularios(Me)

End Sub


Private Sub sbConsultaDPTramite()
Dim i As Long

On Error GoTo vError

If cboBanco.ListCount = 0 Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select Tra.*, Bn.Descripcion as 'BancoDesc'" _
        & " From TES_DEPOSITOS_TRAMITE Tra inner join Tes_Bancos Bn on Tra.id_banco = Bn.id_Banco" _
        & " Where Tra.ID_REQUERIDA = 1 And Tra.IDENTIFICADO = 0" _
        & " and  fecha between '" & Format(dtpId_Inicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpId_Corte.Value, "yyyy/mm/dd") & " 23:59:59'"


If Len(Trim(txtId_NumDoc.Text)) > 0 Then
    strSQL = strSQL & " and Tra.Documento like '%" & txtId_NumDoc.Text & "%'"
End If

strSQL = strSQL & " and Tra.Id_Banco = " & cboBanco.ItemData(cboBanco.ListIndex)

Call OpenRecordSet(rs, strSQL)

vGridId.MaxRows = 0


  Do While Not rs.EOF
    vGridId.MaxRows = vGridId.MaxRows + 1
    vGridId.Row = vGridId.MaxRows
         
    vGridId.col = 1

    For i = 2 To vGridId.MaxCols
      vGridId.col = i
      Select Case i
         Case 2 'Id
            vGridId.Text = CStr(rs!NSolicitud)
         Case 3 'Cuenta
            vGridId.Text = rs!BancoDesc & ""
            vGridId.CellTag = rs!Id_Banco
         Case 4 ' Tipo
            vGridId.Text = "DP"
         Case 5 'Num Documento
            vGridId.Text = rs!Documento
         Case 6 'Fecha del Documento
            vGridId.Text = Format(rs!fecha, "dd/mm/yyyy")
         Case 7 'Monto
            vGridId.Text = Format(rs!Monto, "Standard")
         Case 8 'Descripcion
            vGridId.Text = rs!DESCRIPCION
         Case 9 'Registro Fecha
            vGridId.Text = rs!Registro_Fecha & ""
         Case 10 'Registro Usuario
            vGridId.Text = rs!Registro_Usuario & ""
      
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGridId_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

txtId_Cedula.Text = GLOBALES.gTag
txtId_Nombre.Text = GLOBALES.gTag2


With vGridId
    .Row = Row
    .col = 2
    txtId_NSolicitud = .Text
    .col = 3
    txtId_Banco.Text = .Text
    txtId_Banco.Tag = .CellTag
    
    .col = 5
    txtId_NumDocId.Text = .Text
    .col = 6
    txtId_Fecha.Text = .Text
    .col = 7
    txtId_Monto.Text = .Text
    .col = 8
    txtId_Descripcion.Text = .Text
End With

fraIdentifica.Visible = True

End Sub

