VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDExclusionMulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Exclusiones de Multa"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   13455
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5895
      Left            =   0
      TabIndex        =   9
      Top             =   2640
      Width           =   13455
      _Version        =   1572864
      _ExtentX        =   23733
      _ExtentY        =   10398
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
      Sorted          =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   495
      Left            =   10200
      TabIndex        =   31
      Top             =   0
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Todas "
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
      Value           =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   13455
      _Version        =   1572864
      _ExtentX        =   23733
      _ExtentY        =   2355
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
      Item(0).Caption =   "Filtros"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "gbFiltros"
      Item(1).Caption =   "Exclusión"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gbExcluye"
      Begin XtremeSuiteControls.GroupBox gbFiltros 
         Height          =   975
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   13575
         _Version        =   1572864
         _ExtentX        =   23945
         _ExtentY        =   1720
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtC_Cedula 
            Height          =   330
            Left            =   2640
            TabIndex        =   12
            Top             =   480
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtC_Contrato 
            Height          =   330
            Left            =   1440
            TabIndex        =   14
            Top             =   480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
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
         Begin XtremeSuiteControls.FlatEdit txtC_Nombre 
            Height          =   330
            Left            =   4680
            TabIndex        =   13
            Top             =   480
            Width           =   5055
            _Version        =   1572864
            _ExtentX        =   8916
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
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   495
            Left            =   10200
            TabIndex        =   28
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
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
            Picture         =   "frmFNDExclusionMulta.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   495
            Left            =   11520
            TabIndex        =   29
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
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
            Picture         =   "frmFNDExclusionMulta.frx":0700
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBarX 
            Height          =   135
            Left            =   10200
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   238
            _StockProps     =   93
            BackColor       =   -2147483633
            Scrolling       =   1
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   2640
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   7
            Left            =   4680
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   1440
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox gbExcluye 
         Height          =   975
         Left            =   -70000
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   13455
         _Version        =   1572864
         _ExtentX        =   23733
         _ExtentY        =   1720
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.CheckBox chkExcluye 
            Height          =   255
            Left            =   9960
            TabIndex        =   19
            Top             =   480
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Excluye"
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
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   330
            Left            =   2640
            TabIndex        =   20
            Top             =   480
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   330
            Left            =   4680
            TabIndex        =   21
            Top             =   480
            Width           =   5055
            _Version        =   1572864
            _ExtentX        =   8916
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnRegistra 
            Height          =   495
            Left            =   11760
            TabIndex        =   22
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Registra"
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
            Picture         =   "frmFNDExclusionMulta.frx":086A
         End
         Begin XtremeSuiteControls.FlatEdit txtContrato 
            Height          =   330
            Left            =   1440
            TabIndex        =   23
            ToolTipText     =   "Presione F4"
            Top             =   480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
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
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cédula"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   2640
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   4680
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   1440
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _Version        =   1572864
      _ExtentX        =   10821
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Presione F4"
      Top             =   480
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   4935
      _Version        =   1572864
      _ExtentX        =   8705
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   8760
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   582
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   8760
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   582
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption scResultados 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   2280
      Width           =   13455
      _Version        =   1572864
      _ExtentX        =   23733
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lista de Casos Filtrados con Exclusiones Gestionadas"
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   7920
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   7920
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmFNDExclusionMulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbLimpia()
 
 tcMain.Item(0).Selected = True
 
 txtContrato.Text = ""
 txtCedula.Text = ""
 txtNombre.Text = ""
 chkExcluye.Value = xtpUnchecked
 
 lsw.ListItems.Clear

End Sub


Private Sub btnBuscar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

Dim pContrato As String, pFecha As Date

txtC_Cedula.Text = fxSysCleanTxtInject(txtC_Cedula.Text)
txtC_Nombre.Text = fxSysCleanTxtInject(txtC_Nombre.Text)

If Not IsNumeric(txtC_Contrato.Text) Or txtC_Contrato.Text = "" Or txtC_Contrato.Text = "0" Then
    pContrato = "Null"
Else
    pContrato = txtC_Contrato.Text
End If

If chkFechas.Value = xtpUnchecked Then
    strSQL = "exec spFnd_Exclusion_Multas_List " & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo.Text _
           & "','" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00', '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
           & " 23:59:59', " & pContrato & ", '" & txtC_Cedula.Text & "', '" & txtC_Nombre.Text & "'"
Else
    pFecha = fxFechaServidor
    
    strSQL = "exec spFnd_Exclusion_Multas_List " & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo.Text _
           & "','1900-01-01 00:00:00', '" & Format(pFecha, "yyyy-mm-dd") _
           & " 23:59:59', " & pContrato & ", '" & txtC_Cedula.Text & "', '" & txtC_Nombre.Text & "'"
End If

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!IdRegistro)
     itmX.SubItems(1) = rs!Cedula
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = rs!COD_Contrato
     itmX.SubItems(4) = rs!Cod_Plan
     itmX.SubItems(5) = rs!Plan_Desc
     itmX.SubItems(6) = rs!Excluye_Desc
     itmX.SubItems(7) = rs!FECHA_REGISTRO & ""
     itmX.SubItems(8) = rs!USUARIO_REGISTRO & ""
     itmX.SubItems(9) = rs!FECHA_ACTUALIZA & ""
     itmX.SubItems(10) = rs!USUARIO_ACTUALIZA & ""

 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnRegistra_Click()
On Error GoTo vError

If Not IsNumeric(txtContrato.Text) Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Exclusion_Multas_Add " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & txtCodigo.Text & "', " & txtContrato.Text _
       & ", '" & txtCedula.Text & "', " & chkExcluye.Value & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Exclusión de Multa: Op." & cboOperadora.ItemData(cboOperadora.ListIndex) _
        & ", Plan: " & txtCodigo.Text & ", Cnt. " & txtContrato.Text & ", Ced. " & txtCedula.Text _
        & ", Excluye: " & IIf(chkExcluye.Value = xtpChecked, "Sí", "No"))

Me.MousePointer = vbDefault

MsgBox "Caso de Exclusión de Multa, registrado satisfactoriamente!", vbInformation

Call btnBuscar_Click

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = xtpChecked Then
    dtpInicio.Enabled = False
Else
    dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled
End Sub

Private Sub Form_Load()

vModulo = 18

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'Idx' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -60, dtpCorte.Value)

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id.", 1500
    .Add , , "Cédula", 2500, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Contrato", 2100, vbCenter
    .Add , , "Plan", 1200, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Excluye?", 1200, vbCenter
    .Add , , "Reg.Fecha", 2500
    .Add , , "Reg.Usuario", 2500, vbCenter
    .Add , , "Act.Fecha", 2500
    .Add , , "Act.Usuario", 2500, vbCenter
End With

tcMain.Item(0).Selected = True
Call chkFechas_Click

End Sub



Private Sub sbConsulta_Contrato()

   gBusquedas.Columna = "CEDULA"
   gBusquedas.Orden = "CEDULA"
   gBusquedas.Filtro = "And Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
                  & " and cod_Plan = '" & txtCodigo.Text & "' and Estado = 'A'"
   gBusquedas.Consulta = "select COD_CONTRATO, CEDULA, NOMBRE From vFnd_Contratos"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtContrato.Text = Trim(gBusquedas.Resultado)
      txtCedula.Text = Trim(gBusquedas.Resultado2)
      txtNombre.Text = Trim(gBusquedas.Resultado3)
   End If
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End Sub



Private Sub Form_Resize()
On Error Resume Next

lsw.Width = Me.Width - 250
tcMain.Width = lsw.Width
gbExcluye.Width = tcMain.Width
gbFiltros.Width = tcMain.Width

scResultados.Width = lsw.Width
lsw.Height = Me.Height - (lsw.top + 450)



End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

txtCedula.Text = Item.SubItems(1)
txtNombre.Text = Item.SubItems(2)
txtContrato.Text = Item.SubItems(3)

If Item.SubItems(6) = "Sí" Then
  chkExcluye.Value = xtpChecked
Else
  chkExcluye.Value = xtpUnchecked
End If
     
tcMain.Item(1).Selected = True
     
Exit Sub

vError:
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call Form_Resize
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  Call sbConsulta_Contrato
End If
End Sub

Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  Call sbConsulta_Contrato
End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  Call sbConsulta_Contrato
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo.Text = Trim(gBusquedas.Resultado)
      txtDescripcion.Text = Trim(gBusquedas.Resultado2)
      
      Call sbLimpia
      
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub


Private Sub txtCodigo_LostFocus()

If Trim(txtCodigo) <> "" Then
   strSQL = "Select Descripcion" _
          & " from fnd_planes where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " And cod_plan = '" & Trim(txtCodigo) & "'"
   Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
           txtDescripcion.Text = Trim(rs!Descripcion)
           
           Call sbLimpia
        Else
           MsgBox "Codigo incorrecto", vbExclamation
           txtCodigo.Text = ""
           txtDescripcion.Text = ""
           txtCodigo.SetFocus
        End If
     rs.Close

Else
  txtDescripcion.Text = ""
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If


End Sub

