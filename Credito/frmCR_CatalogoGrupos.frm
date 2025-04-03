VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCR_CatalogoGrupos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recursos Presupuestarios"
   ClientHeight    =   7944
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12756
   Icon            =   "frmCR_CatalogoGrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7944
   ScaleWidth      =   12756
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6492
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   12612
      _Version        =   1245187
      _ExtentX        =   22246
      _ExtentY        =   11451
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
      ItemCount       =   3
      Item(0).Caption =   "Consulta"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "Label6(0)"
      Item(0).Control(1)=   "cbo"
      Item(0).Control(2)=   "Label6(1)"
      Item(0).Control(3)=   "dtpInicio"
      Item(0).Control(4)=   "dtpCorte"
      Item(0).Control(5)=   "cmdFlujo"
      Item(0).Control(6)=   "cmdBuscar"
      Item(0).Control(7)=   "chkTodos"
      Item(0).Control(8)=   "chkActivos"
      Item(0).Control(9)=   "lswPD"
      Item(1).Caption =   "Definición"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tcAux"
      Item(2).Caption =   "Asignación"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "lsw"
      Item(2).Control(1)=   "lbl"
      Item(2).Control(2)=   "cmdReporte"
      Item(2).Control(3)=   "lswAsg"
      Begin XtremeSuiteControls.ListView lswPD 
         Height          =   5292
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   12372
         _Version        =   1245187
         _ExtentX        =   21823
         _ExtentY        =   9334
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswAsg 
         Height          =   5292
         Left            =   -63880
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   6012
         _Version        =   1245187
         _ExtentX        =   10604
         _ExtentY        =   9334
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5292
         Left            =   -69880
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   5892
         _Version        =   1245187
         _ExtentX        =   10393
         _ExtentY        =   9334
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
         HotTracking     =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   6132
         Left            =   -69160
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   10452
         _Version        =   1245187
         _ExtentX        =   18436
         _ExtentY        =   10816
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Recursos"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "vGrid"
         Item(1).Caption =   "Presupuesto diario"
         Item(1).ControlCount=   7
         Item(1).Control(0)=   "Label6(2)"
         Item(1).Control(1)=   "cboRecurso"
         Item(1).Control(2)=   "dtpPresupDia"
         Item(1).Control(3)=   "btnRegistrar"
         Item(1).Control(4)=   "Label6(3)"
         Item(1).Control(5)=   "txtPresupDia"
         Item(1).Control(6)=   "lswAsignacionRecursos"
         Begin XtremeSuiteControls.ListView lswAsignacionRecursos 
            Height          =   4812
            Left            =   -68440
            TabIndex        =   27
            Top             =   1320
            Visible         =   0   'False
            Width           =   8652
            _Version        =   1245187
            _ExtentX        =   15261
            _ExtentY        =   8488
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
            Appearance      =   16
         End
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   5532
            Left            =   480
            TabIndex        =   20
            Top             =   480
            Width           =   9132
            _Version        =   524288
            _ExtentX        =   16108
            _ExtentY        =   9758
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
            MaxCols         =   496
            ScrollBars      =   2
            SpreadDesigner  =   "frmCR_CatalogoGrupos.frx":030A
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRecurso 
            Height          =   312
            Left            =   -68440
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   4812
            _Version        =   1245187
            _ExtentX        =   8488
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
         Begin XtremeSuiteControls.DateTimePicker dtpPresupDia 
            Height          =   312
            Left            =   -64960
            TabIndex        =   23
            Top             =   840
            Visible         =   0   'False
            Width           =   1332
            _Version        =   1245187
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.PushButton btnRegistrar 
            Height          =   312
            Left            =   -63520
            TabIndex        =   24
            Top             =   840
            Visible         =   0   'False
            Width           =   1092
            _Version        =   1245187
            _ExtentX        =   1926
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Registrar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtPresupDia 
            Height          =   315
            Left            =   -68440
            TabIndex        =   26
            Top             =   840
            Visible         =   0   'False
            Width           =   3372
            _Version        =   1245187
            _ExtentX        =   5948
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   252
            Index           =   3
            Left            =   -69880
            TabIndex        =   25
            Top             =   840
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1245187
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Presupuesto Día:"
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   252
            Index           =   2
            Left            =   -69880
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1245187
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Referencia:"
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
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   252
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Marcar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton cmdFlujo 
         Height          =   312
         Left            =   8280
         TabIndex        =   8
         Top             =   480
         Width           =   972
         _Version        =   1245187
         _ExtentX        =   1714
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Flujo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   3372
         _Version        =   1245187
         _ExtentX        =   5948
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Height          =   312
         Left            =   6840
         TabIndex        =   7
         Top             =   480
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   312
         Left            =   9240
         TabIndex        =   9
         Top             =   480
         Width           =   972
         _Version        =   1245187
         _ExtentX        =   1714
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Consultar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkActivos 
         Height          =   252
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activos?"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   372
         Left            =   -59200
         TabIndex        =   16
         Top             =   6120
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   312
         Left            =   -69880
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   12012
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   1
         Left            =   4800
         TabIndex        =   5
         Top             =   480
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fechas:"
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
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Referencia:"
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
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   0
      Top             =   7812
      Width           =   12756
      _ExtentX        =   22500
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   372
      Left            =   0
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1245187
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton cmdModifica 
      Height          =   372
      Left            =   480
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1245187
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recursos Presupuestarios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   300
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmCR_CatalogoGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vPaso As Boolean


Private Sub btnRegistrar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Byte

On Error GoTo vError

strSQL = "select isnull(count(*),0) as Existe from catalogo_grupo_diario" _
       & " where cod_grupo = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "' and fecha = '" _
       & Format(dtpPresupDia.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)



If rs!Existe = 0 Then
    strSQL = "insert catalogo_grupo_diario(fecha,presupuesto,usuario,fechai,cod_grupo) values('" & Format(dtpPresupDia.Value, "yyyy/mm/dd") _
           & "'," & CCur(txtPresupDia) & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'" & cboRecurso.ItemData(cboRecurso.ListIndex) & "')"
    Call ConectionExecute(strSQL)
    Call Bitacora("Aplica", "Recurso Diario Fecha: " & Format(dtpPresupDia.Value, "yyyy/mm/dd") & " Rec:" & cboRecurso.ItemData(cboRecurso.ListIndex))
Else
  i = MsgBox("Ya existe un monto presupuestario definido para este día, desea reemplazarlo ?", vbYesNo)
  If i = vbYes Then
     strSQL = "update catalogo_grupo_diario set presupuesto = " & CCur(txtPresupDia) _
            & ",usuario = '" & glogon.Usuario & "',fechai = dbo.MyGetdate()" _
            & " where cod_grupo = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "' and fecha = '" _
            & Format(dtpPresupDia.Value, "yyyy/mm/dd") & "'"
     Call ConectionExecute(strSQL)
     Call Bitacora("Modifica", "Recurso Diario Fecha: " & Format(dtpPresupDia.Value, "yyyy/mm/dd") & " Rec:" & cboRecurso.ItemData(cboRecurso.ListIndex))
  End If
End If

rs.Close

txtPresupDia = 0

Call cboRecurso_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboRecurso_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

With lswAsignacionRecursos
  .ListItems.Clear
  strSQL = "select Top 20 * from catalogo_grupo_diario" _
         & " where cod_grupo = '" & cboRecurso.ItemData(cboRecurso.ListIndex) _
         & "' order by fecha desc"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , Format(rs!fecha, "dd/mm/yyyy"))
        itmX.SubItems(1) = Format(rs!Presupuesto, "Standard")
        itmX.SubItems(2) = rs!Usuario
        itmX.SubItems(3) = rs!fechaI
    rs.MoveNext
  Loop
  rs.Close

End With

End Sub

Private Sub chkActivos_Click()
Call sbConsulta(0)
End Sub

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lswPD.ListItems.Count
  lswPD.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Private Sub sbPresupuestoAcumCorte()
Dim strSQL As String, rs As New ADODB.Recordset
Dim x As Integer, y As Integer, vFecha As Date, i As Integer
Dim curPresu As Currency, curTempo As Currency, vFechaX  As Date

For i = 1 To lswPD.ListItems.Count

    x = DateDiff("d", dtpInicio.Value, dtpCorte.Value) + 1
    curPresu = 0
     
'    If Trim(lswPD.ListItems.Item(i).Text) = "TESOR" Then MsgBox "ya"
     
    For y = 1 To x
     strSQL = "select presu_diario from catalogo_grupos where cod_grupo = '" _
            & lswPD.ListItems.Item(i).Text & "'"
     Call OpenRecordSet(rs, strSQL)
         curTempo = IIf(IsNull(rs!presu_diario), 0, rs!presu_diario)
     rs.Close
     
     'Busca y Reemplazo Por Presupuesto Ajustado
     vFecha = DateAdd("d", y, DateAdd("d", -1, dtpInicio.Value))
     
     strSQL = "select presupuesto from catalogo_grupo_diario where cod_grupo = '" _
            & lswPD.ListItems.Item(i).Text & "' and fecha = '" & Format(vFecha, "yyyy/mm/dd") & "'"
     Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF And Not rs.BOF Then
         curTempo = rs!Presupuesto
     End If
     rs.Close
     curPresu = curPresu + curTempo
    Next y
    
    lswPD.ListItems.Item(i).SubItems(2) = Format(curPresu, "Standard")

Next i

End Sub


Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curMonto As Currency


Me.MousePointer = vbHourglass

Call sbPresupuestoAcumCorte

With lswPD.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then

''          & " inner join Catalogo C on C.codigo = R.codigo" _
''          & " inner join CATALOGO_ASIGNAGRP G on C.codigo = G.codigo"

   strSQL = "select R.id_solicitud,R.monto_girado,isnull(sum(monto),0) as Desembolso" _
          & " from reg_creditos R left join Desembolsos D on R.id_solicitud = D.id_solicitud and D.retener = 0" _
          & " where " & IIf((Mid(cbo.Text, 1, 2) = "01"), "R.fechaforp", "R.fecha_Inicio_Calculo") _
          & " between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
          & " and R.estadosol = 'F' and R.monto_girado >= 0  and R.cod_grupo = '" & .Item(i).Text & "'" _
          & " group by R.id_solicitud,R.monto_girado"
    Call OpenRecordSet(rs, strSQL)
    curMonto = 0
    prgBar.Max = rs.RecordCount + 1
    prgBar.Value = 1
    Do While Not rs.EOF
     curMonto = curMonto + rs!monto_girado + rs!desembolso
     prgBar.Value = prgBar.Value + 1
     rs.MoveNext
    Loop
    rs.Close
    
    .Item(i).SubItems(3) = Format(curMonto, "Standard")
    .Item(i).SubItems(4) = Format((CCur(.Item(i).SubItems(2)) - curMonto), "Standard")
    
    If CCur(.Item(i).SubItems(4)) < 0 Then
      .Item(i).ForeColor = vbRed
    Else
      .Item(i).ForeColor = vbBlack
    End If
    
 End If
Next i

End With

Me.MousePointer = vbDefault

Exit Sub
 
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdFlujo_Click()
Dim i As Byte

i = MsgBox("Desea visualizar flujos Agrupados x Fechas", vbYesNo)


Me.MousePointer = vbHourglass

With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Crédito"

   .Connect = glogon.ConectRPT

   .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
   .Formulas(1) = "Fecha= '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
   .Formulas(2) = "fxSubTitulo = 'Inicio : " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Corte : " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
   
   If i = vbNo Then
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_RecursosFlujoDiario.rpt")
   Else
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_RecursosFlujoDiarioGrp.rpt")
   End If
   .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy/mm/dd")
   .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd")
   
   .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Crédito"

   .Connect = glogon.ConectRPT

   .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
   .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoGrupos.rpt")
   .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

vModulo = 3


vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True


cbo.AddItem "01 - Formalización"
cbo.AddItem "02 - Desembolso"
cbo.Text = "02 - Desembolso"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpPresupDia.Value = dtpInicio.Value

With lswPD.ColumnHeaders
    .Clear
    .Add , , "Código", 1050
    .Add , , "Descripción", 3500
    .Add , , "Presupuesto", 2100, vbRightJustify
    .Add , , "Real", 2100, vbRightJustify
    .Add , , "Diferencia", 2100, vbRightJustify
End With

With lswAsignacionRecursos.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2000
    .Add , , "Monto", 2500, vbRightJustify
    .Add , , "Usuario", 2000, vbCenter
    .Add , , "Fec.Sys.", 2000, vbCenter
End With


With lsw.ColumnHeaders
  .Clear
  .Add , , "Código", 1440
  .Add , , "Descripción", 3600
End With


With lswAsg.ColumnHeaders
  .Clear
  .Add , , "Código", 1440
  .Add , , "Descripción", 3600
End With


Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdModifica.Enabled
btnRegistrar.Enabled = cmdModifica.Enabled
lswAsg.Enabled = cmdActualiza.Enabled

Call sbConsulta(0)

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If vGrid.Text = "" Then Exit Function


If Not fxExiste(vGrid.Text) Then
   vGrid.col = 1
   strSQL = "insert CATALOGO_GRUPOS(cod_grupo,descripcion,presu_mensual,presu_diario,estado)" _
          & " values('" & vGrid.Text & "','"
   vGrid.col = 2
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.col = 3
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 4
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 5
   strSQL = strSQL & vGrid.Value & ")"
   
   
   
   Call ConectionExecute(strSQL)
   vGrid.col = 1
   Call Bitacora("Registra", "Grupo Catalogo Adicional Cod: " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.col = 2
    strSQL = "update CATALOGO_GRUPOS set descripcion = '" & vGrid.Text
    vGrid.col = 3
    strSQL = strSQL & "',presu_mensual = " & CCur(vGrid.Text) & ",presu_diario = "
    vGrid.col = 4
    strSQL = strSQL & CCur(vGrid.Text) & ",estado = "
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value
    vGrid.col = 1
    strSQL = strSQL & " where cod_grupo = '" & vGrid.Text & "'"
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Grupo Catalogo Cod: " & vGrid.Text)
    
End If

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function

Private Function fxExiste(vCod As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from CATALOGO_GRUPOS" _
       & " where cod_grupo = '" & vCod & "'"
Call OpenRecordSet(rs, strSQL)
fxExiste = IIf((rs!Existe = 1), True, False)
rs.Close
End Function


Private Sub sbCargaLswAdicional()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

lswAsg.ListItems.Clear

vPaso = True

strSQL = "select codigo,descripcion,retencion,poliza,convenio" _
       & " From catalogo " _
       & " where codigo in(select codigo from CATALOGO_ASIGNAGRP where cod_grupo = '" _
       & vCodigo & "') order by codigo"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  Set itmX = lswAsg.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.Checked = True
  If itmX.Checked Then itmX.ForeColor = vbBlue
      
   If rs!retencion = "S" Or rs!Poliza = "S" Then
     itmX.SubItems(2) = "Retencion"
   Else
     itmX.SubItems(2) = "Cartera"
   End If

   If rs!Convenio = "S" Then itmX.SubItems(2) = Mid(itmX.SubItems(2), 1, 3) & ".Convenio"
  
  
  rs.MoveNext
Loop
rs.Close



strSQL = "select codigo,descripcion,retencion,poliza,convenio" _
       & " From catalogo " _
       & " where codigo not in(select codigo from CATALOGO_ASIGNAGRP where cod_grupo = '" _
       & vCodigo & "') order by codigo"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  Set itmX = lswAsg.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.Checked = False
  
   If rs!retencion = "S" Or rs!Poliza = "S" Then
     itmX.SubItems(2) = "Retencion"
   Else
     itmX.SubItems(2) = "Cartera"
   End If

   If rs!Convenio = "S" Then itmX.SubItems(2) = Mid(itmX.SubItems(2), 1, 3) & ".Convenio"
  
  
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
  
  vCodigo = Item.Text
  lbl.Caption = Item.Text & " ¦ " & Item.SubItems(1)
  Call sbCargaLswAdicional

End Sub


Private Sub lswAsg_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert CATALOGO_ASIGNAGRP(codigo,cod_grupo) values('" & Item.Text _
          & "','" & vCodigo & "')"
Else
   strSQL = "delete CATALOGO_ASIGNAGRP where codigo = '" & Item.Text & "' and cod_grupo = '" _
          & vCodigo & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbConsulta(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Me.MousePointer = vbHourglass


Select Case Index
   Case 0 'Presupuesto
        
        lswPD.ListItems.Clear
        strSQL = "select cod_grupo,descripcion,presu_diario from catalogo_grupos " _
               & " where estado = " & chkActivos.Value & " order by cod_grupo"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
         Set itmX = lswPD.ListItems.Add(, , rs!cod_Grupo)
             itmX.SubItems(1) = rs!Descripcion
             itmX.SubItems(2) = Format(rs!presu_diario, "Standard")
             itmX.SubItems(3) = 0
             itmX.SubItems(4) = Format(rs!presu_diario, "Standard")
             
         rs.MoveNext
        Loop
        rs.Close
   
   
   Case 1 'Definicion
        
        tcAux.Item(0).Selected = True
        
        strSQL = "select COD_GRUPO,descripcion,presu_mensual,presu_diario,estado from CATALOGO_GRUPOS" _
               & " order by cod_grupo"
        Call sbCargaGrid(vGrid, 5, strSQL)
   
   Case 2 'Asignacion
                
        vCodigo = ""
        lbl.Caption = ""
        lswAsg.ListItems.Clear
        
        strSQL = "select cod_grupo,descripcion from catalogo_grupos order by cod_grupo"
        Call OpenRecordSet(rs, strSQL, 0)
        lsw.ListItems.Clear
        Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!cod_Grupo)
             itmX.SubItems(1) = rs!Descripcion & ""
         rs.MoveNext
        Loop
        rs.Close

End Select

Me.MousePointer = vbDefault

End Sub

Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 1 Then
  
  vPaso = True
  strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX' from catalogo_grupos"
  Call sbCbo_Llena_New(cboRecurso, strSQL, False, True)
  
  vPaso = False
  
  Call cboRecurso_Click
End If

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call sbConsulta(Item.Index)
End Sub

Private Sub txtPresupDia_GotFocus()
On Error GoTo vError
 txtPresupDia.Text = CCur(txtPresupDia)
vError:
End Sub

Private Sub txtPresupDia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpPresupDia.SetFocus
End Sub

Private Sub txtPresupDia_LostFocus()
On Error GoTo vError
 txtPresupDia.Text = Format(CCur(txtPresupDia.Text), "Standard")
vError:
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

End Sub



