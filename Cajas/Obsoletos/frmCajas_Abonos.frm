VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajas_Abonos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplica Abonos"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOperacion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   52
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   51
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   50
      ToolTipText     =   "Nombre Completo del Socio (Apellidos y Nombre)"
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox txtCedula 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   49
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame fraAbono 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   42
      Top             =   3600
      Width           =   6615
      Begin VB.OptionButton optAbono 
         BackColor       =   &H00C00000&
         Caption         =   "Ordinario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H00C00000&
         Caption         =   "Extra Ordinario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H00C00000&
         Caption         =   "Cancelación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFechaCancelacion 
         Height          =   315
         Left            =   4920
         TabIndex        =   46
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   175112195
         CurrentDate     =   40310
      End
      Begin VB.Label lblFechaCancelacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Abono (Real) por parte del cliente para cancelación...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Abono   >"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraDatosAbono 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   6255
      Begin VB.ComboBox cboTipoPago 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCajas_Abonos.frx":0000
         Left            =   4080
         List            =   "frmCajas_Abonos.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCuotas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   31
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtDatosAmortiza 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "frmCajas_Abonos.frx":0037
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCajas_Abonos.frx":003B
         Left            =   4080
         List            =   "frmCajas_Abonos.frx":003D
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtTotalPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   $"frmCajas_Abonos.frx":003F
         Top             =   1440
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar tblDesgloce 
         Height          =   330
         Left            =   5760
         TabIndex        =   58
         ToolTipText     =   "Detalle de pago"
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Aplicar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label23 
         Caption         =   "Tipo - Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "# Cuotas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDatosInteres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo - Doc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDatosAnticipo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Cargo Anticipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3120
         X2              =   6000
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.CommandButton CmdAbono 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Picture         =   "frmCajas_Abonos.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdReporte 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Reporte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CheckBox chkRecalculaCuota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Recalcular Cuota"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   23
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Timer TimerVerificaPlanPagos 
      Left            =   7680
      Top             =   360
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "r0$psjFdix"
      Top             =   1680
      Width           =   7935
      Begin VB.Label lblFecUltMov 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCuota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblAmortiza 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         ToolTipText     =   "Si es menor a la fecha de proceso se Utiliza la Fecha de Proceso"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortiza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Estado Actual ....:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Estado Resultante....:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblSaldoR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortiza."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblAmortizaR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblInteresR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblCuotaR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblFecUltMovR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Abonos.frx":017E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblOpex 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7440
      TabIndex        =   57
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   960
      TabIndex        =   56
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   55
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H8000000E&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2640
      TabIndex        =   54
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   53
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8280
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmCajas_Abonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vCuotasDeducidas As Integer, vCuotasDirectas As Integer
Dim vInteres As Currency, vPlazo As Integer, vSaldoMes As Currency, vUltimoRecibo As Long
Dim vRetencion As Boolean, vBaseCalculo As String, vPrideduc As Long
Dim pDatos() As Currency, vFechaHoy As Date

Private Function fxVerificaMorosidad(lngOperacion As Long) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select coalesce(count(*),0) as Existe from Morosidad" _
       & " where estado = 'A' and id_solicitud = " & lngOperacion
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
    fxVerificaMorosidad = IIf((rsX!existe = 0), False, True)
rsX.Close
End Function


Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then CmdAbono.SetFocus
End Sub


Private Sub chkRecalculaCuota_Click()

If vRetencion Then
   chkRecalculaCuota.Value = vbUnchecked
   MsgBox "Las retenciones no se pueden Ajustar para Recálculos, verifique...", vbExclamation
   Exit Sub
End If

Call txtTotalPagar_Change

End Sub

Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngRecibo As Long, vCuenta As String
Dim vTipo As String, vFecha As Date, vTipoDoc As String
Dim vFechaProceso As Long, i As Integer, vConcepto As String


Me.MousePointer = vbHourglass


On Error GoTo vError

lngRecibo = 0
vFecha = fxFechaServidor

'Configuracion del Documento
'vTipo = fxTipoASEDoc(cboTipo.Text)
vTipo = SIFGlobal.fxSIFCodText(cboTipo.Text)
vCuenta = Trim(fxDocumentoCuenta(vTipo))
lngRecibo = fxDocumentoConsecutivo(vTipo)

vUltimoRecibo = lngRecibo

If vAseDocValido = False Then
  MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
        & " válida para esta operación...", vbCritical
  Exit Sub
End If


'Genera el Comprobante
Select Case True
  Case optAbono(0) 'Abono Ordinario
      vConcepto = "CRD001"
  
  Case optAbono(1) 'Abono Extraordinario
      vConcepto = "CRD002"
  
  Case optAbono(2) 'Abono De Cancelacion
      vConcepto = "CRD003"
End Select


If CLng(lblFecUltMovR) < GLOBALES.glngFechaCR Then
  lblFecUltMovR.Caption = GLOBALES.glngFechaCR
End If
vFechaProceso = lblFecUltMovR.Caption


If optAbono.Item(1).Value Then
  vTipo = "E"
  vFechaProceso = Format(dtpFechaCancelacion.Value, "yyyymm")
Else
  vTipo = "O"
End If


'Inicia Transaccion
glogon.Conection.BeginTrans

If optAbono.Item(0).Value And CLng(txtCuotas.Text) > 1 Then
    'Varias Cuotas
'          pDatos(i, 1) = curTmpInteres
'          pDatos(i, 2) = curTmpAmortiza
'          pDatos(i, 3) = lngFecha
'          pDatos(i, 4) = curSaldo
'          pDatos(i, 5) = curCuota
   For i = 1 To CLng(txtCuotas.Text)
        strSQL = "insert into creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
               & "fechas,fechap,tcon,ncon,saldo,usuario,cod_concepto,cod_Caja) values('" & txtCodigo & "'," & vOperacion & "," _
               & IIf((vTipo = "E"), 0, pDatos(i, 5)) & "," & pDatos(i, 1) + pDatos(i, 2) & "," _
               & pDatos(i, 1) & "," & pDatos(i, 2) & "," _
               & "'" & Format(vFecha, "yyyy/mm/dd") & "'," & CLng(pDatos(i, 3)) & ",'" & fxTipoASENumero(SIFGlobal.fxSIFCodText(cboTipo.Text)) _
               & "','" & IIf((lngRecibo = 0), "null", lngRecibo) & "'," & pDatos(i, 4) & ",'" & glogon.Usuario & "','" & vConcepto & "','')"
        glogon.Conection.Execute strSQL
   Next i
Else
    strSQL = "insert into creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
           & "fechas,fechap,tcon,ncon,saldo,usuario,cod_concepto,cod_Caja) values('" & txtCodigo & "'," & vOperacion & "," _
           & IIf((vTipo = "E"), 0, CCur(lblCuota.Caption)) & "," & CCur(txtTotalPagar.Text) & "," _
           & CCur(lblDatosInteres.Caption) & "," & CCur(txtDatosAmortiza) & "," _
           & "'" & Format(vFecha, "yyyy/mm/dd") & "'," & vFechaProceso & ",'" & fxTipoASENumero(SIFGlobal.fxSIFCodText(cboTipo.Text)) _
           & "','" & IIf((lngRecibo = 0), "null", lngRecibo) & "'," & CCur(lblSaldoR.Caption) & ",'" & glogon.Usuario & "','" & vConcepto & "','')"
    glogon.Conection.Execute strSQL
End If


If Not vRetencion Then
     strSQL = "UPDATE REG_CREDITOS set saldo= saldo - " & CCur(txtDatosAmortiza) _
            & ",saldo_mes= saldo_mes - " & CCur(txtDatosAmortiza) _
            & ",amortiza=amortiza + " & CCur(txtDatosAmortiza) _
            & ",interesc=interesc + " & CCur(lblDatosInteres.Caption)
            
'            & ",cuota = " & CCur(lblCuotaR.Caption)
     
     If vTipo = "E" Then
       strSQL = strSQL & ",cuotas_directas= " & vCuotasDirectas + 1
     Else
       strSQL = strSQL & ",cuotas_planilla= " & vCuotasDeducidas + CLng(txtCuotas) _
              & ",fecult=" & lblFecUltMovR.Caption
     End If
     
     If CCur(lblSaldoR.Caption) = 0 Then strSQL = strSQL & ",estado='C'"
     
     
    If chkRecalculaCuota.Value = vbChecked Then  'Recalculo de Cuota
      
      'Se reinicia la secuencia en cuotas_planila
      strSQL = strSQL & ",monto_recalculo=" & CCur(lblSaldo.Caption) _
           & ",fecha_recalculo='" & Format(vFecha, "yyyy/mm/dd") & "'" _
           & ",plazo_recalculo=" & vPlazo - vCuotasDeducidas _
           & ",indicador_recalculo=1" _
           & ",cuota = " & CCur(lblCuotaR.Caption)
    End If
     
    strSQL = strSQL & " where id_solicitud=" & vOperacion

Else
 'Retencion
     strSQL = "UPDATE REG_CREDITOS set amortiza=amortiza + " & CCur(txtDatosAmortiza) _
            & ",interesc=interesc+" & CCur(lblDatosInteres.Caption)
     
     If vTipo = "E" Then
       strSQL = strSQL & ",cuotas_directas= " & vCuotasDirectas + 1
     Else
       strSQL = strSQL & ",cuotas_planilla= " & vCuotasDeducidas + CLng(txtCuotas) _
              & ",fecult=" & lblFecUltMovR.Caption
     End If
     
   If CCur(txtDatosAmortiza) >= ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) Then
     strSQL = strSQL & ",estado = 'C' where id_solicitud=" & vOperacion
   Else
     strSQL = strSQL & " where id_solicitud=" & vOperacion
   End If

End If

glogon.Conection.Execute strSQL

'Cierra Transaccion
glogon.Conection.CommitTrans

'vTipo = fxTipoASEDoc(cboTipo.Text)
vTipo = SIFGlobal.fxSIFCodText(cboTipo.Text)
'Genera el Comprobante
Select Case True
  Case optAbono(0) 'Abono Ordinario
      Call Bitacora("Registra", "Abono Ordinario a la Operacion : " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("ABONO ORDINARIO", vTipo, CStr(lngRecibo), vConcepto, vCuenta)
  
  Case optAbono(1) 'Abono Extraordinario
      Call Bitacora("Registra", "Abono ExtraOrd. " & IIf((chkRecalculaCuota.Value = 1), "Con Recal.", "Sin Recal") & " a la Op.: " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("ABONO EXTRAORDINARIO", vTipo, CStr(lngRecibo), vConcepto, vCuenta)
  
  Case optAbono(2) 'Abono De Cancelacion
      Call Bitacora("Registra", "Cancelación de la Operacion : " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("CANCELACION DE DEUDA", vTipo, CStr(lngRecibo), vConcepto, vCuenta)
End Select

'IMPRIMIR RECIBO
If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, SIFGlobal.fxSIFCodText(cboTipo.Text))
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 MsgBox Err.Description, vbCritical
End Sub

Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""


'Verifica el proceso
If txtProceso.Tag = "J" Then
   If Not fxCRDAbonosAutorizados(txtCodigo.Text, txtProceso.Tag) Then
      vMensaje = vMensaje & "- El usuario actual no cuenta con permisos para realizar abonos a Creditos en Cobro Judicial, verifique..." & vbCrLf
   End If
End If

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  vMensaje = vMensaje & "- Esta Persona se encuentra CONGELADA, verifique..." & vbCrLf
End If

If vOperacion = 0 Then
  vMensaje = vMensaje & "- Número de Operacion no es válido..." & vbCrLf
End If
 
 
'Verifica Saldo Actual
If Not fxCRDSaldoVerifica(vOperacion, CCur(lblSaldo.Caption)) Then
   vMensaje = vMensaje & "- Esta Operación ha sido modificada, actualice los datos nuevamente antes de realizar el abono..." & vbCrLf
End If
 
If Not vRetencion Then
    If CCur(txtDatosAmortiza) > CCur(lblSaldo.Caption) Then
       vMensaje = vMensaje & "- La Amortización es mayor al Saldo Actual..." & vbCrLf
    End If
Else
    If CCur(txtDatosAmortiza) > ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) Then
        vMensaje = vMensaje & "- La Amortización es mayor que el Remanente a Recaudar : " _
              & ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) & vbCrLf
     End If
End If

If Not IsNumeric(txtTotalPagar.Text) Then
  vMensaje = vMensaje & "- El total a pagar no es un dato válido...verifique...!" & vbCrLf
Else
 If CCur(txtTotalPagar.Text) <= 0 Then
      vMensaje = vMensaje & "- El total a pagar no es un dato válido...verifique...!" & vbCrLf
 End If
End If



If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub CmdAbono_Click()
Dim iRespuesta As Integer

If Not fxVerifica Then Exit Sub

 iRespuesta = MsgBox("Esta seguro de realizar el abono a esta Operación " & vOperacion, vbYesNo)
 If iRespuesta = vbYes Then
  
  Call sbAbono
  If vAseDocValido Then MsgBox "Abono Realizado ... " & cboTipo.Text & " #" & vUltimoRecibo, vbInformation
  
  Call sbConsultaOperacion
 
 Else 'Respuesta
  
  MsgBox "Transacción Cancelada...", vbInformation
 CmdAbono.Enabled = False
 End If

End Sub

Private Sub sbReporte(vTitulo As String)
If vOperacion = 0 Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Módulo de Crédito"
 
 .ReportFileName = SIFGlobal.fxSIFPathReportes("CrdBoletaAbono.rpt")
 
 .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "usuario='" & glogon.Usuario & "'"
 If optAbono(0).Value = True Then
  .Formulas(4) = "tipo_abono='ABONO ORDINARIO : CUOTAS: " & txtCuotas & "'"
 Else
  .Formulas(4) = "tipo_abono='ABONO EXTRAORDINARIO'"
 End If
 .Formulas(5) = "saldo_actual='" & lblSaldo.Caption & "'"
 .Formulas(6) = "amortizacion='" & Me.lblAmortiza.Caption & "'"
 .Formulas(7) = "interesc='" & lblInteres.Caption & "'"
 .Formulas(8) = "fecult='" & Format(lblFecUltMov.Caption, "####-##") & "'"

 .Formulas(9) = "saldo_res='" & lblSaldoR.Caption & "'"
 .Formulas(10) = "amortizacion_res='" & Me.lblAmortizaR.Caption & "'"
 .Formulas(11) = "interesc_res='" & lblInteresR.Caption & "'"
 .Formulas(12) = "fecult_res='" & Format(lblFecUltMovR.Caption, "####-##") & "'"
 
 .Formulas(13) = "abono_amortizacion='" & txtDatosAmortiza & "'"
 .Formulas(14) = "abono_interes='" & lblDatosInteres.Caption & "'"
 .Formulas(15) = "abono_total='" & txtTotalPagar.Text & "'"
 
 .Formulas(16) = "titulo='" & vTitulo & "'"
 .Formulas(17) = "operacion='" & vOperacion & "'"
 .Formulas(18) = "cedula='" & txtCedula & " - " & txtNombre & "'"
 .Formulas(19) = "codigo='" & txtCodigo & " - " & lblDescripcion.Caption & "'"
 
 .PrintReport
End With
Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()

Call sbReporte("ABONO A REALIZAR")

End Sub


Private Sub dtpFechaCancelacion_Change()

If dtpFechaCancelacion.Enabled Then
   'Refresca información base para Cancelación y/o Abonos Extraordinarios
   Select Case True
      Case optAbono.Item(1).Value 'Abono Extraordinario
            Call optAbono_Click(1)
      Case optAbono.Item(2).Value 'Cancelación
            Call optAbono_Click(2)
   End Select
End If
End Sub

Private Sub Form_Activate()
 vModulo = 5
 Call RefrescaTags(Me)
End Sub

Private Sub Form_Load()
Dim iDias As Integer
Dim strSQL As String

 vModulo = 5
 vOperacion = 0
 
 If GLOBALES.SysPlanPagos = 1 Then
    TimerVerificaPlanPagos.Interval = 10
 Else
   'Carga Load Normalmente
    Call Formularios(Me)
    strSQL = "select C.tipo_documento + ' - ' + D.Descripcion as itmx  from SIF_DOCUMENTOS D" _
           & " inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
           & " AND  C.cod_caja =  '" & ModuloCajas.mCaja & "' order by C.tipo_documento"
     
   Call sbLlenaCbo(cboTipo, strSQL, False, False)
'    cboTipo.Clear
'    cboTipo.AddItem "Recibo"
'    cboTipo.AddItem "Nota Credito"
'    cboTipo.Text = "Recibo"
'    Call sbLimpiaDatos
 End If
 
vFechaHoy = fxFechaServidor
iDias = fxCRDParametro("32")

dtpFechaCancelacion.Value = vFechaHoy
dtpFechaCancelacion.MinDate = DateAdd("d", (iDias * -1), dtpFechaCancelacion.Value)
dtpFechaCancelacion.MaxDate = dtpFechaCancelacion.Value

CmdAbono.Enabled = False

ModuloCajas.mAplicarSF = 0

End Sub


Private Sub sbConsultaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curSaldo As Currency

Me.MousePointer = vbHourglass

 
strSQL = "select R.id_solicitud,R.saldo, R.saldo - coalesce(V.amortiza,0) As Saldo_mes,R.proceso" _
       & ",R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult,R.Prideduc" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas,R.montoApr" _
       & ",S.nombre,C.descripcion,C.retencion,C.poliza,R.fechaforp,C.PORC_CARGO_CANCELACION,R.Base_Calculo" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " where R.estado = 'A' and R.saldo > 0" _
       & " and R.ID_SOLICITUD = " & vOperacion
       
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF And Not rs.BOF Then
  vBaseCalculo = Trim(rs!Base_Calculo)
  vPrideduc = rs!Prideduc
  vOperacion = rs!ID_SOLICITUD
  vPlazo = rs!Plazo
  
  
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  If IsNull(rs!saldo_mes) Then
    vSaldoMes = rs!Saldo
    strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!ID_SOLICITUD
    glogon.Conection.Execute strSQL
  Else
    If rs!saldo_mes = 0 Then
        vSaldoMes = rs!Saldo
        strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!ID_SOLICITUD
        glogon.Conection.Execute strSQL
    Else
       vSaldoMes = rs!saldo_mes
    End If
  
  End If
  
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     lblDatosInteres.Caption = 0
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
    
    
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")
    
     lblSaldo.Tag = rs!fechaforp
     lblSaldo.Caption = Format(rs!Saldo, "Standard")
     lblSaldoR.Caption = 0
    
     txtCuotas = 0
     txtOperacion = rs!ID_SOLICITUD
     cboTipoPago.Text = "Efectivo"
     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    txtCedula = rs!Cedula
    txtNombre = rs!Nombre
    txtCodigo = rs!Codigo
    
    txtProceso.Tag = rs!Proceso
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
    
    
    lblDescripcion.Caption = rs!Descripcion
    
    lblDatosAnticipo.ToolTipText = "% de Comision : " & rs!PORC_CARGO_CANCELACION
    lblDatosAnticipo.Tag = rs!PORC_CARGO_CANCELACION
    
    optAbono(0).Enabled = True
    optAbono(1).Enabled = True
    
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
     If rs!Plazo < 900 Then
        lblSaldo.Caption = Format((rs!montoapr * rs!Plazo) - rs!Amortiza, "Standard")
        lblSaldoR.Caption = 0
        vSaldoMes = CCur(lblSaldo.Caption)
     End If
    Else
      vRetencion = False
    End If
        
        Select Case True
         Case optAbono(0).Value
           Call optAbono_Click(0)
         Case optAbono(1).Value
           Call optAbono_Click(1)
         Case optAbono(2).Value
           Call optAbono_Click(2)
        End Select

Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation
 Call sbLimpiaDatos

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
 lblDatosAnticipo.Caption = 0
 
 lblAmortiza.Caption = 0
 lblAmortizaR.Caption = 0
 lblCuota = 0
 lblCuotaR.Caption = 0
 txtDatosAmortiza = 0
 lblDatosInteres.Caption = 0
 lblFecUltMov.Caption = 0
 lblFecUltMovR.Caption = 0
 lblInteres.Caption = 0
 lblInteresR.Caption = 0
 lblDescripcion.Caption = ""
 lblOpex.Caption = ""
 lblSaldo.Caption = 0
 lblSaldoR.Caption = 0
 txtCedula = ""
 txtCodigo = ""
 txtCuotas = 0
 txtNombre = ""
 txtOperacion = ""
 cboTipoPago.Text = "Efectivo"
 txtTotalPagar.Text = 0
 
 txtProceso.Tag = ""
 txtProceso.Text = ""
 
 fraAbono.Enabled = False
 fraDatosAbono.Enabled = False
 
 chkRecalculaCuota.Value = vbUnchecked
 
dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False
dtpFechaCancelacion.Value = vFechaHoy
 
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
gBusquedas.Consulta = "Select R.id_solicitud as Operacion,R.Codigo,S.Cedula,S.Nombre,C.Descripcion" _
          & " from REG_CREDITOS R inner join SOCIOS S on R.cedula = S.cedula" _
          & " inner join Catalogo C on R.codigo = C.codigo"
gBusquedas.Columna = "R.CEDULA"
gBusquedas.Orden = "R.CEDULA"
gBusquedas.Filtro = " AND R.ESTADO = 'A'"

frmBusquedas.Show vbModal

txtOperacion = Trim(gBusquedas.Resultado)
vOperacion = txtOperacion

gBusquedas.Consulta = ""
gBusquedas.Columna = ""
gBusquedas.Orden = ""
gBusquedas.Resultado = ""
gBusquedas.Filtro = ""

Call sbConsultaOperacion

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

End Sub

Private Sub sbCargaOperacionCodCed(vCedula As String, vCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select R.id_solicitud,R.saldo,R.saldo_mes,R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas,C.retencion,C.poliza " _
       & "from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & "where R.estado = 'A' and R.proceso <> 'N' and R.saldo > 0 " _
       & "and R.cedula = '" & txtCedula & "' and R.codigo = '" & txtCodigo & "'"
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  vOperacion = rs!ID_SOLICITUD
  vPlazo = rs!Plazo
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  vSaldoMes = IIf(IsNull(rs!saldo_mes), rs!Saldo, rs!saldo_mes)
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     lblDatosInteres.Caption = 0
    
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")
     lblSaldo.Caption = Format(vSaldoMes, "Standard")
     lblSaldoR.Caption = 0
     txtCuotas = 0
     txtOperacion = rs!ID_SOLICITUD
     cboTipoPago.Text = "Efectivo"
     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    
    optAbono(0).Enabled = True
    optAbono(1).Enabled = True
    
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    If Not vRetencion Then
        Select Case True
         Case optAbono(0).Value
           Call optAbono_Click(0)
         Case optAbono(1).Value
           Call optAbono_Click(1)
        End Select
    Else
           Call optAbono_Click(0)
           optAbono(1).Enabled = False
    End If
    
    
Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontrarón operaciones para abonos con esta cédula y código", vbInformation
End If
rs.Close

End Sub




Private Sub tblDesgloce_ButtonClick(ByVal Button As MSComctlLib.Button)

ModuloCajas.mTotalAplicar = txtTotalPagar.Text
ModuloCajas.mCliente = txtCedula.Text & "-" & txtNombre.Text
ModuloCajas.mServicio = "Aplica Abonos"

frmCajas_DetallePago.Show vbModal
    
If ModuloCajas.mTiquete = Empty Then
    CmdAbono.Enabled = False
    MsgBox "Aún no ha desglosado la transacción"
    Exit Sub
Else
    CmdAbono.Enabled = True
    Call RefrescaTags(Me)
    If optAbono.Item(1).Value = True Then
      txtTotalPagar.Text = Format(ModuloCajas.mTotalAplicar, "Standard")
    End If
End If
End Sub

Private Sub TimerVerificaPlanPagos_Timer()

TimerVerificaPlanPagos.Interval = 0
Call sbSIFForms("frmCR_AbonosNew", 0, , , False)

Unload Me

End Sub

Private Sub txtTotalPagar_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim ProcesosTmp As Long, lngFecha As Long, iPlazoRst As Integer, curCuota As Currency

On Error Resume Next

If chkRecalculaCuota.Value = vbChecked Then
  
' strSQL = "select plazo + DATEDIFF(mm,  getdate(), CONVERT(DATETIME, substring(convert(varchar(6), prideduc), 1,4) + '/' + substring(convert(varchar(6), prideduc), 5,2) + '/28' )) as PlazoFaltante" _
'       & " from reg_creditos where id_solicitud = " & txtOperacion
' rs.Open strSQL, glogon.Conection, adOpenStatic
'    lblCuotaR.Caption = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), rs!PlazoFaltante, vInteres)
' rs.Close


       lngFecha = lblFecUltMovR.Caption
       If lngFecha < vPrideduc Then lngFecha = vPrideduc
      
       ProcesosTmp = vPrideduc
       iPlazoRst = 1
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       
       If iPlazoRst <= 0 Then iPlazoRst = 1
       
       curCuota = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), iPlazoRst, vInteres)
       lblCuotaR.Caption = Format(curCuota, "Standard")
Else
  lblCuotaR.Caption = lblCuota.Caption
End If

End Sub

Private Sub optAbono_Click(Index As Integer)
Dim curInteres As Currency, vFecha As Date
Dim vProceso As Long

If fxVerificaMorosidad(vOperacion) Then
  txtCuotas = 0
  MsgBox "No se Pueden Aplicar Abonos, Esta operación se encuentra morosa", vbInformation
  vOperacion = 0
  Call sbLimpiaDatos
  Exit Sub
End If

fraDatosAbono.Enabled = True
chkRecalculaCuota.Enabled = False
chkRecalculaCuota.Value = vbUnchecked

'&H00C0FFC0&
txtTotalPagar.BackColor = &HC0FFC0
txtCuotas.BackColor = txtTotalPagar.BackColor

txtTotalPagar.Locked = True
txtCuotas.Enabled = False

dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False


Select Case Index

 Case 0 'Ordinario
   lblDatosAnticipo.Caption = 0
   txtDatosAmortiza = 0
   txtCuotas.Enabled = True
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
   txtCuotas.SetFocus
   
   txtCuotas.BackColor = vbWhite
   ModuloCajas.mAplicarSF = 0
 Case 1 'Extraordinario
 
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para Ab.Extraordinario:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
 
   txtCuotas = 0
   lblDatosInteres.Caption = 0
   lblDatosAnticipo.Caption = 0
   txtDatosAmortiza = 0
   
   txtCuotas.Enabled = False
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus
   
   chkRecalculaCuota.Enabled = True
   ModuloCajas.mAplicarSF = 1

   
   
Case 2 'Cancelacion
   'Le Calcula los intereses del proceso mensual + el saldo
   
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para cancelación...:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   txtDatosAmortiza = 0
   txtCuotas.Enabled = False
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
'   txtCuotas.SetFocus
   
   txtDatosAmortiza = lblSaldo.Caption
   
   'Cobra intereses del mes, pero verificar la fecha de proceso que sea igual
   'o menor
   vFecha = dtpFechaCancelacion.Value
   vProceso = Year(vFecha) & Format(Month(vFecha), "00")
   
   
'   '1er Paso de Validacion de Pago de Intereses
'   'Que la fecha de Proceso sea mayor al ultumo movimiento
'   If vProceso > CLng(lblFecUltMov.Caption) Then
'     curInteres = (CCur(lblSaldo.Caption) * vInteres / 36000) * Day(vFecha)
'   Else
'     curInteres = 0
'   End If
'
'   '2do Paso de Validacion de Pago de Intereses
'   'Que la fecha de Primer Deduccion sea mayor al ultimo abono (No ha iniciado plan de pago)
'   If curInteres > 0 And (vPrideduc > vProceso Or vPrideduc > CLng(lblFecUltMov.Caption)) Then
'     curInteres = 0
'   End If
   
   
   If (vProceso >= vPrideduc) And (vProceso > CLng(lblFecUltMov.Caption)) Then
     curInteres = (CCur(lblSaldo.Caption) * vInteres / 36000) * Day(vFecha)
   Else
     curInteres = 0
   End If
   
   
   '3er Paso de Validacion de Pago de Intereses
   'Verifica que no sea un credito del mismo mes
   If curInteres > 0 And Month(CDate(lblSaldo.Tag)) = Month(vFecha) _
        And Year(CDate(lblSaldo.Tag)) = Year(vFecha) Then
      curInteres = 0
   End If
   
   If vRetencion Then
      lblDatosAnticipo.Caption = "0.00"
   Else
      lblDatosAnticipo.Caption = Format(CCur(lblSaldo.Caption) * (CCur(lblDatosAnticipo.Tag) / 100), "Standard")
   End If
   
   lblDatosInteres.Caption = Format(curInteres, "Standard")
   txtTotalPagar.Text = Format(CCur(txtDatosAmortiza) + curInteres + CCur(lblDatosAnticipo.Caption), "Standard")
   ModuloCajas.mAplicarSF = 0
End Select


lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - CCur(txtDatosAmortiza), "Standard")
Call RefrescaTags(Me)

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtNombre = fxNombre(txtCedula)
  If txtCodigo <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtCodigo.SetFocus
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtCodigo = UCase(txtCodigo)
  lblDescripcion.Caption = fxDescribeCodigo(txtCodigo)
  If txtCedula <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtOperacion.SetFocus
End If

End Sub

Private Sub txtCuotas_Change()
Dim curSaldo As Currency, curAmortiza As Currency, curInteres As Currency
Dim curTmpAmortiza As Currency, curTmpInteres As Currency, i As Integer
Dim lngFecha As Long, lngCuotas As Long, lngCuotaMaxima As Long


Dim iDias As Integer, vFecha As Date, curCuota As Currency, iPlazoRst As Integer, ProcesosTmp As Long

On Error Resume Next

If txtCuotas = "" Or Not IsNumeric(txtCuotas.Text) Or txtCuotas.Text = "0" Then
 lngCuotas = 1
Else
 lngCuotas = txtCuotas
End If

If txtOperacion.Text = "" Then Exit Sub


ReDim pDatos(lngCuotas, 5) As Currency


lngFecha = CLng(lblFecUltMov.Caption)

If Not vRetencion Then
    curSaldo = vSaldoMes
Else
  'En las retenciones hay que proyectar el saldo del mes
  curSaldo = ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption))
End If

curAmortiza = 0
curInteres = 0
curCuota = lblCuota.Caption


If lngFecha < vPrideduc Then lngFecha = fxFechaProcesoAnterior(vPrideduc)

If vBaseCalculo = "01" Then
    For i = 1 To lngCuotas
    '360 / 360
        If curSaldo > 0 Then
          
          lngCuotaMaxima = i
          curTmpInteres = (curSaldo * vInteres) / 1200
          curTmpAmortiza = CCur(lblCuota.Caption) - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
        
          pDatos(i, 1) = curTmpInteres
          pDatos(i, 2) = curTmpAmortiza
          pDatos(i, 3) = lngFecha
          pDatos(i, 4) = curSaldo
          pDatos(i, 5) = curCuota
        
        End If
        
        If curSaldo < 0 Then
            pDatos(i, 1) = 0
            pDatos(i, 2) = curSaldo
            pDatos(i, 3) = lngFecha
            pDatos(i, 4) = 0
            pDatos(i, 5) = curSaldo
           
           curAmortiza = curAmortiza + curSaldo
           curSaldo = 0
        End If
     Next i
 
 Else
   '365 / 360
   
       'Calcula el Plazo Restante
       ProcesosTmp = vPrideduc
       iPlazoRst = 0
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       
       'Saca el formato fecha del ultimo movimiento para calculo de dias
       vFecha = Mid(CStr(lngFecha), 1, 4) & "/" & Right(CStr(lngFecha), 2) & "/01"
       
       For i = 1 To lngCuotas
          lngCuotaMaxima = i

          If iPlazoRst = 1 Or iPlazoRst = vPlazo Then
            iDias = 30
          Else
            iDias = fxMesDias(Month(vFecha), Year(vFecha))
          End If
        
          curTmpInteres = curSaldo * (vInteres / 100) * iDias / 360
          curTmpAmortiza = curCuota - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
          vFecha = DateAdd("m", 1, vFecha)
          
          pDatos(i, 1) = curTmpInteres
          pDatos(i, 2) = curTmpAmortiza
          pDatos(i, 3) = lngFecha
          pDatos(i, 4) = curSaldo
          pDatos(i, 5) = curCuota
          
          
          iPlazoRst = iPlazoRst - 1
          curCuota = fxCalcula_Cuota(CDbl(curSaldo), iPlazoRst, vInteres)
          
       Next i
    
   
 End If 'Base

lblDatosInteres.Caption = Format(curInteres, "Standard")
txtDatosAmortiza = Format(curAmortiza, "Standard")
lblFecUltMovR.Caption = lngFecha
lblCuotaR.Caption = Format(curCuota, "Standard")

If Not vRetencion Then 'El proceso nuevo de retenciones no toca los saldos
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard")
End If

lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + curAmortiza, "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + curInteres, "Standard")

If lngCuotas > lngCuotaMaxima Then txtCuotas = lngCuotaMaxima

End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub txtDatosAmortiza_Change()
On Error Resume Next
If Not vRetencion Then
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - CCur(txtDatosAmortiza), "Standard")
Else
    lblSaldoR.Caption = lblCuota.Caption
End If
lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + CCur(txtDatosAmortiza), "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + CCur(lblDatosInteres), "Standard")
txtTotalPagar.Text = Format(CCur(txtDatosAmortiza) + CCur(lblDatosInteres.Caption), "Standard")
End Sub

Private Sub txtDatosAmortiza_GotFocus()
On Error Resume Next
txtDatosAmortiza = CCur(txtDatosAmortiza)
End Sub

Private Sub txtDatosAmortiza_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
 cboTipoPago.SetFocus
End If
End Sub

Private Sub txtDatosAmortiza_LostFocus()
On Error Resume Next
txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyPress(vbKeyReturn)
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 vOperacion = txtOperacion
 Call sbConsultaOperacion
End If
End Sub

'Private Function fxDocumentoAbono(strTipo As String, vTipoDoc As String) As Long
'Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
'Dim lngRecibo As Long, strCliente As String, vCuenta As String
'
'vCuenta = Trim(fxDocumentoCuenta(vTipoDoc))
'
'If vAseDocValido = False Then
'  Exit Function
'End If
'
'lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
'fxDocumentoAbono = lngRecibo
'
'
'If txtProceso.Tag = "J" Then
'      strSQL = "select ctaCamort as ctaAmortiza,ctaCintc as ctaintc from catalogo"
'Else
'    If UCase(lblOpex.Caption) = "OPEX" Then
'      strSQL = "select ctaOamort as ctaAmortiza,ctaOintc as ctaintc from catalogo"
'    Else
'      strSQL = "select ctaNamort as ctaAmortiza,ctaNintc as ctaintc from catalogo"
'    End If
'End If
'strSQL = strSQL & " where codigo = '" & txtCodigo & "'"
'
'
'rs.Open strSQL, glogon.Conection, adOpenStatic
'
'strLinea(1) = "# Cuotas          " & txtCuotas
'strLinea(2) = "Saldo Anterior    " & lblSaldo.Caption
'strLinea(3) = "Interes Corriente " & lblDatosInteres.Caption
'strLinea(4) = "Amortizacion      " & txtDatosAmortiza
'strLinea(5) = "Saldo Actual      " & IIf(vRetencion, lblSaldo.Caption, lblSaldoR.Caption)
'strLinea(6) = "Cargo x Anticipo  " & lblDatosAnticipo.Caption
'strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " L.:" & txtCodigo & "-" & UCase(lblOpex.Caption)
'strLinea(8) = "Descripción       " & fxDescribeCodigo(txtCodigo)
'strLinea(9) = "Usuario           " & glogon.Usuario
'strLinea(10) = "Proc. Retencion  " & IIf(vRetencion, "SI", "NO")
'
'strCliente = Trim(txtCedula) & " - " & Trim(txtNombre)
'strCliente = Mid(strCliente, 1, 45)
'
'strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
'        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,dp)" _
'        & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") & "','" & strCliente & "','" _
'        & strTipo & " Op:" & vOperacion & "'," & CCur(txtTotalPagar.Text) & ",'" & glogon.Usuario & "','P','" _
'        & fxTipoPago(cboTipoPago.Text) & "','" & strLinea(1) & "','" _
'        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
'        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
'        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
'        & vAseDocDetalle & "','" & vAseDocDeposito & "')"
'glogon.Conection.Execute strSQL
'
''ASIENTO
'If CCur(lblDatosInteres.Caption) > 0 Then
'  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
'          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaintc) & "'," & CCur(lblDatosInteres.Caption) & ",'H')"
'  glogon.Conection.Execute strSQL
'End If
'
'If CCur(lblDatosAnticipo.Caption) > 0 Then
'  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
'          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(fxCRDParametro("17")) & "'," & CCur(lblDatosAnticipo.Caption) & ",'H')"
'  glogon.Conection.Execute strSQL
'End If
'
'
'If CCur(txtDatosAmortiza) > 0 Then
'  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
'          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & CCur(txtDatosAmortiza) & ",'H')"
'  glogon.Conection.Execute strSQL
'End If
'
'
'If CCur(lblDatosInteres.Caption) + CCur(txtDatosAmortiza) + CCur(lblDatosAnticipo.Caption) > 0 Then
'  strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
'          & " values(" & lngRecibo & ",'" & vTipoDoc & "','" & vCuenta & "'," & CCur(lblDatosInteres.Caption) + CCur(txtDatosAmortiza) + CCur(lblDatosAnticipo.Caption) & ",'D')"
'  glogon.Conection.Execute strSQL
'End If
'
'rs.Close
'
'End Function

Private Function fxDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pComprobante As String _
                                , pConcepto As String, pCuenta As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency

vCuenta = pCuenta

lngRecibo = CLng(pComprobante)

fxDocumentoAbono = lngRecibo

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
rs.Open strSQL, glogon.Conection, adOpenStatic

      
curIntC = lblDatosInteres.Caption
curIntM = 0
curAmortiza = txtDatosAmortiza.Text
curCargo = lblDatosAnticipo.Caption


strLinea(1) = "# Cuotas          " & txtCuotas
strLinea(2) = "Saldo Anterior    " & lblSaldo.Caption
strLinea(3) = "Interes Corriente " & lblDatosInteres.Caption
strLinea(4) = "Amortizacion      " & txtDatosAmortiza.Text
strLinea(5) = "Saldo Actual      " & IIf(vRetencion, lblSaldo.Caption, lblSaldoR.Caption)
strLinea(6) = "Cargo x Anticipo  " & lblDatosAnticipo.Caption
strLinea(7) = "Cargos [General]  " & "0.00"

strLinea(8) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " L.:" & txtCodigo & "-" & UCase(lblOpex.Caption)
strLinea(9) = "Descripción       " & lblDescripcion.Caption
strLinea(10) = "Proc. Retencion  " & IIf(vRetencion, "SI", "NO")

If dtpFechaCancelacion.Enabled Then
    strLinea(10) = "Fecha Real Abono " & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
End If

If GLOBALES.SysDocVersion = 1 Then

        strCliente = Trim(txtCedula) & " - " & Trim(txtNombre)
        strCliente = Mid(strCliente, 1, 45)
        
        strSQL = "insert ase_documentos(id_documento,tipo,fecha,cliente,concepto,monto,usuario,estado,tipo_pago" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,dp)" _
                & " values(" & lngRecibo & ",'" & pTipoDoc & "',getdate(),'" & strCliente & "','" _
                & pTipoAbono & " Op:" & vOperacion & "'," & CCur(txtTotalPagar.Text) & ",'" & glogon.Usuario & "','P','" _
                & fxTipoPago(cboTipoPago.Text) & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "')"
        glogon.Conection.Execute strSQL
        
        'ASIENTO
        If CCur(lblDatosInteres.Caption) > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(rs!ctaintc) & "'," & CCur(lblDatosInteres.Caption) & ",'H')"
          glogon.Conection.Execute strSQL
        End If
        
        If CCur(lblDatosAnticipo.Caption) > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(fxCRDParametro("17")) & "'," & CCur(lblDatosAnticipo.Caption) & ",'H')"
          glogon.Conection.Execute strSQL
        End If
        
        Select Case CCur(txtDatosAmortiza)
          Case Is < 0
                strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                        & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & Abs(CCur(txtDatosAmortiza)) & ",'D')"
                glogon.Conection.Execute strSQL
          Case Is > 0
                strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                        & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & Trim(rs!ctaamortiza) & "'," & CCur(txtDatosAmortiza) & ",'H')"
                glogon.Conection.Execute strSQL
        End Select
        
        
        
        If CCur(lblDatosInteres.Caption) + CCur(txtDatosAmortiza) + CCur(lblDatosAnticipo.Caption) > 0 Then
          strSQL = "insert ase_asientos(id_documento,tipo,recas_cuenta,recas_monto,recas_debehaber)" _
                  & " values(" & lngRecibo & ",'" & pTipoDoc & "','" & vCuenta & "'," & CCur(lblDatosInteres.Caption) + CCur(txtDatosAmortiza) + CCur(lblDatosAnticipo.Caption) & ",'D')"
          glogon.Conection.Execute strSQL
        End If
        
Else
  'Control de Documentos v2
   
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
                & " values('" & lngRecibo & "','" & pTipoDoc & "',getdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
                & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
                & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "')"
        glogon.Conection.Execute strSQL
        
        'ASIENTO
        If curIntC + curIntM + curAmortiza + curCargo > 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC + curIntM + curCargo + curAmortiza & ",'D','" & rs!cod_divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
                 & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          glogon.Conection.Execute strSQL
          'Crea asientos para las formas de pago
          
          strSQL = "exec dbo.spCajasPorcesos '" & ModuloCajas.mTiquete & "','" & pTipoDoc & "','" & pConcepto & "'" _
             & ",'" & lngRecibo & "'," & CCur(curIntC + curIntM + curCargo + curAmortiza) & "" _
             & " ,'" & ModuloCajas.mOficina & "','" & GLOBALES.gOficinaCentroCosto & "'," & GLOBALES.gEnlace & "" _
             & "," & ModuloCajas.mApertura & ",'" & ModuloCajas.mCaja & "','" & Trim(txtCedula.Text) & "'"
             
          glogon.Conection.Execute strSQL
          
          'en caso que se haya utilizado saldos a favor
          If ModuloCajas.mCasosSFAplicados > 0 Then Call sbActualizaSaldosFavor(ModuloCajas.mCasosSFAplicados, txtCedula.Text)
          
      End If
      
      
'        If curIntC > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
'
'        If curIntM > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntM & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
'
'        If curCargo > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curCargo & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!CtaCargos _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
'
'
'        If curAmortiza > 0 Then
'          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza & ",'C','" & rs!cod_divisa _
'                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
'                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          glogon.Conection.Execute strSQL
'        End If
  

      
End If
rs.Close


End Function




Private Sub txtTotalPagar_GotFocus()

On Error GoTo vError
 txtTotalPagar.Text = CCur(txtTotalPagar.Text)
vError:

End Sub

Private Sub txtTotalPagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If CmdAbono.Enabled Then
    CmdAbono.SetFocus
  Else
    cmdReporte.SetFocus
  End If
End If

End Sub

Private Sub txtTotalPagar_LostFocus()
Dim vFecha As Date, vProceso As Long
Dim curInteres As Currency, curAmortiza As Currency
Dim i As Integer
 
On Error GoTo vError
 
'ExtraOrdinario
If optAbono.Item(1).Value = True Then
   'Cobra intereses del mes, pero verificar la fecha de proceso que sea igual o menor
  
   vFecha = dtpFechaCancelacion.Value
   vProceso = Year(vFecha) & Format(Month(vFecha), "00")
   
   If vProceso >= vPrideduc And vProceso > CLng(lblFecUltMov.Caption) Then
     curInteres = (CCur(txtTotalPagar.Text) * vInteres / 36000) * Day(vFecha)
   Else
     curInteres = 0
   End If
   
'   '2do Paso de Validacion de Pago de Intereses
'   'Que la fecha de Primer Deduccion sea mayor al ultimo abono (No ha iniciado plan de pago)
'   If curInteres > 0 And (vPrideduc > vProceso Or vPrideduc > CLng(lblFecUltMov.Caption)) Then
'     curInteres = 0
'   End If
   
   
   'Verifica que no sea un credito del mismo mes
   If curInteres > 0 And Month(CDate(lblSaldo.Tag)) = Month(vFecha) _
        And Year(CDate(lblSaldo.Tag)) = Year(vFecha) Then
      curInteres = 0
   End If
   
   'Se re-calculan intereses para ajustar y relacionar segun porcion amortizada
   'Previamente sobre el monto a cancelar
   
   If curInteres > 0 Then
      'Hacer 10 aproximaciones
      For i = 1 To 10
            curAmortiza = CCur(txtTotalPagar.Text) - curInteres
            curInteres = (curAmortiza * vInteres / 36000) * Day(vFecha)
      Next i
   End If
   
   lblDatosInteres.Caption = Format(curInteres, "Standard")
   txtDatosAmortiza.Text = Format(CCur(txtTotalPagar.Text) - curInteres, "Standard")

End If


txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text), "Standard")

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


