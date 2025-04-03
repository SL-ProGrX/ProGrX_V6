VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmTES_Bloqueos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo de Solicitudes"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   Icon            =   "frmTES_Bloqueos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10860
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   10575
      _Version        =   1441793
      _ExtentX        =   18653
      _ExtentY        =   11880
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
      Item(0).Caption =   "Bloqueo"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "Label1(1)"
      Item(0).Control(1)=   "Label1(0)"
      Item(0).Control(2)=   "txtCodigo"
      Item(0).Control(3)=   "txtDatos"
      Item(0).Control(4)=   "txtBloqueo"
      Item(0).Control(5)=   "cmdDesBloquear"
      Item(0).Control(6)=   "cmdBloquear"
      Item(0).Control(7)=   "Label1(2)"
      Item(0).Control(8)=   "GroupBox1(3)"
      Item(1).Caption =   "Desbloqueo"
      Item(1).ControlCount=   13
      Item(1).Control(0)=   "chkFechas"
      Item(1).Control(1)=   "txtSolicitudCorte"
      Item(1).Control(2)=   "txtSolicitudInicial"
      Item(1).Control(3)=   "lsw"
      Item(1).Control(4)=   "chkMarcas"
      Item(1).Control(5)=   "btnBuscar"
      Item(1).Control(6)=   "dtpCorte"
      Item(1).Control(7)=   "Label3(1)"
      Item(1).Control(8)=   "chkSolicitudes"
      Item(1).Control(9)=   "dtpInicio"
      Item(1).Control(10)=   "Label1(3)"
      Item(1).Control(11)=   "cmdDesBloqueoLote"
      Item(1).Control(12)=   "lblX"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4812
         Left            =   -70000
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   10572
         _Version        =   1441793
         _ExtentX        =   18648
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkMarcas 
         Height          =   204
         Left            =   -69640
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   204
         _Version        =   1441793
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   672
         Left            =   -62680
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1182
         _StockProps     =   79
         Caption         =   "&Buscar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bloqueos.frx":6852
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -66640
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -68080
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtSolicitudInicial 
         Height          =   312
         Left            =   -68080
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.FlatEdit txtSolicitudCorte 
         Height          =   312
         Left            =   -66640
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   432
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
      Begin XtremeSuiteControls.FlatEdit txtDatos 
         Height          =   3315
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15055
         _ExtentY        =   5847
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdDesBloquear 
         Height          =   675
         Left            =   6600
         TabIndex        =   14
         Top             =   5940
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   1191
         _StockProps     =   79
         Caption         =   "&Des-bloquear"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bloqueos.frx":7270
      End
      Begin XtremeSuiteControls.PushButton cmdBloquear 
         Height          =   675
         Left            =   8280
         TabIndex        =   15
         Top             =   5940
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   1191
         _StockProps     =   79
         Caption         =   "&Bloquear"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bloqueos.frx":7A4E
      End
      Begin XtremeSuiteControls.PushButton cmdDesBloqueoLote 
         Height          =   672
         Left            =   -61360
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   1185
         _StockProps     =   79
         Caption         =   "&Des-bloquear"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bloqueos.frx":83E3
      End
      Begin XtremeSuiteControls.FlatEdit txtBloqueo 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   4440
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15049
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   324
         Left            =   -64960
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1044
         _Version        =   1441793
         _ExtentX        =   1841
         _ExtentY        =   572
         _StockProps     =   79
         Caption         =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkSolicitudes 
         Height          =   324
         Left            =   -64960
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   1044
         _Version        =   1441793
         _ExtentX        =   1841
         _ExtentY        =   572
         _StockProps     =   79
         Caption         =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   975
         Index           =   3
         Left            =   360
         TabIndex        =   24
         Top             =   4800
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Notas: "
         ForeColor       =   4210752
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   615
            Left            =   1080
            TabIndex        =   25
            Top             =   240
            Width           =   8535
            _Version        =   1441793
            _ExtentX        =   15055
            _ExtentY        =   1085
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
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueo"
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
         Index           =   2
         Left            =   360
         TabIndex        =   21
         Top             =   4440
         Width           =   1335
      End
      Begin XtremeShortcutBar.ShortcutCaption lblX 
         Height          =   372
         Left            =   -70000
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   10572
         _Version        =   1441793
         _ExtentX        =   18648
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Solicitudes Bloqueadas"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitud"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   540
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1020
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitudes"
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
         Left            =   -69160
         TabIndex        =   7
         Top             =   900
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas"
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
         Index           =   1
         Left            =   -69160
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   972
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   0
      Top             =   7965
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bloqueo de Solicitudes"
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
      Height          =   312
      Index           =   10
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   3972
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmTES_Bloqueos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnBuscar_Click()
  Call sbBuscar
End Sub

Private Sub cmdBloquear_Click()
Dim strSQL As String

On Error GoTo vError

If txtDatos.Tag <> "S" Then
  MsgBox "Bloqueo no procede verifique...", vbExclamation
  Exit Sub
End If

strSQL = "update Tes_Transacciones set user_hold = '" & glogon.Usuario & "',fecha_hold = dbo.MyGetdate()" _
       & " where nsolicitud = " & txtCodigo
Call ConectionExecute(strSQL)

Call sbTesBitacoraEspecial(txtCodigo.Text, "05", "")

Call Bitacora("Aplica", "Bloqueo de Solicitud : " & txtCodigo)

MsgBox "Solicitud Bloqueada Satisfactoriamente...", vbInformation

Call txtCodigo_LostFocus

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDesBloqueo(vSolicitud As Long)
Dim strSQL As String

On Error GoTo vError

strSQL = "update Tes_Transacciones set user_hold = Null,fecha_hold = null" _
       & " where nsolicitud = " & vSolicitud
Call ConectionExecute(strSQL)

Call sbTesBitacoraEspecial(vSolicitud, "06", "")

Call Bitacora("Aplica", "Desbloqueo de Solicitud : " & vSolicitud)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdDesBloquear_Click()
Dim strSQL As String

On Error GoTo vError

If txtDatos.Tag <> "S" Then
  MsgBox "Des-Bloqueo no procede verifique...", vbExclamation
  Exit Sub
End If

Call sbDesBloqueo(txtCodigo)

MsgBox "Solicitud Des-Bloqueada Satisfactoriamente...", vbInformation

Call txtCodigo_LostFocus

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdDesBloqueoLote_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError


Me.MousePointer = vbHourglass

PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1

PrgBar.Visible = True

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
     Call sbDesBloqueo(lsw.ListItems.Item(i).Text)
  End If
  
  PrgBar.Value = PrgBar.Value + 1
Next i
   
PrgBar.Visible = False
Call sbBuscar


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSolicitudInicial.SetFocus
vError:
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
vError:
End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio

Call chkFechas_Click
Call chkSolicitudes_Click

Call Formularios(Me)
Call RefrescaTags(Me)

tcMain.Item(0).Selected = True

End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkMarcas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
 lsw.ListItems.Item(i).Checked = chkMarcas.Value
Next i

End Sub

Private Sub chkSolicitudes_Click()

If chkSolicitudes.Value = vbChecked Then
   txtSolicitudInicial.Enabled = False
Else
   txtSolicitudInicial.Enabled = True
End If

txtSolicitudCorte.Enabled = txtSolicitudInicial.Enabled

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub txtCodigo_Change()
txtDatos = ""
txtDatos.Tag = "N"
txtBloqueo = ""

cmdBloquear.Enabled = False
cmdDesBloquear.Enabled = False
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDatos.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

txtDatos = ""
txtDatos.Tag = "N"
txtBloqueo = ""

strSQL = "select C.*,B.descripcion as BancoX,X.descripcion as ConceptoX" _
       & ",U.descripcion as UnidadX,T.descripcion as TipoX" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_Banco = B.id_Banco" _
       & " inner join tes_conceptos X on C.cod_concepto = X.cod_concepto" _
       & " inner join CntX_unidades U on C.cod_unidad = U.cod_unidad and U.cod_contabilidad = " & GLOBALES.gEnlace _
       & " inner join tes_Tipos_doc T on C.Tipo = T.tipo" _
       & " where C.estado = 'P' and C.nsolicitud = " & txtCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
   txtDatos.Tag = "S"
   
   txtDatos = "Código ...........: " & vbTab & rs!Codigo & vbCrLf
   txtDatos = txtDatos & "Beneficiario .....: " & vbTab & rs!Beneficiario & vbCrLf
   txtDatos = txtDatos & "Banco ............: " & vbTab & rs!BancoX & vbCrLf
   txtDatos = txtDatos & "Monto ............: " & vbTab & Format(rs!Monto, "Standard") & vbCrLf & vbCrLf
   
   txtDatos = txtDatos & "Tipo .............: " & vbTab & rs!TipoX & vbCrLf
   txtDatos = txtDatos & "Concepto..........: " & vbTab & rs!ConceptoX & vbCrLf
   txtDatos = txtDatos & "Unidad............: " & vbTab & rs!UnidadX & vbCrLf & vbCrLf
   
   txtDatos = txtDatos & "Fecha ............: " & vbTab & Format(rs!fecha_solicitud, "dd/mm/yyyy") & vbCrLf
   txtDatos = txtDatos & "Usuario ..........: " & vbTab & rs!user_solicita & vbCrLf & vbCrLf
   
   txtDatos = txtDatos & "Detalle...........: " & vbTab & Trim(rs!Detalle1 & "") & vbCrLf
   txtDatos = txtDatos & "                    " & vbTab & Trim(rs!Detalle2 & "") & vbCrLf
   txtDatos = txtDatos & "                    " & vbTab & Trim(rs!Detalle3 & "") & vbCrLf
   txtDatos = txtDatos & "                    " & vbTab & Trim(rs!Detalle4 & "") & vbCrLf
   txtDatos = txtDatos & "                    " & vbTab & Trim(rs!Detalle5 & "") & vbCrLf
   
   If Not IsNull(rs!fecha_hold) Then
      txtBloqueo.Text = rs!user_hold & " - " & rs!fecha_hold
      cmdBloquear.Enabled = False
      cmdDesBloquear.Enabled = True
   Else
      cmdBloquear.Enabled = True
      cmdDesBloquear.Enabled = False
   End If

Else
  MsgBox "No se encontró número de solicitud, o no se encuentra pendiente", vbExclamation

 
End If
rs.Close

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtSolicitudInicial_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSolicitudCorte.SetFocus
vError:
End Sub


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass



strSQL = "select C.nsolicitud,C.codigo,C.beneficiario,C.monto,C.fecha_solicitud" _
       & ",C.fecha_Hold,B.descripcion as BancoX,C.Tipo" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_Banco = B.id_Banco" _
       & " where user_hold is not null"
       
If chkFechas.Value = vbUnchecked Then
   strSQL = strSQL & " and fecha_hold between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

If chkSolicitudes.Value = vbUnchecked Then
  If Trim(txtSolicitudCorte) = Empty Or Trim(txtSolicitudInicial) = Empty Then
    MsgBox "Debe indicar la solicitud inicial y la final..", vbCritical
    Exit Sub
  Else
   strSQL = strSQL & " and (nsolicitud >= " & CCur(txtSolicitudInicial) & " and nsolicitud <=" _
          & CCur(txtSolicitudCorte) & ")"
  End If
End If

lsw.ListItems.Clear
With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Solicitud", 1400
    .Add , , "Código", 1400, vbCenter
    .Add , , "Beneficiario", 3200
    .Add , , "Monto", 1400, vbRightJustify
    .Add , , "Fec.Sol", 1200, vbCenter
    .Add , , "Tipo", 900, vbCenter
    .Add , , "Cuenta", 3000
    .Add , , "Fec.Bloqueo", 1200, vbCenter
End With
lsw.Checkboxes = True

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!Beneficiario
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     itmX.SubItems(4) = Format(rs!fecha_solicitud, "yyyy-mm-dd")
     itmX.SubItems(5) = rs!Tipo
     itmX.SubItems(6) = rs!BancoX
     itmX.SubItems(7) = Format(rs!fecha_hold & "", "yyyy-mm-dd")
     
     itmX.Checked = chkMarcas.Value
     
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
