VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmSYS_APL_Integrador 
   Caption         =   "APL: Integrador"
   ClientHeight    =   9936
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   12312
   LinkTopic       =   "Form1"
   ScaleHeight     =   9936
   ScaleWidth      =   12312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox Box_Tramite 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12012
      _Version        =   1245185
      _ExtentX        =   21188
      _ExtentY        =   8700
      _StockProps     =   79
      Caption         =   "Solicitudes en Tramite"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkAtendidasSol 
         Height          =   492
         Left            =   6960
         TabIndex        =   10
         Top             =   600
         Width           =   1212
         _Version        =   1245185
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Atendidas?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin VB.Timer Timer_Inicial 
         Interval        =   1000
         Left            =   10320
         Top             =   240
      End
      Begin VB.Timer TimerX 
         Interval        =   60000
         Left            =   10680
         Top             =   240
      End
      Begin XtremeSuiteControls.PushButton btnRefrescar_Tramite 
         Height          =   492
         Left            =   8280
         TabIndex        =   3
         Top             =   600
         Width           =   1692
         _Version        =   1245185
         _ExtentX        =   2984
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Refrescar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmSYS_APL_Integrador.frx":0000
      End
      Begin FPSpreadADO.fpSpread vGrid_Tramite 
         Height          =   3492
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   11892
         _Version        =   524288
         _ExtentX        =   20976
         _ExtentY        =   6160
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
         MaxCols         =   504
         SpreadDesigner  =   "frmSYS_APL_Integrador.frx":098D
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado_Tramite 
         Height          =   312
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.ComboBox cboLinea_Tramite 
         Height          =   312
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   3972
         _Version        =   1245185
         _ExtentX        =   7006
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
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Línea:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.GroupBox Box_PreSolicitud 
      Height          =   4932
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   12012
      _Version        =   1245185
      _ExtentX        =   21188
      _ExtentY        =   8700
      _StockProps     =   79
      Caption         =   "Consultas"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnRefresh_Solicitud 
         Height          =   492
         Left            =   8280
         TabIndex        =   5
         Top             =   600
         Width           =   1692
         _Version        =   1245185
         _ExtentX        =   2984
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Refrescar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmSYS_APL_Integrador.frx":1483
      End
      Begin FPSpreadADO.fpSpread vGrid_Solicitud 
         Height          =   3492
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   11892
         _Version        =   524288
         _ExtentX        =   20976
         _ExtentY        =   6160
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
         MaxCols         =   499
         SpreadDesigner  =   "frmSYS_APL_Integrador.frx":1E10
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.CheckBox chkAtendidasCon 
         Height          =   492
         Left            =   6960
         TabIndex        =   11
         Top             =   600
         Width           =   1212
         _Version        =   1245185
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Atendidas?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboLinea_Solicitud 
         Height          =   312
         Left            =   2880
         TabIndex        =   14
         Top             =   720
         Width           =   3972
         _Version        =   1245185
         _ExtentX        =   7006
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
      Begin XtremeSuiteControls.ComboBox cboEstado_Solicitud 
         Height          =   312
         Left            =   600
         TabIndex        =   15
         Top             =   720
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Línea:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   2880
         TabIndex        =   8
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   1092
      End
   End
End
Attribute VB_Name = "frmSYS_APL_Integrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vDominio As String
Dim db As New ADODB.Connection
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub btnRefrescar_Tramite_Click()
Call sbLista_Tramite
End Sub

Private Sub btnRefresh_Solicitud_Click()
Call sbLista_Solicitudes
End Sub

Private Sub Form_Load()
On Error GoTo vError


strSQL = "select APL_Dominio from SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL)
vDominio = rs!APL_Dominio & ""
rs.Close

vPaso = True
cboEstado_Tramite.Clear
cboEstado_Tramite.AddItem "Recibidas"
cboEstado_Tramite.AddItem "Pendientes"
cboEstado_Tramite.AddItem "Autorizadas"
cboEstado_Tramite.AddItem "Denegadas"
cboEstado_Tramite.AddItem "Formalizadas"
cboEstado_Tramite.Text = "Recibidas"


cboEstado_Solicitud.Clear
cboEstado_Solicitud.AddItem "Recibidas"
cboEstado_Solicitud.AddItem "Pendientes"
cboEstado_Solicitud.AddItem "Autorizadas"
cboEstado_Solicitud.AddItem "Denegadas"
cboEstado_Solicitud.Text = "Recibidas"

cboLinea_Tramite.AddItem "TODAS"
cboLinea_Solicitud.AddItem "TODAS"

cboLinea_Tramite.Text = "TODAS"
cboLinea_Solicitud.Text = "TODAS"

vPaso = False


'Establece Conexion
strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=progrx.centralus.cloudapp.azure.com" _
       & ";Database=APL;APP=PGX_APL_Access;tcp:progrx.centralus.cloudapp.azure.com" _
       & "," & SIFGlobal.PuertosDisponibles & ";"
       
db.ConnectionString = strSQL
db.Open , "APL_Integrador", "m1t-1F1l$yJ3r!*.e#"

strSQL = "exec spProGrX_APL_Lineas '" & vDominio & "'"
rs.Open strSQL, db, adOpenStatic
Do While Not rs.EOF
 
    cboLinea_Tramite.AddItem rs!Item_Desc & ""
    cboLinea_Tramite.ItemData(cboLinea_Tramite.ListCount - 1) = CStr(rs!Item_Id)
  
    cboLinea_Solicitud.AddItem rs!Item_Desc & ""
    cboLinea_Solicitud.ItemData(cboLinea_Solicitud.ListCount - 1) = CStr(rs!Item_Id)
  
  
  rs.MoveNext
  
Loop
rs.Close

vGrid_Tramite.MaxCols = 18
vGrid_Solicitud.MaxCols = 14

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Resize()
On Error Resume Next

Box_Tramite.Width = Me.Width - 250
Box_Tramite.Height = (Me.Height - 500) / 2


Box_PreSolicitud.Top = Box_Tramite.Top + Box_Tramite.Height + 150
Box_PreSolicitud.Width = Box_Tramite.Width
Box_PreSolicitud.Height = Box_Tramite.Height


vGrid_Tramite.Width = Box_Tramite.Width - 150
vGrid_Tramite.Height = Box_Tramite.Height - (vGrid_Tramite.Top + 250)

vGrid_Solicitud.Width = vGrid_Tramite.Width
vGrid_Solicitud.Height = vGrid_Tramite.Height - 100


End Sub

Private Sub sbLista_Solicitudes()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spProGrX_APL_SOLICITUDES_Lista '" & vDominio & "',Null,'" _
       & Mid(cboEstado_Solicitud.Text, 1, 1) & "','" & glogon.Usuario & "','" & glogon.Usuario & "'"
       
If cboLinea_Solicitud.Text <> "TODAS" Then
   strSQL = strSQL & ",'" & cboLinea_Solicitud.ItemData(cboLinea_Solicitud.ListIndex) & "'"
Else
   strSQL = strSQL & ",Null"
End If

   strSQL = strSQL & ",Null, Null," & chkAtendidasCon.Value

  
With vGrid_Solicitud
  .MaxRows = 0
  
  rs.Open strSQL, db, adOpenStatic
  
  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    .Col = 2
    .Text = CStr(rs!Num_Solicitud)
    .Col = 3
    .Text = rs!Linea_Desc
    .Col = 4
    .Text = rs!Cedula
    .Col = 5
    .Text = rs!Nombre
    .Col = 6
    .Text = Format(rs!Monto, "Standard")
    .Col = 7
    .Text = CStr(rs!Plazo)
    .Col = 8
    .Text = rs!Institucion_desc
    .Col = 9
    .Text = rs!notas & ""
    .Col = 10
    .Text = rs!registro_Fecha & ""
    .Col = 11
    .Text = rs!registro_usuario & ""
    .Col = 12
    .Text = rs!Atiende_Fecha & ""
    .Col = 13
    .Text = rs!Atiende_Usuario & ""
    .Col = 14
    .Text = CStr(rs!Minutos)
    
    rs.MoveNext
  Loop
  rs.Close
  
End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbLista_Tramite()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spProGrX_APL_OPERACIONES_Lista '" & vDominio & "',Null,'" _
       & Mid(cboEstado_Tramite.Text, 1, 1) & "','" & glogon.Usuario & "','" & glogon.Usuario & "'"
       
If cboLinea_Tramite.Text <> "TODAS" Then
   strSQL = strSQL & ",'" & cboLinea_Tramite.ItemData(cboLinea_Tramite.ListIndex) & "'"
Else
   strSQL = strSQL & ",Null"
End If

   strSQL = strSQL & ",Null, Null," & chkAtendidasSol.Value

With vGrid_Tramite
  .MaxRows = 0
  
  rs.Open strSQL, db, adOpenStatic
  
  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    .Col = 2
    .Text = CStr(rs!APL_OPERACION)
    .Col = 3
    .Text = rs!Linea_Desc
    .Col = 4
    .Text = rs!Cedula
    .Col = 5
    .Text = rs!Nombre
    .Col = 6
    .Text = Format(rs!Factura_Monto, "Standard")
    .Col = 7
    .Text = CStr(rs!Plazo)
    
    .Col = 8
    .Text = Format(rs!Tasa, "Standard")
    .Col = 9
    .Text = Format(rs!Cuota, "Standard")
    
    .Col = 10 'Plan Desc
    .Text = rs!Plan_desc
    .Col = 11
    .Text = rs!Institucion_desc
    .Col = 12 'Factura
    .Text = rs!FACTURA_NUMERO & ""
    .Col = 13 'ProGrx Operacion
    .Text = CStr(rs!Operacion & "")
    .Col = 14
    .Text = rs!registro_Fecha & ""
    .Col = 15
    .Text = rs!registro_usuario & ""
    .Col = 16
    .Text = rs!Atiende_Fecha & ""
    .Col = 17
    .Text = rs!Atiende_Usuario & ""
    
    .Col = 18
    .Text = CStr(rs!Minutos)
    
    rs.MoveNext
  Loop
  rs.Close
  
End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Timer_Inicial_Timer()
Timer_Inicial.Interval = 0
Timer_Inicial.Enabled = False
Call sbLista_Solicitudes
Call sbLista_Tramite

End Sub

Private Sub TimerX_Timer()
Dim pSQL As String, pRs As New ADODB.Recordset
Dim pTrayIcon As XtremeSuiteControls.TrayIcon

On Error GoTo vError

Set pTrayIcon = frmContenedor.TrayIcon

pSQL = "exec spProGrX_APL_Monitor '" & vDominio & "'"

pRs.Open pSQL, db, adOpenStatic

If pRs!Solicitud + pRs!Tramite > 0 Then
   pTrayIcon.ShowBalloonTip 25, "APL: Notificación" _
             , "Tienes (" & pRs!Solicitud + pRs!Tramite & ") Solicitudes por Resolver" _
            , xtpToolTipIconInfo
End If
pRs.Close

Exit Sub

vError:

End Sub


Private Sub vGrid_Solicitud_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, i As Integer
Dim pOperacion As Long

On Error GoTo vError

 i = MsgBox("Desea Registrar como Atendiendo! el Caso?", vbYesNo)
 If i = vbNo Then
    Exit Sub
 End If
 vGrid_Solicitud.Row = Row
 vGrid_Solicitud.Col = 2
 pOperacion = vGrid_Solicitud.Text
 
 strSQL = "exec spProGrX_APL_SOLICITUDES_Atiende '" & vDominio & "'," & pOperacion & ",'" & glogon.Usuario & "'"
 db.Execute strSQL
  
 MsgBox "Caso atendido!", vbInformation
 
Exit Sub


vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub vGrid_Tramite_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, i As Integer
Dim pOperacion As Long

On Error GoTo vError

 i = MsgBox("Desea Registrar como Atendiendo! el Caso?", vbYesNo)
 If i = vbNo Then
    Exit Sub
 End If
 vGrid_Tramite.Row = Row
 vGrid_Tramite.Col = 2
 pOperacion = vGrid_Tramite.Text
 
 strSQL = "exec spProGrX_APL_OPERACIONES_Atiende '" & vDominio & "'," & pOperacion & ",'" & glogon.Usuario & "'"
 db.Execute strSQL

 MsgBox "Caso atendido!", vbInformation

Exit Sub


vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub
