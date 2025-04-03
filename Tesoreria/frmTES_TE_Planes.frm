VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmTES_TE_Planes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Planes Empresariales para Transferencias"
   ClientHeight    =   3588
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9792
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3588
   ScaleWidth      =   9792
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnScroll 
      Height          =   372
      Index           =   0
      Left            =   5520
      TabIndex        =   9
      Top             =   1440
      Width           =   372
      _Version        =   1245187
      _ExtentX        =   656
      _ExtentY        =   656
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   14
      Picture         =   "frmTES_TE_Planes.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnRegistra 
      Height          =   492
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Registrar"
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
      Picture         =   "frmTES_TE_Planes.frx":08D1
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtPlan 
      Height          =   372
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   3492
      _Version        =   1245187
      _ExtentX        =   6159
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "PLB"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConsecId 
      Height          =   372
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   3492
      _Version        =   1245187
      _ExtentX        =   6159
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConsecInt 
      Height          =   372
      Left            =   1920
      TabIndex        =   7
      Top             =   2880
      Width           =   3492
      _Version        =   1245187
      _ExtentX        =   6159
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnScroll 
      Height          =   372
      Index           =   1
      Left            =   5880
      TabIndex        =   10
      Top             =   1440
      Width           =   372
      _Version        =   1245187
      _ExtentX        =   656
      _ExtentY        =   656
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   14
      Picture         =   "frmTES_TE_Planes.frx":0FF8
   End
   Begin XtremeSuiteControls.PushButton btnBorra 
      Height          =   492
      Left            =   7800
      TabIndex        =   11
      Top             =   2760
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Eliminar"
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
      Picture         =   "frmTES_TE_Planes.frx":18C9
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   732
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Consecutivo Interno"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   732
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Número de Ultima Transferencia"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Plan"
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
   Begin XtremeShortcutBar.ShortcutCaption scBanco 
      Height          =   612
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9852
      _Version        =   1245187
      _ExtentX        =   17378
      _ExtentY        =   1080
      _StockProps     =   14
      Caption         =   "Banco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.21
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scCuenta 
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9852
      _Version        =   1245187
      _ExtentX        =   17378
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Cuenta Bancaria"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmTES_TE_Planes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnBorra_Click()


Call sbBorrar

End Sub

Private Sub btnRegistra_Click()
Dim vMensaje As String

vMensaje = ""

If Not IsNumeric(txtConsecId.Text) Then
    vMensaje = vMensaje & vbCrLf & " - El Consecutivo de la Transferencia no es válido!"
End If

If Not IsNumeric(txtConsecInt.Text) Then
    vMensaje = vMensaje & vbCrLf & " - El Consecutivo Interno para Transferencia no es válido!"
End If

If Len(vMensaje) = 0 Then
    Call sbGuardar
Else
    MsgBox vMensaje, vbExclamation
End If


End Sub

Private Sub btnScroll_Click(Index As Integer)

On Error GoTo vError

strSQL = "select Top 1 cod_Plan from TES_BANCO_PLANES_TE"

If Index = 1 Then
   strSQL = strSQL & " where id_banco = " & scCuenta.Tag & " and cod_Plan > '" & txtPlan & "' order by cod_Plan asc"
Else
   strSQL = strSQL & " where id_banco = " & scCuenta.Tag & " and cod_Plan < '" & txtPlan & "' order by cod_Plan desc"
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtPlan.Text = rs!cod_Plan
  Call sbConsulta
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

vModulo = 9


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbInicializa()

On Error GoTo vError

strSQL = "select B.ID_BANCO, B.COD_GRUPO, B.DESCRIPCION , B.DESC_CORTA" _
       & ", Bg.DESCRIPCION as 'Banco_Desc', Bg.DESC_CORTA as 'Banco_Desc_Corta'" _
       & "  from TES_BANCOS B inner join TES_BANCOS_GRUPOS Bg on B.COD_GRUPO = Bg.COD_GRUPO" _
       & " Where B.ID_Banco = " & GLOBALES.gTag

Call OpenRecordSet(rs, strSQL)

scBanco.Tag = rs!Cod_Grupo
scBanco.Caption = rs!Banco_Desc

scCuenta.Tag = rs!ID_BANCO
scCuenta.Caption = rs!Descripcion

txtPlan.Text = ""
txtConsecId.Text = "0"
txtConsecInt.Text = "0"

rs.Close

If GLOBALES.gTag2 <> "" Then
    txtPlan.Text = GLOBALES.gTag2
End If

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub


Private Sub sbConsulta()

On Error GoTo vError

If txtPlan.Text = "" Then Exit Sub

strSQL = "exec spTes_Planes_Consulta " & scCuenta.Tag & ",'" & txtPlan.Text & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtPlan.Text = rs!cod_Plan
    
    txtConsecId.Text = CStr(rs!NUMERO_TE)
    txtConsecInt.Text = CStr(rs!NUMERO_INTERNO)
End If

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbGuardar()

On Error GoTo vError

strSQL = "exec spTes_Planes_Registro " & scCuenta.Tag & ",'" & txtPlan.Text & "', " & CLng(txtConsecId.Text) _
       & ", " & CLng(txtConsecInt.Text) & ", '" & glogon.Usuario & "', 'A'"

Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Cta Id: " & scCuenta.Tag & ", Plan: " & txtPlan.Text & ", Consec Id: " _
        & txtConsecId.Text & ", Consec Interno: " & txtConsecInt.Text)

MsgBox "Plan Registrado Satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbBorrar()

On Error GoTo vError

strSQL = "exec spTes_Planes_Registro " & scCuenta.Tag & ",'" & txtPlan.Text & "', " & CLng(txtConsecId.Text) _
       & ", " & CLng(txtConsecInt.Text) & ", '" & glogon.Usuario & "', 'E'"

Call ConectionExecute(strSQL)

Call Bitacora("Elimina", "Cta Id: " & scCuenta.Tag & ", Plan: " & txtPlan.Text & ", Consec Id: " _
        & txtConsecId.Text & ", Consec Interno: " & txtConsecInt.Text)

MsgBox "Plan Eliminado Satisfactoriamente!", vbInformation

txtPlan.Text = ""
txtConsecId.Text = "0"
txtConsecInt.Text = "0"

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



