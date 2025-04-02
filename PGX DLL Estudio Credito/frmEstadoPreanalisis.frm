VERSION 5.00
Begin VB.Form frmPreaEstadoPreanalisis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estado del preanalisis"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OpSolicitado 
      Caption         =   "Solicitado"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Tag             =   "S"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmb_Aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1065
   End
   Begin VB.CommandButton Cmd_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1065
   End
   Begin VB.OptionButton OpDenegado 
      Caption         =   "Denegado"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Tag             =   "D"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton OpAutorizado 
      Caption         =   "Autorizado"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Tag             =   "A"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   3660
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   3630
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label 
      Caption         =   "Gestión del preanalisis."
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1725
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmEstadoPreanalisis.frx":0000
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmPreaEstadoPreanalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_estadoPreanalisis As String
Public m_codPreanalisis As String
Private clsEntidad As New ASE_PreAnalisis.clsEntidad

Private Sub cmb_Aceptar_Click()
Call sbGuardar
End Sub

Private Sub Cmd_Cancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    If (m_estadoPreanalisis = "" Or m_estadoPreanalisis = "S") Then
       OpSolicitado.Value = True
    ElseIf m_estadoPreanalisis = "A" Then
        OpAutorizado.Value = True
    Else
        OpDenegado.Value = True
    End If
End Sub
Private Sub sbGuardar()
Dim StrUpdate As String
Dim StrSet As String
On Error GoTo VError
StrUpdate = "Update CRD_PREA_PREANALISIS SET "

If OpAutorizado.Value = True Then
    StrSet = "ESTADO = " & fxFormatearValor(OpAutorizado.Tag, Caracter)
    StrSet = StrSet & ", USUARIO_GESTION = " & fxFormatearValor(glogon.usuario, Caracter) & ",  FECHA_GESTION = " & fxFormatearValor(clsEntidad.fxTraerFechaServidor, Fecha)
ElseIf OpDenegado.Value = True Then
    StrSet = "ESTADO = " & fxFormatearValor(OpDenegado.Tag, Caracter)
    StrSet = StrSet & ", USUARIO_GESTION = " & fxFormatearValor(glogon.usuario, Caracter) & ",  FECHA_GESTION = " & fxFormatearValor(clsEntidad.fxTraerFechaServidor, Fecha)
ElseIf OpSolicitado.Value = True Then
    StrSet = "ESTADO = " & fxFormatearValor(OpSolicitado.Tag, Caracter)
End If

StrUpdate = StrUpdate & StrSet & " where COD_PREANALISIS = " & fxFormatearValor(m_codPreanalisis, Caracter)

If execSql(StrUpdate, False) Then
    MsgBox "La información fue actualizada correctamente.", vbInformation, gTituloMsg
    Unload Me
End If

salir:
    Exit Sub
VError:
    MsgBox Err.Description
    Resume salir
End Sub

