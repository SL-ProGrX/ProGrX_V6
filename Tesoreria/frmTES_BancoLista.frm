VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmTES_BancoLista 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tesorería : Bancos Disponibles"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7605
   ControlBox      =   0   'False
   HelpContextID   =   1004
   Icon            =   "frmTES_BancoLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optX 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Código"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optX 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Descripción"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
   Begin VB.OptionButton optX 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Cuenta"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_BancoLista.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_BancoLista.frx":0BE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16711680
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CUENTA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DESCRIPCION"
         Object.Width           =   8114
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CODIGO"
         Object.Width           =   1834
      EndProperty
   End
End
Attribute VB_Name = "frmTES_BancoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLlenaLista()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla los Tes_Bancos diponibles para el modulo de tesoreria.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

strSQL = "Select Id_Banco,Cta,Descripcion From Tes_Bancos Where Estado = 'A'"
Select Case True
  Case optX.Item(0).Value
    strSQL = strSQL & " order by cta"
  Case optX.Item(1).Value
    strSQL = strSQL & " order by descripcion"
  Case optX.Item(2).Value
    strSQL = strSQL & " order by id_banco"
End Select

rs.Open strSQL, glogon.Conection, adOpenStatic
   
lsw.ListItems.Clear

Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , CStr(rs!Cta))
   itmX.SmallIcon = 1
   itmX.SubItems(1) = UCase(rs!Descripcion)
   itmX.SubItems(2) = rs!id_banco
   rs.MoveNext
Loop
rs.Close
     
Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub Form_Load()
 Call sbLlenaLista
End Sub

Private Sub lsw_Click()
 Unload Me
End Sub

Private Sub optX_Click(Index As Integer)
Call sbLlenaLista
End Sub
