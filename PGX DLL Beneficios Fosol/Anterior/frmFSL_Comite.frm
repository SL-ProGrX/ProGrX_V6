VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFSL_Comite 
   Caption         =   "Miembros de Comite"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "frmFSL_Comite.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Historico"
      TabPicture(1)   =   "frmFSL_Comite.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vGridHistorico"
      Tab(1).ControlCount=   1
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9015
         _Version        =   524288
         _ExtentX        =   15901
         _ExtentY        =   6376
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Comite.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridHistorico 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   9015
         _Version        =   524288
         _ExtentX        =   15901
         _ExtentY        =   6165
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Comite.frx":0637
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmFSL_Comite.frx":0BD8
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros de Cómite"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmFSL_Comite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vCodigo As Integer
Dim vDescripcion As String
Dim vEstado As Integer

Private Sub Form_Activate()
  'vModulo =
End Sub

Private Sub sbCargaMiembros()
On Error GoTo error
   
  strSQL = "Select COD_MIEMBRO, NOMBRE, ACTIVO, FECHA_REGISTRO  " _
         & "from FSL_COMITE where ACTIVO = 1"
         
  Call sbCargaGridCheckIni(vGrid, 4, strSQL)
   
Exit Sub

error:
  MsgBox Err.Description
End Sub

Private Sub sbGuardaMiembro()
On Error GoTo error
   
   strSQL = "Insert FSL_COMITE (COD_MIEMBRO, NOMBRE, ACTIVO, USUARIO_REGISTRA, FECHA_REGISTRO) " _
          & " values ()"
          
   glogon.Conection.Execute strSQL
  
Exit Sub

error:
  MsgBox Err.Description
  
End Sub

Private Sub sbModificaMiembro()
On Error GoTo error
   
   strSQL = "Update FSL_COMITE set NOMBRE, ACTIVO, USUARIO_REGISTRA, FECHA_REGISTRO " _
          & "Where COD_MIEMBRO"
   
   glogon.Conection.Execute strSQL

Exit Sub

error:
  MsgBox Err.Description
End Sub

Private Sub sbCargaHistorico()
On Error GoTo error
   
  strSQL = "Select COD_MIEMBRO, NOMBRE, ACTIVO, FECHA_REGISTRO  " _
         & "from FSL_COMITE where ACTIVO = 1"
         
  Call sbCargaGridCheckIni(vGridHistorico, 4, strSQL)
  
Exit Sub

error:
  MsgBox Err.Description
  
End Sub

Private Function fxValidaMiembro() As Boolean

On Error GoTo error
   
   With vGrid
   
      .Row = .ActiveRow
      .Col = 1
      
      If .Text = Empty Then
        fxValidaMiembro = False
      Else
        fxValidaMiembro = True
        vCodigo = CInt(.Text)
        
        .Col = 2
        vDescripcion = .Text
        
        .Col = 3
        vEstado = CInt(.Text)
      End If
   
   End With
   
Exit Function

error:
  MsgBox Err.Description
  
End Function

Private Sub Form_Load()
   Call sbCargaMiembros
   Call sbCargaHistorico
End Sub

