VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_Categorias_Credito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categorías de Crédito"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11595
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   11535
      _Version        =   1572864
      _ExtentX        =   20346
      _ExtentY        =   9975
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
      ItemCount       =   4
      Item(0).Caption =   "Probabilidad Default"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid(0)"
      Item(1).Caption =   "Probabilidad Mora"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid(1)"
      Item(2).Caption =   "Segmentos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGrid(2)"
      Item(3).Caption =   "Carga Masiva Probabilidad"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "txtArchivo"
      Item(3).Control(1)=   "btnArchivo(0)"
      Item(3).Control(2)=   "btnArchivo(1)"
      Item(3).Control(3)=   "btnArchivo(2)"
      Item(3).Control(4)=   "Label1(2)"
      Item(3).Control(5)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4335
         Left            =   -69640
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   10815
         _Version        =   1572864
         _ExtentX        =   19076
         _ExtentY        =   7646
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
         Appearance      =   21
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4815
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   8493
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
         SpreadDesigner  =   "frmCR_Categorias_Credito.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5055
         Index           =   1
         Left            =   -70000
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   8916
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_Categorias_Credito.frx":0921
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5055
         Index           =   2
         Left            =   -70000
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   8916
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_Categorias_Credito.frx":11AC
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   435
         Left            =   -67720
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1572864
         _ExtentX        =   12086
         _ExtentY        =   762
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   0
         Left            =   -60760
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmCR_Categorias_Credito.frx":1A0D
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   1
         Left            =   -60280
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmCR_Categorias_Credito.frx":210D
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   2
         Left            =   -59800
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmCR_Categorias_Credito.frx":2826
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -69040
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento Categoría Créditicia"
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
      Height          =   480
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmCR_Categorias_Credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mTipo As String
Dim vPaso As Boolean


Private Sub sbLista(pTipo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

mTipo = UCase(pTipo)


Select Case pTipo
    Case "P"
        strSQL = "SELECT ID_PROBABILIDAD_DEF, DESCRIPCION, CATEGORIA ,VALOR_INICIAL, VALOR_FINAL" _
               & ", USUARIO_REGISTRA, FEC_REGISTRA, USUARIO_MODIFICA, FEC_MODIFICA" _
               & "  FROM CRD_CATALOGO_PROBABILIDAD_DEF"
        Call sbCargaGrid(vGrid(0), 9, strSQL)
    Case "M"
       strSQL = "SELECT ID_PROBABILIDAD_MORA, DESCRIPCION, TIPO_MORA, PORC_PROBABILIDAD" _
              & ", USUARIO_REGISTRA, FEC_REGISTRA, USUARIO_MODIFICA, FEC_MODIFICA" _
              & " From CRD_CATALOGO_PROBABILIDAD_MORA"
              
        Call sbCargaGrid(vGrid(1), 8, strSQL)
    
    Case "S"
        strSQL = "select Id_Segmento, COD_SEGMENTO, DESCRIPCION, PORC_SEGMENTO" _
               & " , USUARIO_REGISTRA, FEC_REGISTRA, USUARIO_MODIFICA, FEC_MODIFICA" _
               & " From CRD_SEGMENTOS_PROBABILIDAD"
        Call sbCargaGrid(vGrid(2), 8, strSQL)
    
    Case Else
        
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

tcMain.Item(0).Selected = True

Call sbLista("P")

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
Case 0
    Call sbLista("P")
Case 1
    Call sbLista("M")
Case 2
    Call sbLista("S")
    

End Select

End Sub
