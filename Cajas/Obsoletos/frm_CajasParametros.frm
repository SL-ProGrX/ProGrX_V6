VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frm_CajasParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros para cajas"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_CajasParametros.frx":0000
   ScaleHeight     =   5610
   ScaleWidth      =   7395
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7335
      _Version        =   524288
      _ExtentX        =   12938
      _ExtentY        =   8070
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      ColHeaderDisplay=   0
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frm_CajasParametros.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Parámetros de Cajas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frm_CajasParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
