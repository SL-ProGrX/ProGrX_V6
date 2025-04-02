VERSION 5.00
Begin VB.Form frmTES_Imagenes 
   Caption         =   "Fondos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image img1024x768 
      Height          =   1095
      Left            =   1800
      Picture         =   "frmTES_Imagenes.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image img800x600 
      Height          =   975
      Left            =   360
      Picture         =   "frmTES_Imagenes.frx":59F87
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmTES_Imagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

