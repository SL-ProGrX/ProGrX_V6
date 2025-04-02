VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_CalculoOperacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de la Operación"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   3016
   Icon            =   "frmCR_CalculoOperacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9780
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4572
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   9492
      _Version        =   1310723
      _ExtentX        =   16743
      _ExtentY        =   8064
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
      Item(0).Caption =   "Refundiciones"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "lblColFaltaCancela"
      Item(0).Control(2)=   "lblColNoRefunde"
      Item(0).Control(3)=   "lblColRefunde"
      Item(1).Caption =   "Cargos"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "txtDesembolsos"
      Item(1).Control(1)=   "txtDias"
      Item(1).Control(2)=   "chkCuota"
      Item(1).Control(3)=   "Label11"
      Item(1).Control(4)=   "Label10"
      Item(1).Control(5)=   "lswCargos"
      Item(2).Caption =   "Resultados"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "Frame2"
      Item(2).Control(1)=   "Frame1"
      Item(2).Control(2)=   "cmdCalcular"
      Item(3).Caption =   "Disponibles"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswDisp"
      Begin XtremeSuiteControls.ListView lswCargos 
         Height          =   3372
         Left            =   -69880
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   9252
         _Version        =   1310723
         _ExtentX        =   16319
         _ExtentY        =   5948
         _StockProps     =   77
         BackColor       =   -2147483643
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
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDisp 
         Height          =   3972
         Left            =   -69880
         TabIndex        =   65
         Top             =   480
         Visible         =   0   'False
         Width           =   9372
         _Version        =   1310723
         _ExtentX        =   16531
         _ExtentY        =   7006
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3492
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   9252
         _Version        =   1310723
         _ExtentX        =   16319
         _ExtentY        =   6159
         _StockProps     =   77
         BackColor       =   -2147483643
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
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdCalcular 
         Height          =   372
         Left            =   -62680
         TabIndex        =   63
         Top             =   4200
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Calcular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkCuota 
         Height          =   252
         Left            =   -64120
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1310723
         _ExtentX        =   4466
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Deducir Primer Cuota"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cálculo de Disponible según garantía Sobre Ahorros"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3612
         Left            =   -65200
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   0
            X2              =   5040
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label lblNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   51
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label lblSaldos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   50
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   49
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblGAR_Ahorros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   48
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblGAR_Porcentaje 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   47
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Disponible Neto"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   46
            Top             =   2760
            Width           =   2292
         End
         Begin VB.Label Label13 
            Caption         =   "Disponible Bruto"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   45
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Aporte Obrero"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   44
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "(-) Saldo en Préstamos "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   43
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Porcentajes S/Ahorros"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "Monto Max. del Préstamo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   41
            Top             =   3120
            Width           =   2175
         End
         Begin VB.Label lblMontoPrestamo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   40
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label lblGAR_Pat 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3120
            TabIndex        =   39
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblGAR_Cap 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3720
            TabIndex        =   38
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblGAR_Patronal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   37
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblGAR_Capitaliza 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   36
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Aporte Patronal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   35
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Capitalización"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   34
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Obr"
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
            Index           =   1
            Left            =   2520
            TabIndex        =   54
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Pat"
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
            Index           =   2
            Left            =   3120
            TabIndex        =   53
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Cap"
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
            Index           =   3
            Left            =   3720
            TabIndex        =   52
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cálculos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3612
         Left            =   -69760
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Monto Préstamo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   6
            Left            =   240
            TabIndex        =   32
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblMontoCalculado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            TabIndex        =   31
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Total Rebajos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   30
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblTotalRebajos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            TabIndex        =   29
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Refundiciones"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   312
            Index           =   4
            Left            =   240
            TabIndex        =   28
            Top             =   1116
            Width           =   1572
         End
         Begin VB.Label lblRefundiciones 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   312
            Left            =   1800
            TabIndex        =   27
            Top             =   1116
            Width           =   2172
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Total Morosidad"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Top             =   675
            Width           =   1575
         End
         Begin VB.Label lblMorosidad 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            TabIndex        =   25
            Top             =   675
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Total Cargos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblCargos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Monto a Girar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   2355
            Width           =   1575
         End
         Begin VB.Label lblMontoAGirar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   2355
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Intereses"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   312
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   1428
            Width           =   1572
         End
         Begin VB.Label lblMontoInteres 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   312
            Left            =   1800
            TabIndex        =   19
            Top             =   1428
            Width           =   2172
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDesembolsos 
         Height          =   312
         Left            =   -68800
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDias 
         Height          =   312
         Left            =   -65440
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin VB.Label lblColRefunde 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Refundibles"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   120
         TabIndex        =   59
         Top             =   4080
         Width           =   3012
      End
      Begin VB.Label lblColNoRefunde 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "No Refundibles"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   3120
         TabIndex        =   58
         Top             =   4080
         Width           =   3132
      End
      Begin VB.Label lblColFaltaCancela 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Falta Cancelación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   6240
         TabIndex        =   57
         Top             =   4080
         Width           =   3132
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Días de Intereses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   -66880
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Desembolsos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   -69880
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   1080
   End
   Begin ComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   6
      Top             =   6444
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.Tag             =   ""
            Object.ToolTipText     =   "Cuotas Liberadas"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.Tag             =   ""
            Object.ToolTipText     =   "Diferencia en Cuotas"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboCalculoAdd 
      Height          =   288
      Left            =   1920
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2990
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   5772
      _Version        =   1310723
      _ExtentX        =   10181
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3600
      TabIndex        =   11
      Top             =   960
      Width           =   5772
      _Version        =   1310723
      _ExtentX        =   10181
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   12
      Top             =   960
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   312
      Left            =   7680
      TabIndex        =   13
      Top             =   1320
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPlazo 
      Height          =   312
      Left            =   4440
      TabIndex        =   14
      Top             =   1320
      Width           =   612
      _Version        =   1310723
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtTasa 
      Height          =   312
      Left            =   5760
      TabIndex        =   15
      Top             =   1320
      Width           =   612
      _Version        =   1310723
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtMontoSolicitado 
      Height          =   312
      Left            =   1920
      TabIndex        =   16
      Top             =   1320
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label8 
      Caption         =   "Monto"
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
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   612
   End
   Begin VB.Image imgMonto 
      Height          =   240
      Left            =   3630
      ToolTipText     =   "Calcular Monto para Giro en Cero"
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label8 
      Caption         =   "Linea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Cuota"
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
      Left            =   6840
      TabIndex        =   3
      Top             =   1320
      Width           =   492
   End
   Begin VB.Label Label4 
      Caption         =   "%"
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
      Left            =   6480
      TabIndex        =   2
      Top             =   1320
      Width           =   132
   End
   Begin VB.Label Label3 
      Caption         =   "Tasa"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "Plazo"
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
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   492
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCR_CalculoOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vActualiza As Boolean, vMovCuota As Boolean
Dim mFrecuenciaPago As String


Private Sub chkCuota_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iMes As Integer, lngAnio As Long
Dim vFecha As Date, vFechaCalculo As Date

vFecha = fxFechaServidor

If chkCuota.Value = vbChecked Then

'    If Day(vFecha) > 15 Then
         iMes = Month(vFecha)
         lngAnio = Year(vFecha)
         If iMes = 12 Then
            iMes = 1
            lngAnio = lngAnio + 1
         Else
            iMes = iMes + 1
         End If
         'Calcular Intereses Hasta el Ultimo día del Mes
         vFechaCalculo = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
         vFechaCalculo = DateAdd("d", -1, vFechaCalculo)
         
         txtDias = (Abs(DateDiff("d", vFechaCalculo, vFecha)) + 1)
    
'    Else 'Fecha dia 15
'      txtDias = 0
'    End If

Else 'Cuota
    
    'Carga Descripcion del Codigo, y Sus Rangos
    strSQL = "select fechacortealterna,fechacorte,dbo.MyGetdate() as Fecha" _
           & " from catalogo where codigo = '" & txtCodigo & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
      rs.Close
      MsgBox "No se encontró el código especificado...", vbExclamation
      Exit Sub
    End If
    
    If rs!FechaCorteAlterna = "S" Then
      txtDias = DateDiff("d", rs!fecha, rs!FechaCorte) + 1
      If txtDias < 0 Then txtDias = 0
    Else
      strSQL = "select cr_fecha_calculo,dbo.MyGetdate() as fecha from par_ahcr"
      rs.Close
      Call OpenRecordSet(rs, strSQL)
      txtDias = DateDiff("d", rs!fecha, rs!cr_fecha_calculo) + 1
      If txtDias < 0 Then txtDias = 0
    End If
     
    rs.Close
   
End If

End Sub

Private Sub cmdCalcular_Click()
Dim curRefundir As Currency, i As Integer, curMonto As Currency
Dim curInteres As Currency, lng As Long, itmX As ListViewItem
Dim curCargos As Currency, curMora As Currency

'Es el Monto por Diferencial para Créditos Sobre Excedentes
Dim curMontoRebajar As Currency, curMontoPrestamo As Currency
Dim curCuotaRef As Currency

On Error Resume Next


vMovCuota = False

curRefundir = 0
curCargos = 0
curMora = 0
curMontoRebajar = 0
curMontoPrestamo = 0
curCuotaRef = 0


'Procesa Refundiciones
With lsw
  For lng = 1 To .ListItems.Count
    If .ListItems.Item(lng).Checked Then
     'Saldo Menos - Amortizacion Atrasada, ya que se Toma en cuenta en Mora
      curRefundir = curRefundir + (CCur(.ListItems.Item(lng).SubItems(2)) - CCur(.ListItems.Item(lng).SubItems(6)))
      curCuotaRef = curCuotaRef + CCur(.ListItems.Item(lng).SubItems(7))
'      If txtCodigo.Tag = 1 And UCase(Trim(lsw.ListItems(lng).SubItems(1))) = UCase(Trim(txtCodigo)) Then
'         curMontoResta = curMontoResta + CCur((lsw.ListItems(lng).SubItems(3)))
'      End If
    Else
    End If
   Next lng
End With


'Procesa Mora
With lsw
  For lng = 1 To .ListItems.Count
     'Saldo Menos - Amortizacion Atrasada, ya que se Toma en cuenta en Mora
      curMora = curMora + (CCur(.ListItems.Item(lng).SubItems(4)) + CCur(.ListItems.Item(lng).SubItems(5)) + CCur(.ListItems.Item(lng).SubItems(6)))
   Next lng
End With

'Si es Prestamo Sobre Excedentes Descontar el Saldo (Se Asume que es Igual al Monto Aprobado)
'Ya que no recibe Abonos.
If fxCreditoExcedente(txtCodigo) Then
  For lng = 1 To lsw.ListItems.Count
     If UCase(Trim(lsw.ListItems.Item(lng).SubItems(1))) = UCase(Trim(txtCodigo)) Then
      curMontoRebajar = curMontoRebajar + CCur(lsw.ListItems.Item(lng).SubItems(2))
     End If
   Next lng
End If

'Procesa Cargos Adicionales
With lswCargos
   For lng = 1 To .ListItems.Count
    If Mid(.ListItems.Item(lng).SubItems(2), 1, 1) = "P" Then
         curCargos = curCargos + (CCur(txtMontoSolicitado) * (CCur(.ListItems.Item(lng).SubItems(3)) / 100))
    Else
         curCargos = curCargos + CCur(.ListItems.Item(lng).SubItems(3))
    End If
   Next lng
 End With

'Otros Cargos
curCargos = curCargos + IIf((txtDesembolsos = ""), 0, txtDesembolsos)
curCargos = curCargos + IIf((chkCuota.Value = 1), txtCuota, 0)

lblRefundiciones.Caption = Format(curRefundir, "Standard")
lblCargos.Caption = Format(curCargos, "Standard")
lblMorosidad.Caption = Format(curMora, "Standard")

'Bruto Sobre Ahorros - Saldos Sobre Ahorros
lblNeto.Caption = Format(CCur(lblBruto.Caption) - CCur(lblSaldos.Caption), "Standard")

'Monto Max del Prestamos Neto + Refundiciones
lblMontoPrestamo.Caption = Format(CCur(lblNeto.Caption) + CCur(lblRefundiciones.Caption), "Standard")

'Si el Tag del Codigo es 2, es Sobre Ahorros, entonces indicar el monto Maximo del prestamo
If txtCodigo.Tag = 2 And vActualiza Then
  txtMontoSolicitado = lblMontoPrestamo.Caption
  
  If IsNumeric(txtMontoSolicitado) Then
      If CCur(txtMontoSolicitado) > fxRangoMaximo(txtCodigo) Then
        i = MsgBox("El monto disponible es mayor que monto maximo establecido en la Tabla de Rangos de la Línea del Préstamo," _
                 & " desea restaurar el monto máximo reglamentario ?", vbYesNo)
        If i = vbYes Then
           txtMontoSolicitado = fxRangoMaximo(txtCodigo)
        End If
      End If
  End If
  
  Call txtMontoSolicitado_LostFocus
  Call txtMontoSolicitado_KeyDown(vbKeyReturn, 0)
End If

If IsNumeric(txtMontoSolicitado) And IsNumeric(txtPlazo) _
  And IsNumeric(txtTasa) Then
'  txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado) - curMontoRebajar, txtPlazo, txtTasa)
  txtCuota = fxCalcula_Cuota(CCur(txtMontoSolicitado), txtPlazo, txtTasa, mFrecuenciaPago)
End If


If Val(txtCuota) > 0 Then
 curMonto = txtMontoSolicitado
 curMonto = curMonto - (curRefundir + curCargos + curMora)
  
  If txtDias <> "" Then
   curInteres = ((CCur(txtMontoSolicitado) - curMontoRebajar) * txtTasa / 36000) * txtDias
   curMonto = curMonto - curInteres
   lblMontoInteres.Caption = Format(curInteres, "Standard")
  End If
  
  lblTotalRebajos.Caption = Format(CCur(lblCargos) _
         + CCur(lblRefundiciones) + CCur(lblMorosidad) + CCur(lblMontoInteres), "Standard")
  
  If curMonto < 0 Then
    lblMontoAGirar.ForeColor = vbRed
  Else
    lblMontoAGirar.ForeColor = vbBlack
  End If
  
  lblMontoAGirar.Caption = Format(curMonto, "Standard")
  
  'Calcular el Monto del Prestamo, para poder sacarlo Bruto el Monto
'  If CCur(txtMontoSolicitado) > (curRefundir + curCargos + curMora + curInteres) Then
    curMonto = curMonto + (curRefundir + curCargos + curMora + curInteres)
    Do While (curMonto - (curRefundir + curCargos + curMora + curInteres)) <= CCur(txtMontoSolicitado)
      curMonto = curMonto + 100
      
'      curMonto = CCur(txtMontoSolicitado.Text) - (curMonto - (curRefundir + curCargos + curMora + curInteres))
      
      If txtDias <> "" Then
       curInteres = (curMonto * txtTasa / 36000) * txtDias
      End If
    Loop
'  End If
  
  lblMontoCalculado.Caption = Format(curMonto, "Standard")

 
End If

StatusBarX.Panels.Item(1) = Format(curCuotaRef, "Standard")
StatusBarX.Panels.Item(2) = Format(curCuotaRef - CCur(txtCuota.Text), "Standard")

vMovCuota = True

End Sub


Private Sub Form_Load()

tcMain.Item(0).Selected = True

imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mFrecuenciaPago = "M"

With lsw.ColumnHeaders
  .Clear
  .Add , , "Operación", 1400
  .Add , , "Linea", 1100, vbCenter
  .Add , , "Saldo", 1800, vbRightJustify
  .Add , , "Garantía", 2100

  .Add , , "Mora.IntCor", 1300, vbRightJustify
  .Add , , "Mora.IntMor", 1300, vbRightJustify
  .Add , , "Mora.Principal", 1300, vbRightJustify
  .Add , , "Cuota", 1400, vbRightJustify
End With


With lswCargos.ColumnHeaders
  .Clear
  .Add , , "Cargo", 1100, vbCenter
  .Add , , "Descripción", 3100
  .Add , , "Tipo", 1500, vbCenter
  .Add , , "Valor", 1800, vbRightJustify
End With


With lswDisp.ColumnHeaders
  .Clear
  .Add , , "Garantía", 3000
  .Add , , "Monto", 2100, vbRightJustify
  .Add , , "Saldo", 2100, vbRightJustify
  .Add , , "Disponible", 2100, vbRightJustify

End With

lblColRefunde.ForeColor = vbBlack
lblColNoRefunde.ForeColor = vbRed
lblColFaltaCancela.ForeColor = vbBlue

 cboCalculoAdd.AddItem "Monto del Crédito"
 cboCalculoAdd.AddItem "Monto a Girar"
 cboCalculoAdd.AddItem "Giro en Cero"
 cboCalculoAdd.Text = "Monto del Crédito"


End Sub

Private Sub imgMonto_Click()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim curMonto As Currency, curRebajos As Currency, curCargos As Currency, curMntAdd As Currency
'Dim curIntereses As Currency, curPrimerCuota As Currency, iDias As Long
'Dim curPoliza As Currency, curPolizaBase As Currency
'
'Dim vGarantia As String, vConvenio As String
'Dim vCobraTasaFormaliza As Boolean, vCreditoExcedentes As Boolean
'Dim i As Integer, vFecha As Date, curTemp As Currency
'
'On Error GoTo vError
'
'Me.MousePointer = vbHourglass
'
'If cboCalculoAdd.Text = "Monto del Crédito" Then
'  Me.MousePointer = vbDefault
'  Exit Sub
'End If
'
'
''Calcula valores fijos
'curMntAdd = CCur(txtMonto.Text)
'curRebajos = fxMontoEnRefundiciones(Operacion.Operacion) + fxMontoEnDesembolsos(Operacion.Operacion) + fxMontoEnRetenciones(Operacion.Operacion)
'vCobraTasaFormaliza = fxCobraTasaFormaliza(fxCodigoCbo(cboDestino))
'vCreditoExcedentes = fxCreditoExcedente(Operacion.Codigo)
'
'
'strSQL = "select R.Garantia,R.cuota,R.int,C.convenio,R.FECHA_CALCULO_INT,R.FECHA_INICIO_CALCULO" _
'       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
'       & " where R.id_solicitud =" & Operacion.Operacion
'Call OpenRecordSet(rs, strSQL)
'    vGarantia = rs!Garantia
'    vConvenio = rs!Convenio
'    If IsNull(rs!fecha_calculo_int) Then
'       vFecha = fxFechaServidor
'       vFecha = DateAdd("m", 1, vFecha)
'       vFecha = DateAdd("d", -1, CDate(Year(vFecha) & "/" & Format(Month(vFecha), "00") & "/01"))
'    Else
'       vFecha = rs!fecha_calculo_int
'    End If
'
'
'    If vFecha < rs!fecha_inicio_calculo Then
'     iDias = 0
'    Else
'     iDias = vFecha - rs!fecha_inicio_calculo + 1
'    End If
'rs.Close
'
'
''Base de la Poliza
'strSQL = "select CR_PSDMNT from par_ahcr"
'Call OpenRecordSet(rs, strSQL)
'If rs.EOF And rs.BOF Then
'  curPolizaBase = 0
'Else
'  curPolizaBase = IIf(IsNull(rs!cr_PsdMnt), 0, rs!cr_PsdMnt)
'End If
'rs.Close
'
'curCargos = 0
'curMonto = curRebajos
'curTemp = 0
'
'
'' cboCalculoAdd.AddItem "Monto del Crédito"
'' cboCalculoAdd.AddItem "Monto a Girar"
'' cboCalculoAdd.AddItem "Giro en Cero"
'
'If cboCalculoAdd.Text = "Monto a Girar" Then
'  curMonto = curMonto + curMntAdd
'End If
'
'i = 5 'Acercamientos
'
''Inicio de Calculos y Variaciones
'For i = 1 To 5
'    curIntereses = 0
'    If vCobraTasaFormaliza Then
'         If Operacion.EstadoSolicitud = "F" Then
'            If vCreditoExcedentes Then
'                 curIntereses = fxInteresesHastaFormalizar(dtpDesembolso.Value, curMonto)
'            Else
'                 curIntereses = ((curMonto * CCur(txtTasa.Text)) / (36000)) * iDias
'            End If
'
'
'         Else
'             If chkPrimera.Value = vbChecked Then
'                 curIntereses = fxInteresesDiasPrimerCuota(dtpDesembolso.Value, curMonto, txtTasa)
'             Else
'                 curIntereses = fxInteresesHastaFormalizar(dtpDesembolso.Value, curMonto)
'             End If
'         End If
'    End If
'
'
'
'    curPrimerCuota = IIf((chkPrimera.Value = vbChecked), txtCuota, 0)
'
'
'    'Calcula Poliza
'    curPoliza = 0
'    If vCobraTasaFormaliza Then
'        If vGarantia <> "H" And vConvenio = "N" Then
'            curPoliza = (curMonto / 1000000) * curPolizaBase
'        End If
'    End If
'
'
'    'Definir el Monto Base del Credito
'
'
'    If cboCalculoAdd.Text = "Monto a Girar" Then
'        curMonto = Round(curPoliza, 2) + curPrimerCuota + Round(curIntereses, 2) + curRebajos + Round(curCargos, 2) + curMntAdd
'    Else
'        curMonto = Round(curPoliza, 2) + curPrimerCuota + Round(curIntereses, 2) + curRebajos + Round(curCargos, 2)
'    End If
'
'    curTemp = curCargos
'
'    'Procesar Cargos y Recuperarlos
'    Call sbCargosAdicionales(Operacion.Operacion, Operacion.Codigo, Round(curMonto, 2))
'    curCargos = fxMontoEnCargos(Operacion.Operacion)
'
'    curMonto = curMonto + (curCargos - curTemp)
'    txtCuota.Text = fxCalcula_Cuota(CDbl(Round(curMonto, 2)), txtPlazo, txtTasa)
'
'Next i
'
'cboCalculoAdd.Text = "Monto del Crédito"
'txtMonto.Text = Format(curMonto, "Standard")
'
'txtMonto.SetFocus
'
'Me.MousePointer = vbDefault
'
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDisponibles()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Me.MousePointer = vbHourglass

On Error GoTo vError

lswDisp.ListItems.Clear

strSQL = "exec spVoxAhorros '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
 Set itmX = lswDisp.ListItems.Add(, , "Sobre Ahorros")
     itmX.SubItems(1) = Format(rs!Disponible, "Standard")
     itmX.SubItems(2) = Format(rs!Saldos, "Standard")
     itmX.SubItems(3) = Format(rs!Disponible - rs!Saldos, "Standard")
rs.Close
        
strSQL = "exec spVoxFiduciario '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
 Set itmX = lswDisp.ListItems.Add(, , "Fiduciaria")
     itmX.SubItems(1) = Format(rs!Disponible, "Standard")
     itmX.SubItems(2) = Format(rs!Saldos, "Standard")
     itmX.SubItems(3) = Format(rs!Disponible - rs!Saldos, "Standard")
rs.Close
        
'Vivienda

'Excedentes
strSQL = "exec spVoxExcedenteCredito '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
 Set itmX = lswDisp.ListItems.Add(, , "Excedentes")
     itmX.SubItems(1) = Format(rs!Base, "Standard")
     itmX.SubItems(2) = Format(rs!Saldos, "Standard")
     itmX.SubItems(3) = Format(rs!Base - rs!Saldos, "Standard")
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
 Call cmdCalcular_Click
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 2
     Call cmdCalcular_Click
  Case 3
     Call sbDisponibles
End Select

End Sub

Private Sub Timer1_Timer()

Timer1.Interval = 0

If GLOBALES.gCedulaActual <> "" Then
   txtCedula.Text = GLOBALES.gCedulaActual
   Call txtCedula_KeyDown(vbKeyReturn, 0)
End If

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbLimpiaDatos
  Call sbCargaDatos
  txtCodigo.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Cedula"
  gBusquedas.Orden = "Cedula"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCedula = gBusquedas.Resultado
    txtNombre = gBusquedas.Resultado2
    Call sbLimpiaDatos
    Call sbCargaDatos
    txtCodigo.SetFocus
  End If
End If

End Sub


Private Sub sbLimpiaDatos()
 txtNombre = ""
 txtMontoSolicitado = ""
 txtPlazo = ""
 txtTasa = ""
 txtCuota = ""
 txtDesembolsos = 0
 txtDias = 0
 txtCodigo = ""
 txtDescripcion = ""
  
 chkCuota.Value = 0
 lblGAR_Porcentaje.Caption = ""
 lblGAR_Pat.Caption = ""
 lblGAR_Cap.Caption = ""
 
 lblSaldos.Caption = ""
 lblGAR_Ahorros.Caption = ""
 lblBruto.Caption = ""
 lblNeto.Caption = ""
 
 lswCargos.ListItems.Clear
 lsw.ListItems.Clear
 
 lblCargos.Caption = 0
 lblMorosidad.Caption = 0
 lblRefundiciones.Caption = 0
 
 lblMontoInteres.Caption = 0
 lblMontoAGirar.Caption = 0
 
 

End Sub

Private Function fxTotalPeriodo(vPrimer As Long) As Integer
Dim i As Currency, x  As Integer

x = 1
i = vPrimer
Do While i < GLOBALES.glngFechaCR
  i = fxFechaProcesoSiguiente(i)
  x = x + 1
Loop

fxTotalPeriodo = x

End Function

Private Sub sbCargaDatos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curSaldos As Currency
   
On Error GoTo vError
   
Me.MousePointer = vbHourglass


strSQL = "exec spCrdGarantiaPatDetalle '" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
  txtNombre.Text = rs!Nombre
  lblGAR_Porcentaje.Caption = rs!Porc_Obrero
  lblGAR_Pat.Caption = rs!Porc_Patronal
  lblGAR_Cap.Caption = rs!Porc_Capitaliza
  lblGAR_Patronal.Caption = Format(rs!Mnt_Patronal, "Standard")
  lblGAR_Ahorros.Caption = Format(rs!Mnt_Obrero, "Standard")
  lblGAR_Capitaliza.Caption = Format(rs!Mnt_Capitaliza, "Standard")
  
  'Se sobre escriben más adelante
  lblSaldos.Caption = Format(rs!Saldo, "Standard")
  lblBruto.Caption = Format(rs!Monto, "Standard")
  lblNeto.Caption = Format(rs!Neto, "Standard")
rs.Close


 strSQL = "Select R.id_solicitud as Operacion,R.codigo,R.saldo,R.garantia,R.plazo,R.montoapr,Gar.Descripcion as 'GarantiaDesc'" _
        & ",R.cuota,R.amortiza as 'Recaudado',C.retencion,C.poliza,isnull(V.intc,0) as 'MoraIntc'" _
        & ",isnull(V.intm,0) as 'MoraIntm', isnull(V.amortiza,0) as 'MoraPrincipal'" _
        & ",C.REFUNDE,C.ACEPTAREFUN,C.refunde_tipo,C.refunde_porc,R.prideduc" _
        & ", datediff(m, dbo.fxSIFCorteAFechaInicio(R.PRIDEDUC), dbo.MyGETDATE()) / convert(float,R.PLAZO) as 'TiempoTranscurrido'" _
        & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
        & " inner join crd_garantia_tipos Gar on R.garantia = Gar.Garantia" _
        & " left join Vista_morosidad V on R.id_solicitud = V.id_solicitud" _
        & " where R.cedula = '" & Trim(txtCedula) & "' and R.saldo > 0 and R.proceso <> 'J' and R.estado = 'A'"
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  'Se cambia por la siguiente ya que la mora y otros rebajos los contemplan
  'ya que ahora se paga la mora de todos los prestamos y no solo del sobre ahorros
  'If !garantia = "A" Then curSaldos = curSaldos + !Saldo + !MoraIntm + !MoraIntc
  If rs!Garantia = "A" Then curSaldos = curSaldos + rs!Saldo
  
  Set itmX = lsw.ListItems.Add(, , CStr(rs!Operacion))
   itmX.SubItems(1) = UCase(rs!Codigo)
   
   If rs!retencion = "N" And rs!Poliza = "N" Then
       itmX.SubItems(2) = Format(rs!Saldo, "Standard")
   Else
     If rs!Plazo < 999 Then
       itmX.SubItems(2) = Format((rs!Cuota * rs!Plazo) - rs!Recaudado, "Standard")
     Else
       itmX.SubItems(2) = Format(rs!Saldo, "Standard")
     End If
   End If
   itmX.SubItems(3) = rs!GarantiaDesc
   itmX.SubItems(4) = Format(rs!MoraIntC, "Standard")
   itmX.SubItems(5) = Format(rs!MoraIntM, "Standard")
   itmX.SubItems(6) = Format(rs!MoraPrincipal, "Standard")
   itmX.SubItems(7) = Format(rs!Cuota, "Standard")
  
   'Verifica Colores
   'Negro = Refundible
   'Azul = Falta Periodo de Cancelacion
   'Rojo = No Refundible
   If rs!aceptarefun = "S" Then
      If rs!refunde_tipo = "P" Then
         If rs!TiempoTranscurrido > (rs!refunde_porc / 100) Then
           'Nada
         Else
           itmX.ForeColor = vbBlue
         End If
      Else
         If (rs!Saldo / rs!montoapr) > (rs!refunde_porc / 100) Then
           'Nada
         Else
           itmX.ForeColor = vbBlue
         End If
      
      End If
   Else
     itmX.ForeColor = vbRed
   End If
  
  
  rs.MoveNext
 Loop
  rs.Close

lblSaldos.Caption = Format(curSaldos, "Standard")
lblNeto.Caption = Format(CCur(lblBruto.Caption) - CCur(lblSaldos.Caption), "Standard")
lblMontoPrestamo.Caption = Format(CCur(lblNeto.Caption) + CCur(lblRefundiciones.Caption), "Standard")

vError:

Me.MousePointer = vbDefault

End Sub

Private Function fxRangoMaximo(vCodigo As String) As Currency
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select isnull(max(hasta),0) as Maximo from rangos where codigo = '" & vCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
  fxRangoMaximo = 0
Else
  fxRangoMaximo = rsX!Maximo
End If
rsX.Close
End Function



Private Sub sbCargaCodigo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, lng As Integer

On Error GoTo vError

txtCodigo.Tag = 0
vActualiza = False


mFrecuenciaPago = "M"
'Carga Descripcion del Codigo, y Sus Rangos
strSQL = "select descripcion,fechacortealterna,fechacorte,dbo.MyGetdate() as Fecha, Base_Calculo" _
       & ",refunde,operaciones_activas  from catalogo where codigo = '" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "No se encontró el código especificado...", vbExclamation
  Exit Sub
End If

txtDescripcion = rs!Descripcion & ""

If rs!Base_Calculo = "06" Then
    mFrecuenciaPago = "Q"
End If

'Limpia Refundiciones Anteriores ya que se cambio de codigo
For lng = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(lng).Checked = False
Next lng


If rs!Refunde = "S" And rs!operaciones_activas = 1 Then
    'Procesa Refundiciones
    With lsw
      For lng = 1 To .ListItems.Count
        If UCase(Trim(.ListItems.Item(lng).SubItems(1))) = UCase(Trim(txtCodigo)) _
           And .ListItems.Item(lng).ForeColor <> vbRed Then
          .ListItems.Item(lng).Checked = True
        End If
       Next lng
    End With
End If

If rs!FechaCorteAlterna = "S" Then
  txtDias = DateDiff("d", rs!fecha, rs!FechaCorte) + 1
  If txtDias < 0 Then txtDias = 0
Else
  strSQL = "select cr_fecha_calculo,dbo.MyGetdate() as fecha from par_ahcr"
  rs.Close
  Call OpenRecordSet(rs, strSQL)
  txtDias = DateDiff("d", rs!fecha, rs!cr_fecha_calculo) + 1
  If txtDias < 0 Then txtDias = 0
End If
 
rs.Close

'Identifica el Tipo de Garantía
strSQL = "SELECT Gar.GARANTIA,Gar.FORMULARIO " _
       & " FROM CRD_CATALOGO_GARANTIAS Cat inner join CRD_GARANTIA_TIPOS Gar on Cat.GARANTIA = Gar.GARANTIA" _
       & " where Cat.CODIGO = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case Trim(rs!formulario)
    Case "F01" 'Sobre Ahorros
        txtCodigo.Tag = 2
        
    Case "F08" 'Excedentes
        txtCodigo.Tag = 1
        txtMontoSolicitado = Format(fxExcedenteDisponible(txtCedula), "Standard")

 
    Case "F06" 'Fondos de Ahorros Extraordinarios
       txtMontoSolicitado = Format(fxDisponibleFondos(txtCedula, ""), "Standard")
 
 End Select
 rs.MoveNext
Loop
rs.Close



'Sacar Cargos Adicionales
strSQL = "select C.* " _
       & " from cargos_adicionales C inner join cargos_asignacion A" _
       & " on C.cod_cargo = A.cod_cargo" _
       & " where A.codigo = '" & txtCodigo & "' and C.Automatico = 1"
Call OpenRecordSet(rs, strSQL, 0)
lswCargos.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lswCargos.ListItems.Add(, , rs!COD_CARGO)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.SubItems(2) = IIf((rs!Tipo = "P"), "PORCENTUAL", "MONTO")
      itmX.SubItems(3) = Format(rs!Valor, "Standard")
  rs.MoveNext
Loop
rs.Close



'Ver si es Calculo de Excedente y Ponerlo aqui
If txtCodigo.Tag <> 1 Then
    strSQL = "select ase_codigo from excedentes_parametros"
    Call OpenRecordSet(rs, strSQL)
    If UCase(Trim(rs!ase_codigo)) = UCase(Trim(txtCodigo)) Then
       txtCodigo.Tag = 1
       txtMontoSolicitado = Format(fxExcedenteDisponible(txtCedula), "Standard")
    End If
    rs.Close
End If 'txtCodigo.Tag <> 1

Call chkCuota_Click

If txtCodigo.Tag = 2 Then
  vActualiza = True
  Call cmdCalcular_Click
  
  tcMain.Item(2).Selected = True
  
End If

vError:

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbCargaCodigo
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Codigo"
  gBusquedas.Orden = "Codigo"
  gBusquedas.Consulta = "select Codigo,Descripcion from catalogo"
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCodigo = gBusquedas.Resultado
    txtDescripcion = gBusquedas.Resultado2
    Call sbCargaCodigo
  End If
End If

End Sub


Private Sub txtCuota_Change()
If vMovCuota Then Call cmdCalcular_Click
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoSolicitado.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select Codigo,Descripcion from catalogo"
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCodigo = gBusquedas.Resultado
    txtDescripcion = gBusquedas.Resultado2
    Call sbCargaCodigo
  End If
End If

End Sub

Private Sub txtDesembolsos_KeyPress(KeyAscii As Integer)

On Error GoTo vError

If KeyAscii = vbKeyReturn Then
 txtDesembolsos = Format(txtDesembolsos, "Standard")
 txtDias.SetFocus
End If

vError:

End Sub
Private Sub txtDias_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtDias = Format(txtDias, "##0")
 Else
 KeyAscii = Validacion(KeyAscii)
End If
End Sub

Private Sub txtTasa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Then
 vActualiza = False
 txtTasa = Format(txtTasa, "#0.00")
 
 If txtPlazo = "" Or txtTasa = "" Or txtMontoSolicitado = "" Or CCur(txtPlazo) = 0 _
        Or CCur(txtTasa) = 0 Or CCur(txtMontoSolicitado) = 0 Then
  'nada
 Else
     txtCuota.Text = fxCalcula_Cuota(CCur(txtMontoSolicitado), txtPlazo, CCur(txtTasa), mFrecuenciaPago)
 End If

Else
  vActualiza = False
End If
vError:
End Sub


Private Sub txtMontoSolicitado_GotFocus()
On Error GoTo vError
txtMontoSolicitado = CCur(txtMontoSolicitado)
vError:
End Sub

Private Sub txtMontoSolicitado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And txtMontoSolicitado <> "" Then
   txtPlazo = fxCatalogoRango(txtCodigo, txtMontoSolicitado, "P")
   txtTasa = fxCatalogoRango(txtCodigo, txtMontoSolicitado, "I")
   txtPlazo.SetFocus
 End If
vError:
End Sub

Private Sub txtMontoSolicitado_LostFocus()
On Error GoTo vError
txtMontoSolicitado = Format(txtMontoSolicitado, "Standard")
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCedula = gBusquedas.Resultado
    txtNombre = gBusquedas.Resultado2
    Call sbLimpiaDatos
    Call sbCargaDatos
    txtCodigo.SetFocus
  End If
End If
End Sub


Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
On Error GoTo vError

If KeyAscii = vbKeyReturn Then
 txtPlazo = Format(txtPlazo, "##0")
 txtTasa.SetFocus
End If

vError:

End Sub
