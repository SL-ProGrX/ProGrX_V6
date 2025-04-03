VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaTipos_Prenda_GastosHonorarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gastos, Honorarios y Examenes: Prendarios"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   16425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   16335
      _Version        =   1572864
      _ExtentX        =   28813
      _ExtentY        =   12303
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
      Item(0).Caption =   "Gastos y Honorarios Prendarios"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "cboTipo"
      Item(0).Control(1)=   "vGrid(0)"
      Item(0).Control(2)=   "Label2"
      Item(1).Caption =   "Exámenes"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid(2)"
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   3120
         TabIndex        =   1
         Top             =   480
         Width           =   4575
         _Version        =   1572864
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6135
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   16335
         _Version        =   524288
         _ExtentX        =   28813
         _ExtentY        =   10821
         _StockProps     =   64
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
         MaxCols         =   10
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaTipos_Prenda_GastosHonorarios.frx":0000
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6375
         Index           =   2
         Left            =   -70000
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   16335
         _Version        =   524288
         _ExtentX        =   28813
         _ExtentY        =   11245
         _StockProps     =   64
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
         MaxCols         =   12
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frmPreaTipos_Prenda_GastosHonorarios.frx":0A13
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Concepto"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimientos de Gastos, Honorarios y Examenes: Prendarios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   11295
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   16455
   End
End
Attribute VB_Name = "frmPreaTipos_Prenda_GastosHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, mTipo As String


Private Sub sbLista(pTipo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

mTipo = UCase(pTipo)

strSQL = "exec spCrd_Prea_Config_Prenda_Listas '" & pTipo & "'"

Select Case pTipo
    Case "C", "T" 'Constitución
        Call sbCargaGrid(vGrid(0), vGrid(0).MaxCols, strSQL)
    Case "E" 'Examenes
        Call sbCargaGrid(vGrid(2), vGrid(2).MaxCols, strSQL)
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub

Call sbLista(UCase(Mid(cboTipo.Text, 1, 1)))

End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
    cboTipo.Clear
    cboTipo.AddItem "Constitución"
    cboTipo.AddItem "Traspaso"
    cboTipo.Text = "Constitución"
vPaso = False


tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

Call cboTipo_Click

End Sub

Private Function fxGuardar(Index As Integer) As Long

On Error GoTo vError

fxGuardar = 0

vGrid(Index).Row = vGrid(Index).ActiveRow
vGrid(Index).Col = 1


Dim pId As Long, pMontoMin As Currency, pMontoMax As Currency, pGastos As Currency, pHonorarios As Currency, pImpuesto As Currency
Dim pRangoEdad As String, pEdadMin As Integer, pEdadMax As Integer, pEdadDesc As String, pEstado As String

pId = 0

With vGrid(Index)

Select Case Index
    Case 0
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pMontoMin = CCur(.Text)
      .Col = 3
      pMontoMax = CCur(.Text)
      .Col = 4
      pGastos = CCur(.Text)
      .Col = 5
      pHonorarios = CCur(.Text)
      .Col = 6
      pEstado = Mid(.Text, 1, 1)
    
    Case 2
      .Col = 1
      pId = IIf(.Text = "", 0, .Text)
      .Col = 2
      pRangoEdad = .Text
      .Col = 3
      pEdadMin = .Text
      .Col = 4
      pEdadMax = .Text
      .Col = 5
      pMontoMin = CCur(.Text)
      .Col = 6
      pMontoMax = CCur(.Text)
      .Col = 7
      pEdadDesc = .Text
      .Col = 8
      pEstado = Mid(.Text, 1, 1)
    
End Select

Select Case mTipo
    Case "C", "T"
        strSQL = "exec spCrd_Prea_Config_Prenda_Add " & pId & ", " & pMontoMin & ", " & pMontoMax _
               & ", " & pGastos & ", " & pHonorarios & ", '" & pEstado & "', '" & glogon.Usuario & "', '" & mTipo & "'"
    Case "E"
        strSQL = "exec spCrd_Prea_Config_Examen_Prenda_Requisito_Add " & pId & ", '" & pRangoEdad _
               & "', " & pEdadMin & ", " & pEdadMax & ", " & pMontoMin & ", " & pMontoMax _
               & ", '" & pEdadDesc & "', '" & pEstado & "', '" & glogon.Usuario & "'"
End Select

Call OpenRecordSet(rs, strSQL)
If rs!Pass = 1 Then
  
  pId = rs!IdLlave
  
  .Col = 1
  .Text = rs!IdLlave
  
  Call Bitacora(rs!Movimiento, rs!Mensaje)
  MsgBox rs!Mensaje & ", " & rs!Movimiento & " satisfactoriamente!", vbInformation
Else
   MsgBox rs!Mensaje, vbExclamation
End If


End With

fxGuardar = pId

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 0
        Call sbLista(UCase(Mid(cboTipo.Text, 1, 1)))
    Case 1
        Call sbLista("E")
End Select
End Sub

Private Sub vGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i As Long, MaxCol As Integer, Tabla As String

On Error GoTo vError

Select Case Index
    Case 0
        MaxCol = 6
        Tabla = cboTipo.Text
    Case 2
        MaxCol = 8
        Tabla = "Requisitos de Examenes"
End Select


If vGrid(Index).ActiveCol = MaxCol And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar(Index)
  If i = 0 Then Exit Sub
  vGrid(Index).Row = vGrid(Index).ActiveRow
  If vGrid(Index).MaxRows <= vGrid(Index).ActiveRow Then
    vGrid(Index).MaxRows = vGrid(Index).MaxRows + 1
    vGrid(Index).Row = vGrid(Index).MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid(Index).MaxRows = vGrid(Index).MaxRows + 1
    vGrid(Index).InsertRows vGrid(Index).ActiveRow, 1
    vGrid(Index).Row = vGrid(Index).ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar de " & Tabla, vbYesNo)
     If i = vbYes Then
        vGrid(Index).Row = vGrid(Index).ActiveRow
        vGrid(Index).Col = 1
        
        
        Select Case mTipo
            Case "C", "T"
                strSQL = "exec spCrd_Prea_Config_Prenda_Del '" & vGrid(Index).Text & "', '" & glogon.Usuario & "', '" & mTipo & "'"
            Case "E"
                strSQL = "exec spCrd_Prea_Config_Examen_Prenda_Requisito_Del '" & vGrid(Index).Text & "', '" & glogon.Usuario & "'"
        End Select
        
        
        Call OpenRecordSet(rs, strSQL)
        
        If rs!Pass = 1 Then
                    
            vGrid(Index).Col = 1
            strSQL = vGrid(Index).Text
    
            vGrid(Index).DeleteRows vGrid(Index).ActiveRow, 1
            vGrid(Index).MaxRows = vGrid(Index).MaxRows - 1
            
            If vGrid(Index).MaxRows <= 0 Then
              vGrid(Index).MaxRows = 1
            End If
            
            Call Bitacora(rs!Movimiento, rs!Mensaje)
            
            MsgBox rs!Mensaje & ", Eliminado Satisfactoriamente!", vbInformation
        Else
            MsgBox rs!Mensaje, vbExclamation
        End If


     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



