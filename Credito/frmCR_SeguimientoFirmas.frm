VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_SeguimientoFirmas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Firmas Deudor, Fiadores y Co-Deudores"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9255
   HelpContextID   =   3025
   Icon            =   "frmCR_SeguimientoFirmas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   1812
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   8772
      _Version        =   524288
      _ExtentX        =   15473
      _ExtentY        =   3196
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
      MaxCols         =   497
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_SeguimientoFirmas.frx":030A
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Firmas"
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
      Height          =   492
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   4812
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCR_SeguimientoFirmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbGuardar()
Dim i As Integer, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With vGrid
    For i = 1 To .MaxRows
      .Row = i
      .col = 1
      If .Text = "Deudor" Then
        .col = 4
        strSQL = "update reg_creditos set firma_deudor = " & .Value & ",fechaforf = dbo.MyGetdate()" _
               & " where id_solicitud = " & Operacion.Operacion
      
      Else
         .col = 4
         strSQL = "update fiadores set firma = '" & IIf((.Value = 1), "S", "N") _
                & "' where ID_SOLiCITUD = " & Operacion.Operacion & " and CedulaF = '"
                
         .col = 2
         strSQL = strSQL & .Text & "'"
      End If
      Call ConectionExecute(strSQL)
    
    
    Next i
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

Me.MousePointer = vbHourglass


strSQL = "select 'Deudor' as Tipo,R.cedula,S.nombre,isnull(R.firma_deudor,0) as Firma" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " Where R.ID_SOLICITUD = " & Operacion.Operacion _
       & " Union " _
       & " select 'Fiador' as Tipo,S.cedula,S.nombre,case when F.firma = 'N' then 0 else 1 end as Firma" _
       & " from fiadores F inner join Socios S on F.cedulaf = S.cedula" _
       & " where F.estado = 'A' and F.id_solicitud = " & Operacion.Operacion
Call sbCargaGrid(vGrid, 4, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call sbGuardar
End Sub
