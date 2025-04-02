VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_Genera 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generando..."
   ClientHeight    =   4524
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   11064
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4524
   ScaleWidth      =   11064
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3132
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10812
      _Version        =   1245187
      _ExtentX        =   19071
      _ExtentY        =   5524
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Sorted          =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdCancelar 
      Height          =   732
      Left            =   7800
      TabIndex        =   1
      Top             =   3480
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmTES_Genera.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdGenera 
      Height          =   732
      Left            =   9240
      TabIndex        =   2
      Top             =   3480
      Width           =   1692
      _Version        =   1245187
      _ExtentX        =   2984
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Continuar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmTES_Genera.frx":068A
   End
   Begin VB.Label Label1 
      Caption         =   "Se Generaron Correctamente los Documentos, Desea Continuar?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   3855
   End
End
Attribute VB_Name = "frmTES_Genera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Indicar mediante el valor de la variable gblnContinua que las solicitudes no
'               se imprimieron correctamente.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

gblnContinua = False
Unload Me

End Sub

Private Sub cmdGenera_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Indicar mediante el valor de la variable gblnContinua que las solicitudes se
'               imprimieron correctamente.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

gblnContinua = True
Unload Me
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario. Ademas despliega en
'               pantalla los datos principales de las solicitudes que se estan imprimiendo,
'               las cuales estan contenidas en la variable de arreglo gstrGrid.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'               CentrarFrm - (Centra el formulario dentro del formulario MDI)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim i As Integer, itmX As ListViewItem
Dim frm As Form

On Error GoTo vError
'(.lsw.Tag + 1)

For Each frm In Forms
  If (UCase(frm.Name) = UCase("frmTES_EmisionDocumentos")) Then
    Exit For
  End If
Next frm


With lsw.ColumnHeaders
  .Clear
  .Add , , "No. Solicitud", 1200
  .Add , , "Beneficiario", 3200
  .Add , , "No.Documento", 1400
  .Add , , "Monto", 1400, vbRightJustify
  .Add , , "Fecha", 1400
End With

With frm
 For i = 1 To .lsw.ListItems.Count
    Set itmX = lsw.ListItems.Add(, , .lsw.ListItems(i).Text)
        itmX.SubItems(1) = .lsw.ListItems(i).SubItems(1)
        itmX.SubItems(2) = .lsw.ListItems(i).SubItems(2)
        itmX.SubItems(3) = .lsw.ListItems(i).SubItems(3)
        itmX.SubItems(4) = .lsw.ListItems(i).SubItems(4)
 Next i
 .Hide
End With

Me.Refresh

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

'Private Sub lsw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo vError
'
'    lsw.SortKey = ColumnHeader.Index - 1
'
'    If (lsw.SortOrder = lvwAscending) Then
'        lsw.SortOrder = lvwDescending
'    Else
'        lsw.SortOrder = lvwAscending
'    End If
'
'    lsw.Sorted = True
'    Exit Sub
'
'vError:
'   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical
'
'End Sub
