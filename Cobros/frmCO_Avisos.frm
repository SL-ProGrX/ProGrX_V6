VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCO_Avisos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Avisos"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   6732
   Icon            =   "frmCO_Avisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2544
   ScaleWidth      =   6732
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   960
      MaxLength       =   6
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame fraAviso2 
      Height          =   735
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
      Begin VB.ComboBox cboCuotasAviso2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCO_Avisos.frx":000C
         Left            =   840
         List            =   "frmCO_Avisos.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCuotasAviso2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cuotas"
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.OptionButton optAvisos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Primer Aviso"
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
      Height          =   615
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optAvisos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Segundo Aviso"
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
      Height          =   615
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame fraAviso1 
      Height          =   735
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      Begin VB.ComboBox cboCuotasAviso1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCO_Avisos.frx":0028
         Left            =   840
         List            =   "frmCO_Avisos.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCuotasAviso1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCuota 
         Caption         =   "Cuotas"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   168
      Left            =   0
      TabIndex        =   1
      Top             =   2376
      Visible         =   0   'False
      Width           =   6732
      _ExtentX        =   11875
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar tlbAviso 
      Height          =   456
      Left            =   4800
      TabIndex        =   0
      Top             =   1560
      Width           =   1896
      _ExtentX        =   3344
      _ExtentY        =   804
      ButtonWidth     =   2879
      ButtonHeight    =   804
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Crear &Avisos"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   4200
      Top             =   960
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Avisos.frx":0044
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   192
      Left            =   372
      TabIndex        =   14
      Top             =   120
      Width           =   456
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmCO_Avisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vAplica As Boolean

Function fxAvisoPoliza(vOP As Long) As String
Dim rs As New ADODB.Recordset
Dim strSQL As String

strSQL = "Select isnull(Count(*),0) as Cuotas From Morosidad Where id_solicitud=" & vOP
strSQL = strSQL & " And Estado='A'"

With rs
  .Open strSQL, glogon.Conection, adOpenStatic
     If !Cuotas > 0 Then
       fxAvisoPoliza = "El estado de su Poliza saldo deudor se encuentra atrasada."
     Else
       fxAvisoPoliza = "El estado de su Poliza saldo deudor se encuentra al día."
     End If
  .Close
End With

End Function

Private Sub Form_Load()
Set Me.Icon = imgLista.ListImages(1).Picture
cboCuotasAviso1.Text = "="
txtCuotasAviso1.Text = 1

cboCuotasAviso2.Text = ">="
txtCuotasAviso2.Text = 2

End Sub

Private Sub optAvisos_Click(Index As Integer)
Select Case Index
  Case 0
    fraAviso2.Visible = False
    fraAviso1.Visible = True
  Case 1
    fraAviso2.Visible = True
    fraAviso1.Visible = False
End Select
End Sub

Private Sub tlbAviso_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rs As New ADODB.Recordset, strSQL As String, x As Integer
Dim strFecha As String, vMes As Integer, strRuta As String

Dim rsTmp As New ADODB.Recordset, vArchivo1 As String, vArchivo2 As String

If Trim(txtCodigo) = "" Then
   MsgBox "Debe Especificar Un Código", vbExclamation
   Exit Sub
ElseIf vAplica = False Then
   MsgBox "Este Código No Aplica Avisos", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

strFecha = Format(fxFechaServidor, "yyyy/mm/dd")

strSQL = "select * from par_ahcr"
Call OpenRecordSet(rs, strSQL)
vMes = Mid(GLOBALES.glngFechaCR, 5, 2)
If rs!cr_apl = 0 Then
 If vMes = 1 Then
   vMes = 12
 Else
   vMes = vMes - 1
 End If
End If
rs.Close



Select Case True
    Case optAvisos(0).Value
         x = MsgBox("Esta seguro que desea crear Primer Aviso", vbYesNo)
         If x = vbNo Then Exit Sub
         
         x = 1
         vArchivo1 = "CbrPrimerAviso.rpt"
         vArchivo2 = "CbrPrimerAvisoFiador.rpt"
         
         strSQL = " And V.Cuota " & cboCuotasAviso1 & " " & txtCuotasAviso1
         
    Case optAvisos(1)
         x = MsgBox("Esta seguro que desea crear Segundo Aviso", vbYesNo)

         If x = vbNo Then Exit Sub
         
         strSQL = " And V.Cuota " & cboCuotasAviso2 & " " & txtCuotasAviso2
         
         x = 2
         vArchivo1 = "CbrSegundoAviso.rpt"
         vArchivo2 = "CbrSegundoAvisoFiador.rpt"
End Select
 
 
With frmContenedor.Crt

   strSQL = "Select R.id_solicitud,R.poliza,R.garantia,R.codigo" _
          & " From Vista_Morosidad V inner join Reg_Creditos R on V.id_solicitud = R.id_solicitud" _
          & " Where R.Codigo='" & Trim(txtCodigo) & "' and R.proceso <> 'J'" & strSQL
   rs.Open strSQL, glogon.Conection
   Do While Not rs.EOF
       .Reset
       .Destination = crptToPrinter
       
       .Connect = glogon.ConectRPT
       
       .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
       .ReportFileName = strRuta & vArchivo1
       .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & rs!id_solicitud
       
       If rs!Garantia = "F" Then
           .Formulas(1) = "Copia='cc. Fiadores'"
        If IsNull(rs!Poliza) Then
           .Formulas(2) = "Poliza='El estado de su Poliza saldo deudor se encuentra al día.'"
        Else
           .Formulas(2) = "Poliza='" & fxAvisoPoliza(!Poliza) & "'"
        End If
       End If
       frmContenedor.Crt.PrintReport

       If rs!Garantia = "F" Then
          strSQL = "Select fia_consec From Fiadores Where Id_solicitud=" & rs!id_solicitud
          Call OpenRecordSet(rsTmp, strSQL, 0)
            Do While Not rsTmp.EOF
               .Reset
               
               .Connect = glogon.ConectRPT
               
               .ReportFileName = SIFGlobal.fxPathReportes(vArchivo2)
               .Destination = crptToPrinter
               .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
               .SelectionFormula = "{FIADORES.FIA_CONSEC} = " & rsTmp!fia_consec
               .PrintReport
               rsTmp.MoveNext
            Loop
          rsTmp.Close
       End If


       strSQL = "Insert Cbr_Avisos(ID_SOLICITUD,Codigo,TIPO_AVISO," _
              & "FECHA_AVISO) Values(" & rs!id_solicitud & ",'" _
              & rs!Codigo & "'," & x & ",dbo.MyGetdate())"
       Call ConectionExecute(strSQL)
       
       rs.MoveNext
    Loop
  rs.Close

End With

Me.MousePointer = vbDefault
MsgBox "Avisos realizados satisfactoriamente ...", vbExclamation

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 13
   txtCodigo_LostFocus
End Select
End Sub


Private Sub txtCodigo_LostFocus()
Dim rs As New ADODB.Recordset
Dim strSQL As String

If Trim(txtCodigo) <> "" Then
   strSQL = "Select * from Catalogo where Codigo='" & Trim(txtCodigo) & "'"
   With rs
     .Open strSQL, glogon.Conection, adOpenStatic
       If .EOF = False Then
          lblDescripcion = Trim(!Descripcion)
          If !Poliza = "S" Or !retencion = "S" Then
             vAplica = False
          Else
             vAplica = True
          End If
       Else
          txtCodigo = ""
          lblDescripcion = ""
          txtCodigo.SetFocus
       End If
     .Close
   End With
Else
   txtCodigo = ""
   lblDescripcion = ""
   txtCodigo.SetFocus
End If

End Sub


Private Sub txtCuotasAviso1_Change()
If Trim(txtCuotasAviso1) = "" Then
   txtCuotasAviso1 = 1
End If
End Sub


Private Sub txtCuotasAviso2_Change()
If Trim(txtCuotasAviso2) = "" Then
   txtCuotasAviso2 = 1
End If
End Sub


