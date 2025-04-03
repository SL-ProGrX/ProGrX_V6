VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.0#0"; "TTF16.OCX"
Begin VB.Form frmAH_REP 
   Caption         =   "Reporte de Detalle de Socios"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "&Tipo de impresion"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
      Begin VB.OptionButton Optimp 
         Caption         =   "Grid"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Optimp 
         Caption         =   "Cristal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Rep.frx":0000
            Key             =   "Ejecutar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Rep.frx":031C
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Rep.frx":0638
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblrep 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   1535
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ejecutar"
            Key             =   "Ejecutar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9000
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cboMes 
      Height          =   315
      ItemData        =   "frmAH_Rep.frx":0954
      Left            =   2760
      List            =   "frmAH_Rep.frx":097F
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtAno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7223
      _0              =   $"frmAH_Rep.frx":09E8
      _1              =   $"frmAH_Rep.frx":0DF2
      _2              =   $"frmAH_Rep.frx":11FB
      _3              =   $"frmAH_Rep.frx":1604
      _4              =   $"frmAH_Rep.frx":1A0D
      _count          =   5
      _ver            =   2
   End
End
Attribute VB_Name = "frmAH_REP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gnRow As Long, gncol As Long
Dim gencontrado As Boolean
Sub imprimir()
On Error GoTo CapturaError
 'If Optimp.Item(0).Value = True Then
    F1Book1.FilePrint True
 'Else
  'If Optimp.Item(1).Value = True Then
   'frmCC_MenuPrincipal.Crt.ReportFileName = "C:\proyectos\is-ase\ahorro\reportes\ahtiso.rpt"
   'frmCC_MenuPrincipal.Crt.SubreportToChange = strSql
   'frmCC_MenuPrincipal.Crt.PrintReport
  'End If
 'End If
 
    Exit Sub
CapturaError:
 MsgBox "La impresion fue cancelada", 0
 Exit Sub
End Sub
Sub habilita()
tblrep.Buttons.Item(1).Enabled = True 'EJECUTAR
tblrep.Buttons.Item(2).Enabled = True 'IMPRIMIR ojo es falso
tblrep.Buttons.Item(3).Enabled = True 'SALIR
End Sub
Function procesar()
Dim recregs As New ADODB.Recordset, recregs1 As New ADODB.Recordset
Dim strsql As String, strfepro As String
Dim i As Integer, j As Integer, pFound As Long, lngactual As Long
Dim curahant As Currency, curahmes As Currency, cursaldoah As Currency
Dim curapant As Currency, curapmes As Currency, cursaldoap As Currency
Dim curtoahant As Currency, curtoapant As Currency, curtoahmes As Currency, curtoapmes As Currency
Dim curtoahsal As Currency, curtoapsal As Currency
procesar = 0
curtoahant = 0
curtoapant = 0
curtoahmes = 0
curtoapmes = 0
curtoahsal = 0
curtoapsal = 0

On Error GoTo CapturaError

Adodc1.ConnectionString = "driver={SQL Server};" _
            & "uid=sa;server=perseus;database=aseccss"
Adodc1.CommandTimeout = 600
Adodc1.CursorType = adOpenStatic

If cboMes.ListIndex <> -1 And Trim(txtAno.Text) <> "" Then
      strmes = Trim(str(cboMes.ItemData(cboMes.ListIndex)))
      stranio = Trim(txtAno.Text)
      Select Case strmes
       Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            strmes = "0" & Trim(strmes)
       Case 10, 11, 12
            strmes = strmes
      End Select
      strfepro = stranio & strmes 'fecha de variable de combo
      strfepro1 = strmes & "/" & "01" & "/" & stranio
Else
      MsgBox "Introdusca la fecha de proceso", 0
      Exit Function
End If

strsql = "select max(fechaproc) from ahorro_detallado "


Me.MousePointer = 11
strsql = "select socios.cedula,nombre, sum(monto) as ahorro, tipo from AHORRO_DETALLADO,socios WHERE "
strsql = strsql & "ahorro_detallado.cedula=socios.cedula "
strsql = strsql & "and fechaproc<" & strfepro & " "
strsql = strsql & "and (tipo='O' or tipo='P') "
strsql = strsql & "and (estado='A' or estado='J') "
strsql = strsql & "and socios.estadoactual='S' "
strsql = strsql & "group by socios.cedula,socios.nombre,tipo "
strsql = strsql & "order by socios.cedula "
Adodc1.RecordSource = strsql
Adodc1.Refresh


F1Book1.MaxRow = 2
F1Book1.MaxCol = 8
F1Book1.ClearRange -1, -1, -1, -1, F1ClearAll
F1Book1.Row = 1
While Not Adodc1.Recordset.EOF
 With F1Book1
        DoEvents
       .Col = 1
       strCedu = "'" & Trim(Adodc1.Recordset!cedula) & "'"
       F1Book1.Text = strCedu
       .Col = 2
       F1Book1.Text = Adodc1.Recordset!nombre
   If strceduante = Trim(Adodc1.Recordset!cedula) And Adodc1.Recordset!tipo = "P" Then
       .Col = 4
       F1Book1.Text = Adodc1.Recordset!ahorro
       curtoapant = curtoapant + CCur(Adodc1.Recordset!ahorro)
       F1Book1.MaxRow = F1Book1.MaxRow + 1
       .Row = .Row + 1
   Else
       .Col = 3
       F1Book1.Text = Adodc1.Recordset!ahorro
       curtoahant = curtoahant + CCur(Adodc1.Recordset!ahorro)
   End If
       strceduante = Trim(Adodc1.Recordset!cedula)
       Me.Refresh
       Adodc1.Recordset.MoveNext
 End With
Wend

'Adodc1.Recordset.MoveFirst

Adodc1.Recordset.Close
F1Book1.Col = 1
F1Book1.Row = 1

strsql = "select ahorro_detallado.cedula,sum(monto) as monto,tipo,nombre from AHORRO_DETALLADO,SOCIOS WHERE "
strsql = strsql & "socios.estadoactual='S' "
strsql = strsql & "and fechaproc=" & strfepro & " "
strsql = strsql & "and ahorro_detallado.cedula=socios.cedula "
strsql = strsql & "and (tipo='P' or tipo='O') "
strsql = strsql & "and (estado='A' or estado='J') "
strsql = strsql & "group by ahorro_detallado.cedula,monto,tipo,nombre "
strsql = strsql & "order by ahorro_detallado.cedula "
Adodc1.CursorType = adOpenStatic
Adodc1.RecordSource = strsql
Adodc1.Refresh

    
 While Not Adodc1.Recordset.EOF
    gencontrado = False
    Me.Refresh
    DoEvents
    strCedu = "'" & Trim(Adodc1.Recordset!cedula) & "'"
    F1Book1.Find strCedu, 1, 1, 1, F1Book1.LastRow, 1, 0, pFound
    If pFound > 1 Then
     'Si este error se produce la solucion es introducir delimitadores en la cedula
     'al copiarlose en el grid. Pues el metodo find del grid es como un like '%string%'
     MsgBox "Error Cedula: " & strCedu, 0
    End If
    F1Book1.SetActiveCell gnRow, 1
    strCedu = Trim(F1Book1.Text)
    If gencontrado = True Then
        If Adodc1.Recordset!tipo = "O" Then
         F1Book1.Col = 5
         F1Book1.Text = Trim(Adodc1.Recordset!monto)
         curtoahmes = curtoahmes + CCur(Adodc1.Recordset!monto)
         Adodc1.Recordset.MoveNext
        End If
        If strCedu = "'" & Trim(Adodc1.Recordset!cedula) & "'" And Adodc1.Recordset!tipo = "P" Then
         F1Book1.Col = 6
         F1Book1.Text = Trim(Adodc1.Recordset!monto)
         curtoapmes = curtoapmes + CCur(Adodc1.Recordset!monto)
         Adodc1.Recordset.MoveNext
        End If
    Else
         'MsgBox "no encontrado" & strCedu
         lngactual = F1Book1.Row
         'Do While strCedu = "'" & Trim(Adodc1.Recordset!cedula) & "'"
         '   Adodc1.Recordset.MoveNext
         'Loop
         strCedu1 = "'" & Trim(Adodc1.Recordset!cedula) & "'"
         F1Book1.Row = F1Book1.LastRow + 1
         F1Book1.Col = 1
         F1Book1.Text = "'" & Trim(Adodc1.Recordset!cedula) & "'"
         F1Book1.Col = 2
         F1Book1.Text = Adodc1.Recordset!nombre
         F1Book1.Col = 3
         F1Book1.Text = "0"
         F1Book1.Col = 4
         F1Book1.Text = "0"
         If Adodc1.Recordset!tipo = "O" Then
           F1Book1.Col = 5
           F1Book1.Text = Adodc1.Recordset!monto
           Adodc1.Recordset.MoveNext
         End If
         If strCedu1 = "'" & Trim(Adodc1.Recordset!cedula) & "'" And Adodc1.Recordset!tipo = "P" Then
            F1Book1.Col = 6
            F1Book1.Text = Adodc1.Recordset!monto
            Adodc1.Recordset.MoveNext
         End If
         F1Book1.Row = lngactual
    End If
    F1Book1.MaxRow = F1Book1.MaxRow + 1
    F1Book1.Row = F1Book1.Row + 1
Wend

F1Book1.Col = 1
F1Book1.Row = 1
    With F1Book1
        For i = 1 To F1Book1.LastRow
        DoEvents
          Me.Refresh
          .Col = 3
          curahant = CCur(F1Book1.Text)
          .Col = 5
          If Trim(F1Book1.Text) <> "" Then
           curahmes = CCur(F1Book1.Text)
          Else
           curapmes = 0
          End If
          .Col = 7
          cursaldoah = curahant + curahmes
          F1Book1.Text = str(cursaldoah)
          curtoahsal = curtoahsal + cursaldoah
          .Col = 4
          curapant = CCur(F1Book1.Text)
          .Col = 6
          If Trim(F1Book1.Text) <> "" Then
          curapmes = CCur(F1Book1.Text)
          Else
          curapmes = 0
          End If
          .Col = 8
          cursaldoap = curapant + curapmes
          F1Book1.Text = str(cursaldoap)
          curtoapsal = curtoapsal + cursaldoap
          curapant = 0
          curapmes = 0
          cursaldoap = 0
          curahant = 0
          curahmes = 0
          cursaldoah = 0
          F1Book1.MaxRow = F1Book1.MaxRow + 1
          F1Book1.Row = F1Book1.Row + 1
        Next i
    End With
 Me.MousePointer = 1
 
'F1Book1.FilePrint True
'sumatorios
With F1Book1
lngul = F1Book1.LastRow
F1Book1.Row = lngul + 1
.Col = 3
.Text = Format(curtoahant, "#####################.##")
.Col = 4
.Text = Format(curtoapant, "#####################.##")
.Col = 5
.Text = Format(curtoahmes, "#####################.##")
.Col = 6
.Text = Format(curtoapmes, "#####################.##")
.Col = 7
.Text = Format(curtoahsal, "#####################.##")
.Col = 8
.Text = Format(curtoapsal, "#####################.##")

End With
procesar = 1

Exit Function

CapturaError:
 Call ProcedimientoErrores(Me.Name)
 Resume
End Function
Function feproxima(strFecha As String) As String
Dim strmes As String
Dim stranio As String
Dim imes As Integer, ianio As Integer

     feproxima = ""
     stranio = Mid(strFecha, 1, 4)
     strmes = Mid(strFecha, 5, 2)
     ianio = CInt(stranio)
     imes = CInt(strmes)
     If CInt(strmes) = 12 Then
         ianio = ianio + 1
         stranio = Trim(str(ianio))
         strmes = "01"
     Else
         imes = imes + 1
         strmes = "0" & Trim(str(imes))
     End If
     feproxima = Trim(stranio) & Trim(strmes)

End Function
Function feanterior(strFecha As String) As String
Dim strmes As String
Dim stranio As String
Dim imes As Integer, ianio As Integer

     feanterior = ""
     stranio = Mid(strFecha, 1, 4)
     strmes = Mid(strFecha, 5, 2)
     ianio = CInt(stranio)
     imes = CInt(strmes)
     If CInt(strmes) = 1 Then
         ianio = ianio - 1
         stranio = Trim(str(ianio))
         strmes = "12"
     Else
         imes = imes - 1
         strmes = "0" & Trim(str(imes))
     End If
     feanterior = Trim(stranio) & Trim(strmes)
          
End Function

Private Sub cmdrep_Click()

End Sub

Private Sub Command1_Click()
 
End Sub


Private Sub F1Book1_Found(ByVal nSheet As Long, ByVal nRow As Long, ByVal nCol As Long, pCancel As Integer)
gnRow = nRow
gncol = nCol
gencontrado = True
End Sub

Private Sub Form_Load()
Call habilita
End Sub

Private Sub Option1_Click(Index As Integer)

End Sub

Private Sub tblrep_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
 Case "Ejecutar"
      If procesar = 1 Then
        tblrep.Buttons.Item(1).Enabled = False 'EJECUTAR
        tblrep.Buttons.Item(2).Enabled = True 'imprimir
      End If
 Case "Imprimir"
      Call imprimir
      tblrep.Buttons.Item(2).Enabled = True 'imprimir
 Case "Salir"
      Unload Me
 End Select
End Sub

Private Sub txtAno_Change()
Dim strCadena As String, blnNum As Boolean
Dim strmes As String, stranio As String, strfpro As String
'If cboMes.ListIndex <> -1 And Trim(txtAno.Text) <> "" Then
'   tblrep.Buttons.Item(1).Enabled = True
'End If
strCadena = txtAno
If strCadena <> "" Then
   blnNum = IsNumeric(strCadena)
   If blnNum = False Then
     txtAno = Mid(strCadena, 1, Len(strCadena) - 1)
     txtAno.SelStart = Len(txtAno)
   End If
End If
If IsNumeric(strCadena) Then
  If CInt(strCadena) > 2200 Then
     txtAno = Mid(strCadena, 1, Len(strCadena) - 1)
     txtAno.SelStart = Len(txtAno)
  End If

End If

End Sub
