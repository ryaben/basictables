VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form principal 
   Caption         =   "BasicTables"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8655
   Icon            =   "principal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7680
      TabIndex        =   5
      Text            =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Text            =   "1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7680
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Text            =   "English"
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   6240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6135
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4471
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "by Rama Studios"
            TextSave        =   "by Rama Studios"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu newtable 
         Caption         =   "&New table"
         Shortcut        =   ^N
      End
      Begin VB.Menu opentable 
         Caption         =   "&Open table"
         Shortcut        =   ^O
      End
      Begin VB.Menu savetable 
         Caption         =   "&Save table"
         Shortcut        =   ^S
      End
      Begin VB.Menu savetableas 
         Caption         =   "Save table &as ..."
         Shortcut        =   ^A
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu printtable 
         Caption         =   "&Print table"
         Shortcut        =   ^P
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu closebtn 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu newcolumn 
         Caption         =   "Add &column"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu newentrybtn 
         Caption         =   "Add &entry"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu editentrybtn 
         Caption         =   "E&dit selected entry"
         Shortcut        =   ^E
      End
      Begin VB.Menu editcolumntitlebtn 
         Caption         =   "Edit column &title"
         Shortcut        =   ^T
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
      End
      Begin VB.Menu deletecolumnbtn 
         Caption         =   "Delete co&lumn"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu deleteentry 
         Caption         =   "Delete &selected entry"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu deleteall 
         Caption         =   "Delete &all entries"
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu optionsbtn 
         Caption         =   "&Options"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu aboutbtn 
         Caption         =   "&About ..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu officialsite 
         Caption         =   "&Official site"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu status 
      Caption         =   "Status: Unsaved"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable para los subitems del listview
Dim item As ListItem

Private Const WM_SETREDRAW As Long = &HB&
' DEclaración de la Función Api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1

Const APPLICATION As String = "Data"

'Variables de datos generales
Dim form_state As Single
Dim backgroundcolor As String
Dim backgroundindex As String
Dim fontcolor As String
Dim fontindex As String
Dim gridlines As Single
Dim fonttype As String
Dim language As String

'Variables de msgbox, inputbox y commondialog
Dim opentabletitle_comdialog As String
Dim savetabletitle_comdialog As String
Dim savetableastitle_comdialog As String
Dim error_msgbox As String
Dim newtabletitle As String
Dim newtableprompt As String
Dim newcolumntitle As String
Dim newcolumnprompt As String
Dim deleteentrytitle As String
Dim deleteentryprompt As String
Dim deleteentriestitle As String
Dim deleteentriesprompt As String
 
'Variables de lenguajes
Dim file_menu As String
Dim edit_menu As String
Dim tools_menu As String
Dim help_menu As String
'--------------------------
Dim columnsamount As String
Dim entriesamount As String
Dim newtable_menu As String
Dim opentable_menu As String
Dim savetable_menu As String
Dim savetableas_menu As String
Dim printtable_menu As String
Dim close_menu As String
Dim addcolumn_menu As String
Dim addentry_menu As String
Dim editentry_menu As String
Dim editcolumn_menu As String
Dim deletecolumn_menu As String
Dim deleteentry_menu As String
Dim deleteallentries_menu As String
Dim options_menu As String
Dim about_menu As String
Dim officialsite_menu As String
Dim byramastudios As String
Dim tableunsaved As String


'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
    
Private Function Grabar_Ini(Path_INI As String, Key As String, Valor As Variant) As String

WritePrivateProfileString APPLICATION, _
                                         Key, _
                                         Valor, _
                                         Path_INI

End Function

Private Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        Leer_Ini = Left$(bufer, Len_Value)

End Function

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

Public Sub Importar_ListView(ListView As ListView, PathArchivo As String)

    On Error GoTo errsub
    
    Dim Linea As String, x As Integer, i As Integer, sStr() As String, it As Long
    
    'Opcional

    ListView.ListItems.Clear
    ListView.ColumnHeaders.Clear
    ListView.View = lvwReport
    
    'Abre el archivo para leer los datos
    Open PathArchivo For Input As #1
    'Leemos la primer linea que es el de los encabezados
    Line Input #1, Linea
    
    
    sStr = Split(Linea, vbTab)
    'Agregamos el texto en los encabezados del ListView
    
'*****************BUCLE AGREGADO***************************
' Esto crea las columnas necesarias en el listView
    For i = 1 To UBound(sStr)
        ListView.ColumnHeaders.Add
    Next
'******************FIN*********************************
    
    For i = LBound(sStr) To UBound(sStr) - 1
        ListView.ColumnHeaders(i + 1).Text = sStr(i)
    Next
    
    'Recorremos todo el archivo de texto
    While Not EOF(1)
    
    'Leemos la siguientes lineas del archivo
    Line Input #1, Linea
    
    
    'Separamos los datos
    sStr = Split(Linea, vbTab)
    'Agregamos el Item
    ListView.ListItems.Add , , sStr(LBound(sStr))
    it = it + 1
    For i = LBound(sStr) To UBound(sStr) - 1
        'Agregamos el Subitem
        ListView.ListItems(it).ListSubItems.Add , , sStr(i + 1)
        
    Next
    
    Wend
    'Cerramos el archivo abierto
    Close

Exit Sub
errsub:
MsgBox Err.Description, vbCritical

End Sub

Public Sub Exportar_ListView(ListView As ListView, PathArchivo As String)
    On Error GoTo errsub
    Dim Linea As String, x As Integer, i As Integer
    
    'Abrimos un archivo para guardar los datos del ListView
    Open PathArchivo For Output As #1
    
    'Recorremos los encabezados para guardar el caption
    For i = 1 To ListView1.ColumnHeaders.Count
        Linea = Linea & ListView1.ColumnHeaders(i).Text & vbTab
    Next
    'Imprimimos la línea
    Print #1, Linea
    
    'recorremos cada Item y Subitem
    For i = 1 To ListView.ListItems.Count
        'texto del Item
        Linea = ListView.ListItems(i) & vbTab
        'texto de los SubItems
        For x = 1 To ListView1.ColumnHeaders.Count - 1
            Linea = Linea & ListView.ListItems.item(i).SubItems(x) & vbTab
        Next
    'Imprimimos la linea
    Print #1, Linea
    Next
    
    'Cerramos
    Close

Exit Sub
errsub:
MsgBox Err.Description, vbCritical

End Sub

'A esta función se le envía el control LV a imprimir
Public Sub Imprimir_ListView(ListView As ListView)

Dim i As Integer, AnchoCol As Single, Espacio As Integer, x As Integer
  
  AnchoCol = 0
  'Recorremos desde la primer columna hasta la última para almacenar el ancho total
  For i = 1 To ListView.ColumnHeaders.Count
     AnchoCol = AnchoCol + ListView.ColumnHeaders(i).Width
  Next
  
  Espacio = 0
  
  Printer.Print
  
  'Imprime una línea
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  
  With ListView
  
  'Acá se imprimen los encabezados del ListView
  For i = 1 To .ColumnHeaders.Count
      Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
      Printer.Print ListView.ColumnHeaders(i).Text;
      Printer.CurrentX = Espacio
  Next

  Printer.Print
  
  'Imprime una línea
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  
  'Imprime Línea en blanco
  Printer.Print
  
  'Este bucle recorre los items y subitems del ListView  y los imprime
  For i = 1 To .ListItems.Count
       Espacio = 0
       
       Set item = .ListItems(i)
       Printer.Print item.Text;
       'Recorremos las columnas
       For x = 1 To .ColumnHeaders.Count - 1
             Espacio = Espacio + CInt(.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
             Printer.CurrentX = Espacio
             Printer.Print item.SubItems(x);
       Next
       
       'Otro espacio en blanco
       Printer.Print
  Next
  
  End With
  
  Printer.Print
  'Imprime la línea de final de impresión
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  Printer.Print

  
  'Comenzamos la impresión
  Printer.EndDoc
End Sub

Private Sub aboutbtn_Click()
about.Show 1
End Sub

Private Sub closebtn_Click()
End
End Sub

Private Sub deleteall_Click()
If ListView1.ListItems.Count = 0 Then
MsgBox error_msgbox, vbExclamation
Else
If MsgBox(deleteentriesprompt, vbYesNo + vbExclamation, deleteentriestitle) = vbYes Then
ListView1.ListItems.Clear
End If
End If
End Sub

Private Sub deletecolumnbtn_Click()
deletecolumn.Show 1
End Sub

Private Sub deleteentry_Click()
If ListView1.ListItems.Count = 0 Then
MsgBox error_msgbox, vbExclamation
Else
If MsgBox(deleteentryprompt, vbYesNo + vbExclamation, deleteentrytitle) = vbYes Then
ListView1.ListItems.Remove ListView1.SelectedItem.Index
End If
End If
End Sub

Private Sub editcolumntitlebtn_Click()
editcolumn.Show 1
End Sub

Private Sub editentrybtn_Click()
If ListView1.ListItems.Count = 0 Then
MsgBox error_msgbox, vbExclamation
Else
editentry.Show 1
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim fso As Object

form_state = Leer_Ini(App.Path & "\config.ini", "State", 0)
backgroundcolor = Leer_Ini(App.Path & "\config.ini", "Background", &HFFFFFF)
backgroundindex = Leer_Ini(App.Path & "\config.ini", "Background index", 0)
fontcolor = Leer_Ini(App.Path & "\config.ini", "Font", &H0&)
fontindex = Leer_Ini(App.Path & "\config.ini", "Font index", 1)
gridlines = Leer_Ini(App.Path & "\config.ini", "Gridlines", 1)
fonttype = Leer_Ini(App.Path & "\config.ini", "Font2", "MS Sans Serif")
language = Leer_Ini(App.Path & "\config.ini", "Language", "English")

Me.WindowState = form_state
ListView1.BackColor = backgroundcolor
Text2.Text = backgroundindex
ListView1.ForeColor = fontcolor
Text3.Text = fontindex
ListView1.gridlines = gridlines
ListView1.Font = fonttype
Text1.Text = language

file_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "file_menu", "File")
file.Caption = file_menu
edit_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "edit_menu", "Edit")
edit.Caption = edit_menu
tools_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "tools_menu", "Tools")
tools.Caption = tools_menu
help_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "help_menu", "Help")
help.Caption = help_menu
'--------------------------------------------------------------------------------------
opentabletitle_comdialog = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "opentabletitle_comdialog", "Open table")
savetabletitle_comdialog = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "savetabletitle_comdialog", "Save table")
savetableastitle_comdialog = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "savetableastitle_comdialog", "Save table as ...")
error_msgbox = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "error_msgbox", "There are no items in the table")
newtabletitle = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "newtabletitle", "New table")
newtableprompt = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "newtableprompt", "Are you sure you want to start a new table?")
newcolumntitle = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "newcolumntitle", "Add column")
newcolumnprompt = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "newcolumnprompt", "Choose a title name for the new column")
deleteentrytitle = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deleteentrytitle", "Delete entry")
deleteentryprompt = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deleteentryprompt", "Are you sure you want to delete selected entry?")
deleteentriestitle = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deleteentriestitle", "Delete all entries")
deleteentriesprompt = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deleteentriesprompt", "Are you sure you want to delete all entries?")
'--------------------------------------------------------------------------------------
columnsamount = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "columnsamount", "Columns:")
entriesamount = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "entriesamount", "Entries:")
newtable_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "newtable_menu", "New table")
newtable.Caption = newtable_menu
opentable_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "opentable_menu", "Open table")
opentable.Caption = opentable_menu
savetable_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "savetable_menu", "Save table")
savetable.Caption = savetable_menu
printtable_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "printtable_menu", "Print table")
printtable.Caption = printtable_menu
close_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "close_menu", "Close")
closebtn.Caption = close_menu
addcolumn_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "addcolumn_menu", "Add column")
newcolumn.Caption = addcolumn_menu
addentry_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "addentry_menu", "Add entry")
newentrybtn.Caption = addentry_menu
editentry_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "editentry_menu", "Edit selected entry")
editentrybtn.Caption = editentry_menu
editcolumn_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "editcolumncaption_frm", "Edit column title")
editcolumntitlebtn.Caption = editcolumn_menu
savetableas_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "savetableastitle_comdialog", "Save table as ...")
savetableas.Caption = savetableas_menu
deletecolumn_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deletecolumn_menu", "Delete column")
deletecolumnbtn.Caption = deletecolumn_menu
deleteentry_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deleteentry_menu", "Delete selected entry")
deleteentry.Caption = deleteentry_menu
deleteallentries_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "deleteallentries_menu", "Delete all entries")
deleteall.Caption = deleteallentries_menu
options_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "options_menu", "Options")
optionsbtn.Caption = options_menu
about_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "about_menu", "About")
aboutbtn.Caption = about_menu
officialsite_menu = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "officialsite_menu", "Official site")
officialsite.Caption = officialsite_menu
byramastudios = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "byramastudios", "by Rama Studios")
StatusBar1.Panels(5).Text = byramastudios
tableunsaved = Leer_Ini(App.Path & "\Languages\" & Text1 & ".lng", "tableunsaved", "Table not saved")
StatusBar1.Panels(3).Text = tableunsaved

StatusBar1.Panels(4).Text = "Table1.bdt"

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(App.Path & "\Languages\" & Text1 & ".lng") Then
options.Label6.Caption = ""
ElseIf fso.FileExists(App.Path & "\Languages\" & Text1 & ".lng") = False And Text1.Text <> "English" Then
options.Label6.Caption = "Error on load"
options.Label6.Visible = True
End If

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
ListView1.Height = Me.Height - 1220
ListView1.Width = Me.Width - 100
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Grabar_Ini(App.Path & "\config.ini", "State", Me.WindowState)
Call Grabar_Ini(App.Path & "\config.ini", "Background", ListView1.BackColor)
Call Grabar_Ini(App.Path & "\config.ini", "Background index", Text2.Text)
Call Grabar_Ini(App.Path & "\config.ini", "Font", ListView1.ForeColor)
Call Grabar_Ini(App.Path & "\config.ini", "Font index", Text3.Text)
Call Grabar_Ini(App.Path & "\config.ini", "Gridlines", Text4.Text)
Call Grabar_Ini(App.Path & "\config.ini", "Font2", ListView1.Font)
Call Grabar_Ini(App.Path & "\config.ini", "Language", Text1.Text)
End
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    On Error Resume Next
    
    
    With ListView1
    
        Dim i As Long
        Dim Formato As String
        Dim strData() As String
        
        Dim Columna As Long
        
        Call SendMessage(Me.hwnd, WM_SETREDRAW, 0&, 0&)
        
        
        Columna = ColumnHeader.Index - 1
        
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Tipo de dato a ordenar
        ''''''''''''''''''''''''''''''''''''''''''''''
        
        Select Case UCase$(ColumnHeader.Tag)
    
        
        ' Fecha
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "DATE"
        
            Formato = "YYYYMMDDHhNnSs"
        
            ' Ordena alfabéticamente la columna con Fechas _
              ( es la columna que tiene en el tag el valor DATE )
        
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    Formato)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
            
        ' Datos de numéricos
        '''''''''''''''''''''''''''''''''''''''''''''
        Case "NUMBER"
        
            ' Ordena alfabéticamente la columna con números _
              ( es la columna que tiene en el tag el valor NUMBER )
        
            Formato = String(30, "0") & "." & String(30, "0")
                
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        Formato)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        Formato))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next i
                End If
            End With
            
            ' Ordena alfabéticamente
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                If (Columna > 0) Then
                    For i = 1 To .Count
                        With .item(i).ListSubItems(Columna)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                Else
                    For i = 1 To .Count
                        With .item(i)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next i
                End If
            End With
        
        Case Else
                    
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        End Select
    
    End With
    
    Call SendMessage(Me.hwnd, WM_SETREDRAW, 1&, 0&)
    ListView1.Refresh
    
End Sub

Private Sub newcolumn_Click()
Dim inputcolumn As String
inputcolumn = InputBox(newcolumnprompt, newcolumntitle)
If inputcolumn <> "" Then
ListView1.ColumnHeaders.Add , , inputcolumn
End If
End Sub

Private Sub newentrybtn_Click()
newentry.Show 1
End Sub

Private Sub newtable_Click()
If MsgBox(newtableprompt, vbYesNo + vbExclamation, newtabletitle) = vbYes Then
ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear
StatusBar1.Panels(3).Text = tableunsaved
StatusBar1.Panels(4).Text = "Table1.bdt"
End If
End Sub

Private Sub officialsite_Click()
Dim Z
Z = ShellExecute(Me.hwnd, "Open", "http://adf.ly/Kk0PI", &O0, &O0, SW_NORMAL)
End Sub

Private Sub opentable_Click()
On Error Resume Next
With CommonDialog1
.DialogTitle = opentabletitle_comdialog
.FileName = ""
.InitDir = App.Path
.Filter = "Basic Data tables (*.bdt)|*.bdt|All files|*.*|"
.ShowOpen
If .FileName <> "" Then
Call Importar_ListView(ListView1, .FileName)
StatusBar1.Panels(3).Text = .FileName
StatusBar1.Panels(4).Text = .FileTitle
End If
End With
Set ListView1.SelectedItem = ListView1.ListItems(1)
End Sub

Private Sub optionsbtn_Click()
options.Show 1
End Sub

Private Sub printtable_Click()
'Le enviamos el control ListView como parámetro
Imprimir_ListView ListView1
End Sub

Private Sub savetable_Click()
On Error Resume Next
If StatusBar1.Panels(3).Text = tableunsaved Then
With CommonDialog1
.DialogTitle = savetabletitle_comdialog
.FileName = ""
.InitDir = App.Path
.Filter = "Basic Data tables (*.bdt)|*.bdt|All files|*.*|"
.ShowSave
If .FileName <> "" Then
Call Exportar_ListView(ListView1, .FileName)
StatusBar1.Panels(3).Text = .FileName
StatusBar1.Panels(4).Text = .FileTitle
End If
End With
Else
Call Exportar_ListView(ListView1, StatusBar1.Panels(3).Text)
End If
End Sub

Private Sub savetableas_Click()
With CommonDialog1
.DialogTitle = savetableastitle_comdialog
.FileName = ""
.InitDir = App.Path
.Filter = "Basic Data tables (*.bdt)|*.bdt|All files|*.*|"
.ShowSave
If .FileName <> "" Then
Call Exportar_ListView(ListView1, .FileName)
StatusBar1.Panels(3).Text = .FileName
StatusBar1.Panels(4).Text = .FileTitle
End If
End With
End Sub

Private Sub status_Click()

End Sub

Private Sub Timer1_Timer()
Me.Caption = "BasicTables - " & StatusBar1.Panels(4).Text
StatusBar1.Panels(1).Text = columnsamount & " " & ListView1.ColumnHeaders.Count
StatusBar1.Panels(2).Text = entriesamount & " " & ListView1.ListItems.Count
End Sub
