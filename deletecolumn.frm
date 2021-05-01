VERSION 5.00
Begin VB.Form deletecolumn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete column"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
   Icon            =   "deletecolumn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   555
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "deletecolumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const APPLICATION As String = "Data"

Dim deletecolumncaption_frm As String
Dim deletecolumndelete_frm As String

'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

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

Private Sub Command1_Click()
principal.ListView1.ColumnHeaders.Remove Combo1.ListIndex + 1
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer
Dim table_columns As Integer

deletecolumncaption_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "deletecolumncaption_frm", "Delete column")
deletecolumn.Caption = deletecolumncaption_frm
deletecolumndelete_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "deletecolumndelete_frm", "Delete")
deletecolumn.Command1.Caption = deletecolumndelete_frm

table_columns = principal.ListView1.ColumnHeaders.Count

For x = 1 To table_columns
Combo1.AddItem principal.ListView1.ColumnHeaders(x).Text
Next

If Combo1.ListCount = 0 Then
Command1.Enabled = False
End If
End Sub
