VERSION 5.00
Begin VB.Form editcolumn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit column title"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Default         =   -1  'True
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "editcolumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const APPLICATION As String = "Data"

Dim editcolumncaption_frm As String
Dim editcolumnedit_frm As String


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
principal.ListView1.ColumnHeaders(Combo1.ListIndex + 1).Text = Text1.Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim x As Integer
Dim table_columns As Integer

editcolumncaption_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editcolumncaption_frm", "Edit column title")
editcolumn.Caption = editcolumncaption_frm
editcolumnedit_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editcolumnedit_frm", "Edit")
editcolumn.Command1.Caption = editcolumnedit_frm

table_columns = principal.ListView1.ColumnHeaders.Count

For x = 1 To table_columns
Combo1.AddItem principal.ListView1.ColumnHeaders(x).Text
Next

Combo1.Text = principal.ListView1.ColumnHeaders(1).Text

If Combo1.ListCount = 0 Then
Command1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
If Text1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
