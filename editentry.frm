VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form editentry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit entry"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   Icon            =   "editentry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Edit entry"
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set field text"
      Default         =   -1  'True
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Text"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Column"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Text"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "editentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const APPLICATION As String = "Data"

Dim editentrycaption_frm As String
Dim editentrycolumnlabel_frm As String
Dim editentrytextlabel_frm As String
Dim editentryset_frm As String
Dim editentrycancel_frm As String
Dim editentryedit_frm As String

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
Dim y As Integer
Dim items_entry As Integer

items_entry = ListView1.ListItems.Count

For y = 1 To items_entry
If y = 1 Then
principal.ListView1.SelectedItem.Text = ListView1.ListItems(y).ListSubItems(1).Text
Else
principal.ListView1.SelectedItem.ListSubItems(y - 1) = ListView1.ListItems(y).ListSubItems(1).Text
End If
Next

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo erraction
ListView1.ListItems(Combo1.ListIndex + 1).ListSubItems(1).Text = Text1
Combo1.ListIndex = Combo1.ListIndex + 1
Text1 = ListView1.ListItems(Combo1.ListIndex + 1).ListSubItems(1).Text
Exit Sub
erraction:
Combo1.ListIndex = 0
Text1 = ListView1.ListItems(Combo1.ListIndex + 1).ListSubItems(1).Text
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim subelemento As ListItem
Dim x As Integer
Dim table_items As Integer

editentrycaption_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editentrycaption_frm", "Edit entry")
editentry.Caption = editentrycaption_frm
editentrycolumnlabel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editentrycolumnlabel_frm", "Column")
editentry.Label1.Caption = editentrycolumnlabel_frm
editentry.ListView1.ColumnHeaders(1).Text = editentrycolumnlabel_frm
editentrytextlabel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editentrytextlabel_frm", "Text")
editentry.Label2.Caption = editentrytextlabel_frm
editentry.ListView1.ColumnHeaders(2).Text = editentrytextlabel_frm
editentryset_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editentryset_frm", "Set field text")
editentry.Command3.Caption = editentryset_frm
editentrycancel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editentrycancel_frm", "Cancel")
editentry.Command2.Caption = editentrycancel_frm
editentryedit_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "editentryedit_frm", "Edit entry")
editentry.Command1.Caption = editentryedit_frm

table_items = principal.ListView1.ColumnHeaders.Count

For x = 1 To table_items
Combo1.AddItem principal.ListView1.ColumnHeaders(x).Text
Set subelemento = ListView1.ListItems.Add(, , principal.ListView1.ColumnHeaders(x).Text)
If x = 1 Then
subelemento.SubItems(1) = principal.ListView1.SelectedItem.Text
Else
subelemento.SubItems(1) = principal.ListView1.SelectedItem.ListSubItems(x - 1).Text
End If
Next

Combo1.Text = principal.ListView1.ColumnHeaders(1).Text
Text1.Text = ListView1.ListItems(1).ListSubItems(1).Text

End Sub
