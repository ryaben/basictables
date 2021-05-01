VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3255
   Icon            =   "options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLanguage 
      Height          =   1935
      Left            =   120
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "options.frx":48FA
         Left            =   120
         List            =   "options.frx":4910
         TabIndex        =   28
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "*Language reset requires restart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Interface language:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame frmFont 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   2400
         Top             =   1800
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "options.frx":494E
         Left            =   1320
         List            =   "options.frx":495E
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "White"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Black"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Grey"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Yellow"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Red"
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Green"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Orange"
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Pink"
         ForeColor       =   &H00FFC0FF&
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Font type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Font color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   2775
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   3000
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame frmBackground 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   1200
         Top             =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Pink"
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Orange"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Green"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Red"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Blue"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gridlines"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Yellow"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grey"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "Black"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "White"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3000
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Table color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Background"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Language"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const APPLICATION As String = "Data"

Dim optionsbackgroundcaption_strip As String
Dim optionsfontcaption_strip As String
Dim optionslanguagecaption_strip As String
'-----------------------------------------
Dim optionsaccept_frm As String
Dim optionscaption_frm As String
'-----------------------------------------
Dim optionsbackgroundlabel_frm As String
Dim optionsbackgroundgridlines_frm As String
'-----------------------------------------
Dim optionswhite_frm As String
Dim optionsblack_frm As String
Dim optionsgray_frm As String
Dim optionsyellow_frm As String
Dim optionsblue_frm As String
Dim optionsred_frm As String
Dim optionsgreen_frm As String
Dim optionsorange_frm As String
Dim optionspink_frm As String
'-----------------------------------------
Dim optionsfontcolorlabel_frm As String
Dim optionsfonttypelabel_frm As String
'-----------------------------------------
Dim optionslanguagelabel_frm As String
Dim optionslanguagereset_frm As String

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
Unload Me
End Sub

Private Sub Form_Load()
optionsbackgroundcaption_strip = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsbackgroundcaption_strip", "Background")
options.TabStrip1.Tabs(1).Caption = optionsbackgroundcaption_strip
optionsfontcaption_strip = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsfontcaption_strip", "Font")
options.TabStrip1.Tabs(2).Caption = optionsfontcaption_strip
optionslanguagecaption_strip = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionslanguagecaption_strip", "Language")
options.TabStrip1.Tabs(3).Caption = optionslanguagecaption_strip
optionsaccept_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsaccept_frm", "Accept")
options.Command1.Caption = optionsaccept_frm
optionscaption_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optioncaption_frm", "Options")
options.Caption = optionscaption_frm
'-----------------------------------------------------------------
optionsbackgroundlabel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsbackgroundlabel_frm", "Table color:")
options.Label1 = optionsbackgroundlabel_frm
optionswhite_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionswhite_frm", "White")
options.Option1(0).Caption = optionswhite_frm
options.Option2(0).Caption = optionswhite_frm
optionsblack_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsblack_frm", "Black")
options.Option1(1).Caption = optionsblack_frm
options.Option2(1).Caption = optionsblack_frm
optionsgray_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsgray_frm", "Gray")
options.Option1(2).Caption = optionsgray_frm
options.Option2(2).Caption = optionsgray_frm
optionsyellow_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsyellow_frm", "Yellow")
options.Option1(3).Caption = optionsyellow_frm
options.Option2(3).Caption = optionsyellow_frm
optionsblue_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsblue_frm", "Blue")
options.Option1(4).Caption = optionsblue_frm
options.Option2(4).Caption = optionsblue_frm
optionsred_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsred_frm", "Red")
options.Option1(5).Caption = optionsred_frm
options.Option2(5).Caption = optionsred_frm
optionsgreen_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsgreen_frm", "Green")
options.Option1(6).Caption = optionsgreen_frm
options.Option2(6).Caption = optionsgreen_frm
optionsorange_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsorange_frm", "Orange")
options.Option1(7).Caption = optionsorange_frm
options.Option2(7).Caption = optionsorange_frm
optionspink_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionspink_frm", "Pink")
options.Option1(8).Caption = optionspink_frm
options.Option2(8).Caption = optionspink_frm
optionsbackgroundgridlines_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsbackgroundgridlines_frm", "Gridlines")
options.Check1.Caption = optionsbackgroundgridlines_frm
'-----------------------------------------------------------------
optionsfontcolorlabel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsfontcolorlabel_frm", "Font color:")
options.Label2 = optionsfontcolorlabel_frm
optionsfonttypelabel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionsfonttypelabel_frm", "Font type:")
options.Label3 = optionsfonttypelabel_frm
'-----------------------------------------------------------------
optionslanguagelabel_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionslanguagelabel_frm", "Interface language:")
options.Label4 = optionslanguagelabel_frm
optionslanguagereset_frm = Leer_Ini(App.Path & "\Languages\" & principal.Text1 & ".lng", "optionslanguagereset_frm", "*Language reset requires restart (click)")
options.Label5 = optionslanguagereset_frm

Option1(principal.Text2.Text).Value = True
Option2(principal.Text3.Text).Value = True

Check1.Value = principal.Text4

Combo1.Text = principal.ListView1.Font

If principal.Text1.Text = "English" Then
List1.ListIndex = 0
ElseIf principal.Text1.Text = "Español" Then
List1.ListIndex = 1
ElseIf principal.Text1.Text = "Deutsch" Then
List1.ListIndex = 2
ElseIf principal.Text1.Text = "Français" Then
List1.ListIndex = 3
ElseIf principal.Text1.Text = "Italiano" Then
List1.ListIndex = 4
ElseIf principal.Text1.Text = "Português" Then
List1.ListIndex = 5
End If
End Sub

Private Sub List1_Click()
principal.Text1.Text = List1.List(List1.ListIndex)
End Sub

Private Sub Option1_Click(Index As Integer)
principal.ListView1.BackColor = Option1(Index).BackColor
principal.Text2.Text = Option1(Index).Index
End Sub

Private Sub Option2_Click(Index As Integer)
principal.ListView1.ForeColor = Option2(Index).ForeColor
principal.Text3.Text = Option2(Index).Index
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(1).Selected = True Then
frmBackground.Visible = True
frmFont.Visible = False
frmLanguage.Visible = False
ElseIf TabStrip1.Tabs(2).Selected = True Then
frmBackground.Visible = False
frmFont.Visible = True
frmLanguage.Visible = False
ElseIf TabStrip1.Tabs(3).Selected = True Then
frmBackground.Visible = False
frmFont.Visible = False
frmLanguage.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
If Check1.Value = 1 Then
principal.ListView1.gridlines = True
principal.Text4.Text = 1
Else
principal.ListView1.gridlines = False
principal.Text4.Text = 0
End If
End Sub

Private Sub Timer2_Timer()
principal.ListView1.Font = Combo1.Text
End Sub
