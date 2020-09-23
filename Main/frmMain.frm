VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Main Program"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Text            =   "Version"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Label3 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3360
      Width           =   4335
   End
   Begin VB.CommandButton Browse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMain.frx":0000
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Title"
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   1980
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000002&
      Height          =   1980
      Left            =   120
      Pattern         =   "*.dll"
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Plugin Directory: "
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Available Plugins:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'use a variable to store the selected plugin name
Dim PluginName
'use a variable to assign the message box message to
Dim pMsg

Private Sub Browse_Click()
Dim TmpPath As String
Dim path As String
TmpPath = Label3
If Len(TmpPath) > 0 Then
    If Not Right$(TmpPath, 1) <> "\" Then TmpPath = Left$(TmpPath, Len(TmpPath) - 1) ' Remove "\" if the user added
End If
TmpPath = BrowseForFolder(TmpPath) ' Browse for folder
If TmpPath = "" Then
    path = "No folder selected !" ' If the user pressed cancel
Else
    path = TmpPath & "\" ' If the user selected a folder
End If
Label3 = path
List1.Clear
File1.path = path
File1.Pattern = "*.dll;*.exe"
RemoveSelf
End Sub

Private Sub Command1_Click()

On Error GoTo errhandler

    'use a variable to define the plugin
    Dim objPlugIn As Object
    'Variable contains plugin's response
    Dim Response As String
    Dim Indentity As String
    'Load the plugin
    'The format for CreateObject is [Project name].[Class module name]
    'Note: do not specify a filename, only the name of the project or class module
    Set objPlugIn = CreateObject(PluginName(0) + ".plugMain")
    'Call the entry function
    Response = objPlugIn.Run(Me)
    Indentity = objPlugIn.Identify
    'if the plugin contains an error, show us in a message box
    If Response <> vbNullString Then
        MsgBox Response
    End If
    Text1.Text = Indentity
Exit Sub

errhandler:
    Select Case Err.Number
        Case 438 'Couldn't find the entry function
            MsgBox "There was an error running the selected plugin." + vbCrLf + "Please contact your program vendor if the problem continues.", vbExclamation
            
        Case 13 'Didnt select a plugin from the list
            MsgBox "You did not select a plugin from the list."
            
    End Select
End Sub

Private Sub Command2_Click()

End Sub
Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim TmpPath As String * 256
With Browse_for_folder
    .hOwner = Me.hWnd ' Window Handle
    .lpszTitle = "Browse for folders" ' Dialog Title
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, TmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(TmpPath, InStr(TmpPath, vbNullChar) - 1) ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function
Private Sub Form_Load()
File1.path = App.path
File1.Pattern = "*.dll;*.exe"
RemoveSelf
Label3.Text = App.path
End Sub

Sub RemoveSelf()
Dim i As Integer
' use a variable for the loop we are going to do
For i = 0 To File1.ListCount - 1
' loop while variable i is a number not more than the number of plugins
    File1.ListIndex = i
    ' select the i option in the list
    If Not LCase(File1.FileName) = LCase(App.EXEName + ".exe") Then
    ' both compared in lower case
    ' if the filename selected in the file list is not the filename of
    ' the main program then
        List1.AddItem File1.FileName, 0
        ' add the filename to the plugins list
    End If
Next i
' push back to the top of the loop
End Sub

Private Sub List1_Click()
On Error GoTo errhandler
    PluginName = Split(List1.List(List1.ListIndex), ".")
    
    Dim objPlugIn As Object
    Dim Response As String
    Dim Indentity As String
    Dim Description As String
    Dim Version As String
    
    Set objPlugIn = CreateObject(PluginName(0) + ".plugMain")
    
    Indentity = objPlugIn.Identify
    Description = objPlugIn.Description
    Version = objPlugIn.Version
    
    If Response <> vbNullString Then
        MsgBox Response
    End If
    
    Text1.Text = Indentity
    Text2.Text = Description
    Text3.Text = Version
    
errhandler:
    Select Case Err.Number
        Case 438 'Couldn't find the entry function
            MsgBox "There was an error running the selected plugin." + vbCrLf + "Please contact your program vendor if the problem continues.", vbExclamation
            
        Case 429 'Not A Plugin"
            MsgBox "This plugin is not compatible with this software", , "Not Compatible"
            
    End Select
    
Exit Sub
End Sub
