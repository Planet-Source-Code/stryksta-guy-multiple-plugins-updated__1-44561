VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Once again, by using a generic entry function, we can do just about anything with our
'plug-in. This one opens up a user interface


Public Function Run(FormX As Object) As String
On Error GoTo errtrap
    Form1.Show
    Run = vbNullString
    Set Form1.objHost = FormX
Exit Function
errtrap:
    Run = Err.Description
End Function
Public Function Identify() As String

  Identify = "Show Form"

End Function
Public Function Description() As String

  Description = "This is an example of displaying a description for the plugins and so on and so on, now you have an example to show name the plugin rather than just the filename"

End Function
Public Function Version() As String

  Version = "1.02"

End Function
