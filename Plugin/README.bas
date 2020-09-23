Attribute VB_Name = "README"
' note that this project is an ActiveX DLL project
' you can't use forms in these because they compile
' as DLLs

' make sure you have at least one class module in each
' plugin called 'plugMain'. NOTE: the filename does not
' have to be 'plugMain' as well
' NOTE: NEVER CALL ANYTHING 'MAIN'

' make sure that there is a Public Function Run() in
' each plugins 'plugMain' class module
' the name of the function can be changed as long as
' you change the code in the main form

' when you compile, make sure the filename is the
' same as the project name
