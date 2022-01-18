Attribute VB_Name = "VME_Export"
Option Explicit


' Export function
Private Sub VME_Export(targetBook As Workbook, folderPath As String)
    Dim module As VBComponent           ' vba module
    Dim moduleList As VBComponents      ' vba module list in xlsm file
    Dim extension As String             ' extension of vba module
    Dim filePath As String              ' export folder path
    
    ' Get target modules list
    Set moduleList = targetBook.VBProject.VBComponents
    
    ' All modules export
    For Each module In moduleList
        extension = VME_CheckModuleExtension(module)
        
        If Not extension = "" Then
            ' Target module export!!
            filePath = folderPath & "\" & module.Name & "." & extension
            Call module.export(filePath)
        End If
        
        ' Display output filePath
        Debug.Print filePath
    Next
End Sub

' Check target module extension
Private Function VME_CheckModuleExtension(targetModule As VBComponent) As String
    Dim extension As String

    Select Case targetModule.Type
        Case vbext_ct_ClassModule
            ' Class module
            extension = "cls"
        Case vbext_ct_MSForm
            ' Form module
            ' [!!!Attention!!! ".frx" is also exported together!!]
            extension = "frm"
        Case vbext_ct_StdModule
            ' Standard Module
            extension = "bas"
        Case Else
            ' Other Module
            extension = ""
    End Select
    
    VME_CheckModuleExtension = extension
End Function
