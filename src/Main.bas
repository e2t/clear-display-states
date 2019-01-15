Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Dim currentDoc As ModelDoc2

Sub Main()
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If Not currentDoc Is Nothing Then
        If currentDoc.GetType <> swDocPART Then
            MsgBox "Only parts supported.", vbCritical
        Else
            MainForm.Show
        End If
    End If
End Sub

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function

Sub Run(isOneColor As Boolean)
    Dim part As PartDoc
    Dim i As Variant
    Dim materials As Collection 'of MaterialValue
    Dim confs(0) As String
    Dim currentConf As String
    
    Set part = currentDoc
    Set materials = New Collection
    
    Dim vs As MaterialVisualPropertiesData
    Set vs = part.GetMaterialVisualProperties
    
    currentConf = currentDoc.ConfigurationManager.ActiveConfiguration.Name
    
    If isOneColor Then
        RememberMaterial currentConf, currentConf, materials, part
    Else
        For Each i In currentDoc.GetConfigurationNames
            RememberMaterial i, i, materials, part
        Next
    End If
    
    part.RemoveAllDisplayStates  ' Material erased after this
    currentDoc.Extension.LinkedDisplayState = Not isOneColor
    For Each i In materials
        currentDoc.ShowConfiguration2 i.toConf
        part.SetMaterialPropertyName2 i.toConf, i.DataBase, i.Material
        confs(0) = i.toConf
        part.SetMaterialVisualProperties vs, swSpecifyConfiguration, confs
    Next
    currentDoc.ShowConfiguration2 currentConf
End Sub

Sub RememberMaterial(ByVal fromConf As String, ByVal toConf As String, _
                     ByRef materials As Collection, part As PartDoc)
    Dim value As MaterialValue
    Dim db As String
    
    Set value = New MaterialValue
    value.Material = part.GetMaterialPropertyName2(fromConf, db)
    value.DataBase = db
    value.toConf = toConf
    materials.Add value
End Sub
