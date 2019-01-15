Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object

Sub Main()
    Dim doc As ModelDoc2
    Dim part As PartDoc
    Dim material As String
    Dim matdb As String
    
    Set swApp = Application.SldWorks
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Exit Sub
    If doc.GetType <> swDocPART Then
        MsgBox "Only parts supported.", vbCritical
        Exit Sub
    End If
    Set part = doc
    material = part.GetMaterialPropertyName2(doc.ConfigurationManager.ActiveConfiguration.Name, matdb)
    part.RemoveAllDisplayStates 'material erased after this
    part.SetMaterialPropertyName2 "", matdb, material
    doc.ForceRebuild3 True
End Sub
