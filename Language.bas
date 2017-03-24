Public Sub InitLanguage()
'Configure les champs en fonction de la langue
    Langue = Language
    If Langue = "EN" Then
        nQt = "Quantity"
        nRef = "Part Number"
        nRev = "Revision"
        nDef = "Definition"
        nNom = "Nomenclature"
        nDesc = "Product Description"
        nSrce = "Source"
        nParamActivate = "Component Activation State"
    Else
        nQt = "Quantité"
        nRef = "Référence"
        nRev = "Révision"
        nDef = "Définition"
        nNom = "Nomenclature"
        nDesc = "Description du produit"
        nSrce = "Source"
        nParamActivate = "Etat d'activation du composant"
    End If
End Sub

Public Function Language() As String
'Détecte la langue de l'interface Catia
'Ouvre un part vierge et test le nom du "Main Body"
Dim oFolder, ofs
Dim EmptyPartFolder, EmptyPartFile
Dim oEmptyPart  As PartDocument

On Error Resume Next
Set ofs = CreateObject("Scripting.FileSystemObject")
Set oFolder = ofs.GetFolder(CATIA.Parent.Path)
Set EmptyPartFolder = ofs.GetFolder(oFolder.ParentFolder.ParentFolder.Path & "\startup\templates") ' dossier relatif des modèles vides
Set EmptyPartFile = ofs.GetFile(EmptyPartFolder.Path & "\empty.CATPart")

If Err.Number = 0 Then
    On Error GoTo 0
    Set oEmptyPart = CATIA.Documents.Open(EmptyPartFile.Path)
    If oEmptyPart.Part.MainBody.Name = "PartBody" Then
        Language = "EN"
    Else
        Language = "FR"
    End If
End If

    oEmptyPart.Close
 Set oEmptyPart = Nothing
 Set EmptyPartFile = Nothing
 Set EmptyPartFolder = Nothing
 Set oFolder = Nothing
 Set ofs = Nothing
 On Error GoTo 0

End Function