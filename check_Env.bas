Public Function check_Env(Env As String) As Boolean
'Check si l'environnement est conforme aux prérequis de lancement des macros
'Env = "Part", "Product"
On Error Resume Next
Dim mPart As PartDocument
Dim mprod As ProductDocument

Select Case Env
    Case "Parts"
        Set mPart = CATIA.ActiveDocument
        If err.Number <> 0 Then
            MsgBox "Activez un CATPart avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            err.Clear
            End
        Else
            check_Env = True
        End If
    Case "Product" 'Test si un CatProduct est actif
        Set mprod = CATIA.ActiveDocument
        If err.Number <> 0 Then
            MsgBox "Activez un Catproduct avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            err.Clear
            End
        Else
            check_Env = True
        End If
    End Select
    
    On Error GoTo 0
End Function