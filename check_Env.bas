Public Function check_Env(Env As String) As Boolean
'Check si l'environnement est conforme aux prérequis de lancement des macros
'Env = "Part", "Product", "CatDrawing"
On Error Resume Next
Dim mPart As PartDocument
Dim mprod As ProductDocument
Dim mDraw As DrawingDocument
check_Env = False

Select Case UCase(Env)
    Case "PART" 'Test si un CatPart est actif
        Set mPart = CATIA.ActiveDocument
        If Err.Number <> 0 Then
            MsgBox "Activez un CATPart avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            Err.Clear
            End
        Else
            check_Env = True
        End If
    Case "PRODUCT" 'Test si un CatProduct est actif
        Set mprod = CATIA.ActiveDocument
        If Err.Number <> 0 Then
            MsgBox "Activez un Catproduct avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            Err.Clear
            End
        Else
            check_Env = True
        End If
    Case "DRAWING" 'Test si un CatDrawing est actif
        Set mDraw = CATIA.ActiveDocument
        If Err.Number <> 0 Then
            MsgBox "Activez un CatDrawing avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            Err.Clear
            End
        Else
            check_Env = True
        End If
    End Select
    
    On Error GoTo 0
End Function