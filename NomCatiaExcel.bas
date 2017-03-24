Public Function NoDebLstPieces(objWBk) As Long
'recherche la 1ere ligne des attributs des pièces
' la ligne commence par "Liste des pièces"
    Dim NomSeparateur As String
    Dim NoLigne As Long
    NoLigne = 1
    NomSeparateur = "Liste des pièces"
    While Left(objWBk.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebLstPieces = NoLigne

End Function

Public Function NoDebRecap(objWBk As Variant) As Long
'recherche la 1ere ligne du récapitulatif des pièces
' la ligne commence par "Nomenclature de" ou "Recapitulation of:"
    Dim NomSeparateur As String
    Dim NoLigne As Integer
    NoLigne = 1
    If Langue = "EN" Then
        NomSeparateur = "Recapitulation of:"
    ElseIf Langue = "FR" Then
        NomSeparateur = "Récapitulatif sur"
    End If
    While Left(objWBk.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebRecap = NoLigne
End Function

Public Function NoDerniereLigne(objWBk As Variant) As Long
'recherche la dernière ligne du fichier excel
'On part du principe que 2 lignes vide indiquent la fin du fichier
Dim NoLigne As Integer, NbLigVide As Integer
    NoLigne = 1
    NbLigVide = 0
    While NbLigVide < 2
        If objWBk.ActiveSheet.cells(NoLigne, 1).Value = "" Then
            NbLigVide = NbLigVide + 1
        Else
            NbLigVide = 0
        End If
    NoLigne = NoLigne + 1
    Wend
    NoDerniereLigne = NoLigne - 2
End Function

Public Function TestEstSSE(Ligne As String) As String
'test si la ligne correspond a une entète de sous ensemble
' la ligne commence par "Nomenclature de" ou "Bill of Material"
Dim NomSeparateur As String
Dim tmpNomSSE As String

    If Langue = "EN" Then
        NomSeparateur = "Bill of Material: "
    ElseIf Langue = "FR" Then
        NomSeparateur = "Nomenclature de "
    End If
    On Error Resume Next 'Test si la chaine est vide ou inférieur a len(nomséparateur)
    tmpNomSSE = Right(Ligne, Len(Ligne) - Len(NomSeparateur))
    If Err.Number <> 0 Then
         TestEstSSE = "False"
    Else
        If Left(Ligne, Len(NomSeparateur)) = NomSeparateur Then
            TestEstSSE = tmpNomSSE
        Else
            TestEstSSE = "False"
        End If
    End If
End Function

Public Function FormatSource(str As String) As String
'formate le contenu du champs source
'remplace "Inconu" ou "Unknown" par une chaine vide.
'remplace les codes champs sources par une string
FormatSource = str
Select Case str
    Case "Inconnu", "Unknown"
        FormatSource = ""
    Case "Bought", catProductBought
        FormatSource = "Acheté"
    Case "Made", catProductMade
        FormatSource = "Fabriqué"
End Select
End Function