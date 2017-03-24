'Fonction de récupération du username
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'Version de la macro
Public Const VMacro As String = "xxxxxxxxxxx"
Public Const nMacro As String = "xxxxxxxxxxx"
Public Const nPath As String = "\\srvxsiordo\xLogs\01_CatiaMacros"
Public Const nFicLog As String = "logUtilMacro.txt"


Public Sub LogUtilMacro(ByVal mPath As String, ByVal mFic As String, ByVal mMacro As String, ByVal mModule As String, ByVal mVersion As String)
'Log l'utilisation de la macro
'Ecrit une ligne dans un fichier de log sur le serveur
'mPath = localisation du fichier de log ("\\serveur\partage")
'mFic = Nom du fichier de log ("logUtilMacro.txt")
'mMacro = nom de la macro ("NomGSE")
'mVersion = Version de la macro ("version 9.1.4")
'mModule = Nom du module ("_Info_Outillage")

Dim mDate As String
Dim mUser As String
Dim nFicLog As String
Dim LigLog As String
Const ForWriting = 2, ForAppending = 8

    mDate = Date & " " & Time()
    mUser = ReturnUserName()
    nFicLog = mPath & "\" & mFic

    nliglog = mDate & ";" & mUser & ";" & mMacro & ";" & mModule & ";" & mVersion

    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile(nFicLog)
    If err.Number <> 0 Then
        Set f = fs.opentextfile(nFicLog, ForWriting, 1)
    Else
        Set f = fs.opentextfile(nFicLog, ForAppending, 1)
    End If
    
    f.Writeline nliglog
    f.Close
    On Error GoTo 0

End Sub

Function ReturnUserName() As String 'extrait d'un code de Paul, Dave Peterson Exelabo
'Renvoi le user name de l'utilisateur de la station
'fonctionne avec la fonction GetUserName dans l'entète de déclaration
    Dim Buffer As String * 256
    Dim BuffLen As Long
    BuffLen = 256
    If GetUserName(Buffer, BuffLen) Then _
    ReturnUserName = Left(Buffer, BuffLen - 1)
End Function