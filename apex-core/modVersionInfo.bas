' Migrated to apex-core - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Attribute VB_Name = "modVersionInfo"
Option Explicit
' ==========================================================================
' Module : modVersionInfo
' Version : 1.0
' Purpose : Gestion des informations de version du framework APEX
' Date    : 10/04/2025
' ==========================================================================

' --- Constantes ---
Private Const MODULE_NAME As String = "modVersionInfo"

Public Const FRAMEWORK_NAME As String = "APEX VBA Framework"
Public Const FRAMEWORK_VERSION As String = "1.0.0"
Public Const FRAMEWORK_BUILD As String = "20250410"
Public Const FRAMEWORK_COPYRIGHT As String = "Copyright © 2025 APEX Team"
Public Const FRAMEWORK_RELEASE_DATE As String = "10/04/2025"
Public Const FRAMEWORK_COMPATIBILITY As String = "Excel 2016+"

Private Const VERSION_FILE As String = "VERSION.txt"
Private Const VERSION_REGEX_PATTERN As String = "^(\d+)\.(\d+)\.(\d+)(?:\-(alpha|beta|rc)\.(\d+))?$"

' --- Propriétés publiques ---
Public Function GetFrameworkVersion() As String
    GetFrameworkVersion = FRAMEWORK_VERSION
End Function

Public Function GetFrameworkVersionFull() As String
    GetFrameworkVersionFull = FRAMEWORK_NAME & " v" & FRAMEWORK_VERSION & " (Build " & FRAMEWORK_BUILD & ")"
End Function

Public Function GetFrameworkVersionInfo() As String
    GetFrameworkVersionInfo = "Nom: " & FRAMEWORK_NAME & vbCrLf & _
                              "Version: " & FRAMEWORK_VERSION & vbCrLf & _
                              "Build: " & FRAMEWORK_BUILD & vbCrLf & _
                              "Date de sortie: " & FRAMEWORK_RELEASE_DATE & vbCrLf & _
                              "Compatibilité: " & FRAMEWORK_COMPATIBILITY & vbCrLf & _
                              "Copyright: " & FRAMEWORK_COPYRIGHT
End Function

Public Function IsPreRelease() As Boolean
    ' Vérifier si la version actuelle est une préversion (alpha, beta, rc)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = VERSION_REGEX_PATTERN
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(FRAMEWORK_VERSION)
    
    ' Si correspondance et capture du groupe de préversion
    If matches.Count > 0 Then
        Dim match As Object
        Set match = matches(0)
        
        ' Le groupe 4 contient alpha/beta/rc si présent
        If match.SubMatches.Count >= 4 Then
            IsPreRelease = (match.SubMatches(3) <> "")
        Else
            IsPreRelease = False
        End If
    Else
        IsPreRelease = False
    End If
End Function

Public Function CompareVersions(ByVal version1 As String, ByVal version2 As String) As Long
    ' Compare deux versions et retourne:
    '  1 si version1 > version2
    '  0 si version1 = version2
    ' -1 si version1 < version2
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = VERSION_REGEX_PATTERN
    End With
    
    ' Vérifier si les deux versions sont valides
    If Not regex.Test(version1) Or Not regex.Test(version2) Then
        ' Version non valide
        CompareVersions = 0
        Exit Function
    End If
    
    ' Extraire les composants de version1
    Dim matches1 As Object
    Set matches1 = regex.Execute(version1)
    Dim major1 As Long, minor1 As Long, patch1 As Long
    Dim preType1 As String, preNum1 As Long
    
    With matches1(0)
        major1 = CLng(.SubMatches(0))
        minor1 = CLng(.SubMatches(1))
        patch1 = CLng(.SubMatches(2))
        
        If .SubMatches.Count >= 5 And .SubMatches(3) <> "" Then
            preType1 = .SubMatches(3)
            preNum1 = CLng(.SubMatches(4))
        Else
            preType1 = ""
            preNum1 = 0
        End If
    End With
    
    ' Extraire les composants de version2
    Dim matches2 As Object
    Set matches2 = regex.Execute(version2)
    Dim major2 As Long, minor2 As Long, patch2 As Long
    Dim preType2 As String, preNum2 As Long
    
    With matches2(0)
        major2 = CLng(.SubMatches(0))
        minor2 = CLng(.SubMatches(1))
        patch2 = CLng(.SubMatches(2))
        
        If .SubMatches.Count >= 5 And .SubMatches(3) <> "" Then
            preType2 = .SubMatches(3)
            preNum2 = CLng(.SubMatches(4))
        Else
            preType2 = ""
            preNum2 = 0
        End If
    End With
    
    ' Comparer les versions
    If major1 <> major2 Then
        CompareVersions = IIf(major1 > major2, 1, -1)
    ElseIf minor1 <> minor2 Then
        CompareVersions = IIf(minor1 > minor2, 1, -1)
    ElseIf patch1 <> patch2 Then
        CompareVersions = IIf(patch1 > patch2, 1, -1)
    ElseIf preType1 <> preType2 Then
        ' Une version sans prérelease est supérieure à une version avec prérelease
        If preType1 = "" Then
            CompareVersions = 1
        ElseIf preType2 = "" Then
            CompareVersions = -1
        Else
            ' Ordre: alpha < beta < rc
            Select Case preType1
                Case "alpha"
                    CompareVersions = -1
                Case "beta"
                    CompareVersions = IIf(preType2 = "alpha", 1, -1)
                Case "rc"
                    CompareVersions = 1
            End Select
        End If
    ElseIf preNum1 <> preNum2 Then
        CompareVersions = IIf(preNum1 > preNum2, 1, -1)
    Else
        CompareVersions = 0 ' Versions identiques
    End If
End Function

Public Function GetMajorVersion() As Long
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = VERSION_REGEX_PATTERN
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(FRAMEWORK_VERSION)
    
    If matches.Count > 0 Then
        GetMajorVersion = CLng(matches(0).SubMatches(0))
    Else
        GetMajorVersion = 0
    End If
End Function

Public Function GetMinorVersion() As Long
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = VERSION_REGEX_PATTERN
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(FRAMEWORK_VERSION)
    
    If matches.Count > 0 Then
        GetMinorVersion = CLng(matches(0).SubMatches(1))
    Else
        GetMinorVersion = 0
    End If
End Function

Public Function GetPatchVersion() As Long
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = VERSION_REGEX_PATTERN
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(FRAMEWORK_VERSION)
    
    If matches.Count > 0 Then
        GetPatchVersion = CLng(matches(0).SubMatches(2))
    Else
        GetPatchVersion = 0
    End If
End Function

Public Function ReadVersionFromFile(Optional ByVal filePath As String = "") As String
    On Error Resume Next
    
    ' Utiliser le chemin par défaut si non spécifié
    If filePath = "" Then filePath = VERSION_FILE
    
    ' Vérifier si le fichier existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        ReadVersionFromFile = ""
        Exit Function
    End If
    
    ' Lire le fichier
    Dim ts As Object
    Dim content As String
    Dim lines() As String
    Dim i As Long
    
    Set ts = fso.OpenTextFile(filePath, 1) ' ForReading = 1
    content = ts.ReadAll
    ts.Close
    
    ' Analyser le fichier ligne par ligne
    lines = Split(content, vbCrLf)
    
    For i = 0 To UBound(lines)
        ' Rechercher une ligne commençant par "Version:"
        If Left(Trim(lines(i)), 8) = "Version:" Then
            ReadVersionFromFile = Trim(Mid(lines(i), 9))
            Exit Function
        End If
    Next i
    
    ' Si aucune version trouvée
    ReadVersionFromFile = ""
    On Error GoTo 0
End Function

Public Function GetVersionLabel() As String
    ' Retourne un libellé pour la version actuelle
    If IsPreRelease() Then
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        
        With regex
            .Global = False
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = VERSION_REGEX_PATTERN
        End With
        
        Dim matches As Object
        Set matches = regex.Execute(FRAMEWORK_VERSION)
        
        If matches.Count > 0 Then
            Dim preType As String
            preType = matches(0).SubMatches(3)
            
            Select Case LCase(preType)
                Case "alpha"
                    GetVersionLabel = "Alpha"
                Case "beta"
                    GetVersionLabel = "Bêta"
                Case "rc"
                    GetVersionLabel = "Release Candidate"
                Case Else
                    GetVersionLabel = "Pré-version"
            End Select
        Else
            GetVersionLabel = "Pré-version"
        End If
    Else
        GetVersionLabel = "Version stable"
    End If
End Function

Public Function ShowVersionDialog()
    ' Afficher une boîte de dialogue avec les informations de version
    MsgBox GetFrameworkVersionInfo(), vbInformation, FRAMEWORK_NAME & " - Informations de version"
End Function
