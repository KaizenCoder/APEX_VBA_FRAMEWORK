' Migrated to apex-core/utils - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Module : modTextUtils
' Version : 1.0
' Purpose : Utilitaires pour la manipulation des chaînes de texte
' Date : 10/04/2025
' ==========================================================================

' --- Variables privées ---
Private m_Logger As Object ' ILoggerBase
Private m_LastError As String

' --- Initialisation ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize(Optional ByVal logger As Object = Nothing)
    ' Initialise le module avec un logger optionnel
    Set m_Logger = logger
    m_LastError = ""
    
    If Not m_Logger Is Nothing Then
        ' TODO: Logger l'initialisation
        ' m_Logger.LogInfo "modTextUtils initialisé", "TEXT"
    End If
End Sub

' --- Fonctions de base sur les chaînes ---
'@Description: 
'@Param: 
'@Returns: 

Public Function IsNullOrEmpty(ByVal text As Variant) As Boolean
    ' Vérifie si une chaîne est nulle ou vide
    On Error GoTo ErrorHandler
    
    ' Vérifier si la variable est de type chaîne
    If VarType(text) <> vbString Then
        ' Si ce n'est pas une chaîne, vérifier si c'est Null ou Empty
        IsNullOrEmpty = IsNull(text) Or IsEmpty(text)
    Else
        ' Si c'est une chaîne, vérifier si elle est vide
        IsNullOrEmpty = (Len(text) = 0)
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si la chaîne est nulle ou vide: " & Err.Description
    LogError m_LastError
    IsNullOrEmpty = True
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function IsNullOrWhiteSpace(ByVal text As Variant) As Boolean
    ' Vérifie si une chaîne est nulle, vide ou ne contient que des espaces
    On Error GoTo ErrorHandler
    
    ' Si la chaîne est nulle ou vide, retourner vrai
    If IsNullOrEmpty(text) Then
        IsNullOrWhiteSpace = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si la chaîne ne contient que des espaces
    IsNullOrWhiteSpace = (Len(Trim(text)) = 0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si la chaîne est nulle, vide ou ne contient que des espaces: " & Err.Description
    LogError m_LastError
    IsNullOrWhiteSpace = True
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function IfNullOrEmpty(ByVal value As Variant, ByVal defaultValue As Variant) As Variant
    ' Retourne la valeur ou une valeur par défaut si elle est nulle ou vide
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(value) Then
        IfNullOrEmpty = defaultValue
    Else
        IfNullOrEmpty = value
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du remplacement d'une valeur nulle ou vide: " & Err.Description
    LogError m_LastError
    IfNullOrEmpty = defaultValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function SafeLeft(ByVal text As String, ByVal length As Long) As String
    ' Version sécurisée de la fonction Left
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        SafeLeft = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la longueur demandée est négative, retourner une chaîne vide
    If length <= 0 Then
        SafeLeft = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la longueur demandée est supérieure à la longueur de la chaîne, retourner la chaîne entière
    If length >= Len(text) Then
        SafeLeft = text
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Sinon, utiliser la fonction Left standard
    SafeLeft = Left(text, length)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction de caractères à gauche: " & Err.Description
    LogError m_LastError
    SafeLeft = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function SafeRight(ByVal text As String, ByVal length As Long) As String
    ' Version sécurisée de la fonction Right
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        SafeRight = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la longueur demandée est négative, retourner une chaîne vide
    If length <= 0 Then
        SafeRight = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la longueur demandée est supérieure à la longueur de la chaîne, retourner la chaîne entière
    If length >= Len(text) Then
        SafeRight = text
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Sinon, utiliser la fonction Right standard
    SafeRight = Right(text, length)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction de caractères à droite: " & Err.Description
    LogError m_LastError
    SafeRight = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function SafeMid(ByVal text As String, ByVal start As Long, Optional ByVal length As Long = -1) As String
    ' Version sécurisée de la fonction Mid
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        SafeMid = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la position de départ est négative ou nulle, la fixer à 1
    If start <= 0 Then
        start = 1
    End If
    
    ' Si la position de départ est au-delà de la longueur de la chaîne, retourner une chaîne vide
    If start > Len(text) Then
        SafeMid = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la longueur n'est pas spécifiée ou est négative, extraire jusqu'à la fin de la chaîne
    If length < 0 Then
        SafeMid = Mid(text, start)
    Else
        ' Sinon, utiliser la fonction Mid standard
        SafeMid = Mid(text, start, length)
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction de caractères au milieu: " & Err.Description
    LogError m_LastError
    SafeMid = ""
End Function

' --- Fonctions de manipulation de chaînes ---
'@Description: 
'@Param: 
'@Returns: 

Public Function RemoveAccents(ByVal text As String) As String
    ' Supprime les accents d'une chaîne
    Dim i As Long
    Dim char As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        RemoveAccents = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    result = ""
    
    ' Parcourir chaque caractère de la chaîne
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        ' Remplacer les caractères accentués
        Select Case AscW(char)
            ' A accentué
            Case 192 To 197
                result = result & "A"
            ' a accentué
            Case 224 To 229
                result = result & "a"
            ' E accentué
            Case 200 To 203
                result = result & "E"
            ' e accentué
            Case 232 To 235
                result = result & "e"
            ' I accentué
            Case 204 To 207
                result = result & "I"
            ' i accentué
            Case 236 To 239
                result = result & "i"
            ' O accentué
            Case 210 To 214
                result = result & "O"
            ' o accentué
            Case 242 To 246
                result = result & "o"
            ' U accentué
            Case 217 To 220
                result = result & "U"
            ' u accentué
            Case 249 To 252
                result = result & "u"
            ' C cédille
            Case 199
                result = result & "C"
            ' c cédille
            Case 231
                result = result & "c"
            ' N tilde
            Case 209
                result = result & "N"
            ' n tilde
            Case 241
                result = result & "n"
            ' Autres caractères spéciaux
            Case 198
                result = result & "AE"
            Case 230
                result = result & "ae"
            Case 338
                result = result & "OE"
            Case 339
                result = result & "oe"
            ' Caractère normal
            Case Else
                result = result & char
        End Select
    Next i
    
    RemoveAccents = result
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression des accents: " & Err.Description
    LogError m_LastError
    RemoveAccents = text
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function Contains(ByVal text As String, ByVal subString As String, _
                         Optional ByVal caseSensitive As Boolean = True) As Boolean
    ' Vérifie si une chaîne contient une sous-chaîne
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Or IsNullOrEmpty(subString) Then
        Contains = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si la chaîne contient la sous-chaîne
    If caseSensitive Then
        Contains = (InStr(1, text, subString, vbBinaryCompare) > 0)
    Else
        Contains = (InStr(1, text, subString, vbTextCompare) > 0)
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si la chaîne contient une sous-chaîne: " & Err.Description
    LogError m_LastError
    Contains = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function StartsWith(ByVal text As String, ByVal prefix As String, _
                           Optional ByVal caseSensitive As Boolean = True) As Boolean
    ' Vérifie si une chaîne commence par un préfixe
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        StartsWith = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    If IsNullOrEmpty(prefix) Then
        StartsWith = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si le préfixe est plus long que la chaîne, retourner faux
    If Len(prefix) > Len(text) Then
        StartsWith = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si la chaîne commence par le préfixe
    If caseSensitive Then
        StartsWith = (Left(text, Len(prefix)) = prefix)
    Else
        StartsWith = (StrComp(Left(text, Len(prefix)), prefix, vbTextCompare) = 0)
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si la chaîne commence par un préfixe: " & Err.Description
    LogError m_LastError
    StartsWith = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function EndsWith(ByVal text As String, ByVal suffix As String, _
                         Optional ByVal caseSensitive As Boolean = True) As Boolean
    ' Vérifie si une chaîne se termine par un suffixe
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        EndsWith = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    If IsNullOrEmpty(suffix) Then
        EndsWith = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si le suffixe est plus long que la chaîne, retourner faux
    If Len(suffix) > Len(text) Then
        EndsWith = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si la chaîne se termine par le suffixe
    If caseSensitive Then
        EndsWith = (Right(text, Len(suffix)) = suffix)
    Else
        EndsWith = (StrComp(Right(text, Len(suffix)), suffix, vbTextCompare) = 0)
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si la chaîne se termine par un suffixe: " & Err.Description
    LogError m_LastError
    EndsWith = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function Truncate(ByVal text As String, ByVal maxLength As Long, _
                         Optional ByVal ellipsis As String = "...") As String
    ' Tronque une chaîne à une longueur maximale avec ellipsis
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        Truncate = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la chaîne est déjà plus courte que la longueur maximale, la retourner telle quelle
    If Len(text) <= maxLength Then
        Truncate = text
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la longueur maximale est trop petite pour l'ellipsis, tronquer sans ellipsis
    If maxLength <= Len(ellipsis) Then
        Truncate = Left(text, maxLength)
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Tronquer la chaîne avec ellipsis
    Truncate = Left(text, maxLength - Len(ellipsis)) & ellipsis
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la troncature de la chaîne: " & Err.Description
    LogError m_LastError
    Truncate = text
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function PadLeft(ByVal text As String, ByVal totalWidth As Long, _
                        Optional ByVal paddingChar As String = " ") As String
    ' Complète une chaîne à gauche avec un caractère de remplissage
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        PadLeft = String(totalWidth, paddingChar)
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la chaîne est déjà plus longue que la largeur totale, la retourner telle quelle
    If Len(text) >= totalWidth Then
        PadLeft = text
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Compléter la chaîne à gauche
    PadLeft = String(totalWidth - Len(text), paddingChar) & text
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du remplissage à gauche de la chaîne: " & Err.Description
    LogError m_LastError
    PadLeft = text
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function PadRight(ByVal text As String, ByVal totalWidth As Long, _
                         Optional ByVal paddingChar As String = " ") As String
    ' Complète une chaîne à droite avec un caractère de remplissage
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        PadRight = String(totalWidth, paddingChar)
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Si la chaîne est déjà plus longue que la largeur totale, la retourner telle quelle
    If Len(text) >= totalWidth Then
        PadRight = text
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Compléter la chaîne à droite
    PadRight = text & String(totalWidth - Len(text), paddingChar)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du remplissage à droite de la chaîne: " & Err.Description
    LogError m_LastError
    PadRight = text
End Function

' --- Fonctions utilitaires avancées ---
'@Description: 
'@Param: 
'@Returns: 

Public Function ExtractNumbers(ByVal text As String) As String
    ' Extrait tous les chiffres d'une chaîne
    Dim i As Long
    Dim char As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        ExtractNumbers = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    result = ""
    
    ' Parcourir chaque caractère de la chaîne
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        ' Ne conserver que les chiffres
        If char >= "0" And char <= "9" Then
            result = result & char
        End If
    Next i
    
    ExtractNumbers = result
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction des chiffres: " & Err.Description
    LogError m_LastError
    ExtractNumbers = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function ExtractLetters(ByVal text As String) As String
    ' Extrait toutes les lettres d'une chaîne
    Dim i As Long
    Dim char As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        ExtractLetters = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    result = ""
    
    ' Parcourir chaque caractère de la chaîne
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        ' Ne conserver que les lettres
        If (char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Then
            result = result & char
        End If
    Next i
    
    ExtractLetters = result
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'extraction des lettres: " & Err.Description
    LogError m_LastError
    ExtractLetters = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function IsNumeric2(ByVal text As String, Optional ByVal decimalSeparator As String = ",") As Boolean
    ' Version améliorée de la fonction IsNumeric
    Dim i As Long
    Dim char As String
    Dim hasSeparator As Boolean
    
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        IsNumeric2 = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si la chaîne est vide après avoir supprimé les espaces
    If Len(Trim(text)) = 0 Then
        IsNumeric2 = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Gérer le signe +/-
    If Left(text, 1) = "+" Or Left(text, 1) = "-" Then
        text = Mid(text, 2)
    End If
    
    hasSeparator = False
    
    ' Parcourir chaque caractère de la chaîne
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        ' Vérifier si le caractère est un chiffre
        If char >= "0" And char <= "9" Then
            ' OK
        ' Vérifier si le caractère est un séparateur décimal
        ElseIf char = decimalSeparator Then
            ' Si on a déjà rencontré un séparateur, la chaîne n'est pas un nombre
            If hasSeparator Then
                IsNumeric2 = False
                Exit'@Description: 
'@Param: 
'@Returns: 

 Function
            End If
            
            hasSeparator = True
        ' Si ce n'est ni un chiffre ni un séparateur, la chaîne n'est pas un nombre
        Else
            IsNumeric2 = False
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
    
    IsNumeric2 = True
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification si la chaîne est un nombre: " & Err.Description
    LogError m_LastError
    IsNumeric2 = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function CountOccurrences(ByVal text As String, ByVal subString As String, _
                                Optional ByVal caseSensitive As Boolean = True) As Long
    ' Compte le nombre d'occurrences d'une sous-chaîne dans une chaîne
    Dim compareMethod As VbCompareMethod
    Dim pos As Long
    Dim count As Long
    
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Or IsNullOrEmpty(subString) Then
        CountOccurrences = 0
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Définir la méthode de comparaison
    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    
    count = 0
    pos = 1
    
    ' Compter les occurrences
    Do
        pos = InStr(pos, text, subString, compareMethod)
        
        If pos = 0 Then
            Exit Do
        End If
        
        count = count + 1
        pos = pos + Len(subString)
    Loop While pos <= Len(text)
    
    CountOccurrences = count
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du comptage des occurrences: " & Err.Description
    LogError m_LastError
    CountOccurrences = 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function SplitToArray(ByVal text As String, ByVal delimiter As String, _
                            Optional ByVal caseSensitive As Boolean = True) As Variant
    ' Split une chaîne en tableau en respectant la casse si demandé
    Dim compareMethod As VbCompareMethod
    
    On Error GoTo ErrorHandler
    
    If IsNullOrEmpty(text) Then
        SplitToArray = Array()
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Définir la méthode de comparaison
    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    
    ' Split en utilisant la méthode de comparaison
    If caseSensitive Then
        SplitToArray = Split(text, delimiter)
    Else
        ' Si la comparaison n'est pas sensible à la casse, implémenter un split personnalisé
        Dim result() As String
        Dim tmpText As String
        Dim pos As Long
        Dim count As Long
        
        tmpText = text
        count = 0
        ReDim result(0 To 100)
        
        ' Trouver les positions des délimiteurs
        Do
            pos = InStr(1, tmpText, delimiter, vbTextCompare)
            
            If pos = 0 Then
                ' Ajouter la dernière partie
                If Len(tmpText) > 0 Then
                    If count > UBound(result) Then
                        ReDim Preserve result(0 To count * 2)
                    End If
                    
                    result(count) = tmpText
                    count = count + 1
                End If
                
                Exit Do
            End If
            
            ' Ajouter la partie avant le délimiteur
            If count > UBound(result) Then
                ReDim Preserve result(0 To count * 2)
            End If
            
            result(count) = Left(tmpText, pos - 1)
            count = count + 1
            
            ' Supprimer la partie traitée
            tmpText = Mid(tmpText, pos + Len(delimiter))
        Loop
        
        ' Redimensionner le tableau au nombre exact d'éléments
        If count > 0 Then
            ReDim Preserve result(0 To count - 1)
            SplitToArray = result
        Else
            SplitToArray = Array()
        End If
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du split de la chaîne: " & Err.Description
    LogError m_LastError
    SplitToArray = Array()
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function JoinArray(ByVal arr As Variant, ByVal delimiter As String) As String
    ' Join un tableau en chaîne
    On Error GoTo ErrorHandler
    
    ' Vérifier si le tableau est vide
    If Not IsArray(arr) Then
        JoinArray = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    JoinArray = Join(arr, delimiter)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du join du tableau: " & Err.Description
    LogError m_LastError
    JoinArray = ""
End Function

' --- Propriétés ---
Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

' --- Fonctions privées ---
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "TEXT"
    End If
End Sub
