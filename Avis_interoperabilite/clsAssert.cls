VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsAssert.cls
' Classe d'assertions pour le framework de test
'
' Cette classe fournit des méthodes pour valider différentes conditions
' dans les tests unitaires, facilitant la vérification des résultats attendus.
'
' @class clsAssert
' @author APEX Framework Team
' @version 1.0
' @date 2024-07-27
Option Explicit

' États internes
Private Type TAssertState
    FailureCount As Long             ' Nombre d'échecs d'assertions
    LastFailureMessage As String     ' Dernier message d'échec
    Initialized As Boolean           ' Indicateur d'initialisation
    StopOnFirstFailure As Boolean    ' Arrêter l'exécution au premier échec
    ThrowOnAssertFail As Boolean     ' Lever une erreur lors d'un échec d'assertion
End Type

Private this As TAssertState

'------------------------------------------------------------------------------
' Initialise l'objet d'assertions
'
' @param stopOnFirstFailure Arrêter l'exécution après le premier échec
' @param throwOnAssertFail Lever une erreur lors d'un échec d'assertion
'------------------------------------------------------------------------------
Public Sub Initialize(Optional ByVal stopOnFirstFailure As Boolean = False, _
                     Optional ByVal throwOnAssertFail As Boolean = False)
    this.FailureCount = 0
    this.LastFailureMessage = ""
    this.StopOnFirstFailure = stopOnFirstFailure
    this.ThrowOnAssertFail = throwOnAssertFail
    this.Initialized = True
End Sub

'------------------------------------------------------------------------------
' Accesseur pour le nombre d'échecs
'
' @return Long Nombre d'échecs d'assertions
'------------------------------------------------------------------------------
Public Property Get FailureCount() As Long
    FailureCount = this.FailureCount
End Property

'------------------------------------------------------------------------------
' Accesseur pour le dernier message d'échec
'
' @return String Dernier message d'échec
'------------------------------------------------------------------------------
Public Property Get LastFailureMessage() As String
    LastFailureMessage = this.LastFailureMessage
End Property

'------------------------------------------------------------------------------
' Vérifie si une condition est vraie
'
' @param condition Condition booléenne à vérifier
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsTrue(ByVal condition As Boolean, Optional ByVal message As String = "La condition doit être vraie") As Boolean
    If Not condition Then
        RecordFailure message
        IsTrue = False
    Else
        IsTrue = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une condition est fausse
'
' @param condition Condition booléenne à vérifier
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsFalse(ByVal condition As Boolean, Optional ByVal message As String = "La condition doit être fausse") As Boolean
    If condition Then
        RecordFailure message
        IsFalse = False
    Else
        IsFalse = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si deux valeurs sont égales
'
' @param expected Valeur attendue
' @param actual Valeur réelle
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function AreEqual(ByVal expected As Variant, ByVal actual As Variant, _
                        Optional ByVal message As String = "") As Boolean
    
    ' Gérer différemment les tableaux et les objets
    If IsObject(expected) Or IsObject(actual) Then
        If IsObject(expected) And IsObject(actual) Then
            ' Les deux sont des objets
            If expected Is actual Then
                AreEqual = True
                Exit Function
            End If
        End If
        
        ' Si on arrive ici, les objets ne sont pas les mêmes
        If Len(message) = 0 Then
            message = "Les objets doivent être identiques."
        End If
        RecordFailure message & " Expected: [Object], Actual: [Object]"
        AreEqual = False
        Exit Function
    End If
    
    ' Traiter les tableaux
    If IsArray(expected) And IsArray(actual) Then
        If AreArraysEqual(expected, actual) Then
            AreEqual = True
        Else
            If Len(message) = 0 Then
                message = "Les tableaux doivent être égaux."
            End If
            RecordFailure message & " Expected: " & ArrayToString(expected) & ", Actual: " & ArrayToString(actual)
            AreEqual = False
        End If
        Exit Function
    End If
    
    ' Valeurs scalaires
    If expected = actual Then
        AreEqual = True
    Else
        If Len(message) = 0 Then
            message = "Les valeurs doivent être égales."
        End If
        RecordFailure message & " Expected: " & CStr(expected) & ", Actual: " & CStr(actual)
        AreEqual = False
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si deux valeurs ne sont pas égales
'
' @param notExpected Valeur non attendue
' @param actual Valeur réelle
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function AreNotEqual(ByVal notExpected As Variant, ByVal actual As Variant, _
                           Optional ByVal message As String = "") As Boolean
    
    ' Inverson de AreEqual
    If AreEqual(notExpected, actual, "") Then
        If Len(message) = 0 Then
            message = "Les valeurs ne doivent pas être égales."
        End If
        RecordFailure message & " NotExpected: " & CStr(notExpected) & ", Actual: " & CStr(actual)
        AreNotEqual = False
    Else
        AreNotEqual = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une valeur est Nothing
'
' @param obj Objet à vérifier
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsNothing(ByVal obj As Object, Optional ByVal message As String = "L'objet doit être Nothing") As Boolean
    If Not obj Is Nothing Then
        RecordFailure message
        IsNothing = False
    Else
        IsNothing = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une valeur n'est pas Nothing
'
' @param obj Objet à vérifier
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsNotNothing(ByVal obj As Object, Optional ByVal message As String = "L'objet ne doit pas être Nothing") As Boolean
    If obj Is Nothing Then
        RecordFailure message
        IsNotNothing = False
    Else
        IsNotNothing = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une valeur est nulle (Empty, Null, vbNullString, Nothing)
'
' @param value Valeur à vérifier
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsNull(ByVal value As Variant, Optional ByVal message As String = "La valeur doit être nulle") As Boolean
    If IsObject(value) Then
        If Not value Is Nothing Then
            RecordFailure message
            IsNull = False
            Exit Function
        End If
    ElseIf Not IsEmpty(value) And Not IsNull(value) And value <> vbNullString And Not IsMissing(value) Then
        RecordFailure message
        IsNull = False
        Exit Function
    End If
    
    IsNull = True
End Function

'------------------------------------------------------------------------------
' Vérifie si une valeur n'est pas nulle
'
' @param value Valeur à vérifier
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsNotNull(ByVal value As Variant, Optional ByVal message As String = "La valeur ne doit pas être nulle") As Boolean
    If IsObject(value) Then
        If value Is Nothing Then
            RecordFailure message
            IsNotNull = False
            Exit Function
        End If
    ElseIf IsEmpty(value) Or IsNull(value) Or value = vbNullString Or IsMissing(value) Then
        RecordFailure message
        IsNotNull = False
        Exit Function
    End If
    
    IsNotNull = True
End Function

'------------------------------------------------------------------------------
' Vérifie si une chaîne contient une sous-chaîne
'
' @param str Chaîne à vérifier
' @param substring Sous-chaîne à rechercher
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function Contains(ByVal str As String, ByVal substring As String, _
                        Optional ByVal message As String = "") As Boolean
    If InStr(1, str, substring, vbTextCompare) = 0 Then
        If Len(message) = 0 Then
            message = "La chaîne doit contenir '" & substring & "'"
        End If
        RecordFailure message
        Contains = False
    Else
        Contains = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une chaîne ne contient pas une sous-chaîne
'
' @param str Chaîne à vérifier
' @param substring Sous-chaîne à ne pas trouver
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function NotContains(ByVal str As String, ByVal substring As String, _
                           Optional ByVal message As String = "") As Boolean
    If InStr(1, str, substring, vbTextCompare) > 0 Then
        If Len(message) = 0 Then
            message = "La chaîne ne doit pas contenir '" & substring & "'"
        End If
        RecordFailure message
        NotContains = False
    Else
        NotContains = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une valeur est entre deux limites (incluses)
'
' @param value Valeur à vérifier
' @param lowerBound Limite inférieure
' @param upperBound Limite supérieure
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsBetween(ByVal value As Variant, ByVal lowerBound As Variant, _
                         ByVal upperBound As Variant, _
                         Optional ByVal message As String = "") As Boolean
    If value < lowerBound Or value > upperBound Then
        If Len(message) = 0 Then
            message = "La valeur doit être entre " & CStr(lowerBound) & " et " & CStr(upperBound) & " (inclus)"
        End If
        RecordFailure message & " Actual: " & CStr(value)
        IsBetween = False
    Else
        IsBetween = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si une valeur est proche d'une autre avec une tolérance
'
' @param expected Valeur attendue
' @param actual Valeur réelle
' @param tolerance Tolérance acceptée
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function IsCloseTo(ByVal expected As Double, ByVal actual As Double, _
                         ByVal tolerance As Double, _
                         Optional ByVal message As String = "") As Boolean
    If Abs(expected - actual) > tolerance Then
        If Len(message) = 0 Then
            message = "La valeur doit être proche de " & CStr(expected) & " avec une tolérance de " & CStr(tolerance)
        End If
        RecordFailure message & " Actual: " & CStr(actual)
        IsCloseTo = False
    Else
        IsCloseTo = True
    End If
End Function

'------------------------------------------------------------------------------
' Force un échec d'assertion avec un message personnalisé
'
' @param message Message d'erreur
' @return Boolean Toujours False
'------------------------------------------------------------------------------
Public Function Fail(ByVal message As String) As Boolean
    RecordFailure message
    Fail = False
End Function

'------------------------------------------------------------------------------
' Vérifie si une collection contient un élément
'
' @param coll Collection à vérifier
' @param item Élément à rechercher (par valeur pour les types simples, par référence pour les objets)
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function CollectionContains(ByVal coll As Collection, ByVal item As Variant, _
                                 Optional ByVal message As String = "") As Boolean
    Dim i As Long
    Dim found As Boolean
    Dim collItem As Variant
    
    found = False
    
    For i = 1 To coll.Count
        If IsObject(item) Then
            If IsObject(coll(i)) Then
                If coll(i) Is item Then
                    found = True
                    Exit For
                End If
            End If
        Else
            If coll(i) = item Then
                found = True
                Exit For
            End If
        End If
    Next i
    
    If Not found Then
        If Len(message) = 0 Then
            message = "La collection doit contenir l'élément spécifié"
        End If
        RecordFailure message
        CollectionContains = False
    Else
        CollectionContains = True
    End If
End Function

'------------------------------------------------------------------------------
' Vérifie si un dictionnaire contient une clé
'
' @param dict Dictionnaire à vérifier
' @param key Clé à rechercher
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function DictionaryContainsKey(ByVal dict As Object, ByVal key As Variant, _
                                    Optional ByVal message As String = "") As Boolean
    On Error Resume Next
    Dim value As Variant
    value = dict(key)
    
    If Err.Number <> 0 Then
        If Len(message) = 0 Then
            message = "Le dictionnaire doit contenir la clé spécifiée: " & CStr(key)
        End If
        RecordFailure message
        DictionaryContainsKey = False
    Else
        DictionaryContainsKey = True
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Vérifie si une opération lève une erreur
'
' @param errorNumber Numéro d'erreur attendu (0 pour n'importe quelle erreur)
' @param procedureName Nom de la procédure contenant l'opération
' @param message Message d'erreur optionnel
' @return Boolean Résultat de l'assertion
'------------------------------------------------------------------------------
Public Function Throws(ByVal errorNumber As Long, ByVal procedureName As String, _
                      Optional ByVal message As String = "") As Boolean
    On Error Resume Next
    
    ' Exécuter le code qui devrait lever une erreur
    Application.Run procedureName
    
    ' Vérifier si une erreur a été levée
    If Err.Number = 0 Then
        ' Aucune erreur n'a été levée
        If Len(message) = 0 Then
            message = "Une erreur était attendue mais aucune n'a été levée"
        End If
        Err.Clear
        RecordFailure message
        Throws = False
    ElseIf errorNumber <> 0 And Err.Number <> errorNumber Then
        ' Erreur incorrecte levée
        If Len(message) = 0 Then
            message = "Erreur attendue: " & errorNumber & ", Erreur réelle: " & Err.Number & " - " & Err.Description
        End If
        Err.Clear
        RecordFailure message
        Throws = False
    Else
        ' Bonne erreur levée
        Err.Clear
        Throws = True
    End If
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Enregistre un échec d'assertion
'
' @param message Message d'erreur
'------------------------------------------------------------------------------
Private Sub RecordFailure(ByVal message As String)
    this.FailureCount = this.FailureCount + 1
    this.LastFailureMessage = message
    
    Debug.Print "ÉCHEC D'ASSERTION: " & message
    
    If this.ThrowOnAssertFail Then
        Err.Raise vbObjectError + 9000, "clsAssert", message
    End If
    
    If this.StopOnFirstFailure Then
        Stop ' Arrête l'exécution en mode debug
    End If
End Sub

'------------------------------------------------------------------------------
' Vérifie si deux tableaux sont égaux
'
' @param arr1 Premier tableau
' @param arr2 Deuxième tableau
' @return Boolean Résultat de la comparaison
'------------------------------------------------------------------------------
Private Function AreArraysEqual(ByVal arr1 As Variant, ByVal arr2 As Variant) As Boolean
    ' Vérification des dimensions
    Dim dims1 As Long, dims2 As Long
    
    On Error Resume Next
    dims1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
    If Err.Number <> 0 Then
        AreArraysEqual = False
        Exit Function
    End If
    
    dims2 = UBound(arr2, 1) - LBound(arr2, 1) + 1
    If Err.Number <> 0 Then
        AreArraysEqual = False
        Exit Function
    End If
    On Error GoTo 0
    
    If dims1 <> dims2 Then
        AreArraysEqual = False
        Exit Function
    End If
    
    ' Vérification des éléments
    Dim i As Long, j As Long
    Dim isEqual As Boolean
    
    isEqual = True
    
    ' Tableau unidimensionnel
    If dims1 = 1 Then
        If UBound(arr1) - LBound(arr1) <> UBound(arr2) - LBound(arr2) Then
            AreArraysEqual = False
            Exit Function
        End If
        
        For i = LBound(arr1) To UBound(arr1)
            If arr1(i) <> arr2(i + LBound(arr2) - LBound(arr1)) Then
                isEqual = False
                Exit For
            End If
        Next i
    Else
        ' Tableau bidimensionnel
        If UBound(arr1, 1) - LBound(arr1, 1) <> UBound(arr2, 1) - LBound(arr2, 1) Or _
           UBound(arr1, 2) - LBound(arr1, 2) <> UBound(arr2, 2) - LBound(arr2, 2) Then
            AreArraysEqual = False
            Exit Function
        End If
        
        For i = LBound(arr1, 1) To UBound(arr1, 1)
            For j = LBound(arr1, 2) To UBound(arr1, 2)
                If arr1(i, j) <> arr2(i + LBound(arr2, 1) - LBound(arr1, 1), j + LBound(arr2, 2) - LBound(arr1, 2)) Then
                    isEqual = False
                    Exit For
                End If
            Next j
            If Not isEqual Then Exit For
        Next i
    End If
    
    AreArraysEqual = isEqual
End Function

'------------------------------------------------------------------------------
' Convertit un tableau en chaîne de caractères pour l'affichage
'
' @param arr Tableau à convertir
' @return String Représentation du tableau
'------------------------------------------------------------------------------
Private Function ArrayToString(ByVal arr As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim i As Long, j As Long
    
    result = "["
    
    ' Tableau unidimensionnel
    If UBound(arr, 1) - LBound(arr, 1) = UBound(arr) - LBound(arr) Then
        For i = LBound(arr) To UBound(arr)
            If i > LBound(arr) Then result = result & ", "
            result = result & CStr(arr(i))
        Next i
    Else
        ' Tableau bidimensionnel
        For i = LBound(arr, 1) To UBound(arr, 1)
            If i > LBound(arr, 1) Then result = result & "; "
            result = result & "["
            
            For j = LBound(arr, 2) To UBound(arr, 2)
                If j > LBound(arr, 2) Then result = result & ", "
                result = result & CStr(arr(i, j))
            Next j
            
            result = result & "]"
        Next i
    End If
    
    result = result & "]"
    ArrayToString = result
    Exit Function
    
ErrorHandler:
    ArrayToString = "[Conversion de tableau impossible]"
End Function 