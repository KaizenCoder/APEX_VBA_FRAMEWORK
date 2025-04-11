' Migrated to apex-core/testing - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Attribute VB_Name = "modTestAssertions"
Option Explicit
' ==========================================================================
' Module : modTestAssertions
' Version : 2.0
' Purpose : Module de fonctions d'assertion pour les tests unitaires
' Date    : 10/04/2025
' ==========================================================================

' --- Constantes ---
Private Const MODULE_NAME As String = "modTestAssertions"
Private Const PRECISION_DOUBLE As Double = 0.0000000001 ' 10^-10

' --- Variables privées ---
Private m_successCount As Long
Private m_failureCount As Long
Private m_logger As Object
Private m_configManager As Object
Private m_traceAssertions As Boolean

' --- Initialisation ---
Public Sub Initialize()
    m_successCount = 0
    m_failureCount = 0
    
    ' Initialiser les dépendances
    On Error Resume Next
    Set m_logger = CreateObject("APEX.Logger")
    If Err.Number <> 0 Then Set m_logger = Nothing
    
    Set m_configManager = CreateObject("APEX.ConfigManager")
    If Err.Number <> 0 Then Set m_configManager = Nothing
    Err.Clear
    On Error GoTo 0
    
    ' Vérifier si le traçage des assertions est activé
    If Not m_configManager Is Nothing Then
        m_traceAssertions = (m_configManager.GetSetting("Debug", "TraceAssertions", "True") = "True")
    Else
        m_traceAssertions = True
    End If
End Sub

' --- Assertions de base ---
Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "La condition devrait être True")
    If condition Then
        OnSuccess "AssertTrue", message
    Else
        OnFailure "AssertTrue", message
        Err.Raise 9999, "AssertTrue", message
    End If
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "La condition devrait être False")
    If Not condition Then
        OnSuccess "AssertFalse", message
    Else
        OnFailure "AssertFalse", message
        Err.Raise 9999, "AssertFalse", message
    End If
End Sub

Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    Dim isEqual As Boolean
    Dim actualType As VbVarType
    Dim expectedType As VbVarType
    
    ' Traiter les types particuliers
    actualType = VarType(actual)
    expectedType = VarType(expected)
    
    ' Si les types sont numériques
    If IsNumeric(expected) And IsNumeric(actual) Then
        ' Utiliser AssertAlmostEqual pour les doubles
        If actualType = vbDouble Or expectedType = vbDouble Then
            AssertAlmostEqual CDbl(expected), CDbl(actual), PRECISION_DOUBLE, IIf(message = "", "Les valeurs devraient être égales", message)
            Exit Sub
        Else
            isEqual = (CDbl(expected) = CDbl(actual))
        End If
    ' Si les types sont des dates
    ElseIf IsDate(expected) And IsDate(actual) Then
        isEqual = (CDate(expected) = CDate(actual))
    ' Si les types sont des objets
    ElseIf (actualType >= vbObject) And (expectedType >= vbObject) Then
        ' Impossible de comparer des objets directement, utiliser Is
        If expected Is actual Then
            isEqual = True
        Else
            isEqual = False
        End If
    ' Comparaison standard pour les autres types
    Else
        isEqual = (expected = actual)
    End If
    
    If isEqual Then
        OnSuccess "AssertEqual", "Valeurs égales: " & ToString(expected)
    Else
        Dim errorMsg As String
        errorMsg = "Valeurs différentes - Attendu: " & ToString(expected) & ", Obtenu: " & ToString(actual)
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertEqual", errorMsg
        Err.Raise 9999, "AssertEqual", errorMsg
    End If
End Sub

Public Sub AssertNotEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    Dim isEqual As Boolean
    Dim actualType As VbVarType
    Dim expectedType As VbVarType
    
    ' Traiter les types particuliers
    actualType = VarType(actual)
    expectedType = VarType(expected)
    
    ' Si les types sont numériques
    If IsNumeric(expected) And IsNumeric(actual) Then
        isEqual = (CDbl(expected) = CDbl(actual))
    ' Si les types sont des dates
    ElseIf IsDate(expected) And IsDate(actual) Then
        isEqual = (CDate(expected) = CDate(actual))
    ' Si les types sont des objets
    ElseIf (actualType >= vbObject) And (expectedType >= vbObject) Then
        ' Impossible de comparer des objets directement, utiliser Is
        If expected Is actual Then
            isEqual = True
        Else
            isEqual = False
        End If
    ' Comparaison standard pour les autres types
    Else
        isEqual = (expected = actual)
    End If
    
    If Not isEqual Then
        OnSuccess "AssertNotEqual", "Valeurs différentes comme attendu"
    Else
        Dim errorMsg As String
        errorMsg = "Les valeurs ne devraient pas être égales: " & ToString(expected)
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertNotEqual", errorMsg
        Err.Raise 9999, "AssertNotEqual", errorMsg
    End If
End Sub

Public Sub AssertIsNothing(ByVal obj As Object, Optional ByVal message As String = "L'objet devrait être Nothing")
    If obj Is Nothing Then
        OnSuccess "AssertIsNothing", "L'objet est Nothing"
    Else
        OnFailure "AssertIsNothing", message
        Err.Raise 9999, "AssertIsNothing", message
    End If
End Sub

Public Sub AssertIsNotNothing(ByVal obj As Object, Optional ByVal message As String = "L'objet ne devrait pas être Nothing")
    If Not obj Is Nothing Then
        OnSuccess "AssertIsNotNothing", "L'objet n'est pas Nothing"
    Else
        OnFailure "AssertIsNotNothing", message
        Err.Raise 9999, "AssertIsNotNothing", message
    End If
End Sub

Public Sub AssertGreaterThan(ByVal value1 As Variant, ByVal value2 As Variant, Optional ByVal message As String = "")
    If IsNumeric(value1) And IsNumeric(value2) Then
        If CDbl(value1) > CDbl(value2) Then
            OnSuccess "AssertGreaterThan", "La valeur " & ToString(value1) & " est supérieure à " & ToString(value2)
        Else
            Dim errorMsg As String
            errorMsg = "La valeur " & ToString(value1) & " n'est pas supérieure à " & ToString(value2)
            If message <> "" Then errorMsg = message & " - " & errorMsg
            
            OnFailure "AssertGreaterThan", errorMsg
            Err.Raise 9999, "AssertGreaterThan", errorMsg
        End If
    Else
        OnFailure "AssertGreaterThan", "Les valeurs comparées doivent être numériques"
        Err.Raise 9999, "AssertGreaterThan", "Les valeurs comparées doivent être numériques"
    End If
End Sub

Public Sub AssertLessThan(ByVal value1 As Variant, ByVal value2 As Variant, Optional ByVal message As String = "")
    If IsNumeric(value1) And IsNumeric(value2) Then
        If CDbl(value1) < CDbl(value2) Then
            OnSuccess "AssertLessThan", "La valeur " & ToString(value1) & " est inférieure à " & ToString(value2)
        Else
            Dim errorMsg As String
            errorMsg = "La valeur " & ToString(value1) & " n'est pas inférieure à " & ToString(value2)
            If message <> "" Then errorMsg = message & " - " & errorMsg
            
            OnFailure "AssertLessThan", errorMsg
            Err.Raise 9999, "AssertLessThan", errorMsg
        End If
    Else
        OnFailure "AssertLessThan", "Les valeurs comparées doivent être numériques"
        Err.Raise 9999, "AssertLessThan", "Les valeurs comparées doivent être numériques"
    End If
End Sub

Public Sub AssertAlmostEqual(ByVal expected As Double, ByVal actual As Double, _
                            Optional ByVal tolerance As Double = 0.0001, _
                            Optional ByVal message As String = "")
    If Abs(expected - actual) <= tolerance Then
        OnSuccess "AssertAlmostEqual", "Les valeurs " & expected & " et " & actual & " sont presque égales (tolérance: " & tolerance & ")"
    Else
        Dim errorMsg As String
        errorMsg = "Les valeurs " & expected & " et " & actual & " ne sont pas assez proches (différence: " & Abs(expected - actual) & ", tolérance: " & tolerance & ")"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertAlmostEqual", errorMsg
        Err.Raise 9999, "AssertAlmostEqual", errorMsg
    End If
End Sub

' --- Assertions de chaînes ---
Public Sub AssertStringContains(ByVal str As String, ByVal subStr As String, Optional ByVal caseSensitive As Boolean = True, Optional ByVal message As String = "")
    Dim result As Boolean
    
    If caseSensitive Then
        result = (InStr(1, str, subStr, vbBinaryCompare) > 0)
    Else
        result = (InStr(1, str, subStr, vbTextCompare) > 0)
    End If
    
    If result Then
        OnSuccess "AssertStringContains", "La chaîne contient '" & subStr & "'"
    Else
        Dim errorMsg As String
        errorMsg = "La chaîne ne contient pas '" & subStr & "'"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertStringContains", errorMsg
        Err.Raise 9999, "AssertStringContains", errorMsg
    End If
End Sub

Public Sub AssertStringStartsWith(ByVal str As String, ByVal prefix As String, Optional ByVal caseSensitive As Boolean = True, Optional ByVal message As String = "")
    Dim result As Boolean
    
    If Len(str) < Len(prefix) Then
        result = False
    Else
        If caseSensitive Then
            result = (Left(str, Len(prefix)) = prefix)
        Else
            result = (StrComp(Left(str, Len(prefix)), prefix, vbTextCompare) = 0)
        End If
    End If
    
    If result Then
        OnSuccess "AssertStringStartsWith", "La chaîne commence par '" & prefix & "'"
    Else
        Dim errorMsg As String
        errorMsg = "La chaîne ne commence pas par '" & prefix & "'"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertStringStartsWith", errorMsg
        Err.Raise 9999, "AssertStringStartsWith", errorMsg
    End If
End Sub

Public Sub AssertStringEndsWith(ByVal str As String, ByVal suffix As String, Optional ByVal caseSensitive As Boolean = True, Optional ByVal message As String = "")
    Dim result As Boolean
    
    If Len(str) < Len(suffix) Then
        result = False
    Else
        If caseSensitive Then
            result = (Right(str, Len(suffix)) = suffix)
        Else
            result = (StrComp(Right(str, Len(suffix)), suffix, vbTextCompare) = 0)
        End If
    End If
    
    If result Then
        OnSuccess "AssertStringEndsWith", "La chaîne se termine par '" & suffix & "'"
    Else
        Dim errorMsg As String
        errorMsg = "La chaîne ne se termine pas par '" & suffix & "'"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertStringEndsWith", errorMsg
        Err.Raise 9999, "AssertStringEndsWith", errorMsg
    End If
End Sub

Public Sub AssertStringMatches(ByVal str As String, ByVal pattern As String, Optional ByVal message As String = "")
    Dim regex As Object
    Dim result As Boolean
    
    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp")
    
    If Err.Number <> 0 Then
        OnFailure "AssertStringMatches", "Impossible de créer l'objet RegExp. Vérifiez que Microsoft VBScript Regular Expressions est référencé."
        Err.Raise 9999, "AssertStringMatches", "Impossible de créer l'objet RegExp"
        Exit Sub
    End If
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = pattern
    End With
    
    result = regex.Test(str)
    On Error GoTo 0
    
    If result Then
        OnSuccess "AssertStringMatches", "La chaîne correspond au motif '" & pattern & "'"
    Else
        Dim errorMsg As String
        errorMsg = "La chaîne ne correspond pas au motif '" & pattern & "'"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertStringMatches", errorMsg
        Err.Raise 9999, "AssertStringMatches", errorMsg
    End If
End Sub

' --- Assertions de collections ---
Public Sub AssertCollectionContains(ByVal coll As Collection, ByVal item As Variant, Optional ByVal message As String = "")
    Dim found As Boolean
    Dim i As Long
    Dim itemStr As String
    
    found = False
    itemStr = ToString(item)
    
    ' Parcourir la collection
    For i = 1 To coll.Count
        If ToString(coll(i)) = itemStr Then
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        OnSuccess "AssertCollectionContains", "La collection contient l'élément '" & itemStr & "'"
    Else
        Dim errorMsg As String
        errorMsg = "La collection ne contient pas l'élément '" & itemStr & "'"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertCollectionContains", errorMsg
        Err.Raise 9999, "AssertCollectionContains", errorMsg
    End If
End Sub

Public Sub AssertCollectionNotEmpty(ByVal coll As Collection, Optional ByVal message As String = "La collection ne devrait pas être vide")
    If coll.Count > 0 Then
        OnSuccess "AssertCollectionNotEmpty", "La collection contient " & coll.Count & " élément(s)"
    Else
        OnFailure "AssertCollectionNotEmpty", message
        Err.Raise 9999, "AssertCollectionNotEmpty", message
    End If
End Sub

Public Sub AssertCollectionEmpty(ByVal coll As Collection, Optional ByVal message As String = "La collection devrait être vide")
    If coll.Count = 0 Then
        OnSuccess "AssertCollectionEmpty", "La collection est vide"
    Else
        Dim errorMsg As String
        errorMsg = "La collection contient " & coll.Count & " élément(s) alors qu'elle devrait être vide"
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertCollectionEmpty", errorMsg
        Err.Raise 9999, "AssertCollectionEmpty", errorMsg
    End If
End Sub

Public Sub AssertCollectionCount(ByVal coll As Collection, ByVal expectedCount As Long, Optional ByVal message As String = "")
    If coll.Count = expectedCount Then
        OnSuccess "AssertCollectionCount", "La collection contient " & expectedCount & " élément(s) comme attendu"
    Else
        Dim errorMsg As String
        errorMsg = "La collection contient " & coll.Count & " élément(s) au lieu de " & expectedCount
        If message <> "" Then errorMsg = message & " - " & errorMsg
        
        OnFailure "AssertCollectionCount", errorMsg
        Err.Raise 9999, "AssertCollectionCount", errorMsg
    End If
End Sub

' --- Assertions d'exception ---
Public Function AssertRaises(ByVal expectedErrNumber As Long) As Boolean
    ' Cette fonction doit être utilisée avec On Error Resume Next
    ' Exemple:
    '   On Error Resume Next
    '   Set obj = Nothing
    '   obj.Method ' Génère une erreur
    '   AssertRaises 91 ' Vérifie que l'erreur 91 s'est produite
    '   On Error GoTo 0
    
    If Err.Number = expectedErrNumber Then
        OnSuccess "AssertRaises", "L'erreur " & expectedErrNumber & " s'est bien produite: " & Err.Description
        AssertRaises = True
    Else
        Dim errorMsg As String
        If Err.Number = 0 Then
            errorMsg = "Aucune erreur ne s'est produite alors que l'erreur " & expectedErrNumber & " était attendue"
        Else
            errorMsg = "L'erreur " & Err.Number & " s'est produite au lieu de l'erreur " & expectedErrNumber & " - " & Err.Description
        End If
        
        OnFailure "AssertRaises", errorMsg
        AssertRaises = False
    End If
    
    Err.Clear
End Function

' --- Assertions d'exécution ---
Public Sub AssertExecutionTime(ByVal procedure As String, Optional ByVal maxTimeMs As Long = 1000, Optional ByVal params As Variant = Null)
    Dim startTime As Double
    Dim endTime As Double
    Dim durationMs As Long
    
    startTime = Timer
    
    On Error Resume Next
    If IsNull(params) Then
        Application.Run procedure
    Else
        Application.Run procedure, params
    End If
    
    If Err.Number <> 0 Then
        OnFailure "AssertExecutionTime", "Erreur lors de l'exécution de " & procedure & ": " & Err.Description
        Err.Raise 9999, "AssertExecutionTime", "Erreur lors de l'exécution: " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    
    endTime = Timer
    durationMs = (endTime - startTime) * 1000
    
    If durationMs <= maxTimeMs Then
        OnSuccess "AssertExecutionTime", "Exécution de " & procedure & " en " & durationMs & " ms (max: " & maxTimeMs & " ms)"
    Else
        OnFailure "AssertExecutionTime", "Exécution de " & procedure & " trop lente: " & durationMs & " ms (max: " & maxTimeMs & " ms)"
        Err.Raise 9999, "AssertExecutionTime", "Exécution trop lente: " & durationMs & " ms"
    End If
End Sub

' --- Statistiques ---
Public Function GetAssertionCounts() As String
    GetAssertionCounts = "Assertions: " & (m_successCount + m_failureCount) & _
                        " (Réussies: " & m_successCount & ", Échouées: " & m_failureCount & ")"
End Function

Public Sub ResetAssertionCounts()
    m_successCount = 0
    m_failureCount = 0
End Sub

' --- Fonctions utilitaires ---
Private Function ToString(ByVal value As Variant) As String
    Dim result As String
    
    On Error Resume Next
    
    Select Case VarType(value)
        Case vbNull
            result = "<Null>"
        Case vbEmpty
            result = "<Empty>"
        Case vbObject
            If value Is Nothing Then
                result = "<Nothing>"
            Else
                result = "<Object: " & TypeName(value) & ">"
            End If
        Case vbBoolean
            result = IIf(value, "True", "False")
        Case vbDate
            result = Format(value, "yyyy-mm-dd hh:nn:ss")
        Case vbArray To vbArray + vbByte
            result = "<Array>"
        Case Else
            result = CStr(value)
    End Select
    
    If Err.Number <> 0 Then
        result = "<Impossible de convertir: " & TypeName(value) & ">"
        Err.Clear
    End If
    
    On Error GoTo 0
    ToString = result
End Function

Private Sub OnSuccess(ByVal assertion As String, ByVal message As String)
    m_successCount = m_successCount + 1
    
    ' Tracer l'assertion si activé
    If m_traceAssertions Then
        LogMessage assertion & ": RÉUSSI - " & message, "debug"
    End If
End Sub

Private Sub OnFailure(ByVal assertion As String, ByVal message As String)
    m_failureCount = m_failureCount + 1
    
    ' Toujours tracer les échecs
    LogMessage assertion & ": ÉCHEC - " & message, "error"
End Sub

Private Sub LogMessage(ByVal message As String, ByVal logLevel As String)
    ' Écrire dans le journal si disponible
    On Error Resume Next
    If Not m_logger Is Nothing Then
        Select Case LCase(logLevel)
            Case "debug"
                m_logger.LogDebug MODULE_NAME, message
            Case "info"
                m_logger.LogInfo MODULE_NAME, message
            Case "warning"
                m_logger.LogWarning MODULE_NAME, message
            Case "error"
                m_logger.LogError MODULE_NAME, message
            Case Else
                m_logger.LogInfo MODULE_NAME, message
        End Select
    Else
        ' Écrire dans la fenêtre de débogage
        Debug.Print message
    End If
    On Error GoTo 0
End Sub 