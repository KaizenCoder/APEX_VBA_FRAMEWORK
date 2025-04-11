' Module   : modRibbonCallbacks
' Purpose  : Gestionnaire des callbacks du ruban personnalisé Apex
' Date     : 10/04/2025
' Version  : 1.0
' Author   : Équipe Apex Framework
' ===========================================================

Option Explicit

' Référence au logger global
Private Logger As clsLogger

' Initialisation du module
Private Sub Initialize()
    If Logger Is Nothing Then
        Set Logger = New clsLogger
        Logger.Initialize "RibbonCallbacks"
    End If
End Sub

' Exécute tous les tests unitaires
Public Sub OnAction_RunAllTests(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Initialize
    Logger.LogInfo "Ruban > Tests : Lancement de tous les tests"
    
    ' Appeler le module de test
    Call modTestRunner.RunAllTests
    
    Exit Sub
ErrorHandler:
    Logger.LogError "OnAction_RunAllTests : " & Err.Description
End Sub

' Affiche les logs du système
Public Sub OnAction_ViewLogs(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Initialize
    Logger.LogInfo "Ruban > Logs : Affichage du journal"
    
    ' Afficher le formulaire de logs (à implémenter)
    MsgBox "Fonctionnalité de visualisation des logs à implémenter", vbInformation
    
    Exit Sub
ErrorHandler:
    Logger.LogError "OnAction_ViewLogs : " & Err.Description
End Sub

' Chiffre une chaîne avec DPAPI
Public Sub OnAction_EncryptString(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Initialize
    Logger.LogInfo "Ruban > Sécurité : Cryptage d'une chaîne"
    
    Dim inputString As String
    Dim encryptedString As String
    
    ' Afficher une boîte de dialogue pour saisir la chaîne
    inputString = InputBox("Entrez la chaîne à crypter:", "Cryptage DPAPI")
    
    If Len(inputString) > 0 Then
        ' Appeler la fonction de cryptage
        encryptedString = modSecurityDPAPI.EncryptString(inputString)
        
        ' Afficher le résultat
        Debug.Print "Chaîne cryptée: " & encryptedString
        
        ' Copier dans le presse-papier
        With CreateObject("Forms.TextBox.1")
            .Text = encryptedString
            .SelStart = 0
            .SelLength = Len(.Text)
            .Copy
        End With
        
        MsgBox "La chaîne cryptée a été copiée dans le presse-papier", vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    Logger.LogError "OnAction_EncryptString : " & Err.Description
End Sub

' Lance la comparaison de recettes
Public Sub OnAction_RunRecipe(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Initialize
    Logger.LogInfo "Ruban > Recette : Lancement de la comparaison"
    
    ' Appeler le module de recette
    Call modRecipeComparer.CompareRecipes
    
    Exit Sub
ErrorHandler:
    Logger.LogError "OnAction_RunRecipe : " & Err.Description
End Sub

' Analyse un fichier XML
Public Sub OnAction_ParseXml(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Initialize
    Logger.LogInfo "Ruban > XML : Analyse d'un fichier XML"
    
    ' Code pour analyse XML (à implémenter)
    MsgBox "Fonctionnalité d'analyse XML à implémenter", vbInformation
    
    Exit Sub
ErrorHandler:
    Logger.LogError "OnAction_ParseXml : " & Err.Description
End Sub 