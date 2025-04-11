' Migrated to apex-metier/xml - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Class : clsXmlParser
' Version : 1.0
' Purpose : Analyse et manipulation de documents XML
' Date : 10/04/2025
' ==========================================================================

' --- Références requises ---
' Microsoft XML (MSXML2)

' --- Variables privées ---
Private m_xmlDocument As Object    ' MSXML2.DOMDocument60
Private m_xmlNamespaces As Object  ' Dictionary de namespaces
Private m_LastError As String
Private m_Logger As Object         ' ILoggerBase
Private m_IsInitialized As Boolean
Private m_PreserveWhitespace As Boolean
Private m_ValidateOnParse As Boolean
Private m_ResolveExternals As Boolean

' --- Initialisation ---
Private Sub Class_Initialize()
    ' Initialisation par défaut
    m_LastError = ""
    m_IsInitialized = False
    m_PreserveWhitespace = False
    m_ValidateOnParse = False
    m_ResolveExternals = False
    Set m_xmlNamespaces = CreateObject("Scripting.Dictionary")
    
    InitializeXmlDocument
End Sub

Private Sub Class_Terminate()
    ' Nettoyage
    Set m_xmlDocument = Nothing
    Set m_xmlNamespaces = Nothing
    Set m_Logger = Nothing
End Sub

' --- Propriétés ---
Public Property Get XmlDocument() As Object
    ' Retourne le document XML sous-jacent (MSXML2.DOMDocument)
    Set XmlDocument = m_xmlDocument
End Property

Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

Public Property Let PreserveWhitespace(ByVal value As Boolean)
    ' Définit si les espaces blancs sont préservés
    m_PreserveWhitespace = value
    If Not m_xmlDocument Is Nothing Then
        m_xmlDocument.preserveWhiteSpace = value
    End If
End Property

Public Property Let ValidateOnParse(ByVal value As Boolean)
    ' Définit si la validation est effectuée lors de l'analyse
    m_ValidateOnParse = value
    If Not m_xmlDocument Is Nothing Then
        m_xmlDocument.validateOnParse = value
    End If
End Property

Public Property Let ResolveExternals(ByVal value As Boolean)
    ' Définit si les références externes sont résolues
    m_ResolveExternals = value
    If Not m_xmlDocument Is Nothing Then
        m_xmlDocument.resolveExternals = value
    End If
End Property

Public Property Set Logger(ByVal value As Object)
    ' Définit le logger à utiliser
    Set m_Logger = value
End Property

' --- Méthodes publiques ---
Public Function LoadXml(ByVal xmlString As String) As Boolean
    ' Charge du XML à partir d'une chaîne
    On Error GoTo ErrorHandler
    
    InitializeXmlDocument
    
    m_xmlDocument.LoadXML xmlString
    LoadXml = m_xmlDocument.parseError.errorCode = 0
    
    If Not LoadXml Then
        m_LastError = "Erreur d'analyse XML: " & m_xmlDocument.parseError.reason & _
                     " Ligne: " & m_xmlDocument.parseError.Line & _
                     " Position: " & m_xmlDocument.parseError.linepos
        LogError m_LastError
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur de chargement XML: " & Err.Description
    LogError m_LastError
    LoadXml = False
End Function

Public Function LoadXmlFromFile(ByVal filePath As String) As Boolean
    ' Charge du XML à partir d'un fichier
    On Error GoTo ErrorHandler
    
    InitializeXmlDocument
    
    m_xmlDocument.Load filePath
    LoadXmlFromFile = m_xmlDocument.parseError.errorCode = 0
    
    If Not LoadXmlFromFile Then
        m_LastError = "Erreur d'analyse XML depuis fichier: " & m_xmlDocument.parseError.reason & _
                     " Ligne: " & m_xmlDocument.parseError.Line & _
                     " Position: " & m_xmlDocument.parseError.linepos
        LogError m_LastError
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur de chargement XML depuis fichier: " & Err.Description
    LogError m_LastError
    LoadXmlFromFile = False
End Function

Public Function SaveToFile(ByVal filePath As String) As Boolean
    ' Sauvegarde le document XML dans un fichier
    On Error GoTo ErrorHandler
    
    If m_xmlDocument Is Nothing Then
        m_LastError = "Aucun document XML à sauvegarder"
        LogError m_LastError
        SaveToFile = False
        Exit Function
    End If
    
    m_xmlDocument.Save filePath
    SaveToFile = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur de sauvegarde XML: " & Err.Description
    LogError m_LastError
    SaveToFile = False
End Function

Public Function GetXmlString() As String
    ' Retourne le document XML sous forme de chaîne
    On Error GoTo ErrorHandler
    
    If m_xmlDocument Is Nothing Then
        m_LastError = "Aucun document XML à convertir en chaîne"
        LogError m_LastError
        GetXmlString = ""
        Exit Function
    End If
    
    GetXmlString = m_xmlDocument.xml
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la conversion XML en chaîne: " & Err.Description
    LogError m_LastError
    GetXmlString = ""
End Function

Public Function SelectSingleNode(ByVal xPath As String) As Object
    ' Sélectionne un nœud XML unique via XPath
    On Error GoTo ErrorHandler
    
    If m_xmlDocument Is Nothing Then
        m_LastError = "Aucun document XML pour sélectionner un nœud"
        LogError m_LastError
        Set SelectSingleNode = Nothing
        Exit Function
    End If
    
    If m_xmlNamespaces.Count > 0 Then
        ' Utiliser la sélection avec namespaces
        Set SelectSingleNode = SelectSingleNodeWithNamespaces(xPath)
    Else
        ' Sélection standard
        Set SelectSingleNode = m_xmlDocument.SelectSingleNode(xPath)
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la sélection d'un nœud: " & Err.Description
    LogError m_LastError
    Set SelectSingleNode = Nothing
End Function

Public Function SelectNodes(ByVal xPath As String) As Object
    ' Sélectionne plusieurs nœuds XML via XPath
    On Error GoTo ErrorHandler
    
    If m_xmlDocument Is Nothing Then
        m_LastError = "Aucun document XML pour sélectionner des nœuds"
        LogError m_LastError
        Set SelectNodes = Nothing
        Exit Function
    End If
    
    If m_xmlNamespaces.Count > 0 Then
        ' Utiliser la sélection avec namespaces
        Set SelectNodes = SelectNodesWithNamespaces(xPath)
    Else
        ' Sélection standard
        Set SelectNodes = m_xmlDocument.SelectNodes(xPath)
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la sélection de nœuds: " & Err.Description
    LogError m_LastError
    Set SelectNodes = Nothing
End Function

Public Sub AddNamespace(ByVal prefix As String, ByVal uri As String)
    ' Ajoute un namespace pour les requêtes XPath
    m_xmlNamespaces(prefix) = uri
End Sub

Public Function CreateNode(ByVal nodeType As Long, ByVal nodeName As String, Optional ByVal namespaceURI As String = "") As Object
    ' Crée un nouveau nœud
    On Error GoTo ErrorHandler
    
    If m_xmlDocument Is Nothing Then
        InitializeXmlDocument
    End If
    
    Set CreateNode = m_xmlDocument.createNode(nodeType, nodeName, namespaceURI)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la création d'un nœud: " & Err.Description
    LogError m_LastError
    Set CreateNode = Nothing
End Function

Public Function CreateElement(ByVal elementName As String, Optional ByVal namespaceURI As String = "") As Object
    ' Crée un nouvel élément
    On Error GoTo ErrorHandler
    
    If m_xmlDocument Is Nothing Then
        InitializeXmlDocument
    End If
    
    If namespaceURI = "" Then
        Set CreateElement = m_xmlDocument.createElement(elementName)
    Else
        Set CreateElement = m_xmlDocument.createNode(1, elementName, namespaceURI) ' 1 = NODE_ELEMENT
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la création d'un élément: " & Err.Description
    LogError m_LastError
    Set CreateElement = Nothing
End Function

Public Function GetNodeValue(ByVal node As Object) As String
    ' Obtient la valeur d'un nœud
    On Error GoTo ErrorHandler
    
    If node Is Nothing Then
        GetNodeValue = ""
        Exit Function
    End If
    
    ' Vérifier le type de nœud
    Select Case node.nodeType
        Case 1 ' NODE_ELEMENT
            GetNodeValue = node.Text
        Case 2 ' NODE_ATTRIBUTE
            GetNodeValue = node.value
        Case 3 ' NODE_TEXT
            GetNodeValue = node.nodeValue
        Case Else
            GetNodeValue = node.nodeValue
    End Select
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention de la valeur d'un nœud: " & Err.Description
    LogError m_LastError
    GetNodeValue = ""
End Function

Public Function GetAttribute(ByVal node As Object, ByVal attributeName As String) As String
    ' Obtient la valeur d'un attribut
    On Error GoTo ErrorHandler
    
    If node Is Nothing Then
        GetAttribute = ""
        Exit Function
    End If
    
    ' Vérifier que c'est un élément
    If node.nodeType <> 1 Then ' NODE_ELEMENT
        GetAttribute = ""
        Exit Function
    End If
    
    ' Obtenir l'attribut
    If node.hasAttribute(attributeName) Then
        GetAttribute = node.getAttribute(attributeName)
    Else
        GetAttribute = ""
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention d'un attribut: " & Err.Description
    LogError m_LastError
    GetAttribute = ""
End Function

' --- Méthodes privées ---
Private Sub InitializeXmlDocument()
    ' Initialise ou réinitialise le document XML
    On Error GoTo ErrorHandler
    
    ' Créer un nouveau document
    Set m_xmlDocument = CreateObject("MSXML2.DOMDocument.6.0")
    
    ' Appliquer les options
    m_xmlDocument.async = False
    m_xmlDocument.preserveWhiteSpace = m_PreserveWhitespace
    m_xmlDocument.validateOnParse = m_ValidateOnParse
    m_xmlDocument.resolveExternals = m_ResolveExternals
    
    m_IsInitialized = True
    
    Exit Sub
    
ErrorHandler:
    m_LastError = "Erreur d'initialisation du document XML: " & Err.Description
    LogError m_LastError
    m_IsInitialized = False
    
    ' Essayer avec une version plus ancienne de MSXML en cas d'échec
    On Error Resume Next
    Set m_xmlDocument = CreateObject("MSXML2.DOMDocument")
    
    If Not m_xmlDocument Is Nothing Then
        m_xmlDocument.async = False
        m_xmlDocument.preserveWhiteSpace = m_PreserveWhitespace
        m_xmlDocument.validateOnParse = m_ValidateOnParse
        m_xmlDocument.resolveExternals = m_ResolveExternals
        m_IsInitialized = True
    End If
End Sub

Private Function SelectSingleNodeWithNamespaces(ByVal xPath As String) As Object
    ' Sélectionne un nœud avec support des namespaces
    Dim nsManager As Object
    
    ' Créer un gestionnaire de namespaces
    Set nsManager = m_xmlDocument.createNode(1, "nsManager", "")
    
    ' Ajouter tous les namespaces
    Dim prefix As Variant
    For Each prefix In m_xmlNamespaces.Keys
        nsManager.setAttributeNS "xmlns:" & prefix, "", m_xmlNamespaces(prefix)
    Next
    
    ' Effectuer la sélection
    Set SelectSingleNodeWithNamespaces = m_xmlDocument.SelectSingleNode(xPath, nsManager)
End Function

Private Function SelectNodesWithNamespaces(ByVal xPath As String) As Object
    ' Sélectionne plusieurs nœuds avec support des namespaces
    Dim nsManager As Object
    
    ' Créer un gestionnaire de namespaces
    Set nsManager = m_xmlDocument.createNode(1, "nsManager", "")
    
    ' Ajouter tous les namespaces
    Dim prefix As Variant
    For Each prefix In m_xmlNamespaces.Keys
        nsManager.setAttributeNS "xmlns:" & prefix, "", m_xmlNamespaces(prefix)
    Next
    
    ' Effectuer la sélection
    Set SelectNodesWithNamespaces = m_xmlDocument.SelectNodes(xPath, nsManager)
End Function

Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "XML"
    End If
End Sub
