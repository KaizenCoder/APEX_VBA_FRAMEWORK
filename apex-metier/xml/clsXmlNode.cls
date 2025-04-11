' Migrated to apex-metier/xml - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit
' ==========================================================================
' Class : clsXmlNode
' Version : 1.0
' Purpose : Encapsulation d'un nœud XML avec fonctionnalités étendues
' Date : 10/04/2025
' ==========================================================================

' --- Variables privées ---
Private m_Node As Object       ' MSXML2.IXMLDOMNode
Private m_Logger As Object     ' ILoggerBase
Private m_LastError As String

' --- Initialisation ---
Private Sub Class_Initialize()
    m_LastError = ""
    Set m_Node = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_Node = Nothing
    Set m_Logger = Nothing
End Sub

' --- Propriétés ---
Public Property Get Node() As Object
    ' Retourne le nœud XML sous-jacent
    Set Node = m_Node
End Property

Public Property Set Node(ByVal value As Object)
    ' Définit le nœud XML sous-jacent
    Set m_Node = value
End Property

Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

Public Property Set Logger(ByVal value As Object)
    ' Définit le logger à utiliser
    Set m_Logger = value
End Property

Public Property Get NodeType() As Integer
    ' Retourne le type du nœud
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        NodeType = -1
        Exit Property
    End If
    
    NodeType = m_Node.nodeType
    
    Exit Property
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention du type de nœud: " & Err.Description
    LogError m_LastError
    NodeType = -1
End Property

Public Property Get NodeName() As String
    ' Retourne le nom du nœud
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        NodeName = ""
        Exit Property
    End If
    
    NodeName = m_Node.nodeName
    
    Exit Property
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention du nom de nœud: " & Err.Description
    LogError m_LastError
    NodeName = ""
End Property

Public Property Get NodeValue() As String
    ' Retourne la valeur du nœud
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        NodeValue = ""
        Exit Property
    End If
    
    ' Le type de nœud détermine comment obtenir la valeur
    Select Case m_Node.nodeType
        Case 1 ' NODE_ELEMENT
            NodeValue = m_Node.Text
        Case 2 ' NODE_ATTRIBUTE
            NodeValue = m_Node.value
        Case 3 ' NODE_TEXT
            NodeValue = m_Node.nodeValue
        Case Else
            NodeValue = m_Node.nodeValue
    End Select
    
    Exit Property
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention de la valeur de nœud: " & Err.Description
    LogError m_LastError
    NodeValue = ""
End Property

Public Property Let NodeValue(ByVal value As String)
    ' Définit la valeur du nœud
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour affecter une valeur"
        LogError m_LastError
        Exit Property
    End If
    
    ' Le type de nœud détermine comment définir la valeur
    Select Case m_Node.nodeType
        Case 1 ' NODE_ELEMENT
            m_Node.Text = value
        Case 2 ' NODE_ATTRIBUTE
            m_Node.value = value
        Case 3 ' NODE_TEXT
            m_Node.nodeValue = value
        Case Else
            m_Node.nodeValue = value
    End Select
    
    Exit Property
    
ErrorHandler:
    m_LastError = "Erreur lors de la définition de la valeur de nœud: " & Err.Description
    LogError m_LastError
End Property

Public Property Get ParentNode() As clsXmlNode
    ' Retourne le nœud parent encapsulé dans un objet clsXmlNode
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        Set ParentNode = Nothing
        Exit Property
    End If
    
    If m_Node.parentNode Is Nothing Then
        Set ParentNode = Nothing
        Exit Property
    End If
    
    ' Créer un nouvel objet clsXmlNode pour le parent
    Set ParentNode = New clsXmlNode
    ParentNode.Node = m_Node.parentNode
    If Not m_Logger Is Nothing Then
        Set ParentNode.Logger = m_Logger
    End If
    
    Exit Property
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention du nœud parent: " & Err.Description
    LogError m_LastError
    Set ParentNode = Nothing
End Property

' --- Méthodes publiques ---
Public Function GetChildNodes() As Collection
    ' Retourne une collection de tous les nœuds enfants
    Dim result As Collection
    Dim childNode As Object
    Dim xmlNode As clsXmlNode
    
    On Error GoTo ErrorHandler
    
    Set result = New Collection
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour obtenir les enfants"
        LogError m_LastError
        Set GetChildNodes = result
        Exit Function
    End If
    
    ' Parcourir tous les nœuds enfants
    For Each childNode In m_Node.childNodes
        Set xmlNode = New clsXmlNode
        Set xmlNode.Node = childNode
        If Not m_Logger Is Nothing Then
            Set xmlNode.Logger = m_Logger
        End If
        result.Add xmlNode
    Next childNode
    
    Set GetChildNodes = result
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention des nœuds enfants: " & Err.Description
    LogError m_LastError
    Set GetChildNodes = result
End Function

Public Function GetChildNodeByName(ByVal nodeName As String) As clsXmlNode
    ' Retourne le premier nœud enfant avec le nom spécifié
    Dim childNode As Object
    
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour obtenir l'enfant"
        LogError m_LastError
        Set GetChildNodeByName = Nothing
        Exit Function
    End If
    
    ' Rechercher l'enfant par nom
    For Each childNode In m_Node.childNodes
        If childNode.nodeName = nodeName Then
            Set GetChildNodeByName = New clsXmlNode
            Set GetChildNodeByName.Node = childNode
            If Not m_Logger Is Nothing Then
                Set GetChildNodeByName.Logger = m_Logger
            End If
            Exit Function
        End If
    Next childNode
    
    ' Aucun enfant trouvé avec ce nom
    Set GetChildNodeByName = Nothing
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la recherche de l'enfant par nom: " & Err.Description
    LogError m_LastError
    Set GetChildNodeByName = Nothing
End Function

Public Function GetAttribute(ByVal attributeName As String) As String
    ' Obtient la valeur d'un attribut
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        GetAttribute = ""
        Exit Function
    End If
    
    ' Vérifier que c'est un élément
    If m_Node.nodeType <> 1 Then ' NODE_ELEMENT
        GetAttribute = ""
        Exit Function
    End If
    
    ' Obtenir l'attribut
    If m_Node.hasAttribute(attributeName) Then
        GetAttribute = m_Node.getAttribute(attributeName)
    Else
        GetAttribute = ""
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'obtention de l'attribut: " & Err.Description
    LogError m_LastError
    GetAttribute = ""
End Function

Public Function SetAttribute(ByVal attributeName As String, ByVal attributeValue As String) As Boolean
    ' Définit la valeur d'un attribut
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour définir l'attribut"
        LogError m_LastError
        SetAttribute = False
        Exit Function
    End If
    
    ' Vérifier que c'est un élément
    If m_Node.nodeType <> 1 Then ' NODE_ELEMENT
        m_LastError = "Le nœud n'est pas un élément, impossible de définir un attribut"
        LogError m_LastError
        SetAttribute = False
        Exit Function
    End If
    
    ' Définir l'attribut
    m_Node.setAttribute attributeName, attributeValue
    SetAttribute = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la définition de l'attribut: " & Err.Description
    LogError m_LastError
    SetAttribute = False
End Function

Public Function RemoveAttribute(ByVal attributeName As String) As Boolean
    ' Supprime un attribut
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour supprimer l'attribut"
        LogError m_LastError
        RemoveAttribute = False
        Exit Function
    End If
    
    ' Vérifier que c'est un élément
    If m_Node.nodeType <> 1 Then ' NODE_ELEMENT
        m_LastError = "Le nœud n'est pas un élément, impossible de supprimer un attribut"
        LogError m_LastError
        RemoveAttribute = False
        Exit Function
    End If
    
    ' Vérifier si l'attribut existe
    If Not m_Node.hasAttribute(attributeName) Then
        ' L'attribut n'existe pas, considérer comme réussi
        RemoveAttribute = True
        Exit Function
    End If
    
    ' Supprimer l'attribut
    m_Node.removeAttribute attributeName
    RemoveAttribute = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression de l'attribut: " & Err.Description
    LogError m_LastError
    RemoveAttribute = False
End Function

Public Function HasAttribute(ByVal attributeName As String) As Boolean
    ' Vérifie si un attribut existe
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        HasAttribute = False
        Exit Function
    End If
    
    ' Vérifier que c'est un élément
    If m_Node.nodeType <> 1 Then ' NODE_ELEMENT
        HasAttribute = False
        Exit Function
    End If
    
    ' Vérifier l'attribut
    HasAttribute = m_Node.hasAttribute(attributeName)
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la vérification de l'attribut: " & Err.Description
    LogError m_LastError
    HasAttribute = False
End Function

Public Function AppendChild(ByVal childNode As clsXmlNode) As Boolean
    ' Ajoute un nœud enfant
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud parent défini pour ajouter un enfant"
        LogError m_LastError
        AppendChild = False
        Exit Function
    End If
    
    If childNode Is Nothing Then
        m_LastError = "Aucun nœud enfant fourni à ajouter"
        LogError m_LastError
        AppendChild = False
        Exit Function
    End If
    
    If childNode.Node Is Nothing Then
        m_LastError = "Le nœud enfant fourni ne contient pas de nœud XML"
        LogError m_LastError
        AppendChild = False
        Exit Function
    End If
    
    ' Ajouter l'enfant
    m_Node.appendChild childNode.Node
    AppendChild = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'ajout d'un nœud enfant: " & Err.Description
    LogError m_LastError
    AppendChild = False
End Function

Public Function RemoveChild(ByVal childNode As clsXmlNode) As Boolean
    ' Supprime un nœud enfant
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud parent défini pour supprimer un enfant"
        LogError m_LastError
        RemoveChild = False
        Exit Function
    End If
    
    If childNode Is Nothing Then
        m_LastError = "Aucun nœud enfant fourni à supprimer"
        LogError m_LastError
        RemoveChild = False
        Exit Function
    End If
    
    If childNode.Node Is Nothing Then
        m_LastError = "Le nœud enfant fourni ne contient pas de nœud XML"
        LogError m_LastError
        RemoveChild = False
        Exit Function
    End If
    
    ' Supprimer l'enfant
    m_Node.removeChild childNode.Node
    RemoveChild = True
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la suppression d'un nœud enfant: " & Err.Description
    LogError m_LastError
    RemoveChild = False
End Function

Public Function CloneNode(Optional ByVal deep As Boolean = True) As clsXmlNode
    ' Clone le nœud
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini à cloner"
        LogError m_LastError
        Set CloneNode = Nothing
        Exit Function
    End If
    
    ' Cloner le nœud
    Dim clonedNode As Object
    Set clonedNode = m_Node.cloneNode(deep)
    
    ' Créer un nouvel objet clsXmlNode pour le clone
    Set CloneNode = New clsXmlNode
    Set CloneNode.Node = clonedNode
    If Not m_Logger Is Nothing Then
        Set CloneNode.Logger = m_Logger
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors du clonage d'un nœud: " & Err.Description
    LogError m_LastError
    Set CloneNode = Nothing
End Function

Public Function SelectSingleNode(ByVal xPath As String) As clsXmlNode
    ' Sélectionne un nœud via XPath
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour sélectionner un nœud"
        LogError m_LastError
        Set SelectSingleNode = Nothing
        Exit Function
    End If
    
    ' Sélectionner le nœud
    Dim selectedNode As Object
    Set selectedNode = m_Node.SelectSingleNode(xPath)
    
    If selectedNode Is Nothing Then
        Set SelectSingleNode = Nothing
        Exit Function
    End If
    
    ' Créer un nouvel objet clsXmlNode pour le nœud sélectionné
    Set SelectSingleNode = New clsXmlNode
    Set SelectSingleNode.Node = selectedNode
    If Not m_Logger Is Nothing Then
        Set SelectSingleNode.Logger = m_Logger
    End If
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la sélection d'un nœud par XPath: " & Err.Description
    LogError m_LastError
    Set SelectSingleNode = Nothing
End Function

Public Function SelectNodes(ByVal xPath As String) As Collection
    ' Sélectionne plusieurs nœuds via XPath
    Dim result As Collection
    Dim nodeList As Object
    Dim node As Object
    Dim xmlNode As clsXmlNode
    
    On Error GoTo ErrorHandler
    
    Set result = New Collection
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour sélectionner des nœuds"
        LogError m_LastError
        Set SelectNodes = result
        Exit Function
    End If
    
    ' Sélectionner les nœuds
    Set nodeList = m_Node.SelectNodes(xPath)
    
    If nodeList Is Nothing Then
        Set SelectNodes = result
        Exit Function
    End If
    
    ' Parcourir tous les nœuds sélectionnés
    For Each node In nodeList
        Set xmlNode = New clsXmlNode
        Set xmlNode.Node = node
        If Not m_Logger Is Nothing Then
            Set xmlNode.Logger = m_Logger
        End If
        result.Add xmlNode
    Next node
    
    Set SelectNodes = result
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la sélection de nœuds par XPath: " & Err.Description
    LogError m_LastError
    Set SelectNodes = result
End Function

Public Function ToXmlString() As String
    ' Convertit le nœud en chaîne XML
    On Error GoTo ErrorHandler
    
    If m_Node Is Nothing Then
        m_LastError = "Aucun nœud défini pour la conversion en chaîne"
        LogError m_LastError
        ToXmlString = ""
        Exit Function
    End If
    
    ToXmlString = m_Node.xml
    
    Exit Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la conversion en chaîne XML: " & Err.Description
    LogError m_LastError
    ToXmlString = ""
End Function

' --- Méthodes privées ---
Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "XML"
    End If
End Sub
