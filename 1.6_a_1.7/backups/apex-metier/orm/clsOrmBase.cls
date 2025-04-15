' Migrated to apex-metier/orm - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Class : clsOrmBase
' Version : 2.0 - Phase 4 - Relations and Composite Keys
' Implements: IRelationalObject (v1.0)
' Purpose : Base class for ORM objects with relations support
' Requires : clsDbAccessor, clsQueryBuilder, IRelationMetadata
' ==========================================================================

Implements IRelationalObject

' --- Variables membres ---
Private m_DbAccessor As clsDbAccessor        ' Accès à la base de données
Private m_TableName As String                ' Nom de la table
Private m_PrimaryKeyColumns() As String      ' Colonnes de clé primaire (composite)
Private m_IsLoaded As Boolean                ' État chargé
Private m_IsDirty As Boolean                 ' État modifié
Private m_LastError As String                ' Dernière erreur
Private m_Relations As Object                ' Dictionary<String, Collection<IRelationMetadata>>
Private m_RelatedObjects As Object           ' Dictionary<String, Collection<Object>>
Private m_Logger As ILoggerBase              ' Logger d'instance

' --- Constantes ---
Private Const MAX_FETCH_DEPTH As Integer = 3  ' Profondeur maximale de chargement pour éviter les boucles
Private Const ORM_TAG As String = "ORM"       ' Tag pour le logging

' --- Initialisation ---
Private Sub Class_Initialize()
    ' Initialiser les dictionnaires pour les relations
    Set m_Relations = CreateObject("Scripting.Dictionary")
    Set m_RelatedObjects = CreateObject("Scripting.Dictionary")
    
    ' Valeurs par défaut
    m_IsLoaded = False
    m_IsDirty = False
    m_LastError = ""
    
    ' Par défaut, une seule clé primaire - à surcharger dans les classes dérivées
    ReDim m_PrimaryKeyColumns(0)
    m_PrimaryKeyColumns(0) = "ID"
End Sub

' --- Méthodes protégées à surcharger ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize(ByVal dbAccessor As clsDbAccessor, Optional ByVal tableName As String = "", Optional ByVal logger As ILoggerBase = Nothing)
    ' Initialiser les références et propriétés de base
    Set m_DbAccessor = dbAccessor
    Set m_Logger = logger
    
    If tableName <> "" Then m_TableName = tableName
    
    LogIfAble "Initialize", LogLevelDebug, "ORM initialisé pour table: " & m_TableName
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Public Property Let TableName(ByVal value As String)
    m_TableName = value
End Property

Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Get PrimaryKeyColumns() As String()
    PrimaryKeyColumns = m_PrimaryKeyColumns
End Property

'@Description: 
'@Param: 
'@Returns: 

Public Sub SetPrimaryKeys(ParamArray keyColumns() As Variant)
    Dim i As Long
    
    ' Redimensionner le tableau pour correspondre au nombre de clés
    If UBound(keyColumns) >= 0 Then
        ReDim m_PrimaryKeyColumns(UBound(keyColumns))
        
        ' Copier les noms des colonnes
        For i = 0 To UBound(keyColumns)
            m_PrimaryKeyColumns(i) = CStr(keyColumns(i))
        Next i
        
        LogIfAble "SetPrimaryKeys", LogLevelDebug, "Clé(s) primaire(s) définie(s): " & Join(m_PrimaryKeyColumns, ", ")
    Else
        ReDim m_PrimaryKeyColumns(0)
        m_PrimaryKeyColumns(0) = "ID"
        LogIfAble "SetPrimaryKeys", LogLevelDebug, "Aucune clé définie, utilisation de ID par défaut"
    End If
End Sub

' --- Propriétés d'état ---
Public Property Get IsLoaded() As Boolean
    IsLoaded = m_IsLoaded
End Property

Public Property Let IsLoaded(ByVal value As Boolean)
    m_IsLoaded = value
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = m_IsDirty
End Property

Public Property Let IsDirty(ByVal value As Boolean)
    m_IsDirty = value
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

' --- CRUD Operations ---
'@Description: 
'@Param: 
'@Returns: 

Public Function Load(ByVal id As Variant) As Boolean
    ' Cette méthode devrait être surchargée dans les classes dérivées
    ' Vérifier les préconditions
    If m_DbAccessor Is Nothing Then
        m_LastError = "Erreur: DbAccessor n'est pas initialisé"
        LogIfAble "Load", LogLevelError, m_LastError
        Load = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    If m_TableName = "" Then
        m_LastError = "Erreur: Nom de table non défini"
        LogIfAble "Load", LogLevelError, m_LastError
        Load = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Pour les clés composites, cette méthode accepte un variant mais n'implémente que le cas simple
    ' Pour une implémentation complète, surcharger dans la classe dérivée
    
    ' Préparer la requête
    Dim qb As New clsQueryBuilder
    
    qb.SelectColumns("*") _
      .FromTable(m_TableName) _
      .AddWhere(m_PrimaryKeyColumns(0), "=", id)
      
    Dim result As Variant
    result = qb.Build()
    
    ' Exécuter la requête
    Dim rs As ADODB.Recordset
    Set rs = m_DbAccessor.GetRecordset(result(0), result(1))
    
    ' Traiter les résultats
    If Not rs.EOF Then
        ' LoadFromRecordset devrait être surchargée dans les classes dérivées
        Load = LoadFromRecordset(rs)
        m_IsLoaded = Load
        m_IsDirty = False
    Else
        m_LastError = "Aucun enregistrement trouvé avec ID = " & id
        Load = False
    End If
    
    ' Nettoyer
    rs.Close
    Set rs = Nothing
    Set qb = Nothing
    
    LogIfAble "Load", IIf(Load, LogLevelInfo, LogLevelWarning), _
              IIf(Load, "Chargé avec succès ID " & id, "Échec du chargement ID " & id & ": " & m_LastError)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
LoadError:
    m_LastError = "Erreur lors du chargement: " & Err.Description
    LogIfAble "Load", LogLevelError, m_LastError
    Load = False
    Resume LoadCleanup
    
LoadCleanup:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set qb = Nothing
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function LoadFromRecordset(rs As ADODB.Recordset) As Boolean
    ' Cette méthode doit être surchargée dans les classes dérivées
    ' pour mapper les champs du recordset aux propriétés de l'objet
    
    ' Implémentation de base
    LogIfAble "LoadFromRecordset", LogLevelWarning, "Méthode appelée sans être surchargée"
    LoadFromRecordset = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function Save() As Boolean
    ' Cette méthode devrait être surchargée dans les classes dérivées pour une implémentation complète
    ' Vérifier les préconditions
    If m_DbAccessor Is Nothing Then
        m_LastError = "Erreur: DbAccessor n'est pas initialisé"
        LogIfAble "Save", LogLevelError, m_LastError
        Save = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    If m_TableName = "" Then
        m_LastError = "Erreur: Nom de table non défini"
        LogIfAble "Save", LogLevelError, m_LastError
        Save = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si l'objet est déjà chargé
    If m_IsLoaded Then
        ' Update (à implémenter par surcharge)
        Save = False
        m_LastError = "Update not implemented in base class"
    Else
        ' Insert (à implémenter par surcharge)
        Save = False
        m_LastError = "Insert not implemented in base class"
    End If
    
    LogIfAble "Save", IIf(Save, LogLevelInfo, LogLevelWarning), _
              IIf(Save, "Sauvegardé avec succès", "Échec de la sauvegarde: " & m_LastError)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
SaveError:
    m_LastError = "Erreur lors de la sauvegarde: " & Err.Description
    LogIfAble "Save", LogLevelError, m_LastError
    Save = False
    Resume Next
End Function

' --- Nouvelles méthodes pour le support des clés primaires (Phase 4) ---
'@Description: 
'@Param: 
'@Returns: 

Public Function GetPrimaryKeyValue() As Variant
    ' Retourne la valeur de la clé primaire simple
    ' Cette méthode doit être surchargée pour les clés composites
    If UBound(m_PrimaryKeyColumns) = 0 Then
        ' À surcharger pour retourner la propriété correspondante
        GetPrimaryKeyValue = Null
    Else
        GetPrimaryKeyValue = Null
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetPrimaryKeyValues() As Variant()
    ' Retourne les valeurs de la clé primaire composite
    ' Cette méthode doit être surchargée par la classe dérivée
    ReDim result(UBound(m_PrimaryKeyColumns))
    Dim i As Long
    
    For i = 0 To UBound(m_PrimaryKeyColumns)
        ' À surcharger pour retourner les propriétés correspondantes
        result(i) = Null
    Next i
    
    GetPrimaryKeyValues = result
End Function

' --- Gestion des relations ---
'@Description: 
'@Param: 
'@Returns: 

Public Function AddRelation(ByVal relationMetadata As IRelationMetadata) As Boolean
    Dim relationName As String
    relationName = relationMetadata.RelationName
    
    ' Vérifier si la relation existe déjà
    If Not m_Relations.Exists(relationName) Then
        ' Créer une nouvelle collection
        Dim relationCollection As New Collection
        relationCollection.Add relationMetadata
        
        ' Ajouter au dictionnaire
        m_Relations.Add relationName, relationCollection
    Else
        ' Ajouter à la collection existante
        Dim existingCollection As Collection
        Set existingCollection = m_Relations(relationName)
        existingCollection.Add relationMetadata
    End If
    
    ' Initialiser la collection d'objets si nécessaire
    If Not m_RelatedObjects.Exists(relationName) Then
        Set m_RelatedObjects(relationName) = New Collection
    End If
    
    LogIfAble "AddRelation", LogLevelDebug, "Relation ajoutée: " & relationName & " (" & _
              RelationTypeToString(relationMetadata.RelationType) & ") -> " & relationMetadata.TargetEntity
    
    AddRelation = True
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function RelationTypeToString(ByVal relationType As RelationTypeEnum) As String
    Select Case relationType
        Case RelationTypeEnum.OneToOne: RelationTypeToString = "1:1"
        Case RelationTypeEnum.OneToMany: RelationTypeToString = "1:N"
        Case RelationTypeEnum.ManyToOne: RelationTypeToString = "N:1"
        Case RelationTypeEnum.ManyToMany: RelationTypeToString = "N:M"
        Case Else: RelationTypeToString = "Unknown"
    End Select
End Function

' --- IRelationalObject Implementation ---
'@Description: 
'@Param: 
'@Returns: 

Private Function IRelationalObject_GetRelatedObjects(ByVal relationName As String) As Collection
    Dim result As New Collection
    
    ' Vérifier si la relation existe
    If Not m_Relations.Exists(relationName) Then
        LogIfAble "GetRelatedObjects", LogLevelWarning, "Relation non trouvée: " & relationName
        Set IRelationalObject_GetRelatedObjects = result
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si des objets sont déjà chargés
    If m_RelatedObjects.Exists(relationName) Then
        Set result = m_RelatedObjects(relationName)
    Else
        ' Si pas d'objets chargés mais relation définie, charger via lazy loading
        Dim metadata As IRelationMetadata
        Set metadata = m_Relations(relationName)(1) ' Prendre la première définition
        
        If metadata.LazyLoading Then
            LoadRelatedObjects relationName
            If m_RelatedObjects.Exists(relationName) Then
                Set result = m_RelatedObjects(relationName)
            End If
        End If
    End If
    
    Set IRelationalObject_GetRelatedObjects = result
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IRelationalObject_GetRelatedObject(ByVal relationName As String, Optional ByVal relatedId As Variant) As Object
    Set IRelationalObject_GetRelatedObject = Nothing
    
    ' Vérifier si la relation existe
    If Not m_Relations.Exists(relationName) Then
        LogIfAble "GetRelatedObject", LogLevelWarning, "Relation non trouvée: " & relationName
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Vérifier si des objets sont déjà chargés
    Dim objects As Collection
    
    If m_RelatedObjects.Exists(relationName) Then
        Set objects = m_RelatedObjects(relationName)
    Else
        ' Si pas d'objets chargés mais relation définie, charger via lazy loading
        Dim metadata As IRelationMetadata
        Set metadata = m_Relations(relationName)(1) ' Prendre la première définition
        
        If metadata.LazyLoading Then
            LoadRelatedObjects relationName
            If m_RelatedObjects.Exists(relationName) Then
                Set objects = m_RelatedObjects(relationName)
            Else
                Exit'@Description: 
'@Param: 
'@Returns: 

 Function
            End If
        Else
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    End If
    
    ' Si ID spécifié, chercher l'objet correspondant
    If Not IsMissing(relatedId) And Not IsEmpty(relatedId) Then
        Dim obj As Object
        For Each obj In objects
            ' Vérifier si l'objet a une méthode GetPrimaryKeyValue
            On Error Resume Next
            If obj.GetPrimaryKeyValue = relatedId Then
                Set IRelationalObject_GetRelatedObject = obj
                Exit'@Description: 
'@Param: 
'@Returns: 

 Function
            End If
            On Error GoTo 0
        Next obj
    ElseIf objects.Count > 0 Then
        ' Sinon retourner le premier objet
        Set IRelationalObject_GetRelatedObject = objects(1)
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IRelationalObject_AddRelatedObject(ByVal relationName As String, ByVal relatedObject As Object) As Boolean
    ' Vérifier si la relation existe
    If Not m_Relations.Exists(relationName) Then
        LogIfAble "AddRelatedObject", LogLevelWarning, "Relation non trouvée: " & relationName
        IRelationalObject_AddRelatedObject = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Initialiser la collection si nécessaire
    If Not m_RelatedObjects.Exists(relationName) Then
        Set m_RelatedObjects(relationName) = New Collection
    End If
    
    ' Ajouter l'objet
    On Error Resume Next
    m_RelatedObjects(relationName).Add relatedObject
    
    If Err.Number <> 0 Then
        LogIfAble "AddRelatedObject", LogLevelError, "Erreur lors de l'ajout: " & Err.Description
        IRelationalObject_AddRelatedObject = False
        Err.Clear
    Else
        IRelationalObject_AddRelatedObject = True
        m_IsDirty = True
    End If
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IRelationalObject_RemoveRelatedObject(ByVal relationName As String, ByVal relatedObject As Object) As Boolean
    ' Vérifier si la relation existe
    If Not m_Relations.Exists(relationName) Or Not m_RelatedObjects.Exists(relationName) Then
        LogIfAble "RemoveRelatedObject", LogLevelWarning, "Relation non trouvée: " & relationName
        IRelationalObject_RemoveRelatedObject = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Trouver et supprimer l'objet
    Dim objects As Collection
    Set objects = m_RelatedObjects(relationName)
    
    Dim i As Long
    Dim found As Boolean
    
    ' Recherche par référence d'objet
    For i = 1 To objects.Count
        If objects(i) Is relatedObject Then
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        On Error Resume Next
        objects.Remove i
        
        If Err.Number <> 0 Then
            LogIfAble "RemoveRelatedObject", LogLevelError, "Erreur lors de la suppression: " & Err.Description
            IRelationalObject_RemoveRelatedObject = False
            Err.Clear
        Else
            IRelationalObject_RemoveRelatedObject = True
            m_IsDirty = True
        End If
        On Error GoTo 0
    Else
        IRelationalObject_RemoveRelatedObject = False
        LogIfAble "RemoveRelatedObject", LogLevelWarning, "Objet non trouvé dans la relation"
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IRelationalObject_HasRelatedObjects(ByVal relationName As String) As Boolean
    ' Vérifier si la relation existe et contient des objets
    If Not m_Relations.Exists(relationName) Then
        IRelationalObject_HasRelatedObjects = False
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    If m_RelatedObjects.Exists(relationName) Then
        IRelationalObject_HasRelatedObjects = (m_RelatedObjects(relationName).Count > 0)
    Else
        IRelationalObject_HasRelatedObjects = False
    End If
End Function

' --- Chargement et sauvegarde en cascade ---
'@Description: 
'@Param: 
'@Returns: 

Private Function IRelationalObject_LoadWithRelations(Optional ByVal depth As Integer = 1) As Boolean
    ' Cette méthode doit d'abord charger l'objet lui-même
    ' puis charger ses relations jusqu'à la profondeur spécifiée
    
    ' Vérifier la profondeur maximale pour éviter les boucles infinies
    If depth > MAX_FETCH_DEPTH Then depth = MAX_FETCH_DEPTH
    
    ' Pour l'instant simple validation - la classe dérivée doit implémenter sa propre logique
    IRelationalObject_LoadWithRelations = (m_IsLoaded Or depth <= 0)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IRelationalObject_SaveWithRelations(Optional ByVal cascadingMode As Integer = 1) As Boolean
    ' Cette méthode doit d'abord sauvegarder l'objet lui-même
    ' puis sauvegarder ses relations selon le mode spécifié
    ' cascadingMode: 1 = save all, 2 = save new only, 3 = save modified only
    
    ' Pour l'instant simple validation - la classe dérivée doit implémenter sa propre logique
    IRelationalObject_SaveWithRelations = False
End Function

' --- Méthode pour chargement différé des relations ---
'@Description: 
'@Param: 
'@Returns: 

Private Function LoadRelatedObjects(ByVal relationName As String) As Boolean
    ' Cette méthode devrait être surchargée dans les classes dérivées
    ' Elle doit charger les objets liés à la relation spécifiée
    
    LogIfAble "LoadRelatedObjects", LogLevelWarning, "Méthode non implémentée dans la classe de base"
    LoadRelatedObjects = False
End Function

' --- Utilitaires ---
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogIfAble(ByVal method As String, ByVal level As LogLevelEnum, ByVal message As String)
    If m_Logger Is Nothing Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
    On Error Resume Next
    m_Logger.LogMessage message, level, ORM_TAG, TypeName(Me) & "." & method
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Class_Terminate()
    ' Nettoyage des collections et références
    Set m_Relations = Nothing
    Set m_RelatedObjects = Nothing
    Set m_DbAccessor = Nothing
    Set m_Logger = Nothing
End Sub 