' Migrated to apex-metier/orm - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Class : clsRelationMetadata
' Version : 1.0
' Implements: IRelationMetadata (v1.0)
' Purpose : Implémentation des métadonnées de relation pour l'ORM
' Requires : RelationTypeEnum, IRelationMetadata
' ==========================================================================

Implements IRelationMetadata

' --- Variables membres ---
Private m_RelationName As String          ' Nom unique de la relation
Private m_RelationType As RelationTypeEnum ' Type de relation (1:1, 1:N, N:M, etc.)
Private m_SourceEntity As String          ' Entité source (nom de la classe/table)
Private m_TargetEntity As String          ' Entité cible (nom de la classe/table)
Private m_LocalKeyColumn As String        ' Colonne clé locale (dans l'entité source)
Private m_ForeignKeyColumn As String      ' Colonne clé étrangère (dans l'entité cible)
Private m_JunctionTable As String         ' Table de liaison (pour N:M)
Private m_JunctionSourceColumn As String  ' Colonne source dans la table de liaison
Private m_JunctionTargetColumn As String  ' Colonne cible dans la table de liaison
Private m_LazyLoading As Boolean          ' Chargement différé
Private m_CascadeDelete As Boolean        ' Suppression en cascade
Private m_CascadeSave As Boolean          ' Sauvegarde en cascade

' --- Initialisation ---
Private Sub Class_Initialize()
    ' Valeurs par défaut
    m_RelationName = ""
    m_RelationType = RelationTypeEnum.OneToMany ' Par défaut
    m_SourceEntity = ""
    m_TargetEntity = ""
    m_LocalKeyColumn = "ID"
    m_ForeignKeyColumn = ""
    m_JunctionTable = ""
    m_JunctionSourceColumn = ""
    m_JunctionTargetColumn = ""
    m_LazyLoading = True      ' Activé par défaut
    m_CascadeDelete = False   ' Désactivé par défaut
    m_CascadeSave = True      ' Activé par défaut
End Sub

' --- Constructeur pour relation 1:1 ou 1:N ---
'@Description: 
'@Param: 
'@Returns: 

Public Function CreateSimpleRelation(ByVal relationName As String, ByVal relationType As RelationTypeEnum, _
                                  ByVal sourceEntity As String, ByVal targetEntity As String, _
                                  ByVal localKeyColumn As String, ByVal foreignKeyColumn As String) As clsRelationMetadata
    
    ' Définir les propriétés de base
    m_RelationName = relationName
    m_RelationType = relationType
    m_SourceEntity = sourceEntity
    m_TargetEntity = targetEntity
    m_LocalKeyColumn = localKeyColumn
    m_ForeignKeyColumn = foreignKeyColumn
    
    ' Retourner l'instance configurée
    Set CreateSimpleRelation = Me
End Function

' --- Constructeur pour relation N:M ---
'@Description: 
'@Param: 
'@Returns: 

Public Function CreateManyToManyRelation(ByVal relationName As String, _
                                      ByVal sourceEntity As String, ByVal targetEntity As String, _
                                      ByVal junctionTable As String, _
                                      ByVal junctionSourceColumn As String, ByVal junctionTargetColumn As String) As clsRelationMetadata
    
    ' Définir les propriétés de base
    m_RelationName = relationName
    m_RelationType = RelationTypeEnum.ManyToMany
    m_SourceEntity = sourceEntity
    m_TargetEntity = targetEntity
    m_JunctionTable = junctionTable
    m_JunctionSourceColumn = junctionSourceColumn
    m_JunctionTargetColumn = junctionTargetColumn
    
    ' Retourner l'instance configurée
    Set CreateManyToManyRelation = Me
End Function

' --- Configuration avancée ---
'@Description: 
'@Param: 
'@Returns: 

Public Function WithLazyLoading(ByVal enabled As Boolean) As clsRelationMetadata
    m_LazyLoading = enabled
    Set WithLazyLoading = Me
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function WithCascadeDelete(ByVal enabled As Boolean) As clsRelationMetadata
    m_CascadeDelete = enabled
    Set WithCascadeDelete = Me
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function WithCascadeSave(ByVal enabled As Boolean) As clsRelationMetadata
    m_CascadeSave = enabled
    Set WithCascadeSave = Me
End Function

' --- Implémentation IRelationMetadata ---
Private Property Get IRelationMetadata_RelationName() As String
    IRelationMetadata_RelationName = m_RelationName
End Property

Private Property Get IRelationMetadata_RelationType() As RelationTypeEnum
    IRelationMetadata_RelationType = m_RelationType
End Property

Private Property Get IRelationMetadata_SourceEntity() As String
    IRelationMetadata_SourceEntity = m_SourceEntity
End Property

Private Property Get IRelationMetadata_TargetEntity() As String
    IRelationMetadata_TargetEntity = m_TargetEntity
End Property

Private Property Get IRelationMetadata_LocalKeyColumn() As String
    IRelationMetadata_LocalKeyColumn = m_LocalKeyColumn
End Property

Private Property Get IRelationMetadata_ForeignKeyColumn() As String
    IRelationMetadata_ForeignKeyColumn = m_ForeignKeyColumn
End Property

Private Property Get IRelationMetadata_JunctionTable() As String
    IRelationMetadata_JunctionTable = m_JunctionTable
End Property

Private Property Get IRelationMetadata_JunctionSourceColumn() As String
    IRelationMetadata_JunctionSourceColumn = m_JunctionSourceColumn
End Property

Private Property Get IRelationMetadata_JunctionTargetColumn() As String
    IRelationMetadata_JunctionTargetColumn = m_JunctionTargetColumn
End Property

Private Property Get IRelationMetadata_LazyLoading() As Boolean
    IRelationMetadata_LazyLoading = m_LazyLoading
End Property

Private Property Get IRelationMetadata_CascadeDelete() As Boolean
    IRelationMetadata_CascadeDelete = m_CascadeDelete
End Property

Private Property Get IRelationMetadata_CascadeSave() As Boolean
    IRelationMetadata_CascadeSave = m_CascadeSave
End Property

' --- Accesseurs publics ---
Public Property Get RelationName() As String
    RelationName = m_RelationName
End Property

Public Property Get RelationType() As RelationTypeEnum
    RelationType = m_RelationType
End Property

Public Property Get SourceEntity() As String
    SourceEntity = m_SourceEntity
End Property

Public Property Get TargetEntity() As String
    TargetEntity = m_TargetEntity
End Property

Public Property Get LocalKeyColumn() As String
    LocalKeyColumn = m_LocalKeyColumn
End Property

Public Property Get ForeignKeyColumn() As String
    ForeignKeyColumn = m_ForeignKeyColumn
End Property

Public Property Get JunctionTable() As String
    JunctionTable = m_JunctionTable
End Property

Public Property Get JunctionSourceColumn() As String
    JunctionSourceColumn = m_JunctionSourceColumn
End Property

Public Property Get JunctionTargetColumn() As String
    JunctionTargetColumn = m_JunctionTargetColumn
End Property

Public Property Get LazyLoading() As Boolean
    LazyLoading = m_LazyLoading
End Property

Public Property Get CascadeDelete() As Boolean
    CascadeDelete = m_CascadeDelete
End Property

Public Property Get CascadeSave() As Boolean
    CascadeSave = m_CascadeSave
End Property 