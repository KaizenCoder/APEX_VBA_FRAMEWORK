' Migrated to apex-metier/orm/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Interface : IRelationMetadata
' Version : 1.0
' Purpose : Interface pour décrire une relation entre entités ORM
' ==========================================================================

' --- Propriétés de base de la relation ---
Public Property Get RelationName() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get RelationType() As RelationTypeEnum: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get SourceEntity() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get TargetEntity() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property

' --- Colonnes et clés ---
Public Property Get LocalKeyColumn() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get ForeignKeyColumn() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get JunctionTable() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get JunctionSourceColumn() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get JunctionTargetColumn() As String: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property

' --- Comportement ---
Public Property Get LazyLoading() As Boolean: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get CascadeDelete() As Boolean: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property
Public Property Get CascadeSave() As Boolean: Err.Raise vbObjectError + 1502, "IRelationMetadata": End Property 