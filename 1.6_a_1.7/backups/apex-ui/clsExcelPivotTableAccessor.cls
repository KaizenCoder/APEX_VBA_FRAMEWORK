VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelPivotTableAccessor"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.UI.Excel")
'@ModuleDescription("Implémentation de l'interface IPivotTableAccessor pour manipuler les tableaux croisés dynamiques Excel")
'@Implements IPivotTableAccessor
Option Explicit

Implements IPivotTableAccessor

' Variables privées
Private mobjSheet As Object ' ISheetAccessor
Private mobjNativePivot As Object ' PivotTable
Private mstrPivotTableName As String

' Constantes d'erreur
Private Const ERR_INVALID_PIVOT As Long = 2001
Private Const ERR_INVALID_FIELD As Long = 2002
Private Const ERR_INVALID_ITEM As Long = 2003
Private Const ERR_INVALID_OPERATION As Long = 2004

' Constantes pour les zones du tableau croisé
Private Const AREA_ROWS As Long = 1
Private Const AREA_COLUMNS As Long = 2
Private Const AREA_PAGES As Long = 3
Private Const AREA_DATA As Long = 4

'*******************************************************************************
' Module : clsExcelPivotTableAccessor
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Implémentation de l'interface IPivotTableAccessor pour manipuler les
'          tableaux croisés dynamiques Excel
'*******************************************************************************

''
' Initialise un nouvel accesseur de tableau croisé dynamique
' @param sheetAccessor L'accesseur de feuille qui contient le tableau croisé
' @param pivotTableName Le nom du tableau croisé dynamique à manipuler
' @return Boolean True si l'initialisation a réussi
''
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal sheetAccessor As Object, ByVal pivotTableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier et stocker l'accesseur de feuille
    Set mobjSheet = sheetAccessor
    
    ' Récupérer la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Rechercher le tableau croisé dynamique par son nom
    Dim pvt As Object
    For Each pvt In nativeSheet.PivotTables
        If pvt.Name = pivotTableName Then
            Set mobjNativePivot = pvt
            mstrPivotTableName = pivotTableName
            Initialize = True
            Exit Function
        End If
    Next pvt
    
    ' Tableau croisé non trouvé
    Err.Raise ERR_INVALID_PIVOT, "clsExcelPivotTableAccessor", "Le tableau croisé '" & pivotTableName & "' n'existe pas dans la feuille"
    
    Exit Function
ErrorHandler:
    Initialize = False
    If Err.Number <> ERR_INVALID_PIVOT Then
        Err.Raise Err.Number, "clsExcelPivotTableAccessor.Initialize", Err.Description
    End If
End Function

''
' Crée un nouveau tableau croisé dynamique à partir d'une source de données
' @param sheetAccessor L'accesseur de la feuille cible
' @param dataSource La source de données (plage ou tableau)
' @param targetRange La plage cible pour le nouveau tableau croisé
' @param pivotName Le nom à donner au nouveau tableau croisé
' @return clsExcelPivotTableAccessor L'accesseur pour le tableau croisé nouvellement créé
''
'@Description: 
'@Param: 
'@Returns: 

Public Function CreatePivotTableFromData(ByVal sheetAccessor As Object, ByVal dataSource As Object, _
                                      ByVal targetRange As String, ByVal pivotName As String) As clsExcelPivotTableAccessor
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native cible
    Dim nativeTargetSheet As Object
    Set nativeTargetSheet = sheetAccessor.GetNativeSheet
    
    ' Obtenir la plage cible
    Dim destRange As Object
    Set destRange = nativeTargetSheet.Range(targetRange)
    
    ' Créer le tableau croisé à partir de la source
    Dim sourceType As XlPivotTableSourceType
    Dim sourceData As Object
    
    ' Déterminer le type de source
    If TypeOf dataSource Is Object Then
        ' Si c'est un tableau Excel
        If TypeName(dataSource) = "ListObject" Then
            sourceType = xlDatabase
            Set sourceData = dataSource
        ' Si c'est un accesseur de tableau
        ElseIf TypeOf dataSource Is clsExcelTableAccessor Then
            sourceType = xlDatabase
            Set sourceData = dataSource.GetNativeTable
        ' Sinon c'est probablement une plage
        Else
            sourceType = xlDatabase
            Set sourceData = dataSource
        End If
    ElseIf VarType(dataSource) = vbString Then
        ' Si c'est une référence de plage sous forme de chaîne
        sourceType = xlDatabase
        Set sourceData = sheetAccessor.GetNativeSheet.Range(CStr(dataSource))
    End If
    
    ' Créer le tableau croisé
    Dim pivotCache As Object
    Set pivotCache = sheetAccessor.GetNativeSheet.Parent.PivotCaches.Create( _
                        SourceType:=sourceType, _
                        SourceData:=sourceData)
                        
    Dim newPivotTable As Object
    Set newPivotTable = pivotCache.CreatePivotTable( _
                        TableDestination:=destRange, _
                        TableName:=pivotName)
    
    ' Créer et initialiser un nouvel accesseur pour ce tableau croisé
    Dim pivotAccessor As New clsExcelPivotTableAccessor
    pivotAccessor.Initialize sheetAccessor, pivotName
    
    Set CreatePivotTableFromData = pivotAccessor
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelPivotTableAccessor.CreatePivotTableFromData", Err.Description
End Function

' ==================== Implémentation de IPivotTableAccessor ====================

' ----- Propriétés -----

Private Property Get IPivotTableAccessor_PivotTableName() As String
    IPivotTableAccessor_PivotTableName = mstrPivotTableName
End Property

Private Property Get IPivotTableAccessor_ParentSheet() As Object
    Set IPivotTableAccessor_ParentSheet = mobjSheet
End Property

Private Property Get IPivotTableAccessor_HasPageFields() As Boolean
    On Error Resume Next
    IPivotTableAccessor_HasPageFields = (mobjNativePivot.PageFields.Count > 0)
    If Err.Number <> 0 Then IPivotTableAccessor_HasPageFields = False
End Property

Private Property Get IPivotTableAccessor_DataFieldsCount() As Long
    On Error Resume Next
    IPivotTableAccessor_DataFieldsCount = mobjNativePivot.DataFields.Count
    If Err.Number <> 0 Then IPivotTableAccessor_DataFieldsCount = 0
End Property

Private Property Get IPivotTableAccessor_FieldNames() As Variant
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim fields() As String
    Dim fieldCount As Long
    
    fieldCount = mobjNativePivot.PivotFields.Count
    ReDim fields(1 To fieldCount)
    
    For i = 1 To fieldCount
        fields(i) = mobjNativePivot.PivotFields(i).Name
    Next i
    
    IPivotTableAccessor_FieldNames = fields
    Exit Property
    
ErrorHandler:
    Dim emptyArray() As String
    ReDim emptyArray(0)
    IPivotTableAccessor_FieldNames = emptyArray
End Property

' ----- Méthodes de structure -----

'@Description: 
'@Param: 
'@Returns: 

Private Function IPivotTableAccessor_AddRowField(ByVal fieldName As String, Optional ByVal position As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.AddRowField", "Champ non trouvé: " & fieldName
    End If
    
    ' Ajouter comme champ de ligne
    If position >= 0 Then
        field.Orientation = xlRowField
        field.Position = position
    Else
        field.Orientation = xlRowField
    End If
    
    IPivotTableAccessor_AddRowField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_AddRowField = False
End Function

Private Function IPivotTableAccessor_AddColumnField(ByVal fieldName As String, Optional ByVal position As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.AddColumnField", "Champ non trouvé: " & fieldName
    End If
    
    ' Ajouter comme champ de colonne
    If position >= 0 Then
        field.Orientation = xlColumnField
        field.Position = position
    Else
        field.Orientation = xlColumnField
    End If
    
    IPivotTableAccessor_AddColumnField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_AddColumnField = False
End Function

Private Function IPivotTableAccessor_AddPageField(ByVal fieldName As String, Optional ByVal position As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.AddPageField", "Champ non trouvé: " & fieldName
    End If
    
    ' Ajouter comme champ de page (filtre de rapport)
    If position >= 0 Then
        field.Orientation = xlPageField
        field.Position = position
    Else
        field.Orientation = xlPageField
    End If
    
    IPivotTableAccessor_AddPageField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_AddPageField = False
End Function

Private Function IPivotTableAccessor_AddDataField(ByVal fieldName As String, Optional ByVal caption As String = "", _
                                               Optional ByVal function As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.AddDataField", "Champ non trouvé: " & fieldName
    End If
    
    ' Définir la fonction d'agrégation par défaut si non spécifiée
    If function = -1 Then
        function = xlSum
    End If
    
    ' Ajouter comme champ de données
    Dim dataField As Object
    Set dataField = mobjNativePivot.AddDataField(field, caption, function)
    
    ' Si une légende personnalisée a été fournie, l'appliquer
    If caption <> "" Then
        dataField.Caption = caption
    End If
    
    IPivotTableAccessor_AddDataField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_AddDataField = False
End Function

Private Function IPivotTableAccessor_RemoveField(ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe et est actif
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.RemoveField", "Champ non trouvé: " & fieldName
    End If
    
    ' Réinitialiser l'orientation pour retirer le champ
    field.Orientation = xlHidden
    
    IPivotTableAccessor_RemoveField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_RemoveField = False
End Function

Private Function IPivotTableAccessor_MoveField(ByVal fieldName As String, ByVal targetArea As Long, _
                                            Optional ByVal position As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.MoveField", "Champ non trouvé: " & fieldName
    End If
    
    ' Mapper la zone cible à l'orientation Excel
    Dim orientation As XlPivotFieldOrientation
    
    Select Case targetArea
        Case AREA_ROWS
            orientation = xlRowField
        Case AREA_COLUMNS
            orientation = xlColumnField
        Case AREA_PAGES
            orientation = xlPageField
        Case AREA_DATA
            orientation = xlDataField
        Case Else
            Err.Raise ERR_INVALID_OPERATION, "clsExcelPivotTableAccessor.MoveField", "Zone cible invalide"
    End Select
    
    ' Appliquer l'orientation
    field.Orientation = orientation
    
    ' Appliquer la position si spécifiée
    If position >= 0 Then
        field.Position = position
    End If
    
    IPivotTableAccessor_MoveField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_MoveField = False
End Function

' ----- Méthodes de filtre et de format -----

'@Description: 
'@Param: 
'@Returns: 

Private Function IPivotTableAccessor_ApplyFilter(ByVal fieldName As String, ByVal values As Variant, _
                                              Optional ByVal include As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.ApplyFilter", "Champ non trouvé: " & fieldName
    End If
    
    ' Définir le type de filtre
    If include Then
        field.EnableMultiplePageItems = True
        
        ' Effacer d'abord tous les filtres
        field.ClearAllFilters
        
        ' Appliquer les valeurs d'inclusion
        If IsArray(values) Then
            Dim i As Long
            For i = LBound(values) To UBound(values)
                field.PivotItems(values(i)).Visible = True
            Next i
        Else
            field.PivotItems(values).Visible = True
        End If
    Else
        field.EnableMultiplePageItems = True
        
        ' D'abord montrer tous les éléments
        Dim j As Long
        For j = 1 To field.PivotItems.Count
            field.PivotItems(j).Visible = True
        Next j
        
        ' Ensuite cacher les éléments demandés
        If IsArray(values) Then
            Dim k As Long
            For k = LBound(values) To UBound(values)
                field.PivotItems(values(k)).Visible = False
            Next k
        Else
            field.PivotItems(values).Visible = False
        End If
    End If
    
    IPivotTableAccessor_ApplyFilter = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_ApplyFilter = False
End Function

Private Function IPivotTableAccessor_ClearFilters(ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.ClearFilters", "Champ non trouvé: " & fieldName
    End If
    
    ' Effacer les filtres
    field.ClearAllFilters
    
    IPivotTableAccessor_ClearFilters = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_ClearFilters = False
End Function

Private Function IPivotTableAccessor_ClearAllFilters() As Boolean
    On Error GoTo ErrorHandler
    
    ' Effacer tous les filtres de tous les champs
    Dim field As Object
    For Each field In mobjNativePivot.PivotFields
        field.ClearAllFilters
    Next field
    
    IPivotTableAccessor_ClearAllFilters = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_ClearAllFilters = False
End Function

Private Function IPivotTableAccessor_FormatDataField(ByVal dataFieldName As String, ByVal numberFormat As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Rechercher le champ de données
    Dim dataField As Object
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 1 To mobjNativePivot.DataFields.Count
        If mobjNativePivot.DataFields(i).Name = dataFieldName Then
            Set dataField = mobjNativePivot.DataFields(i)
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.FormatDataField", "Champ de données non trouvé: " & dataFieldName
    End If
    
    ' Appliquer le format numérique
    dataField.NumberFormat = numberFormat
    
    IPivotTableAccessor_FormatDataField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_FormatDataField = False
End Function

Private Function IPivotTableAccessor_SetSubtotal(ByVal fieldName As String, ByVal showSubtotal As Boolean, _
                                              Optional ByVal function As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.SetSubtotal", "Champ non trouvé: " & fieldName
    End If
    
    ' Vérifier que le champ est un champ de ligne ou de colonne
    If field.Orientation <> xlRowField And field.Orientation <> xlColumnField Then
        Err.Raise ERR_INVALID_OPERATION, "clsExcelPivotTableAccessor.SetSubtotal", _
                    "Seuls les champs de ligne ou de colonne peuvent avoir des sous-totaux"
    End If
    
    ' Définir les sous-totaux
    If showSubtotal Then
        If function >= 0 Then
            field.Subtotals(function) = True
        Else
            ' Utiliser le sous-total automatique
            field.Subtotals(1) = True  ' 1 = Automatic
        End If
    Else
        field.Subtotals(0) = False  ' 0 = No subtotals
    End If
    
    IPivotTableAccessor_SetSubtotal = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_SetSubtotal = False
End Function

' ----- Méthodes d'actions -----

'@Description: 
'@Param: 
'@Returns: 

Private Function IPivotTableAccessor_Refresh() As Boolean
    On Error GoTo ErrorHandler
    
    mobjNativePivot.RefreshTable
    IPivotTableAccessor_Refresh = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_Refresh = False
End Function

Private Function IPivotTableAccessor_ExpandItem(ByVal fieldName As String, ByVal itemName As String, _
                                             ByVal expand As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.ExpandItem", "Champ non trouvé: " & fieldName
    End If
    
    ' Vérifier que l'élément existe
    Dim item As Object
    On Error Resume Next
    Set item = field.PivotItems(itemName)
    On Error GoTo ErrorHandler
    
    If item Is Nothing Then
        Err.Raise ERR_INVALID_ITEM, "clsExcelPivotTableAccessor.ExpandItem", "Élément non trouvé: " & itemName
    End If
    
    ' Développer ou réduire l'élément
    If expand Then
        item.DrillTo fieldName
    Else
        item.ShowDetail = False
    End If
    
    IPivotTableAccessor_ExpandItem = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_ExpandItem = False
End Function

Private Function IPivotTableAccessor_ExpandField(ByVal fieldName As String, ByVal expand As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.ExpandField", "Champ non trouvé: " & fieldName
    End If
    
    ' Développer ou réduire tous les éléments du champ
    Dim item As Object
    For Each item In field.PivotItems
        On Error Resume Next
        item.ShowDetail = expand
        On Error GoTo 0
    Next item
    
    IPivotTableAccessor_ExpandField = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_ExpandField = False
End Function

Private Function IPivotTableAccessor_ExpandAll(ByVal expand As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Option applicable à tout le tableau croisé
    mobjNativePivot.ShowDrillIndicators = expand
    
    ' Pour chaque champ de ligne ou colonne, développer/réduire tous les éléments
    Dim field As Object
    For Each field In mobjNativePivot.PivotFields
        If field.Orientation = xlRowField Or field.Orientation = xlColumnField Then
            Dim item As Object
            For Each item In field.PivotItems
                On Error Resume Next
                item.ShowDetail = expand
                On Error GoTo 0
            Next item
        End If
    Next field
    
    IPivotTableAccessor_ExpandAll = True
    Exit Function
    
ErrorHandler:
    IPivotTableAccessor_ExpandAll = False
End Function

' ----- Méthodes de données -----

'@Description: 
'@Param: 
'@Returns: 

Private Function IPivotTableAccessor_GetDataFieldValues(ByVal dataFieldName As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Rechercher le champ de données
    Dim dataField As Object
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 1 To mobjNativePivot.DataFields.Count
        If mobjNativePivot.DataFields(i).Name = dataFieldName Then
            Set dataField = mobjNativePivot.DataFields(i)
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.GetDataFieldValues", "Champ de données non trouvé: " & dataFieldName
    End If
    
    ' Récupérer la plage de données
    IPivotTableAccessor_GetDataFieldValues = dataField.DataRange.Value
    Exit Function
    
ErrorHandler:
    Dim emptyArray() As Variant
    ReDim emptyArray(0, 0)
    IPivotTableAccessor_GetDataFieldValues = emptyArray
End Function

Private Function IPivotTableAccessor_GetAllValues() As Variant
    On Error GoTo ErrorHandler
    
    ' Récupérer tout le contenu du tableau croisé
    IPivotTableAccessor_GetAllValues = mobjNativePivot.TableRange2.Value
    Exit Function
    
ErrorHandler:
    Dim emptyArray() As Variant
    ReDim emptyArray(0, 0)
    IPivotTableAccessor_GetAllValues = emptyArray
End Function

Private Function IPivotTableAccessor_GetFieldItems(ByVal fieldName As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Vérifier que le champ existe
    Dim field As Object
    Set field = GetPivotFieldByName(fieldName)
    
    If field Is Nothing Then
        Err.Raise ERR_INVALID_FIELD, "clsExcelPivotTableAccessor.GetFieldItems", "Champ non trouvé: " & fieldName
    End If
    
    ' Récupérer les éléments du champ
    Dim items() As String
    ReDim items(1 To field.PivotItems.Count)
    
    Dim i As Long
    For i = 1 To field.PivotItems.Count
        items(i) = field.PivotItems(i).Name
    Next i
    
    IPivotTableAccessor_GetFieldItems = items
    Exit Function
    
ErrorHandler:
    Dim emptyArray() As String
    ReDim emptyArray(0)
    IPivotTableAccessor_GetFieldItems = emptyArray
End Function

' ==================== Méthodes privées de support ====================

''
' Recherche un champ de tableau croisé par son nom
' @param fieldName Nom du champ à rechercher
' @return Object Le champ PivotField trouvé ou Nothing si non trouvé
''
'@Description: 
'@Param: 
'@Returns: 

Private Function GetPivotFieldByName(ByVal fieldName As String) As Object
    On Error Resume Next
    
    ' Essayer de récupérer directement
    Set GetPivotFieldByName = mobjNativePivot.PivotFields(fieldName)
    
    ' Si non trouvé, parcourir la collection
    If GetPivotFieldByName Is Nothing Then
        Dim i As Long
        For i = 1 To mobjNativePivot.PivotFields.Count
            If mobjNativePivot.PivotFields(i).Name = fieldName Then
                Set GetPivotFieldByName = mobjNativePivot.PivotFields(i)
                Exit Function
            End If
        Next i
    End If
End Function

''
' Retourne l'objet PivotTable sous-jacent pour un accès direct par d'autres composants
' @return Object L'objet PivotTable natif d'Excel
''
'@Description: 
'@Param: 
'@Returns: 

Public Function GetNativePivotTable() As Object
    Set GetNativePivotTable = mobjNativePivot
End Function