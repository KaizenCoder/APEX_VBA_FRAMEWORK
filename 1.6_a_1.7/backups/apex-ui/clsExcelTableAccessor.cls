VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelTableAccessor"

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
'@ModuleDescription("Implémentation de l'interface ITableAccessor pour manipuler les Tables Excel")
'@Implements ITableAccessor
Option Explicit

Implements ITableAccessor

' Variables privées
Private mobjSheet As Object ' ISheetAccessor
Private mobjNativeTable As Object ' ListObject
Private mstrTableName As String

' Constantes d'erreur
Private Const ERR_INVALID_TABLE As Long = 1001
Private Const ERR_INVALID_COLUMN As Long = 1002
Private Const ERR_INVALID_ROW As Long = 1003
Private Const ERR_INVALID_DATA As Long = 1004

'*******************************************************************************
' Module : clsExcelTableAccessor
' Author : [Votre nom]
' Date   : 12/04/2025
' Purpose: Implémentation de l'interface ITableAccessor pour manipuler les
'          Tables Excel (ListObject) avec opérations CRUD et formatage
'*******************************************************************************

''
' Initialise un nouvel accesseur de table Excel
' @param sheetAccessor L'accesseur de feuille qui contient la table
' @param tableName Le nom de la table à manipuler
' @return Boolean True si l'initialisation a réussi
''
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal sheetAccessor As Object, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier et stocker l'accesseur de feuille
    Set mobjSheet = sheetAccessor
    
    ' Récupérer la table native par son nom
    Dim nativeSheet As Object
    Set nativeSheet = mobjSheet.GetNativeSheet
    
    ' Rechercher la table par son nom
    Dim tbl As Object
    For Each tbl In nativeSheet.ListObjects
        If tbl.Name = tableName Then
            Set mobjNativeTable = tbl
            mstrTableName = tableName
            Initialize = True
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next tbl
    
    ' Table non trouvée
    Err.Raise ERR_INVALID_TABLE, "clsExcelTableAccessor", "La table '" & tableName & "' n'existe pas dans la feuille"
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
ErrorHandler:
    Initialize = False
    If Err.Number <> ERR_INVALID_TABLE Then
        Err.Raise Err.Number, "clsExcelTableAccessor.Initialize", Err.Description
    End If
End Function

''
' Crée une nouvelle table Excel à partir d'une plage de données
' @param sheetAccessor L'accesseur de feuille où créer la table
' @param rangeName La plage contenant les données (ex: "A1:D10")
' @param tableName Le nom à donner à la nouvelle table
' @param hasHeaders Indique si la première ligne contient des en-têtes
' @return Object L'accesseur pour la table nouvellement créée
''
'@Description: 
'@Param: 
'@Returns: 

Public Function CreateTableFromRange(ByVal sheetAccessor As Object, ByVal rangeName As String, _
                                   ByVal tableName As String, Optional ByVal hasHeaders As Boolean = True) As clsExcelTableAccessor
    On Error GoTo ErrorHandler
    
    ' Obtenir la feuille native
    Dim nativeSheet As Object
    Set nativeSheet = sheetAccessor.GetNativeSheet
    
    ' Créer la table à partir de la plage
    Dim nativeRange As Object
    Set nativeRange = nativeSheet.Range(rangeName)
    
    Dim tableHeaderSetting As XlYesNoGuess
    If hasHeaders Then
        tableHeaderSetting = xlYes
    Else
        tableHeaderSetting = xlNo
    End If
    
    ' Créer la table
    Dim newTable As Object
    Set newTable = nativeSheet.ListObjects.Add(xlSrcRange, nativeRange, , tableHeaderSetting)
    newTable.Name = tableName
    
    ' Créer et initialiser un nouvel accesseur pour cette table
    Dim tableAccessor As New clsExcelTableAccessor
    tableAccessor.Initialize sheetAccessor, tableName
    
    Set CreateTableFromRange = tableAccessor
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.CreateTableFromRange", Err.Description
End Function

' ==================== Implémentation de ITableAccessor ====================

' ----- Propriétés -----

Private Property Get ITableAccessor_TableName() As String
    ITableAccessor_TableName = mstrTableName
End Property

Private Property Get ITableAccessor_ParentSheet() As Object
    Set ITableAccessor_ParentSheet = mobjSheet
End Property

Private Property Get ITableAccessor_RowCount() As Long
    On Error Resume Next
    ITableAccessor_RowCount = mobjNativeTable.DataBodyRange.Rows.Count
    If Err.Number <> 0 Then ITableAccessor_RowCount = 0
End Property

Private Property Get ITableAccessor_ColumnCount() As Long
    ITableAccessor_ColumnCount = mobjNativeTable.ListColumns.Count
End Property

Private Property Get ITableAccessor_ColumnNames() As Variant
    Dim cols As Long
    Dim i As Long
    Dim result() As String
    
    cols = mobjNativeTable.ListColumns.Count
    ReDim result(1 To cols)
    
    For i = 1 To cols
        result(i) = mobjNativeTable.ListColumns(i).Name
    Next i
    
    ITableAccessor_ColumnNames = result
End Property

' ----- Méthodes de lecture -----

'@Description: 
'@Param: 
'@Returns: 

Private Function ITableAccessor_ReadAllData() As Variant
    On Error GoTo ErrorHandler
    
    If mobjNativeTable.DataBodyRange Is Nothing Then
        ' Table vide, retourner un tableau vide
        Dim emptyArray() As Variant
        ReDim emptyArray(0, 0)
        ITableAccessor_ReadAllData = emptyArray
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Lire toutes les données (sans l'en-tête)
    ITableAccessor_ReadAllData = mobjNativeTable.DataBodyRange.Value
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.ReadAllData", Err.Description
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_ReadHeader() As Variant
    On Error GoTo ErrorHandler
    
    Dim cols As Long
    Dim i As Long
    Dim result() As String
    
    cols = mobjNativeTable.ListColumns.Count
    ReDim result(1 To cols)
    
    For i = 1 To cols
        result(i) = mobjNativeTable.ListColumns(i).Name
    Next i
    
    ITableAccessor_ReadHeader = result
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.ReadHeader", Err.Description
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_ReadRow(ByVal rowIndex As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index de ligne est valide
    If rowIndex < 1 Or rowIndex > Me.ITableAccessor_RowCount Then
        Err.Raise ERR_INVALID_ROW, "clsExcelTableAccessor.ReadRow", "Index de ligne invalide: " & rowIndex
    End If
    
    ' Lire une ligne spécifique
    Dim rowData As Variant
    rowData = mobjNativeTable.DataBodyRange.Rows(rowIndex).Value
    
    ' Convertir en tableau 1D
    Dim cols As Long
    Dim result() As Variant
    cols = UBound(rowData, 2)
    ReDim result(1 To cols)
    
    Dim j As Long
    For j = 1 To cols
        result(j) = rowData(1, j)
    Next j
    
    ITableAccessor_ReadRow = result
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.ReadRow", Err.Description
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_ReadColumn(ByVal columnName As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.ReadColumn", "Colonne non trouvée: " & columnName
    End If
    
    ' Lire une colonne spécifique
    Dim colData As Variant
    colData = mobjNativeTable.ListColumns(columnName).DataBodyRange.Value
    
    ' Convertir en tableau 1D si nécessaire
    If Me.ITableAccessor_RowCount = 1 Then
        Dim singleValue(1 To 1) As Variant
        singleValue(1) = colData
        ITableAccessor_ReadColumn = singleValue
    Else
        Dim rows As Long
        Dim result() As Variant
        rows = UBound(colData, 1)
        ReDim result(1 To rows)
        
        Dim i As Long
        For i = 1 To rows
            result(i) = colData(i, 1)
        Next i
        
        ITableAccessor_ReadColumn = result
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.ReadColumn", Err.Description
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_ReadCell(ByVal rowIndex As Long, ByVal columnName As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index de ligne est valide
    If rowIndex < 1 Or rowIndex > Me.ITableAccessor_RowCount Then
        Err.Raise ERR_INVALID_ROW, "clsExcelTableAccessor.ReadCell", "Index de ligne invalide: " & rowIndex
    End If
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.ReadCell", "Colonne non trouvée: " & columnName
    End If
    
    ' Lire une cellule spécifique
    ITableAccessor_ReadCell = mobjNativeTable.ListColumns(columnName).DataBodyRange.Cells(rowIndex, 1).Value
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.ReadCell", Err.Description
End Function

' ----- Méthodes d'écriture -----

'@Description: 
'@Param: 
'@Returns: 

Private Function ITableAccessor_WriteRow(ByVal rowIndex As Long, ByVal rowData As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index de ligne est valide
    If rowIndex < 1 Or rowIndex > Me.ITableAccessor_RowCount Then
        Err.Raise ERR_INVALID_ROW, "clsExcelTableAccessor.WriteRow", "Index de ligne invalide: " & rowIndex
    End If
    
    ' Vérifier que les données sont valides
    If Not IsArray(rowData) Then
        Err.Raise ERR_INVALID_DATA, "clsExcelTableAccessor.WriteRow", "Les données doivent être un tableau"
    End If
    
    ' Convertir en tableau 2D pour l'affectation de la plage
    Dim cols As Long
    cols = UBound(rowData)
    
    If cols <> Me.ITableAccessor_ColumnCount Then
        Err.Raise ERR_INVALID_DATA, "clsExcelTableAccessor.WriteRow", _
                "Le nombre de colonnes dans les données (" & cols & ") ne correspond pas au nombre de colonnes de la table (" & _
                Me.ITableAccessor_ColumnCount & ")"
    End If
    
    Dim data2D() As Variant
    ReDim data2D(1 To 1, 1 To cols)
    
    Dim j As Long
    For j = 1 To cols
        data2D(1, j) = rowData(j)
    Next j
    
    ' Écrire les données
    mobjNativeTable.DataBodyRange.Rows(rowIndex).Value = data2D
    ITableAccessor_WriteRow = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_WriteRow = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_WriteColumn(ByVal columnName As String, ByVal columnData As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.WriteColumn", "Colonne non trouvée: " & columnName
    End If
    
    ' Vérifier que les données sont valides
    If Not IsArray(columnData) Then
        Err.Raise ERR_INVALID_DATA, "clsExcelTableAccessor.WriteColumn", "Les données doivent être un tableau"
    End If
    
    Dim rows As Long
    rows = UBound(columnData)
    
    If rows <> Me.ITableAccessor_RowCount Then
        Err.Raise ERR_INVALID_DATA, "clsExcelTableAccessor.WriteColumn", _
                "Le nombre de lignes dans les données (" & rows & ") ne correspond pas au nombre de lignes de la table (" & _
                Me.ITableAccessor_RowCount & ")"
    End If
    
    ' Convertir en tableau 2D pour l'affectation de la plage
    Dim data2D() As Variant
    ReDim data2D(1 To rows, 1 To 1)
    
    Dim i As Long
    For i = 1 To rows
        data2D(i, 1) = columnData(i)
    Next i
    
    ' Écrire les données
    mobjNativeTable.ListColumns(columnName).DataBodyRange.Value = data2D
    ITableAccessor_WriteColumn = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_WriteColumn = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_WriteCell(ByVal rowIndex As Long, ByVal columnName As String, ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index de ligne est valide
    If rowIndex < 1 Or rowIndex > Me.ITableAccessor_RowCount Then
        Err.Raise ERR_INVALID_ROW, "clsExcelTableAccessor.WriteCell", "Index de ligne invalide: " & rowIndex
    End If
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.WriteCell", "Colonne non trouvée: " & columnName
    End If
    
    ' Écrire la valeur
    mobjNativeTable.ListColumns(columnName).DataBodyRange.Cells(rowIndex, 1).Value = value
    ITableAccessor_WriteCell = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_WriteCell = False
End Function

' ----- Méthodes de manipulation de la structure -----

'@Description: 
'@Param: 
'@Returns: 

Private Function ITableAccessor_AddRow(Optional ByVal rowData As Variant) As Long
    On Error GoTo ErrorHandler
    
    ' Ajouter une ligne à la fin de la table
    Dim newRow As Object
    Dim rowCount As Long
    
    ' Déterminer l'index de la nouvelle ligne
    rowCount = Me.ITableAccessor_RowCount
    
    ' Ajouter une ligne vide
    Set newRow = mobjNativeTable.ListRows.Add
    
    ' Si des données ont été fournies, les écrire
    If Not IsMissing(rowData) Then
        If IsArray(rowData) Then
            ' Convertir en tableau 2D pour l'affectation de la plage
            Dim cols As Long
            cols = UBound(rowData)
            
            If cols <> Me.ITableAccessor_ColumnCount Then
                Err.Raise ERR_INVALID_DATA, "clsExcelTableAccessor.AddRow", _
                        "Le nombre de colonnes dans les données (" & cols & ") ne correspond pas au nombre de colonnes de la table (" & _
                        Me.ITableAccessor_ColumnCount & ")"
            End If
            
            Dim data2D() As Variant
            ReDim data2D(1 To 1, 1 To cols)
            
            Dim j As Long
            For j = 1 To cols
                data2D(1, j) = rowData(j)
            Next j
            
            ' Écrire les données dans la nouvelle ligne
            newRow.Range.Value = data2D
        End If
    End If
    
    ' Retourner l'index de la nouvelle ligne (1-based)
    ITableAccessor_AddRow = rowCount + 1
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsExcelTableAccessor.AddRow", Err.Description
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_DeleteRow(ByVal rowIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que l'index de ligne est valide
    If rowIndex < 1 Or rowIndex > Me.ITableAccessor_RowCount Then
        Err.Raise ERR_INVALID_ROW, "clsExcelTableAccessor.DeleteRow", "Index de ligne invalide: " & rowIndex
    End If
    
    ' Supprimer la ligne
    mobjNativeTable.ListRows(rowIndex).Delete
    ITableAccessor_DeleteRow = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_DeleteRow = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_AddColumn(ByVal columnName As String, Optional ByVal columnData As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne n'existe pas déjà
    If GetColumnIndexByName(columnName) > 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.AddColumn", "La colonne existe déjà: " & columnName
    End If
    
    ' Ajouter une nouvelle colonne
    Dim newCol As Object
    Set newCol = mobjNativeTable.ListColumns.Add
    newCol.Name = columnName
    
    ' Si des données ont été fournies, les écrire
    If Not IsMissing(columnData) Then
        If IsArray(columnData) Then
            Dim rows As Long
            rows = UBound(columnData)
            
            If rows <> Me.ITableAccessor_RowCount Then
                Err.Raise ERR_INVALID_DATA, "clsExcelTableAccessor.AddColumn", _
                        "Le nombre de lignes dans les données (" & rows & ") ne correspond pas au nombre de lignes de la table (" & _
                        Me.ITableAccessor_RowCount & ")"
            End If
            
            ' Convertir en tableau 2D pour l'affectation de la plage
            Dim data2D() As Variant
            ReDim data2D(1 To rows, 1 To 1)
            
            Dim i As Long
            For i = 1 To rows
                data2D(i, 1) = columnData(i)
            Next i
            
            ' Écrire les données
            newCol.DataBodyRange.Value = data2D
        End If
    End If
    
    ITableAccessor_AddColumn = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_AddColumn = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_DeleteColumn(ByVal columnName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.DeleteColumn", "Colonne non trouvée: " & columnName
    End If
    
    ' Supprimer la colonne
    mobjNativeTable.ListColumns(columnName).Delete
    ITableAccessor_DeleteColumn = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_DeleteColumn = False
End Function

' ----- Méthodes de filtrage et tri -----

'@Description: 
'@Param: 
'@Returns: 

Private Function ITableAccessor_ApplyFilter(ByVal columnName As String, ByVal criteria As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.ApplyFilter", "Colonne non trouvée: " & columnName
    End If
    
    ' Activer le filtrage s'il n'est pas déjà activé
    If Not mobjNativeTable.ShowAutoFilter Then
        mobjNativeTable.ShowAutoFilter = True
    End If
    
    ' Appliquer le filtre
    mobjNativeTable.Range.AutoFilter Field:=colIndex, Criteria1:=criteria
    ITableAccessor_ApplyFilter = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_ApplyFilter = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_ClearFilters() As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si le filtrage est actif
    If mobjNativeTable.ShowAutoFilter Then
        ' Effacer tous les filtres
        mobjNativeTable.Range.AutoFilter
    End If
    
    ITableAccessor_ClearFilters = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_ClearFilters = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_SortByColumn(ByVal columnName As String, Optional ByVal ascending As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.SortByColumn", "Colonne non trouvée: " & columnName
    End If
    
    ' Déterminer l'ordre de tri
    Dim sortOrder As XlSortOrder
    If ascending Then
        sortOrder = xlAscending
    Else
        sortOrder = xlDescending
    End If
    
    ' Trier la table
    With mobjNativeTable.Sort
        .SortFields.Clear
        .SortFields.Add Key:=mobjNativeTable.ListColumns(columnName).Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=sortOrder, _
                        DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ITableAccessor_SortByColumn = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_SortByColumn = False
End Function

' ----- Méthodes de mise en forme -----

'@Description: 
'@Param: 
'@Returns: 

Private Function ITableAccessor_ApplyTableStyle(ByVal styleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Appliquer un style prédéfini à la table
    mobjNativeTable.TableStyle = styleName
    ITableAccessor_ApplyTableStyle = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_ApplyTableStyle = False
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function ITableAccessor_SetConditionalFormatting(ByVal columnName As String, ByVal formulaString As String, _
                                                  ByVal formatSettings As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que la colonne existe
    Dim colIndex As Long
    colIndex = GetColumnIndexByName(columnName)
    
    If colIndex = 0 Then
        Err.Raise ERR_INVALID_COLUMN, "clsExcelTableAccessor.SetConditionalFormatting", "Colonne non trouvée: " & columnName
    End If
    
    ' Supprimer toute mise en forme conditionnelle existante pour cette colonne
    mobjNativeTable.ListColumns(columnName).DataBodyRange.FormatConditions.Delete
    
    ' Ajouter la nouvelle mise en forme conditionnelle
    Dim fc As Object
    Set fc = mobjNativeTable.ListColumns(columnName).DataBodyRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaString)
    
    ' Appliquer les paramètres de format
    If Not IsMissing(formatSettings) Then
        If IsObject(formatSettings) Then
            ' Si c'est un dictionnaire ou un objet avec des propriétés
            ' Implémenter selon le format de formatSettings
        ElseIf VarType(formatSettings) = vbString Then
            ' Si c'est une chaîne représentant une couleur
            fc.Interior.Color = formatSettings
        End If
    End If
    
    ITableAccessor_SetConditionalFormatting = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    ITableAccessor_SetConditionalFormatting = False
End Function

' ==================== Méthodes privées de support ====================

''
' Obtient l'index d'une colonne par son nom
' @param columnName Nom de la colonne à rechercher
' @return Long Index de la colonne (0 si non trouvée)
''
'@Description: 
'@Param: 
'@Returns: 

Private Function GetColumnIndexByName(ByVal columnName As String) As Long
    On Error Resume Next
    
    Dim i As Long
    For i = 1 To mobjNativeTable.ListColumns.Count
        If mobjNativeTable.ListColumns(i).Name = columnName Then
            GetColumnIndexByName = i
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
    
    GetColumnIndexByName = 0
End Function

''
' Retourne l'objet ListObject sous-jacent pour un accès direct par d'autres composants
' @return Object L'objet ListObject natif d'Excel
''
'@Description: 
'@Param: 
'@Returns: 

Public Function GetNativeTable() As Object
    Set GetNativeTable = mobjNativeTable
End Function