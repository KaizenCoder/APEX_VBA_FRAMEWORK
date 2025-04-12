VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCachedSheetAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Classe : clsCachedSheetAccessor
' Version : 1.0
' Purpose : Accesseur de feuille avec mécanisme de cache pour optimiser les performances
' Author  : APEX Framework Team
' Date    : 2024-04-11
' ==========================================================================

Private Const CLASS_NAME As String = "clsCachedSheetAccessor"
Private Const LOG_CATEGORY As String = "CachedSheetAccessor"

' Objet implémentant l'interface ISheetAccessor
Private m_baseAccessor As ISheetAccessor

' Mécanisme de cache
Private m_cacheEnabled As Boolean
Private m_cacheData As Variant      ' Tableau 2D contenant les données en cache
Private m_cacheValid As Boolean     ' Indique si le cache est à jour
Private m_cacheStartRow As Long     ' Première ligne du cache
Private m_cacheStartCol As Long     ' Première colonne du cache
Private m_cacheEndRow As Long       ' Dernière ligne du cache
Private m_cacheEndCol As Long       ' Dernière colonne du cache

' Statistiques
Private m_cacheHits As Long
Private m_cacheMisses As Long
Private m_cacheInvalidations As Long

' Services
Private m_logger As ILogger

' ============================================================================
' Function: Initialize
' Purpose: Initialise l'accesseur en cache avec un accesseur de base
' ============================================================================
Public Function Initialize(ByVal baseAccessor As ISheetAccessor, Optional ByVal enableCache As Boolean = True) As clsCachedSheetAccessor
    If baseAccessor Is Nothing Then
        ReportError ERR_INVALID_ARGUMENT, "Initialize: baseAccessor cannot be Nothing", CLASS_NAME
        Exit Function
    End If
    
    ' Récupérer les services
    Dim appContext As IApplicationContext
    Set appContext = GetApplicationContext()
    Set m_logger = appContext.Logger
    
    ' Initialiser l'objet
    Set m_baseAccessor = baseAccessor
    m_cacheEnabled = enableCache
    m_cacheValid = False
    
    ' Initialiser les statistiques
    m_cacheHits = 0
    m_cacheMisses = 0
    m_cacheInvalidations = 0
    
    m_logger.LogInfo "CachedSheetAccessor initialisé, cache " & IIf(m_cacheEnabled, "activé", "désactivé"), LOG_CATEGORY
    
    Set Initialize = Me
End Function

' ============================================================================
' Property: IsCacheEnabled
' Purpose: Active/désactive le cache
' ============================================================================
Public Property Get IsCacheEnabled() As Boolean
    IsCacheEnabled = m_cacheEnabled
End Property

Public Property Let IsCacheEnabled(ByVal value As Boolean)
    m_cacheEnabled = value
    If Not m_cacheEnabled Then
        ClearCache
    End If
    m_logger.LogDebug "Cache " & IIf(m_cacheEnabled, "activé", "désactivé"), LOG_CATEGORY
End Property

' ============================================================================
' Method: ClearCache
' Purpose: Vide le cache
' ============================================================================
Public Sub ClearCache()
    m_cacheValid = False
    Erase m_cacheData
    m_cacheStartRow = 0
    m_cacheStartCol = 0
    m_cacheEndRow = 0
    m_cacheEndCol = 0
    m_cacheInvalidations = m_cacheInvalidations + 1
    m_logger.LogDebug "Cache vidé", LOG_CATEGORY
End Sub

' ============================================================================
' Method: GetCacheStatistics
' Purpose: Renvoie les statistiques du cache
' ============================================================================
Public Sub GetCacheStatistics(ByRef hits As Long, ByRef misses As Long, ByRef invalidations As Long)
    hits = m_cacheHits
    misses = m_cacheMisses
    invalidations = m_cacheInvalidations
End Sub

' ============================================================================
' Method: ReadRange
' Purpose: Lit les données d'une plage, en utilisant le cache si possible
' ============================================================================
Public Function ReadRange(ByVal startRow As Long, ByVal startCol As Long, _
                         ByVal endRow As Long, ByVal endCol As Long) As Variant
    Dim result As Variant
    Dim i As Long, j As Long, cacheI As Long, cacheJ As Long
    
    On Error GoTo ErrorHandler
    
    ' Si le cache est désactivé, déléguer directement à l'accesseur de base
    If Not m_cacheEnabled Then
        ReadRange = m_baseAccessor.ReadRange(startRow, startCol, endRow, endCol)
        Exit Function
    End If
    
    ' Vérifier si la demande est entièrement dans le cache
    If m_cacheValid And IsRangeInCache(startRow, startCol, endRow, endCol) Then
        ' Cache hit - extraire les données du cache
        m_cacheHits = m_cacheHits + 1
        m_logger.LogDebug "Cache hit pour la plage R" & startRow & "C" & startCol & ":R" & endRow & "C" & endCol, LOG_CATEGORY
        
        ' Créer un sous-tableau à partir du cache
        ReDim result(1 To (endRow - startRow + 1), 1 To (endCol - startCol + 1))
        
        For i = startRow To endRow
            For j = startCol To endCol
                ' Calculer les indices relatifs dans le cache
                cacheI = i - m_cacheStartRow + 1
                cacheJ = j - m_cacheStartCol + 1
                
                ' Copier la valeur du cache vers le résultat
                result(i - startRow + 1, j - startCol + 1) = m_cacheData(cacheI, cacheJ)
            Next j
        Next i
    Else
        ' Cache miss - lire les données de la source et mettre à jour le cache
        m_cacheMisses = m_cacheMisses + 1
        m_logger.LogDebug "Cache miss pour la plage R" & startRow & "C" & startCol & ":R" & endRow & "C" & endCol, LOG_CATEGORY
        
        ' Déterminer une zone étendue à mettre en cache (préchargement)
        Dim cacheStartRow As Long, cacheStartCol As Long
        Dim cacheEndRow As Long, cacheEndCol As Long
        
        ' Stratégie de préchargement : ajouter 20% de lignes supplémentaires autour de la zone demandée
        ' avec un minimum de 10 lignes et un maximum de 1000 lignes
        Dim extraRows As Long
        extraRows = WorksheetFunction.Max(10, WorksheetFunction.Min(1000, (endRow - startRow + 1) * 0.2))
        
        cacheStartRow = WorksheetFunction.Max(1, startRow - extraRows)
        cacheEndRow = endRow + extraRows
        
        ' Colonnes : prendre toutes les colonnes de la plage demandée + 5 colonnes de chaque côté
        cacheStartCol = WorksheetFunction.Max(1, startCol - 5)
        cacheEndCol = endCol + 5
        
        ' Lire les données étendues depuis la source
        m_cacheData = m_baseAccessor.ReadRange(cacheStartRow, cacheStartCol, cacheEndRow, cacheEndCol)
        
        ' Mettre à jour les informations du cache
        m_cacheStartRow = cacheStartRow
        m_cacheStartCol = cacheStartCol
        m_cacheEndRow = cacheEndRow
        m_cacheEndCol = cacheEndCol
        m_cacheValid = True
        
        ' Extraire le sous-ensemble demandé
        ReDim result(1 To (endRow - startRow + 1), 1 To (endCol - startCol + 1))
        
        For i = startRow To endRow
            For j = startCol To endCol
                ' Calculer les indices relatifs dans le cache
                cacheI = i - m_cacheStartRow + 1
                cacheJ = j - m_cacheStartCol + 1
                
                ' Copier la valeur du cache vers le résultat
                result(i - startRow + 1, j - startCol + 1) = m_cacheData(cacheI, cacheJ)
            Next j
        Next i
    End If
    
    ReadRange = result
    Exit Function
    
ErrorHandler:
    ' En cas d'erreur, faire une lecture directe sans cache
    m_logger.LogError "Erreur ReadRange: " & Err.Number & " - " & Err.Description, LOG_CATEGORY
    ReadRange = m_baseAccessor.ReadRange(startRow, startCol, endRow, endCol)
End Function

' ============================================================================
' Method: WriteRange
' Purpose: Écrit des données dans une plage et invalide le cache si nécessaire
' ============================================================================
Public Sub WriteRange(ByRef data As Variant, ByVal startRow As Long, ByVal startCol As Long)
    Dim endRow As Long, endCol As Long
    
    On Error GoTo ErrorHandler
    
    ' Calculer la fin de la plage d'après les dimensions du tableau de données
    endRow = startRow + UBound(data, 1) - 1
    endCol = startCol + UBound(data, 2) - 1
    
    ' Écrire les données via l'accesseur de base
    m_baseAccessor.WriteRange data, startRow, startCol
    
    ' Si le cache est activé, vérifier s'il faut l'invalider
    If m_cacheEnabled And m_cacheValid Then
        If RangesOverlap(startRow, startCol, endRow, endCol, _
                        m_cacheStartRow, m_cacheStartCol, m_cacheEndRow, m_cacheEndCol) Then
            ' La plage écrite chevauche la plage en cache, invalider le cache
            m_logger.LogDebug "Invalidation du cache suite à écriture dans la plage R" & startRow & "C" & startCol & ":R" & endRow & "C" & endCol, LOG_CATEGORY
            ClearCache
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    m_logger.LogError "Erreur WriteRange: " & Err.Number & " - " & Err.Description, LOG_CATEGORY
    ReportError Err.Number, "WriteRange: " & Err.Description, CLASS_NAME
End Sub

' ============================================================================
' Method: RefreshCache
' Purpose: Force une actualisation du cache
' ============================================================================
Public Sub RefreshCache()
    If m_cacheEnabled And m_cacheValid Then
        ' Recharger les données actuelles du cache
        m_cacheData = m_baseAccessor.ReadRange(m_cacheStartRow, m_cacheStartCol, m_cacheEndRow, m_cacheEndCol)
        m_logger.LogDebug "Cache actualisé pour la plage R" & m_cacheStartRow & "C" & m_cacheStartCol & ":R" & m_cacheEndRow & "C" & m_cacheEndCol, LOG_CATEGORY
    End If
End Sub

' ============================================================================
' Method: IsRangeInCache
' Purpose: Vérifie si une plage est entièrement dans le cache
' ============================================================================
Private Function IsRangeInCache(ByVal startRow As Long, ByVal startCol As Long, _
                              ByVal endRow As Long, ByVal endCol As Long) As Boolean
    IsRangeInCache = (m_cacheValid And _
                    startRow >= m_cacheStartRow And _
                    startCol >= m_cacheStartCol And _
                    endRow <= m_cacheEndRow And _
                    endCol <= m_cacheEndCol)
End Function

' ============================================================================
' Method: RangesOverlap
' Purpose: Vérifie si deux plages se chevauchent
' ============================================================================
Private Function RangesOverlap(ByVal r1StartRow As Long, ByVal r1StartCol As Long, _
                            ByVal r1EndRow As Long, ByVal r1EndCol As Long, _
                            ByVal r2StartRow As Long, ByVal r2StartCol As Long, _
                            ByVal r2EndRow As Long, ByVal r2EndCol As Long) As Boolean
    RangesOverlap = Not (r1EndRow < r2StartRow Or r1StartRow > r2EndRow Or _
                       r1EndCol < r2StartCol Or r1StartCol > r2EndCol)
End Function

' ============================================================================
' Method: Class_Terminate
' Purpose: Nettoyage à la destruction de l'objet
' ============================================================================
Private Sub Class_Terminate()
    If m_cacheEnabled And m_logger Is Not Nothing Then
        ' Logguer les statistiques du cache avant destruction
        m_logger.LogInfo "Statistiques du cache - Hits: " & m_cacheHits & ", Misses: " & m_cacheMisses & ", Invalidations: " & m_cacheInvalidations, LOG_CATEGORY
    End If
    
    ' Libérer les références
    Set m_baseAccessor = Nothing
    Set m_logger = Nothing
End Sub 