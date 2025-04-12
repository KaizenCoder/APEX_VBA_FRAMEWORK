VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMigrationManager"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Database.Migration")
Option Explicit

'==========================================================================
' Classe    : clsMigrationManager
' Purpose   : Gestion des migrations de base de données
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : DB-002
'==========================================================================

Private Type TMigrationState
    CurrentVersion As String
    TargetVersion As String
    IsRunning As Boolean
    HasError As Boolean
    ErrorMessage As String
    StartTime As Date
    ScriptsPath As String
    BackupPath As String
End Type

Private Type TMigrationStats
    TotalMigrations As Long
    SuccessfulMigrations As Long
    FailedMigrations As Long
    AverageMigrationTime As Double
    LongestMigration As Double
    LastMigrationDate As Date
End Type

Private Type TMigrationScript
    Version As String
    Name As String
    Path As String
    Type As String ' UP ou DOWN
    Dependencies() As String
    IsApplied As Boolean
End Type

Private this As TMigrationState
Private m_Stats As TMigrationStats
Private m_Scripts As Collection
Private m_Logger As ILoggerBase
Private m_MetricsCollector As clsMetricsCollector
Private m_TransactionManager As clsTransactionManager

'@Description("Initialise le gestionnaire de migration")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(ByVal scriptsPath As String, _
                         ByVal backupPath As String, _
                         Optional ByVal logger As ILoggerBase = Nothing, _
                         Optional ByVal metricsCollector As clsMetricsCollector = Nothing, _
                         Optional ByVal transactionManager As clsTransactionManager = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier les chemins
    If Not FolderExists(scriptsPath) Then
        LogError "Initialize", "Dossier de scripts introuvable: " & scriptsPath
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    If Not FolderExists(backupPath) Then
        CreateFolder backupPath
    End If
    
    ' Initialiser l'état
    this.ScriptsPath = scriptsPath
    this.BackupPath = backupPath
    this.IsRunning = False
    this.HasError = False
    
    ' Initialiser les dépendances
    Set m_Logger = logger
    Set m_MetricsCollector = metricsCollector
    Set m_TransactionManager = transactionManager
    Set m_Scripts = New Collection
    
    ' Charger les scripts
    If Not LoadMigrationScripts Then
        LogError "Initialize", "Échec du chargement des scripts"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Initialize = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "Initialize", Err.Description
    Initialize = False
End Function

'@Description("Charge les scripts de migration")
'@Description: 
'@Param: 
'@Returns: 

Private Function LoadMigrationScripts() As Boolean
    On Error GoTo ErrorHandler
    
    ' Réinitialiser la collection
    Set m_Scripts = New Collection
    
    ' Parcourir les fichiers
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(this.ScriptsPath)
    
    Dim file As Object
    For Each file In folder.Files
        If file.Name Like "V*__*.sql" Then
            ' Parser le nom du fichier
            Dim script As TMigrationScript
            If ParseScriptFileName(file.Name, script) Then
                script.Path = file.Path
                script.IsApplied = False
                m_Scripts.Add script
            End If
        End If
    Next file
    
    LoadMigrationScripts = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "LoadMigrationScripts", Err.Description
    LoadMigrationScripts = False
End Function

'@Description("Parse le nom d'un fichier de migration")
'@Description: 
'@Param: 
'@Returns: 

Private Function ParseScriptFileName(ByVal fileName As String, ByRef script As TMigrationScript) As Boolean
    ' Format attendu : V1.0.0__Description.sql
    
    Dim parts() As String
    parts = Split(fileName, "__")
    
    If UBound(parts) <> 1 Then Exit Function
    
    ' Version
    script.Version = Mid$(parts(0), 2) ' Enlever le V
    
    ' Nom
    script.Name = Left$(parts(1), Len(parts(1)) - 4) ' Enlever .sql
    
    ' Type (par défaut UP)
    script.Type = "UP"
    
    ParseScriptFileName = True
End Function

'@Description("Migre la base de données vers une version spécifique")
'@Description: 
'@Param: 
'@Returns: 

Public Function MigrateTo(ByVal targetVersion As String, Optional ByVal backup As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier si une migration est en cours
    If this.IsRunning Then
        LogError "MigrateTo", "Une migration est déjà en cours"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Initialiser l'état
    this.IsRunning = True
    this.HasError = False
    this.ErrorMessage = ""
    this.StartTime = Now
    this.TargetVersion = targetVersion
    
    ' Créer une sauvegarde si demandé
    If backup Then
        If Not CreateBackup Then
            LogError "MigrateTo", "Échec de la sauvegarde"
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    End If
    
    ' Déterminer les scripts à exécuter
    Dim scriptsToRun As Collection
    Set scriptsToRun = GetScriptsToRun(targetVersion)
    
    If scriptsToRun.Count = 0 Then
        LogInfo "Aucun script à exécuter"
        MigrateTo = True
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Exécuter les scripts dans une transaction
    If Not m_TransactionManager Is Nothing Then
        m_TransactionManager.BeginTransaction
    End If
    
    Dim script As TMigrationScript
    For Each script In scriptsToRun
        If Not ExecuteScript(script) Then
            If Not m_TransactionManager Is Nothing Then
                m_TransactionManager.RollbackTransaction
            End If
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next script
    
    If Not m_TransactionManager Is Nothing Then
        m_TransactionManager.CommitTransaction
    End If
    
    ' Mettre à jour les statistiques
    UpdateMigrationMetrics True
    
    MigrateTo = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "MigrateTo", Err.Description
    If Not m_TransactionManager Is Nothing Then
        m_TransactionManager.RollbackTransaction
    End If
    MigrateTo = False
End Function

'@Description("Crée une sauvegarde de la base de données")
'@Description: 
'@Param: 
'@Returns: 

Private Function CreateBackup() As Boolean
    On Error GoTo ErrorHandler
    
    ' Générer le nom du fichier
    Dim backupFile As String
    backupFile = this.BackupPath & "\backup_" & _
                 Format(Now, "yyyymmdd_hhnnss") & ".bak"
    
    ' TODO: Implémenter la sauvegarde selon le SGBD utilisé
    ' Exemple pour SQL Server :
    ' BACKUP DATABASE [DB] TO DISK = 'backupFile'
    
    CreateBackup = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "CreateBackup", Err.Description
    CreateBackup = False
End Function

'@Description("Détermine les scripts à exécuter")
'@Description: 
'@Param: 
'@Returns: 

Private Function GetScriptsToRun(ByVal targetVersion As String) As Collection
    Dim result As New Collection
    
    ' Déterminer la direction
    Dim direction As String
    If CompareVersions(targetVersion, this.CurrentVersion) > 0 Then
        direction = "UP"
    Else
        direction = "DOWN"
    End If
    
    ' Filtrer les scripts
    Dim script As TMigrationScript
    For Each script In m_Scripts
        If direction = "UP" Then
            If CompareVersions(script.Version, this.CurrentVersion) > 0 And _
               CompareVersions(script.Version, targetVersion) <= 0 And _
               Not script.IsApplied Then
                result.Add script
            End If
        Else
            If CompareVersions(script.Version, this.CurrentVersion) <= 0 And _
               CompareVersions(script.Version, targetVersion) > 0 And _
               script.IsApplied Then
                result.Add script
            End If
        End If
    Next script
    
    Set GetScriptsToRun = result
End Function

'@Description("Compare deux numéros de version")
'@Description: 
'@Param: 
'@Returns: 

Private Function CompareVersions(ByVal version1 As String, ByVal version2 As String) As Long
    Dim parts1() As String
    Dim parts2() As String
    
    parts1 = Split(version1, ".")
    parts2 = Split(version2, ".")
    
    Dim i As Long
    For i = 0 To UBound(parts1)
        If i > UBound(parts2) Then
            CompareVersions = 1
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
        
        If CLng(parts1(i)) > CLng(parts2(i)) Then
            CompareVersions = 1
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        ElseIf CLng(parts1(i)) < CLng(parts2(i)) Then
            CompareVersions = -1
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next i
    
    If UBound(parts2) > UBound(parts1) Then
        CompareVersions = -1
    Else
        CompareVersions = 0
    End If
End Function

'@Description("Exécute un script de migration")
'@Description: 
'@Param: 
'@Returns: 

Private Function ExecuteScript(ByRef script As TMigrationScript) As Boolean
    On Error GoTo ErrorHandler
    
    ' Lire le contenu du script
    Dim content As String
    content = ReadTextFile(script.Path)
    
    If Len(content) = 0 Then
        LogError "ExecuteScript", "Script vide: " & script.Path
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Exécuter les commandes
    Dim commands() As String
    commands = Split(content, "GO")
    
    Dim i As Long
    For i = 0 To UBound(commands)
        If Len(Trim$(commands(i))) > 0 Then
            ' TODO: Exécuter la commande selon le SGBD utilisé
            ' Exemple :
            ' connection.Execute commands(i)
        End If
    Next i
    
    ' Marquer comme appliqué
    script.IsApplied = True
    
    ' Mettre à jour la version courante
    this.CurrentVersion = script.Version
    
    LogInfo "Script exécuté: " & script.Name
    ExecuteScript = True
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "ExecuteScript", Err.Description
    ExecuteScript = False
End Function

'@Description("Met à jour les métriques de migration")
'@Description: 
'@Param: 
'@Returns: 

Private Sub UpdateMigrationMetrics(ByVal success As Boolean)
    ' Calculer la durée
    Dim duration As Double
    duration = DateDiff("s", this.StartTime, Now)
    
    ' Mettre à jour les statistiques
    m_Stats.TotalMigrations = m_Stats.TotalMigrations + 1
    If success Then
        m_Stats.SuccessfulMigrations = m_Stats.SuccessfulMigrations + 1
    Else
        m_Stats.FailedMigrations = m_Stats.FailedMigrations + 1
    End If
    
    ' Moyenne mobile
    m_Stats.AverageMigrationTime = (m_Stats.AverageMigrationTime * (m_Stats.TotalMigrations - 1) + duration) / m_Stats.TotalMigrations
    
    ' Plus longue migration
    If duration > m_Stats.LongestMigration Then
        m_Stats.LongestMigration = duration
    End If
    
    m_Stats.LastMigrationDate = Now
    
    ' Métriques
    If Not m_MetricsCollector Is Nothing Then
        m_MetricsCollector.RecordMetric "MigrationDuration", duration
        m_MetricsCollector.RecordMetric "MigrationSuccess", IIf(success, 1, 0)
    End If
End Sub

'@Description("Obtient les statistiques des migrations")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetStats() As String
    Dim successRate As Double
    If m_Stats.TotalMigrations > 0 Then
        successRate = m_Stats.SuccessfulMigrations / m_Stats.TotalMigrations
    End If
    
    GetStats = "=== Statistiques Migrations ===" & vbNewLine & _
               "Total: " & m_Stats.TotalMigrations & vbNewLine & _
               "Réussies: " & m_Stats.SuccessfulMigrations & vbNewLine & _
               "Échouées: " & m_Stats.FailedMigrations & vbNewLine & _
               "Taux de succès: " & Format(successRate, "0.0%") & vbNewLine & _
               "Temps moyen: " & Format(m_Stats.AverageMigrationTime, "0.000") & " s" & vbNewLine & _
               "Plus longue: " & Format(m_Stats.LongestMigration, "0.000") & " s" & vbNewLine & _
               "Dernière: " & Format(m_Stats.LastMigrationDate, "yyyy-mm-dd hh:nn:ss")
End Function

'@Description("Vérifie si un dossier existe")
'@Description: 
'@Param: 
'@Returns: 

Private Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(path) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

'@Description("Crée un dossier")
'@Description: 
'@Param: 
'@Returns: 

Private Sub CreateFolder(ByVal path As String)
    On Error Resume Next
    MkDir path
    On Error GoTo 0
End Sub

'@Description("Lit un fichier texte")
'@Description: 
'@Param: 
'@Returns: 

Private Function ReadTextFile(ByVal path As String) As String
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1) ' ForReading
    
    ReadTextFile = ts.ReadAll
    ts.Close
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    LogError "ReadTextFile", Err.Description
    ReadTextFile = ""
End Function

'@Description("Journalise une erreur")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogError(ByVal method As String, ByVal description As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log "ERROR", "MigrationManager." & method, description
    End If
End Sub

'@Description("Journalise une information")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogInfo(ByVal message As String)
    If Not m_Logger Is Nothing Then
        m_Logger.Log "INFO", "MigrationManager", message
    End If
End Sub 