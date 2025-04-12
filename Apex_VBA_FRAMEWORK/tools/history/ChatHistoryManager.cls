Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team


'@Folder("Tools.History")
'@ModuleDescription("Gestionnaire d'historisation des chats")

Private Type THistoryManager
    LogPath As String
    ArchivePath As String
End Type

Private this As THistoryManager

'@Description("Initialise le gestionnaire")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize()
    this.LogPath = "tools/workflow/sessions/"
    this.ArchivePath = "tools/workflow/archive/"
End Sub

'@Description("Archive les sessions de chat spécifiées")
'@Description: 
'@Param: 
'@Returns: 

Public Sub ArchiveChatSessions(ByVal startSession As Long, ByVal endSession As Long)
    On Error GoTo ErrorHandler
    
    'Validation des paramètres
    If startSession > endSession Then
        Err.Raise 5, "ArchiveChatSessions", "La session de début doit être inférieure à la session de fin"
    End If
    
    'Configuration de l'archivage
    Dim archiveDate As String
    archiveDate = Format(Now, "YYYY_MM_DD")
    
    'Création du dossier d'archive si nécessaire
    CreateArchiveFolder archiveDate
    
    'Archivage des sessions
    Dim sessionNum As Long
    For sessionNum = startSession To endSession
        ArchiveSingleSession sessionNum, archiveDate
    Next sessionNum
    
    'Génération du rapport d'archivage
    GenerateArchiveReport startSession, endSession, archiveDate
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
ErrorHandler:
    Err.Raise Err.Number, "ArchiveChatSessions", "Erreur lors de l'archivage: " & Err.Description
End Sub

'@Description("Archive une session unique")
'@Description: 
'@Param: 
'@Returns: 

Private Sub ArchiveSingleSession(ByVal sessionNum As Long, ByVal archiveDate As String)
    Dim sessionFile As String
    sessionFile = this.LogPath & Format(sessionNum, "000") & ".md"
    
    If Not FileExists(sessionFile) Then Exit Sub
    
    'Lecture et traitement du contenu
    Dim content As String
    content = ReadFileContent(sessionFile)
    
    'Ajout des métadonnées d'archivage
    content = AddArchiveMetadata(content, sessionNum, archiveDate)
    
    'Sauvegarde dans l'archive
    Dim archiveFile As String
    archiveFile = this.ArchivePath & archiveDate & "/session_" & Format(sessionNum, "000") & ".md"
    
    SaveFileContent archiveFile, content
End Sub

'@Description("Ajoute les métadonnées d'archivage")
'@Description: 
'@Param: 
'@Returns: 

Private Function AddArchiveMetadata(ByVal content As String, ByVal sessionNum As Long, ByVal archiveDate As String) As String
    Dim metadata As String
    metadata = "---" & vbNewLine & _
              "Archive_Date: " & archiveDate & vbNewLine & _
              "Session_Number: " & Format(sessionNum, "000") & vbNewLine & _
              "Archive_Type: Chat_History" & vbNewLine & _
              "---" & vbNewLine & vbNewLine
    
    AddArchiveMetadata = metadata & content
End Function

'@Description("Génère le rapport d'archivage")
'@Description: 
'@Param: 
'@Returns: 

Private Sub GenerateArchiveReport(ByVal startSession As Long, ByVal endSession As Long, ByVal archiveDate As String)
    Dim reportFile As String
    reportFile = this.ArchivePath & archiveDate & "/archive_report.md"
    
    Dim report As String
    report = "# Rapport d'Archivage des Sessions" & vbNewLine & _
            "Date: " & archiveDate & vbNewLine & _
            "Sessions: " & startSession & " à " & endSession & vbNewLine & _
            "Statut: Complété" & vbNewLine
    
    SaveFileContent reportFile, report
End Sub 