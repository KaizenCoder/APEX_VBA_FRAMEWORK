' Migrated to apex-core/utils - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Module : modDateUtils
' Version : 1.0
' Purpose : Utilitaires pour la manipulation et le formatage des dates
' Date : 10/04/2025
' ==========================================================================

' --- Variables privées ---
Private m_Logger As Object ' ILoggerBase
Private m_LastError As String

' --- Constantes ---
Private Const DATE_FORMAT_FR As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_US As String = "mm/dd/yyyy"
Private Const DATE_FORMAT_ISO As String = "yyyy-mm-dd"
Private Const DATE_FORMAT_SQL As String = "yyyy-mm-dd hh:nn:ss"

' --- Initialisation ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize(Optional ByVal logger As Object = Nothing)
    ' Initialise le module avec un logger optionnel
    Set m_Logger = logger
    m_LastError = ""
    
    If Not m_Logger Is Nothing Then
        ' TODO: Logger l'initialisation
        ' m_Logger.LogInfo "modDateUtils initialisé", "DATE"
    End If
End Sub

' --- Fonctions de formatage de date ---
'@Description: 
'@Param: 
'@Returns: 

Public Function FormatDate(ByVal dateValue As Date, Optional ByVal format As String = DATE_FORMAT_FR) As String
    ' Formate une date selon un format spécifié
    On Error GoTo ErrorHandler
    
    FormatDate = Format(dateValue, format)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du formatage de la date: " & Err.Description
    LogError m_LastError
    FormatDate = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetISODate(ByVal dateValue As Date) As String
    ' Retourne une date au format ISO (YYYY-MM-DD)
    GetISODate = FormatDate(dateValue, DATE_FORMAT_ISO)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetSQLDate(ByVal dateValue As Date) As String
    ' Retourne une date au format SQL (YYYY-MM-DD HH:NN:SS)
    GetSQLDate = FormatDate(dateValue, DATE_FORMAT_SQL)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetLocalizedDate(ByVal dateValue As Date, Optional ByVal locale As String = "FR") As String
    ' Retourne une date formatée selon la locale spécifiée
    On Error GoTo ErrorHandler
    
    Select Case UCase(locale)
        Case "FR"
            GetLocalizedDate = FormatDate(dateValue, DATE_FORMAT_FR)
        Case "US"
            GetLocalizedDate = FormatDate(dateValue, DATE_FORMAT_US)
        Case "ISO"
            GetLocalizedDate = FormatDate(dateValue, DATE_FORMAT_ISO)
        Case Else
            GetLocalizedDate = FormatDate(dateValue, DATE_FORMAT_FR)
    End Select
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du formatage de la date localisée: " & Err.Description
    LogError m_LastError
    GetLocalizedDate = ""
End Function

' --- Fonctions de conversion de date ---
'@Description: 
'@Param: 
'@Returns: 

Public Function ParseDate(ByVal dateString As String, Optional ByVal format As String = "") As Date
    ' Convertit une chaîne en date selon un format spécifié
    On Error GoTo ErrorHandler
    
    ' Si le format n'est pas spécifié, essayer de déterminer automatiquement
    If format = "" Then
        ' Tenter de détecter le format
        If InStr(dateString, "-") > 0 Then
            ' Probablement ISO ou similaire
            If Len(dateString) > 10 Then
                format = DATE_FORMAT_SQL
            Else
                format = DATE_FORMAT_ISO
            End If
        ElseIf InStr(dateString, "/") > 0 Then
            ' Probablement FR ou US
            ' Tenter de déterminer en fonction de la position des séparateurs
            Dim parts() As String
            parts = Split(dateString, "/")
            
            If UBound(parts) >= 2 Then
                If Len(parts(0)) = 2 And Val(parts(0)) <= 31 Then
                    format = DATE_FORMAT_FR
                ElseIf Len(parts(0)) = 2 And Val(parts(0)) <= 12 Then
                    format = DATE_FORMAT_US
                Else
                    format = DATE_FORMAT_FR ' Par défaut
                End If
            Else
                format = DATE_FORMAT_FR ' Par défaut
            End If
        Else
            ' Format inconnu, utiliser la conversion par défaut de VBA
            ParseDate = CDate(dateString)
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    End If
    
    ' Convertir en fonction du format
    Select Case format
        Case DATE_FORMAT_ISO
            ' Format ISO: YYYY-MM-DD
            Dim isoparts() As String
            isoparts = Split(dateString, "-")
            
            If UBound(isoparts) >= 2 Then
                ParseDate = DateSerial(Val(isoparts(0)), Val(isoparts(1)), Val(isoparts(2)))
            Else
                ParseDate = CDate(dateString)
            End If
            
        Case DATE_FORMAT_SQL
            ' Format SQL: YYYY-MM-DD HH:NN:SS
            Dim sqldateparts() As String
            sqldateparts = Split(dateString, " ")
            
            If UBound(sqldateparts) >= 1 Then
                ' Partie date
                Dim sqldateisoparts() As String
                sqldateisoparts = Split(sqldateparts(0), "-")
                
                ' Partie heure
                Dim sqltimeparts() As String
                sqltimeparts = Split(sqldateparts(1), ":")
                
                If UBound(sqldateisoparts) >= 2 And UBound(sqltimeparts) >= 2 Then
                    ParseDate = DateSerial(Val(sqldateisoparts(0)), Val(sqldateisoparts(1)), Val(sqldateisoparts(2))) + _
                               TimeSerial(Val(sqltimeparts(0)), Val(sqltimeparts(1)), Val(sqltimeparts(2)))
                Else
                    ParseDate = CDate(dateString)
                End If
            Else
                ParseDate = CDate(dateString)
            End If
            
        Case DATE_FORMAT_FR
            ' Format FR: JJ/MM/AAAA
            Dim frparts() As String
            frparts = Split(dateString, "/")
            
            If UBound(frparts) >= 2 Then
                ParseDate = DateSerial(Val(frparts(2)), Val(frparts(1)), Val(frparts(0)))
            Else
                ParseDate = CDate(dateString)
            End If
            
        Case DATE_FORMAT_US
            ' Format US: MM/DD/YYYY
            Dim usparts() As String
            usparts = Split(dateString, "/")
            
            If UBound(usparts) >= 2 Then
                ParseDate = DateSerial(Val(usparts(2)), Val(usparts(0)), Val(usparts(1)))
            Else
                ParseDate = CDate(dateString)
            End If
            
        Case Else
            ' Utiliser la conversion par défaut de VBA
            ParseDate = CDate(dateString)
    End Select
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la conversion de la chaîne en date: " & Err.Description
    LogError m_LastError
    ' Retourner une date nulle en cas d'erreur
    ParseDate = DateSerial(1900, 1, 1)
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function IsValidDate(ByVal dateString As String, Optional ByVal format As String = "") As Boolean
    ' Vérifie si une chaîne représente une date valide
    On Error GoTo ErrorHandler
    
    Dim testDate As Date
    testDate = ParseDate(dateString, format)
    
    ' Vérifie si la date est dans une plage raisonnable
    IsValidDate = (Year(testDate) >= 1900 And Year(testDate) <= 2100)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la validation de la date: " & Err.Description
    LogError m_LastError
    IsValidDate = False
End Function

' --- Fonctions de calcul de date ---
'@Description: 
'@Param: 
'@Returns: 

Public Function AddWorkDays(ByVal startDate As Date, ByVal workDays As Long) As Date
    ' Ajoute un nombre de jours ouvrés à une date
    Dim currentDate As Date
    Dim remainingDays As Long
    
    On Error GoTo ErrorHandler
    
    currentDate = startDate
    remainingDays = workDays
    
    ' Si on ajoute 0 jours, retourner la date de départ
    If remainingDays = 0 Then
        AddWorkDays = startDate
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Ajouter les jours un par un
    Do While remainingDays <> 0
        ' Ajouter un jour
        If remainingDays > 0 Then
            currentDate = currentDate + 1
            remainingDays = remainingDays - 1
        Else
            currentDate = currentDate - 1
            remainingDays = remainingDays + 1
        End If
        
        ' Si c'est un weekend, ne pas compter le jour
        If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
            If remainingDays > 0 Then
                remainingDays = remainingDays + 1
            Else
                remainingDays = remainingDays - 1
            End If
        End If
    Loop
    
    AddWorkDays = currentDate
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'ajout des jours ouvrés: " & Err.Description
    LogError m_LastError
    AddWorkDays = startDate
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetWorkDaysBetween(ByVal startDate As Date, ByVal endDate As Date) As Long
    ' Calcule le nombre de jours ouvrés entre deux dates
    Dim currentDate As Date
    Dim workDays As Long
    
    On Error GoTo ErrorHandler
    
    ' S'assurer que startDate est inférieure à endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    workDays = 0
    currentDate = startDate
    
    ' Compter les jours ouvrés
    Do While currentDate <= endDate
        If Weekday(currentDate) <> vbSaturday And Weekday(currentDate) <> vbSunday Then
            workDays = workDays + 1
        End If
        
        currentDate = currentDate + 1
    Loop
    
    GetWorkDaysBetween = workDays
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul des jours ouvrés: " & Err.Description
    LogError m_LastError
    GetWorkDaysBetween = 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetQuarter(ByVal dateValue As Date) As Integer
    ' Retourne le trimestre (1-4) pour une date donnée
    On Error GoTo ErrorHandler
    
    GetQuarter = Int((Month(dateValue) - 1) / 3) + 1
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du trimestre: " & Err.Description
    LogError m_LastError
    GetQuarter = 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetFirstDayOfMonth(ByVal dateValue As Date) As Date
    ' Retourne le premier jour du mois pour une date donnée
    On Error GoTo ErrorHandler
    
    GetFirstDayOfMonth = DateSerial(Year(dateValue), Month(dateValue), 1)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du premier jour du mois: " & Err.Description
    LogError m_LastError
    GetFirstDayOfMonth = dateValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetLastDayOfMonth(ByVal dateValue As Date) As Date
    ' Retourne le dernier jour du mois pour une date donnée
    On Error GoTo ErrorHandler
    
    GetLastDayOfMonth = DateSerial(Year(dateValue), Month(dateValue) + 1, 0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du dernier jour du mois: " & Err.Description
    LogError m_LastError
    GetLastDayOfMonth = dateValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetFirstDayOfQuarter(ByVal dateValue As Date) As Date
    ' Retourne le premier jour du trimestre pour une date donnée
    Dim quarter As Integer
    
    On Error GoTo ErrorHandler
    
    quarter = GetQuarter(dateValue)
    GetFirstDayOfQuarter = DateSerial(Year(dateValue), (quarter - 1) * 3 + 1, 1)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du premier jour du trimestre: " & Err.Description
    LogError m_LastError
    GetFirstDayOfQuarter = dateValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetLastDayOfQuarter(ByVal dateValue As Date) As Date
    ' Retourne le dernier jour du trimestre pour une date donnée
    Dim quarter As Integer
    
    On Error GoTo ErrorHandler
    
    quarter = GetQuarter(dateValue)
    GetLastDayOfQuarter = DateSerial(Year(dateValue), quarter * 3 + 1, 0)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du dernier jour du trimestre: " & Err.Description
    LogError m_LastError
    GetLastDayOfQuarter = dateValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetFirstDayOfYear(ByVal dateValue As Date) As Date
    ' Retourne le premier jour de l'année pour une date donnée
    On Error GoTo ErrorHandler
    
    GetFirstDayOfYear = DateSerial(Year(dateValue), 1, 1)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du premier jour de l'année: " & Err.Description
    LogError m_LastError
    GetFirstDayOfYear = dateValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetLastDayOfYear(ByVal dateValue As Date) As Date
    ' Retourne le dernier jour de l'année pour une date donnée
    On Error GoTo ErrorHandler
    
    GetLastDayOfYear = DateSerial(Year(dateValue), 12, 31)
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul du dernier jour de l'année: " & Err.Description
    LogError m_LastError
    GetLastDayOfYear = dateValue
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function GetAge(ByVal birthDate As Date, Optional ByVal referenceDate As Date = 0) As Integer
    ' Calcule l'âge en années à partir d'une date de naissance
    On Error GoTo ErrorHandler
    
    ' Si aucune date de référence n'est spécifiée, utiliser la date du jour
    If referenceDate = 0 Then referenceDate = Date
    
    ' Calcul simple de l'âge
    Dim age As Integer
    age = Year(referenceDate) - Year(birthDate)
    
    ' Ajustement si l'anniversaire n'est pas encore passé cette année
    If Month(referenceDate) < Month(birthDate) Or _
       (Month(referenceDate) = Month(birthDate) And Day(referenceDate) < Day(birthDate)) Then
        age = age - 1
    End If
    
    GetAge = age
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du calcul de l'âge: " & Err.Description
    LogError m_LastError
    GetAge = 0
End Function

' --- Fonctions temporelles ---
'@Description: 
'@Param: 
'@Returns: 

Public Function SecondsToTime(ByVal seconds As Long) As String
    ' Convertit un nombre de secondes en format hh:mm:ss
    On Error GoTo ErrorHandler
    
    Dim hours As Long
    Dim minutes As Long
    Dim remainingSeconds As Long
    
    hours = seconds \ 3600
    minutes = (seconds Mod 3600) \ 60
    remainingSeconds = seconds Mod 60
    
    SecondsToTime = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00")
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la conversion des secondes en temps: " & Err.Description
    LogError m_LastError
    SecondsToTime = "00:00:00"
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function TimeToSeconds(ByVal timeString As String) As Long
    ' Convertit un format hh:mm:ss en nombre de secondes
    On Error GoTo ErrorHandler
    
    Dim timeParts() As String
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    
    ' Diviser la chaîne de temps
    timeParts = Split(timeString, ":")
    
    ' Calculer le nombre de secondes
    If UBound(timeParts) >= 2 Then
        hours = Val(timeParts(0))
        minutes = Val(timeParts(1))
        seconds = Val(timeParts(2))
        
        TimeToSeconds = hours * 3600 + minutes * 60 + seconds
    ElseIf UBound(timeParts) = 1 Then
        minutes = Val(timeParts(0))
        seconds = Val(timeParts(1))
        
        TimeToSeconds = minutes * 60 + seconds
    Else
        seconds = Val(timeParts(0))
        
        TimeToSeconds = seconds
    End If
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de la conversion du temps en secondes: " & Err.Description
    LogError m_LastError
    TimeToSeconds = 0
End Function

' --- Propriétés ---
Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

' --- Fonctions privées ---
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "DATE"
    End If
End Sub
