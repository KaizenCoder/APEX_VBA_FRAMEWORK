VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockLogger"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("APEX.Tests.Mocks")
Option Explicit

Implements ILoggerBase

Private Type TLogEntry
    Level As String
    Source As String
    Message As String
    Timestamp As Date
End Type

Private m_Logs As Collection

Private Sub Class_Initialize()
    Set m_Logs = New Collection
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Class_Terminate()
    Set m_Logs = Nothing
End Sub

'@Description("Enregistre un message dans le journal")
'@Description: 
'@Param: 
'@Returns: 

Private Sub ILoggerBase_Log(ByVal level As String, ByVal source As String, ByVal message As String)
    Dim entry As TLogEntry
    
    With entry
        .Level = level
        .Source = source
        .Message = message
        .Timestamp = Now
    End With
    
    m_Logs.Add entry
End Sub

'@Description("Obtient tous les messages du journal")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetLogs() As Collection
    Set GetLogs = m_Logs
End Function

'@Description("Obtient les messages d'un certain niveau")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetLogsByLevel(ByVal level As String) As Collection
    Dim result As New Collection
    Dim entry As TLogEntry
    Dim var As Variant
    
    For Each var In m_Logs
        entry = var
        If entry.Level = level Then
            result.Add entry
        End If
    Next var
    
    Set GetLogsByLevel = result
End Function

'@Description("Obtient les messages d'une certaine source")
'@Description: 
'@Param: 
'@Returns: 

Public Function GetLogsBySource(ByVal source As String) As Collection
    Dim result As New Collection
    Dim entry As TLogEntry
    Dim var As Variant
    
    For Each var In m_Logs
        entry = var
        If entry.Source = source Then
            result.Add entry
        End If
    Next var
    
    Set GetLogsBySource = result
End Function

'@Description("Vérifie si un message existe dans les logs")
'@Description: 
'@Param: 
'@Returns: 

Public Function ContainsMessage(ByVal message As String) As Boolean
    Dim entry As TLogEntry
    Dim var As Variant
    
    For Each var In m_Logs
        entry = var
        If InStr(1, entry.Message, message) > 0 Then
            ContainsMessage = True
            Exit'@Description: 
'@Param: 
'@Returns: 

 Function
        End If
    Next var
    
    ContainsMessage = False
End Function

'@Description("Efface tous les logs")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Clear()
    Set m_Logs = New Collection
End Sub 