' Migrated to apex-core/interfaces - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

' ==========================================================================
' Interface : ILoggerBase
' Version : 3.0
' Purpose : Defines the standard contract for logger implementations.
' ==========================================================================

' --- Initialization & Configuration ---
Public Sub Initialize(Optional ByVal minLevel As LogLevelEnum = LogLevelInfo, Optional ByVal logSheetName As String = "Logs", Optional ByVal logFileNamePattern As String = "{WorkbookName}\_{Date}.log", Optional ByVal maxLogFileSizeKB As Long = 5120, Optional ByVal targetWorkbook As Workbook = Nothing, Optional ByVal enabledCategories As String = "\*", Optional ByVal disabledCategories As String = "", Optional ByVal bufferSize As Long = 1, Optional ByVal crashLogBufferSize As Long = 10)
Err.Raise vbObjectError + 1001, "ILoggerBase", "Initialize method not implemented."
End Sub
Public Sub SetLogger(ByVal loggerInstance As ILoggerBase) ' Usually NOOP for logger itself
Err.Raise vbObjectError + 1001, "ILoggerBase", "SetLogger method not implemented."
End Sub

' --- Logging Methods ---
Public Sub LogMessage(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = False, Optional ByVal toFile As Boolean = True)
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogMessage method not implemented."
End Sub
Public Sub LogConsole(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogConsole method not implemented."
End Sub
Public Sub LogSheet(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogSheet method not implemented."
End Sub
Public Sub LogFile(ByVal msg As String, Optional ByVal level As LogLevelEnum = LogLevelInfo, Optional ByVal category As String = "", Optional ByVal source As String = "", Optional ByVal user As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogFile method not implemented."
End Sub
Public Sub LogError(ByVal errObject As ErrObject, Optional ByVal level As LogLevelEnum = LogLevelError, Optional ByVal sourceRoutine As String = "", Optional ByVal category As String = "ERROR", Optional ByVal user As String = "", Optional ByVal toConsole As Boolean = True, Optional ByVal toSheet As Boolean = True, Optional ByVal toFile As Boolean = True)
Err.Raise vbObjectError + 1001, "ILoggerBase", "LogError method not implemented."
End Sub

' --- Control Methods ---
Public Sub FlushLogs()
Err.Raise vbObjectError + 1001, "ILoggerBase", "FlushLogs method not implemented."
End Sub
Public Sub GenerateCrashReport(Optional ByVal crashFilePath As String = "")
Err.Raise vbObjectError + 1001, "ILoggerBase", "GenerateCrashReport method not implemented."
End Sub

' --- Read-only Properties ---
Public Property Get MinLogLevel() As LogLevelEnum
Err.Raise vbObjectError + 1001, "ILoggerBase", "MinLogLevel property not implemented."
End Property
Public Property Get IsInitialized() As Boolean
Err.Raise vbObjectError + 1001, "ILoggerBase", "IsInitialized property not implemented."
End Property

' --- Log Level Enum (Defined here for interface context) ---
Public Enum LogLevelEnum
LogLevelDebug = 10
LogLevelInfo = 20
LogLevelWarning = 30
LogLevelError = 40
LogLevelFatal = 50
End Enum 