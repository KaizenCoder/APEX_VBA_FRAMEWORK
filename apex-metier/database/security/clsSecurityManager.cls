VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSecurityManager"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Database.Security")
Option Explicit

'==========================================================================
' Classe    : clsSecurityManager
' Purpose   : Gestion de la sécurité pour l'accès aux données
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : SEC-001
'==========================================================================

Private Type TSecuritySettings
    EnableParameterization As Boolean
    ValidateInputs As Boolean
    SanitizeOutputs As Boolean
    UseSecureConnections As Boolean
    EncryptionKey As String
    MaxRetries As Long
    LockoutThreshold As Long
    PasswordPolicy As String
    AuditLevel As String
End Type

Private this As TSecuritySettings
Private m_Logger As ILoggerBase

'@Description("Initialise le gestionnaire de sécurité")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize(Optional ByVal logger As ILoggerBase = Nothing) As Boolean
    If Not logger Is Nothing Then Set m_Logger = logger
    
    ' Configuration par défaut sécurisée
    With this
        .EnableParameterization = True
        .ValidateInputs = True
        .SanitizeOutputs = True
        .UseSecureConnections = True
        .MaxRetries = 3
        .LockoutThreshold = 5
        .PasswordPolicy = "STRONG"
        .AuditLevel = "HIGH"
    End With
    
    Initialize = True
End Function

'@Description("Chiffre une chaîne de connexion")
'@Description: 
'@Param: 
'@Returns: 

Public Function EncryptConnectionString(ByVal connectionString As String) As String
    If Len(this.EncryptionKey) = 0 Then
        GenerateEncryptionKey
    End If
    
    ' Chiffrement AES-256
    EncryptConnectionString = modCrypto.EncryptAES256( _
        plainText:=connectionString, _
        key:=this.EncryptionKey _
    )
End Function

'@Description("Déchiffre une chaîne de connexion")
'@Description: 
'@Param: 
'@Returns: 

Public Function DecryptConnectionString(ByVal encryptedString As String) As String
    If Len(this.EncryptionKey) = 0 Then
        Err.Raise 5, "clsSecurityManager", "Clé de chiffrement non initialisée"
    End If
    
    ' Déchiffrement AES-256
    DecryptConnectionString = modCrypto.DecryptAES256( _
        cipherText:=encryptedString, _
        key:=this.EncryptionKey _
    )
End Function

'@Description("Valide et nettoie une entrée SQL")
'@Description: 
'@Param: 
'@Returns: 

Public Function SanitizeSqlInput(ByVal input As String) As String
    If Not this.ValidateInputs Then
        SanitizeSqlInput = input
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Nettoyage basique
    Dim sanitized As String
    sanitized = Replace(input, "'", "''")
    sanitized = Replace(sanitized, ";", "")
    sanitized = Replace(sanitized, "--", "")
    sanitized = Replace(sanitized, "/*", "")
    sanitized = Replace(sanitized, "*/", "")
    
    ' Validation supplémentaire
    If ContainsSqlInjection(sanitized) Then
        LogSecurityEvent "Tentative d'injection SQL détectée", "HIGH"
        Err.Raise 5, "clsSecurityManager", "Entrée SQL non valide détectée"
    End If
    
    SanitizeSqlInput = sanitized
End Function

'@Description("Vérifie si une chaîne contient une tentative d'injection SQL")
'@Description: 
'@Param: 
'@Returns: 

Private Function ContainsSqlInjection(ByVal input As String) As Boolean
    ' Liste de mots-clés SQL dangereux
    Const DANGEROUS_KEYWORDS As String = "UNION|SELECT|INSERT|UPDATE|DELETE|DROP|TRUNCATE|ALTER|EXEC|EXECUTE"
    
    ' Vérification des mots-clés
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = DANGEROUS_KEYWORDS
        .IgnoreCase = True
        .Global = True
        ContainsSqlInjection = .Test(input)
    End With
End Function

'@Description("Génère une nouvelle clé de chiffrement")
'@Description: 
'@Param: 
'@Returns: 

Private Sub GenerateEncryptionKey()
    ' Génération d'une clé AES-256 aléatoire
    this.EncryptionKey = modCrypto.GenerateRandomKey(32)
End Sub

'@Description("Journalise un événement de sécurité")
'@Description: 
'@Param: 
'@Returns: 

Private Sub LogSecurityEvent(ByVal message As String, ByVal level As String)
    If m_Logger Is Nothing Then Exit'@Description: 
'@Param: 
'@Returns: 

 Sub
    
    m_Logger.LogEvent _
        component:="SecurityManager", _
        eventType:="SECURITY", _
        message:=message, _
        level:=level
End Sub

' Propriétés de configuration
'@Description("Active/désactive la paramétrisation des requêtes")
Public Property Let EnableParameterization(ByVal value As Boolean)
    this.EnableParameterization = value
End Property

'@Description("Active/désactive la validation des entrées")
Public Property Let ValidateInputs(ByVal value As Boolean)
    this.ValidateInputs = value
End Property

'@Description("Active/désactive le nettoyage des sorties")
Public Property Let SanitizeOutputs(ByVal value As Boolean)
    this.SanitizeOutputs = value
End Property

'@Description("Active/désactive l'utilisation de connexions sécurisées")
Public Property Let UseSecureConnections(ByVal value As Boolean)
    this.UseSecureConnections = value
End Property

'@Description("Définit le niveau d'audit")
Public Property Let AuditLevel(ByVal value As String)
    this.AuditLevel = UCase$(value)
End Property

Private Sub Class_Terminate()
    Set m_Logger = Nothing
End Sub 