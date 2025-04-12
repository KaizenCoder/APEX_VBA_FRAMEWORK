' Migrated to apex-core - 2025-04-09
' Part of the APEX Framework v1.1 architecture refactoring
Option Explicit

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

' ==========================================================================
' Module : modSecurityDPAPI
' Version : 1.0
' Purpose : Sécurisation des données sensibles via Windows DPAPI
' Date : 10/04/2025
' ==========================================================================

' --- API Windows pour le chiffrement DPAPI ---
#If VBA7 Then
    ' Pour VBA 64 bits (Office 64 bits)
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function CryptProtectData Lib "crypt32.dll" (pDataIn As DATA_BLOB, _
        ByVal szDataDescr As LongPtr, pOptionalEntropy As DATA_BLOB, _
        ByVal pvReserved As LongPtr, pPromptStruct As Any, _
        ByVal dwFlags As Long, pDataOut As DATA_BLOB) As Long
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function CryptUnprotectData Lib "crypt32.dll" (pDataIn As DATA_BLOB, _
        ByVal ppszDataDescr As LongPtr, pOptionalEntropy As DATA_BLOB, _
        ByVal pvReserved As LongPtr, pPromptStruct As Any, _
        ByVal dwFlags As Long, pDataOut As DATA_BLOB) As Long
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As LongPtr)
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, _
        ByVal uBytes As LongPtr) As LongPtr
    
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
#Else
    ' Pour VBA 32 bits (Office 32 bits)
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function CryptProtectData Lib "crypt32.dll" (pDataIn As DATA_BLOB, _
        ByVal szDataDescr As Long, pOptionalEntropy As DATA_BLOB, _
        ByVal pvReserved As Long, pPromptStruct As Any, _
        ByVal dwFlags As Long, pDataOut As DATA_BLOB) As Long
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function CryptUnprotectData Lib "crypt32.dll" (pDataIn As DATA_BLOB, _
        ByVal ppszDataDescr As Long, pOptionalEntropy As DATA_BLOB, _
        ByVal pvReserved As Long, pPromptStruct As Any, _
        ByVal dwFlags As Long, pDataOut As DATA_BLOB) As Long
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, _
        ByVal uBytes As Long) As Long
    
    Private Declare'@Description: 
'@Param: 
'@Returns: 

 Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
#End If

' --- Constantes pour l'API DPAPI ---
Private Const CRYPTPROTECT_UI_FORBIDDEN As Long = &H1
Private Const CRYPTPROTECT_LOCAL_MACHINE As Long = &H4
Private Const LMEM_FIXED As Long = &H0
Private Const LMEM_ZEROINIT As Long = &H40

' --- Structure DATA_BLOB pour DPAPI ---
Private Type DATA_BLOB
    cbData As Long
    #If VBA7 Then
        pbData As LongPtr
    #Else
        pbData As Long
    #End If
End Type

' --- Variables privées ---
Private m_Logger As Object ' ILoggerBase
Private m_Entropy As String ' Salt optionnel pour renforcer le chiffrement
Private m_LastError As String ' Dernière erreur

' --- Initialisation ---
'@Description: 
'@Param: 
'@Returns: 

Public Sub Initialize(Optional ByVal logger As Object = Nothing, Optional ByVal entropy As String = "")
    ' Initialise le module avec un logger et une entropie optionnels
    Set m_Logger = logger
    m_Entropy = entropy
    m_LastError = ""
    
    If Not m_Logger Is Nothing Then
        ' TODO: Logger l'initialisation
        ' m_Logger.LogInfo "modSecurityDPAPI initialisé", "SECURITY"
    End If
End Sub

' --- Fonctions publiques ---
'@Description: 
'@Param: 
'@Returns: 

Public Function EncryptString(ByVal plainText As String) As Byte()
    ' Chiffre une chaîne en utilisant DPAPI
    ' Retourne un tableau de bytes chiffrés
    
    Dim dataIn As DATA_BLOB
    Dim dataOut As DATA_BLOB
    Dim dataEntropy As DATA_BLOB
    Dim bytesIn() As Byte
    Dim bytesOut() As Byte
    Dim bytesEntropy() As Byte
    Dim result As Long
    
    On Error GoTo ErrorHandler
    
    ' Préparer les données à chiffrer
    bytesIn = StringToBytes(plainText)
    dataIn = CreateBlob(bytesIn)
    
    ' Préparer l'entropie si spécifiée
    If m_Entropy <> "" Then
        bytesEntropy = StringToBytes(m_Entropy)
        dataEntropy = CreateBlob(bytesEntropy)
    End If
    
    ' Appeler l'API Windows pour chiffrer
    result = CryptProtectData(dataIn, 0, dataEntropy, 0, ByVal 0, CRYPTPROTECT_UI_FORBIDDEN, dataOut)
    
    If result = 0 Then
        m_LastError = "Échec du chiffrement par DPAPI"
        LogError m_LastError
        EncryptString = vbNullString
        GoTo Cleanup
    End If
    
    ' Récupérer les données chiffrées
    ReDim bytesOut(0 To dataOut.cbData - 1)
    CopyMemory bytesOut(0), ByVal dataOut.pbData, dataOut.cbData
    
    ' Nettoyer et retourner
    EncryptString = bytesOut
    
Cleanup:
    ' Libérer la mémoire allouée
    LocalFree dataOut.pbData
    LocalFree dataIn.pbData
    If m_Entropy <> "" Then LocalFree dataEntropy.pbData
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du chiffrement: " & Err.Description
    LogError m_LastError
    EncryptString = vbNullString
    Resume Cleanup
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function DecryptString(ByVal encryptedBytes() As Byte) As String
    ' Déchiffre un tableau de bytes chiffrés avec DPAPI
    ' Retourne la chaîne déchiffrée
    
    Dim dataIn As DATA_BLOB
    Dim dataOut As DATA_BLOB
    Dim dataEntropy As DATA_BLOB
    Dim bytesOut() As Byte
    Dim bytesEntropy() As Byte
    Dim result As Long
    
    On Error GoTo ErrorHandler
    
    ' Préparer les données à déchiffrer
    dataIn = CreateBlob(encryptedBytes)
    
    ' Préparer l'entropie si spécifiée
    If m_Entropy <> "" Then
        bytesEntropy = StringToBytes(m_Entropy)
        dataEntropy = CreateBlob(bytesEntropy)
    End If
    
    ' Appeler l'API Windows pour déchiffrer
    result = CryptUnprotectData(dataIn, 0, dataEntropy, 0, ByVal 0, CRYPTPROTECT_UI_FORBIDDEN, dataOut)
    
    If result = 0 Then
        m_LastError = "Échec du déchiffrement par DPAPI"
        LogError m_LastError
        DecryptString = vbNullString
        GoTo Cleanup
    End If
    
    ' Récupérer les données déchiffrées
    ReDim bytesOut(0 To dataOut.cbData - 1)
    CopyMemory bytesOut(0), ByVal dataOut.pbData, dataOut.cbData
    
    ' Convertir en chaîne
    DecryptString = BytesToString(bytesOut)
    
Cleanup:
    ' Libérer la mémoire allouée
    LocalFree dataOut.pbData
    LocalFree dataIn.pbData
    If m_Entropy <> "" Then LocalFree dataEntropy.pbData
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du déchiffrement: " & Err.Description
    LogError m_LastError
    DecryptString = vbNullString
    Resume Cleanup
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function EncryptStringToBase64(ByVal plainText As String) As String
    ' Chiffre une chaîne et retourne le résultat encodé en Base64
    ' Pratique pour le stockage dans des fichiers texte
    
    Dim encryptedBytes() As Byte
    
    encryptedBytes = EncryptString(plainText)
    If Not IsArrayEmpty(encryptedBytes) Then
        EncryptStringToBase64 = BytesToBase64(encryptedBytes)
    Else
        EncryptStringToBase64 = ""
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Function DecryptStringFromBase64(ByVal base64Text As String) As String
    ' Déchiffre une chaîne préalablement chiffrée et encodée en Base64
    
    Dim encryptedBytes() As Byte
    
    encryptedBytes = Base64ToBytes(base64Text)
    If Not IsArrayEmpty(encryptedBytes) Then
        DecryptStringFromBase64 = DecryptString(encryptedBytes)
    Else
        DecryptStringFromBase64 = ""
    End If
End'@Description: 
'@Param: 
'@Returns: 

 Function

Public Property Get LastError() As String
    ' Retourne la dernière erreur survenue
    LastError = m_LastError
End Property

' --- Fonctions utilitaires privées ---
'@Description: 
'@Param: 
'@Returns: 

Private Function CreateBlob(ByRef bytes() As Byte) As DATA_BLOB
    ' Crée une structure DATA_BLOB à partir d'un tableau de bytes
    Dim blob As DATA_BLOB
    
    blob.cbData = UBound(bytes) - LBound(bytes) + 1
    blob.pbData = LocalAlloc(LMEM_FIXED Or LMEM_ZEROINIT, blob.cbData)
    CopyMemory ByVal blob.pbData, bytes(LBound(bytes)), blob.cbData
    
    CreateBlob = blob
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function StringToBytes(ByVal text As String) As Byte()
    ' Convertit une chaîne en tableau de bytes
    Dim bytes() As Byte
    bytes = text
    StringToBytes = bytes
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function BytesToString(ByRef bytes() As Byte) As String
    ' Convertit un tableau de bytes en chaîne
    Dim result As String
    
    If IsArrayEmpty(bytes) Then
        BytesToString = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Convertir en chaîne en s'assurant que la longueur est correcte
    result = Space$(UBound(bytes) - LBound(bytes) + 1)
    CopyMemory ByVal StrPtr(result), bytes(LBound(bytes)), Len(result)
    
    BytesToString = result
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function BytesToBase64(ByRef bytes() As Byte) As String
    ' Convertit un tableau de bytes en chaîne Base64
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    On Error GoTo ErrorHandler
    
    If IsArrayEmpty(bytes) Then
        BytesToBase64 = ""
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Utiliser XML pour l'encodage Base64
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("base64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = bytes
    BytesToBase64 = xmlNode.Text
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors de l'encodage Base64: " & Err.Description
    LogError m_LastError
    BytesToBase64 = ""
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function Base64ToBytes(ByVal base64Text As String) As Byte()
    ' Convertit une chaîne Base64 en tableau de bytes
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim bytes() As Byte
    
    On Error GoTo ErrorHandler
    
    If Trim(base64Text) = "" Then
        ReDim bytes(0 To 0)
        bytes(0) = 0
        Base64ToBytes = bytes
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Utiliser XML pour le décodage Base64
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("base64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = base64Text
    Base64ToBytes = xmlNode.nodeTypedValue
    
    Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    
ErrorHandler:
    m_LastError = "Erreur lors du décodage Base64: " & Err.Description
    LogError m_LastError
    ReDim bytes(0 To 0)
    bytes(0) = 0
    Base64ToBytes = bytes
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Function IsArrayEmpty(ByRef arr As Variant) As Boolean
    ' Vérifie si un tableau est vide ou non initialisé
    On Error Resume Next
    IsArrayEmpty = (UBound(arr) < LBound(arr))
    On Error GoTo 0
End'@Description: 
'@Param: 
'@Returns: 

 Function

Private Sub LogError(ByVal errorMessage As String)
    ' Log les erreurs si un logger est disponible
    If Not m_Logger Is Nothing Then
        ' TODO: Utiliser le logger pour enregistrer l'erreur
        ' m_Logger.LogError errorMessage, "SECURITY"
    End If
End Sub
