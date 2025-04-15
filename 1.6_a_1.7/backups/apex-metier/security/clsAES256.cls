VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAES256"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("APEX.Metier.Security")
Option Explicit

'==========================================================================
' Classe    : clsAES256
' Purpose   : Implémentation AES-256 pour le chiffrement des données
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : SEC-002
'==========================================================================

#If VBA7 Then
    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptOpenAlgorithmProvider Lib "bcrypt" ( _
        ByRef phAlgorithm As LongPtr, _
        ByVal pszAlgId As LongPtr, _
        ByVal pszImplementation As LongPtr, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptGetProperty Lib "bcrypt" ( _
        ByVal hObject As LongPtr, _
        ByVal pszProperty As LongPtr, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByRef pcbResult As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptSetProperty Lib "bcrypt" ( _
        ByVal hObject As LongPtr, _
        ByVal pszProperty As LongPtr, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptGenerateSymmetricKey Lib "bcrypt" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef phKey As LongPtr, _
        ByRef pbKeyObject As Any, _
        ByVal cbKeyObject As Long, _
        ByRef pbSecret As Any, _
        ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptEncrypt Lib "bcrypt" ( _
        ByVal hKey As LongPtr, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal pPaddingInfo As LongPtr, _
        ByRef pbIV As Any, _
        ByVal cbIV As Long, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByRef pcbResult As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptDecrypt Lib "bcrypt" ( _
        ByVal hKey As LongPtr, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal pPaddingInfo As LongPtr, _
        ByRef pbIV As Any, _
        ByVal cbIV As Long, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByRef pcbResult As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptDestroyKey Lib "bcrypt" ( _
        ByVal hKey As LongPtr) As Long

    Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function BCryptCloseAlgorithmProvider Lib "bcrypt" ( _
        ByVal hAlgorithm As LongPtr, _
        ByVal dwFlags As Long) As Long
#End If

' Constantes
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = &H8
Private Const BCRYPT_BLOCK_PADDING As Long = &H1

Private Const BCRYPT_AES_ALGORITHM As String = "AES"
Private Const BCRYPT_CHAIN_MODE_CBC As String = "ChainingModeCBC"
Private Const BCRYPT_KEY_LENGTH As String = "KeyLength"

Private Type TAES256State
    AlgorithmHandle As LongPtr
    KeyHandle As LongPtr
    KeyObject() As Byte
    IV() As Byte
    IsInitialized As Boolean
End Type

Private this As TAES256State

'@Description("Initialise le fournisseur AES-256")
'@Description: 
'@Param: 
'@Returns: 

Public Function Initialize() As Boolean
    Dim status As Long
    Dim cbResult As Long
    Dim cbKeyObject As Long
    
    ' Ouvrir le fournisseur AES
    status = BCryptOpenAlgorithmProvider(this.AlgorithmHandle, _
                                       StrPtr(BCRYPT_AES_ALGORITHM), _
                                       0, 0)
    If status <> 0 Then
        Debug.Print "Erreur BCryptOpenAlgorithmProvider: " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Configurer le mode CBC
    status = BCryptSetProperty(this.AlgorithmHandle, _
                             StrPtr(BCRYPT_CHAIN_MODE_CBC), _
                             ByVal StrPtr(BCRYPT_CHAIN_MODE_CBC), _
                             LenB(BCRYPT_CHAIN_MODE_CBC), _
                             0)
    If status <> 0 Then
        Debug.Print "Erreur BCryptSetProperty: " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Obtenir la taille du buffer de clé
    status = BCryptGetProperty(this.AlgorithmHandle, _
                             StrPtr("ObjectLength"), _
                             cbKeyObject, 4, cbResult, 0)
    If status <> 0 Then
        Debug.Print "Erreur BCryptGetProperty: " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Allouer le buffer de clé
    ReDim this.KeyObject(0 To cbKeyObject - 1)
    
    this.IsInitialized = True
    Initialize = True
End Function

'@Description("Génère une clé AES-256 à partir d'une clé secrète")
'@Description: 
'@Param: 
'@Returns: 

Public Function GenerateKey(ByRef secretKey() As Byte) As Boolean
    If Not this.IsInitialized Then
        Debug.Print "AES non initialisé"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Générer la clé
    Dim status As Long
    status = BCryptGenerateSymmetricKey(this.AlgorithmHandle, _
                                      this.KeyHandle, _
                                      this.KeyObject(0), _
                                      UBound(this.KeyObject) + 1, _
                                      secretKey(0), _
                                      UBound(secretKey) + 1, _
                                      0)
    
    If status <> 0 Then
        Debug.Print "Erreur BCryptGenerateSymmetricKey: " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    GenerateKey = True
End Function

'@Description("Chiffre des données avec AES-256")
'@Description: 
'@Param: 
'@Returns: 

Public Function Encrypt(ByRef data() As Byte, ByRef IV() As Byte) As Byte()
    If Not this.IsInitialized Then
        Debug.Print "AES non initialisé"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Calculer la taille du buffer de sortie
    Dim cbOutput As Long
    Dim status As Long
    status = BCryptEncrypt(this.KeyHandle, _
                          data(0), _
                          UBound(data) + 1, _
                          0, _
                          IV(0), _
                          UBound(IV) + 1, _
                          ByVal 0&, _
                          0, _
                          cbOutput, _
                          BCRYPT_BLOCK_PADDING)
    
    If status <> 0 Then
        Debug.Print "Erreur BCryptEncrypt (calcul taille): " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Allouer le buffer de sortie
    Dim output() As Byte
    ReDim output(0 To cbOutput - 1)
    
    ' Chiffrer les données
    status = BCryptEncrypt(this.KeyHandle, _
                          data(0), _
                          UBound(data) + 1, _
                          0, _
                          IV(0), _
                          UBound(IV) + 1, _
                          output(0), _
                          cbOutput, _
                          cbOutput, _
                          BCRYPT_BLOCK_PADDING)
    
    If status <> 0 Then
        Debug.Print "Erreur BCryptEncrypt: " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Encrypt = output
End Function

'@Description("Déchiffre des données avec AES-256")
'@Description: 
'@Param: 
'@Returns: 

Public Function Decrypt(ByRef data() As Byte, ByRef IV() As Byte) As Byte()
    If Not this.IsInitialized Then
        Debug.Print "AES non initialisé"
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Calculer la taille du buffer de sortie
    Dim cbOutput As Long
    Dim status As Long
    status = BCryptDecrypt(this.KeyHandle, _
                          data(0), _
                          UBound(data) + 1, _
                          0, _
                          IV(0), _
                          UBound(IV) + 1, _
                          ByVal 0&, _
                          0, _
                          cbOutput, _
                          BCRYPT_BLOCK_PADDING)
    
    If status <> 0 Then
        Debug.Print "Erreur BCryptDecrypt (calcul taille): " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    ' Allouer le buffer de sortie
    Dim output() As Byte
    ReDim output(0 To cbOutput - 1)
    
    ' Déchiffrer les données
    status = BCryptDecrypt(this.KeyHandle, _
                          data(0), _
                          UBound(data) + 1, _
                          0, _
                          IV(0), _
                          UBound(IV) + 1, _
                          output(0), _
                          cbOutput, _
                          cbOutput, _
                          BCRYPT_BLOCK_PADDING)
    
    If status <> 0 Then
        Debug.Print "Erreur BCryptDecrypt: " & status
        Exit'@Description: 
'@Param: 
'@Returns: 

 Function
    End If
    
    Decrypt = output
End Function

'@Description("Nettoie les ressources")
'@Description: 
'@Param: 
'@Returns: 

Public Sub Cleanup()
    If this.KeyHandle <> 0 Then
        BCryptDestroyKey this.KeyHandle
        this.KeyHandle = 0
    End If
    
    If this.AlgorithmHandle <> 0 Then
        BCryptCloseAlgorithmProvider this.AlgorithmHandle, 0
        this.AlgorithmHandle = 0
    End If
    
    this.IsInitialized = False
End'@Description: 
'@Param: 
'@Returns: 

 Sub

Private Sub Class_Terminate()
    Cleanup
End Sub 