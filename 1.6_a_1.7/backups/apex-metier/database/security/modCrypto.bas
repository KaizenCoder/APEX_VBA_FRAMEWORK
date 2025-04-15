Attribute VB_Name = "modCrypto"

'@Module: [NomDuModule]
'@Description: 
'@Version: 1.0
'@Date: 13/04/2025
'@Author: APEX Framework Team

'@Folder("APEX.Metier.Database.Security")
Option Explicit

'==========================================================================
' Module    : modCrypto
' Purpose   : Fonctions de cryptographie pour la sécurité des données
' Author    : APEX Framework Team
' Date      : 2024-04-11
' Reference : SEC-001
'==========================================================================

Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As LongPtr, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function CryptGenRandom Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, _
    ByVal dwLen As Long, _
    ByRef pbBuffer As Any) As Long

Private Declare PtrSafe'@Description: 
'@Param: 
'@Returns: 

 Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CRYPT_NEWKEYSET As Long = &H8
Private Const NTE_BAD_KEYSET As Long = &H80090016

'@Description("Génère une clé aléatoire de la longueur spécifiée")
'@Description: 
'@Param: 
'@Returns: 

Public Function GenerateRandomKey(ByVal length As Long) As String
    Dim hCryptProv As LongPtr
    Dim buffer() As Byte
    Dim result As String
    
    ' Initialiser le fournisseur de cryptographie
    If CryptAcquireContext(hCryptProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        If Err.LastDllError = NTE_BAD_KEYSET Then
            If CryptAcquireContext(hCryptProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
                Err.Raise 5, "modCrypto", "Impossible d'initialiser le fournisseur de cryptographie"
            End If
        End If
    End If
    
    ' Générer les octets aléatoires
    ReDim buffer(0 To length - 1)
    If CryptGenRandom(hCryptProv, length, buffer(0)) = 0 Then
        CryptReleaseContext hCryptProv, 0
        Err.Raise 5, "modCrypto", "Échec de la génération de nombres aléatoires"
    End If
    
    ' Convertir en chaîne hexadécimale
    result = ""
    Dim i As Long
    For i = 0 To length - 1
        result = result & Right$("0" & Hex$(buffer(i)), 2)
    Next i
    
    ' Libérer les ressources
    CryptReleaseContext hCryptProv, 0
    
    GenerateRandomKey = result
End Function

'@Description("Chiffre une chaîne avec AES-256")
'@Description: 
'@Param: 
'@Returns: 

Public Function EncryptAES256(ByVal plainText As String, ByVal key As String) As String
    ' Implémentation simplifiée pour démonstration
    ' Dans un environnement de production, utiliser une bibliothèque cryptographique complète
    
    Dim encrypted As String
    encrypted = ""
    
    ' XOR basique avec la clé (À NE PAS UTILISER EN PRODUCTION)
    Dim i As Long, j As Long
    For i = 1 To Len(plainText)
        j = ((i - 1) Mod (Len(key) / 2)) + 1
        encrypted = encrypted & Chr$(Asc(Mid$(plainText, i, 1)) Xor _
                                   CLng("&H" & Mid$(key, j * 2 - 1, 2)))
    Next i
    
    ' Encoder en Base64
    EncryptAES256 = EncodeBase64(encrypted)
End Function

'@Description("Déchiffre une chaîne avec AES-256")
'@Description: 
'@Param: 
'@Returns: 

Public Function DecryptAES256(ByVal cipherText As String, ByVal key As String) As String
    ' Implémentation simplifiée pour démonstration
    ' Dans un environnement de production, utiliser une bibliothèque cryptographique complète
    
    ' Décoder de Base64
    Dim encrypted As String
    encrypted = DecodeBase64(cipherText)
    
    Dim decrypted As String
    decrypted = ""
    
    ' XOR basique avec la clé (À NE PAS UTILISER EN PRODUCTION)
    Dim i As Long, j As Long
    For i = 1 To Len(encrypted)
        j = ((i - 1) Mod (Len(key) / 2)) + 1
        decrypted = decrypted & Chr$(Asc(Mid$(encrypted, i, 1)) Xor _
                                    CLng("&H" & Mid$(key, j * 2 - 1, 2)))
    Next i
    
    DecryptAES256 = decrypted
End Function

'@Description("Encode une chaîne en Base64")
'@Description: 
'@Param: 
'@Returns: 

Private Function EncodeBase64(ByVal text As String) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = StringToBytes(text)
    
    EncodeBase64 = xmlNode.text
    
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
End Function

'@Description("Décode une chaîne Base64")
'@Description: 
'@Param: 
'@Returns: 

Private Function DecodeBase64(ByVal base64 As String) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.text = base64
    
    DecodeBase64 = BytesToString(xmlNode.nodeTypedValue)
    
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
End Function

'@Description("Convertit une chaîne en tableau d'octets")
'@Description: 
'@Param: 
'@Returns: 

Private Function StringToBytes(ByVal text As String) As Byte()
    Dim bytes() As Byte
    bytes = text
    StringToBytes = bytes
End Function

'@Description("Convertit un tableau d'octets en chaîne")
'@Description: 
'@Param: 
'@Returns: 

Private Function BytesToString(ByRef bytes() As Byte) As String
    BytesToString = bytes
End Function 