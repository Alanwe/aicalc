using System;
using System.Security.Cryptography;
using System.Text;

namespace AiCalc.Services.AI;

/// <summary>
/// Secure credential storage using Windows DPAPI
/// </summary>
public static class CredentialService
{
    /// <summary>
    /// Encrypt a plain text string using DPAPI (CurrentUser scope)
    /// </summary>
    public static string Encrypt(string plainText)
    {
        if (string.IsNullOrEmpty(plainText))
            return string.Empty;
        
        try
        {
            var plainBytes = Encoding.UTF8.GetBytes(plainText);
            var encryptedBytes = ProtectedData.Protect(
                plainBytes,
                optionalEntropy: null,
                scope: DataProtectionScope.CurrentUser
            );
            return Convert.ToBase64String(encryptedBytes);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to encrypt data", ex);
        }
    }
    
    /// <summary>
    /// Decrypt an encrypted string using DPAPI
    /// </summary>
    public static string Decrypt(string encryptedText)
    {
        if (string.IsNullOrEmpty(encryptedText))
            return string.Empty;
        
        try
        {
            var encryptedBytes = Convert.FromBase64String(encryptedText);
            var plainBytes = ProtectedData.Unprotect(
                encryptedBytes,
                optionalEntropy: null,
                scope: DataProtectionScope.CurrentUser
            );
            return Encoding.UTF8.GetString(plainBytes);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to decrypt data", ex);
        }
    }
    
    /// <summary>
    /// Test if a string is encrypted (base64 encoded)
    /// </summary>
    public static bool IsEncrypted(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;
        
        try
        {
            Convert.FromBase64String(text);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
