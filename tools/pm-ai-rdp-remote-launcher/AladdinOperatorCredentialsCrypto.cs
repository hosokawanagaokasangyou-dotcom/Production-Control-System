using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace PmAi.RdpRemoteLauncher;

internal static class AladdinOperatorCredentialsCrypto
{
    internal const int FormatVersion = 1;
    internal const int DefaultIterations = 480_000;
    internal const string DefaultPassphrase = "pm-ai-aladdin-operator";

    internal static string DecryptFromPayload(JsonElement payload, string? passphrase = null)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException("暗号化ペイロードが不正です。");
        }

        var version = payload.TryGetProperty("v", out var vNode) ? vNode.GetInt32() : 0;
        if (version != FormatVersion)
        {
            throw new InvalidOperationException("未対応の暗号化形式です: v=" + version);
        }

        var phrase = string.IsNullOrWhiteSpace(passphrase) ? DefaultPassphrase : passphrase.Trim();
        var iterations = payload.TryGetProperty("iterations", out var iterNode)
            ? iterNode.GetInt32()
            : DefaultIterations;
        var salt = Convert.FromBase64String(payload.GetProperty("salt_b64").GetString() ?? "");
        var iv = Convert.FromBase64String(payload.GetProperty("iv_b64").GetString() ?? "");
        var ciphertext = Convert.FromBase64String(payload.GetProperty("ciphertext_b64").GetString() ?? "");

        var key = DeriveKey(phrase, salt, iterations);
        using var aes = Aes.Create();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.Key = key;
        aes.IV = iv;
        using var decryptor = aes.CreateDecryptor();
        var plain = decryptor.TransformFinalBlock(ciphertext, 0, ciphertext.Length);
        return Encoding.UTF8.GetString(plain);
    }

    private static byte[] DeriveKey(string passphrase, byte[] salt, int iterations)
    {
        using var derive = new Rfc2898DeriveBytes(
            passphrase,
            salt,
            iterations,
            HashAlgorithmName.SHA256);
        return derive.GetBytes(32);
    }
}
