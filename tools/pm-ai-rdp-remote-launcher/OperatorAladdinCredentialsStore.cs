using System.Text.Json;

namespace PmAi.RdpRemoteLauncher;

internal sealed record OperatorAladdinCredentials(string LoginId, string Password);

internal static class OperatorAladdinCredentialsStore
{
    internal const string FileName = "operator-aladdin-credentials.launcher.json";
    internal const string OperatorEnvVar = "PM_AI_OPERATOR_USER";
    internal const string FactoryEnvVar = "PM_AI_FACTORY_SITE";

    internal static OperatorAladdinCredentials? Resolve(string iniPath, string? operatorFromIni)
    {
        var operatorName = ResolveOperatorName(operatorFromIni);
        if (string.IsNullOrWhiteSpace(operatorName))
        {
            LauncherLog.Info("操作者名が未設定のためアラジン資格情報を解決しません。");
            return null;
        }

        var jsonPath = ResolveJsonPath(iniPath);
        if (!File.Exists(jsonPath))
        {
            LauncherLog.Error("アラジン資格情報 JSON が見つかりません: " + jsonPath);
            return null;
        }

        try
        {
            using var stream = File.OpenRead(jsonPath);
            using var doc = JsonDocument.Parse(stream);
            var root = doc.RootElement;
            var factory = ResolveFactorySite();
            if (!TryReadOperatorEntry(root, factory, operatorName, out var loginId, out var passwordPayload))
            {
                LauncherLog.Error(
                    "操作者のアラジン資格情報が未設定です: operator="
                        + operatorName
                        + " factory="
                        + factory);
                return null;
            }

            var password = AladdinOperatorCredentialsCrypto.DecryptFromPayload(passwordPayload);
            if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            {
                LauncherLog.Error("操作者のアラジン資格情報が不完全です: " + operatorName);
                return null;
            }

            LauncherLog.Info("アラジン資格情報を解決しました: operator=" + operatorName);
            return new OperatorAladdinCredentials(loginId.Trim(), password);
        }
        catch (Exception ex)
        {
            LauncherLog.Error("アラジン資格情報 JSON の読込に失敗: " + ex.Message);
            return null;
        }
    }

    private static string ResolveOperatorName(string? operatorFromIni)
    {
        if (!string.IsNullOrWhiteSpace(operatorFromIni))
        {
            return operatorFromIni.Trim();
        }

        var fromEnv = Environment.GetEnvironmentVariable(OperatorEnvVar);
        return string.IsNullOrWhiteSpace(fromEnv) ? "" : fromEnv.Trim();
    }

    private static string ResolveFactorySite()
    {
        var fromEnv = Environment.GetEnvironmentVariable(FactoryEnvVar);
        if (!string.IsNullOrWhiteSpace(fromEnv))
        {
            return fromEnv.Trim().ToUpperInvariant();
        }

        return "KONAN";
    }

    private static string ResolveJsonPath(string iniPath)
    {
        var dir = Path.GetDirectoryName(iniPath);
        if (string.IsNullOrWhiteSpace(dir))
        {
            return FileName;
        }

        return Path.Combine(dir, FileName);
    }

    private static bool TryReadOperatorEntry(
        JsonElement root,
        string factory,
        string operatorName,
        out string loginId,
        out JsonElement passwordPayload)
    {
        loginId = "";
        passwordPayload = default;
        if (root.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        if (root.TryGetProperty("factories", out var factories) && factories.ValueKind == JsonValueKind.Object)
        {
            if (!factories.TryGetProperty(factory, out var factoryNode)
                || factoryNode.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            return TryReadOperatorObject(factoryNode, operatorName, out loginId, out passwordPayload);
        }

        if (root.TryGetProperty("factory", out var legacyFactory)
            && legacyFactory.ValueKind == JsonValueKind.String
            && !string.Equals(legacyFactory.GetString(), factory, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (root.TryGetProperty("operators", out var legacyOps))
        {
            return TryReadOperatorObject(legacyOps, operatorName, out loginId, out passwordPayload);
        }

        return false;
    }

    private static bool TryReadOperatorObject(
        JsonElement operators,
        string operatorName,
        out string loginId,
        out JsonElement passwordPayload)
    {
        loginId = "";
        passwordPayload = default;
        if (operators.ValueKind != JsonValueKind.Object
            || !operators.TryGetProperty(operatorName, out var row)
            || row.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        loginId = row.TryGetProperty("loginId", out var idNode) ? idNode.GetString() ?? "" : "";
        if (!row.TryGetProperty("password", out passwordPayload) || passwordPayload.ValueKind != JsonValueKind.Object)
        {
            loginId = "";
            passwordPayload = default;
            return false;
        }

        return true;
    }
}
