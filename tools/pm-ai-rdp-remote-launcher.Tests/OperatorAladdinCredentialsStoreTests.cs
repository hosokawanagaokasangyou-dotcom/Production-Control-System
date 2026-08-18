using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public sealed class OperatorAladdinCredentialsStoreTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string? _priorFactorySite;

    public OperatorAladdinCredentialsStoreTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "PmAiRdpCredTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _priorFactorySite = Environment.GetEnvironmentVariable(
            OperatorAladdinCredentialsStore.FactoryEnvVar);
        Environment.SetEnvironmentVariable(OperatorAladdinCredentialsStore.FactoryEnvVar, "KONAN");
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable(
            OperatorAladdinCredentialsStore.FactoryEnvVar,
            _priorFactorySite);
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // ignore cleanup failures on shared temp
        }
    }

    [Fact]
    public void Resolve_fallsBackToRdpLauncherFactoryWhenKonanMissing()
    {
        var iniPath = Path.Combine(_tempDir, "細川_RPA設定.ini");
        File.WriteAllText(iniPath, "[RPA]\n操作者=細川\n");
        var jsonPath = Path.Combine(_tempDir, OperatorAladdinCredentialsStore.FileName);
        File.WriteAllText(
            jsonPath,
            """
            {
              "schemaVersion": 1,
              "factories": {
                "RDP_LAUNCHER": {
                  "細川": {
                    "loginId": "000585",
                    "password": {
                      "v": 1,
                      "kdf": "pbkdf2_sha256",
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                }
              }
            }
            """);

        var credentials = OperatorAladdinCredentialsStore.Resolve(iniPath, "細川");

        Assert.NotNull(credentials);
        Assert.Equal("000585", credentials!.LoginId);
        Assert.False(string.IsNullOrWhiteSpace(credentials.Password));
    }

    [Fact]
    public void Resolve_prefersPrimaryFactoryOverRdpLauncher()
    {
        var iniPath = Path.Combine(_tempDir, "砂田_RPA設定.ini");
        File.WriteAllText(iniPath, "[RPA]\n操作者=砂田\n");
        var jsonPath = Path.Combine(_tempDir, OperatorAladdinCredentialsStore.FileName);
        File.WriteAllText(
            jsonPath,
            """
            {
              "schemaVersion": 1,
              "factories": {
                "KONAN": {
                  "砂田": {
                    "loginId": "111111",
                    "password": {
                      "v": 1,
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                },
                "RDP_LAUNCHER": {
                  "砂田": {
                    "loginId": "999999",
                    "password": {
                      "v": 1,
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                }
              }
            }
            """);

        var credentials = OperatorAladdinCredentialsStore.Resolve(iniPath, "砂田");

        Assert.NotNull(credentials);
        Assert.Equal("111111", credentials!.LoginId);
    }

    [Fact]
    public void Resolve_usesKokubuBlockWhenPrimaryFactoryMissing()
    {
        var iniPath = Path.Combine(_tempDir, "細川_RPA設定.ini");
        File.WriteAllText(iniPath, "[RPA]\n操作者=細川\n");
        var jsonPath = Path.Combine(_tempDir, OperatorAladdinCredentialsStore.FileName);
        File.WriteAllText(
            jsonPath,
            """
            {
              "schemaVersion": 1,
              "factories": {
                "KOKUBU": {
                  "細川": {
                    "loginId": "00585",
                    "password": {
                      "v": 1,
                      "kdf": "pbkdf2_sha256",
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                }
              }
            }
            """);

        var credentials = OperatorAladdinCredentialsStore.Resolve(iniPath, "細川");

        Assert.NotNull(credentials);
        Assert.Equal("00585", credentials!.LoginId);
        Assert.False(string.IsNullOrWhiteSpace(credentials.Password));
    }

    [Fact]
    public void Resolve_prefersPrimaryFactoryOverKokubu()
    {
        var iniPath = Path.Combine(_tempDir, "細川_RPA設定.ini");
        File.WriteAllText(iniPath, "[RPA]\n操作者=細川\n");
        var jsonPath = Path.Combine(_tempDir, OperatorAladdinCredentialsStore.FileName);
        File.WriteAllText(
            jsonPath,
            """
            {
              "schemaVersion": 1,
              "factories": {
                "KONAN": {
                  "細川": {
                    "loginId": "000585",
                    "password": {
                      "v": 1,
                      "kdf": "pbkdf2_sha256",
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                },
                "KOKUBU": {
                  "細川": {
                    "loginId": "00585",
                    "password": {
                      "v": 1,
                      "kdf": "pbkdf2_sha256",
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                }
              }
            }
            """);

        var credentials = OperatorAladdinCredentialsStore.Resolve(iniPath, "細川");

        Assert.NotNull(credentials);
        Assert.Equal("000585", credentials!.LoginId);
    }

    [Fact]
    public void Resolve_prefersUniqueOperatorsOverFactoryBlocks()
    {
        var iniPath = Path.Combine(_tempDir, "細川_RPA設定.ini");
        File.WriteAllText(iniPath, "[RPA]\n操作者=細川\n");
        var jsonPath = Path.Combine(_tempDir, OperatorAladdinCredentialsStore.FileName);
        File.WriteAllText(
            jsonPath,
            """
            {
              "schemaVersion": 2,
              "operators": {
                "細川": {
                  "loginId": "00585",
                  "password": {
                    "v": 1,
                    "kdf": "pbkdf2_sha256",
                    "iterations": 480000,
                    "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                    "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                    "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                  }
                }
              },
              "factories": {
                "KONAN": {
                  "細川": {
                    "loginId": "000585",
                    "password": {
                      "v": 1,
                      "kdf": "pbkdf2_sha256",
                      "iterations": 480000,
                      "salt_b64": "MPcAl3n25c7Y8/Emh5LDjg==",
                      "iv_b64": "JFzysmQq0ZpyVlQ3KtS/cA==",
                      "ciphertext_b64": "Wc8dqMsVZrNykYMOnm4PDg=="
                    }
                  }
                }
              }
            }
            """);

        var credentials = OperatorAladdinCredentialsStore.Resolve(iniPath, "細川");

        Assert.NotNull(credentials);
        Assert.Equal("00585", credentials!.LoginId);
    }
}
