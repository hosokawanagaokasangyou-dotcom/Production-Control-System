package jp.co.pm.ai.desktop.io;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * Windows の {@code rdpsign.exe} を使い .rdp ファイルへデジタル署名する。
 *
 * <p>証明書は CurrentUser / LocalMachine の Personal ストアから列挙する（秘密鍵付き・有効期限内）。
 */
public final class RdpFileSigner {

    private static final int PROCESS_TIMEOUT_SEC = 120;
    private static final int OUTPUT_CAPTURE_MAX = 256_000;
    private static final Pattern THUMBPRINT_PATTERN = Pattern.compile("[0-9A-Fa-f]{40}");
    private static final DateTimeFormatter BACKUP_STAMP =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");

    public record SigningCertificate(
            String thumbprintSha1, String subject, String storeLabel, String notAfter, String usageLabel) {

        public SigningCertificate {
            usageLabel = usageLabel != null && !usageLabel.isBlank() ? usageLabel : "—";
        }

        /** rdpsign.exe が使える証明書か（コード署名 EKU、または TLS 以外のデジタル署名）。 */
        public boolean rdpSignCapable() {
            return !"除外".equals(usageLabel);
        }

        @Override
        public String toString() {
            String subj = subject != null && subject.length() > 72 ? subject.substring(0, 69) + "..." : subject;
            return storeLabel + " | " + subj + " | " + usageLabel + " | 有効期限 " + notAfter;
        }
    }

    public record CertificateListResult(List<SigningCertificate> eligible, int skippedIneligibleCount) {

        /** 作成直後など一覧に載らない証明書を先頭にマージする。 */
        public CertificateListResult withEnsuredEligible(SigningCertificate ensurePresent) {
            if (ensurePresent == null) {
                return this;
            }
            List<SigningCertificate> merged = new ArrayList<>();
            merged.add(ensurePresent);
            for (SigningCertificate c : eligible) {
                if (!ensurePresent.thumbprintSha1().equals(c.thumbprintSha1())) {
                    merged.add(c);
                }
            }
            return new CertificateListResult(List.copyOf(merged), skippedIneligibleCount);
        }
    }

    public record CommandResult(int exitCode, String output) {
        public boolean success() {
            return exitCode == 0;
        }
    }

    public static final String SIGNED_OUTPUT_SUFFIX = ".pm-ai-signed.rdp";

    /** UAC キャンセル時の Windows 終了コード（ERROR_CANCELLED / 0x4C7）。 */
    public static final int UAC_CANCELLED_EXIT_CODE = 1223;

    /** 署名処理の実際の対象パス（%TEMP% 上の作業コピー）。 */
    public record SigningTarget(
            Path signingPath, Path profilePath, Path sourcePath, boolean createsNewProfileFile) {

        /** 起動・環境変数に使う .rdp のパス（署名済み新規ファイル）。 */
        public Path effectiveProfilePath() {
            return profilePath != null ? profilePath : signingPath;
        }
    }

    public record SignAttemptResult(CommandResult result, SigningTarget target) {}

    private RdpFileSigner() {}

    public static boolean isSupportedPlatform() {
        return isWindows();
    }

    private static Path resolveRdpsignExe() throws IOException {
        if (!isSupportedPlatform()) {
            throw new IOException("RDP 署名は Windows のみ対応です。");
        }
        List<Path> candidates = new ArrayList<>();
        String windir = System.getenv("SystemRoot");
        if (windir != null && !windir.isBlank()) {
            String root = windir.trim();
            candidates.add(Path.of(root, "System32", "rdpsign.exe"));
            if (is32BitJvmOn64BitWindows()) {
                candidates.add(0, Path.of(root, "Sysnative", "rdpsign.exe"));
            }
        }
        candidates.add(Path.of("C:\\Windows\\System32\\rdpsign.exe"));
        if (is32BitJvmOn64BitWindows()) {
            candidates.add(0, Path.of("C:\\Windows\\Sysnative\\rdpsign.exe"));
        }
        for (Path p : candidates) {
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
        }
        throw new IOException("rdpsign.exe が見つかりません。");
    }

    private static boolean is32BitJvmOn64BitWindows() {
        String arch = System.getProperty("os.arch", "").toLowerCase(Locale.ROOT);
        boolean jvm32 = arch.contains("86") || "i386".equals(arch);
        String pfx86 = System.getenv("ProgramFiles(x86)");
        return jvm32 && pfx86 != null && !pfx86.isBlank();
    }

    /** {@code signature:s:} 行の有無で署名済みか判定する（UTF-8 / UTF-16 LE を問わず検索）。 */
    public static boolean isSigned(Path rdpProfile) throws IOException {
        Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
        byte[] bytes = Files.readAllBytes(abs);
        if (containsSignatureMarker(bytes) || containsUtf16LeAsciiMarker(bytes, "signature:s:")) {
            return true;
        }
        for (String line : Files.readAllLines(abs, StandardCharsets.UTF_8)) {
            if (line != null && line.stripLeading().toLowerCase(Locale.ROOT).startsWith("signature:s:")) {
                return true;
            }
        }
        return false;
    }

    static boolean containsSignatureMarker(byte[] bytes) {
        return indexOfAscii(bytes, "signature:s:", true) >= 0;
    }

    static boolean containsUtf16LeAsciiMarker(byte[] bytes, String ascii) {
        if (ascii == null || ascii.isEmpty()) {
            return false;
        }
        byte[] pattern = new byte[ascii.length() * 2];
        for (int i = 0; i < ascii.length(); i++) {
            pattern[i * 2] = (byte) ascii.charAt(i);
        }
        return indexOf(bytes, pattern) >= 0;
    }

    private static int indexOfAscii(byte[] bytes, String ascii, boolean caseInsensitive) {
        byte[] marker = ascii.getBytes(StandardCharsets.US_ASCII);
        if (bytes.length < marker.length) {
            return -1;
        }
        outer:
        for (int i = 0; i <= bytes.length - marker.length; i++) {
            for (int j = 0; j < marker.length; j++) {
                byte fileByte = bytes[i + j];
                byte markerByte = marker[j];
                if (caseInsensitive) {
                    if (toLowerAscii(fileByte) != toLowerAscii(markerByte)) {
                        continue outer;
                    }
                } else if (fileByte != markerByte) {
                    continue outer;
                }
            }
            return i;
        }
        return -1;
    }

    private static byte toLowerAscii(byte b) {
        if (b >= 'A' && b <= 'Z') {
            return (byte) (b + 32);
        }
        return b;
    }

    private static int indexOf(byte[] bytes, byte[] pattern) {
        if (bytes.length < pattern.length) {
            return -1;
        }
        outer:
        for (int i = 0; i <= bytes.length - pattern.length; i++) {
            for (int j = 0; j < pattern.length; j++) {
                if (bytes[i + j] != pattern[j]) {
                    continue outer;
                }
            }
            return i;
        }
        return -1;
    }

    public record ProfileTrustDiagnosis(
            Path profilePath,
            boolean signaturePresent,
            boolean hkcuTrustConfigured,
            boolean hklmTrustConfigured,
            String hkcuThumbprints,
            String hklmThumbprints,
            String expectedThumbprint) {

        public boolean expectedThumbprintTrusted() {
            if (expectedThumbprint == null || expectedThumbprint.isBlank()) {
                return false;
            }
            String expected = expectedThumbprint.toUpperCase(Locale.ROOT);
            return thumbprintListContains(hkcuThumbprints, expected)
                    || thumbprintListContains(hklmThumbprints, expected);
        }

        private static boolean thumbprintListContains(String list, String thumb) {
            if (list == null || list.isBlank()) {
                return false;
            }
            for (String part : list.split(";")) {
                if (thumb.equals(part.replaceAll("\\s+", "").toUpperCase(Locale.ROOT))) {
                    return true;
                }
            }
            return false;
        }

        public String summary() {
            StringBuilder sb = new StringBuilder();
            sb.append("ファイル: ").append(profilePath).append('\n');
            sb.append("signature:s: 行: ").append(signaturePresent ? "あり" : "なし（未署名）").append('\n');
            sb.append("HKCU 信頼設定: ")
                    .append(hkcuTrustConfigured ? "あり" : "なし")
                    .append(hkcuThumbprints != null && !hkcuThumbprints.isBlank()
                            ? " / " + hkcuThumbprints
                            : "")
                    .append('\n');
            sb.append("HKLM 信頼設定: ")
                    .append(hklmTrustConfigured ? "あり" : "なし")
                    .append(hklmThumbprints != null && !hklmThumbprints.isBlank()
                            ? " / " + hklmThumbprints
                            : "")
                    .append('\n');
            if (!expectedThumbprint.isBlank()) {
                sb.append("期待サムプリント: ").append(expectedThumbprint);
                sb.append(expectedThumbprintTrusted() ? " → 登録済み" : " → 未登録");
            }
            if (!signaturePresent) {
                sb.append("\n\n【対処】RDP 署名ウィザードで本署名をやり直してください。");
            } else if (!expectedThumbprintTrusted()) {
                sb.append("\n\n【対処】ステップ4で「信頼設定を適用（HKCU+HKLM）」を実行してください。");
            } else if (!hkcuTrustConfigured && hklmTrustConfigured) {
                sb.append("\n\n【状態】HKLM に期待サムプリントは登録済みです。");
                sb.append(" HKCU は未設定ですが、通常は HKLM のみでも警告は抑止されます。");
                sb.append("\n社内 GPO により HKCU\\...\\Policies への書込が拒否されることがあります（IT 管理下では正常）。");
                sb.append("\n\n【警告が続く場合】");
                sb.append("\n1) gpupdate /force のあと PC 再起動、または mstsc 終了後に .rdp を開き直す");
                sb.append("\n2) 起動ファイルが Default.pm-ai-signed.rdp（署名済み）であることを確認");
                sb.append("\n3) gpresult /Scope User /v で Remote Desktop Connection Client のユーザー GPO を確認");
            } else if (!hkcuTrustConfigured) {
                sb.append("\n\n【対処】HKCU にも同じサムプリントを登録してください（「現在ユーザー（HKCU）のみ」）。");
            } else {
                sb.append("\n\n設定は揃っています。mstsc を終了してから .rdp を開き直してください。");
            }
            return sb.toString();
        }
    }

    /** 署名済み .rdp と HKCU/HKLM 信頼ポリシーの状態を診断する。 */
    public static ProfileTrustDiagnosis diagnoseProfileTrust(Path profilePath, String expectedThumbprintSha1)
            throws IOException {
        if (!isSupportedPlatform()) {
            throw new IOException("RDP 信頼診断は Windows のみ対応です。");
        }
        Path abs = RemoteDesktopLauncher.validateRdpProfile(profilePath);
        String expected =
                expectedThumbprintSha1 != null && !expectedThumbprintSha1.isBlank()
                        ? normalizeThumbprintSha1(expectedThumbprintSha1)
                        : "";
        boolean signed = isSigned(abs);
        TrustPolicyState hkcu = readTrustPolicyState(false);
        TrustPolicyState hklm = readTrustPolicyState(true);
        return new ProfileTrustDiagnosis(
                abs,
                signed,
                hkcu.configured(),
                hklm.configured(),
                hkcu.thumbprints(),
                hklm.thumbprints(),
                expected);
    }

    private record TrustPolicyState(boolean configured, String thumbprints) {}

    private static TrustPolicyState readTrustPolicyState(boolean machineWide) throws IOException {
        String hive = machineWide ? "HKLM" : "HKCU";
        String script =
                """
                $path = '%s:\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Terminal Services'
                $p = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                if ($null -eq $p) { Write-Output '#NONE#'; exit 0 }
                $allow = $p.AllowSignedFiles
                $thumb = [string]$p.TrustedCertThumbprints
                Write-Output ($allow.ToString() + '|' + $thumb)
                """
                        .formatted(hive);
        CommandResult result = runPowerShell(script);
        if (!result.success() || result.output().isBlank() || result.output().contains("#NONE#")) {
            return new TrustPolicyState(false, "");
        }
        String[] parts = result.output().strip().split("\\|", 2);
        boolean configured = parts.length > 0 && "1".equals(parts[0].strip());
        String thumbs = parts.length > 1 ? parts[1].strip() : "";
        return new TrustPolicyState(configured || !thumbs.isBlank(), thumbs);
    }

    public static String normalizeThumbprintSha1(String raw) {
        if (raw == null) {
            return "";
        }
        String compact = raw.replaceAll("\\s+", "").toUpperCase(Locale.ROOT);
        if (!THUMBPRINT_PATTERN.matcher(compact).matches()) {
            throw new IllegalArgumentException("証明書サムプリント（SHA-1・40桁）が不正です: " + raw);
        }
        return compact;
    }

    /**
     * Personal ストアから {@code rdpsign.exe} が使える証明書のみ列挙する。
     *
     * <p>SSL/TLS（サーバー認証）証明書などは一覧から除外する。除外件数は {@link CertificateListResult#skippedIneligibleCount()}。
     *
     * @throws IOException PowerShell 実行失敗
     */
    public static CertificateListResult listSigningCertificates() throws IOException {
        if (!isSupportedPlatform()) {
            throw new IOException("RDP 署名は Windows のみ対応です。");
        }
        String script =
                """
                function Get-CertEkuOids([System.Security.Cryptography.X509Certificates.X509Certificate2]$cert) {
                  $oids = New-Object System.Collections.Generic.List[string]
                  foreach ($eku in $cert.EnhancedKeyUsageList) {
                    if ($eku.Value) { [void]$oids.Add($eku.Value) }
                  }
                  if ($oids.Count -gt 0) { return $oids }
                  foreach ($ext in $cert.Extensions) {
                    if ($ext.Oid.Value -ne '2.5.29.37') { continue }
                    $ekuExt = New-Object System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension -ArgumentList $ext, $false
                    foreach ($oid in $ekuExt.EnhancedKeyUsages) {
                      if ($oid.Value) { [void]$oids.Add($oid.Value) }
                    }
                  }
                  return $oids
                }
                function Test-IsRdpSigningCertificate([System.Security.Cryptography.X509Certificates.X509Certificate2]$cert) {
                  if (-not $cert.HasPrivateKey -or $cert.NotAfter -le (Get-Date)) { return $false }
                  $codeSign = '1.3.6.1.5.5.7.3.3'
                  $serverAuth = '1.3.6.1.5.5.7.3.1'
                  $eku = @(Get-CertEkuOids $cert)
                  if ($eku -contains $codeSign) { return $true }
                  $kuExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.15' } | Select-Object -First 1
                  if (-not $kuExt) { return $false }
                  $ku = New-Object System.Security.Cryptography.X509Certificates.X509KeyUsageExtension -ArgumentList $kuExt, $false
                  $ds = [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature
                  if (-not ($ku.KeyUsages -band $ds)) { return $false }
                  if ($eku.Count -eq 0) { return $true }
                  if ($eku -contains $serverAuth) { return $false }
                  return $false
                }
                function Get-UsageLabel([System.Security.Cryptography.X509Certificates.X509Certificate2]$cert) {
                  $codeSign = '1.3.6.1.5.5.7.3.3'
                  foreach ($oid in (Get-CertEkuOids $cert)) {
                    if ($oid -eq $codeSign) { return 'コード署名' }
                  }
                  return 'デジタル署名'
                }
                [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
                $stores = @('Cert:\\CurrentUser\\My', 'Cert:\\LocalMachine\\My')
                $skipped = 0
                $seen = New-Object System.Collections.Generic.HashSet[string]
                foreach ($storePath in $stores) {
                  $label = if ($storePath -like '*CurrentUser*') { 'CurrentUser' } else { 'LocalMachine' }
                  Get-ChildItem $storePath -ErrorAction SilentlyContinue |
                    Where-Object { $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
                    ForEach-Object {
                      if (-not (Test-IsRdpSigningCertificate $_)) {
                        $script:skipped++
                        return
                      }
                      $t = ($_.Thumbprint -replace '\\s','').ToUpperInvariant()
                      if (-not $seen.Add($t)) { return }
                      $s = $_.Subject -replace '\\|','/'
                      $d = $_.NotAfter.ToString('yyyy-MM-dd')
                      $u = Get-UsageLabel $_
                      Write-Output ($t + '|' + $s + '|' + $label + '|' + $d + '|' + $u)
                    }
                }
                Write-Output ('#SKIPPED=' + $skipped)
                """;
        CommandResult result = runPowerShell(script);
        if (!result.success() && result.output().isBlank()) {
            throw new IOException("証明書一覧の取得に失敗しました（終了コード " + result.exitCode() + "）。");
        }
        List<SigningCertificate> out = new ArrayList<>();
        int skipped = 0;
        for (String line : result.output().split("\\R")) {
            if (line != null && line.startsWith("#SKIPPED=")) {
                try {
                    skipped = Integer.parseInt(line.substring("#SKIPPED=".length()).strip());
                } catch (NumberFormatException ignored) {
                    // ignore
                }
                continue;
            }
            SigningCertificate cert = parseCertificateLine(line);
            if (cert != null) {
                out.add(cert);
            }
        }
        return new CertificateListResult(List.copyOf(out), skipped);
    }

    /**
     * RDP ファイル署名用の自己署名証明書を CurrentUser\\My に作成する。
     *
     * @param commonName 例: {@code 湖南工場 RDP Signing}
     */
    public static SigningCertificate createRdpSigningCertificate(String commonName) throws IOException {
        if (!isSupportedPlatform()) {
            throw new IOException("RDP 署名は Windows のみ対応です。");
        }
        String cn = commonName != null && !commonName.isBlank() ? commonName.strip() : "RDP Signing";
        if (cn.contains("|")) {
            throw new IllegalArgumentException("Common Name に | は使えません。");
        }
        String subject = cn.startsWith("CN=") ? cn : "CN=" + cn;
        String escapedSubject = subject.replace("'", "''");
        String script =
                """
                [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
                $subject = '%s'
                $cert = New-SelfSignedCertificate `
                  -Type CodeSigningCert `
                  -Subject $subject `
                  -CertStoreLocation 'Cert:\\CurrentUser\\My' `
                  -KeyAlgorithm RSA `
                  -KeyLength 2048 `
                  -HashAlgorithm SHA256 `
                  -NotAfter (Get-Date).AddYears(3) `
                  -KeyExportPolicy Exportable
                $cert = Get-Item -Path ('Cert:\\CurrentUser\\My\\' + $cert.Thumbprint)
                $t = ($cert.Thumbprint -replace '\\s','').ToUpperInvariant()
                $s = $cert.Subject -replace '\\|','/'
                $d = $cert.NotAfter.ToString('yyyy-MM-dd')
                Write-Output ($t + '|' + $s + '|CurrentUser|' + $d + '|コード署名')
                """
                        .formatted(escapedSubject);
        CommandResult result = runPowerShell(script);
        if (!result.success()) {
            throw new IOException(
                    "RDP 署名用証明書の作成に失敗しました（終了コード "
                            + result.exitCode()
                            + "）: "
                            + abbreviate(result.output(), 400));
        }
        for (String line : result.output().split("\\R")) {
            SigningCertificate cert = parseCertificateLine(line);
            if (cert != null) {
                return cert;
            }
        }
        throw new IOException("RDP 署名用証明書を作成しましたが、結果を読み取れませんでした。");
    }

    /** rdpsign 失敗時のよくある原因を日本語で補足する。 */
    public static String explainSignFailure(CommandResult result) {
        if (result == null) {
            return "";
        }
        if (result.success()) {
            return result.output();
        }
        String out = result.output() != null ? result.output() : "";
        String lower = out.toLowerCase(Locale.ROOT);
        StringBuilder sb = new StringBuilder(out);
        if (lower.contains("0x8007000d") || lower.contains("unable to use the certificate")) {
            sb.append(
                    """

                    【補足】この証明書は RDP ファイル署名に使えません。
                    rdpsign.exe には「コード署名」用途（EKU 1.3.6.1.5.5.7.3.3）の証明書が必要です。
                    SSL/TLS・メール・スマートカード用証明書は署名に失敗します。
                    ウィザードの「RDP署名用証明書を作成」で正しい証明書を作成してください。""");
        } else if (lower.contains("0x80092004") || lower.contains("unable to find") || lower.contains("unable locate")) {
            sb.append(
                    """

                    【補足】指定サムプリントの証明書が Personal ストアに見つかりません。
                    秘密鍵付きでインポートされているか、ストア（CurrentUser / LocalMachine）を確認してください。""");
        } else if (lower.contains("0x80070003") || lower.contains("path not found") || lower.contains("指定されたパス")) {
            sb.append(
                    """

                    【補足】rdpsign.exe が .rdp ファイルのパスを開けません（0x80070003）。
                    本署名は %TEMP%\\PM-AI-rdp-sign 上の ASCII 作業ファイルに対して UAC 昇格で実行し、
                    成功後にリポジトリルートへコピーします。作業ファイルが存在するか確認してください。""");
        } else if (lower.contains("0x80070005") || lower.contains("access is denied") || lower.contains("アクセスが拒否")) {
            sb.append(
                    """

                    【補足】ファイルの上書き権限がありません（0x80070005）。
                    本署名は UAC 昇格 PowerShell 内で ProgramData\\PM-AI\\rdp-sign に作業ファイルを作成して rdpsign します。
                    Windows セキュリティの「ウイルスと脅威の防止」→「ランサムウェアの防止」→
                    「フォルダー アクセスの制御」で rdpsign.exe の許可、または PM-AI 作業フォルダの除外を確認してください。""");
        } else if (result.exitCode() == UAC_CANCELLED_EXIT_CODE) {
            sb.append(

                    """

                    【補足】UAC（管理者権限の確認）がキャンセルされました。本署名を完了するには UAC を許可してください。""");
        }
        return sb.toString().strip();
    }

    static SigningCertificate parseCertificateLine(String line) {
        if (line == null || line.isBlank() || line.startsWith("#")) {
            return null;
        }
        String[] parts = line.split("\\|", 5);
        if (parts.length < 4) {
            return null;
        }
        try {
            String thumb = normalizeThumbprintSha1(parts[0]);
            String usage = parts.length >= 5 ? parts[4].strip() : "—";
            return new SigningCertificate(thumb, parts[1].strip(), parts[2].strip(), parts[3].strip(), usage);
        } catch (IllegalArgumentException ex) {
            return null;
        }
    }

    private static String abbreviate(String text, int max) {
        if (text == null || text.length() <= max) {
            return text != null ? text : "";
        }
        return text.substring(0, max) + "…";
    }

    /**
     * OneDrive / Dropbox 等、rdpsign が上書きしにくい場所か。
     */
    public static boolean isRestrictedSigningLocation(Path rdpProfile) {
        if (rdpProfile == null) {
            return false;
        }
        String normalized = rdpProfile.toAbsolutePath().normalize().toString().toLowerCase(Locale.ROOT);
        return normalized.contains("onedrive")
                || normalized.contains("dropbox")
                || normalized.contains("icloud")
                || normalized.contains("\\my drive\\")
                || normalized.contains("/my drive/");
    }

    /** 署名済み .rdp の出力先ディレクトリ（リポジトリルート）。 */
    public static Path resolveSignedOutputDir(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of()).toAbsolutePath().normalize();
    }

    /** 元 .rdp から作る署名済み新規出力パス（リポジトリルート直下。元ファイルは書き換えない）。 */
    public static Path resolveSignedOutputPath(Path sourceRdp, Map<String, String> ui) {
        Path source = sourceRdp.toAbsolutePath().normalize();
        String base = signedOutputBaseName(source.getFileName().toString());
        return resolveSignedOutputDir(ui)
                .resolve(base + SIGNED_OUTPUT_SUFFIX)
                .toAbsolutePath()
                .normalize();
    }

    /**
     * {@code *.pm-ai-signed.rdp} を入力しても {@code Default.pm-ai-signed.rdp} に正規化する。
     * 二重サフィックス（{@code Default.pm-ai-signed.pm-ai-signed.rdp}）も解消する。
     */
    static String signedOutputBaseName(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return "profile";
        }
        String base = fileName.strip();
        while (base.endsWith(SIGNED_OUTPUT_SUFFIX)) {
            base = base.substring(0, base.length() - SIGNED_OUTPUT_SUFFIX.length());
        }
        int dot = base.lastIndexOf('.');
        return dot > 0 ? base.substring(0, dot) : base;
    }

    /**
     * 環境変数等で設定された .rdp を起動・署名に使う前に正規化する。
     * 二重サフィックスや {@code .pm-ai-signed.rdp} 再指定時に、正しい署名済み1ファイルへ寄せる。
     * ディレクトリや Windows 既定 {@code Default.rdp} が指定されたときは、署名済みプロファイルを探索する。
     */
    public static Path resolvePreferredSignedProfilePath(Path configured, Map<String, String> ui) {
        if (configured == null) {
            throw new IllegalArgumentException("configured");
        }
        Path configuredAbs = configured.toAbsolutePath().normalize();
        Path signedDir = resolveSignedOutputDir(ui);
        if (Files.isDirectory(configuredAbs)) {
            Path found = findSignedProfileNear(configuredAbs, signedDir);
            if (found != null) {
                return found;
            }
            return configuredAbs;
        }
        if (AppPaths.isWindowsDefaultRdpProfile(configuredAbs)) {
            Path found = findSignedProfileNear(configuredAbs.getParent(), signedDir);
            if (found != null) {
                return found;
            }
        }
        Path canonical = resolveSignedOutputPath(configuredAbs, ui);
        try {
            if (Files.isRegularFile(canonical) && isSigned(canonical)) {
                return canonical;
            }
            if (Files.isRegularFile(configuredAbs) && isSigned(configuredAbs)) {
                return configuredAbs;
            }
        } catch (IOException ignored) {
            // fall through
        }
        if (Files.isRegularFile(canonical)) {
            return canonical;
        }
        Path nearby = findSignedProfileNear(configuredAbs.getParent(), signedDir);
        if (nearby != null) {
            return nearby;
        }
        return configuredAbs;
    }

    /**
     * UI 環境変数から起動に使う .rdp を決める。未設定時はリポジトリの署名済みプロファイルを探す（Default.rdp には落とさない）。
     */
    public static Path resolvePreferredSignedProfilePathFromUi(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.getOrDefault(AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE, "");
        Path signedDir = resolveSignedOutputDir(u);
        Path configured =
                raw == null || raw.isBlank()
                        ? signedDir.resolve("Default.rdp")
                        : Path.of(raw.strip());
        return resolvePreferredSignedProfilePath(configured, u);
    }

    static Path findSignedProfileNear(Path directory, Path signedOutputDir) {
        Path found = findDefaultSignedProfileIn(directory);
        if (found != null) {
            return found;
        }
        return findDefaultSignedProfileIn(signedOutputDir);
    }

    private static Path findDefaultSignedProfileIn(Path directory) {
        if (directory == null || !Files.isDirectory(directory)) {
            return null;
        }
        Path named = directory.resolve("Default" + SIGNED_OUTPUT_SUFFIX);
        if (Files.isRegularFile(named)) {
            return named.toAbsolutePath().normalize();
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory, "*" + SIGNED_OUTPUT_SUFFIX)) {
            for (Path p : stream) {
                if (Files.isRegularFile(p)) {
                    return p.toAbsolutePath().normalize();
                }
            }
        } catch (IOException ignored) {
            return null;
        }
        return null;
    }

    /** {@link #resolveSignedOutputPath(Path, Map)} の ui 未指定版。 */
    public static Path resolveSignedOutputPath(Path sourceRdp) {
        return resolveSignedOutputPath(sourceRdp, Map.of());
    }

    /** rdpsign の in-place 署名作業ファイル名（ASCII のみ。昇格プロセスから確実に開ける）。 */
    public static final String RDPSIGN_WORK_FILENAME = "pm-ai-signing.rdp";

    /** UAC 昇格 rdpsign 成功後に Java が読む署名済み出力（ProgramData 配下・ASCII パス）。 */
    public static final String RDPSIGN_ELEVATED_OUTPUT_FILENAME = "pm-ai-signed-output.rdp";

    /** UAC 昇格 rdpsign の作業・出力フォルダ（{@code %ProgramData%\\PM-AI\\rdp-sign}）。 */
    public static Path resolveRdpsignElevatedWorkDir() throws IOException {
        String programData = System.getenv("ProgramData");
        Path base =
                programData != null && !programData.isBlank()
                        ? Path.of(programData.trim())
                        : Path.of("C:\\ProgramData");
        Path dir = base.resolve("PM-AI").resolve("rdp-sign");
        Files.createDirectories(dir);
        return dir.toAbsolutePath().normalize();
    }

    /** UAC 昇格 rdpsign が書き出す署名済み .rdp（通常権限 Java がリポジトリへコピーする）。 */
    public static Path resolveRdpsignElevatedOutputPath() throws IOException {
        return resolveRdpsignElevatedWorkDir().resolve(RDPSIGN_ELEVATED_OUTPUT_FILENAME);
    }

    /** 非昇格側のステージング（元 .rdp のコピー先）。 */
    public static Path resolveRdpsignWorkDir(Map<String, String> ui) throws IOException {
        return resolveRdpsignLogDir();
    }

    /** {@link #resolveRdpsignWorkDir(Map)} の ui 未指定版。 */
    public static Path resolveRdpsignWorkDir() throws IOException {
        return resolveRdpsignWorkDir(Map.of());
    }

    /**
     * 署名対象を解決する。
     *
     * <p>元 .rdp は読み取り専用として触らない。未署名コピーは {@link #resolveRdpsignLogDir()} に置く。
     * 本署名は UAC 昇格 PowerShell が ProgramData 上でコピー・rdpsign・出力まで行い、
     * 成功後 {@link #resolveSignedOutputPath} のリポジトリルートへコピーする。
     */
    public static SigningTarget prepareSigningTarget(Path rdpProfile, Map<String, String> ui) throws IOException {
        Path source = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
        Path profilePath = resolveSignedOutputPath(source, ui);
        Path signingPath = prepareSigningWorkFile(source);
        boolean createsNew = !source.toAbsolutePath().normalize().equals(profilePath);
        return new SigningTarget(signingPath, profilePath, source, createsNew);
    }

    /** {@link #prepareSigningTarget(Path, Map)} の ui 未指定版。 */
    public static SigningTarget prepareSigningTarget(Path rdpProfile) throws IOException {
        return prepareSigningTarget(rdpProfile, Map.of());
    }

    /** 元 .rdp を %TEMP%\\PM-AI-rdp-sign 上の ASCII 作業ファイルへコピーする。 */
    static Path prepareSigningWorkFile(Path source) throws IOException {
        Path workDir = resolveRdpsignLogDir();
        Path work = workDir.resolve(RDPSIGN_WORK_FILENAME);
        Files.copy(source, work, StandardCopyOption.REPLACE_EXISTING);
        clearWindowsReadOnly(work);
        if (!Files.isWritable(work)) {
            throw new IOException("作業ファイルに書き込めません: " + work);
        }
        if (!Files.isWritable(workDir)) {
            throw new IOException("作業フォルダに書き込めません: " + workDir);
        }
        return work.toAbsolutePath().normalize();
    }

    static Path prepareUnsignedCopyAtOutput(Path source, Path outputPath) throws IOException {
        Path parent = outputPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.copy(source, outputPath, StandardCopyOption.REPLACE_EXISTING);
        clearWindowsReadOnly(outputPath);
        if (!Files.isWritable(outputPath)) {
            throw new IOException("出力ファイルに書き込めません: " + outputPath);
        }
        if (parent != null && !Files.isWritable(parent)) {
            throw new IOException("出力フォルダに書き込めません: " + parent);
        }
        return outputPath.toAbsolutePath().normalize();
    }

    static void clearWindowsReadOnly(Path path) throws IOException {
        if (path == null || !Files.exists(path)) {
            return;
        }
        try {
            Files.setAttribute(path, "dos:readonly", false);
        } catch (IOException | UnsupportedOperationException ignored) {
            // ignore
        }
    }

    /**
     * 署名テスト（{@code /l}）。元ファイルは変更しない。
     */
    public static CommandResult testSign(Path rdpProfile, String thumbprintSha1) throws IOException {
        return attemptSign(rdpProfile, thumbprintSha1, true, false).result();
    }

    public static CommandResult testSign(Path rdpProfile, String thumbprintSha1, Map<String, String> ui)
            throws IOException {
        return attemptSign(rdpProfile, thumbprintSha1, true, false, ui).result();
    }

    /**
     * 署名を実行する。{@code backupBeforeSign} が true のとき {@code .unsigned-タイムスタンプ.bak} を作成する。
     */
    public static CommandResult sign(Path rdpProfile, String thumbprintSha1, boolean backupBeforeSign)
            throws IOException {
        return attemptSign(rdpProfile, thumbprintSha1, false, backupBeforeSign).result();
    }

    public static CommandResult sign(
            Path rdpProfile, String thumbprintSha1, boolean backupBeforeSign, Map<String, String> ui)
            throws IOException {
        return attemptSign(rdpProfile, thumbprintSha1, false, backupBeforeSign, ui).result();
    }

    /** 同期フォルダ回避を含む署名／テスト署名。 */
    public static SignAttemptResult attemptSign(
            Path rdpProfile, String thumbprintSha1, boolean testOnly, boolean backupBeforeSign)
            throws IOException {
        return attemptSign(rdpProfile, thumbprintSha1, testOnly, backupBeforeSign, Map.of());
    }

    /** リポジトリルート出力を含む署名／テスト署名。 */
    public static SignAttemptResult attemptSign(
            Path rdpProfile,
            String thumbprintSha1,
            boolean testOnly,
            boolean backupBeforeSign,
            Map<String, String> ui)
            throws IOException {
        SigningTarget target = prepareSigningTarget(rdpProfile, ui);
        Path backupSource = target.profilePath();
        if (backupBeforeSign && !testOnly && Files.isRegularFile(backupSource)) {
            backupUnsignedCopy(backupSource);
        }
        try {
            CommandResult result = runRdpsign(target.signingPath(), thumbprintSha1, testOnly, !testOnly);
            if (result.success() && !testOnly) {
                Path signedSource = resolveRdpsignElevatedOutputPath();
                Path finalized = writeNewSignedProfile(signedSource, target.profilePath());
                clearWindowsReadOnly(finalized);
                SigningTarget finalizedTarget =
                        new SigningTarget(
                                finalized,
                                finalized,
                                target.sourcePath(),
                                target.createsNewProfileFile());
                return new SignAttemptResult(result, finalizedTarget);
            }
            return new SignAttemptResult(result, target);
        } finally {
            try {
                Files.deleteIfExists(target.signingPath());
                Files.deleteIfExists(resolveRdpsignElevatedOutputPath());
            } catch (IOException ignored) {
                // ignore
            }
        }
    }

    /** 署名済み内容を新規出力 .rdp へ書き込む（元 .rdp は変更しない）。 */
    static Path writeNewSignedProfile(Path signedWork, Path outputPath) throws IOException {
        Objects.requireNonNull(signedWork, "signedWork");
        Objects.requireNonNull(outputPath, "outputPath");
        Path parent = outputPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        if (Files.isRegularFile(outputPath)) {
            clearWindowsReadOnly(outputPath);
            try {
                Files.delete(outputPath);
            } catch (IOException ignored) {
                // REPLACE_EXISTING で上書きする
            }
        }
        clearWindowsReadOnly(signedWork);
        Files.copy(signedWork, outputPath, StandardCopyOption.REPLACE_EXISTING);
        clearWindowsReadOnly(outputPath);
        try {
            Files.deleteIfExists(signedWork);
        } catch (IOException ignored) {
            // ignore
        }
        return outputPath.toAbsolutePath().normalize();
    }

    /** GPO「信頼できる .rdp 発行元」用 PowerShell スクリプト（管理者・HKLM）。 */
    public static String buildTrustedPublisherRegistryScript(String thumbprintSha1) {
        return buildTrustedPublisherRegistryScript(thumbprintSha1, true);
    }

    /** 信頼ポリシー登録 PowerShell（{@code machineWide=true} で HKLM、false で HKCU）。 */
    public static String buildTrustedPublisherRegistryScript(String thumbprintSha1, boolean machineWide) {
        String thumb = normalizeThumbprintSha1(thumbprintSha1);
        String hive = machineWide ? "HKLM" : "HKCU";
        String note =
                machineWide
                        ? "管理者 PowerShell で実行 — 信頼できる .rdp 発行元（全ユーザー / HKLM）"
                        : "現在ユーザー向け — 信頼できる .rdp 発行元（HKCU・管理者不要）";
        return """
                # %s
                $path = '%s:\\SOFTWARE\\Policies\\Microsoft\\Windows NT\\Terminal Services'
                New-Item -Path $path -Force | Out-Null
                New-ItemProperty -Path $path -Name 'AllowSignedFiles' -PropertyType DWord -Value 1 -Force | Out-Null
                New-ItemProperty -Path $path -Name 'TrustedCertThumbprints' -PropertyType String -Value '%s' -Force | Out-Null
                """
                .formatted(note, hive, thumb);
    }

    /**
     * 信頼ポリシーを HKCU と HKLM の両方へ書き込む（推奨）。
     * HKCU は通常権限、HKLM は UAC 昇格。
     */
    public static CommandResult applyTrustedPublisherPolicyAllScopes(String thumbprintSha1) throws IOException {
        CommandResult hkcu = applyTrustedPublisherPolicy(thumbprintSha1, false);
        if (!hkcu.success()) {
            return hkcu;
        }
        CommandResult hklm = applyTrustedPublisherPolicy(thumbprintSha1, true);
        if (!hklm.success()) {
            return new CommandResult(
                    hklm.exitCode(),
                    ("HKCU は適用済み。HKLM: " + hklm.output()).strip());
        }
        return new CommandResult(0, "HKCU と HKLM の両方に信頼設定を適用しました。");
    }

    /**
     * 信頼ポリシーをレジストリへ書き込む。
     *
     * @param machineWide true のとき HKLM（UAC 昇格）、false のとき HKCU（現在ユーザー）
     */
    public static CommandResult applyTrustedPublisherPolicy(String thumbprintSha1, boolean machineWide)
            throws IOException {
        if (!isSupportedPlatform()) {
            throw new IOException("RDP 信頼設定は Windows のみ対応です。");
        }
        String script = buildTrustedPublisherRegistryScript(thumbprintSha1, machineWide);
        if (machineWide) {
            return runPowerShellScriptElevated(script);
        }
        Path scriptFile = resolveRdpsignLogDir().resolve("trust-policy-hkcu.ps1");
        writePowerShellScript(scriptFile, script.strip() + System.lineSeparator());
        return runPowerShellFile(scriptFile);
    }

    static CommandResult runPowerShellScriptElevated(String scriptContent) throws IOException {
        Path logDir = resolveRdpsignLogDir();
        Path innerScript = logDir.resolve("trust-policy-inner.ps1");
        Path launcherScript = logDir.resolve("trust-policy-launcher.ps1");
        writePowerShellScript(innerScript, scriptContent.strip() + System.lineSeparator());
        writePowerShellScript(launcherScript, buildElevatedRdpsignLauncherScript(innerScript));
        return runPowerShellFile(launcherScript);
    }

    static void backupUnsignedCopy(Path rdpFile) throws IOException {
        String stamp = LocalDateTime.now().format(BACKUP_STAMP);
        Path parent = rdpFile.getParent();
        String base = rdpFile.getFileName().toString();
        String backupName = base + ".unsigned-" + stamp + ".bak";
        Path backup = parent != null ? parent.resolve(backupName) : Path.of(backupName);
        Files.copy(rdpFile, backup, StandardCopyOption.REPLACE_EXISTING);
    }

    /** rdpsign ログ／UAC 昇格スクリプト用（%TEMP%\\PM-AI-rdp-sign）。 */
    static Path resolveRdpsignLogDir() throws IOException {
        String temp = System.getenv("TEMP");
        Path dir =
                temp != null && !temp.isBlank()
                        ? Path.of(temp.trim(), "PM-AI-rdp-sign")
                        : Path.of(System.getProperty("java.io.tmpdir"), "PM-AI-rdp-sign");
        Files.createDirectories(dir);
        return dir.toAbsolutePath().normalize();
    }

    static String escapePowerShellSingleQuoted(String value) {
        if (value == null) {
            return "''";
        }
        return "'" + value.replace("'", "''") + "'";
    }

    /** UAC 昇格後に実行する rdpsign 用 PowerShell スクリプト本文（作業ファイルは昇格プロセス内で ProgramData に作成）。 */
    static String buildElevatedRdpsignInnerScript(
            Path rdpsignExe,
            String thumbprintSha1,
            Path stagingFile,
            Path elevatedWorkDir,
            Path elevatedOutputFile,
            Path logFile,
            Path exitFile) {
        return """
                $ErrorActionPreference = 'Continue'
                $logPath = %s
                $exitPath = %s
                $staging = %s
                $workDir = %s
                $output = %s
                $code = 1
                function Clear-FileReadOnly([string]$Path) {
                  if (-not (Test-Path -LiteralPath $Path)) { return }
                  try {
                    [System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal)
                  } catch {}
                }
                try {
                  if (-not (Test-Path -LiteralPath $staging)) {
                    throw "Staging file not found: $staging"
                  }
                  New-Item -ItemType Directory -Force -Path $workDir | Out-Null
                  $work = Join-Path $workDir ('pm-ai-signing-' + [guid]::NewGuid().ToString('N') + '.rdp')
                  Copy-Item -LiteralPath $staging -Destination $work -Force
                  Clear-FileReadOnly $work
                  Set-Location -LiteralPath $workDir
                  $logLines = @()
                  & %s /sha256 %s /v $work 2>&1 | ForEach-Object { $logLines += $_ }
                  $logLines | Out-File -FilePath $logPath -Encoding utf8
                  $code = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { 0 }
                  if ($code -eq 0) {
                    Copy-Item -LiteralPath $work -Destination $output -Force
                    Clear-FileReadOnly $output
                  }
                  Remove-Item -LiteralPath $work -Force -ErrorAction SilentlyContinue
                } catch {
                  $_ | Out-File -FilePath $logPath -Append -Encoding utf8
                  $code = 1
                }
                Set-Content -Path $exitPath -Value $code -Encoding ascii -NoNewline
                exit $code
                """
                .formatted(
                        escapePowerShellSingleQuoted(logFile.toString()),
                        escapePowerShellSingleQuoted(exitFile.toString()),
                        escapePowerShellSingleQuoted(stagingFile.toString()),
                        escapePowerShellSingleQuoted(elevatedWorkDir.toString()),
                        escapePowerShellSingleQuoted(elevatedOutputFile.toString()),
                        escapePowerShellSingleQuoted(rdpsignExe.toString()),
                        escapePowerShellSingleQuoted(thumbprintSha1));
    }

    static String buildElevatedRdpsignLauncherScript(Path innerScript) {
        return """
                $inner = %s
                $p = Start-Process -FilePath 'powershell.exe' `
                  -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File',$inner) `
                  -Verb RunAs -Wait -PassThru
                if ($null -eq $p) { exit %d }
                exit $p.ExitCode
                """
                .formatted(escapePowerShellSingleQuoted(innerScript.toString()), UAC_CANCELLED_EXIT_CODE);
    }

    static void writePowerShellScript(Path scriptFile, String content) throws IOException {
        byte[] bom = new byte[] {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        byte[] body = content.getBytes(StandardCharsets.UTF_8);
        byte[] withBom = new byte[bom.length + body.length];
        System.arraycopy(bom, 0, withBom, 0, bom.length);
        System.arraycopy(body, 0, withBom, bom.length, body.length);
        Files.write(scriptFile, withBom);
    }

    private static CommandResult runRdpsignElevated(Path stagingFile, String thumbprintSha1) throws IOException {
        Path rdpsign = resolveRdpsignExe();
        Path abs = RemoteDesktopLauncher.validateRdpProfile(stagingFile);
        String thumb = normalizeThumbprintSha1(thumbprintSha1);
        Path elevatedWorkDir = resolveRdpsignElevatedWorkDir();
        Path elevatedOutput = resolveRdpsignElevatedOutputPath();
        Path logDir = resolveRdpsignLogDir();
        Path logFile = logDir.resolve("rdpsign-last.log");
        Path exitFile = logDir.resolve("rdpsign-last.exit");
        Path innerScript = logDir.resolve("rdpsign-elevated-inner.ps1");
        Path launcherScript = logDir.resolve("rdpsign-elevated-launcher.ps1");
        Files.deleteIfExists(logFile);
        Files.deleteIfExists(exitFile);
        Files.deleteIfExists(elevatedOutput);
        writePowerShellScript(
                innerScript,
                buildElevatedRdpsignInnerScript(
                        rdpsign, thumb, abs, elevatedWorkDir, elevatedOutput, logFile, exitFile));
        writePowerShellScript(launcherScript, buildElevatedRdpsignLauncherScript(innerScript));
        CommandResult launcher = runPowerShellFile(launcherScript);
        int exitCode = launcher.exitCode();
        String output = "";
        if (Files.isRegularFile(logFile)) {
            output = Files.readString(logFile, StandardCharsets.UTF_8);
        } else if (!launcher.output().isBlank()) {
            output = launcher.output();
        }
        if (Files.isRegularFile(exitFile)) {
            try {
                exitCode = Integer.parseInt(Files.readString(exitFile, StandardCharsets.US_ASCII).strip());
            } catch (NumberFormatException ignored) {
                // launcher の終了コードを使う
            }
        }
        if (exitCode == UAC_CANCELLED_EXIT_CODE && output.isBlank()) {
            output = "UAC（管理者権限の確認）がキャンセルされました。";
        }
        if (output.isBlank() && !launcher.output().isBlank()) {
            output = launcher.output();
        }
        return new CommandResult(exitCode, output.strip());
    }

    private static CommandResult runRdpsign(
            Path rdpProfile, String thumbprintSha1, boolean testOnly, boolean elevate) throws IOException {
        if (elevate && !testOnly) {
            return runRdpsignElevated(rdpProfile, thumbprintSha1);
        }
        return runRdpsignDirect(rdpProfile, thumbprintSha1, testOnly);
    }

    private static CommandResult runRdpsignDirect(Path rdpProfile, String thumbprintSha1, boolean testOnly)
            throws IOException {
        Path rdpsign = resolveRdpsignExe();
        Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
        String thumb = normalizeThumbprintSha1(thumbprintSha1);
        List<String> cmd = new ArrayList<>();
        cmd.add(rdpsign.toString());
        cmd.add("/sha256");
        cmd.add(thumb);
        if (testOnly) {
            cmd.add("/l");
        }
        cmd.add("/v");
        cmd.add(abs.toString());
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        Path parent = abs.getParent();
        if (parent != null && Files.isDirectory(parent)) {
            pb.directory(parent.toFile());
        }
        return runProcess(pb, windowsConsoleCharset());
    }

    private static CommandResult runProcess(ProcessBuilder pb) throws IOException {
        return runProcess(pb, windowsConsoleCharset());
    }

    private static CommandResult runProcess(ProcessBuilder pb, Charset outputCharset) throws IOException {
        Objects.requireNonNull(pb, "processBuilder");
        Process process = null;
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        try {
            process = pb.start();
            Thread drain = drainStream(process.getInputStream(), buf);
            boolean finished = process.waitFor(PROCESS_TIMEOUT_SEC, TimeUnit.SECONDS);
            joinQuietly(drain, 5_000);
            if (!finished) {
                process.destroyForcibly();
                throw new IOException("コマンドがタイムアウトしました。");
            }
            Charset cs = outputCharset != null ? outputCharset : StandardCharsets.UTF_8;
            String output = buf.toString(cs);
            return new CommandResult(process.exitValue(), output.strip());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("コマンド実行が中断されました。", e);
        } finally {
            if (process != null && process.isAlive()) {
                process.destroyForcibly();
            }
        }
    }

    private static CommandResult runPowerShell(String script) throws IOException {
        List<String> cmd =
                List.of(
                        "powershell.exe",
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-Command",
                        script);
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        return runProcess(pb, StandardCharsets.UTF_8);
    }

    private static CommandResult runPowerShellFile(Path scriptFile) throws IOException {
        List<String> cmd =
                List.of(
                        "powershell.exe",
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        scriptFile.toAbsolutePath().normalize().toString());
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        return runProcess(pb, StandardCharsets.UTF_8);
    }

    private static Charset windowsConsoleCharset() {
        String lang = System.getenv("LANG");
        if (lang != null && lang.toUpperCase(Locale.ROOT).contains("UTF")) {
            return StandardCharsets.UTF_8;
        }
        return Charset.forName("MS932");
    }

    private static Thread drainStream(InputStream in, ByteArrayOutputStream dest) {
        Thread t =
                new Thread(
                        () -> {
                            try (in) {
                                byte[] chunk = new byte[8192];
                                int n;
                                while ((n = in.read(chunk)) >= 0) {
                                    synchronized (dest) {
                                        int room = OUTPUT_CAPTURE_MAX - dest.size();
                                        if (room <= 0) {
                                            continue;
                                        }
                                        dest.write(chunk, 0, Math.min(n, room));
                                    }
                                }
                            } catch (IOException ignored) {
                                // ignore
                            }
                        },
                        "rdp-signer-drain");
        t.setDaemon(true);
        t.start();
        return t;
    }

    private static void joinQuietly(Thread t, long millis) {
        if (t == null) {
            return;
        }
        try {
            t.join(millis);
        } catch (InterruptedException ie) {
            Thread.currentThread().interrupt();
        }
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
    }
}
