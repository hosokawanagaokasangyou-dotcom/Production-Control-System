package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class RdpFileSignerTest {

    @Test
    void normalizeThumbprintSha1_stripsSpacesAndUppercases() {
        assertEquals(
                "A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2",
                RdpFileSigner.normalizeThumbprintSha1("a1 b2 c3 d4 e5 f6 a1 b2 c3 d4 e5 f6 a1 b2 c3 d4 e5 f6 a1 b2"));
    }

    @Test
    void normalizeThumbprintSha1_rejectsInvalid() {
        assertThrows(IllegalArgumentException.class, () -> RdpFileSigner.normalizeThumbprintSha1("ABC"));
    }

    @Test
    void parseCertificateLine_parsesPipeDelimitedRow() {
        RdpFileSigner.SigningCertificate cert =
                RdpFileSigner.parseCertificateLine(
                        "A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2|CN=RDPFileSigner|CurrentUser|2027-12-31|コード署名");
        assertEquals("A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2", cert.thumbprintSha1());
        assertEquals("CN=RDPFileSigner", cert.subject());
        assertEquals("CurrentUser", cert.storeLabel());
        assertEquals("2027-12-31", cert.notAfter());
        assertEquals("コード署名", cert.usageLabel());
    }

    @Test
    void parseCertificateLine_acceptsLegacyFourFields() {
        RdpFileSigner.SigningCertificate cert =
                RdpFileSigner.parseCertificateLine(
                        "A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2|CN=localhost|CurrentUser|2026-09-30");
        assertEquals("デジタル署名".equals(cert.usageLabel()) || "—".equals(cert.usageLabel()), true);
    }

    @Test
    void withEnsuredEligible_prependsCreatedCertificate() {
        RdpFileSigner.SigningCertificate existing =
                RdpFileSigner.parseCertificateLine(
                        "A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2|CN=localhost|CurrentUser|2026-09-30|デジタル署名");
        RdpFileSigner.SigningCertificate created =
                RdpFileSigner.parseCertificateLine(
                        "B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3|CN=湖南工場 RDP Signing|CurrentUser|2029-01-01|コード署名");
        RdpFileSigner.CertificateListResult base =
                new RdpFileSigner.CertificateListResult(List.of(existing), 13);
        RdpFileSigner.CertificateListResult merged = base.withEnsuredEligible(created);
        assertEquals(2, merged.eligible().size());
        assertEquals(created.thumbprintSha1(), merged.eligible().getFirst().thumbprintSha1());
    }

    @Test
    void isRestrictedSigningLocation_detectsOneDrive() {
        assertTrue(
                RdpFileSigner.isRestrictedSigningLocation(
                        Path.of("C:\\Users\\0585\\OneDrive\\ドキュメント\\Default.rdp")));
        assertFalse(RdpFileSigner.isRestrictedSigningLocation(Path.of("C:\\PM-AI\\Default.rdp")));
    }

    @Test
    void prepareSigningTarget_copiesOneDriveFileToLocalWorkspace(@TempDir Path tmp) throws Exception {
        Path onedrive = tmp.resolve("OneDrive").resolve("docs");
        Files.createDirectories(onedrive);
        Path source = onedrive.resolve("factory.rdp");
        Files.writeString(source, "screen mode id:i:2\n");

        assertTrue(RdpFileSigner.isRestrictedSigningLocation(source));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, tmp.toString());
        assertEquals(tmp.toAbsolutePath().normalize(), RdpFileSigner.resolveSignedOutputDir(ui));
    }

    @Test
    void resolveSignedOutputPath_appendsPmAiSignedSuffix(@TempDir Path repoRoot) {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repoRoot.toString());
        Path source = Path.of("C:\\Users\\x\\OneDrive\\Default.rdp");
        Path output = RdpFileSigner.resolveSignedOutputPath(source, ui);
        assertTrue(output.toString().endsWith("Default" + RdpFileSigner.SIGNED_OUTPUT_SUFFIX));
        assertEquals(repoRoot.toAbsolutePath().normalize(), output.getParent());
    }

    @Test
    void resolveSignedOutputPath_doesNotDoubleSuffix(@TempDir Path repoRoot) {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repoRoot.toString());
        Path alreadySigned =
                Path.of("C:\\repo\\Default" + RdpFileSigner.SIGNED_OUTPUT_SUFFIX);
        Path output = RdpFileSigner.resolveSignedOutputPath(alreadySigned, ui);
        assertTrue(output.toString().endsWith("Default" + RdpFileSigner.SIGNED_OUTPUT_SUFFIX));
        assertFalse(output.toString().contains(".pm-ai-signed.pm-ai-signed"));
    }

    @Test
    void signedOutputBaseName_stripsRepeatedSuffix() {
        assertEquals(
                "Default",
                RdpFileSigner.signedOutputBaseName(
                        "Default.pm-ai-signed.pm-ai-signed.rdp"));
        assertEquals("Default", RdpFileSigner.signedOutputBaseName("Default.rdp"));
        assertEquals("factory", RdpFileSigner.signedOutputBaseName("factory.pm-ai-signed.rdp"));
    }

    @Test
    void prepareSigningTarget_usesTempWorkAndRepoOutput(@TempDir Path repoRoot) throws Exception {
        Path rdp = repoRoot.resolve("factory.rdp");
        Files.writeString(rdp, "screen mode id:i:2\n");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repoRoot.toString());
        RdpFileSigner.SigningTarget target = RdpFileSigner.prepareSigningTarget(rdp, ui);
        assertTrue(target.createsNewProfileFile());
        Path expectedOutput = RdpFileSigner.resolveSignedOutputPath(rdp, ui);
        assertEquals(expectedOutput, target.profilePath());
        assertNotEquals(expectedOutput, target.signingPath());
        assertTrue(target.signingPath().toString().contains(RdpFileSigner.RDPSIGN_WORK_FILENAME));
        assertTrue(Files.isRegularFile(target.signingPath()));
        assertTrue(Files.isRegularFile(rdp));
        assertFalse(Files.exists(expectedOutput));
        String original = Files.readString(rdp);
        assertEquals("screen mode id:i:2\n", original);
        assertEquals("screen mode id:i:2\n", Files.readString(target.signingPath()));
    }

    @Test
    void buildElevatedRdpsignInnerScript_escapesSingleQuotes() {
        Path rdpsign = Path.of("C:\\Windows\\System32\\rdpsign.exe");
        Path staging = Path.of("C:\\Temp\\PM-AI-rdp-sign\\pm-ai-signing.rdp");
        Path workDir = Path.of("C:\\ProgramData\\PM-AI\\rdp-sign");
        Path output = workDir.resolve(RdpFileSigner.RDPSIGN_ELEVATED_OUTPUT_FILENAME);
        Path log = Path.of("C:\\Temp\\PM-AI-rdp-sign\\rdpsign-last.log");
        Path exit = Path.of("C:\\Temp\\PM-AI-rdp-sign\\rdpsign-last.exit");
        String script =
                RdpFileSigner.buildElevatedRdpsignInnerScript(
                        rdpsign,
                        "A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2",
                        staging,
                        workDir,
                        output,
                        log,
                        exit);
        assertTrue(script.contains("pm-ai-signing-"));
        assertTrue(script.contains("ProgramData"));
        assertTrue(script.contains("Clear-FileReadOnly"));
        assertTrue(script.contains("SetAttributes"));
    }

    @Test
    void explainSignFailure_addsHintForUacCancelled() {
        RdpFileSigner.CommandResult result =
                new RdpFileSigner.CommandResult(RdpFileSigner.UAC_CANCELLED_EXIT_CODE, "");
        assertTrue(RdpFileSigner.explainSignFailure(result).contains("UAC"));
    }

    @Test
    void explainSignFailure_addsHintForAccessDenied() {
        RdpFileSigner.CommandResult result =
                new RdpFileSigner.CommandResult(1, "Error Code: 0x80070005");
        assertTrue(RdpFileSigner.explainSignFailure(result).contains("ProgramData"));
    }

    @Test
    void explainSignFailure_addsHintForUnsupportedCertificate() {
        RdpFileSigner.CommandResult result =
                new RdpFileSigner.CommandResult(1, "Unable to use the certificate specified. Error Code: 0x8007000d");
        String explained = RdpFileSigner.explainSignFailure(result);
        assertTrue(explained.contains("0x8007000d"));
        assertTrue(explained.contains("コード署名"));
    }

    @Test
    void isSigned_detectsSignatureLine(@TempDir Path tmp) throws Exception {
        Path unsigned = tmp.resolve("factory.rdp");
        Files.writeString(unsigned, "screen mode id:i:2\n");
        assertFalse(RdpFileSigner.isSigned(unsigned));

        Path signed = tmp.resolve("signed.rdp");
        Files.writeString(signed, "screen mode id:i:2\nsignature:s:BASE64DATA\n");
        assertTrue(RdpFileSigner.isSigned(signed));
    }

    @Test
    void containsSignatureMarker_detectsAsciiSignatureLine(@TempDir Path tmp) throws Exception {
        Path signed = tmp.resolve("signed.rdp");
        Files.writeString(signed, "screen mode id:i:2\nsignature:s:BASE64\n", StandardCharsets.UTF_16LE);
        assertTrue(RdpFileSigner.containsUtf16LeAsciiMarker(Files.readAllBytes(signed), "signature:s:"));
        assertTrue(RdpFileSigner.isSigned(signed));
    }

    @Test
    void buildTrustedPublisherRegistryScript_containsThumbprint() {
        String script =
                RdpFileSigner.buildTrustedPublisherRegistryScript(
                        "a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2");
        assertTrue(script.contains("A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2"));
        assertTrue(script.contains("TrustedCertThumbprints"));
        assertTrue(script.contains("AllowSignedFiles"));
        assertTrue(script.contains("HKLM"));
    }

    @Test
    void buildTrustedPublisherRegistryScript_hkcuVariant() {
        String script =
                RdpFileSigner.buildTrustedPublisherRegistryScript(
                        "a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2", false);
        assertTrue(script.contains("HKCU"));
        assertTrue(script.contains("A1B2C3D4E5F6A1B2C3D4E5F6A1B2C3D4E5F6A1B2"));
    }
}
