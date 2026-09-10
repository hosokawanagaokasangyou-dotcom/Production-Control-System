package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;

/**
 * Gemini 資格情報ファイル用。中身・鍵を出さず固定文言のみ返す。
 */
public final class GeminiCredentialsConflictDiffSummarizer implements ConflictDiffSummarizer {

    @Override
    public String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            List<Path> mismatched) {
        return "認証ファイルが更新または置換されています";
    }
}
