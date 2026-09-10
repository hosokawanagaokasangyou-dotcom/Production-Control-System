package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/** バイナリ／表形式など、ファイル単位の変更を業務語で要約する。 */
public final class NamedFileConflictDiffSummarizer implements ConflictDiffSummarizer {

    private final String screenLabel;

    public NamedFileConflictDiffSummarizer(String screenLabel) {
        this.screenLabel = screenLabel != null ? screenLabel : "ファイル";
    }

    @Override
    public String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            List<Path> mismatched) {
        List<String> lines = new ArrayList<>();
        lines.add(screenLabel + " の保存先が外部で変更されています。");
        for (Path p : mismatched) {
            String name = p.getFileName() != null ? p.getFileName().toString() : p.toString();
            byte[] base = SnapshotPresence.get(baselineSnapshots, p);
            byte[] disk = SnapshotPresence.get(diskBytes, p);
            if (SnapshotPresence.isAbsent(base) && SnapshotPresence.isPresent(disk)) {
                lines.add("・" + name + " が新規に作成されています");
            } else if (SnapshotPresence.isPresent(base) && SnapshotPresence.isAbsent(disk)) {
                lines.add("・" + name + " がディスク上にありません");
            } else if (SnapshotPresence.isAbsent(base) && SnapshotPresence.isAbsent(disk)) {
                lines.add("・" + name + " は読込時・保存時ともディスク上にありません");
            } else {
                String hex = FileContentFingerprint.sha256Hex(disk);
                String shortHex = hex.length() <= 12 ? hex : hex.substring(0, 12);
                lines.add("・" + name + " の内容が変更されています（hash=" + shortHex + "…）");
            }
        }
        return String.join("\n", lines);
    }
}
