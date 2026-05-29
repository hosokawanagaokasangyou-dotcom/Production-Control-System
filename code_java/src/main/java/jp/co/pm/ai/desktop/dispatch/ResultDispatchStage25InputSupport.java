package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import jp.co.pm.ai.desktop.config.AppPaths;

/** 段階2.5 前景実行前の入力 JSON 同期（整列前退避）。 */
public final class ResultDispatchStage25InputSupport {

    private ResultDispatchStage25InputSupport() {}

    /**
     * 整列前の {@code 結果_配台表.json} を一時ファイルへコピーする。
     *
     * @return 退避ファイルの絶対パス
     */
    public static Path copyDispatchJsonBeforeAlign(Path dispatchJson) throws java.io.IOException {
        if (dispatchJson == null || !Files.isRegularFile(dispatchJson)) {
            throw new java.io.FileNotFoundException("結果_配台表.json が見つかりません: " + dispatchJson);
        }
        Path parent = dispatchJson.getParent();
        if (parent == null) {
            parent = AppPaths.resolveResultDispatchTableDir(java.util.Map.of());
        }
        String name = dispatchJson.getFileName().toString();
        Path raw = parent.resolve(name + ".stage2_raw." + java.util.UUID.randomUUID() + ".json");
        Files.copy(dispatchJson, raw, StandardCopyOption.COPY_ATTRIBUTES);
        return raw.toAbsolutePath().normalize();
    }
}
