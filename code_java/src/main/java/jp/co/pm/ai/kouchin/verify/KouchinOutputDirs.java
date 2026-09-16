package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 結果 Excel / メール / 月トレンドの出力先。現工場に関係なく国分と湖南の両方へ書く。
 */
public final class KouchinOutputDirs {

    private KouchinOutputDirs() {}

    public static Path kokubuDir() {
        return Paths.get(AppPaths.DEFAULT_KOUCHIN_BASE_DIR);
    }

    public static Path konanDir() {
        return Paths.get(AppPaths.DEFAULT_KOUCHIN_KONAN_OUTPUT_DIR);
    }

    /**
     * 固定2先（国分 ●自動検証、湖南 共有DATA\●自動検証）。{@code PM_AI_KOUCHIN_OUTPUT_DIR} が非空なら第3先を追加。
     */
    public static List<Path> resolveAll(Map<String, String> ui) {
        LinkedHashSet<Path> dirs = new LinkedHashSet<>();
        dirs.add(kokubuDir().toAbsolutePath().normalize());
        dirs.add(konanDir().toAbsolutePath().normalize());
        String extra =
                ui == null ? "" : ui.getOrDefault(AppPaths.KEY_PM_AI_KOUCHIN_OUTPUT_DIR, "");
        if (extra != null && !extra.isBlank()) {
            dirs.add(Paths.get(extra.trim()).toAbsolutePath().normalize());
        }
        return List.copyOf(dirs);
    }

    public static List<Path> resolveAll() {
        return resolveAll(Map.of());
    }

    public static List<Path> withNames(List<Path> dirs, String fileName) {
        List<Path> out = new ArrayList<>();
        for (Path dir : dirs) {
            out.add(dir.resolve(fileName));
        }
        return List.copyOf(out);
    }
}
