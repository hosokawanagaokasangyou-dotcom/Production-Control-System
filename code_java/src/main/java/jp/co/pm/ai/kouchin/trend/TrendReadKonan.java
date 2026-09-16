package jp.co.pm.ai.kouchin.trend;

import java.nio.file.Path;

/** 湖南②の読み取り。実装は {@link TrendKonanReader}。 */
public final class TrendReadKonan {

    private TrendReadKonan() {}

    public static TrendMonthData read(Path path) {
        return TrendKonanReader.read(path);
    }
}
