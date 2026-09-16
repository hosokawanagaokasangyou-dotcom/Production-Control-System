package jp.co.pm.ai.kouchin.trend;

import java.nio.file.Path;

/** 国分②の読み取り。実装は {@link TrendKokubuReader}。 */
public final class TrendReadKokubu {

    private TrendReadKokubu() {}

    public static TrendMonthData read(Path path) {
        return TrendKokubuReader.read(path);
    }
}
