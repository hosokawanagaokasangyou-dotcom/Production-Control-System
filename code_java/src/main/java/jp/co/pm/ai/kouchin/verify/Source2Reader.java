package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Path;

/**
 * 工場プロファイルに応じて②を読み分ける。
 * 国分 = {@link NagaokaReader}（東レまとめ）、湖南 = {@link ShisanReader}（東レ3シート）。
 */
public final class Source2Reader {

    private Source2Reader() {
    }

    public static MoneyMaps read(FactoryProfile profile, Path path) {
        return profile.id() == FactoryId.KOKUBU ? NagaokaReader.read(path) : ShisanReader.read(path);
    }
}
