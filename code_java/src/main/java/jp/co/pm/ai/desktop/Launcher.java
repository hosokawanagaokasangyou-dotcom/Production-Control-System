package jp.co.pm.ai.desktop;

/**
 * {@code javafx:run} 用エントリ（{@code org.openjfx:javafx-maven-plugin} の {@code runtimePathOption=CLASSPATH} 時、
 * メインクラスは JavaFX の {@code javafx.application.Application} を継承してはならない制約のため）。
 *
 * <p>パッケージ済み起動・{@code exec:exec@pm-ai-desktop} は引き続き {@link PmAiFxApp} を直接指定する。
 */
public final class Launcher {

    private Launcher() {}

    public static void main(String[] args) {
        PmAiFxApp.main(args);
    }
}
