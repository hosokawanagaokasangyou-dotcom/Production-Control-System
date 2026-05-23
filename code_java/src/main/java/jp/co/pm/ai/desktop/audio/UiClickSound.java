package jp.co.pm.ai.desktop.audio;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import javax.sound.sampled.AudioFormat;
import javax.sound.sampled.AudioInputStream;
import javax.sound.sampled.AudioSystem;
import javax.sound.sampled.Clip;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * ボタン押下の短いクリック音。WAV は起動時にメモリへ読み込み、{@link #playClick()} は Clip を再利用してレイテンシを抑える。
 *
 * <p>探索順: {@code code/sounds/ui_button_click.wav} → 組み込み合成クリック。
 */
public final class UiClickSound {

    /** {@link AppPaths#resolveRepoRoot(Map)} からの相対パス（任意）。 */
    public static final String RELATIVE_PATH = "code/sounds/ui_button_click.wav";

    private static final Object LOCK = new Object();
    private static volatile Clip clickClip;

    private UiClickSound() {}

    /** 起動直後にバックグラウンドで Clip を用意する（初回クリックの遅延を避ける）。 */
    public static void warmUpAsync() {
        Thread t =
                new Thread(
                        () -> {
                            try {
                                warmUp(Map.of());
                            } catch (Exception ignored) {
                                // オーディオデバイス無し等
                            }
                        },
                        "ui-click-sound-warmup");
        t.setDaemon(true);
        t.start();
    }

    /** Clip をメモリ上に用意する（未準備なら同期的に読み込む）。 */
    public static void warmUp(Map<String, String> ui) {
        synchronized (LOCK) {
            if (clickClip != null) {
                return;
            }
            try {
                clickClip = openClickClip(ui != null ? ui : Map.of());
            } catch (Exception ignored) {
                clickClip = null;
            }
        }
    }

    /** 事前読込済み Clip を先頭から再生。失敗時は無音。 */
    public static void playClick() {
        Clip clip = clickClip;
        if (clip == null) {
            warmUp(Map.of());
            clip = clickClip;
        }
        if (clip == null) {
            return;
        }
        synchronized (LOCK) {
            try {
                if (clip.isRunning()) {
                    clip.stop();
                }
                clip.setFramePosition(0);
                clip.start();
            } catch (Exception ignored) {
                // 再生中の race 等
            }
        }
    }

    private static Clip openClickClip(Map<String, String> ui) throws Exception {
        Path root = AppPaths.resolveRepoRoot(ui);
        Path wav = root.resolve(RELATIVE_PATH).toAbsolutePath().normalize();
        if (Files.isRegularFile(wav)) {
            try (AudioInputStream ais = AudioSystem.getAudioInputStream(wav.toFile())) {
                Clip clip = AudioSystem.getClip();
                clip.open(ais);
                return clip;
            }
        }
        return openSyntheticClickClip();
    }

    /** 短い「ピコ」系クリック（約 35ms・880Hz・指数減衰）。 */
    private static Clip openSyntheticClickClip() throws Exception {
        byte[] pcm = synthesizeClickPcm();
        AudioFormat format = new AudioFormat(44100f, 16, 1, true, false);
        Clip clip = AudioSystem.getClip();
        clip.open(format, pcm, 0, pcm.length);
        return clip;
    }

    private static byte[] synthesizeClickPcm() {
        int sampleRate = 44100;
        int durationMs = 35;
        int samples = sampleRate * durationMs / 1000;
        byte[] data = new byte[samples * 2];
        for (int i = 0; i < samples; i++) {
            double t = i / (double) sampleRate;
            double env = Math.exp(-t * 140.0);
            double wave = Math.sin(2.0 * Math.PI * 880.0 * t) * env;
            short amp = (short) (wave * 14000.0);
            data[i * 2] = (byte) (amp & 0xff);
            data[i * 2 + 1] = (byte) ((amp >> 8) & 0xff);
        }
        return data;
    }
}
