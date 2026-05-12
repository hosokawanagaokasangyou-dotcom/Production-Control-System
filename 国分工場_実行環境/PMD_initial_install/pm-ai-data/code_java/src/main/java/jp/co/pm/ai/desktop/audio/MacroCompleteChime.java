package jp.co.pm.ai.desktop.audio;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import javax.sound.sampled.AudioSystem;
import javax.sound.sampled.Clip;
import javax.sound.sampled.LineEvent;
import javax.sound.sampled.LineListener;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * Short completion chime for successful button-started runs (stage 1/2). Audio file:
 * {@code code/sounds/macro_complete_chime.wav} under the repo root.
 */
public final class MacroCompleteChime {

    /** Path relative to {@link AppPaths#resolveRepoRoot(Map)}. */
    public static final String RELATIVE_PATH = "code/sounds/macro_complete_chime.wav";

    private MacroCompleteChime() {}

    /**
     * Plays the WAV if it exists under the resolved repo root. Failures are ignored so logs/dialogs are
     * unaffected.
     */
    public static void playIfAvailable(Map<String, String> ui) {
        try {
            Path root = AppPaths.resolveRepoRoot(ui != null ? ui : Map.of());
            Path wav = root.resolve(RELATIVE_PATH).toAbsolutePath().normalize();
            if (!Files.isRegularFile(wav)) {
                return;
            }
            try (var ais = AudioSystem.getAudioInputStream(wav.toFile())) {
                Clip clip = AudioSystem.getClip();
                clip.open(ais);
                clip.addLineListener(
                        (LineListener)
                                event -> {
                                    if (event.getType() == LineEvent.Type.STOP) {
                                        clip.close();
                                    }
                                });
                clip.start();
            }
        } catch (Exception ignored) {
            // Missing audio device or unsupported format: ignore
        }
    }
}
