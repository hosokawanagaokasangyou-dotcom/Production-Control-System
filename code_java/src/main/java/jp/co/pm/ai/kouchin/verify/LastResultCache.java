package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 工場別の直近 {@link MailSnapshot} を JSON で保存・読み込みする。
 * 保存先は {@code ~/.pm-ai-desktop}（{@code ~/.kouchin-verify-desktop} は使わない）。
 */
public final class LastResultCache {

    private static final ObjectMapper MAPPER =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private LastResultCache() {}

    public static Path pathFor(FactoryId factory) {
        return factory == FactoryId.KOKUBU
                ? AppPaths.resolveKouchinLastResultKokubuPath()
                : AppPaths.resolveKouchinLastResultKonanPath();
    }

    public static Optional<MailSnapshot> load(FactoryId factory) {
        Path path = pathFor(factory);
        if (!Files.isRegularFile(path)) {
            return Optional.empty();
        }
        try {
            return Optional.ofNullable(MAPPER.readValue(path.toFile(), MailSnapshot.class));
        } catch (IOException | RuntimeException e) {
            return Optional.empty();
        }
    }

    public static void save(FactoryId factory, MailSnapshot snapshot) throws IOException {
        Files.createDirectories(AppPaths.resolveDesktopAppHomeDir());
        MAPPER.writeValue(pathFor(factory).toFile(), snapshot);
    }

    public static boolean saveQuietly(FactoryId factory, MailSnapshot snapshot) {
        try {
            save(factory, snapshot);
            return true;
        } catch (IOException | RuntimeException e) {
            return false;
        }
    }
}
