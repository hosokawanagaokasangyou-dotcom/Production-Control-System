package jp.co.pm.ai.desktop.config;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import javafx.scene.image.Image;
import javafx.scene.image.ImageView;

/**
 * 工場別ロゴ画像（バンドル PNG およびリポジトリ {@code code/branding/} の上書き）。
 */
public final class FactorySiteLogoSupport {

    private static final String BUNDLED_PREFIX = "/jp/co/pm/ai/desktop/images/factory-";

    private FactorySiteLogoSupport() {}

    /**
     * リポジトリ {@code code/branding/factory-&lt;site&gt;.png} を優先し、無ければクラスパス上のバンドル PNG を返す。
     */
    public static Optional<Image> resolveImage(FactorySite site, Map<String, String> ui) {
        if (site == null) {
            return Optional.empty();
        }
        Optional<Image> override = loadFromRepoBranding(site, ui);
        if (override.isPresent()) {
            return override;
        }
        return loadBundled(site);
    }

    /**
     * {@code code/branding/} の上書き画像のみ ImageView に載せる。バンドル PNG は使わず CSS の立体矩形＋中央ラベルとする。
     */
    public static void applyBrandingOverrideToImageView(
            ImageView view, FactorySite site, Map<String, String> ui) {
        if (view == null) {
            return;
        }
        Optional<Image> img = loadFromRepoBranding(site, ui);
        if (img.isPresent()) {
            view.setImage(img.get());
            view.setVisible(true);
            view.setManaged(true);
        } else {
            view.setImage(null);
            view.setVisible(false);
            view.setManaged(false);
        }
    }

    static String bundledResourcePath(FactorySite site) {
        return BUNDLED_PREFIX + site.name().toLowerCase() + ".png";
    }

    private static Optional<Image> loadFromRepoBranding(FactorySite site, Map<String, String> ui) {
        try {
            Path p =
                    AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                            .resolve("code")
                            .resolve("branding")
                            .resolve("factory-" + site.name().toLowerCase() + ".png")
                            .normalize();
            if (!Files.isRegularFile(p)) {
                return Optional.empty();
            }
            return Optional.of(new Image(p.toUri().toString(), true));
        } catch (Exception ignored) {
            return Optional.empty();
        }
    }

    private static Optional<Image> loadBundled(FactorySite site) {
        String path = bundledResourcePath(site);
        try (InputStream in = FactorySiteLogoSupport.class.getResourceAsStream(path)) {
            if (in == null) {
                return Optional.empty();
            }
            return Optional.of(new Image(in));
        } catch (Exception ignored) {
            return Optional.empty();
        }
    }
}
