package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;

import javafx.scene.image.Image;
import javafx.stage.Stage;

/** アプリ／スプラッシュのウィンドウアイコン（タスクバー・タイトルバー）。 */
public final class AppWindowIconSupport {

    public enum Variant {
        DESKTOP("app-icon"),
        RDP_LAUNCHER("rdp-launcher-icon");

        private final String resourceBase;

        Variant(String resourceBase) {
            this.resourceBase = resourceBase;
        }

        String resourceBase() {
            return resourceBase;
        }
    }

    private static final int[] ICON_SIZES = {16, 32, 48, 64, 128, 256};

    private AppWindowIconSupport() {}

    public static void applyTo(Stage stage, Variant variant) {
        if (stage == null || variant == null) {
            return;
        }
        List<Image> icons = loadIcons(variant);
        if (!icons.isEmpty()) {
            stage.getIcons().setAll(icons);
        }
    }

    static List<Image> loadIcons(Variant variant) {
        List<Image> icons = new ArrayList<>();
        String base = "/jp/co/pm/ai/desktop/images/" + variant.resourceBase();
        for (int size : ICON_SIZES) {
            String path = base + "-" + size + ".png";
            var url = AppWindowIconSupport.class.getResource(path);
            if (url != null) {
                icons.add(new Image(url.toExternalForm(), size, size, true, true));
            }
        }
        return icons;
    }
}
