package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;

/**
 * メインウィンドウ {@link javafx.scene.control.TabPane} の入れ子構成とタブ色をセッションに保存するためのツリー。
 *
 * @param kind {@code tab} または {@code group}
 * @param id {@link jp.co.pm.ai.desktop.MainShellTabId#key()}（kind が tab のとき）
 * @param title グループタブの見出し（kind が group のとき）
 * @param colorHex タブ見出しの背景色（任意、{@code #RRGGBB}）
 * @param children グループ内の子（入れ子グループ可）
 */
public record MainShellTabLayoutNode(
        String kind,
        String id,
        String title,
        String colorHex,
        List<MainShellTabLayoutNode> children) {

    public MainShellTabLayoutNode {
        kind = kind != null ? kind.trim() : "";
        id = id != null ? id.trim() : "";
        title = title != null ? title.trim() : "";
        colorHex = colorHex != null ? colorHex.trim() : "";
        children =
                children == null || children.isEmpty()
                        ? List.of()
                        : List.copyOf(children);
    }

    public static MainShellTabLayoutNode tabNode(String tabId, String colorHex) {
        return new MainShellTabLayoutNode("tab", Objects.requireNonNullElse(tabId, ""), "", colorHex, List.of());
    }

    public static MainShellTabLayoutNode groupNode(
            String title, String colorHex, List<MainShellTabLayoutNode> children) {
        return new MainShellTabLayoutNode(
                "group", "", Objects.requireNonNullElse(title, ""), colorHex, children != null ? children : List.of());
    }

    public boolean isGroup() {
        return "group".equalsIgnoreCase(kind);
    }

    public boolean isTab() {
        return "tab".equalsIgnoreCase(kind);
    }

    /**
     * JSON オブジェクト（1 ノード）から復元。不正な場合は {@code null}。
     */
    public static MainShellTabLayoutNode fromJson(JsonNode o) {
        if (o == null || !o.isObject()) {
            return null;
        }
        String k = text(o, "kind");
        if (k.isBlank()) {
            return null;
        }
        String color = text(o, "color");
        if ("group".equalsIgnoreCase(k)) {
            String title = text(o, "title");
            if (title.isBlank()) {
                title = "グループ";
            }
            JsonNode ch = o.get("children");
            List<MainShellTabLayoutNode> list = new ArrayList<>();
            if (ch != null && ch.isArray()) {
                for (JsonNode el : ch) {
                    MainShellTabLayoutNode n = fromJson(el);
                    if (n != null) {
                        list.add(n);
                    }
                }
            }
            return groupNode(title, color, list);
        }
        if ("tab".equalsIgnoreCase(k)) {
            String id = text(o, "id");
            if (id.isBlank()) {
                return null;
            }
            return tabNode(id, color);
        }
        return null;
    }

    private static String text(JsonNode o, String field) {
        JsonNode n = o.get(field);
        if (n == null || n.isNull()) {
            return "";
        }
        if (n.isTextual()) {
            return n.asText("");
        }
        return n.asText("");
    }
}
