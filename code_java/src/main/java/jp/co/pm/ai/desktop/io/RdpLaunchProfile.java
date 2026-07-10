package jp.co.pm.ai.desktop.io;

import java.util.Objects;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * リモートデスクトップ接続時に選択する起動プロファイル。
 *
 * <p>プロファイル番号は {@link RdpRemoteLauncherIni} のスロット番号（99＝接続先サインアウトのみ、
 * 1～{@link RdpRemoteLauncherIni#MAX_SLOTS}＝RPA）と 1:1 対応する。
 *
 * <p>RPA の exe／引数は {@code RAP設定.ini} のスロット行に保存する（RPA 本体の設定は接続先で別途行う前提）。
 * 名称・説明・区分などのメタデータは {@link RdpLaunchProfileCatalog} が JSON で管理する。
 */
public record RdpLaunchProfile(
        int number,
        String name,
        String description,
        String category,
        Boolean disconnectOnChildExit,
        RdpSessionEndAction sessionEndAction,
        Boolean fullScreen,
        Integer desktopWidth,
        Integer desktopHeight,
        Boolean rpaEternal,
        Boolean deleted) {

    public RdpLaunchProfile {
        if (number != RdpRemoteLauncherIni.SLOT_SIGN_OUT
                && (number < 1 || number > RdpRemoteLauncherIni.MAX_SLOTS)) {
            throw new IllegalArgumentException(
                    "プロファイル番号は 99 または 1～"
                            + RdpRemoteLauncherIni.MAX_SLOTS
                            + " です: "
                            + number);
        }
        name = normalizeOptional(name);
        description = normalizeOptional(description);
        category = normalizeOptional(category);
    }

    public static RdpLaunchProfile empty(int number) {
        return new RdpLaunchProfile(number, "", "", "", null, null, null, null, null, null, null);
    }

    /** 起動プロファイル 99（接続先サインアウトのみ）の既定メタデータ。 */
    public static RdpLaunchProfile signOutOnlyDefault() {
        return new RdpLaunchProfile(
                RdpRemoteLauncherIni.SLOT_SIGN_OUT,
                RdpRemoteLauncherIni.SIGN_OUT_ONLY_PROFILE_NAME,
                "接続先のタスクスケジューラが "
                        + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                        + " "
                        + RdpRemoteLauncherIni.SIGN_OUT_LAUNCHER_ARGS
                        + " を起動し、ini の起動プログラム番号=99 によりサインアウトのみ実行します。"
                        + " alternate shell（接続先ランチャー path 埋め込み）は使いません。",
                "サインアウト",
                null,
                RdpSessionEndAction.SIGN_OUT,
                null,
                null,
                null,
                null,
                null);
    }

    public boolean isSignOutOnlyProfile() {
        return number == RdpRemoteLauncherIni.SLOT_SIGN_OUT;
    }

    /** 論理削除済みか（JSON の {@code deleted: true}）。 */
    public boolean isDeleted() {
        return Boolean.TRUE.equals(deleted);
    }

    /** 削除フラグのみ差し替えたコピー。 */
    public RdpLaunchProfile withDeleted(boolean markDeleted) {
        return new RdpLaunchProfile(
                number,
                name,
                description,
                category,
                disconnectOnChildExit,
                sessionEndAction,
                fullScreen,
                desktopWidth,
                desktopHeight,
                rpaEternal,
                markDeleted ? Boolean.TRUE : null);
    }

    /** プロファイルが ini より優先する終了時セッション操作。未設定なら null。 */
    public RdpSessionEndAction resolvedSessionEndAction() {
        if (sessionEndAction != null) {
            return sessionEndAction;
        }
        if (Boolean.FALSE.equals(disconnectOnChildExit)) {
            return RdpSessionEndAction.NONE;
        }
        if (Boolean.TRUE.equals(disconnectOnChildExit)) {
            return RdpSessionEndAction.SIGN_OUT;
        }
        return null;
    }

    /** 接続ボタン横 ComboBox 向けの短い表示。 */
    public String displayLabel() {
        String labelName = name.isBlank() ? "（名称未設定）" : name;
        String base = number + ": " + labelName;
        return isDeleted() ? base + "（削除済）" : base;
    }

    /** 詳細パネル向けの 1 行要約。 */
    public String detailText() {
        StringBuilder sb = new StringBuilder();
        if (!description.isBlank()) {
            sb.append(description.trim());
        }
        if (!category.isBlank()) {
            if (sb.length() > 0) {
                sb.append('\n');
            }
            sb.append("区分: ").append(category.trim());
        }
        if (fullScreen != null || desktopWidth != null || desktopHeight != null) {
            if (sb.length() > 0) {
                sb.append('\n');
            }
            sb.append("表示: ");
            if (Boolean.TRUE.equals(fullScreen)) {
                sb.append("全画面");
            } else if (desktopWidth != null && desktopHeight != null) {
                sb.append(desktopWidth).append('×').append(desktopHeight);
            } else {
                sb.append("既定（接続タブの表示設定）");
            }
        }
        RdpSessionEndAction resolvedEndAction = resolvedSessionEndAction();
        if (resolvedEndAction != null) {
            if (sb.length() > 0) {
                sb.append('\n');
            }
            sb.append("終了時セッション操作: ")
                    .append(resolvedEndAction.displayLabel())
                    .append("（プロファイル設定）");
        }
        if (Boolean.TRUE.equals(rpaEternal)) {
            if (sb.length() > 0) {
                sb.append('\n');
            }
            sb.append("RPA: --eternal（シナリオなし／終了後もプロセス維持）");
        }
        return sb.toString();
    }

    public boolean hasMetadata() {
        return !name.isBlank() || !description.isBlank() || !category.isBlank();
    }

    private static String normalizeOptional(String value) {
        return value != null ? value.trim() : "";
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (!(obj instanceof RdpLaunchProfile other)) {
            return false;
        }
        return number == other.number
                && Objects.equals(name, other.name)
                && Objects.equals(description, other.description)
                && Objects.equals(category, other.category)
                && Objects.equals(disconnectOnChildExit, other.disconnectOnChildExit)
                && Objects.equals(sessionEndAction, other.sessionEndAction)
                && Objects.equals(fullScreen, other.fullScreen)
                && Objects.equals(desktopWidth, other.desktopWidth)
                && Objects.equals(desktopHeight, other.desktopHeight)
                && Objects.equals(rpaEternal, other.rpaEternal)
                && Objects.equals(deleted, other.deleted);
    }

    @Override
    public int hashCode() {
        return Objects.hash(
                number,
                name,
                description,
                category,
                disconnectOnChildExit,
                sessionEndAction,
                fullScreen,
                desktopWidth,
                desktopHeight,
                rpaEternal,
                deleted);
    }
}
