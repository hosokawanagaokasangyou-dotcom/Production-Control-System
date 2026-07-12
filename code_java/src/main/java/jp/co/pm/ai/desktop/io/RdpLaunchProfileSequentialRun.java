package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/**
 * リモートデスクトップタブの連続実行（複数起動プロファイルを順番に接続）向けロジック。
 */
public final class RdpLaunchProfileSequentialRun {

    private RdpLaunchProfileSequentialRun() {}

    /**
     * プロファイル番号を選択順でトグルする。
     * プロファイル {@link RdpRemoteLauncherIni#SLOT_SIGN_OUT} は先頭（未選択時の追加）のみ可。
     *
     * @return 変更後の選択順リスト（不変コピー）。追加不可のときは {@code currentOrder} のコピー。
     */
    public static List<Integer> toggleSelection(List<Integer> currentOrder, int profileNumber) {
        Objects.requireNonNull(currentOrder, "currentOrder");
        if (profileNumber <= 0) {
            return List.copyOf(currentOrder);
        }
        List<Integer> next = new ArrayList<>(currentOrder);
        if (next.remove(Integer.valueOf(profileNumber))) {
            return List.copyOf(next);
        }
        if (!canAddProfileToSelection(currentOrder, profileNumber)) {
            return List.copyOf(currentOrder);
        }
        next.add(profileNumber);
        return List.copyOf(next);
    }

    /**
     * 連続実行の選択にプロファイルを追加できるか。
     * プロファイル 99 は選択が空のときのみ追加可（先頭限定）。
     */
    public static boolean canAddProfileToSelection(List<Integer> currentOrder, int profileNumber) {
        Objects.requireNonNull(currentOrder, "currentOrder");
        if (!RdpRemoteLauncherIni.isSignOutOnlyProfile(profileNumber)) {
            return true;
        }
        return currentOrder.isEmpty();
    }

    /** プロファイル 99 を先頭以外に置けないことを検証する。問題なければ empty。 */
    public static java.util.Optional<String> validateSignOutOnlyAtHead(List<Integer> selectionOrder) {
        List<Integer> normalized = normalizeSelection(selectionOrder);
        int signOutIdx = normalized.indexOf(RdpRemoteLauncherIni.SLOT_SIGN_OUT);
        if (signOutIdx < 0) {
            return java.util.Optional.empty();
        }
        if (signOutIdx != 0) {
            return java.util.Optional.of(
                    "起動プロファイル "
                            + RdpRemoteLauncherIni.SLOT_SIGN_OUT
                            + "（接続先サインアウトのみ）は連続実行の先頭のみ選択できます。");
        }
        return java.util.Optional.empty();
    }

    /** 連続実行キューに RPA プロファイル（1〜9）が含まれるか。 */
    public static boolean selectionRequiresAladdinCredentials(Iterable<Integer> selectionOrder) {
        for (Integer n : normalizeSelection(selectionOrder)) {
            if (!RdpRemoteLauncherIni.isSignOutOnlyProfile(n)) {
                return true;
            }
        }
        return false;
    }

    /** 選択順における 1 始まりの表示番号。未選択なら empty。 */
    public static int selectionOrderIndex(List<Integer> selectionOrder, int profileNumber) {
        Objects.requireNonNull(selectionOrder, "selectionOrder");
        int idx = selectionOrder.indexOf(profileNumber);
        return idx < 0 ? -1 : idx + 1;
    }

    /** クイック起動ボタン表示用。選択順があるときは先頭に付与する。 */
    public static String quickButtonLabel(String baseLabel, int selectionOrderIndex) {
        if (baseLabel == null) {
            baseLabel = "";
        }
        if (selectionOrderIndex <= 0) {
            return baseLabel;
        }
        return selectionOrderMarker(selectionOrderIndex) + " " + baseLabel;
    }

    /** 1→①, 2→② … 10 以降は数字。 */
    public static String selectionOrderMarker(int oneBasedIndex) {
        if (oneBasedIndex <= 0) {
            return "";
        }
        return switch (oneBasedIndex) {
            case 1 -> "①";
            case 2 -> "②";
            case 3 -> "③";
            case 4 -> "④";
            case 5 -> "⑤";
            case 6 -> "⑥";
            case 7 -> "⑦";
            case 8 -> "⑧";
            case 9 -> "⑨";
            default -> oneBasedIndex + ".";
        };
    }

    public static String launchButtonTextIdle(int selectedCount) {
        if (selectedCount <= 0) {
            return "連続実行するタスクを選択";
        }
        return "連続実行を開始（" + selectedCount + "件）";
    }

    public static String launchButtonTextActive(int currentOneBased, int total) {
        if (total <= 0 || currentOneBased <= 0) {
            return "連続実行 接続中";
        }
        return "連続実行 " + currentOneBased + "/" + total + " 接続中";
    }

    public static String progressStatusText(int currentOneBased, int total, String profileLabel) {
        if (total <= 0 || currentOneBased <= 0) {
            return "";
        }
        String label = profileLabel != null ? profileLabel.trim() : "";
        if (label.isEmpty()) {
            return "連続実行 " + currentOneBased + "/" + total;
        }
        return "連続実行 " + currentOneBased + "/" + total + ": " + label;
    }

    /** 重複を除きつつ選択順を保持する。 */
    public static List<Integer> normalizeSelection(Iterable<Integer> source) {
        if (source == null) {
            return List.of();
        }
        Set<Integer> seen = new LinkedHashSet<>();
        for (Integer n : source) {
            if (n != null && n > 0) {
                seen.add(n);
            }
        }
        return List.copyOf(seen);
    }
}
