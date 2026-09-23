package jp.co.pm.ai.desktop.dispatch;

import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.function.BooleanSupplier;

/**
 * 画面が今見ている配台結果。既定は自分のローカル最新。セッションには保存しない。
 * リスナーの例外は飲み込んで、他タブとアプリ本体を止めない。
 */
public final class DispatchResultSelection {

    public enum Kind {
        LOCAL_LATEST,
        SNAPSHOT
    }

    private final CopyOnWriteArrayList<Runnable> listeners = new CopyOnWriteArrayList<>();
    private volatile BooleanSupplier changeAllowed = () -> true;

    private Kind kind = Kind.LOCAL_LATEST;
    private String operatorDir = "";
    private String generationDir = "";
    private String badgeText = "";
    private String detailText = "自分のローカル最新";
    private boolean ownPast;

    public Kind kind() {
        return kind;
    }

    public boolean isLocalLatest() {
        return kind == Kind.LOCAL_LATEST;
    }

    public String operatorDir() {
        return operatorDir;
    }

    public String generationDir() {
        return generationDir;
    }

    public String badgeText() {
        return badgeText == null ? "" : badgeText;
    }

    public String detailText() {
        return detailText == null ? "" : detailText;
    }

    public boolean ownPast() {
        return ownPast;
    }

    /** 自分の最新のときは空。タブ内バッジと印刷に使う。 */
    public static String badgeTextFor(boolean ownPast, String operatorLabel, String savedAt) {
        String when = savedAt == null ? "" : savedAt.strip();
        if (ownPast) {
            return when.isEmpty() ? "過去の結果" : "過去の結果: " + when;
        }
        String who = operatorLabel == null || operatorLabel.isBlank() ? "他者" : operatorLabel.strip();
        return when.isEmpty() ? "他者の結果: " + who : "他者の結果: " + who + " " + when;
    }

    /** メインタブ見出しに足す短い印。自分の最新のときは空。 */
    public static String tabMark(boolean localLatest, boolean ownPast) {
        if (localLatest) {
            return "";
        }
        return ownPast ? " ［過去の結果］" : " ［他者の結果］";
    }

    public void setChangeAllowed(BooleanSupplier changeAllowed) {
        this.changeAllowed = changeAllowed != null ? changeAllowed : () -> true;
    }

    public void addListener(Runnable listener) {
        if (listener != null) {
            listeners.add(listener);
        }
    }

    public boolean selectLocal() {
        return change(Kind.LOCAL_LATEST, "", "", "", "自分のローカル最新", false);
    }

    public boolean selectSnapshot(
            String operatorDir, String generationDir, String badgeText, String detailText, boolean ownPast) {
        if (operatorDir == null || operatorDir.isBlank() || generationDir == null || generationDir.isBlank()) {
            return false;
        }
        if (!DispatchSnapshotStore.isGenerationDirName(generationDir)) {
            return false;
        }
        return change(
                Kind.SNAPSHOT,
                operatorDir,
                generationDir,
                badgeText == null ? "" : badgeText,
                detailText == null ? "" : detailText,
                ownPast);
    }

    private boolean change(
            Kind nextKind,
            String nextOperator,
            String nextGeneration,
            String nextBadge,
            String nextDetail,
            boolean nextOwnPast) {
        if (kind == nextKind
                && operatorDir.equals(nextOperator)
                && generationDir.equals(nextGeneration)) {
            return true;
        }
        try {
            BooleanSupplier gate = changeAllowed;
            if (gate != null && !gate.getAsBoolean()) {
                return false;
            }
        } catch (RuntimeException ex) {
            return false;
        }
        kind = nextKind;
        operatorDir = nextOperator;
        generationDir = nextGeneration;
        badgeText = nextBadge;
        detailText = nextDetail;
        ownPast = nextOwnPast;
        for (Runnable listener : List.copyOf(listeners)) {
            try {
                listener.run();
            } catch (RuntimeException ex) {
                // 1タブの失敗で残りのタブやシェルを止めない
            }
        }
        return true;
    }
}
