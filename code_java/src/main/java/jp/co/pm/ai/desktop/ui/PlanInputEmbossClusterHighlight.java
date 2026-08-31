package jp.co.pm.ai.desktop.ui;

/**
 * 配台計画タブ: 「エンボスをまとめる」と「保存」のオレンジ点灯。
 *
 * <p>エンボスがある読込直後はまとめボタンを点灯。押下後はまとめを消し、保存を点灯。
 * 保存成功後は保存も消す。再読込までまとめは再点灯しない。
 */
public final class PlanInputEmbossClusterHighlight {

    private boolean clusteredSinceLoad;
    private boolean savePendingAfterCluster;

    public void resetForLoadedTable() {
        clusteredSinceLoad = false;
        savePendingAfterCluster = false;
    }

    public void markClusterApplied() {
        clusteredSinceLoad = true;
        savePendingAfterCluster = true;
    }

    public void markSaved() {
        savePendingAfterCluster = false;
    }

    public boolean clusterHot(boolean hasEligibleEmboss) {
        return hasEligibleEmboss && !clusteredSinceLoad;
    }

    public boolean saveHot() {
        return savePendingAfterCluster;
    }
}
