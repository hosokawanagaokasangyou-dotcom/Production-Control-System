package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.util.Collection;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.function.Consumer;

/**
 * 依頼書原本 Excel の更新日時を追跡し、プレビュー確認前の「更新あり」状態を保持する。
 */
public final class RequestFormOriginalUpdateMonitor {

    private final ConcurrentMap<String, Long> acknowledgedMtimeByKey = new ConcurrentHashMap<>();
    private final Set<String> updatedKeys = ConcurrentHashMap.newKeySet();
    private volatile Consumer<Collection<String>> onUpdatedKeysChanged = __ -> {};

    public void setOnUpdatedKeysChanged(Consumer<Collection<String>> listener) {
        this.onUpdatedKeysChanged = listener != null ? listener : __ -> {};
    }

    static String canonicalKey(File file) {
        if (file == null) {
            return "";
        }
        try {
            return file.getCanonicalPath();
        } catch (IOException e) {
            return file.getAbsolutePath();
        }
    }

    /** 初回は現在の mtime を基準に登録（バッジなし）。 */
    public void ensureTracked(File file) {
        if (file == null || !file.isFile()) {
            return;
        }
        String key = canonicalKey(file);
        if (key.isEmpty()) {
            return;
        }
        long mtime = file.lastModified();
        acknowledgedMtimeByKey.putIfAbsent(key, mtime);
    }

    /** 更新日時を確認し、確認済みより新しければ「更新あり」にする。 */
    public boolean poll(File file) {
        if (file == null || !file.isFile()) {
            return false;
        }
        String key = canonicalKey(file);
        if (key.isEmpty()) {
            return false;
        }
        long mtime = file.lastModified();
        Long ack = acknowledgedMtimeByKey.get(key);
        if (ack == null) {
            acknowledgedMtimeByKey.put(key, mtime);
            return false;
        }
        if (mtime > ack) {
            boolean added = updatedKeys.add(key);
            if (added) {
                onUpdatedKeysChanged.accept(Set.copyOf(updatedKeys));
            }
            return true;
        }
        return updatedKeys.contains(key);
    }

    public void pollAll(Collection<File> files) {
        if (files == null || files.isEmpty()) {
            return;
        }
        for (File file : files) {
            poll(file);
        }
    }

    /** プレビュー表示後: 現在の mtime を確認済みとし、バッジを消す。 */
    public void markPreviewAcknowledged(File file) {
        if (file == null || !file.isFile()) {
            return;
        }
        String key = canonicalKey(file);
        if (key.isEmpty()) {
            return;
        }
        long mtime = file.lastModified();
        acknowledgedMtimeByKey.put(key, mtime);
        if (updatedKeys.remove(key)) {
            onUpdatedKeysChanged.accept(Set.copyOf(updatedKeys));
        }
    }

    public boolean isUpdated(File file) {
        if (file == null) {
            return false;
        }
        String key = canonicalKey(file);
        return !key.isEmpty() && updatedKeys.contains(key);
    }

    public Set<String> updatedKeysSnapshot() {
        return Set.copyOf(updatedKeys);
    }
}
