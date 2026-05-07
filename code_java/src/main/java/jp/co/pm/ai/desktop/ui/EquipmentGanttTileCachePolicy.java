package jp.co.pm.ai.desktop.ui;

/**
 * Policy constants for a future day-tile image cache (plan case 3). No runtime cache is wired yet.
 */
public final class EquipmentGanttTileCachePolicy {

    public static final int MAX_TILES_LRU = 64;

    public static final long MAX_TILE_RGBA_BYTES = 32L * 1024 * 1024;

    public static final boolean INVALIDATE_ON_STYLE_CHANGE = true;

    private EquipmentGanttTileCachePolicy() {}
}
