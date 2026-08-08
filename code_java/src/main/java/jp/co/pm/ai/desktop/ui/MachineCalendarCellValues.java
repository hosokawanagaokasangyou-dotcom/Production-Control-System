package jp.co.pm.ai.desktop.ui;

/** 機械カレンダーセル値の純粋ロジック（JavaFX 非依存）。 */
public final class MachineCalendarCellValues {

    public enum OccupancyMode {
        OCCUPIED("*"),
        AVAILABLE("");

        private final String storedValue;

        OccupancyMode(String storedValue) {
            this.storedValue = storedValue;
        }

        public String storedValue() {
            return storedValue;
        }
    }

    private MachineCalendarCellValues() {}

    public static String toggle(String current) {
        if (current == null || current.isBlank()) {
            return "*";
        }
        return "";
    }

    public static String shortLabel(String val) {
        if (val == null || val.isBlank()) {
            return "·";
        }
        return val.length() > 3 ? val.substring(0, 3) : val;
    }

    public static OccupancyMode resolvePaintModeFromAnchor(String anchorValue) {
        if (anchorValue == null || anchorValue.isBlank()) {
            return OccupancyMode.OCCUPIED;
        }
        return OccupancyMode.AVAILABLE;
    }

    public static boolean isOccupied(String value) {
        return value != null && !value.isBlank();
    }
}
