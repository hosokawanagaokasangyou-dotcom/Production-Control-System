package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.stage.Window;

import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;

/** 段階2直前: 加工途中とアラジン当日対象を同じ「翌日配台量」で一括入力する。 */
public final class Stage2NextDayDispatchDialog {

    static final Stage2NextDayRollDispatchDialogSupport.Theme THEME =
            new Stage2NextDayRollDispatchDialogSupport.Theme(
                    "段階2 — 翌日の配台量",
                    "対象行について、翌日に配台するロール数を指定してください。"
                            + " 0 の行は翌日に配台しません。",
                    "ロール数はコンボボックスから選びます。値は翌日の配台上限です。"
                            + " 配台ロール単位 (m) の整数倍です。"
                            + " 設備能力などにより実際の配台量は入力値より少なくなる場合があります。",
                    "実加工",
                    "翌日配台(ロール)",
                    "",
                    "-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);",
                    true,
                    true);

    private Stage2NextDayDispatchDialog() {}

    public record Result(
            List<Stage2InProgressNextDayDispatchIo.Entry> inProgressEntries,
            List<Stage2AladdinTodayExcludeNextDayDispatchIo.Entry> aladdinExcludeEntries) {
        public Result {
            inProgressEntries = List.copyOf(inProgressEntries);
            aladdinExcludeEntries = List.copyOf(aladdinExcludeEntries);
        }
    }

    private record ConvertedEntry(
            Stage2InProgressNextDayDispatchIo.Entry inProgress,
            Stage2AladdinTodayExcludeNextDayDispatchIo.Entry aladdinExclude) {}

    /** @return 確定時は両種別の保存値。キャンセル時は empty。 */
    public static Optional<Result> prompt(
            Window owner,
            List<Stage2InProgressNextDayDispatchDialog.Row> inProgressRows,
            List<Stage2AladdinTodayExcludeNextDayDispatchDialog.Row> aladdinRows) {
        List<Stage2NextDayRollDispatchDialogSupport.RowModel> rows = new ArrayList<>();
        rows.addAll(safeInProgressRows(inProgressRows));
        rows.addAll(safeAladdinRows(aladdinRows));
        if (rows.isEmpty()) {
            return Optional.of(new Result(List.of(), List.of()));
        }

        Optional<List<ConvertedEntry>> converted =
                Stage2NextDayRollDispatchDialogSupport.prompt(
                        owner,
                        rows,
                        THEME,
                        Stage2NextDayDispatchDialog::convertRow,
                        r ->
                                Stage2InProgressNextDayRollInput.validateRollInput(
                                        r.rollCountProperty().get(),
                                        r.remainingM(),
                                        r.unitInfo()));
        return converted.map(Stage2NextDayDispatchDialog::toResult);
    }

    static Result collectResult(
            List<Stage2InProgressNextDayDispatchDialog.Row> inProgressRows,
            List<Stage2AladdinTodayExcludeNextDayDispatchDialog.Row> aladdinRows) {
        List<ConvertedEntry> converted = new ArrayList<>();
        safeInProgressRows(inProgressRows).forEach(row -> converted.add(convertRow(row)));
        safeAladdinRows(aladdinRows).forEach(row -> converted.add(convertRow(row)));
        return toResult(converted);
    }

    private static ConvertedEntry convertRow(
            Stage2NextDayRollDispatchDialogSupport.RowModel row) {
        if (row instanceof Stage2InProgressNextDayDispatchDialog.Row inProgress) {
            return new ConvertedEntry(inProgress.toEntryFromNextDayInput(), null);
        }
        if (row instanceof Stage2AladdinTodayExcludeNextDayDispatchDialog.Row aladdin) {
            return new ConvertedEntry(null, aladdin.toEntryFromNextDayInput());
        }
        throw new IllegalArgumentException("未対応の段階2翌日配台行です: " + row.getClass().getName());
    }

    private static Result toResult(List<ConvertedEntry> converted) {
        List<Stage2InProgressNextDayDispatchIo.Entry> inProgress = new ArrayList<>();
        List<Stage2AladdinTodayExcludeNextDayDispatchIo.Entry> aladdin = new ArrayList<>();
        for (ConvertedEntry entry : converted) {
            if (entry.inProgress() != null) {
                inProgress.add(entry.inProgress());
            }
            if (entry.aladdinExclude() != null) {
                aladdin.add(entry.aladdinExclude());
            }
        }
        return new Result(inProgress, aladdin);
    }

    private static List<Stage2InProgressNextDayDispatchDialog.Row> safeInProgressRows(
            List<Stage2InProgressNextDayDispatchDialog.Row> rows) {
        return rows != null ? rows : List.of();
    }

    private static List<Stage2AladdinTodayExcludeNextDayDispatchDialog.Row> safeAladdinRows(
            List<Stage2AladdinTodayExcludeNextDayDispatchDialog.Row> rows) {
        return rows != null ? rows : List.of();
    }
}
