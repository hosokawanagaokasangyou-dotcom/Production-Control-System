package jp.co.pm.ai.planning.stage2.core;

import java.time.LocalTime;
import java.util.List;
import java.util.Optional;

import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * マスタ・計画入力から得られる制約の要約（Python 側の need／勤怠／カレンダー解釈と突き合わせる前段のダイジェスト）。
 */
public record Stage2ConstraintDigest(
        int memberCount,
        String factoryStart,
        String factoryEnd,
        int excludeRuleCount,
        int planDataRows,
        String planSheetResolved,
        String masterPath) {

    public static Stage2ConstraintDigest fromSnapshot(Stage2InputSnapshot snap) {
        List<String> members = snap.memberDisplayNames();
        Optional<LocalTime> fs = snap.factoryStart();
        Optional<LocalTime> fe = snap.factoryEnd();
        int rows = snap.planningTasksSheet().rows() != null ? snap.planningTasksSheet().rows().size() : 0;
        return new Stage2ConstraintDigest(
                members != null ? members.size() : 0,
                fs.map(LocalTime::toString).orElse(""),
                fe.map(LocalTime::toString).orElse(""),
                snap.excludeRuleCount(),
                rows,
                snap.planSheetName(),
                snap.masterPath().toString());
    }
}
