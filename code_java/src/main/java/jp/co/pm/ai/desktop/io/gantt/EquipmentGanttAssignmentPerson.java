package jp.co.pm.ai.desktop.io.gantt;

import java.text.Normalizer;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;

/** 担当割当編集用の1名分（フル氏名・役割・バッジ表示・照合キー）。 */
public record EquipmentGanttAssignmentPerson(
        String fullName,
        EquipmentGanttAssignmentRole role,
        String badgeLabel,
        String memberKey) {

    public EquipmentGanttAssignmentPerson {
        fullName = fullName != null ? fullName : "";
        badgeLabel = badgeLabel != null ? badgeLabel : "";
        memberKey = memberKey != null ? memberKey : "";
    }

    public EquipmentGanttAssignmentPerson withRole(EquipmentGanttAssignmentRole newRole) {
        return new EquipmentGanttAssignmentPerson(fullName, newRole, badgeLabel, memberKey);
    }

    public static EquipmentGanttAssignmentPerson fromRawName(
            String raw, EquipmentGanttAssignmentRole role) {
        String full = normalizeFullName(raw);
        String badge = PersonNameBadgeText.badgeTwoFromRawName(raw);
        String key = PersonBadgeStyle.normalizeLabelKey(full);
        return new EquipmentGanttAssignmentPerson(full, role, badge, key);
    }

    static String normalizeFullName(String raw) {
        if (raw == null) {
            return "";
        }
        return Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC);
    }
}
