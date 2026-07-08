package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

/**
 * 転記照合・原反投入日照合で共用する日付の等価判定。
 *
 * <p>複数行（改行区切り）・{@code M/D} 形式・年省略・和暦表記などを {@link JuchuTransferValueNormalizer}
 * で解決したうえで比較する。
 */
final class JuchuTransferDateMatcher {

    private JuchuTransferDateMatcher() {}

    /** 2 つの日付文字列（複数行可）が同一の日付集合を表すかを判定する。 */
    static boolean datesMatch(String left, String right) {
        List<String> leftLines = splitDateLines(left);
        List<String> rightLines = splitDateLines(right);
        if (leftLines.isEmpty() && rightLines.isEmpty()) {
            return true;
        }
        if (leftLines.isEmpty() || rightLines.isEmpty()) {
            return false;
        }
        if (leftLines.size() == rightLines.size()) {
            for (int i = 0; i < leftLines.size(); i++) {
                if (!singleDateMatch(leftLines.get(i), rightLines.get(i))) {
                    return false;
                }
            }
            return true;
        }
        // 原反複数行で片側のみ M/D が複数行・もう一方が1行（同一日）など
        if (rightLines.size() == 1) {
            String rightDate = rightLines.get(0);
            return leftLines.stream().allMatch(o -> singleDateMatch(o, rightDate));
        }
        if (leftLines.size() == 1) {
            String leftDate = leftLines.get(0);
            return rightLines.stream().allMatch(j -> singleDateMatch(leftDate, j));
        }
        return datesMatchSameUniqueResolvedDate(leftLines, rightLines);
    }

    /** 行数は異なるが、解決後の日付がいずれも同一なら一致。 */
    private static boolean datesMatchSameUniqueResolvedDate(
            List<String> leftLines, List<String> rightLines) {
        LocalDate rightRef =
                rightLines.stream()
                        .map(JuchuTransferValueNormalizer::parseLocalDate)
                        .filter(d -> d != null)
                        .findFirst()
                        .orElse(LocalDate.now());
        LocalDate leftRef =
                leftLines.stream()
                        .map(JuchuTransferValueNormalizer::parseLocalDate)
                        .filter(d -> d != null)
                        .findFirst()
                        .orElse(rightRef);
        List<LocalDate> leftResolved = resolveDateLines(leftLines, rightRef);
        List<LocalDate> rightResolved = resolveDateLines(rightLines, leftRef);
        if (leftResolved.contains(null) || rightResolved.contains(null)) {
            return false;
        }
        if (leftResolved.stream().distinct().count() != 1
                || rightResolved.stream().distinct().count() != 1) {
            return false;
        }
        return leftResolved.get(0).equals(rightResolved.get(0));
    }

    private static List<LocalDate> resolveDateLines(List<String> lines, LocalDate yearReference) {
        List<LocalDate> resolved = new ArrayList<>();
        for (String line : lines) {
            resolved.add(JuchuTransferValueNormalizer.parseLocalDate(line, yearReference));
        }
        return resolved;
    }

    private static List<String> splitDateLines(String val) {
        if (JuchuTransferValueNormalizer.isBlank(val)) {
            return List.of();
        }
        List<String> lines = new ArrayList<>();
        for (String line : val.split("\\n", -1)) {
            String t = line != null ? line.strip() : "";
            if (!t.isEmpty()) {
                lines.add(t);
            }
        }
        return List.copyOf(lines);
    }

    private static boolean singleDateMatch(String left, String right) {
        LocalDate rightFull = JuchuTransferValueNormalizer.parseLocalDate(right);
        LocalDate leftFull = JuchuTransferValueNormalizer.parseLocalDate(left);
        LocalDate leftResolved =
                JuchuTransferValueNormalizer.parseLocalDate(
                        left, rightFull != null ? rightFull : LocalDate.now());
        LocalDate rightResolved =
                JuchuTransferValueNormalizer.parseLocalDate(
                        right, leftFull != null ? leftFull : LocalDate.now());
        if (leftResolved != null && rightResolved != null) {
            return leftResolved.equals(rightResolved);
        }
        return JuchuTransferValueNormalizer.normalizeDateVal(left)
                .equals(JuchuTransferValueNormalizer.normalizeDateVal(right));
    }
}
