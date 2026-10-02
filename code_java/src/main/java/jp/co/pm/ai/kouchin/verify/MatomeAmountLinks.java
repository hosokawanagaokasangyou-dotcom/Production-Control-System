package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** 検証Dの場所文字列から、まとめAAと元シートAAのセル番地を取る。 */
public final class MatomeAmountLinks {

    private static final Pattern PLACE = Pattern.compile(
            "(東レ[^!↔/\\s]{0,24})!\\$?([A-Z]{0,3})\\$?(\\d+)");
    private static final int COL_AA = 26;

    private MatomeAmountLinks() {}

    /** 東レまとめのAA。 */
    public static List<String> matomeAa(String place) {
        return addresses(place, true);
    }

    /** 元シートのAA。 */
    public static List<String> sourceAa(String place) {
        return addresses(place, false);
    }

    private static List<String> addresses(String place, boolean matome) {
        if (place == null || place.isBlank()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        Matcher matcher = PLACE.matcher(place);
        while (matcher.find()) {
            String sheet = matcher.group(1);
            boolean isMatome = FactoryProfile.MATOME_SHEET.equals(sheet);
            if (isMatome != matome) {
                continue;
            }
            int row;
            try {
                row = Integer.parseInt(matcher.group(3));
            } catch (NumberFormatException ex) {
                continue;
            }
            if (row < 1) {
                continue;
            }
            out.add(AmountCellRef.address(SourceRawSheets.copiedSheetName(sheet), COL_AA, row));
        }
        return List.copyOf(out);
    }
}
