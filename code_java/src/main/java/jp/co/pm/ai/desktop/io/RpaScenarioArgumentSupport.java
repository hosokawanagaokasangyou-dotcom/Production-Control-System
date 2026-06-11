package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import jp.co.pm.ai.desktop.config.AladdinRpaLaunchArgs;

/**
 * RPA 起動引数内の {@link AladdinRpaLaunchArgs#SCENARIO_FLAG} パスを、空白トークン分割なしで抽出・正規化する。
 */
final class RpaScenarioArgumentSupport {

    private RpaScenarioArgumentSupport() {}

    static String normalizeScenarioArguments(String arguments) {
        return rebuildScenarioArguments(arguments, true);
    }

    static String repairScenarioArguments(String arguments) {
        return rebuildScenarioArguments(arguments, false);
    }

    static List<String> extractScenarioPaths(String arguments) {
        if (arguments == null || arguments.isBlank()) {
            return List.of();
        }
        return parseScenarioAndOtherArguments(arguments.trim(), false).scenarioPaths();
    }

    private static String rebuildScenarioArguments(String arguments, boolean stripEternalFlag) {
        if (arguments == null || arguments.isBlank()) {
            return "";
        }
        ParsedArguments parsed =
                parseScenarioAndOtherArguments(arguments.trim(), stripEternalFlag);
        List<String> normalized = new ArrayList<>(parsed.otherTokens());
        for (String path : parsed.scenarioPaths()) {
            normalized.add(AladdinRpaLaunchArgs.SCENARIO_FLAG);
            normalized.add(UncPathSegmentRepair.repair(path));
        }
        if (normalized.isEmpty()) {
            return "";
        }
        return formatTokensForIniArguments(normalized);
    }

    private static ParsedArguments parseScenarioAndOtherArguments(String input) {
        return parseScenarioAndOtherArguments(input, true);
    }

    private static ParsedArguments parseScenarioAndOtherArguments(
            String input, boolean stripEternalFlag) {
        List<String> scenarioPaths = new ArrayList<>();
        List<String> otherTokens = new ArrayList<>();
        int index = 0;
        while (index < input.length()) {
            index = skipWhitespace(input, index);
            if (index >= input.length()) {
                break;
            }
            if (startsWithFlag(input, index, AladdinRpaLaunchArgs.SCENARIO_FLAG)) {
                index += AladdinRpaLaunchArgs.SCENARIO_FLAG.length();
                index = skipWhitespace(input, index);
                ExtractedPath extracted = extractScenarioPath(input, index);
                if (!extracted.path().isBlank()) {
                    scenarioPaths.add(extracted.path());
                }
                index = extracted.endIndex();
                continue;
            }
            if (tryExtractBareScenarioPath(input, index) instanceof ExtractedPath bare
                    && !bare.path().isBlank()) {
                scenarioPaths.add(bare.path());
                index = bare.endIndex();
                continue;
            }
            if (startsWithFlag(input, index, AladdinRpaLaunchArgs.ID_FLAG)
                    || startsWithFlag(input, index, AladdinRpaLaunchArgs.PASSWORD_FLAG)) {
                index = skipFlagWithValue(input, index);
                continue;
            }
            TokenExtract token = readNextToken(input, index);
            if (!token.token().isEmpty()) {
                if (stripEternalFlag
                        && AladdinRpaLaunchArgs.ETERNAL_FLAG.equalsIgnoreCase(token.token())) {
                    index = token.endIndex();
                    continue;
                }
                otherTokens.add(token.token());
            }
            index = token.endIndex();
        }
        return new ParsedArguments(scenarioPaths, otherTokens);
    }

    private static ExtractedPath extractScenarioPath(String input, int startIndex) {
        if (startIndex >= input.length()) {
            return new ExtractedPath("", startIndex);
        }
        if (input.charAt(startIndex) == '"') {
            int end = findClosingQuote(input, startIndex + 1);
            if (end < 0) {
                return new ExtractedPath("", startIndex);
            }
            String path = input.substring(startIndex + 1, end).replace("\"\"", "\"");
            return new ExtractedPath(path, end + 1);
        }
        int ardrpaIndex = indexOfIgnoreCase(input, ".ardrpa", startIndex);
        if (ardrpaIndex < 0) {
            TokenExtract token = readNextToken(input, startIndex);
            return new ExtractedPath(token.token(), token.endIndex());
        }
        int endIndex = ardrpaIndex + ".ardrpa".length();
        return new ExtractedPath(input.substring(startIndex, endIndex), endIndex);
    }

    private static ExtractedPath tryExtractBareScenarioPath(String input, int startIndex) {
        if (startIndex >= input.length() || input.charAt(startIndex) == '"') {
            return null;
        }
        if (!looksLikePathStart(input, startIndex)) {
            return null;
        }
        int ardrpaIndex = indexOfIgnoreCase(input, ".ardrpa", startIndex);
        if (ardrpaIndex < 0) {
            return null;
        }
        int endIndex = ardrpaIndex + ".ardrpa".length();
        return new ExtractedPath(input.substring(startIndex, endIndex), endIndex);
    }

    private static boolean looksLikePathStart(String input, int startIndex) {
        if (startIndex + 1 < input.length()
                && input.charAt(startIndex) == '\\'
                && input.charAt(startIndex + 1) == '\\') {
            return true;
        }
        return startIndex + 2 < input.length()
                && Character.isLetter(input.charAt(startIndex))
                && input.charAt(startIndex + 1) == ':';
    }

    private static TokenExtract readNextToken(String input, int startIndex) {
        if (startIndex >= input.length()) {
            return new TokenExtract("", startIndex);
        }
        if (input.charAt(startIndex) == '"') {
            int end = findClosingQuote(input, startIndex + 1);
            if (end < 0) {
                return new TokenExtract(input.substring(startIndex), input.length());
            }
            return new TokenExtract(
                    input.substring(startIndex + 1, end).replace("\"\"", "\""), end + 1);
        }
        int index = startIndex;
        while (index < input.length() && !Character.isWhitespace(input.charAt(index))) {
            index++;
        }
        return new TokenExtract(input.substring(startIndex, index), index);
    }

    private static int skipFlagWithValue(String input, int startIndex) {
        TokenExtract token = readNextToken(input, startIndex);
        int index = skipWhitespace(input, token.endIndex());
        if (index >= input.length() || input.charAt(index) == '-') {
            return index;
        }
        return readNextToken(input, index).endIndex();
    }

    private static int skipWhitespace(String input, int startIndex) {
        int index = startIndex;
        while (index < input.length() && Character.isWhitespace(input.charAt(index))) {
            index++;
        }
        return index;
    }

    private static boolean startsWithFlag(String input, int startIndex, String flag) {
        if (startIndex + flag.length() > input.length()) {
            return false;
        }
        return input.regionMatches(true, startIndex, flag, 0, flag.length());
    }

    private static int findClosingQuote(String text, int fromIndex) {
        for (int i = fromIndex; i < text.length(); i++) {
            if (text.charAt(i) != '"') {
                continue;
            }
            if (i + 1 < text.length() && text.charAt(i + 1) == '"') {
                i++;
                continue;
            }
            return i;
        }
        return -1;
    }

    private static int indexOfIgnoreCase(String text, String value, int startIndex) {
        return text.toLowerCase(Locale.ROOT)
                .indexOf(value.toLowerCase(Locale.ROOT), startIndex);
    }

    private static String formatTokensForIniArguments(List<String> tokens) {
        StringBuilder out = new StringBuilder();
        for (String token : tokens) {
            if (token.isEmpty()) {
                continue;
            }
            if (out.length() > 0) {
                out.append(' ');
            }
            out.append(quoteArgumentIfNeeded(token));
        }
        return out.toString();
    }

    private static String quoteArgumentIfNeeded(String token) {
        if (token.indexOf(' ') >= 0 || token.indexOf('\t') >= 0) {
            return "\"" + token.replace("\"", "\"\"") + "\"";
        }
        return token;
    }

    private record ParsedArguments(List<String> scenarioPaths, List<String> otherTokens) {}

    private record ExtractedPath(String path, int endIndex) {}

    private record TokenExtract(String token, int endIndex) {}
}
