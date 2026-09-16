package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Path;
import java.util.List;

/**
 * ②ファイルの当月・過去月・翌月以降の振り分け結果。
 *
 * @param current 当月ファイル（①の対象月と一致するもの）
 * @param prevs   過去月ファイル（新しい月順・直近 {@link FileDiscovery#PREV_MONTH_WINDOW} か月まで）
 * @param nexts   翌月以降のファイル（近い月順）
 */
public record MonthlyFileSet(Path current, List<DatedFile> prevs, List<DatedFile> nexts) {
}
