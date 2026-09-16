package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Path;

/**
 * 年月ラベル付きの②ファイル。
 *
 * @param label 「2026年7月度」
 * @param ym    年月
 * @param path  ファイル
 */
public record DatedFile(String label, YearMonthKey ym, Path path) {
}
