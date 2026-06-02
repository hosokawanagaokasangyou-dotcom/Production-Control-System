package jp.co.pm.ai.desktop.ui;

import java.util.List;

import jp.co.pm.ai.desktop.Stage2UnknownMasterCombinationPrompt.UnknownPair;

/**
 * {@link Stage2UnknownMasterCombinationDialog} の戻り値。
 *
 * <p>ネスト record の {@code $Result.class} が Windows / ネットワークドライブの増分コンパイルで
 * 欠落し FXML ロード時に {@link NoClassDefFoundError} になるのを避けるためトップレベルに分離。
 */
public record Stage2UnknownMasterCombinationDialogResult(List<UnknownPair> markExclude) {}
