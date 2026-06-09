package jp.co.pm.ai.desktop.io.actuals;

import java.util.List;

/** ダッシュボード用の配台予定スナップショット。 */
public record DispatchSnapshot(List<String> headers, List<List<String>> rows) {}
