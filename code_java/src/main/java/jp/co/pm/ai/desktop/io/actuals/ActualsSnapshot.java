package jp.co.pm.ai.desktop.io.actuals;

import java.util.List;

/** ダッシュボード用の加工実績明細スナップショット。 */
public record ActualsSnapshot(List<String> headers, List<List<String>> rows) {}
