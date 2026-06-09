package jp.co.pm.ai.desktop.io.actuals;

import java.util.List;

/** ダッシュボード用のアラジン予定スナップショット。 */
public record AladdinSnapshot(List<String> headers, List<List<String>> rows) {}
