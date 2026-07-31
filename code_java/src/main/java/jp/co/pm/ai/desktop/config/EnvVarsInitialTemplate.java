package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Predicate;

import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;

/**
 * 環境変数タブ「環境変数を初期化」後と同じ値集合（ui_ref + ブートストラップ + 工場既定 + 表示補完 + フォルダ正規化）の期待値。
 *
 * <p>起動時に現在値と突き合わせ、環境変数タブ以外の操作を制限する判定に使う。
 */
public final class EnvVarsInitialTemplate {

  @FunctionalInterface
  public interface KeyDefaultResolver {
    String resolve(String key, Map<String, String> ui);
  }

  private EnvVarsInitialTemplate() {}

  /**
   * {@link jp.co.pm.ai.desktop.MainShellController#applyEnvRowsFullBundledResetAndPersist} と同順序で期待値を構築する。
   */
  public static Map<String, String> buildExpectedMap(
      List<String> bootstrapOrder,
      FactorySite site,
      KeyDefaultResolver bootstrapDefaults,
      KeyDefaultResolver optionalDisplayDefaults,
      Predicate<String> includeKey) {
    Objects.requireNonNull(bootstrapOrder, "bootstrapOrder");
    FactorySite effective = site != null ? site : FactorySite.KONAN;
    KeyDefaultResolver bootstrap =
        bootstrapDefaults != null ? bootstrapDefaults : (k, u) -> "";
    KeyDefaultResolver optional =
        optionalDisplayDefaults != null ? optionalDisplayDefaults : (k, u) -> "";
    Predicate<String> keyFilter = includeKey != null ? includeKey : k -> true;

    LinkedHashMap<String, String> map = new LinkedHashMap<>();
    for (WorkbookEnvSheetReader.RowEntry e : UiRefEnvDefaults.loadOrEmpty()) {
      String k = e.key() != null ? e.key().trim() : "";
      if (k.isEmpty() || !keyFilter.test(k)) {
        continue;
      }
      map.put(k, nz(e.value()));
    }

    Map<String, String> ui = Map.of();
    for (String k : bootstrapOrder) {
      if (!keyFilter.test(k)) {
        continue;
      }
      String cur = nz(map.get(k));
      if (cur.isEmpty()) {
        cur = nz(bootstrap.resolve(k, ui));
      }
      map.put(k, cur);
      ui = withKey(ui, k, cur);
    }

    overlayFactorySiteValues(map, effective, ui);
    ui = copyMap(map);

    for (String k : bootstrapOrder) {
      if (!keyFilter.test(k)) {
        continue;
      }
      if (nz(map.get(k)).isEmpty()) {
        String v = nz(bootstrap.resolve(k, ui));
        if (!v.isEmpty()) {
          map.put(k, v);
        }
      }
    }
    ui = copyMap(map);

    for (String k : new ArrayList<>(map.keySet())) {
      if (!keyFilter.test(k) || !nz(map.get(k)).isEmpty()) {
        continue;
      }
      String v = nz(optional.resolve(k, ui));
      if (!v.isEmpty()) {
        map.put(k, v);
      }
    }

    Map<String, String> overrides = AppPaths.normalizedFolderEnvOverrides(map);
    map.putAll(overrides);
    return Map.copyOf(map);
  }

  /** 現在の環境変数タブ値が期待テンプレートと一致するか（キーはテンプレート側を正とする）。 */
  public static boolean matches(
      Map<String, String> current, Map<String, String> expected, Predicate<String> includeKey) {
    if (expected == null || expected.isEmpty()) {
      return true;
    }
    Predicate<String> keyFilter = includeKey != null ? includeKey : k -> true;
    Map<String, String> cur = current != null ? current : Map.of();
    for (Map.Entry<String, String> e : expected.entrySet()) {
      String k = e.getKey();
      if (k == null || k.isBlank() || !keyFilter.test(k)) {
        continue;
      }
      if (!nz(cur.get(k)).equals(nz(e.getValue()))) {
        return false;
      }
    }
    return true;
  }

  static void overlayFactorySiteValues(
      Map<String, String> map, FactorySite site, Map<String, String> ui) {
    if (map == null || site == null) {
      return;
    }
    Map<String, String> ctx = ui != null ? ui : Map.of();
    putIfManaged(map, AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, site.taskInputSourceDir());
    putIfManaged(map, AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, site.actualDetailSourceDir());
    putIfManaged(map, AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR, site.portableBundleSourceDir());
    putIfManaged(
        map,
        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
        nz(site.pmAiMasterWorkbookEnvValue(ctx)));
    putIfManaged(
        map,
        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
        nz(site.pmAiSummaryAiDispatchWorkbookEnvValue(ctx)));
    putIfManaged(map, AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR, site.aladdinMasterDir());
    putIfManaged(map, AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE, site.requestFormJuchuFile());
    putIfManaged(map, AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR, site.requestFormTpiPdfDir());
    putIfManaged(
        map, AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR, site.rdpPortableBundleSourceDir());
    putIfManaged(map, AppPaths.KEY_PM_AI_FACTORY_SITE, site.name());
    if (site == FactorySite.KONAN) {
      putIfManaged(
          map, AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR, AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR_M);
      putIfManaged(
          map,
          AppPaths.KEY_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR,
          AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR_M);
    }
  }

  private static void putIfManaged(Map<String, String> map, String key, String value) {
    if (map.containsKey(key)) {
      map.put(key, value != null ? value : "");
    }
  }

  private static Map<String, String> withKey(Map<String, String> ui, String key, String value) {
    LinkedHashMap<String, String> next = new LinkedHashMap<>(ui);
    next.put(key, value != null ? value : "");
    return next;
  }

  private static Map<String, String> copyMap(Map<String, String> source) {
    return new LinkedHashMap<>(source);
  }

  private static String nz(String s) {
    return s != null ? s.trim() : "";
  }
}
