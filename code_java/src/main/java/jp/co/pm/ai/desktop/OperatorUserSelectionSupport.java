package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.BiFunction;

import javafx.application.Platform;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ChoiceDialog;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.FactorySiteOperatorAccess;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

/** 操作者名選択・PIN 認証フロー（PMD / リモートデスクトップ配布用シェル共通）。 */
public final class OperatorUserSelectionSupport {

    private OperatorUserSelectionSupport() {}

    /** 操作者 ChoiceDialog 用: 当該工場のユーザー管理に未登録であることを示す疑似行。 */
    static final String UNREGISTERED_OPERATOR_PLACEHOLDER = "【ユーザー登録無し】";

    private static String operatorEventLogPrefix(boolean startup) {
        return startup ? "[startup]" : "[operator]";
    }

    static String operatorSelectionScopeLabel(FactorySite factory) {
        if (RemoteDesktopStandaloneBootstrap.isActivated() || factory == FactorySite.RDP_LAUNCHER) {
            return "";
        }
        return factory.displayLabelJa();
    }

    static String operatorSelectionScopeSuffix(FactorySite factory, String dept) {
        if (RemoteDesktopStandaloneBootstrap.isActivated() || factory == FactorySite.RDP_LAUNCHER) {
            return dept.isBlank() ? "" : "部署 " + dept;
        }
        return factory.displayLabelJa() + (dept.isBlank() ? "" : "・部署 " + dept);
    }

    private static String operatorSelectionLogContext(
            FactorySite factory, String dept, String detailSuffix) {
        String scope = operatorSelectionScopeSuffix(factory, dept);
        String body =
                scope.isBlank()
                        ? (detailSuffix != null ? detailSuffix : "")
                        : scope + (detailSuffix != null ? detailSuffix : "");
        return body.isBlank() ? "" : " （" + body + "）";
    }

    public static void requireOperatorSelectionForFactory(
            DesktopShellHost host, FactorySite site, boolean startup) {
        if (host == null) {
            return;
        }
        Stage primary = host.primaryStageForDialogs();
        if (primary == null || primary.getScene() == null) {
            return;
        }
        FactorySite factory =
                RemoteDesktopStandaloneBootstrap.isActivated()
                        ? FactorySite.RDP_LAUNCHER
                        : (site != null
                                ? site
                                : GlobalInitSettingTarget.loadEffective(host.snapshotUiEnv()));
        FactoryOperatorUserStore.configureForCurrentApp(host.snapshotUiEnv(), factory);
        try {
            FactoryOperatorUserStore.ensureStoreFileOnDisk();
        } catch (IOException ex) {
            host.appendLog(
                    operatorEventLogPrefix(startup)
                            + " 操作者一覧の読込に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
        boolean rdpActive = RemoteDesktopStandaloneBootstrap.isActivated();
        if (rdpActive && !startup) {
            performRdpOperatorChange(
                    host,
                    factory,
                    FactoryOperatorUserStore.sessionOperatorName(),
                    FactoryOperatorUserStore.sessionRdpDepartmentKey());
            host.refreshOperatorUserPresentation();
            return;
        }
        if (!startup && !rdpActive) {
            try {
                String current = FactoryOperatorUserStore.sessionOperatorName();
                if (!current.isBlank()) {
                    if (FactoryOperatorUserStore.loginChoicesForFactory(factory).contains(current)) {
                        host.refreshOperatorUserPresentation();
                        return;
                    }
                }
                if (FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(factory)) {
                    host.appendLog(
                            "[factory] 操作者: "
                                    + FactoryOperatorUserStore.sessionOperatorName()
                                    + " （前回選択を復元）");
                    host.refreshOperatorUserPresentation();
                    return;
                }
            } catch (IOException ex) {
                host.appendLog(
                        "[factory] 操作者の復元をスキップ: "
                                + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }
        }
        if (rdpActive) {
            while (FactoryOperatorUserStore.sessionRdpDepartmentKey().isBlank()) {
                if (startup) {
                    try {
                        if (FactoryOperatorUserStore.tryRestoreSessionRdpDepartmentFromLocal()) {
                            break;
                        }
                    } catch (IOException ex) {
                        host.appendLog(
                                "[startup] 部署の前回選択を復元できませんでした: "
                                        + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
                    }
                }
                if (!ensureRdpDepartmentSelected(host, startup)) {
                    host.refreshOperatorUserPresentation();
                    return;
                }
                break;
            }
        }
        if (startup) {
            try {
                if (FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(factory)) {
                    String restored = FactoryOperatorUserStore.sessionOperatorName();
                    String dept = FactoryOperatorUserStore.sessionRdpDepartmentKey();
                    host.appendLog(
                            operatorEventLogPrefix(startup)
                                    + " 操作者: "
                                    + restored
                                    + operatorSelectionLogContext(factory, dept, "・前回選択を復元")
                                    + (FactoryOperatorUserStore.isGuestOperator(restored)
                                            ? " ※ゲスト"
                                            : ""));
                    host.refreshOperatorUserPresentation();
                    return;
                }
            } catch (IOException ex) {
                host.appendLog(
                        "[startup] 操作者の前回選択を復元できませんでした: "
                                + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }
        }
        while (FactoryOperatorUserStore.sessionOperatorName().isBlank()) {
            Optional<String> chosen = promptOperatorUserChoice(host, factory, startup);
            if (chosen.isEmpty()) {
                String scope = operatorSelectionScopeLabel(factory);
                host.showWarningDialog(
                        "操作者名（必須）",
                        (scope.isBlank() ? "" : scope + " の")
                                + "操作者名を選択してください。\n"
                                + "一覧の編集は「ユーザー管理者」タブから行えます。");
                continue;
            }
            String name = chosen.get();
            if ("【ユーザー登録無し】".equals(name) || UNREGISTERED_OPERATOR_PLACEHOLDER.equals(name)) {
                host.showWarningDialog(
                        "操作者名",
                        "当該工場のユーザー管理に登録された操作者名を選んでください。");
                continue;
            }
            if (!confirmSelectedOperatorWithPin(host, factory, name)) {
                continue;
            }
            try {
                FactoryOperatorUserStore.selectSessionOperator(factory, name);
                String dept = FactoryOperatorUserStore.sessionRdpDepartmentKey();
                host.appendLog(
                        operatorEventLogPrefix(startup)
                                + " 操作者: "
                                + name
                                + operatorSelectionLogContext(factory, dept, "")
                                + (FactoryOperatorUserStore.isGuestOperator(name) ? " ※ゲスト" : ""));
            } catch (Exception ex) {
                host.showWarningDialog(
                        "操作者名", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
        host.refreshOperatorUserPresentation();
    }

    /**
     * セッション中の操作者を変更する。取消時は変更前の操作者を復元する。
     *
     * <p>起動時の {@link #requireOperatorSelectionForFactory} とは異なり、既に操作者が選択済みでも必ず選択ダイアログを出す。
     */
    public static void changeSessionOperator(DesktopShellHost host, FactorySite site) {
        if (host == null) {
            return;
        }
        Stage primary = host.primaryStageForDialogs();
        if (primary == null || primary.getScene() == null) {
            return;
        }
        FactorySite factory =
                RemoteDesktopStandaloneBootstrap.isActivated()
                        ? FactorySite.RDP_LAUNCHER
                        : (site != null
                                ? site
                                : GlobalInitSettingTarget.loadEffective(host.snapshotUiEnv()));
        FactoryOperatorUserStore.configureForCurrentApp(host.snapshotUiEnv(), factory);
        try {
            FactoryOperatorUserStore.ensureStoreFileOnDisk();
        } catch (IOException ex) {
            host.appendLog(
                    "[operator] 操作者一覧の読込に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
        String previousOperator = FactoryOperatorUserStore.sessionOperatorName();
        String previousDepartment = FactoryOperatorUserStore.sessionRdpDepartmentKey();
        if (RemoteDesktopStandaloneBootstrap.isActivated()) {
            performRdpOperatorChange(host, factory, previousOperator, previousDepartment);
        } else {
            performProductionOperatorChange(host, factory, previousOperator);
        }
        host.refreshOperatorUserPresentation();
        host.refreshRemoteDesktopOperatorContext();
    }

    /**
     * 本番 PMD 向けの操作者変更（部署選択なし）。取消時は {@code previousOperator} を復元する。
     */
    static void performProductionOperatorChange(
            DesktopShellHost host, FactorySite factory, String previousOperator) {
        performProductionOperatorChange(host, factory, previousOperator, null);
    }

    /** テスト用: ダイアログ差し替え可能な本番操作者変更フロー。 */
    static void performProductionOperatorChange(
            DesktopShellHost host,
            FactorySite factory,
            String previousOperator,
            BiFunction<DesktopShellHost, FactorySite, Optional<String>> operatorPromptOverride) {
        FactoryOperatorUserStore.clearSessionOperatorName();
        Optional<String> chosen =
                operatorPromptOverride != null
                        ? operatorPromptOverride.apply(host, factory)
                        : promptOperatorUserChoice(host, factory, false);
        if (chosen.isEmpty()) {
            restoreSessionOperatorQuietly(host, factory, previousOperator);
            return;
        }
        String name = chosen.get();
        if (UNREGISTERED_OPERATOR_PLACEHOLDER.equals(name) || "【ユーザー登録無し】".equals(name)) {
            host.showWarningDialog(
                    "操作者名", "当該工場のユーザー管理に登録された操作者名を選んでください。");
            restoreSessionOperatorQuietly(host, factory, previousOperator);
            return;
        }
        if (!confirmSelectedOperatorWithPin(host, factory, name)) {
            restoreSessionOperatorQuietly(host, factory, previousOperator);
            return;
        }
        try {
            FactoryOperatorUserStore.selectSessionOperator(factory, name);
            String dept = FactoryOperatorUserStore.sessionRdpDepartmentKey();
            host.appendLog(
                    operatorEventLogPrefix(false)
                            + " 操作者: "
                            + name
                            + operatorSelectionLogContext(factory, dept, "・変更")
                            + (FactoryOperatorUserStore.isGuestOperator(name) ? " ※ゲスト" : ""));
        } catch (Exception ex) {
            host.showWarningDialog(
                    "操作者名", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            restoreSessionOperatorQuietly(host, factory, previousOperator);
        }
    }

    private static void restoreSessionOperatorQuietly(
            DesktopShellHost host, FactorySite factory, String previousOperator) {
        if (previousOperator == null || previousOperator.isBlank()) {
            return;
        }
        try {
            FactoryOperatorUserStore.selectSessionOperator(factory, previousOperator);
        } catch (Exception ex) {
            host.appendLog(
                    "[operator] 操作者の復元に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    private static void restoreSessionOperatorAndDepartmentQuietly(
            DesktopShellHost host,
            FactorySite factory,
            String previousOperator,
            String previousDepartment) {
        if (previousDepartment != null && !previousDepartment.isBlank()) {
            try {
                FactoryOperatorUserStore.selectSessionRdpDepartment(previousDepartment);
            } catch (Exception ex) {
                host.appendLog(
                        "[operator] 部署の復元に失敗: "
                                + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }
        }
        restoreSessionOperatorQuietly(host, factory, previousOperator);
    }

    /**
     * リモートデスクトップ RPA ランチャーの「操作者を変更」専用フロー。
     * 部署選択 → 当該部署の操作者選択。いずれかで取消したら前回の操作者・部署を復元して終了する。
     */
    private static void performRdpOperatorChange(
            DesktopShellHost host,
            FactorySite factory,
            String previousOperator,
            String previousDepartment) {
        performRdpOperatorChange(host, factory, previousOperator, previousDepartment, null, null);
    }

    /** テスト用: ダイアログ差し替え可能な操作者変更フロー。 */
    static void performRdpOperatorChange(
            DesktopShellHost host,
            FactorySite factory,
            String previousOperator,
            String previousDepartment,
            BiFunction<DesktopShellHost, List<String>, Optional<String>> departmentPromptOverride,
            BiFunction<DesktopShellHost, FactorySite, Optional<String>> operatorPromptOverride) {
        FactoryOperatorUserStore.clearSessionOperatorName();
        FactoryOperatorUserStore.clearSessionRdpDepartmentKey();
        if (!selectRdpDepartmentForOperatorChange(host, departmentPromptOverride)) {
            restoreSessionOperatorAndDepartmentQuietly(
                    host, factory, previousOperator, previousDepartment);
            return;
        }
        Optional<String> chosen =
                operatorPromptOverride != null
                        ? operatorPromptOverride.apply(host, factory)
                        : promptOperatorUserChoice(host, factory, false);
        if (chosen.isEmpty()) {
            restoreSessionOperatorAndDepartmentQuietly(
                    host, factory, previousOperator, previousDepartment);
            return;
        }
        String name = chosen.get();
        if (!confirmSelectedOperatorWithPin(host, factory, name)) {
            restoreSessionOperatorAndDepartmentQuietly(
                    host, factory, previousOperator, previousDepartment);
            return;
        }
        try {
            FactoryOperatorUserStore.selectSessionOperator(factory, name);
            String dept = FactoryOperatorUserStore.sessionRdpDepartmentKey();
            host.appendLog(
                    operatorEventLogPrefix(false)
                            + " 操作者: "
                            + name
                            + operatorSelectionLogContext(factory, dept, "")
                            + (FactoryOperatorUserStore.isGuestOperator(name) ? " ※ゲスト" : ""));
        } catch (Exception ex) {
            host.showWarningDialog(
                    "操作者名", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            restoreSessionOperatorAndDepartmentQuietly(
                    host, factory, previousOperator, previousDepartment);
        }
    }

    private static boolean selectRdpDepartmentForOperatorChange(
            DesktopShellHost host,
            BiFunction<DesktopShellHost, List<String>, Optional<String>> departmentPromptOverride) {
        List<String> departments;
        try {
            departments = FactoryOperatorUserStore.listRdpDepartmentKeys();
        } catch (IOException ex) {
            host.showWarningDialog(
                    "部署", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return false;
        }
        if (departments.isEmpty()) {
            host.showWarningDialog(
                    "部署",
                    "部署が未登録です。ユーザー管理者タブで部署を追加してください。");
            return false;
        }
        if (departments.size() == 1) {
            try {
                FactoryOperatorUserStore.selectSessionRdpDepartment(departments.get(0));
                return true;
            } catch (IOException ex) {
                host.showWarningDialog(
                        "部署", ex.getMessage() != null ? ex.getMessage() : ex.toString());
                return false;
            }
        }
        Optional<String> chosen =
                departmentPromptOverride != null
                        ? departmentPromptOverride.apply(host, departments)
                        : promptRdpDepartmentChoice(host, departments, false);
        if (chosen.isEmpty()) {
            return false;
        }
        try {
            FactoryOperatorUserStore.selectSessionRdpDepartment(chosen.get());
            return true;
        } catch (IOException ex) {
            host.showWarningDialog(
                    "部署", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return false;
        }
    }

    private static boolean confirmSelectedOperatorWithPin(
            DesktopShellHost host, FactorySite factory, String name) {
        try {
            if (FactoryOperatorUserStore.hasPin(factory, name)) {
                if (FactoryOperatorUserStore.isPinLocked(factory, name)) {
                    host.showWarningDialog(
                            "PIN ロック",
                            "操作者「"
                                    + name
                                    + "」は PIN を "
                                    + FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES
                                    + " 回連続で間違えたためロックされています。\n"
                                    + "ユーザー管理者タブでロック解除または PIN 再発行してください。");
                    return false;
                }
                Optional<String> verifiedPin = promptAndVerifyOperatorPin(host, factory, name);
                if (verifiedPin.isEmpty()) {
                    return false;
                }
                if (FactoryOperatorUserStore.mustChangePin(factory, name)) {
                    if (!promptRequiredInitialPinChange(host, factory, name, verifiedPin.get())) {
                        return false;
                    }
                }
            }
            return true;
        } catch (Exception ex) {
            host.showWarningDialog(
                    "操作者名", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return false;
        }
    }

    private static boolean ensureRdpDepartmentSelected(DesktopShellHost host, boolean startup) {
        List<String> departments;
        try {
            departments = FactoryOperatorUserStore.listRdpDepartmentKeys();
        } catch (IOException ex) {
            host.showWarningDialog(
                    "部署", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return false;
        }
        if (departments.isEmpty()) {
            host.showWarningDialog(
                    "部署",
                    "部署が未登録です。ユーザー管理者タブで部署を追加してください。");
            return false;
        }
        if (departments.size() == 1) {
            try {
                FactoryOperatorUserStore.selectSessionRdpDepartment(departments.get(0));
                return true;
            } catch (IOException ex) {
                host.showWarningDialog(
                        "部署", ex.getMessage() != null ? ex.getMessage() : ex.toString());
                return false;
            }
        }
        Optional<String> chosen = promptRdpDepartmentChoice(host, departments, startup);
        if (chosen.isEmpty()) {
            host.showWarningDialog(
                    "部署（必須）",
                    "リモートデスクトップ RPA ランチャーを利用する部署を選択してください。\n"
                            + "一覧の編集は「ユーザー管理者」タブから行えます。");
            return false;
        }
        try {
            FactoryOperatorUserStore.selectSessionRdpDepartment(chosen.get());
            return true;
        } catch (IOException ex) {
            host.showWarningDialog(
                    "部署", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return false;
        }
    }

    private static Optional<String> promptRdpDepartmentChoice(
            DesktopShellHost host, List<String> departments, boolean startup) {
        String pref = "";
        if (!startup) {
            pref = FactoryOperatorUserStore.sessionRdpDepartmentKey();
        }
        if (pref.isBlank()) {
            try {
                pref = FactoryOperatorUserStore.lastSelectedRdpDepartmentLocal();
            } catch (IOException ex) {
                pref = "";
            }
        }
        if (pref.isBlank() || !departments.contains(pref)) {
            pref = departments.get(0);
        }
        ChoiceDialog<String> d = new ChoiceDialog<>(pref, departments);
        host.prepareDialogForMainTheme(d);
        d.setTitle("部署の選択");
        d.setHeaderText(null);
        String intro =
                startup
                        ? RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE + " を利用する部署を選んでください。\n"
                        : "操作者を変更する部署を選んでください。\n";
        d.setContentText(
                intro + "（部署の追加・削除はユーザー管理者タブから行えます。）");
        return d.showAndWait();
    }

    private static Optional<String> promptOperatorUserChoice(
            DesktopShellHost host, FactorySite site, boolean startup) {
        List<String> names;
        try {
            names = FactoryOperatorUserStore.loginChoicesForFactory(site);
        } catch (IOException ex) {
            names = new ArrayList<>(FactoryOperatorUserStore.DEFAULT_NAMES);
            if (!names.contains(FactoryOperatorUserStore.GUEST_OPERATOR_NAME)) {
                names.add(FactoryOperatorUserStore.GUEST_OPERATOR_NAME);
            }
        }
        String pref;
        try {
            pref = FactoryOperatorUserStore.lastSelectedForFactory(site);
        } catch (IOException ex) {
            pref = "";
        }
        if (pref.isBlank() || !names.contains(pref)) {
            pref = names.get(0);
        }
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (!session.isBlank()
                && !names.contains(session)
                && FactorySiteOperatorAccess.isFactorySummaryFolderReachable(
                        host.snapshotUiEnv(), site)) {
            names = new ArrayList<>(names);
            names.add(0, UNREGISTERED_OPERATOR_PLACEHOLDER);
            pref = UNREGISTERED_OPERATOR_PLACEHOLDER;
        }
        String appLabel =
                RemoteDesktopStandaloneBootstrap.isActivated()
                        ? RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE
                        : "配台システム";
        ChoiceDialog<String> d = new ChoiceDialog<>(pref, names);
        host.prepareDialogForMainTheme(d);
        d.setTitle("操作者名の選択");
        d.setHeaderText(null);
        d.setContentText(
                (startup ? appLabel + " を利用する操作者名を選んでください。\n" : "")
                        + (RemoteDesktopStandaloneBootstrap.isActivated()
                                ? "（このアプリ専用のユーザー一覧。配台システムとは別管理です。）\n"
                                        + (FactoryOperatorUserStore.sessionRdpDepartmentKey().isBlank()
                                                ? ""
                                                : "部署: "
                                                        + FactoryOperatorUserStore.sessionRdpDepartmentKey()
                                                        + "\n")
                                : "工場: " + site.displayLabelJa() + "\n")
                        + "（一覧の編集はユーザー管理者タブから行えます。）\n"
                        + "「"
                        + FactoryOperatorUserStore.GUEST_OPERATOR_NAME
                        + "」は PIN 不要です。");
        return d.showAndWait();
    }

    private static Optional<String> promptAndVerifyOperatorPin(
            DesktopShellHost host, FactorySite factory, String operatorName) {
        if (host.primaryStageForDialogs() == null) {
            return Optional.empty();
        }
        try {
            if (FactoryOperatorUserStore.isPinLocked(factory, operatorName)) {
                host.showWarningDialog(
                        "PIN ロック",
                        "操作者「"
                                + operatorName
                                + "」は PIN を "
                                + FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES
                                + " 回連続で間違えたためロックされています。\n"
                                + "ユーザー管理者タブでロック解除または PIN 再発行してください。");
                return Optional.empty();
            }
        } catch (IOException ex) {
            host.showWarningDialog("PIN", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return Optional.empty();
        }
        Dialog<String> dialog = new Dialog<>();
        host.prepareDialogForMainTheme(dialog);
        dialog.setTitle("PIN 認証");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "操作者「"
                                + operatorName
                                + "」の PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を入力してください。");
        hint.setWrapText(true);
        PasswordField pf = new PasswordField();
        pf.setPromptText("PIN");
        VBox box = new VBox(8, hint, new Label("PIN:"), pf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        focusInputWhenDialogShown(dialog, pf);
        dialog.setResultConverter(
                bt -> {
                    if (bt != ButtonType.OK) {
                        return null;
                    }
                    String t = pf.getText();
                    return t != null ? t.strip() : "";
                });
        while (true) {
            Optional<String> ans = dialog.showAndWait();
            if (ans.isEmpty()) {
                return Optional.empty();
            }
            String pin = ans.get();
            if (FactoryOperatorUserStore.normalizePin(pin) == null) {
                host.showWarningDialog(
                        "PIN", FactoryOperatorUserStore.pinLengthRangeDescriptionJa() + "を入力してください。");
                continue;
            }
            try {
                FactoryOperatorUserStore.PinVerificationResult result =
                        FactoryOperatorUserStore.verifyPinAttempt(factory, operatorName, pin);
                switch (result) {
                    case SUCCESS -> {
                        return Optional.of(pin);
                    }
                    case LOCKED -> {
                        host.showWarningDialog(
                                "PIN ロック",
                                "操作者「"
                                        + operatorName
                                        + "」は PIN を "
                                        + FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES
                                        + " 回連続で間違えたためロックされました。\n"
                                        + "ユーザー管理者タブでロック解除または PIN 再発行してください。");
                        return Optional.empty();
                    }
                    case WRONG_PIN -> {
                        int remaining = FactoryOperatorUserStore.remainingPinAttempts(factory, operatorName);
                        host.showWarningDialog(
                                "PIN",
                                remaining > 0
                                        ? "PIN が正しくありません。残り "
                                                + remaining
                                                + " 回でロックされます。"
                                        : "PIN が正しくありません。");
                    }
                    default -> host.showWarningDialog("PIN", "PIN が正しくありません。");
                }
            } catch (IOException ex) {
                host.showWarningDialog("PIN", ex.getMessage() != null ? ex.getMessage() : ex.toString());
                return Optional.empty();
            }
        }
    }

    private static boolean promptRequiredInitialPinChange(
            DesktopShellHost host, FactorySite factory, String operatorName, String currentPin) {
        if (host.primaryStageForDialogs() == null) {
            return false;
        }
        Dialog<ButtonType> dialog = new Dialog<>();
        host.prepareDialogForMainTheme(dialog);
        dialog.setTitle("初回 PIN 変更（必須）");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "操作者「"
                                + operatorName
                                + "」は初回ログインのため、新しい PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を設定してください。");
        hint.setWrapText(true);
        PasswordField newPf = new PasswordField();
        newPf.setPromptText("新しい PIN");
        PasswordField confirmPf = new PasswordField();
        confirmPf.setPromptText("新しい PIN（確認）");
        VBox box = new VBox(8, hint, new Label("新しい PIN:"), newPf, new Label("確認:"), confirmPf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK);
        focusInputWhenDialogShown(dialog, newPf);
        while (true) {
            Optional<ButtonType> ans = dialog.showAndWait();
            if (ans.isEmpty() || ans.get() != ButtonType.OK) {
                return false;
            }
            String newPin = newPf.getText() != null ? newPf.getText().strip() : "";
            String confirmPin = confirmPf.getText() != null ? confirmPf.getText().strip() : "";
            if (!newPin.equals(confirmPin)) {
                host.showWarningDialog("初回 PIN 変更", "新しい PIN と確認入力が一致しません。");
                continue;
            }
            if (FactoryOperatorUserStore.normalizePin(newPin) == null) {
                host.showWarningDialog(
                        "初回 PIN 変更",
                        "新しい PIN は "
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + " です。");
                continue;
            }
            try {
                FactoryOperatorUserStore.changePinOnFirstLogin(factory, operatorName, currentPin, newPin);
                host.appendLog("[operator-user] 初回 PIN を変更しました: " + operatorName);
                host.showInformationDialog("初回 PIN 変更", "PIN を変更しました。ログインを続行します。");
                return true;
            } catch (Exception ex) {
                host.showWarningDialog(
                        "初回 PIN 変更", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
    }

    private static void focusInputWhenDialogShown(Dialog<?> dialog, javafx.scene.Node input) {
        if (dialog == null || input == null) {
            return;
        }
        dialog.setOnShown(e -> Platform.runLater(input::requestFocus));
    }
}
