package jp.co.pm.ai.desktop.io.win32;

import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * Windows 実機 POC チェックリスト（手動）。
 *
 * <ol>
 *   <li>270×200 で接続・表示・操作
 *   <li>キーボード / マウス入力
 *   <li>右ペイン埋め込み / 終了時 Pane 削除
 *   <li>プロファイル毎のサイズ差
 *   <li>セキュリティ警告ダイアログ自動操作
 *   <li>RPA 完了 → mstsc 終了 → 既存 UI コールバック
 * </ol>
 */
@EnabledOnOs(OS.WINDOWS)
class MstscWindowEmbedderWindowsPocTest {

    @Test
    void embedInfrastructureAvailableOnWindows() {
        assertTrue(RemoteDesktopLauncher.isSupportedPlatform());
        assertTrue(new MstscWindowEmbedder().isSupported());
    }
}
