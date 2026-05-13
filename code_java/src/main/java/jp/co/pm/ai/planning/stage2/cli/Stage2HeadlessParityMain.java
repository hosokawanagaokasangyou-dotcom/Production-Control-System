package jp.co.pm.ai.planning.stage2.cli;

import java.util.Map;

import jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv;
import jp.co.pm.ai.planning.stage2.parity.Stage2HeadlessParityRunner;
import jp.co.pm.ai.planning.stage2.parity.Stage2ParityBundle;

/**
 * OS 環境変数だけで段階2 Python→Java 同一検証を実行する CLI（JavaFX 不要）。{@code PM_AI_PLAN_INPUT_PATH} / {@code
 * PM_AI_CODE_PYTHON_DIR} 等は事前にエクスポートしておくこと。
 *
 * <p>任意の第1引数は {@code PM_AI_CODE_PYTHON_DIR} が空のときに使うスクリプトディレクトリ（実行タブ相当のフォールバック）。
 *
 * <p>終了コード: 検証成功（{@link Stage2ParityBundle#allPass()}）で 0、失敗で 1、起動前例外で 2。
 */
public final class Stage2HeadlessParityMain {

    private Stage2HeadlessParityMain() {}

    public static void main(String[] args) {
        try {
            Map<String, String> child = Stage2PythonChildEnv.headlessBaseFromSystemEnv();
            String wb = System.getenv("PM_AI_TASK_INPUT_WORKBOOK");
            if (wb == null) {
                wb = "";
            } else {
                wb = wb.strip();
            }
            String scriptFallback = "";
            if (args != null && args.length > 0 && args[0] != null && !args[0].isBlank()) {
                scriptFallback = args[0].strip();
            }
            Stage2HeadlessParityRunner.Outcome o =
                    Stage2HeadlessParityRunner.run(child, wb, scriptFallback, null, System.out::println);
            if (o.fatalError() != null) {
                System.err.println("[stage2-parity] 致命的エラー: " + o.fatalError().getMessage());
                o.fatalError().printStackTrace(System.err);
                System.exit(2);
                return;
            }
            Stage2ParityBundle b = o.bundle();
            for (String line : b.logLines()) {
                System.out.println(line);
            }
            if (!b.allPass()) {
                System.err.println(b.summary());
                System.exit(1);
                return;
            }
            System.out.println(b.summary());
            System.exit(0);
        } catch (Throwable t) {
            System.err.println("[stage2-parity] 未処理: " + t.getMessage());
            t.printStackTrace(System.err);
            System.exit(2);
        }
    }
}
