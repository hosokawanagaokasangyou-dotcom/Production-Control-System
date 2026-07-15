package jp.co.pm.ai.planning.stage2.source;

/** 段階1の旧bundle無効化と新bundle保存を完了条件として扱う。 */
public final class Stage1SourceBundleCompletionGate {

    @FunctionalInterface
    public interface CheckedAction {
        void run() throws Exception;
    }

    public record Result(boolean completionAllowed, String message) {
        public static Result allowed() {
            return new Result(true, "");
        }

        public static Result blocked(String message) {
            return new Result(false, message != null ? message : "bundle処理に失敗しました");
        }
    }

    private Stage1SourceBundleCompletionGate() {}

    public static Result invalidateBeforeStage1(CheckedAction invalidator) {
        if (invalidator == null) {
            return Result.blocked("旧bundleの削除処理がありません。");
        }
        try {
            invalidator.run();
            return Result.allowed();
        } catch (Exception ex) {
            return Result.blocked(
                    "旧当日配台bundleの削除に失敗しました: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    public static Result persist(
            boolean todayDispatch,
            boolean bundleReady,
            CheckedAction invalidator,
            CheckedAction writer) {
        Result invalidation = invalidateBeforeStage1(invalidator);
        if (!invalidation.completionAllowed()) {
            return invalidation;
        }
        if (!todayDispatch) {
            return Result.allowed();
        }
        if (!bundleReady || writer == null) {
            return Result.blocked("当日配台bundleの保存に必要なソースがありません。");
        }
        try {
            writer.run();
            return Result.allowed();
        } catch (Exception ex) {
            return Result.blocked(
                    "当日配台bundleの保存に失敗しました: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }
}
