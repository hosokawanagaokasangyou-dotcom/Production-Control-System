package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** 保存前競合チェックの結果。 */
public final class ConflictCheckResult {

    public enum Kind {
        OK,
        CONFLICT,
        IO_ERROR
    }

    private final Kind kind;
    private final List<Path> mismatchedPaths;
    private final String errorMessage;

    private ConflictCheckResult(Kind kind, List<Path> mismatchedPaths, String errorMessage) {
        this.kind = Objects.requireNonNull(kind);
        this.mismatchedPaths = mismatchedPaths == null ? List.of() : List.copyOf(mismatchedPaths);
        this.errorMessage = errorMessage;
    }

    public static ConflictCheckResult ok() {
        return new ConflictCheckResult(Kind.OK, List.of(), null);
    }

    public static ConflictCheckResult conflict(List<Path> mismatched) {
        return new ConflictCheckResult(Kind.CONFLICT, mismatched, null);
    }

    public static ConflictCheckResult ioError(String message) {
        return new ConflictCheckResult(Kind.IO_ERROR, List.of(), message);
    }

    public Kind kind() {
        return kind;
    }

    public List<Path> mismatchedPaths() {
        return mismatchedPaths;
    }

    public String errorMessage() {
        return errorMessage;
    }
}
