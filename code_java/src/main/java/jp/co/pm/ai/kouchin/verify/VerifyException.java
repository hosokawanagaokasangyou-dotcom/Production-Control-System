package jp.co.pm.ai.kouchin.verify;

/**
 * 検証を継続できないデータ不備・ファイル不足を表す例外。
 * Python 版 {@code verify_kouchin.py} の {@code sys.exit(...)} に相当する。
 */
public class VerifyException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public VerifyException(String message) {
        super(message);
    }

    public VerifyException(String message, Throwable cause) {
        super(message, cause);
    }
}
