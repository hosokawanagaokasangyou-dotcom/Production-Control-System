package jp.co.pm.ai.planning.stage2.parity;

/**
 * 段階2 Java/Python 同一検証などの二値結果（一致／不一致とユーザー向け説明）。
 */
public record Stage2ParityCheckResult(boolean identical, String summary) {}
