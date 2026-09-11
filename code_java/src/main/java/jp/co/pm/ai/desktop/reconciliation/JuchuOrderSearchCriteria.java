package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
import java.util.Optional;

public record JuchuOrderSearchCriteria(
        LocalDate from,
        LocalDate to,
        String productKeyword,
        String rawMaterialKeyword,
        String machineKeyword,
        String processKeyword) {

    public JuchuOrderSearchCriteria(
            LocalDate from, LocalDate to, String productKeyword, String rawMaterialKeyword) {
        this(from, to, productKeyword, rawMaterialKeyword, "", "");
    }

    public Optional<String> validationError() {
        if (from == null || to == null) {
            return Optional.of("納期範囲を指定してください");
        }
        if (from.isAfter(to)) {
            return Optional.of("開始日が終了日より後です");
        }
        if (isBlankKeyword(productKeyword)
                && isBlankKeyword(rawMaterialKeyword)
                && isBlankKeyword(machineKeyword)
                && isBlankKeyword(processKeyword)) {
            return Optional.of("製品名・投入原反・機械名・工程名のいずれかを入力してください");
        }
        return Optional.empty();
    }

    private static boolean isBlankKeyword(String value) {
        return value == null || value.strip().isEmpty();
    }
}
