package jp.co.pm.ai.desktop.ui;

import java.text.Normalizer;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;

/** 検索付きチェックリストの JavaFX 非依存選択モデル。 */
public final class LimitedOperatorSelectionModel {

    private final List<String> candidates;
    private final List<String> outOfCandidateInitialNames;
    private final LinkedHashSet<String> selected = new LinkedHashSet<>();

    public LimitedOperatorSelectionModel(List<String> candidates, List<String> initiallySelected) {
        this.candidates =
                List.copyOf(
                        new LinkedHashSet<>(
                                candidates != null ? candidates : List.of()));
        LinkedHashSet<String> outOfCandidate = new LinkedHashSet<>();
        if (initiallySelected != null) {
            for (String name : initiallySelected) {
                selected.add(name);
                if (!this.candidates.contains(name)) {
                    outOfCandidate.add(name);
                }
            }
        }
        outOfCandidateInitialNames = List.copyOf(outOfCandidate);
    }

    public List<String> filteredCandidates(String query) {
        return filtered(candidates, query);
    }

    public List<String> filteredDisplayNames(String query) {
        List<String> displayed = new ArrayList<>(candidates);
        displayed.addAll(outOfCandidateInitialNames);
        return filtered(displayed, query);
    }

    public boolean isCandidate(String name) {
        return candidates.contains(name);
    }

    public boolean isSelected(String name) {
        return selected.contains(name);
    }

    public void setSelected(String name, boolean value) {
        if (value) {
            if (isCandidate(name)) {
                selected.add(name);
            }
        } else {
            selected.remove(name);
        }
    }

    public void selectAll(List<String> names) {
        if (names == null) {
            return;
        }
        for (String name : names) {
            setSelected(name, true);
        }
    }

    public void clearAll() {
        selected.clear();
    }

    public List<String> selectedNames() {
        return List.copyOf(selected);
    }

    public List<String> selectedOutOfCandidateNames() {
        List<String> names = new ArrayList<>();
        for (String name : selected) {
            if (!isCandidate(name)) {
                names.add(name);
            }
        }
        return List.copyOf(names);
    }

    public boolean hasSelectedOutOfCandidateNames() {
        return !selectedOutOfCandidateNames().isEmpty();
    }

    public void validateConfirmable() {
        List<String> invalid = selectedOutOfCandidateNames();
        if (!invalid.isEmpty()) {
            throw new IllegalStateException(
                    "資格外/候補外の既存値をチェック解除してください: "
                            + String.join(", ", invalid));
        }
    }

    private static List<String> filtered(List<String> source, String query) {
        String key = normalize(query);
        if (key.isEmpty()) {
            return List.copyOf(source);
        }
        List<String> filtered = new ArrayList<>();
        for (String name : source) {
            if (normalize(name).contains(key)) {
                filtered.add(name);
            }
        }
        return List.copyOf(filtered);
    }

    private static String normalize(String value) {
        String raw = value != null ? value.strip() : "";
        return Normalizer.normalize(raw, Normalizer.Form.NFKC).toLowerCase(Locale.ROOT);
    }
}
