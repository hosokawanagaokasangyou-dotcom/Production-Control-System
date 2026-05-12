package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;

/**
 * Defaults for the environment-variable UI tab, extracted from the UI reference workbook and shipped as
 * {@code /jp/co/pm/ai/desktop/ui_ref_env_defaults.json}. Regenerate with
 * {@link jp.co.pm.ai.desktop.devtool.GenerateUiRefEnvDefaultsJson}.
 */
public final class UiRefEnvDefaults {

    static final String RESOURCE = "/jp/co/pm/ai/desktop/ui_ref_env_defaults.json";

    private UiRefEnvDefaults() {}

    public static List<WorkbookEnvSheetReader.RowEntry> loadOrEmpty() {
        try (InputStream in = UiRefEnvDefaults.class.getResourceAsStream(RESOURCE)) {
            if (in == null) {
                return List.of();
            }
            ObjectMapper om = new ObjectMapper();
            Payload p = om.readValue(in, Payload.class);
            if (p.entries == null) {
                return List.of();
            }
            List<WorkbookEnvSheetReader.RowEntry> list = new ArrayList<>(p.entries.size());
            for (PayloadEntry e : p.entries) {
                if (e == null || e.key == null || e.key.isBlank()) {
                    continue;
                }
                String v = e.value != null ? e.value : "";
                String d = e.description != null && !e.description.isEmpty() ? e.description : null;
                list.add(new WorkbookEnvSheetReader.RowEntry(e.key, v, d));
            }
            return List.copyOf(list);
        } catch (IOException e) {
            return List.of();
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    static final class Payload {
        public String source_workbook;
        public String sheet;
        public List<PayloadEntry> entries;
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    static final class PayloadEntry {
        public String key;
        public String value;
        public String description;
    }
}
