package jp.co.pm.ai.desktop.config;

import java.util.HashMap;
import java.util.Map;

/**
 * Supplemental descriptions derived from {@code workbook_env_bootstrap.py}, {@code planning_core}, and the
 * desktop bridge (not from OS env). Merged with sheet text in the UI.
 *
 * <p>Variables whose names mention xlwings concern Excel COM automation from the Python stack when Excel
 * invokes those scripts (add-in / legacy macro workflows). They are not prerequisites for the JavaFX
 * desktop launcher path, which runs child Python headlessly for stages 1/2 without xlwings.
 */
public final class EnvVarDocs {

    private static final Map<String, String> LOGIC = new HashMap<>();

    static {
        put(
                "PM_AI_PYTHON",
                "\u672c\u30a2\u30d7\u30ea\u306e\u300cPython\u300d\u5165\u529b\u3068\u9023\u52d5\u3002"
                        + "\u5b50\u30d7\u30ed\u30bb\u30b9\u306e python \u547d\u4ee4\u306b\u4f7f\u7528\u3002");
        put(
                "PM_AI_CODE_PYTHON_DIR",
                "\u30b9\u30af\u30ea\u30d7\u30c8\u6839\uff08task_extract_stage1.py \u7b49\uff09\u3002"
                        + "\u81ea\u52d5\u691c\u51fa\u306f user.dir \u304b\u3089 code/python \u3092\u63a2\u3059\u3002");
        put(
                "PM_AI_REPO_ROOT",
                "Production-Control-System \u306e\u89aa\uff08\u30ea\u30dd\u30b8\u30c8\u30ea\u6839\uff09\u3002"
                        + "PM_AI_CODE_PYTHON_DIR \u672a\u6307\u5b9a\u6642\u306e\u63a8\u5b9a\u306b\u4f7f\u7528\u3002");
        put(
                "PM_AI_OUTPUT_DIR",
                "\u6bb5\u968e1/2 \u306e\u51fa\u529b\u5148\uff08plan_input_tasks.xlsx \u7b49\u3001\u5f93\u6765 code/output"
                        + " \u306b\u76f8\u5f53\uff09\u3002\u672a\u8a2d\u5b9a\u6642\u306f PM_AI_REPO_ROOT"
                        + " \u76f4\u4e0b\u306e output\uff08JavaFX \u3068 planning_core.bootstrap \u3068\u540c\u89e3\u6c7a\uff09\u3002");
        put(
                "PM_AI_WORKSPACE",
                "\u914d\u53f0\u4f5c\u696d\u30eb\u30fc\u30c8\uff08Python \u306e cwd\u3001\u30ed\u30b0/output\u3001"
                        + "Gemini \u8a3c\u660e\u66f8\u306e\u641c\u7d22\u5148\u3002JavaFX \u3068 planning_core.bootstrap "
                        + "\u3067\u6700\u512a\u5148\u3055\u308c\u308b\u3002"
                        + "\u672a\u6307\u5b9a\u6642\u306f PM_AI_CODE_PYTHON_DIR \u306e\u89aa\uff08code\uff09\u304b\u3089"
                        + "\u63a8\u5b9a\u3059\u308b\u5834\u5408\u304c\u591a\u3044\u3002");
        put(
                "PM_AI_PROCESSING_PLAN_PATH",
                "\u6bb5\u968e1\u7528\uff1a\u52a0\u5de5\u8a08\u753bDATA\u76f8\u5f53\u306e\u8868\uff08CSV/Parquet/xlsx\uff09\u3002"
                        + "Python \u306f\u672a\u8a2d\u5b9a\u307e\u305f\u306f\u30d5\u30a1\u30a4\u30eb\u7121\u3057\u306e\u3068\u304d"
                        + "\u3001PM_AI_TASK_INPUT_SOURCE_DIR \u5185\u306e\u6700\u65b0\u8868\u3092\u81ea\u52d5\u3067"
                        + "\u3053\u306e\u5909\u6570\u306b\u8a2d\u5b9a\uff08dispatch_workspace.resolve_processing_plan_path_from_env\uff09\u3002"
                        + "run_stage1_extract \u306f\u3053\u306e\u30d1\u30b9\uff08\u307e\u305f\u306f SOURCE_DIR"
                        + "\u89e3\u6c7a\u306e\u5b9f\u5728\u30d5\u30a1\u30a4\u30eb\uff09\u304c\u5fc5\u8981\u3002"
                        + "\u914d\u53f0\u4e0d\u8981\u306f master.xlsm \u304b\u3089 json/stage1_exclude_rules.json \u306b\u66f8\u304d\u51fa\u3057\u3002"
                        + " \u6b63\u5f0f\u306a\u5217\u69cb\u6210\u306f plan/01_\u52a0\u5de5\u8a08\u753bDATA_\u5358\u4e00\u30d5\u30a1\u30a4\u30eb.m"
                        + " \u3068\u540c\u7b49\u306e Power Query \u6210\u5f62\u5f8c\u306e\u52a0\u5de5\u8a08\u753bDATA\u76f8\u5f53\u3002"
                        + "\u751f\u306e\u554f\u5408\u305b xlsx \u3092\u76f4\u63a5\u6307\u5b9a\u3059\u308b\u5834\u5408\u306f Python"
                        + " \u5074\u3067\u30d8\u30c3\u30c0\u30fc\u884c\u30fb\u5217\u540d\u306e\u6551\u6e08\u306e\u307f\uff08PQ"
                        + " \u306e\u8907\u5408\u898b\u51fa\u3057\u3084\u65e5\u4ed8\u5217\u540d\u306e\u5c55\u958b\u306f\u518d\u73fe\u3057\u306a\u3044\uff09\u3002"
                        + "\u78ba\u5b9f\u306b\u5408\u308f\u305b\u308b\u3068\u304d\u306f\u30af\u30a8\u30ea\u66f4\u65b0\u5f8c\u306e CSV"
                        + " \u7b49\u306e\u30d1\u30b9\u3092\u6307\u5b9a\u3059\u308b\u3053\u3068\u3002");
        put(
                "PM_AI_PLAN_INPUT_PATH",
                "\u5c02\u7528UI\u3067\u6307\u5b9a\u3057\u305f\u914d\u53f0\u8a08\u753b\u30bf\u30b9\u30af\u5165\u529b\u30d5\u30a1\u30a4\u30eb\u3078\u306e\u30d1\u30b9"
                        + "\uff08CSV / Parquet / xlsx / xlsm \u7b49\uff09\u3002\u6bb5\u968e2\u306e"
                        + " load_planning_tasks_df \u306f\u8868\u5f62\u5f0f\u3092\u8aad\u3080\u305f\u3081"
                        + "\u3001\u5fc5\u305a\u3057\u3082Excel\u30d6\u30c3\u30af\u3067\u306f\u306a\u3044\u3002"
                        + "JavaFX \u3067\u306e\u904b\u7528\u3067\u306fxlwings\uff08Excel \u30a2\u30c9\u30a4\u30f3\u9023\u643a\u7528\u306e COM\u64cd\u4f5c\uff09\u306f\u672c\u30a2\u30d7\u30ea\u306e\u5fc5\u9808\u3067\u306f\u306a\u3044\u3002"
                        + "Excel \u304b\u3089\u8d77\u52d5\u3059\u308bPython\u7d4c\u8def\u3067\u30d6\u30c3\u30af\u3092\u958b\u304f\u3068\u304d\u306e\u307f"
                        + "\u3001\u5bfe\u8c61\u51e6\u7406\u306b\u5b9f\u5728\u3059\u308bxlsx/xlsm \u304c\u5f79\u306b\u7acb\u3064\u3002"
                        + "\u8a2d\u5b9a\u6642\u306f\u30de\u30af\u30ed\u30d6\u30c3\u30af\u306e\u305d\u306e\u30b7\u30fc\u30c8\u3092\u5143\u306b\u3057\u306a\u3044\u3002");
        put(
                "PM_AI_PROCESSING_PLAN_SHEET",
                "PM_AI_PROCESSING_PLAN_PATH \u304c xlsx \u306e\u3068\u304d\u306e\u30b7\u30fc\u30c8\u6307\u5b9a\u3002\u7a7a\u3067"
                        + "\u5148\u982d\u30b7\u30fc\u30c8\uff080\u756a\uff09\u3002\u5358\u4e00\u30b7\u30fc\u30c8\u306a\u3089"
                        + "\u540d\u524d\u4e0d\u8981\u3002\u8907\u6570\u30b7\u30fc\u30c8\u3067\u540d\u524d\u3092\u6307\u3059\u5834\u5408\u306f"
                        + "\u6587\u5b57\u5217\u3002\u6570\u5024\u306e\u307f\uff08\u4f8b: 1\uff09\u306f 0\u59cb\u307e\u308a\u306e"
                        + "\u30a4\u30f3\u30c7\u30c3\u30af\u30b9\u3002");
        put(
                "PM_AI_PROCESSING_PLAN_HEADER_ROW",
                "xlsx \u8aad\u8fbc\u307f\u6642\u306e\u5217\u540d\u884c\uff08Excel \u306e 1 \u59cb\u307e\u308a\u306e"
                        + "\u884c\u756a\u53f7\uff09\u3002\u7a7a\u3067\u3001\u540c\u4e00\u884c\u306b\u300c\u4f9d\u983cNO\u300d"
                        + "\u3068\u300c\u5de5\u7a0b\u540d\u300d\u3042\u308b\u6700\u4e0a\u4f4d\u306e\u884c\u3092"
                        + "\u81ea\u52d5\u63a2\u77e5\uff08\u5de5\u7a0b\u5225\u751f\u7523\u8a08\u753b\u554f\u5408\u305b"
                        + "\u306a\u3069\u5148\u982d\u306b\u30e1\u30bf\u884c\u304c\u3042\u308b\u30d6\u30c3\u30af\u306f"
                        + "\u901a\u5e38 6 \u884c\u76ee\uff09\u3002");
        put(
                "PM_AI_KOUBAI_INQUIRY_SHAPING",
                "\u5de5\u7a0b\u5225\u554f\u5408\u305b"
                        + " xlsx: "
                        + "6+5"
                        + "\u884c"
                        + "\u8907\u5408\u898b\u51fa\u3057"
                        + "\u3001"
                        + "\u52a0\u5de5\u6642\u9593"
                        + "/"
                        + "\u52a0\u5de5\u901f\u5ea6"
                        + "\u5217\u524a\u9664"
                        + "\u3001"
                        + "\u52a0\u5de5\u6570\u91cf"
                        + "\u306e\u90e8\u5206\u9664\u53bb"
                        + "\uff08\u898b\u51fa\u3057\u304c\u300c\u52a0\u5de5\u6570\u91cf\u300d\u306e\u307f\u306e\u5217\u306f\u5217\u540d\u7dad\u6301\uff09"
                        + "\u3001"
                        + "YYYY/MM/DD"
                        + "\u3002"
                        + "\u7a7a"
                        + "=auto, 0=off, 1=force.");
        put(
                "PM_AI_TABULAR_CSV_ENCODING",
                "PM_AI_PROCESSING_PLAN_PATH \u7b49 CSV \u306e\u6587\u5b57\u30b3\u30fc\u30c9\uff08\u7a7a\u3067 utf-8-sig\uff09\u3002");
        put(
                "PM_AI_GLOBAL_PRIORITY_OVERRIDE_PATH",
                "\u6bb5\u968e2 \u30e1\u30a4\u30f3\u300c\u30b0\u30ed\u30fc\u30d0\u30eb\u30b3\u30e1\u30f3\u30c8\u300d\u4ee3\u66ff: UTF-8"
                        + " \u30c6\u30ad\u30b9\u30c8\u30d5\u30a1\u30a4\u30eb1\u672c\uff08\u30d1\u30b9\u3042\u308c\u3070"
                        + " Excel \u30b7\u30fc\u30c8\u30b9\u30ad\u30e3\u30f3\u306a\u3057\uff09\u3002"
                        + " input_resolution / load_main_sheet_global_priority_override_text\u3002");
        put(
                "PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV",
                "\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7\u306e\u5217\u8a2d\u5b9a\uff08\u5217\u540d\u3001\u8868\u793a"
                        + "\u5217\u3092\u6301\u3064 CSV\u3002\u3042\u308c\u3070\u5217\u8a2d\u5b9a\u30b7\u30fc\u30c8\u8aad\u307f\u3092\u30b9\u30ad\u30c3\u30d7\u3002");
        put(
                "PM_AI_COLUMN_CONFIG_WORKBOOK",
                "\u5217\u8a2d\u5b9a_\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7\u30b7\u30fc\u30c8\u3092\u542b\u3080"
                        + " xlsx/xlsm\u3002PM_AI_PLAN_INPUT_PATH \u3068\u7570\u306a\u308b\u5217\u8a2d\u5b9a\u5c02\u7528\u30d6\u30c3\u30af"
                        + "\u3092\u6307\u3059\u5834\u5408\u3002");
        put(
                "PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK",
                "\u52a0\u5de5\u8a08\u753bDATA\u7b49\u304b\u3089\u300c\u30c7\u30fc\u30bf\u62bd\u51fa\u6642\u9593\u300d\u5217\u3092"
                        + "\u8aad\u3080\u30d6\u30c3\u30af\uff08\u672a\u6307\u5b9a\u6642\u306f planning_core \u306e"
                        + " input_resolution \u306b\u3088\u308b\u63a2\u7d22\u3001PM_AI_PLAN_INPUT_PATH \u306a\u3069\uff09\u3002");
        put(
                "PM_AI_ACTUALS_DATA_WORKBOOK",
                "\u52a0\u5de5\u5b9f\u7e3eDATA \u30b7\u30fc\u30c8\u3092\u8aad\u3080\u30d6\u30c3\u30af\u3002"
                        + "\u672a\u8a2d\u5b9a\u6642\u306f PM_AI_ACTUAL_DETAIL_WORKBOOK \u2192"
                        + " PM_AI_ACTUAL_DETAIL_SOURCE_DIR \u5185\u6700\u65b0 xlsx/xlsm"
                        + " \u2192 PM_AI_PLAN_INPUT_PATH \u304cExcel\u306e\u3068\u304d\u305d\u306e\u30d6\u30c3\u30af"
                        + " \u3068\u5b9f\u7e3e\u660e\u7d30\u3068\u540c\u3058\u65e2\u5b9a\u63a2\u7d22\uff08input_resolution\uff09\u3002");
        put(
                "PM_AI_ACTUALS_DATA_SHEET",
                "PM_AI_ACTUALS_DATA_WORKBOOK \u5185\u306e\u30b7\u30fc\u30c8\u6307\u5b9a\u3002\u7a7a\u3067"
                        + "\u5148\u982d\u30b7\u30fc\u30c8\uff080\u756a\uff09\u3002\u5358\u4e00\u30b7\u30fc\u30c8\u306a\u3089\u540d\u524d\u4e0d\u8981\u3002"
                        + "\u6570\u5024\u306e\u307f\u306f 0\u59cb\u307e\u308a\u306e\u30a4\u30f3\u30c7\u30c3\u30af\u30b9\u3002");
        put(
                "PM_AI_ACTUAL_DETAIL_SHEET",
                "PM_AI_ACTUAL_DETAIL_WORKBOOK \u7b49\u3067\u8aad\u3080\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30\u306e\u30b7\u30fc\u30c8\u6307\u5b9a\u3002\u7a7a\u3067"
                        + "\u5148\u982d\u30b7\u30fc\u30c8\uff080\u756a\uff09\u3002\u5358\u4e00\u30b7\u30fc\u30c8\u306a\u3089\u540d\u524d\u4e0d\u8981\u3002"
                        + "\u6570\u5024\u306e\u307f\u306f 0\u59cb\u307e\u308a\u306e\u30a4\u30f3\u30c7\u30c3\u30af\u30b9\u3002");
        put(
                "PM_AI_TASK_INPUT_SOURCE_DIR",
                "PQ-A \u52a0\u5de5\u8a08\u753bDATA\u53d6\u5f97\u5143\uff08plan/01_*.m \u306e Folder.Files \u3068\u540c\u7cfb\uff09\u3002"
                        + "\u672a\u8a2d\u5b9a\u6642\u306f \\\\192.168.0.101\\\u5171\u6709...\u25cfDATA\\\u751f\u7523\u8a08\u753b\u554f\u5408\u305b\u3002"
                        + "JavaFX \u521d\u671f\u5024\u306f AppPaths.resolveTaskInputSourceDir\u3002"
                        + "Python \u306f PM_AI_PROCESSING_PLAN_PATH \u304c\u672a\u8a2d\u5b9a\u307e\u305f\u306f\u5b58\u5728\u3057\u306a\u3044\u3068\u304d"
                        + "\u3001\u3053\u306e\u30d5\u30a9\u30eb\u30c0\u5185 CSV/Parquet/xlsx \u7b49\u306e\u3046\u3061"
                        + "\u66f4\u65b0\u6642\u523b\u304c\u6700\u65b0\u306e1\u4ef6\u3092\u30bf\u30b9\u30af\u5165\u529b\u306b\u4f7f\u7528\u3002");
        put(
                "PM_AI_ACTUAL_DETAIL_SOURCE_DIR",
                "\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA \u51fa\u529b\u5143\uff08plan/02__q*.m \u306e Folder.Files \u3068\u540c\u7cfb\uff09\u3002"
                        + "planning_core \u306f\u3053\u306e\u30d5\u30a9\u30eb\u30c0\u5185\u306e\u6700\u65b0 xlsx/xlsm"
                        + " \u3092\u5b9f\u7e3e\u660e\u7d30\u8aad\u8fbc\u306e\u5143\u306b\u3059\u308b\u3002"
                        + " PM_AI_ACTUALS_DATA_WORKBOOK \u672a\u8a2d\u5b9a\u6642\u306f"
                        + "\u52a0\u5de5\u5b9f\u7e3eDATA \u8aad\u8fbc\u3082\u540c\u3058\u6700\u65b0\u30d5\u30a1\u30a4\u30eb\u3092\u4f7f\u7528\u3002"
                        + "\u672a\u8a2d\u5b9a\u6642\u306f 002  \u52a0\u5de5G\\\u25cf\u691c\u67fb\u8868\u4f5c\u6210\\\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA\u7cfb UNC\u3002"
                        + "PM_AI_ACTUAL_DETAIL_WORKBOOK \u3067\u5358\u4e00\u30d5\u30a1\u30a4\u30eb\u3092\u512a\u5148\u3002");
        put(
                "PM_AI_ACTUAL_DETAIL_WORKBOOK",
                "\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA\u3092\u8aad\u3080\u30d6\u30c3\u30af\u306e\u30d5\u30eb\u30d1\u30b9\uff08\u6307\u5b9a\u6642\u306f"
                        + " PM_AI_ACTUAL_DETAIL_SOURCE_DIR \u3088\u308a\u512a\u5148\uff09\u3002");
        put(
                "PM_AI_RESULT_DISPATCH_TABLE_DIR",
                "Power Query _q\u7d50\u679c_\u914d\u53f0\u8868 \u53c2\u7167\u7528\u306e"
                        + " \u7d50\u679c_\u914d\u53f0\u8868.xlsx \u51fa\u529b\u5148\uff08\u30de\u30af\u30ed\u30d6\u30c3\u30af\u5074\u306b"
                        + " \u30d5\u30a9\u30eb\u30c0\u30d1\u30b9\u540d\u3092\u5408\u308f\u305b\u308b\u5834\u5408\uff09\u3002"
                        + "\u672a\u8a2d\u5b9a\u6642\u306f\u6bb5\u968e2\u306f PM_AI_WORKSPACE \u307e\u305f\u306f"
                        + " PM_AI_PLAN_INPUT_PATH \u89aa\u968e\u5c64\u306b\u5408\u308f\u305b\u308b\u5834\u5408\u304c\u3042\u308b\u3001"
                        + "JavaFX \u521d\u671f\u5024\u306f PM_AI_REPO_ROOT \u4e0b\u306e code/output\uff08\u4f8b: Production-Control-System/code/output/"
                        + "\u7d50\u679c_\u914d\u53f0\u8868.xlsx \u540c\u968e\u5c64\u306b \u7d50\u679c_\u914d\u53f0\u8868.json \u3082\u51fa\u529b\uff09\u3002");
        put(
                "PM_AI_RESULT_DISPATCH_TABLE_JSON",
                "\u6bb5\u968e2 \u306e \u7d50\u679c_\u914d\u53f0\u8868.json \u51fa\u529b\uff1a"
                        + "0/false/no/off/none \u3067\u7121\u52b9\uff08\u7a7a\u3067\u6709\u52b9\u3001xlsx \u3068\u540c\u30c7\u30fc\u30bf\uff09\u3002");
        put(
                "GEMINI_CREDENTIALS_JSON",
                "Gemini \u6697\u53f7\u5316\u8a3c\u660e\u66f8 JSON\uff08\u4f8b: gemini_credentials.encrypted.json\uff09\u306e"
                        + "\u30d5\u30eb\u30d1\u30b9\u3002planning_core \u3067\u6700\u512a\u5148\u3002"
                        + "JavaFX \u74b0\u5883\u5909\u6570\u30bf\u30d6\u306e\u300c\u30d5\u30a1\u30a4\u30eb...\u300d\u3067\u9078\u629e\u53ef\u3002");
        put(
                "PM_AI_MASTER_WORKBOOK",
                "master \u7cfb .xlsm \u306e\u7d76\u5bfe\u30d1\u30b9\uff08\u5b9f\u5728\u30d5\u30a1\u30a4\u30eb\u306e\u3068\u304d"
                        + " MASTER_WORKBOOK_FILE \u3088\u308a\u512a\u5148\u3002planning_core \u306e"
                        + " \u30de\u30b9\u30bf\u8aad\u8fbc\u30fb\u6a5f\u68b0\u30ab\u30ec\u30f3\u30c0\u30fc\u7b49\u306b\u4f7f\u7528\u3002"
                        + " JavaFX \u306e\u300c\u30de\u30b9\u30bf\u8aad\u8fbc\u30b5\u30de\u30ea\u300d\u30bf\u30d6\u3067\u5185\u5bb9\u3092\u78ba\u8a8d\u53ef\u3002");
        put(
                "PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK",
                "\u5b9f\u884c\u30fb\u30ed\u30b0\u30bf\u30d6\u306e\u300c\u958b\u304f\u300d\uff08\u30b5\u30de\u30ea AI \u914d\u53f0"
                        + " \u7b49\uff09\u304c\u958b\u304f .xlsm \u306e\u7d76\u5bfe\u30d1\u30b9\u3002"
                        + " \u7a7a\u3067 code/ \u4e0b\u306e"
                        + " \u30b5\u30de\u30ea_AI\u914d\u53f0.xlsm\uff08"
                        + "PM_AI_REPO_ROOT \u6e96\u62e0\uff09\u3002"
                        + " \u30d5\u30a1\u30a4\u30eb\u540d\u306e\u307f\u306e\u3068\u304d\u306f code/ \u304b\u3089\u306e\u76f8\u5bfe\u30d1\u30b9"
                        + "\u3068\u3057\u3066\u89e3\u6c7a\u3002");
        put(
                "PM_AI_SKIP_WORKBOOK_ENV_SHEET",
                "1/true \u7b49\u3067 workbook_env_bootstrap \u304c\u30de\u30af\u30ed\u30d6\u30c3\u30af\u306e"
                        + "\u300c\u8a2d\u5b9a_\u74b0\u5883\u5909\u6570\u300d\u30b7\u30fc\u30c8\u3092\u8aad\u307e\u306a\u3044\u3002"
                        + "JavaFX \u74b0\u5883\u5909\u6570\u30bf\u30d6\u304c\u5b50\u30d7\u30ed\u30bb\u30b9\u306e\u6e90\u3002"
                        + " \u7a7a\u306e\u3068\u304d\u30e9\u30f3\u30c1\u30e3\u30fc\u306f 1 \u3092\u4ed8\u4e0e\u3002"
                        + " OS \u74b0\u5883\u5909\u6570\u3078\u306f\u66f8\u304d\u8fbc\u307e\u306a\u3044\u904b\u7528\u3092\u524d\u63d0\u3002");
        put(
                "PM_AI_EXCLUDE_RULES_JSON",
                "\u6bb5\u968e1\uff08run_stage1_extract\uff09\u3067 master.xlsm \u300c\u8a2d\u5b9a_\u914d\u53f0\u4e0d\u8981\u5de5\u7a0b\u300d"
                        + "\u3092 json/stage1_exclude_rules.json \u3078\u66f8\u304d\u51fa\u3057\u3001\u672c\u5909\u6570\u3092"
                        + " \u305d\u306e\u7d76\u5bfe\u30d1\u30b9\u306b\u81ea\u52d5\u8a2d\u5b9a\uff08\u5b50\u30d7\u30ed\u30bb\u30b9\u5185\uff09\u3002"
                        + " \u624b\u52d5\u3067\u3082 UTF-8 JSON\uff08list \u307e\u305f\u306f {\"rules\":[...]}\u3001"
                        + " \u5217\u69cb\u9020\u306f\u8a2d\u5b9a\u30b7\u30fc\u30c8\u3068\u540c\u69d7\u3002"
                        + " \u6709\u52b9\u30d5\u30a1\u30a4\u30eb\u304c\u3042\u308c\u3070 read_excel \u7d4c\u8def\u3092\u7701\u7565\u53ef\u3002"
                        + " JavaFX \u306f\u300c\u30d5\u30a1\u30a4\u30eb...\u300d\u3067\u9078\u629e\u53ef\u3002");
        put(
                "PM_AI_PLAN_RESULT_TASK_JSON",
                "\u6bb5\u968e2 \u51fa\u529b production_plan_*.xlsx \u3068\u540c\u30b8\u30e0\u306e"
                        + " \u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7.json \u8aad\u307f\u66f8\u304d\uff1a"
                        + "0/false/no/off/none \u3067\u7121\u52b9\u3002\u6709\u52b9\u6642\u306f\u518d\u8aad\u8fbc\u3092"
                        + " JSON \u512a\u5148\u306b\u3057\u3066 Excel I/O \u3092\u524a\u6e1b\u3002");
        put(
                "PM_AI_PLAN_RESULT_TASK_JSON_PATH",
                "read_result_task_dataframe \u304c\u8aad\u3080 JSON \u306e\u7d76\u5bfe\u30d1\u30b9"
                        + "\uff08\u5b9f\u5728\u30d5\u30a1\u30a4\u30eb\u306e\u3068\u304d"
                        + " \u51fa\u529b xlsx \u6a2a\u306e\u30b5\u30a4\u30c9\u30ab\u30fc\u30d1\u30b9\u3088\u308a\u512a\u5148\uff09\u3002");
        put(
                "PM_AI_STAGE2_WRITE_EXCEL",
                "\u6bb5\u968e2 \u3067 production_plan / member_schedule \u306e xlsx \u3092\u51fa\u529b\u5148\u306b\u6b8b\u3059\u304b\u3002"
                        + " 0/false/no/off/none \u3067 JSON \u306e\u307f\uff08\u5185\u90e8\u3067\u4e00\u6642 xlsx \u3092\u751f\u6210\u3057 JSON"
                        + " \u51fa\u529b\u5f8c\u306b\u7834\u68c4\uff09\u3002"
                        + " \u672a\u8a2d\u5b9a\u307e\u305f\u306f 1 \u3067\u5f93\u6765\u901a\u308a xlsx \u3082\u51fa\u529b\u3002"
                        + " JavaFX \u306e\u300c\u5b9f\u884c\u30fb\u30ed\u30b0\u300d\u30bf\u30d6\u306e\u30c1\u30a7\u30c3\u30af\u304c\u6bb5\u968e2"
                        + " \u8d77\u52d5\u6642\u306b\u672c\u5909\u6570\u3092\u4e0a\u66f8\u304d\u3059\u308b\u3002"
                        + " 0 \u306e\u3068\u304d\u306f\u8a2d\u5099\u30ac\u30f3\u30c8\uff08\u8a08\u753b\u30fb\u5b9f\u7e3e\u660e\u7d30\uff09\u7cfb\u30b7\u30fc\u30c8\u306f\u4f5c\u6210\u3057\u306a\u3044\uff08\u51e6\u7406\u6642\u9593\u306e\u524a\u6e1b\uff09\u3002");
        put(
                "PM_AI_XLWINGS_STAGE2_DISABLED",
                "1/true/yes/on \u3067\u6bb5\u968e2\u5f8c\u306e xlwings"
                        + "\uff08\u5217\u8a2d\u5b9a\u30b7\u30fc\u30c8\u56f3\u5f62\u8907\u88fd\u7b49\u3001Excel COM/\u30a2\u30c9\u30a4\u30f3\u9023\u643a\u7528\uff09"
                        + " \u3092\u30b9\u30ad\u30c3\u30d7\u3002openpyxl \u306e xlsx \u4fdd\u5b58\u306f\u5f9e\u6765\u901a\u308a\u3002"
                        + "JavaFX \u304b\u3089\u306e\u6bb5\u968e2\u306e\u307f\u306a\u3089\u672c\u6761\u306f\u5b9f\u8cea\u7121\u95a2\u4fc2\u3068\u306a\u308b\u3053\u3068\u304c\u591a\u3044\u3002");
        put(
                "MASTER_WORKBOOK_FILE",
                "master.xlsm \u306e\u30d5\u30a1\u30a4\u30eb\u540d\uff08\u7a7a\u3067 master.xlsm\uff09\u3002"
                        + "\u30de\u30af\u30ed\u30d6\u30c3\u30af\u968e\u5c64\u304b\u3089\u306e\u76f8\u5bfe\u30d1\u30b9\u53ef\u3002"
                        + " PM_AI_MASTER_WORKBOOK \u672a\u6307\u5b9a\u6642\u306e\u89e3\u6c7a\u306b\u4f7f\u7528\u3002"
                        + " \u300c\u30de\u30b9\u30bf\u8aad\u8fbc\u30b5\u30de\u30ea\u300d\u30bf\u30d6\u3068\u9023\u52d5\u3002");
        put(
                "MASTER_USE_SPEED_SHEET",
                "master \u5185 speed \u30b7\u30fc\u30c8\u306b\u3088\u308b\u52a0\u5de5\u901f\u5ea6\u4e0a\u66f8\u304d\u3092\u6709\u52b9\u5316\u3002");
        put(
                "STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST",
                "\u65e5\u5185\u914d\u53f0\u30d5\u30ed\u30fc: 1=\u8a66\u884c\u9806\u512a\u5148\u30de\u30eb\u30c1\u30d1\u30b9\uff08\u65e2\u5b9a\uff09\u3001"
                        + "0=\u5f93\u6765\u30bd\u30fc\u30c8\u3002");
        put(
                "STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT",
                "\u914d\u53f0\u8a66\u884c\u9806\u306e\u300c\u67a0\u300d\u3088\u308a\u5927\u304d\u3044\u9806\u3078\u306e\u5272\u308a\u8fbc\u307f\u5236\u9650\u3002");
        put(
                "STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT",
                "\u6bb5\u968e2\u5f8c\u3001\u5217\u8a2d\u5b9a\u30b7\u30fc\u30c8\u306e\u56f3\u5f62\u3092 xlwings \u3067\u8907\u88fd"
                        + "\uff08Excel \u30a2\u30c9\u30a4\u30f3/\u30de\u30af\u30ed\u9023\u643a\u6642\u3002"
                        + "JavaFX \u3067\u306e headless \u6bb5\u968e2\u306f\u901a\u5e38\u95a2\u4fc2\u306a\u3057\uff09\u3002");
        put(
                "STAGE12_CMD_HIDE_WINDOW",
                "VBA \u7d4c\u7531\u306e\u6bb5\u968e1/2 cmd \u3092\u975e\u8868\u793a\uff081=\u975e\u8868\u793a\uff09\u3002");
        put(
                "PM_AI_CMD_PAUSE_ON_ERROR",
                "CLI \u7d42\u4e86\u6642\u306e pause\uff08Windows\uff09\u3002"
                        + "0/false \u3067\u7121\u52b9\u5316\uff08workbook_env_bootstrap \u540c\u69d8\uff09\u3002");
        put(
                "PYTHONUTF8",
                "\u5b50\u30d7\u30ed\u30bb\u30b9\u3067\u6700\u7d42\u56fa\u5b9a 1\uff08\u672c UI \u3067\u306f\u4e0a\u66f8\u304d\u4e0d\u53ef\uff09\u3002");
        put(
                "PYTHONIOENCODING",
                "\u5b50\u30d7\u30ed\u30bb\u30b9\u3067\u6700\u7d42 utf-8 \u56fa\u5b9a\uff08\u672c UI \u3067\u306f\u4e0a\u66f8\u304d\u4e0d\u53ef\uff09\u3002");
        put(
                "XLWINGS_SUSPEND_AUTO_CALCULATION",
                "xlwings \u304c Excel \u66f8\u304d\u8fbc\u307f\u524d\u5f8c\u3067\u81ea\u52d5\u8a08\u7b97\u3092\u624b\u52d5\u306b\u5207\u66ff\u3048\u308b\u304b"
                        + "\uff08Excel \u30a2\u30c9\u30a4\u30f3\u9023\u643a\u6642\u306e\u307f\u610f\u5473\u304c\u3042\u308b\u3002"
                        + "JavaFX \u304b\u3089\u5b50\u30d7\u30ed\u30bb\u30b9\u3067 Excel \u3092\u64cd\u4f5c\u3057\u306a\u3044\u9650\u308a\u5b9f\u8cea\u672a\u4f7f\u7528\uff09\u3002");
        put(
                "PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY",
                "\u914d\u53f0\u8a66\u884c\u9806\u66f4\u65b0\u6642\u306b post_load\uff08\u4e8b\u5f8c\u5909\u5f62\uff09\u3092\u30b9\u30ad\u30c3\u30d7\u3002");
        put(
                "TASK_PLAN_SHEET",
                "\u914d\u53f0\u8a08\u753b\u30b7\u30fc\u30c8\u540d\uff08\u7a7a\u3067\u65e2\u5b9a\u540d\uff09\u3002");
        put(
                "STAGE2_SERIAL_DISPATCH_BY_TASK_ID",
                "\u65e5\u5185\u914d\u53f0: 1=\u4f9d\u983cNO\u51fa\u73fe\u9806\u3067\u76f4\u5217\uff08\u4ed6\u4f9d\u983c\u306f\u9032\u307e\u306a\u3044\uff09\u3002");
        put(
                "STAGE2_SKIP_SHEET_VISIBILITY_APPLY",
                "\u6bb5\u968e2\u3067\u30b7\u30fc\u30c8\u8868\u793a\u8a2d\u5b9a\u306e\u4e00\u62ec\u9069\u7528\u3092\u30b9\u30ad\u30c3\u30d7\uff08\u9ad8\u901f\u5316\uff09\u3002");
        put(
                "STAGE2_SKIP_SNAPSHOT_EXPORT",
                "\u6bb5\u968e2\u306e pdf/csv \u30b9\u30ca\u30c3\u30d7\u30b7\u30e7\u30c3\u30c8\u51fa\u529b\u3092\u30b9\u30ad\u30c3\u30d7\u3002");
        put(
                "STAGE2_SKIP_MEMBER_SCHEDULE_IMPORT",
                "\u500b\u4eba\u5225\u30b9\u30b1\u30b8\u30e5\u30fc\u30eb\u53d6\u8fbc\u307f\u3092\u30b9\u30ad\u30c3\u30d7\u3002");
        put(
                "PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE",
                "B-2/B-3: \u71b1\u878d\u7740\u691c\u67fb\u306e\u8a2d\u5099\u5360\u6709\u5236\u5fa1\u3002");
        put(
                "PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS",
                "B-2/B-3: EC \u3068\u5f8c\u7d9a\u5de5\u7a0b\u306e\u62c5\u5f53\u8005\u96c6\u5408\u3092\u5206\u96e2\u3002");
        put(
                "WIP_LIMIT_EC_BEFORE_INSP_ROLLS",
                "\u5de5\u7a0b\u9593 WIP: EC\u524d\u301c\u691c\u67fb\u307e\u3067\u306e\u30ed\u30fc\u30eb\u4e0a\u9650\u3002");
        put(
                "RAW_FABRIC_WIDTH_TABLE_PATH",
                "\u539f\u53cd\u5e45 CSV\uff08planning_core \u306e\u5916\u90e8\u8868\u53c2\u7167\uff09\u3002");
        put(
                "PRODUCT_WIDTH_TABLE_PATH",
                "\u88fd\u54c1\u5e45 CSV\u3002\u7a7a\u3060\u3068\u30de\u30af\u30ed\u30d6\u30c3\u30af\u968e\u5c64\u3067\u63a2\u7d22\u3002");
        put(
                "COMPARE_GANTT_SNAPSHOT_DIR",
                "plan_compare_gantt_from_snapshot.py: \u6bd4\u8f03\u5143\u306e\u65e5\u6642\u30d5\u30a9\u30eb\u30c0"
                        + "\uff08pdf \u914d\u4e0b\u306e\u6700\u65b0\u3092\u9078\u629e\u53ef\uff09\u3002");
    }

    private EnvVarDocs() {}

    private static void put(String key, String text) {
        LOGIC.put(key, text);
    }

    /** Logic-derived note, or empty if unknown. */
    public static String logicOnly(String key) {
        if (key == null || key.isBlank()) {
            return "";
        }
        return LOGIC.getOrDefault(key.trim(), "");
    }

    /**
     * Merges sheet/import description with {@link #logicOnly}; avoids duplicate when one contains the other.
     */
    public static String mergeDescriptions(String sheetDescription, String key) {
        String logic = logicOnly(key);
        String s = sheetDescription != null ? sheetDescription.trim() : "";
        if (s.isEmpty()) {
            return logic;
        }
        if (logic.isEmpty()) {
            return s;
        }
        if (s.contains(logic) || logic.contains(s)) {
            return s.length() >= logic.length() ? s : logic;
        }
        return s + " \u2014 " + logic;
    }
}
