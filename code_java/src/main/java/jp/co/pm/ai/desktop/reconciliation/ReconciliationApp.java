package jp.co.pm.ai.desktop.reconciliation;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.scene.image.ImageView;
import javafx.scene.transform.Scale;
import javafx.scene.transform.Translate;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.text.FontWeight;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.io.PoiWorkbookFileWriter;
import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;
import jp.co.pm.ai.desktop.io.PoiWorkbookSaver;
import jp.co.pm.ai.desktop.io.RequestFormJuchuFileBackupStore;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

public class ReconciliationApp {

    private static final String ALADDIN_MASTER_DIR_NAME = AppPaths.ALADDIN_MASTER_DIR_LEAF_NAME;
    private static final String INTEGRATED_MASTER_FILE_NAME = "マスタリレーション統合結果.xlsx";
    private static final String COL_MASTER_BASE_SHOHIN_PRODUCT = "masterBase商品(製品)";
    private static final String COL_MASTER_BASE_SHOHIN_RAW = "masterBase商品(原反)";
    /** 受注ﾌｧｲﾙ: POI lastRowNum が書式だけで膨らんだときの最大走査行数。 */
    private static final int JUCHU_SHEET_MAX_SCAN_ROWS = 20_000;
    private static final double SETTINGS_CARD_WIDTH = 300.0;
    /** 従来 100% 設計からの UI 幅倍率（現状 = 120%）。 */
    private static final double UI_WIDTH_SCALE = 1.2;
    /** 左ペイン幅の 100% 設計値（× {@link #UI_WIDTH_SCALE} が実寸）。 */
    private static final double LEFT_PANE_MIN_WIDTH_BASE = 280.0;
    private static final double LEFT_PANE_PREF_WIDTH_BASE = 480.0;
    private static final double LEFT_PANE_MAX_WIDTH_BASE = 1000.0;
    private static final double LEFT_PANE_MIN_WIDTH = LEFT_PANE_MIN_WIDTH_BASE * UI_WIDTH_SCALE;
    private static final double LEFT_PANE_PREF_WIDTH = LEFT_PANE_PREF_WIDTH_BASE * UI_WIDTH_SCALE;
    private static final double LEFT_PANE_MAX_WIDTH = LEFT_PANE_MAX_WIDTH_BASE * UI_WIDTH_SCALE;
    /** 初回分割位置（左ペイン比率）。120% 表示時は従来 30% × 1.2。 */
    private static final double SPLIT_LEFT_RATIO = Math.min(0.85, 0.30 * UI_WIDTH_SCALE);
    /** フォーム項目名ラベル列の最小幅（100% 設計 × {@link #UI_WIDTH_SCALE}）。 */
    private static final double FORM_LABEL_COLUMN_MIN_WIDTH = 70.0 * UI_WIDTH_SCALE;
    /** 依頼書タブ表示中: 受注ファイル {@code ~$} ロックの確認間隔。 */
    private static final long JUCHU_LOCK_POLL_INTERVAL_MS = 10_000L;
    /** 依頼書タブ表示中: 原本 Excel の更新確認間隔。 */
    private static final long ORIGINAL_FILE_POLL_INTERVAL_MS = 30_000L;
    /** {@link #saveLocalForm()} 後・受注ファイル未転記のステータス。 */
    private static final String STATUS_LOCAL_SAVE_PENDING = "手修正済み (未保存)";

    private Window hostWindow;
    private Path repoRootHint;
    private Path aladdinMasterDir;
    private Map<String, String> uiEnvSnapshot = Map.of();
    private Label embeddedTitleLabel;
    private Consumer<String> originalDirChangeHandler;
    private Consumer<String> juchuFileChangeHandler;
    private TextField txtJuchuPathDisplay;
    private ListView<RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry>
            juchuBackupListView;
    private Label juchuBackupHintLabel;
    private Button btnTransfer;
    private Button btnUndoLastTransfer;
    private Button btnBulkTransferPending;
    private Button btnSaveLocal;
    private Button btnOpenJuchuSide;
    private Button btnJuchuColumnWizardSide;
    private final List<Node> guestMutableControls = new ArrayList<>();
    private Label transferBlockedReasonLabel;
    /** 直前の単票自動転記を受注ファイルで取り消すためのスナップショット（1 件のみ）。 */
    private JuchuTransferUndoState lastJuchuTransferUndo;
    private String targetFolder = "";
    private String juchuFilePath;
    private boolean isLoadingRecord = false;
    /** 受注ファイル転記中（UI スレッドをブロックしないようバックグラウンド実行）。 */
    private volatile boolean juchuTransferInProgress = false;

    /**
     * 自動転記中は数式セルをフォーム値で置換する（既存行の手修正を反映するため）。
     * バックグラウンド転記スレッドでのみ {@code true} にする。
     */
    private static final ThreadLocal<Boolean> JUCHU_TRANSFER_REPLACE_FORMULA =
            ThreadLocal.withInitial(() -> Boolean.FALSE);

    private ComboBox<OrderRecord> comboRecord;
    private TextField txtRecordFilter;
    private RadioButton rbAllRecordsFilter;
    private RadioButton rbExistingOnlyFilter;
    private RadioButton rbNewOnlyFilter;
    private RadioButton rbJuchuWithoutOriginalFilter;
    private ObservableList<OrderRecord> orderRecords = FXCollections.observableArrayList();

    private GridPane sheetGrid;
    private ScrollPane sheetScrollPane;
    
    private Label statusLabel;
    private Label discrepancyLabel;
    private Label lblFormTitle;

    private HBox previewFileNameRow;
    private Label previewFileNameLabel;
    private StackPane previewFileUpdateBadgeHost;
    private File currentPreviewOriginalFile;
    private final RequestFormOriginalUpdateMonitor originalUpdateMonitor =
            new RequestFormOriginalUpdateMonitor();
    private JuchuHeaderAliasRegistry juchuHeaderAliasRegistry =
            JuchuHeaderAliasRegistry.loadDefault();
    private javafx.animation.Timeline juchuLockPollTimeline;
    private javafx.animation.Timeline originalFilePollTimeline;
    private javafx.animation.PauseTransition pollStatusHighlightPause;
    private Label embeddedTabPollStatusLabel;
    private Supplier<RequestFormPreviewBadgeConfig> previewBadgeConfigSupplier =
            RequestFormPreviewBadgeConfig::defaults;
    private Supplier<Boolean> planningPipelineStageBusyChecker = () -> false;
    private static final String PDF_SUPPRESSED_PIPELINE_MSG =
            "段階1～段階3.5 実行中のため PDF プレビュー生成を一時停止しています。";

    // Loading overlay components
    private StackPane mainStackPane;
    private VBox loadingOverlay;
    private Label loadingOverlayLabel;
    
    // Bottom Form inputs (Reconciliation View)
    private TextField txtReqNo;
    private TextField txtItemCode;
    private TextField txtProductCode;
    private TextField txtQty;
    // private TextField txtUser; // unified with newCmbFormUser
    // private TextField txtDelivery; // unified with newDpFormDeliv
    private TextField txtRawMat;
    private TextField txtProcess;
    // private TextField txtContract; // moved to ProductRow#txtKeiyakuNo
    
    private OrderRecord selectedRecord;

    // Master Product list for cache search
    
    // Dynamic multiple rows VBox and list properties
    private VBox productRowsContainer;
    private VBox rawRowsContainer;
    private final List<ProductRow> productRows = new ArrayList<>();
    private final List<RawMaterialRow> rawRows = new ArrayList<>();
private final List<ProductInfo> masterProductList = new ArrayList<>();
    private volatile long masterProductListLoadedMtime = -1L;
    private volatile String masterProductListLoadedPath = "";
    private volatile boolean masterProductListLoadInFlight;

    // New Tab 2 Input Fields (Paper mockup sheet layout)
    private TextField newTxtProdItem;
    private TextField newTxtProdPart;
    private TextField newTxtProdType;
    private TextField newTxtProdWidth;
    private TextField newTxtProdLength;
    private TextField newTxtProdGrade;
    private TextField newTxtProdColor;
    private TextField newTxtProdCategory;
    private TextField newTxtProdEdaban;
    private TextField newTxtProdQty;
    private TextField newTxtRawItem;
    private TextField newTxtRawPart;
    private TextField newTxtRawType;
    private TextField newTxtRawWidth;
    private TextField newTxtRawLength;
    private TextField newTxtRawGrade;
    private TextField newTxtRawColor;
    private TextField newTxtRawCategory;
    private TextField newTxtRawQty;

    // ComboBox fields (user-selectable, editable via Settings tab)
    private ComboBox<String> newCmbProdEcSide;       // ＥＣ面 (製品)
    private ComboBox<String> newCmbProdTrimming;     // ﾄﾘﾐﾝｸﾞ (製品)
    private ComboBox<String> newCmbRawFeedLoc;       // 投入場所 (原反)
    private ComboBox<String> newCmbRawStorageLoc;    // 在庫場所 (原反)
    private ComboBox<String> newCmbFormYoto;
    private ComboBox<String> newCmbFormInputKbn;
    private ComboBox<String> newCmbFormKakoKbn;
    private Label lblFormInputTanto;
    private ComboBox<String> newCmbFormUser;         // ユーザー
    private ComboBox<String> newCmbInputKbn;         // 入力区分 (header)
    private ComboBox<String> newCmbKakoKbn;          // 加工区分 (header)
    /** 設定タブ: 【作業指示】入力区分・加工区分の新規行既定 */
    private ComboBox<String> cmbSettingsDefaultInputKbn;
    private ComboBox<String> cmbSettingsDefaultKakoKbn;
    private ComboBox<String> newCmbWariSu;           // 割数 (product row)

    private TextField newTxtFormTokki1;
    private TextField newTxtFormTokki2;
    private TextField newTxtFormTokki3;
    private DatePicker newDpFormDeliv;
    private DatePicker newDpFormAdjustDeliv;
    private DatePicker newDpFormInputDate;
    private TextField newTxtFormWage;
    
    private TextField newTxtUketsukeNo;
    private TextField newTxtIraiNo;

    // Settings data - combo option lists (editable by user)
    private final ObservableList<String> optInputKbn    = FXCollections.observableArrayList("通常入力", "例外入力");
    private final ObservableList<String> optKakoKbn     = FXCollections.observableArrayList("後加工", "TPI");
    private final ObservableList<String> optWariSu      = FXCollections.observableArrayList("1","2","3","5","6","7","8","9","10");
    private final ObservableList<String> optEcSide      = FXCollections.observableArrayList("Ｈ面","Ｑ面","両面","ｽﾗｲｽ面","ｽｷﾝ面","Ｑ面/-","Ｈ面/-","Ｑ面/Ｑ面/-","H面/H面/-");
    private final ObservableList<String> optTrimming    = FXCollections.observableArrayList("有","無","-");
    private final ObservableList<String> optFeedLoc     = FXCollections.observableArrayList("EC","SEC","ｽﾗｲｽ","ｽﾘｯﾄ","ｴﾝﾎﾞｽ","検反","融着");
    private final ObservableList<String> optStorageLoc  = FXCollections.observableArrayList("滋賀","湖南","滋賀/湖南","湖南/中央","山田","中山","中央湖東","湖南/滋賀","奥田");
    private final ObservableList<String> optYoto        = FXCollections.observableArrayList("W（自動車）","B（輸出）","Y（工材）","V（TPI）","A（TPI）","JR（屋根）","P（TPI）");
    private final ObservableList<String> optUser        = FXCollections.observableArrayList("自動転記","ｵｶﾓﾄ","ﾀﾂﾀ","共和ﾚｻﾞｰ","Scientex","共和興","ｻｶｲﾅｺﾞﾔ","ﾀﾞｲｳﾚ","在ｴﾙ","U4059","U5001","張家港","ｲｽﾞﾐ","盟和","高山産業","中央物産");
    /** 製品行マスタ候補コンボ: 商品コード先頭フィルタ（空なら無制限）。 */
    private final ObservableList<String> optMasterCandidatePrefixProduct = FXCollections.observableArrayList();
    /** 原反行マスタ候補コンボ: 商品コード先頭フィルタ（空なら無制限）。 */
    private final ObservableList<String> optMasterCandidatePrefixRaw = FXCollections.observableArrayList();

    private RequestFormComboChoices comboChoicesState = RequestFormComboChoices.bundledDefaults();

    // Settings & Caching fields
    private volatile long lastInteractionTime = System.currentTimeMillis();
    private static final long IDLE_THRESHOLD_MS = 2000;
    /** バックグラウンド PDF プレビューキャッシュ生成の最小間隔（連続生成で UI/CPU を圧迫しない）。 */
    private static final long BACKGROUND_PDF_CACHE_INTERVAL_MS = 10_000L;
    
    private static class CacheTask {
        final File excelFile;
        final String sheetName;
        final File outputFile;

        CacheTask(File excelFile, String sheetName, File outputFile) {
            this.excelFile = excelFile;
            this.sheetName = sheetName;
            this.outputFile = outputFile;
        }
    }

    private final Queue<CacheTask> cacheQueue = new LinkedList<>();
    private Thread cacheThread = null;

    /** フォルダ選択で依頼書原本パスが変わったとき、環境変数タブへ反映するためのコールバック。 */
    public void setOriginalDirChangeHandler(Consumer<String> handler) {
        this.originalDirChangeHandler = handler;
    }

    /** 受注ファイル選択でパスが変わったとき、環境変数タブへ反映するためのコールバック。 */
    public void setJuchuFileChangeHandler(Consumer<String> handler) {
        this.juchuFileChangeHandler = handler;
    }

    public void setPreviewBadgeConfigSupplier(Supplier<RequestFormPreviewBadgeConfig> supplier) {
        this.previewBadgeConfigSupplier =
                supplier != null ? supplier : RequestFormPreviewBadgeConfig::defaults;
        Platform.runLater(this::refreshPreviewFileHeader);
    }

    /** 段階1～段階3.5 実行中は PDF プレビュー生成を抑制するための判定（メインシェルから注入）。 */
    public void setPlanningPipelineStageBusyChecker(Supplier<Boolean> checker) {
        this.planningPipelineStageBusyChecker = checker != null ? checker : () -> false;
    }

    private boolean isPlanningPipelineStageBusy() {
        try {
            return Boolean.TRUE.equals(planningPipelineStageBusyChecker.get());
        } catch (RuntimeException | Error ex) {
            return false;
        }
    }

    /** デザインタブ変更後にプレビュー上部バッジの見た目を再描画する。 */
    public void refreshPreviewBadgeAppearance() {
        Platform.runLater(this::refreshPreviewFileHeader);
    }

    /**
     * メインシェル「依頼書入力」タブへ埋め込むルートを構築する。
     *
     * @param hostWindow ダイアログの親（未接続時は {@code null} 可。シーン接続後に更新してよい）
     * @param repoRootHint {@link AppPaths#resolveRepoRoot}（作業フォルダの既定探索用）
     * @param uiEnv 環境変数タブ（{@link AppPaths#KEY_PM_AI_ALADDIN_MASTER_DIR}、
     *     {@link AppPaths#KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR} 等）
     */
    public Parent buildEmbeddedRoot(Window hostWindow, Path repoRootHint, Map<String, String> uiEnv) {
        this.hostWindow = hostWindow;
        this.repoRootHint = repoRootHint;
        applyRepoRootAsWorkspaceIfPresent(repoRootHint);
        configureFromUiEnv(uiEnv);
        loadSettings();
        ensureJuchuPathDefault();

        // --- TOP MENU BAR ---
        BorderPane root = new BorderPane();
        root.setPadding(new Insets(0)); // Flush against edges
        
        HBox topBar = new HBox(15);
        topBar.getStyleClass().add("top-bar");
        topBar.setAlignment(Pos.CENTER_LEFT);
        
        embeddedTitleLabel = new Label(embeddedWindowTitle(uiEnv));
        embeddedTitleLabel.getStyleClass().add("top-title");
        
        Button btnReload = new Button("データを再読込");
        btnReload.setOnAction(
                e ->
                        requestReloadData(
                                "受注ファイルと原本データを再読込します。",
                                this::loadMasterProductList));
        btnReload.getStyleClass().add("btn-reload");
        
        Button btnSelectFolder = new Button("フォルダ選択");
        btnSelectFolder.getStyleClass().add("btn-reload");
        btnSelectFolder.setOnAction(e -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            directoryChooser.setTitle("依頼書原本フォルダを選択");
            File defaultDir = new File(targetFolder);
            if (defaultDir.exists() && defaultDir.isDirectory()) {
                directoryChooser.setInitialDirectory(defaultDir);
            }
            File selectedDirectory = directoryChooser.showDialog(hostWindow);
            if (selectedDirectory != null) {
                applySelectedOriginalDir(selectedDirectory.getAbsolutePath());
                statusLabel.setText("選択フォルダ: " + targetFolder);
                saveSettings();
                requestReloadData(
                        "フォルダ変更後、データを再読込します。", this::loadMasterProductList);
            }
        });
        
        statusLabel = new Label("データベース読込中...");
        statusLabel.getStyleClass().add("top-status");

        embeddedTabPollStatusLabel = new Label();
        embeddedTabPollStatusLabel.getStyleClass().add("request-form-poll-status");
        embeddedTabPollStatusLabel.setManaged(false);
        embeddedTabPollStatusLabel.setVisible(false);
        HBox.setHgrow(embeddedTabPollStatusLabel, Priority.ALWAYS);

        topBar.getChildren().addAll(
                embeddedTitleLabel,
                btnSelectFolder,
                btnReload,
                embeddedTabPollStatusLabel,
                statusLabel);
        root.setTop(topBar);

        // --- SPLIT WORKSPACE (Tab 1 Content) ---
        SplitPane splitPane = new SplitPane();
        BorderPane.setMargin(splitPane, new Insets(15));
        
        // 1. LEFT PANE: ScrollPane container for detailed inputs
        ScrollPane leftScroll = new ScrollPane();
        leftScroll.setFitToWidth(true);
        leftScroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        leftScroll.setStyle("-fx-background-color: transparent; -fx-border-color: transparent;");
        leftScroll.setMinWidth(LEFT_PANE_MIN_WIDTH);
        leftScroll.setPrefWidth(LEFT_PANE_PREF_WIDTH);
        leftScroll.setMaxWidth(LEFT_PANE_MAX_WIDTH);
        
        VBox leftContainer = new VBox(15);
        leftContainer.getStyleClass().add("pane-container");
        leftContainer.setPadding(new Insets(10));
        leftContainer.setFillWidth(true);
        leftScroll.setContent(leftContainer);
        
        Label lblLeftTitle = new Label("受注データベース & 選択フィルタ");
        lblLeftTitle.getStyleClass().add("pane-title-left");
        
        HBox filterRowPrimary = new HBox(8);
        filterRowPrimary.setAlignment(Pos.CENTER_LEFT);
        Label lblSearch = new Label("検索・絞り込み:");
        lblSearch.setStyle("-fx-font-weight: bold;");
        txtRecordFilter = new TextField();
        txtRecordFilter.setPromptText("依頼No / ユーザー...");
        txtRecordFilter.setPrefWidth(160 * UI_WIDTH_SCALE);
        txtRecordFilter.textProperty().addListener((obs, oldVal, newVal) -> applyRecordFilter());
        String radioStyle = "-fx-font-size: 11px;";
        rbAllRecordsFilter = new RadioButton("全部");
        rbAllRecordsFilter.setStyle(radioStyle);
        rbExistingOnlyFilter = new RadioButton("既存のみ");
        rbExistingOnlyFilter.setStyle(radioStyle);
        rbNewOnlyFilter = new RadioButton("新規のみ");
        rbNewOnlyFilter.setStyle(radioStyle);
        rbJuchuWithoutOriginalFilter = new RadioButton("原本なし（受注のみ）");
        rbJuchuWithoutOriginalFilter.setStyle(radioStyle);
        rbJuchuWithoutOriginalFilter.setTooltip(
                new Tooltip("依頼書原本にないが受注ファイルには存在するタスクのみを表示（入力日が新しい順）"));
        ToggleGroup recordListFilterGroup = new ToggleGroup();
        for (RadioButton rb :
                List.of(
                        rbAllRecordsFilter,
                        rbExistingOnlyFilter,
                        rbNewOnlyFilter,
                        rbJuchuWithoutOriginalFilter)) {
            rb.setToggleGroup(recordListFilterGroup);
            rb.setMinWidth(Region.USE_PREF_SIZE);
            rb.setMaxWidth(Double.MAX_VALUE);
            installRecordFilterRadioDeselectOnReselect(recordListFilterGroup, rb);
        }
        recordListFilterGroup.selectedToggleProperty().addListener((obs, oldT, newT) -> applyRecordFilter());
        HBox filterModeRow1 = new HBox(12);
        filterModeRow1.setAlignment(Pos.CENTER_LEFT);
        filterModeRow1.getChildren().addAll(rbAllRecordsFilter, rbExistingOnlyFilter);
        HBox filterModeRow2 = new HBox(12);
        filterModeRow2.setAlignment(Pos.CENTER_LEFT);
        filterModeRow2.getChildren().addAll(rbNewOnlyFilter, rbJuchuWithoutOriginalFilter);
        VBox filterModePanel = new VBox(4);
        filterModePanel.setFillWidth(true);
        filterModePanel.getChildren().addAll(filterModeRow1, filterModeRow2);
        filterRowPrimary.getChildren().addAll(lblSearch, txtRecordFilter);
        VBox filterPanel = new VBox(4);
        filterPanel.getChildren().addAll(filterRowPrimary, filterModePanel);
        
        comboRecord = new ComboBox<>();
        comboRecord.setMaxWidth(Double.MAX_VALUE);
        comboRecord.setStyle("-fx-font-weight: bold; -fx-font-size: 13px;");
        comboRecord.setCellFactory(lv -> new ListCell<OrderRecord>() {
            @Override
            protected void updateItem(OrderRecord item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                } else {
                    setText(String.format("%s [%s] - %s", item.getReqNo(), item.getStatus(), item.getUser()));
                }
            }
        });
        comboRecord.setButtonCell(new ListCell<OrderRecord>() {
            @Override
            protected void updateItem(OrderRecord item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                } else {
                    setText(String.format("%s [%s] - %s", item.getReqNo(), item.getStatus(), item.getUser()));
                }
            }
        });
        
        comboRecord.getSelectionModel().selectedItemProperty().addListener((obs, oldSelection, newSelection) -> {
            if (newSelection != null) {
                loadRecordDetails(newSelection);
            }
        });
        
        leftContainer.getChildren().addAll(lblLeftTitle, filterPanel, comboRecord);
        
        // --- 1.1 BASIC INFO SECTION ---
        lblFormTitle = new Label("【基本情報・手修正フォーム】");
        lblFormTitle.setStyle("-fx-font-size: 13px; -fx-font-weight: bold;");
        lblFormTitle.getStyleClass().add("section-title-form");
        
        GridPane basicGrid = new GridPane();
        basicGrid.setHgap(6);
        basicGrid.setVgap(8);
        configureTwoColumnFormGrid(basicGrid);
        
        // Row 0
        Label lblReqNo = new Label("依頼Ｎｏ:");
        styleFormLabel(lblReqNo);
        txtReqNo = new TextField();
        txtReqNo.setStyle("-fx-font-size: 11px;");
        txtReqNo.setEditable(false);
        basicGrid.add(lblReqNo, 0, 0);
        addFormField(basicGrid, txtReqNo, 1, 0);
        
        Label lblUser = new Label("ユーザー:");
        styleFormLabel(lblUser);
        newCmbFormUser = new ComboBox<>(optUser);
        newCmbFormUser.setEditable(true);
        newCmbFormUser.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblUser, 2, 0);
        addFormField(basicGrid, newCmbFormUser, 3, 0);
        
        // Row 1
        Label lblDeliv = new Label("希望納期:");
        styleFormLabel(lblDeliv);
        newDpFormDeliv = new DatePicker();
        newDpFormDeliv.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblDeliv, 0, 1);
        addFormField(basicGrid, newDpFormDeliv, 1, 1);
        
        Label lblAdjustDeliv = new Label("調整納期:");
        styleFormLabel(lblAdjustDeliv);
        newDpFormAdjustDeliv = new DatePicker();
        newDpFormAdjustDeliv.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblAdjustDeliv, 2, 1);
        addFormField(basicGrid, newDpFormAdjustDeliv, 3, 1);
        
        // Row 2
        Label lblWage = new Label("加工賃:");
        styleFormLabel(lblWage);
        newTxtFormWage = new TextField();
        newTxtFormWage.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblWage, 0, 2);
        addFormField(basicGrid, newTxtFormWage, 1, 2);

        Label lblUketsukeNo = new Label("受付Ｎｏ:");
        styleFormLabel(lblUketsukeNo);
        newTxtUketsukeNo = new TextField();
        newTxtUketsukeNo.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblUketsukeNo, 2, 2);
        addFormField(basicGrid, newTxtUketsukeNo, 3, 2);

        // 依頼Ｎｏの実フィールドは txtReqNo。clearInputForm / addNewOrderToExcel が参照する
        // newTxtIraiNo は未生成のままだと NPE になるため生成だけしておく。
        if (newTxtIraiNo == null) {
            newTxtIraiNo = new TextField();
        }

        Label lblInputDate = new Label("入力日:");
        styleFormLabel(lblInputDate);
        newDpFormInputDate = new DatePicker();
        newDpFormInputDate.setStyle("-fx-font-size: 11px;");
        newDpFormInputDate.setEditable(false);
        newDpFormInputDate.setDisable(true);
        basicGrid.add(lblInputDate, 0, 3);
        addFormField(basicGrid, newDpFormInputDate, 1, 3);
        
        // Row 4
        Label lblProcess = new Label("加工内容:");
        styleFormLabel(lblProcess);
        txtProcess = new TextField();
        txtProcess.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblProcess, 2, 3);
        addFormField(basicGrid, txtProcess, 3, 3);
        
        basicGrid.setMaxWidth(Double.MAX_VALUE);
        
        leftContainer.getChildren().addAll(lblFormTitle, basicGrid);
        
        // --- 1.2 PRODUCT INFO SECTION ---
        HBox prodTitleBox = new HBox(10);
        prodTitleBox.setAlignment(Pos.CENTER_LEFT);
        Label lblProdTitle = new Label("【製品（仕上がり）情報】");
        lblProdTitle.setStyle("-fx-font-size: 13px; -fx-font-weight: bold;");
        lblProdTitle.getStyleClass().add("section-title-prod");
        Button btnAddProd = new Button("＋ 追加");
        btnAddProd.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-cursor: hand;");
        btnAddProd.getStyleClass().add("btn-settings-add");
        btnAddProd.setOnAction(e -> addProductRow(null));
        prodTitleBox.getChildren().addAll(lblProdTitle, btnAddProd);
        
        productRowsContainer = new VBox(10);
        productRowsContainer.setFillWidth(true);
        
        leftContainer.getChildren().addAll(prodTitleBox, productRowsContainer);
        
        // --- 1.3 RAW MATERIAL INFO SECTION ---
        HBox rawTitleBox = new HBox(10);
        rawTitleBox.setAlignment(Pos.CENTER_LEFT);
        Label lblRawTitle = new Label("【原反（材料）情報】");
        lblRawTitle.setStyle("-fx-font-size: 13px; -fx-font-weight: bold;");
        lblRawTitle.getStyleClass().add("section-title-raw");
        Button btnAddRaw = new Button("＋ 追加");
        btnAddRaw.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-cursor: hand;");
        btnAddRaw.getStyleClass().add("btn-settings-add");
        btnAddRaw.setOnAction(e -> addRawRow(null));
        rawTitleBox.getChildren().addAll(lblRawTitle, btnAddRaw);
        
        rawRowsContainer = new VBox(10);
        rawRowsContainer.setFillWidth(true);
        
        leftContainer.getChildren().addAll(rawTitleBox, rawRowsContainer);
        
        // --- 1.3.5 WORKFLOW DETAILS SECTION ---
        Label lblWorkflowTitle = new Label("【作業指示・特記事項】");
        lblWorkflowTitle.setStyle("-fx-font-size: 13px; -fx-font-weight: bold;");
        lblWorkflowTitle.getStyleClass().add("section-title-workflow");
        
        GridPane workGrid = new GridPane();
        workGrid.setHgap(8);
        workGrid.setVgap(6);
        configureTwoColumnFormGrid(workGrid);
        
        Label lblInputKbn = new Label("入力区分:");
        styleFormLabel(lblInputKbn);
        newCmbFormInputKbn = new ComboBox<>(optInputKbn);
        newCmbFormInputKbn.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblInputKbn, 0, 0);
        addFormField(workGrid, newCmbFormInputKbn, 1, 0);

        Label lblKakoKbn = new Label("加工区分:");
        styleFormLabel(lblKakoKbn);
        newCmbFormKakoKbn = new ComboBox<>(optKakoKbn);
        newCmbFormKakoKbn.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblKakoKbn, 2, 0);
        addFormField(workGrid, newCmbFormKakoKbn, 3, 0);

        Label lblInputTanto = new Label("入力担当:");
        styleFormLabel(lblInputTanto);
        lblFormInputTanto = new Label();
        lblFormInputTanto.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblInputTanto, 0, 1);
        addFormField(workGrid, lblFormInputTanto, 1, 1);
        refreshFormInputTantoLabel();

        Label lblYoto = new Label("用途:");
        styleFormLabel(lblYoto);
        newCmbFormYoto = new ComboBox<>(optYoto);
        newCmbFormYoto.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblYoto, 2, 1);
        addFormField(workGrid, newCmbFormYoto, 3, 1);
        
        Label lblTokki1 = new Label("特記事項1:");
        styleFormLabel(lblTokki1);
        newTxtFormTokki1 = new TextField();
        newTxtFormTokki1.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblTokki1, 0, 2);
        addFormField(workGrid, newTxtFormTokki1, 1, 2, 3, 1);
        
        Label lblTokki2 = new Label("特記事項2:");
        styleFormLabel(lblTokki2);
        newTxtFormTokki2 = new TextField();
        newTxtFormTokki2.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblTokki2, 0, 3);
        addFormField(workGrid, newTxtFormTokki2, 1, 3, 3, 1);
        
        Label lblTokki3 = new Label("特記事項3:");
        styleFormLabel(lblTokki3);
        newTxtFormTokki3 = new TextField();
        newTxtFormTokki3.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblTokki3, 0, 4);
        addFormField(workGrid, newTxtFormTokki3, 1, 4, 3, 1);
        
        leftContainer.getChildren().addAll(lblWorkflowTitle, workGrid);
        
        // --- 1.4 ACTION BUTTONS SECTION ---
        VBox btnContainer = new VBox(10);
        btnContainer.setPadding(new Insets(10, 0, 0, 0));
        btnContainer.setAlignment(Pos.CENTER);
        
        Button btnNewRecord = new Button("クリア / 新規追加");
        btnNewRecord.setMaxWidth(Double.MAX_VALUE);
        btnNewRecord.setStyle("-fx-font-weight: bold; -fx-cursor: hand; -fx-padding: 8px;");
        btnNewRecord.getStyleClass().add("btn-new-record");
        btnNewRecord.setOnAction(e -> {
            selectedRecord = null;
            if (comboRecord != null) {
                comboRecord.getSelectionModel().clearSelection();
            }
            clearOriginalSheetPreview();
            resetDiscrepancyLabelForNewEntry();
            txtReqNo.clear();
            txtReqNo.setEditable(true);
            txtReqNo.setPromptText("新規の依頼Noを入力");
            
            // txtUser.clear();
            // txtDelivery.clear();
            // txtContract.clear();
            txtProcess.clear();
            
            clearInputForm();
            
            lblFormTitle.setText("【基本情報・新規データ登録モード】");
            lblFormTitle.setStyle("-fx-font-size: 13px; -fx-font-weight: bold;");
            lblFormTitle.getStyleClass().removeAll("section-title-form", "section-title-new-mode");
            lblFormTitle.getStyleClass().add("section-title-new-mode");
        });
        
        btnSaveLocal = new Button("手修正を一時保存");
        configureSideActionButton(btnSaveLocal);
        btnSaveLocal.getStyleClass().add("btn-save-local");
        btnSaveLocal.setOnAction(e -> saveLocalForm());
        registerGuestMutableControl(btnSaveLocal);
        
        btnTransfer = new Button("受注ファイルへ自動転記・更新");
        configureSideActionButton(btnTransfer);
        btnTransfer.getStyleClass().add("btn-transfer");
        btnTransfer.setOnAction(e -> transferToExcel());
        registerGuestMutableControl(btnTransfer);

        btnUndoLastTransfer = new Button("直前の自動転記を取り消し");
        configureSideActionButton(btnUndoLastTransfer);
        btnUndoLastTransfer.getStyleClass().add("btn-undo-transfer");
        btnUndoLastTransfer.setDisable(true);
        btnUndoLastTransfer.setOnAction(e -> undoLastJuchuTransfer());
        registerGuestMutableControl(btnUndoLastTransfer);

        btnBulkTransferPending = new Button("一時保存分一括転記");
        configureSideActionButton(btnBulkTransferPending);
        btnBulkTransferPending.getStyleClass().add("btn-transfer");
        btnBulkTransferPending.setOnAction(e -> transferAllPendingLocalSaves());
        registerGuestMutableControl(btnBulkTransferPending);
        updateTransferButtonState();

        btnOpenJuchuSide = new Button("受注エクセルを開く");
        configureSideActionButton(btnOpenJuchuSide);
        btnOpenJuchuSide.getStyleClass().add("btn-reload");
        btnOpenJuchuSide.setOnAction(evt -> openJuchuExcelExternally());
        registerGuestMutableControl(btnOpenJuchuSide);

        btnJuchuColumnWizardSide = new Button("列定義ウィザード");
        configureSideActionButton(btnJuchuColumnWizardSide);
        btnJuchuColumnWizardSide.getStyleClass().add("btn-reload");
        btnJuchuColumnWizardSide.setOnAction(evt -> openJuchuColumnDefinitionWizard());
        registerGuestMutableControl(btnJuchuColumnWizardSide);

        FlowPane sideBtns =
                new FlowPane(
                        8,
                        8,
                        btnSaveLocal,
                        btnTransfer,
                        btnUndoLastTransfer,
                        btnBulkTransferPending,
                        btnOpenJuchuSide,
                        btnJuchuColumnWizardSide);
        sideBtns.getStyleClass().add("request-form-action-flow");
        sideBtns.setAlignment(javafx.geometry.Pos.CENTER);
        sideBtns.setColumnHalignment(javafx.geometry.HPos.CENTER);
        sideBtns.setMaxWidth(Double.MAX_VALUE);
        sideBtns.prefWrapLengthProperty().bind(btnContainer.widthProperty());

        Tooltip transferButtonsTooltip = new Tooltip();
        sideBtns.addEventFilter(
                javafx.scene.input.MouseEvent.MOUSE_ENTERED,
                e -> {
                    String reason = resolveTransferBlockedReason();
                    if (reason != null && btnTransfer != null && btnTransfer.isDisabled()) {
                        transferButtonsTooltip.setText(reason);
                        Tooltip.install(sideBtns, transferButtonsTooltip);
                    }
                });
        sideBtns.addEventFilter(
                javafx.scene.input.MouseEvent.MOUSE_EXITED,
                e -> Tooltip.uninstall(sideBtns, transferButtonsTooltip));

        transferBlockedReasonLabel = new Label();
        transferBlockedReasonLabel.getStyleClass().add("transfer-blocked-reason");
        transferBlockedReasonLabel.setWrapText(true);
        transferBlockedReasonLabel.setMaxWidth(Double.MAX_VALUE);
        transferBlockedReasonLabel.setManaged(false);
        transferBlockedReasonLabel.setVisible(false);
        
        btnContainer.getChildren().addAll(btnNewRecord, sideBtns, transferBlockedReasonLabel);
        leftContainer.getChildren().add(btnContainer);
        
        // 2. RIGHT PANE: Visual Sheet Viewer & Discrepancies
        VBox rightPane = new VBox(10);
        rightPane.getStyleClass().add("pane-container");
        
        HBox viewerHeaderBox = new HBox(15);
        viewerHeaderBox.setAlignment(Pos.CENTER_LEFT);
        Label lblViewer = new Label("原本加工依頼書プレビュー (PDFキャッシュ)");
        lblViewer.getStyleClass().add("pane-title-right");
        previewFileNameLabel = new Label();
        previewFileNameLabel.getStyleClass().add("top-status");
        previewFileUpdateBadgeHost = new StackPane();
        previewFileUpdateBadgeHost.setManaged(false);
        previewFileUpdateBadgeHost.setVisible(false);
        previewFileNameRow = new HBox(8, previewFileNameLabel, previewFileUpdateBadgeHost);
        previewFileNameRow.setAlignment(Pos.CENTER_LEFT);
        previewFileNameRow.setManaged(false);
        previewFileNameRow.setVisible(false);
        Button btnOpenExcel = new Button("原本をExcelで開く");
        btnOpenExcel.setOnAction(e -> openOriginalExcel());
        Button btnClearPreviewCache = new Button("プレビューキャッシュをクリア");
        btnClearPreviewCache.getStyleClass().add("btn-clear");
        btnClearPreviewCache.setOnAction(e -> confirmAndClearPreviewCache());
        viewerHeaderBox
                .getChildren()
                .addAll(lblViewer, previewFileNameRow, btnOpenExcel, btnClearPreviewCache);
        HBox.setHgrow(previewFileNameRow, Priority.ALWAYS);
        sheetScrollPane = new ScrollPane();
        sheetScrollPane.setFitToWidth(true);
        sheetScrollPane.setFitToHeight(false);
        sheetScrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        sheetScrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        sheetScrollPane.getStyleClass().add("request-form-sheet-scroll");
        VBox.setVgrow(sheetScrollPane, Priority.ALWAYS);
        
        sheetGrid = new GridPane();
        sheetGrid.setPadding(new Insets(10));
        sheetGrid.setHgap(1);
        sheetGrid.setVgap(1);
        sheetScrollPane.setContent(sheetGrid);
        
        discrepancyLabel = new Label("原本と受注ファイルのデータは一致しています。");
        discrepancyLabel.setWrapText(true);
        discrepancyLabel.getStyleClass().add("discrepancy-label-info");
        discrepancyLabel.setMaxWidth(Double.MAX_VALUE);
        
        rightPane.getChildren().addAll(viewerHeaderBox, sheetScrollPane, discrepancyLabel);

        splitPane.getItems().addAll(leftScroll, rightPane);
        splitPane.setDividerPositions(SPLIT_LEFT_RATIO);
        SplitPane.setResizableWithParent(leftScroll, Boolean.TRUE);
        SplitPane.setResizableWithParent(rightPane, Boolean.TRUE);

        // --- DYNAMIC TABS STRUCTURE ---
        TabPane tabPane = new TabPane();
        
        // Tab 1: Database Verification view
        Tab tabVerification = new Tab("一括照合データベース・受注管理");
        tabVerification.setClosable(false);
        BorderPane verificationLayout = new BorderPane();
        verificationLayout.setCenter(splitPane);
        tabVerification.setContent(verificationLayout);
        
        // Tab 2: Settings - edit ComboBox option lists
        Tab tabSettings = createSettingsTab();
        // Tab 3: Post-processing product master editor (right of settings)
        Tab tabPostProcMaster = createPostProcessingProductMasterTab();
        Tab tabPlanMachineDb = createPlanMachineCatalogTab();

        tabPane.getTabs().addAll(tabVerification, tabSettings, tabPostProcMaster, tabPlanMachineDb);
        root.setCenter(tabPane);

        mainStackPane = new StackPane();
        mainStackPane.getChildren().add(root);
        
        // Build premium loading overlay
        loadingOverlay = new VBox(20);
        loadingOverlay.setAlignment(Pos.CENTER);
        loadingOverlay.getStyleClass().add("request-form-loading-overlay");

        ProgressIndicator progressIndicator = new ProgressIndicator();
        progressIndicator.setStyle("-fx-progress-color: -fx-accent;");
        progressIndicator.setPrefSize(60, 60);

        loadingOverlayLabel = new Label("データを読み込んでいます...");
        loadingOverlayLabel.getStyleClass().add("request-form-loading-label");
        
        loadingOverlay.getChildren().addAll(progressIndicator, loadingOverlayLabel);
        loadingOverlay.setVisible(false);
        
        mainStackPane.getChildren().add(loadingOverlay);

        // Track user interaction to pause/resume background caching
        mainStackPane.addEventFilter(javafx.scene.input.InputEvent.ANY, event -> {
            lastInteractionTime = System.currentTimeMillis();
        });

        applyMainShellAlignedStyles(mainStackPane);
        originalUpdateMonitor.setOnUpdatedKeysChanged(
                keys -> Platform.runLater(this::refreshPreviewFileHeader));

        applyGuestSessionRestrictions();
        loadMasterProductListAsync(null);
        reloadData();
        return mainStackPane;
    }

    /** 依頼書入力タブ選択時: 環境変数反映と、変更があったときだけマスタを非同期再読込。 */
    public void onEmbeddedTabActivated(Map<String, String> uiEnv) {
        configureFromUiEnv(uiEnv);
        reloadComboChoicesFromSummarySettings();
        reloadMasterProductListFromDiskIfStale(null);
        startEmbeddedTabPolling();
    }

    /** 依頼書入力タブを離れたとき: バックグラウンド監視を停止。 */
    public void onEmbeddedTabDeactivated() {
        stopEmbeddedTabPolling();
    }

    public void updateHostWindow(Window hostWindow) {
        this.hostWindow = hostWindow;
    }

    private void setupTableColumns() {
    }

    private TextField addFormField(GridPane grid, String labelText, int col, int row) {
        Label label = new Label(labelText);
        label.setStyle("-fx-font-weight: bold;");
        TextField textField = new TextField();
        textField.setPrefWidth(160);
        
        grid.add(label, col * 2, row);
        grid.add(textField, col * 2 + 1, row);
        return textField;
    }

    // --- TAB 2: MOCKUP INPUT FORM TAB LAYOUT CREATION (100% STYLE OFF-LOADED TO CSS) ---
    // createInputFormTab removed

    private ComboBox<String> addPaperMetaCombo(GridPane grid, String labelText, ObservableList<String> options, int row) {
        Label lblKey = new Label(labelText);
        lblKey.getStyleClass().add("paper-meta-label");
        ComboBox<String> combo = new ComboBox<>(options);
        combo.setEditable(true);
        combo.getStyleClass().add("paper-meta-combo");
        if (!options.isEmpty()) combo.setValue(options.get(0));
        grid.add(lblKey, 0, row);
        grid.add(combo, 1, row);
        return combo;
    }

    private TextField addPaperMetaInput(GridPane grid, String labelText, int row) {
        Label lblKey = new Label(labelText);
        lblKey.getStyleClass().add("paper-meta-label");
        TextField tf = new TextField();
        tf.getStyleClass().add("paper-meta-input");
        grid.add(lblKey, 0, row);
        grid.add(tf, 1, row);
        return tf;
    }

    private TextField createFormCell(GridPane grid, int col, int row, String defaultText) {
        TextField tf = new TextField(defaultText);
        tf.getStyleClass().add("paper-input-cell");
        grid.add(tf, col, row);
        return tf;
    }

    private ComboBox<String> createFormCombo(GridPane grid, int col, int row, ObservableList<String> options) {
        ComboBox<String> combo = new ComboBox<>(options);
        combo.setEditable(true);
        combo.getStyleClass().add("paper-combo-cell");
        combo.setMaxWidth(Double.MAX_VALUE);
        if (!options.isEmpty()) combo.setValue(options.get(0));
        grid.add(combo, col, row);
        return combo;
    }

    private ComboBox<String> addFormVerticalCombo(GridPane grid, String labelText, int row, ObservableList<String> options) {
        Label label = new Label(labelText);
        label.getStyleClass().add("paper-vertical-label");
        ComboBox<String> combo = new ComboBox<>(options);
        combo.setEditable(true);
        combo.getStyleClass().add("paper-vertical-combo");
        combo.setMaxWidth(Double.MAX_VALUE);
        if (!options.isEmpty()) combo.setValue(options.get(0));
        grid.add(label, 0, row);
        grid.add(combo, 1, row);
        return combo;
    }

    private void addLeadTimeRow(GridPane grid, String labelText, String dateText, int row) {
        Label label = new Label(labelText);
        label.getStyleClass().add("paper-lead-label");
        Label dateLabel = new Label(dateText);
        dateLabel.getStyleClass().add("paper-lead-value");
        grid.add(label, 0, row);
        grid.add(dateLabel, 1, row);
    }

    private String calculateBusinessDate(int days) {
        Calendar cal = Calendar.getInstance();
        int added = 0;
        while (added < days) {
            cal.add(Calendar.DAY_OF_YEAR, 1);
            int dayOfWeek = cal.get(Calendar.DAY_OF_WEEK);
            if (dayOfWeek != Calendar.SATURDAY && dayOfWeek != Calendar.SUNDAY) {
                added++;
            }
        }
        SimpleDateFormat sdf = new SimpleDateFormat("M/d");
        return sdf.format(cal.getTime());
    }

    private TextField addFormVerticalField(GridPane grid, String labelText, int row, String defaultVal) {
        Label label = new Label(labelText);
        label.getStyleClass().add("paper-vertical-label");
        TextField textField = new TextField(defaultVal);
        textField.getStyleClass().add("paper-vertical-input");
        grid.add(label, 0, row);
        grid.add(textField, 1, row);
        return textField;
    }

    private static final DateTimeFormatter JUCHU_BACKUP_LIST_TS =
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss").withZone(ZoneId.systemDefault());

    // --- TAB 3: SETTINGS (Edit ComboBox option lists) ---
    private Tab createSettingsTab() {
        Tab tab = new Tab("【設定】");
        tab.setClosable(false);
        tab.setOnSelectionChanged(
                evt -> {
                    if (tab.isSelected()) {
                        refreshJuchuBackupList();
                    }
                });

        ScrollPane sp = new ScrollPane();
        sp.setFitToWidth(true);
        sp.getStyleClass().add("form-scroll-pane");

        VBox root = new VBox(12);
        root.getStyleClass().add("form-tab-container");
        root.setFillWidth(false);
        root.setAlignment(Pos.TOP_LEFT);

        Label title = new Label("\u30b3\u30f3\u30dc\u30dc\u30c3\u30af\u30b9\u9805\u76ee\u7de8\u96c6 / Settings - ComboBox \u5019\u88dc\u5024\u7ba1\u7406");
        title.getStyleClass().add("paper-main-title");
        Label subtitle = new Label("\u5404\u30b3\u30f3\u30dc\u30dc\u30c3\u30af\u30b9\u306e\u9078\u629e\u80a2\u3092\u81ea\u7531\u306b\u8ffd\u52a0\u30fb\u524a\u9664\u30fb\u7de8\u96c6\u3067\u304d\u307e\u3059\u3002");
        subtitle.getStyleClass().add("paper-main-subtitle");
        subtitle.setWrapText(true);
        subtitle.setMaxWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        root.getChildren().addAll(title, subtitle);

        Label prefixTitle = new Label("マスタ候補コンボ 先頭文字フィルタ");
        prefixTitle.getStyleClass().add("paper-main-title");
        Label prefixSubtitle =
                new Label(
                        "品番横のマスタ候補コンボに表示する商品コードの先頭文字列を指定します。"
                                + " 複数指定可（いずれかに一致するもののみ表示）。空ならフィルタなし。"
                                + " 製品側と原反側で別々に設定できます。"
                                + " 後加工商品マスタタブの参照検索には製品側・原反側をそれぞれ適用します"
                                + "（片方だけ設定時は、もう片方は無制限。両方設定時はいずれかの先頭一致）。"
                                + " 保存先: サマリ Excel と同じフォルダの "
                                + AppPaths.REQUEST_FORM_INPUT_SETTINGS_JSON_FILENAME
                                + "（追加・削除のたびに自動保存）。");
        prefixSubtitle.getStyleClass().add("paper-main-subtitle");
        prefixSubtitle.setWrapText(true);
        prefixSubtitle.setMaxWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        GridPane prefixEditorsGrid = new GridPane();
        prefixEditorsGrid.setHgap(12);
        prefixEditorsGrid.setVgap(12);
        prefixEditorsGrid.setAlignment(Pos.TOP_LEFT);
        prefixEditorsGrid.add(
                buildListEditor(
                        "製品側（商品コード先頭）",
                        optMasterCandidatePrefixProduct,
                        this::refreshAllMasterCandidateCombos),
                0,
                0);
        prefixEditorsGrid.add(
                buildListEditor(
                        "原反側（商品コード先頭）",
                        optMasterCandidatePrefixRaw,
                        this::refreshAllMasterCandidateCombos),
                1,
                0);
        root.getChildren().addAll(prefixTitle, prefixSubtitle, prefixEditorsGrid);

        GridPane editorsGrid = new GridPane();
        editorsGrid.setHgap(12);
        editorsGrid.setVgap(12);
        editorsGrid.setAlignment(Pos.TOP_LEFT);
        VBox[] editors =
                new VBox[] {
                    buildListEditor("\u5165\u529b\u533a\u5206", optInputKbn),
                    buildListEditor("\u52a0\u5de5\u533a\u5206", optKakoKbn),
                    buildListEditor("\u5272\u6570", optWariSu),
                    buildListEditor("\uff25\uff23\u9762", optEcSide),
                    buildListEditor("\uff84\uff98\uff90\uff9d\uff78\uff9e", optTrimming),
                    buildListEditor("\u6295\u5165\u5834\u6240", optFeedLoc),
                    buildListEditor("\u5728\u5eab\u5834\u6240", optStorageLoc),
                    buildListEditor("\u7528\u9014", optYoto),
                    buildListEditor("\u30e6\u30fc\u30b6\u30fc", optUser)
                };
        for (int i = 0; i < editors.length; i++) {
            editorsGrid.add(editors[i], i % 2, i / 2);
        }
        root.getChildren().add(editorsGrid);

        Label defaultsTitle = new Label("入力欄の既定値（【作業指示・特記事項】）");
        defaultsTitle.getStyleClass().add("paper-main-title");
        Label defaultsSubtitle =
                new Label(
                        "新規追加・クリア時、および受注ファイルへ新規転記する行に適用する"
                                + "「入力区分」「加工区分」の初期選択です。");
        defaultsSubtitle.getStyleClass().add("paper-main-subtitle");
        defaultsSubtitle.setWrapText(true);
        defaultsSubtitle.setMaxWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        root.getChildren().addAll(defaultsTitle, defaultsSubtitle, buildFieldDefaultsSettingsCard());

        // --- 受注ファイル設定カードの追加 ---
        VBox juchuCard = new VBox(10);
        juchuCard.getStyleClass().add("settings-card");
        juchuCard.setMaxWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        juchuCard.setPrefWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        
        Label lblJuchuCardTitle = new Label("受注ファイル（加工計画等）設定");
        lblJuchuCardTitle.getStyleClass().add("settings-card-title");
        
        Label lblJuchuDesc = new Label("自動転記先となる受注ファイル（例：加工計画.xlsm、加工依頼書入力.xlsmなど）を指定し、直接Excelで開くことができます。");
        lblJuchuDesc.setStyle("-fx-font-size: 11px; -fx-text-fill: inherit;");
        lblJuchuDesc.setWrapText(true);
        
        txtJuchuPathDisplay = new TextField(juchuFilePath != null ? juchuFilePath : "");
        txtJuchuPathDisplay.setEditable(false);
        txtJuchuPathDisplay.setStyle("-fx-font-size: 11px;");
        
        Button btnSelectJuchuFile = new Button("ファイル選択");
        btnSelectJuchuFile.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnSelectJuchuFile.getStyleClass().add("btn-reload");
        registerGuestMutableControl(btnSelectJuchuFile);
        
        Button btnOpenJuchuFile = new Button("Excelで開く");
        btnOpenJuchuFile.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnOpenJuchuFile.getStyleClass().add("btn-transfer");
        registerGuestMutableControl(btnOpenJuchuFile);
        
        btnSelectJuchuFile.setOnAction(evt -> {
            javafx.stage.FileChooser fileChooser = new javafx.stage.FileChooser();
            fileChooser.setTitle("受注ファイルを選択");
            fileChooser.getExtensionFilters().addAll(
                new javafx.stage.FileChooser.ExtensionFilter("Excel Files (*.xlsm, *.xlsx)", "*.xlsm", "*.xlsx"),
                new javafx.stage.FileChooser.ExtensionFilter("All Files", "*.*")
            );
            File currentFile = new File(juchuFilePath);
            if (currentFile.exists()) {
                fileChooser.setInitialDirectory(currentFile.getParentFile());
            } else {
                File workspaceDir = new File(targetFolder);
                if (workspaceDir.exists()) fileChooser.setInitialDirectory(workspaceDir);
            }
            File chosenFile = fileChooser.showOpenDialog(hostWindow);
            if (chosenFile != null) {
                applySelectedJuchuFile(chosenFile.getAbsolutePath());
                saveSettings();
                statusLabel.setText("受注ファイルを更新しました: " + juchuFilePath);
                updateTransferButtonState();
            }
        });
        
        btnOpenJuchuFile.setOnAction(evt -> openJuchuExcelExternally());

        Button btnOpenJuchuBackupFolder = new Button("バックアップフォルダを開く");
        btnOpenJuchuBackupFolder.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnOpenJuchuBackupFolder.getStyleClass().add("btn-reload");
        registerGuestMutableControl(btnOpenJuchuBackupFolder);
        btnOpenJuchuBackupFolder.setTooltip(
                new Tooltip(
                        "受注ファイル書き込み前のローカル世代バックアップ（15 分間隔）の保存先フォルダを開きます。"));
        btnOpenJuchuBackupFolder.setOnAction(evt -> openJuchuBackupFolderExternally());

        Button btnJuchuColumnWizard = new Button("列定義ウィザード");
        btnJuchuColumnWizard.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnJuchuColumnWizard.getStyleClass().add("btn-reload");
        btnJuchuColumnWizard.setOnAction(evt -> openJuchuColumnDefinitionWizard());
        registerGuestMutableControl(btnJuchuColumnWizard);

        HBox juchuBtnBox =
                new HBox(10, btnSelectJuchuFile, btnOpenJuchuFile, btnJuchuColumnWizard);
        juchuBtnBox.setAlignment(Pos.CENTER_LEFT);

        Label lblJuchuBackupTitle = new Label("受注ファイル ローカル世代バックアップ");
        lblJuchuBackupTitle.getStyleClass().add("settings-card-subtitle");
        lblJuchuBackupTitle.setStyle("-fx-font-size: 11px; -fx-font-weight: bold;");

        juchuBackupHintLabel = new Label();
        juchuBackupHintLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: inherit;");
        juchuBackupHintLabel.setWrapText(true);

        juchuBackupListView =
                new ListView<>(
                        FXCollections.observableArrayList(
                                List.<RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry>of()));
        juchuBackupListView.setPrefHeight(140);
        juchuBackupListView.setMaxHeight(180);
        juchuBackupListView.setStyle("-fx-font-size: 11px;");
        juchuBackupListView.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(
                                    RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry item,
                                    boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    return;
                                }
                                String ts =
                                        item.createdAtMillis() > 0L
                                                ? JUCHU_BACKUP_LIST_TS.format(
                                                        Instant.ofEpochMilli(item.createdAtMillis()))
                                                : "";
                                String reason =
                                        RequestFormJuchuFileBackupStore.reasonLabelJa(item.reason());
                                setText(ts + "  " + reason + "  " + item.displayLabel());
                            }
                        });

        Button btnRefreshJuchuBackup = new Button("一覧更新");
        btnRefreshJuchuBackup.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnRefreshJuchuBackup.getStyleClass().add("btn-reload");
        btnRefreshJuchuBackup.setOnAction(evt -> refreshJuchuBackupList());

        Button btnRestoreJuchuBackup = new Button("選択した世代を復元");
        btnRestoreJuchuBackup.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnRestoreJuchuBackup.getStyleClass().add("btn-transfer");
        registerGuestMutableControl(btnRestoreJuchuBackup);
        btnRestoreJuchuBackup.setTooltip(
                new Tooltip(
                        "選択したバックアップで現在の受注ファイルを上書きします（復元前に現行ファイルを退避）。"));
        btnRestoreJuchuBackup.setOnAction(evt -> restoreSelectedJuchuBackupFromList());

        HBox juchuBackupBtnBox =
                new HBox(10, btnRefreshJuchuBackup, btnRestoreJuchuBackup, btnOpenJuchuBackupFolder);
        juchuBackupBtnBox.setAlignment(Pos.CENTER_LEFT);

        juchuCard.getChildren()
                .addAll(
                        lblJuchuCardTitle,
                        lblJuchuDesc,
                        txtJuchuPathDisplay,
                        juchuBtnBox,
                        lblJuchuBackupTitle,
                        juchuBackupHintLabel,
                        juchuBackupListView,
                        juchuBackupBtnBox);
        root.getChildren().add(juchuCard);
        refreshJuchuBackupList();

        // --- アラジンマスタ一括統合ツールカードの追加 ---
        VBox toolCard = new VBox(10);
        toolCard.getStyleClass().add("settings-card");
        toolCard.setMaxWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        toolCard.setPrefWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        
        Label lblToolTitle = new Label("アラジンマスタ一括統合・リレーション生成");
        lblToolTitle.getStyleClass().add("settings-card-title");
        
        Label lblToolDesc = new Label("「後加工商品マスタ」「後加工加工内容マスタ」「後加工工程マスタ」の3ファイルを統合し、リレーションを組んだExcelファイル（マスタリレーション統合結果.xlsx）を生成・更新します。");
        lblToolDesc.setStyle("-fx-font-size: 11px; -fx-text-fill: inherit;");
        lblToolDesc.setWrapText(true);
        
        Button btnRunTool = new Button("マスタリレーション統合結果を生成");
        btnRunTool.setMaxWidth(Region.USE_PREF_SIZE);
        btnRunTool.getStyleClass().add("btn-transfer"); // 緑系のボタンスタイルを適用
        registerGuestMutableControl(btnRunTool);
        
        Label lblToolStatus = new Label("待機中...");
        lblToolStatus.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: inherit;");
        
        btnRunTool.setOnAction(
                e -> runIntegratedMasterGeneration(lblToolStatus, btnRunTool, true));
        
        toolCard.getChildren().addAll(lblToolTitle, lblToolDesc, btnRunTool, lblToolStatus);
        root.getChildren().add(toolCard);

        sp.setContent(root);
        tab.setContent(sp);
        return tab;
    }

    private Tab createPostProcessingProductMasterTab() {
        Tab tab = new Tab("後加工商品マスタ");
        tab.setClosable(false);

        StackPane lazyHost = new StackPane();
        lazyHost.setAlignment(Pos.CENTER);
        Label hint = new Label("このタブを開くとマスタ編集画面を読み込みます");
        hint.getStyleClass().add("request-form-tab-loading-label");
        lazyHost.getChildren().add(hint);
        tab.setContent(lazyHost);

        java.util.concurrent.atomic.AtomicBoolean contentMounted =
                new java.util.concurrent.atomic.AtomicBoolean(false);
        tab.selectedProperty()
                .addListener(
                        (obs, wasSelected, selected) -> {
                            if (!selected || !contentMounted.compareAndSet(false, true)) {
                                return;
                            }
                            scheduleLazyTabMount(
                                    lazyHost,
                                    "後加工商品マスタ",
                                    () -> {
                                        ScrollPane sp = new ScrollPane();
                                        sp.setFitToWidth(true);
                                        sp.getStyleClass().add("form-scroll-pane");
                                        VBox content =
                                                PostProcessingProductMasterEditorPane
                                                        .buildTabContent(
                                                                hostWindow,
                                                                new PostProcessingProductMasterEditorPane
                                                                        .Context(
                                                                        () -> uiEnvSnapshot,
                                                                        this
                                                                                ::snapshotMasterProductCatalog,
                                                                        this
                                                                                ::snapshotMasterCandidatePrefixFilters,
                                                                        () -> {
                                                                            PostProcessingProductMasterReferenceCache
                                                                                    .invalidate();
                                                                            PostProcessingKouteiNaiyoMasterLookup
                                                                                    .invalidate();
                                                                            PostProcessingShuruiMasterLookup
                                                                                    .invalidate();
                                                                            PostProcessingKeiriBunruiMasterLookup
                                                                                    .invalidate();
                                                                            PostProcessingYotoMasterLookup
                                                                                    .invalidate();
                                                                            PostProcessingShohinBunrui4MasterLookup
                                                                                    .invalidate();
                                                                            PostProcessingZaikoBunruiMasterLookup
                                                                                    .invalidate();
                                                                            PostProcessingPlanMachineLookup
                                                                                    .invalidate();
                                                                        },
                                                                        msg ->
                                                                                System.out
                                                                                        .println(
                                                                                                msg)));
                                        sp.setContent(content);
                                        lazyHost.getChildren().setAll(sp);
                                    });
                        });
        return tab;
    }

    private Tab createPlanMachineCatalogTab() {
        Tab tab = new Tab("機械コード（加工計画）");
        tab.setClosable(false);

        StackPane lazyHost = new StackPane();
        lazyHost.setAlignment(Pos.CENTER);
        Label hint = new Label("このタブを開くと機械コード一覧を読み込みます");
        hint.getStyleClass().add("request-form-tab-loading-label");
        lazyHost.getChildren().add(hint);
        tab.setContent(lazyHost);

        java.util.concurrent.atomic.AtomicBoolean contentMounted =
                new java.util.concurrent.atomic.AtomicBoolean(false);
        tab.selectedProperty()
                .addListener(
                        (obs, wasSelected, selected) -> {
                            if (!selected || !contentMounted.compareAndSet(false, true)) {
                                return;
                            }
                            scheduleLazyTabMount(
                                    lazyHost,
                                    "機械コード（加工計画）",
                                    () -> {
                                        ScrollPane sp = new ScrollPane();
                                        sp.setFitToWidth(true);
                                        sp.getStyleClass().add("form-scroll-pane");
                                        VBox content =
                                                PostProcessingPlanMachineCatalogPane.build(
                                                        () -> uiEnvSnapshot,
                                                        msg -> System.out.println(msg));
                                        sp.setContent(content);
                                        VBox.setVgrow(
                                                content, javafx.scene.layout.Priority.ALWAYS);
                                        lazyHost.getChildren().setAll(sp);
                                    });
                        });
        return tab;
    }

    /**
     * 遅延タブ: スピナー表示 → バックグラウンドで class preload → FX スレッドでマウント。
     * Windows 増分 compile で inner class が欠けると {@link NoClassDefFoundError} になり得る。
     */
    private void scheduleLazyTabMount(StackPane lazyHost, String tabTitle, Runnable mount) {
        lazyHost.getChildren().setAll(new ProgressIndicator());
        Thread preload =
                new Thread(
                        () -> {
                            String preloadError = preloadLazyTabClasses(tabTitle);
                            Platform.runLater(
                                    () -> {
                                        if (preloadError != null) {
                                            showLazyTabMountError(
                                                    lazyHost, tabTitle, preloadError);
                                            return;
                                        }
                                        mountLazyTabContent(lazyHost, tabTitle, mount);
                                    });
                        },
                        "pm-ai-lazy-tab-" + tabTitle);
        preload.setDaemon(true);
        preload.start();
    }

    private static String preloadLazyTabClasses(String tabTitle) {
        try {
            if ("後加工商品マスタ".equals(tabTitle)) {
                Class.forName(
                        "jp.co.pm.ai.desktop.reconciliation.PostProcessingProductMasterEditorPane");
                Class.forName("jp.co.pm.ai.desktop.reconciliation.ReconciliationApp$ProductRow");
            } else if ("機械コード（加工計画）".equals(tabTitle)) {
                Class.forName(
                        "jp.co.pm.ai.desktop.reconciliation.PostProcessingPlanMachineCatalogPane");
            }
            return null;
        } catch (Throwable ex) {
            return ex.getMessage() != null ? ex.getMessage() : ex.toString();
        }
    }

    private static void showLazyTabMountError(
            StackPane host, String tabTitle, String detail) {
        System.err.println(
                "[request-form] lazy tab preload failed (" + tabTitle + "): " + detail);
        Label err =
                new Label(
                        tabTitle
                                + " の読み込みに失敗しました。\n"
                                + detail
                                + "\n\n対処: code_java で .\\mvnw.cmd clean compile を実行してから再起動してください。");
        err.setWrapText(true);
        err.setMaxWidth(560);
        err.getStyleClass().add("request-form-tab-loading-label");
        host.getChildren().setAll(err);
    }

    /**
     * 遅延タブのマウント失敗時にスピナーが残らないよう、エラーパネルへ切り替える。
     */
    private static void mountLazyTabContent(StackPane host, String tabTitle, Runnable mount) {
        try {
            mount.run();
        } catch (Throwable ex) {
            String detail = ex.getMessage() != null ? ex.getMessage() : ex.toString();
            System.err.println(
                    "[request-form] lazy tab mount failed (" + tabTitle + "): " + detail);
            ex.printStackTrace();
            Label err =
                    new Label(
                            tabTitle
                                    + " の読み込みに失敗しました。\n"
                                    + detail
                                    + "\n\n対処: code_java で .\\mvnw.cmd clean compile を実行してから再起動してください。");
            err.setWrapText(true);
            err.setMaxWidth(560);
            err.getStyleClass().add("request-form-tab-loading-label");
            host.getChildren().setAll(err);
        }
    }

    /**
     * Build a compact editor card for one ObservableList<String>.
     * Changes here update the ComboBoxes live (shared reference).
     */
    private VBox buildListEditor(String fieldName, ObservableList<String> items) {
        return buildListEditor(fieldName, items, null);
    }

    private VBox buildListEditor(String fieldName, ObservableList<String> items, Runnable afterChange) {
        VBox card = new VBox(6);
        card.getStyleClass().add("settings-card");
        card.setMinWidth(SETTINGS_CARD_WIDTH);
        card.setPrefWidth(SETTINGS_CARD_WIDTH);
        card.setMaxWidth(SETTINGS_CARD_WIDTH);

        Label lbl = new Label(fieldName);
        lbl.getStyleClass().add("settings-card-title");

        ListView<String> listView = new ListView<>(items);
        listView.getStyleClass().add("settings-list");
        listView.setPrefHeight(Math.min(Math.max(items.size(), 2) * 26 + 2, 110));
        listView.setMaxHeight(110);
        listView.setEditable(false);

        HBox addRow = new HBox(6);
        addRow.setAlignment(Pos.CENTER_LEFT);
        TextField tfNew = new TextField();
        tfNew.setPromptText("\u65b0\u3057\u3044\u5024\u3092\u5165\u529b...");
        tfNew.getStyleClass().add("settings-new-input");
        tfNew.setPrefWidth(120);
        HBox.setHgrow(tfNew, Priority.ALWAYS);
        tfNew.setMaxWidth(Double.MAX_VALUE);

        Button btnAdd = new Button("+ \u8ffd\u52a0");
        btnAdd.getStyleClass().add("btn-settings-add");
        btnAdd.setMinWidth(Region.USE_PREF_SIZE);
        registerGuestMutableControl(tfNew);
        registerGuestMutableControl(btnAdd);
        btnAdd.setOnAction(
                e -> {
                    String v = tfNew.getText().trim();
                    if (!v.isEmpty() && !items.contains(v)) {
                        items.add(v);
                        listView.setPrefHeight(Math.min(Math.max(items.size(), 2) * 26 + 2, 140));
                        tfNew.clear();
                        saveSettings();
                        if (afterChange != null) {
                            afterChange.run();
                        }
                    }
                });
        tfNew.setOnAction(btnAdd.getOnAction());

        Button btnDel = new Button("\u524a\u9664");
        btnDel.getStyleClass().add("btn-settings-del");
        btnDel.setMinWidth(Region.USE_PREF_SIZE);
        registerGuestMutableControl(btnDel);
        btnDel.setOnAction(
                e -> {
                    String sel = listView.getSelectionModel().getSelectedItem();
                    if (sel != null) {
                        items.remove(sel);
                        listView.setPrefHeight(Math.min(Math.max(items.size(), 2) * 26 + 2, 140));
                        saveSettings();
                        if (afterChange != null) {
                            afterChange.run();
                        }
                    }
                });

        addRow.getChildren().addAll(tfNew, btnAdd, btnDel);
        card.getChildren().addAll(lbl, listView, addRow);
        return card;
    }

    private VBox buildFieldDefaultsSettingsCard() {
        VBox card = new VBox(10);
        card.getStyleClass().add("settings-card");
        card.setMaxWidth(SETTINGS_CARD_WIDTH * 2 + 12);
        card.setPrefWidth(SETTINGS_CARD_WIDTH * 2 + 12);

        GridPane grid = new GridPane();
        grid.setHgap(12);
        grid.setVgap(8);

        Label lblInput = new Label("入力区分（既定）:");
        lblInput.getStyleClass().add("settings-card-title");
        cmbSettingsDefaultInputKbn = new ComboBox<>(optInputKbn);
        cmbSettingsDefaultInputKbn.setMaxWidth(Double.MAX_VALUE);
        registerGuestMutableControl(cmbSettingsDefaultInputKbn);
        cmbSettingsDefaultInputKbn
                .valueProperty()
                .addListener(
                        (obs, oldVal, newVal) -> {
                            if (newVal != null && !newVal.isBlank()) {
                                updateFieldDefaultInState(
                                        RequestFormComboChoices.KEY_INPUT_KBN, newVal);
                            }
                        });

        Label lblKako = new Label("加工区分（既定）:");
        lblKako.getStyleClass().add("settings-card-title");
        cmbSettingsDefaultKakoKbn = new ComboBox<>(optKakoKbn);
        cmbSettingsDefaultKakoKbn.setMaxWidth(Double.MAX_VALUE);
        registerGuestMutableControl(cmbSettingsDefaultKakoKbn);
        cmbSettingsDefaultKakoKbn
                .valueProperty()
                .addListener(
                        (obs, oldVal, newVal) -> {
                            if (newVal != null && !newVal.isBlank()) {
                                updateFieldDefaultInState(
                                        RequestFormComboChoices.KEY_KAKO_KBN, newVal);
                            }
                        });

        grid.add(lblInput, 0, 0);
        grid.add(cmbSettingsDefaultInputKbn, 1, 0);
        grid.add(lblKako, 0, 1);
        grid.add(cmbSettingsDefaultKakoKbn, 1, 1);
        GridPane.setHgrow(cmbSettingsDefaultInputKbn, Priority.ALWAYS);
        GridPane.setHgrow(cmbSettingsDefaultKakoKbn, Priority.ALWAYS);
        card.getChildren().add(grid);
        syncFieldDefaultSelectorCombos();
        return card;
    }

    private void updateFieldDefaultInState(String key, String value) {
        if (key == null || value == null || value.isBlank()) {
            return;
        }
        LinkedHashMap<String, String> nextDefaults =
                new LinkedHashMap<>(comboChoicesState.fieldDefaultsAsMap());
        nextDefaults.put(key, value.strip());
        comboChoicesState =
                RequestFormComboChoices.of(comboChoicesState.asMap(), nextDefaults);
        saveSettings();
    }

    /** サマリ Excel 同フォルダへ依頼書入力設定を書き出す（ComboBox 候補・パス）。 */
    public void persistInputSettings() {
        saveSettings();
    }

    private void syncFieldDefaultSelectorCombos() {
        if (cmbSettingsDefaultInputKbn != null) {
            cmbSettingsDefaultInputKbn.setValue(
                    comboChoicesState.effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
        }
        if (cmbSettingsDefaultKakoKbn != null) {
            cmbSettingsDefaultKakoKbn.setValue(
                    comboChoicesState.effectiveDefaultFor(RequestFormComboChoices.KEY_KAKO_KBN));
        }
    }

    private String defaultInputKbnForNewRow() {
        return comboChoicesState.effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN);
    }

    private String defaultKakoKbnForNewRow() {
        return comboChoicesState.effectiveDefaultFor(RequestFormComboChoices.KEY_KAKO_KBN);
    }

    private void applyWorkInstructionDefaultsToFormCombos() {
        applyWorkInstructionDefaultsToFormCombos(false);
    }

    /** 依頼一覧から新規・未登録行を選んだとき、区分が空なら設定タブの既定値を入れる。 */
    private void applyWorkInstructionDefaultsToFormCombosIfBlank() {
        applyWorkInstructionDefaultsToFormCombos(true);
    }

    private void applyWorkInstructionDefaultsToFormCombos(boolean onlyIfBlank) {
        String inputKbn = defaultInputKbnForNewRow();
        if (!inputKbn.isBlank()) {
            if (newCmbFormInputKbn != null
                    && (!onlyIfBlank || isComboValueBlank(newCmbFormInputKbn))) {
                newCmbFormInputKbn.setValue(inputKbn);
            }
            if (newCmbInputKbn != null && (!onlyIfBlank || isComboValueBlank(newCmbInputKbn))) {
                newCmbInputKbn.setValue(inputKbn);
            }
        }
        String kakoKbn = defaultKakoKbnForNewRow();
        if (!kakoKbn.isBlank()) {
            if (newCmbFormKakoKbn != null
                    && (!onlyIfBlank || isComboValueBlank(newCmbFormKakoKbn))) {
                newCmbFormKakoKbn.setValue(kakoKbn);
            }
            if (newCmbKakoKbn != null && (!onlyIfBlank || isComboValueBlank(newCmbKakoKbn))) {
                newCmbKakoKbn.setValue(kakoKbn);
            }
        }
    }

    private static boolean isComboValueBlank(ComboBox<String> combo) {
        if (combo == null) {
            return true;
        }
        String value = combo.getValue();
        return value == null || value.isBlank();
    }

    /** セッション／プロファイル保存用: 現在の ComboBox 候補リストと入力既定値。 */
    public RequestFormComboChoices snapshotComboChoices() {
        java.util.LinkedHashMap<String, java.util.List<String>> map = new java.util.LinkedHashMap<>();
        map.put(RequestFormComboChoices.KEY_INPUT_KBN, java.util.List.copyOf(optInputKbn));
        map.put(RequestFormComboChoices.KEY_KAKO_KBN, java.util.List.copyOf(optKakoKbn));
        map.put(RequestFormComboChoices.KEY_WARI_SU, java.util.List.copyOf(optWariSu));
        map.put(RequestFormComboChoices.KEY_EC_SIDE, java.util.List.copyOf(optEcSide));
        map.put(RequestFormComboChoices.KEY_TRIMMING, java.util.List.copyOf(optTrimming));
        map.put(RequestFormComboChoices.KEY_FEED_LOC, java.util.List.copyOf(optFeedLoc));
        map.put(RequestFormComboChoices.KEY_STORAGE_LOC, java.util.List.copyOf(optStorageLoc));
        map.put(RequestFormComboChoices.KEY_YOTO, java.util.List.copyOf(optYoto));
        map.put(RequestFormComboChoices.KEY_USER, java.util.List.copyOf(optUser));
        map.put(
                RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT,
                java.util.List.copyOf(optMasterCandidatePrefixProduct));
        map.put(
                RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW,
                java.util.List.copyOf(optMasterCandidatePrefixRaw));
        java.util.LinkedHashMap<String, String> defaults = new java.util.LinkedHashMap<>();
        if (cmbSettingsDefaultInputKbn != null && cmbSettingsDefaultInputKbn.getValue() != null) {
            defaults.put(
                    RequestFormComboChoices.KEY_INPUT_KBN,
                    cmbSettingsDefaultInputKbn.getValue().strip());
        }
        if (cmbSettingsDefaultKakoKbn != null && cmbSettingsDefaultKakoKbn.getValue() != null) {
            defaults.put(
                    RequestFormComboChoices.KEY_KAKO_KBN,
                    cmbSettingsDefaultKakoKbn.getValue().strip());
        }
        if (defaults.isEmpty()) {
            defaults.putAll(comboChoicesState.fieldDefaultsAsMap());
        }
        return RequestFormComboChoices.of(map, defaults);
    }

    /** セッション／プロファイルから ComboBox 候補と入力既定値を復元する。 */
    public void applyComboChoices(RequestFormComboChoices choices) {
        if (choices == null) {
            return;
        }
        comboChoicesState = choices.mergedWithDefaults();
        replaceOptList(
                optInputKbn, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_INPUT_KBN));
        replaceOptList(
                optKakoKbn, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_KAKO_KBN));
        replaceOptList(
                optWariSu, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_WARI_SU));
        replaceOptList(
                optEcSide, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_EC_SIDE));
        replaceOptList(
                optTrimming, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_TRIMMING));
        replaceOptList(
                optFeedLoc, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_FEED_LOC));
        replaceOptList(
                optStorageLoc,
                comboChoicesState.optionsFor(RequestFormComboChoices.KEY_STORAGE_LOC));
        replaceOptList(optYoto, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_YOTO));
        replaceOptList(optUser, comboChoicesState.optionsFor(RequestFormComboChoices.KEY_USER));
        replaceOptListOrClear(
                optMasterCandidatePrefixProduct,
                comboChoicesState.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT));
        replaceOptListOrClear(
                optMasterCandidatePrefixRaw,
                comboChoicesState.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW));
        refreshDynamicRowComboItems();
        refreshAllMasterCandidateCombos();
        syncFieldDefaultSelectorCombos();
    }

    /**
     * 段階1計画データの原反投入場所を候補リストへ追記する。
     *
     * @return 新規追加分の件数（追記なしは 0）
     */
    public int mergeFeedLocOptionsFromPlanning(List<String> fromPlan) {
        if (fromPlan == null || fromPlan.isEmpty()) {
            return 0;
        }
        List<String> before = List.copyOf(optFeedLoc);
        List<String> merged = RequestFormFeedLocPlanSync.mergeDistinctFeedLocations(before, fromPlan);
        int added = RequestFormFeedLocPlanSync.countNewValues(before, merged);
        if (added <= 0) {
            return 0;
        }
        optFeedLoc.setAll(merged);
        java.util.LinkedHashMap<String, java.util.List<String>> map =
                new java.util.LinkedHashMap<>(comboChoicesState.asMap());
        map.put(RequestFormComboChoices.KEY_FEED_LOC, List.copyOf(optFeedLoc));
        comboChoicesState =
                RequestFormComboChoices.of(map, comboChoicesState.fieldDefaultsAsMap())
                        .mergedWithDefaults();
        refreshDynamicRowComboItems();
        return added;
    }

    public JuchuHeaderAliasRegistry juchuHeaderAliasRegistry() {
        return juchuHeaderAliasRegistry;
    }

    public void configureJuchuHeaderAliasRegistry(JuchuHeaderAliasRegistry registry) {
        juchuHeaderAliasRegistry =
                registry != null ? registry : JuchuHeaderAliasRegistry.loadDefault();
    }

    private int juchuHeaderRowIndexFor(String juchuPath) {
        return juchuHeaderAliasRegistry.headerRowIndexFor(juchuPath);
    }

    private int juchuFirstDataRowIndexFor(String juchuPath) {
        return juchuHeaderRowIndexFor(juchuPath) + 1;
    }

    private int juchuHeaderRowIndex0() {
        return juchuHeaderRowIndexFor(juchuFilePath);
    }

    private int juchuFirstDataRowIndex0() {
        return juchuFirstDataRowIndexFor(juchuFilePath);
    }

    private static void replaceOptList(ObservableList<String> target, java.util.List<String> values) {
        if (target == null || values == null || values.isEmpty()) {
            return;
        }
        target.setAll(values);
    }

    private static void replaceOptListOrClear(
            ObservableList<String> target, java.util.List<String> values) {
        if (target == null) {
            return;
        }
        if (values == null || values.isEmpty()) {
            target.clear();
        } else {
            target.setAll(values);
        }
    }

    private void refreshAllMasterCandidateCombos() {
        for (ProductRow pRow : productRows) {
            if (pRow.cmbSearch != null) {
                updateRowProdCandidates(pRow, false, false);
            }
        }
        for (RawMaterialRow rRow : rawRows) {
            if (rRow.cmbSearch != null) {
                updateRowRawCandidates(rRow, false, false);
            }
        }
    }

    private List<ProductInfo> filteredMasterCatalogForProduct() {
        return RequestFormMasterProductCandidateMatcher.filterCatalogByShohinCodePrefixes(
                masterProductList, List.copyOf(optMasterCandidatePrefixProduct));
    }

    private List<ProductInfo> filteredMasterCatalogForRaw() {
        return RequestFormMasterProductCandidateMatcher.filterCatalogByShohinCodePrefixes(
                masterProductList, List.copyOf(optMasterCandidatePrefixRaw));
    }

    private void refreshDynamicRowComboItems() {
        for (ProductRow pRow : productRows) {
            if (pRow.cmbEcSide != null) {
                pRow.cmbEcSide.getItems().setAll(optEcSide);
            }
            if (pRow.cmbTrimming != null) {
                pRow.cmbTrimming.getItems().setAll(optTrimming);
            }
        }
        for (RawMaterialRow rRow : rawRows) {
            if (rRow.cmbWariSu != null) {
                rRow.cmbWariSu.getItems().setAll(optWariSu);
            }
            if (rRow.cmbFeedLoc != null) {
                rRow.cmbFeedLoc.getItems().setAll(optFeedLoc);
            }
            if (rRow.cmbStorageLoc != null) {
                rRow.cmbStorageLoc.getItems().setAll(optStorageLoc);
            }
        }
    }

    // openProductSearchDialog removed

    private void clearInputForm() {
        productRowsContainer.getChildren().clear();
        productRows.clear();
        rawRowsContainer.getChildren().clear();
        rawRows.clear();
        
        // Add default blank rows
        addProductRow(null);
        addRawRow(null);
        
        newTxtFormTokki1.setText("");
        newTxtFormTokki2.setText("");
        newTxtFormTokki3.setText("");
        if (newCmbFormYoto != null) {
            newCmbFormYoto.setValue(null);
        }
        if (!optUser.isEmpty()) newCmbFormUser.setValue(optUser.get(0));
        applyWorkInstructionDefaultsToFormCombos();
        refreshFormInputTantoLabel();
        newDpFormDeliv.setValue(null);
        newDpFormAdjustDeliv.setValue(null);
        newDpFormInputDate.setValue(null);
        newTxtFormWage.setText("9");
        newTxtUketsukeNo.setText("");
        newTxtIraiNo.setText("");
    }

    /** 操作者変更後に依頼書フォームの入力担当表示を更新する。 */
    public void refreshSessionInputTantoLabel() {
        Platform.runLater(
                () -> {
                    refreshFormInputTantoLabel();
                    applyGuestSessionRestrictions();
                });
    }

    private void registerGuestMutableControl(Node node) {
        if (node != null) {
            guestMutableControls.add(node);
        }
    }

    private boolean rejectGuestMutation(String featureLabel) {
        if (FactoryOperatorUserStore.sessionMayMutateRequestFormInput()) {
            return false;
        }
        showAlert("操作不可", featureLabel + "\n\nゲストユーザーはこの操作を実行できません。");
        return true;
    }

    private void applyGuestSessionRestrictions() {
        boolean guest = !FactoryOperatorUserStore.sessionMayMutateRequestFormInput();
        for (Node node : guestMutableControls) {
            node.setDisable(guest);
        }
        updateTransferButtonState();
    }

    private void refreshFormInputTantoLabel() {
        if (lblFormInputTanto == null) {
            return;
        }
        String operator = sessionInputTanto();
        lblFormInputTanto.setText(operator.isBlank() ? "（未ログイン）" : operator);
    }

    private static String sessionInputTanto() {
        return FactoryOperatorUserStore.sessionOperatorName();
    }

    /** 環境変数タブの依頼書入力向けパスを反映する（タブ再選択時も可）。 */
    public void configureFromUiEnv(Map<String, String> uiEnv) {
        uiEnvSnapshot = uiEnv != null ? Map.copyOf(uiEnv) : Map.of();
        refreshEmbeddedWindowTitle();
        RequestFormSheetPreviewPdfRenderer.applyCjkMetricsScaleFromUi(uiEnvSnapshot);
        aladdinMasterDir = AppPaths.resolveAladdinMasterDir(uiEnvSnapshot);
        applyRequestFormOriginalDirFromUiEnv();
        applyJuchuFilePathFromUiEnv();
        refreshJuchuPathDisplay();
        applyGuestSessionRestrictions();
    }

    /** 依頼書入力タブ上部タイトル（工場は環境変数／グローバル設定から解決）。 */
    static String embeddedWindowTitle(Map<String, String> uiEnv) {
        FactorySite site = GlobalInitSettingTarget.loadEffective(uiEnv != null ? uiEnv : Map.of());
        return site.displayLabelJa() + " 統合管理データベース (JavaFX版)";
    }

    private void refreshEmbeddedWindowTitle() {
        if (embeddedTitleLabel != null) {
            embeddedTitleLabel.setText(embeddedWindowTitle(uiEnvSnapshot));
        }
    }

    private void applyRequestFormOriginalDirFromUiEnv() {
        targetFolder = AppPaths.resolveRequestFormOriginalDir(uiEnvSnapshot).toString();
    }

    private void applySelectedOriginalDir(String absolutePath) {
        targetFolder = absolutePath;
        Map<String, String> next = new HashMap<>(uiEnvSnapshot);
        next.put(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, absolutePath);
        uiEnvSnapshot = Map.copyOf(next);
        if (originalDirChangeHandler != null) {
            originalDirChangeHandler.accept(absolutePath);
        }
        ensureJuchuPathDefault();
    }

    private void applySelectedJuchuFile(String absolutePath) {
        juchuFilePath = absolutePath;
        refreshJuchuPathDisplay();
        Map<String, String> next = new HashMap<>(uiEnvSnapshot);
        next.put(AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE, absolutePath);
        uiEnvSnapshot = Map.copyOf(next);
        if (juchuFileChangeHandler != null) {
            juchuFileChangeHandler.accept(absolutePath);
        }
        refreshJuchuBackupList();
    }

    private void applyJuchuFilePathFromUiEnv() {
        AppPaths.resolveRequestFormJuchuFile(uiEnvSnapshot)
                .ifPresent(p -> juchuFilePath = p.toString());
    }

    private void refreshJuchuPathDisplay() {
        if (txtJuchuPathDisplay != null) {
            txtJuchuPathDisplay.setText(juchuFilePath != null ? juchuFilePath : "");
        }
    }

    // --- CACHE LOADER FROM THE ALADDIN INTEGRATED MASTER ---
    private File aladdinMasterDirectory() {
        if (aladdinMasterDir != null) {
            return aladdinMasterDir.toFile();
        }
        return new File(targetFolder, ALADDIN_MASTER_DIR_NAME);
    }

    private File integratedMasterFile() {
        return new File(aladdinMasterDirectory(), INTEGRATED_MASTER_FILE_NAME);
    }

    private void ensureJuchuPathDefault() {
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            juchuFilePath =
                    AppPaths.resolveRequestFormJuchuFile(uiEnvSnapshot)
                            .map(Path::toString)
                            .orElseGet(
                                    () ->
                                            new File(
                                                            targetFolder,
                                                            AppPaths.DEFAULT_REQUEST_FORM_JUCHU_FILE_NAME)
                                                    .getAbsolutePath());
        }
    }

    /** Excel が編集モードで開いているときに作られる {@code ~$<ファイル名>} ロックファイル。 */
    private static File excelLockFileFor(File workbookFile) {
        if (workbookFile == null) {
            return null;
        }
        String name = workbookFile.getName();
        if (name.isBlank()) {
            return null;
        }
        File parent = workbookFile.getParentFile();
        if (parent == null) {
            return null;
        }
        return new File(parent, "~$" + name);
    }

    /**
     * 受注ブックが実際に排他ロックされているか（{@code ~$} の有無だけでは判定しない）。
     * UNC 上に Excel 終了後も残る {@code ~$} だけでは誤検出しやすいため、本体を {@code rw} で開けるかを正とする。
     */
    private static boolean isJuchuWorkbookWritableForTransfer(File workbook) {
        if (workbook == null || !workbook.isFile()) {
            return false;
        }
        try (java.io.RandomAccessFile ignored = new java.io.RandomAccessFile(workbook, "rw")) {
            return true;
        } catch (java.io.IOException ex) {
            return false;
        }
    }

    /**
     * 転記不可の理由（転記処理中フラグは含めない）。転記可能なら {@code null}。
     */
    private String describeJuchuWriteBlockExcludingTransfer() {
        if (!FactoryOperatorUserStore.sessionMayMutateRequestFormInput()) {
            return "ゲストユーザーは受注ファイルへの転記・更新はできません。";
        }
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            return "受注ファイルが未設定です。設定タブまたは環境変数 PM_AI_REQUEST_FORM_JUCHU_FILE を指定してください。";
        }
        File juchuFile = new File(juchuFilePath);
        if (!juchuFile.isFile()) {
            return "受注ファイルが見つかりません: " + juchuFilePath;
        }
        if (isJuchuWorkbookWritableForTransfer(juchuFile)) {
            return null;
        }
        File lockFile = excelLockFileFor(juchuFile);
        if (lockFile != null && lockFile.isFile()) {
            return "受注ファイルが Excel で使用中（読み取り専用）です。保存して閉じてから再試行してください。"
                    + " ロック: "
                    + lockFile.getName()
                    + "（フォルダ: "
                    + lockFile.getParent()
                    + "）";
        }
        return "受注ファイルに書き込めません（読み取り専用・権限・ネットワーク接続を確認してください）: "
                + juchuFile.getAbsolutePath();
    }

    private boolean isJuchuFileLockedForWrite() {
        return describeJuchuWriteBlockExcludingTransfer() != null;
    }

    /**
     * 自動転記ボタンを無効化する理由。転記可能なら {@code null}。
     * JavaFX では disable 中のボタンに Tooltip が表示されないため、{@link #updateTransferButtonState()} でラベルにも出す。
     */
    private String resolveTransferBlockedReason() {
        if (juchuTransferInProgress) {
            return "受注ファイルへの転記処理を実行中です。完了までお待ちください。";
        }
        return describeJuchuWriteBlockExcludingTransfer();
    }

    private void showJuchuWriteBlockedAlert() {
        String reason = describeJuchuWriteBlockExcludingTransfer();
        showAlert("エラー", reason != null ? reason : "受注ファイルへ転記できません。");
    }

    private void updateTransferButtonState() {
        String blockedReason = resolveTransferBlockedReason();
        boolean blocked = blockedReason != null;
        if (btnTransfer != null) {
            btnTransfer.setDisable(blocked);
            btnTransfer.setTooltip(
                    blocked ? new Tooltip(blockedReason) : new Tooltip("現在のフォーム内容を受注ファイルへ転記します。"));
        }
        if (btnBulkTransferPending != null) {
            btnBulkTransferPending.setDisable(blocked);
            btnBulkTransferPending.setTooltip(
                    blocked
                            ? new Tooltip(blockedReason)
                            : new Tooltip("一時保存済みで未転記の全レコードを受注ファイルへ書き込みます。"));
        }
        if (transferBlockedReasonLabel != null) {
            transferBlockedReasonLabel.setText(
                    blocked ? "自動転記不可: " + blockedReason : "");
            transferBlockedReasonLabel.setManaged(blocked);
            transferBlockedReasonLabel.setVisible(blocked);
        }
        updateUndoTransferButtonState();
    }

    private void updateUndoTransferButtonState() {
        if (btnUndoLastTransfer == null) {
            return;
        }
        boolean hasUndo = lastJuchuTransferUndo != null;
        boolean blocked = resolveTransferBlockedReason() != null || juchuTransferInProgress;
        btnUndoLastTransfer.setDisable(!hasUndo || blocked);
        if (hasUndo && !blocked) {
            btnUndoLastTransfer.setTooltip(
                    new Tooltip(
                            "直前に実行した単票の自動転記（依頼No "
                                    + lastJuchuTransferUndo.reqNo()
                                    + "）を受注ファイルで取り消します。"));
        } else if (!hasUndo) {
            btnUndoLastTransfer.setTooltip(new Tooltip("取り消せる自動転記はまだありません。"));
        } else {
            btnUndoLastTransfer.setTooltip(
                    new Tooltip(
                            blocked
                                    ? resolveTransferBlockedReason()
                                    : "転記処理中は取り消せません。"));
        }
    }

    private void clearLastJuchuTransferUndo() {
        lastJuchuTransferUndo = null;
        updateUndoTransferButtonState();
    }

    private void rememberLastJuchuTransferUndo(
            JuchuTransferUndoState undo, OrderRecord recordRef) {
        if (undo == null) {
            clearLastJuchuTransferUndo();
            return;
        }
        String priorStatus = recordRef != null ? recordRef.getStatus() : null;
        String priorDiscrepancy = recordRef != null ? recordRef.getDiscrepancy() : null;
        lastJuchuTransferUndo =
                new JuchuTransferUndoState(
                        undo.reqNo(),
                        undo.juchuFilePath(),
                        undo.insertedNewRow(),
                        undo.rowIndex0(),
                        undo.priorDbValues(),
                        priorStatus,
                        priorDiscrepancy);
        updateUndoTransferButtonState();
    }

    private boolean isPendingLocalSave(OrderRecord record) {
        return record != null && STATUS_LOCAL_SAVE_PENDING.equals(record.getStatus());
    }

    private List<OrderRecord> collectPendingLocalSaveRecords() {
        List<OrderRecord> pending = new ArrayList<>();
        for (OrderRecord record : orderRecords) {
            if (isPendingLocalSave(record)) {
                pending.add(record);
            }
        }
        return pending;
    }

    /**
     * 未転記の一時保存があるときは確認ダイアログを出し、転記または破棄のうえ {@link #reloadData()} する。
     *
     * @param reason ダイアログ本文に付ける操作説明
     * @param afterReload 読込開始後に実行する追加処理（{@code null} 可）
     */
    private void requestReloadData(String reason, Runnable afterReload) {
        List<OrderRecord> pending = collectPendingLocalSaveRecords();
        if (pending.isEmpty()) {
            reloadData();
            if (afterReload != null) {
                afterReload.run();
            }
            return;
        }

        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("一時保存データの確認");
        alert.setHeaderText("未転記の一時保存が " + pending.size() + " 件あります");
        StringBuilder body = new StringBuilder();
        if (reason != null && !reason.isBlank()) {
            body.append(reason.strip()).append("\n\n");
        }
        body.append("一時保存を転記してから読み出しますか？ それとも破棄して読み出しますか？\n\n");
        int shown = 0;
        for (OrderRecord record : pending) {
            if (shown >= 8) {
                body.append("…他 ").append(pending.size() - shown).append(" 件\n");
                break;
            }
            body.append("・").append(record.getReqNo()).append(" [").append(record.getUser()).append("]\n");
            shown++;
        }
        alert.setContentText(body.toString());

        ButtonType transferThenReload =
                new ButtonType("一時保存分を転記してから読込");
        ButtonType discardThenReload = new ButtonType("一時保存を破棄して読込");
        ButtonType cancel = new ButtonType("キャンセル", ButtonBar.ButtonData.CANCEL_CLOSE);
        alert.getButtonTypes().setAll(transferThenReload, discardThenReload, cancel);

        Optional<ButtonType> choice = alert.showAndWait();
        if (choice.isEmpty() || choice.get() == cancel) {
            return;
        }
        if (choice.get() == transferThenReload) {
            transferAllPendingLocalSavesAsync(
                    false,
                    success -> {
                        if (success) {
                            reloadData();
                            if (afterReload != null) {
                                afterReload.run();
                            }
                        }
                    });
            return;
        }
        reloadData();
        if (afterReload != null) {
            afterReload.run();
        }
    }

    /**
     * 一時保存分をバックグラウンドで一括転記する。
     *
     * @param showSuccessDialog {@code true} のとき完了ダイアログを表示
     * @param onComplete 完了時コールバック（FX スレッド）。{@code null} 可
     */
    private void transferAllPendingLocalSavesAsync(
            boolean showSuccessDialog, Consumer<Boolean> onComplete) {
        if (juchuTransferInProgress) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        List<OrderRecord> pending = collectPendingLocalSaveRecords();
        if (pending.isEmpty()) {
            if (showSuccessDialog) {
                showAlert("情報", "一時保存済みで未転記のデータはありません。");
            }
            if (onComplete != null) {
                onComplete.accept(true);
            }
            return;
        }
        if (isJuchuFileLockedForWrite()) {
            showJuchuWriteBlockedAlert();
            updateTransferButtonState();
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }

        File file = new File(juchuFilePath);
        if (!file.isFile()) {
            showAlert("エラー", "受注ファイルが見つかりません。");
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }

        List<String> headerWarnings = readJuchuHeaderWarnings(file);
        if (!headerWarnings.isEmpty()) {
            statusLabel.setText("列定義警告: " + headerWarnings.size() + " 件（詳細はダイアログ）");
        }
        if (!confirmJuchuHeaderWarnings(file, headerWarnings)) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }

        clearLastJuchuTransferUndo();
        beginJuchuTransfer("一時保存分を受注ファイルへ一括転記しています…\n(1/4) 受注ファイルを開いています…");
        statusLabel.setText("一時保存分を受注ファイルへ一括転記中...");
        List<OrderRecord> pendingCopy = new ArrayList<>(pending);
        Thread transferThread =
                new Thread(
                        () -> {
                            List<String> errors = new ArrayList<>();
                            int transferred = 0;
                            Exception fatal = null;
                            try {
                                updateLoadingOverlayText(
                                        "一時保存分を受注ファイルへ一括転記しています…\n(1/4) 受注ファイルを開いています…");
                                try (Workbook wb = PoiWorkbookOpener.open(file)) {
                                    Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
                                    Map<String, Integer> colMap =
                                            buildJuchuColumnMap(sheet, file.getAbsolutePath());
                                    int total = pendingCopy.size();
                                    for (int i = 0; i < total; i++) {
                                        OrderRecord record = pendingCopy.get(i);
                                        updateLoadingOverlayText(
                                                String.format(
                                                        "一時保存分を受注ファイルへ一括転記しています…\n"
                                                                + "(2/4) 転記中 (%d / %d)\n依頼No: %s",
                                                        i + 1,
                                                        total,
                                                        record.getReqNo()));
                                        try {
                                            transferRecordFromDbValues(wb, sheet, colMap, record);
                                            record.setStatus("一致 (転記完了)");
                                            record.setDiscrepancy("一致 (受注ファイルへ転記完了)");
                                            transferred++;
                                        } catch (Exception ex) {
                                            errors.add(record.getReqNo() + ": " + ex.getMessage());
                                        }
                                    }
                                    updateLoadingOverlayText(
                                            "一時保存分を受注ファイルへ一括転記しています…\n(3/4) ファイルを保存しています…");
                                    saveJuchuWorkbookWithLocalBackup(file, wb, "bulk-transfer");
                                }
                                updateLoadingOverlayText(
                                        "一時保存分を受注ファイルへ一括転記しています…\n(4/4) 画面を更新しています…");
                            } catch (Exception ex) {
                                fatal = ex;
                            }

                            final int finalTransferred = transferred;
                            final List<String> finalErrors = errors;
                            final Exception finalFatal = fatal;
                            Platform.runLater(
                                    () -> {
                                        try {
                                            if (finalFatal != null) {
                                                statusLabel.setText("エラー: " + finalFatal.getMessage());
                                                showAlert(
                                                        "エラー",
                                                        "一括転記中にエラーが発生しました:\n"
                                                                + finalFatal.getMessage());
                                                finalFatal.printStackTrace();
                                                if (onComplete != null) {
                                                    onComplete.accept(false);
                                                }
                                                return;
                                            }

                                            refreshComboRecordItems();
                                            if (selectedRecord != null && isPendingLocalSave(selectedRecord)) {
                                                loadRecordDetails(selectedRecord);
                                            } else if (selectedRecord != null) {
                                                discrepancyLabel.setText(selectedRecord.getDiscrepancy());
                                            }

                                            if (!finalErrors.isEmpty()) {
                                                statusLabel.setText(
                                                        "一括転記: 成功 "
                                                                + finalTransferred
                                                                + " 件 / 失敗 "
                                                                + finalErrors.size()
                                                                + " 件");
                                                showAlert(
                                                        "一括転記（一部失敗）",
                                                        "成功: "
                                                                + finalTransferred
                                                                + " 件\n失敗: "
                                                                + finalErrors.size()
                                                                + " 件\n\n"
                                                                + String.join("\n", finalErrors));
                                                if (onComplete != null) {
                                                    onComplete.accept(false);
                                                }
                                                return;
                                            }

                                            statusLabel.setText(
                                                    "一時保存 "
                                                            + finalTransferred
                                                            + " 件の一括転記が完了しました。");
                                            if (showSuccessDialog) {
                                                showAlert(
                                                        "成功",
                                                        "一時保存 "
                                                                + finalTransferred
                                                                + " 件を受注ファイルへ転記しました。");
                                            }
                                            if (onComplete != null) {
                                                onComplete.accept(true);
                                            }
                                        } finally {
                                            endJuchuTransfer();
                                        }
                                    });
                        },
                        "request-form-juchu-bulk-transfer");
        transferThread.setDaemon(true);
        transferThread.start();
    }

    private void transferAllPendingLocalSaves() {
        if (rejectGuestMutation("一時保存分の一括転記")) {
            return;
        }
        transferAllPendingLocalSavesAsync(true, null);
    }

    private Map<String, Integer> buildJuchuColumnMap(Sheet sheet, String juchuPath) {
        Row hRow = sheet.getRow(juchuHeaderRowIndexFor(juchuPath));
        Map<String, Integer> colMap = new HashMap<>();
        if (hRow == null) {
            return colMap;
        }
        for (int c = 0; c < hRow.getLastCellNum(); c++) {
            Cell cell = hRow.getCell(c);
            if (cell != null) {
                colMap.put(cell.getStringCellValue().trim(), c);
            }
        }
        return colMap;
    }

    private void transferRecordFromDbValues(
            Workbook wb, Sheet sheet, Map<String, Integer> colMap, OrderRecord record)
            throws Exception {
        JUCHU_TRANSFER_REPLACE_FORMULA.set(Boolean.TRUE);
        try {
            transferRecordFromDbValuesBody(wb, sheet, colMap, record);
        } finally {
            JUCHU_TRANSFER_REPLACE_FORMULA.remove();
        }
    }

    private void transferRecordFromDbValuesBody(
            Workbook wb, Sheet sheet, Map<String, Integer> colMap, OrderRecord record)
            throws Exception {
        Map<String, String> db = record.getDbValues();
        if (db == null || db.isEmpty()) {
            throw new IllegalStateException("一時保存データが空です");
        }
        String reqNo = record.getReqNo();
        if (reqNo == null || reqNo.isBlank()) {
            throw new IllegalStateException("依頼Ｎｏが空です");
        }

        String normKey = normalize_key(reqNo);
        int lastDataRowIndex = findJuchuSheetLastPopulatedDataRowIndex(sheet);
        Row targetRow = null;
        int targetRowIndex = -1;
        int matchedRowIndex = findJuchuDataRowIndexByReqNo(sheet, normKey, lastDataRowIndex);
        if (matchedRowIndex >= 0) {
            targetRow = sheet.getRow(matchedRowIndex);
            targetRowIndex = matchedRowIndex + 1;
        }

        boolean isNewRow = targetRow == null;
        if (isNewRow) {
            int destRowIdx = insertNewJuchuDataRowPreservingFormulas(sheet, lastDataRowIndex, null);
            targetRow = sheet.getRow(destRowIdx);
            applyDefaultJuchuFormulasIfMissing(targetRow, colMap, destRowIdx + 1);
        }

        setJuchuSheetReqNoIfIncluded(wb, sheet, targetRow, reqNo);

        Map<String, String> mergedDb = new LinkedHashMap<>(db);
        Map<String, String> raw = record.getRawValues();
        if (raw != null && !raw.isEmpty()) {
            if (isNewRow) {
                mergeJuchuDbFromRawDefaults(mergedDb, raw);
            } else {
                mergeJuchuContractNoFromRawWhenBlankOrDifferent(mergedDb, raw);
            }
        }

        writeJuchuRowFromValues(
                targetRow,
                mergedDb,
                isNewRow,
                db.get("入力区分"),
                db.get("加工区分"),
                db.get("入力担当"),
                db.get("特記事項1"),
                db.get("特記事項2"),
                db.get("特記事項3"),
                isNewRow,
                true);

        if (isNewRow) {
            if (raw != null && !raw.isEmpty()) {
                setJuchuNumericOrStringIfIncluded(
                        targetRow,
                        JuchuSheetColumnLayout.Col.KAKOCHIN,
                        firstNonBlank(mergedDb.get("加工賃"), raw.get("加工賃")));
                setJuchuNumericOrStringIfIncluded(
                        targetRow,
                        JuchuSheetColumnLayout.Col.SURYO,
                        firstNonBlank(mergedDb.get("原反数量"), raw.get("原反数量")));
            }
        }
    }

    private static void mergeJuchuDbFromRawDefaults(Map<String, String> db, Map<String, String> raw) {
        for (String key : RequestFormOriginalCellLayout.FORM_EXTRACT_RAW_KEYS) {
            putIfBlank(db, key, raw.get(key));
        }
        if (!db.containsKey("品名1") || db.get("品名1") == null || db.get("品名1").isBlank()) {
            putIfBlank(db, "品名1", raw.get("原反品名"));
        }
    }

    /**
     * 契約Ｎｏがフォーム／受注側で空のときだけ、依頼書原本（E21/L21/S21 連結）の値を補完する。
     * 手入力済みの契約Ｎｏは上書きしない（転記・一覧表示の双方で同じルール）。
     */
    static void mergeJuchuContractNoFromRawWhenBlankOrDifferent(
            Map<String, String> db, Map<String, String> raw) {
        if (db == null || raw == null || raw.isEmpty()) {
            return;
        }
        String rawContract = firstNonBlank(raw.get("契約Ｎｏ"), raw.get("契約No")).trim();
        if (rawContract.isEmpty()) {
            return;
        }
        String dbContract = firstNonBlank(db.get("契約Ｎｏ"), db.get("契約No"));
        if (dbContract.isBlank()) {
            db.put("契約Ｎｏ", rawContract);
        }
    }

    private static void putIfBlank(Map<String, String> db, String key, String value) {
        if (value == null || value.isBlank()) {
            return;
        }
        String current = db.get(key);
        if (current == null || current.isBlank()) {
            db.put(key, value.trim());
        }
    }

    private static void setJuchuSheetDateOrStringFromDb(
            Row row, Map<String, Integer> colMap, String columnHeader, String rawValue) {
        if (rawValue == null || rawValue.isBlank()) {
            setJuchuSheetString(row, colMap, columnHeader, "");
            return;
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            setJuchuSheetDate(row, colMap, columnHeader, sdf.parse(rawValue.strip()));
        } catch (Exception ex) {
            setJuchuSheetString(row, colMap, columnHeader, rawValue.strip());
        }
    }

    private void refreshComboRecordItems() {
        if (comboRecord == null) {
            return;
        }
        OrderRecord selected = comboRecord.getSelectionModel().getSelectedItem();
        applyRecordFilter();
        if (selected != null) {
            comboRecord.getSelectionModel().select(selected);
        }
    }

    /** Excel 起動後にロックファイルが現れるまで少し待ってから転記ボタン状態を更新する。 */
    private void scheduleTransferButtonStateRefresh() {
        javafx.animation.PauseTransition pause =
                new javafx.animation.PauseTransition(javafx.util.Duration.seconds(1.5));
        pause.setOnFinished(e -> updateTransferButtonState());
        pause.play();
    }

    /**
     * リポジトリ直下に {@code アラジンマスタ} があるときは作業フォルダをリポジトリ根に寄せる。
     */
    private void applyRepoRootAsWorkspaceIfPresent(Path repoRoot) {
        if (repoRoot == null) {
            return;
        }
        Path aladdin = repoRoot.resolve(ALADDIN_MASTER_DIR_NAME);
        boolean hasAladdinDir = Files.isDirectory(aladdin);
        boolean savedWorkspaceMissing = !new File(targetFolder).isDirectory();
        if (hasAladdinDir || savedWorkspaceMissing) {
            targetFolder = repoRoot.toAbsolutePath().normalize().toString();
        }
    }

    private static String tailOfChildOutput(String full, int maxChars) {
        if (full == null || full.isBlank()) {
            return "";
        }
        String trimmed = full.strip();
        if (trimmed.length() <= maxChars) {
            return trimmed;
        }
        return trimmed.substring(trimmed.length() - maxChars);
    }

    private String resolveIntegratedMasterScript() {
        File inWorkspace = new File(targetFolder, "create_integrated_master.py");
        if (inWorkspace.isFile()) {
            return inWorkspace.getAbsolutePath();
        }
        if (repoRootHint != null) {
            Path bundled =
                    repoRootHint.resolve("code/python/request_form_input/create_integrated_master.py");
            if (Files.isRegularFile(bundled)) {
                return bundled.toAbsolutePath().normalize().toString();
            }
        }
        return inWorkspace.getAbsolutePath();
    }

    private File previewCacheDirectory() {
        Path repoRoot =
                repoRootHint != null
                        ? repoRootHint.toAbsolutePath().normalize()
                        : AppPaths.resolveRepoRoot(uiEnvSnapshot);
        File cacheDir = repoRoot.resolve("preview_cache").toFile();
        if (!cacheDir.exists()) {
            cacheDir.mkdirs();
        }
        return cacheDir;
    }

    /** 環境変数変更後など、統合マスタをディスクから再読込する。 */
    public void reloadMasterProductListFromDisk() {
        invalidateMasterProductListCache();
        loadMasterProductListAsync(null);
    }

    /** マスタファイルに変更がないときはスキップ（タブ再選択の応答性向上）。 */
    public void reloadMasterProductListFromDiskIfStale(Runnable onComplete) {
        File masterFile = integratedMasterFile();
        String path = masterFile.getAbsolutePath();
        long mtime = masterFile.exists() ? masterFile.lastModified() : -1L;
        synchronized (this) {
            if (masterFile.exists()
                    && !masterProductList.isEmpty()
                    && path.equals(masterProductListLoadedPath)
                    && mtime == masterProductListLoadedMtime) {
                if (onComplete != null) {
                    Platform.runLater(onComplete);
                }
                return;
            }
        }
        loadMasterProductListAsync(onComplete);
    }

    private void invalidateMasterProductListCache() {
        masterProductListLoadedMtime = -1L;
        masterProductListLoadedPath = "";
    }

    private void loadMasterProductList() {
        List<ProductInfo> loaded = loadMasterProductListFromFile();
        applyMasterProductList(loaded, integratedMasterFile());
    }

    private void loadMasterProductListAsync(Runnable onComplete) {
        if (masterProductListLoadInFlight) {
            if (onComplete != null) {
                Platform.runLater(onComplete);
            }
            return;
        }
        masterProductListLoadInFlight = true;
        Thread loadThread =
                new Thread(
                        () -> {
                            File masterFile = integratedMasterFile();
                            List<ProductInfo> loaded = loadMasterProductListFromFile();
                            Platform.runLater(
                                    () -> {
                                        try {
                                            applyMasterProductList(loaded, masterFile);
                                            if (onComplete != null) {
                                                onComplete.run();
                                            }
                                        } finally {
                                            masterProductListLoadInFlight = false;
                                        }
                                    });
                        },
                        "request-form-master-load");
        loadThread.setDaemon(true);
        loadThread.start();
    }

    private void applyMasterProductList(List<ProductInfo> loaded, File masterFile) {
        masterProductList.clear();
        if (loaded != null && !loaded.isEmpty()) {
            masterProductList.addAll(loaded);
        }
        if (masterFile != null && masterFile.exists()) {
            masterProductListLoadedPath = masterFile.getAbsolutePath();
            masterProductListLoadedMtime = masterFile.lastModified();
        } else {
            invalidateMasterProductListCache();
        }
        refreshAllMasterCandidateCombos();
    }

    private List<ProductInfo> loadMasterProductListFromFile() {
        List<ProductInfo> result = new ArrayList<>();
        File masterFile = integratedMasterFile();
        if (!masterFile.exists()) {
            System.err.println("Integrated master file not found: " + masterFile.getAbsolutePath());
            return result;
        }

        try (FileInputStream fis = new FileInputStream(masterFile);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Sheet sheet = wb.getSheet("②商品別・工程展開リスト");
            if (sheet == null) {
                System.err.println("Sheet ②商品別・工程展開リスト not found!");
                return result;
            }

            Row hRow = sheet.getRow(0);
            if (hRow == null) {
                return result;
            }
            Map<String, Integer> colMap = new HashMap<>();
            for (int c = 0; c < hRow.getLastCellNum(); c++) {
                Cell cell = hRow.getCell(c);
                if (cell != null) {
                    colMap.put(cell.getStringCellValue().trim(), c);
                }
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }

                Cell codeCell = row.getCell(colMap.getOrDefault("商品コード", 0));
                String shohinCode = getCellValueAsString(codeCell);
                if (shohinCode.isEmpty()) {
                    continue;
                }

                String seihinCode = getCellValueAsString(row.getCell(colMap.getOrDefault("製品コード", 1)));
                String shohinName1 = getCellValueAsString(row.getCell(colMap.getOrDefault("商品名1", 2)));
                String shohinName2 = getCellValueAsString(row.getCell(colMap.getOrDefault("商品名2", 3)));
                String unitName = getCellValueAsString(row.getCell(colMap.getOrDefault("単位名", 4)));
                String quantityPerCase = getCellValueAsString(row.getCell(colMap.getOrDefault("入数", 5)));
                String selfKakoKbn = getCellValueAsString(row.getCell(colMap.getOrDefault("自社後加工区分", 6)));
                String foamName = getCellValueAsString(row.getCell(colMap.getOrDefault("発泡体品名", 7)));
                String foamPartNo = getCellValueAsString(row.getCell(colMap.getOrDefault("発泡体品番", 8)));
                String foamWidth = getCellValueAsString(row.getCell(colMap.getOrDefault("発泡体幅", 9)));
                String foamLength = getCellValueAsString(row.getCell(colMap.getOrDefault("発泡体長さ", 10)));
                String foamColor = getCellValueAsString(row.getCell(colMap.getOrDefault("発泡体色", 11)));
                String foamThickness = getCellValueAsString(row.getCell(colMap.getOrDefault("発泡体厚み", 12)));

                List<String> steps = new ArrayList<>();
                for (int i = 1; i <= 7; i++) {
                    String colName = "加工内容名" + i;
                    if (colMap.containsKey(colName)) {
                        String stepVal = getCellValueAsString(row.getCell(colMap.get(colName)));
                        if (stepVal != null
                                && !stepVal.trim().isEmpty()
                                && !stepVal.equalsIgnoreCase("None")) {
                            steps.add(stepVal.trim());
                        }
                    }
                }
                String kakoNaiyo = String.join(",", steps);

                result.add(
                        new ProductInfo(
                                shohinCode,
                                seihinCode,
                                shohinName1,
                                shohinName2,
                                unitName,
                                quantityPerCase,
                                selfKakoKbn,
                                foamName,
                                foamPartNo,
                                foamWidth,
                                foamLength,
                                foamColor,
                                foamThickness,
                                kakoNaiyo));
            }
            System.out.println(
                    "Loaded " + result.size() + " products into lookup cache from Integrated Master.");
        } catch (Exception e) {
            System.err.println("Error loading integrated product master: " + e.getMessage());
            e.printStackTrace();
        }
        return result;
    }

    private void addNewOrderToExcel() {
        String reqNo = newTxtIraiNo.getText().trim();
        if (reqNo.isEmpty()) {
            showAlert("エラー", "依頼Ｎｏを入力してください。");
            return;
        }
        
        statusLabel.setText("新規依頼を登録中...");
        
        try {
            File file = new File(juchuFilePath);
            if (!file.exists()) {
                showAlert("エラー", "受注ファイルが見つかりません。");
                return;
            }
            
            try (Workbook wb = PoiWorkbookOpener.open(file)) {
                Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");

                Row hRow = sheet.getRow(2);
                Map<String, Integer> colMap = new HashMap<>();
                for (int c = 0; c < hRow.getLastCellNum(); c++) {
                    Cell cell = hRow.getCell(c);
                    if (cell != null) {
                        colMap.put(cell.getStringCellValue().trim(), c);
                    }
                }

                String normKey = normalize_key(reqNo);
                int lastDataRowIndex = findJuchuSheetLastPopulatedDataRowIndex(sheet);
                if (findJuchuDataRowIndexByReqNo(sheet, normKey, lastDataRowIndex) >= 0) {
                    showAlert(
                            "エラー",
                            "この依頼Ｎｏ（" + reqNo + "）は既に存在します。別の番号を指定してください。");
                    return;
                }

                int destRowIdx = insertNewJuchuDataRowPreservingFormulas(sheet, lastDataRowIndex, null);
                Row targetRow = sheet.getRow(destRowIdx);
                applyDefaultJuchuFormulasIfMissing(targetRow, colMap, destRowIdx + 1);

                setJuchuSheetReqNoIfIncluded(wb, sheet, targetRow, reqNo);

                Map<String, String> db = buildJuchuDbValuesFromForm();
                JUCHU_TRANSFER_REPLACE_FORMULA.set(Boolean.TRUE);
                try {
                    writeJuchuRowFromValues(
                            targetRow,
                            db,
                            true,
                            db.get("入力区分"),
                            db.get("加工区分"),
                            db.get("入力担当"),
                            db.get("特記事項1"),
                            db.get("特記事項2"),
                            db.get("特記事項3"),
                            true,
                            true);
                } finally {
                    JUCHU_TRANSFER_REPLACE_FORMULA.remove();
                }

                saveJuchuWorkbookWithLocalBackup(file, wb, "new-order");
            }
            
            statusLabel.setText("新規依頼の登録が完了しました。");
            showAlert("成功", "新規依頼データが受注ファイルに追加・登録されました！");
            
            clearInputForm();
            requestReloadData("新規登録後、データを再読込します。", null);
            
        } catch (Exception ex) {
            statusLabel.setText("エラー: " + ex.getMessage());
            showAlert("エラー", "依頼の新規追加中にエラーが発生しました:\n" + ex.getMessage());
            ex.printStackTrace();
        }
    }

    // --- LOGIC: DATA RE-READING & PARSING ---
    private void reloadData() {
        showLoadingOverlay(true, "データベースおよび原本ファイルを読み込んでいます...");
        
        Thread loadThread = new Thread(() -> {
            List<OrderRecord> loadedRecords = new ArrayList<>();
            String statusMsg = "";
            boolean success = false;
            List<String> headerWarningsFinal = List.of();
            
            try {
                File juchuFile = new File(juchuFilePath);
                if (!juchuFile.exists()) {
                    showLoadingOverlay(false, "");
                    Platform.runLater(() -> statusLabel.setText("エラー: 受注ファイルが見つかりません。"));
                    return;
                }
                
                FileInputStream fis = new FileInputStream(juchuFile);
                Workbook wbJuchu = PoiWorkbookOpener.open(fis);
                Sheet sJuchu = wbJuchu.getSheet("受注ﾌｧｲﾙ");
                Row hRow = sJuchu.getRow(juchuHeaderRowIndex0());
                List<String> headerWarnings =
                        JuchuSheetColumnLayout.validateHeaders(
                                hRow, juchuHeaderAliasRegistry, juchuFilePath);
                headerWarningsFinal = headerWarnings;

                Map<String, Map<String, String>> dbRows = new HashMap<>();
                int lastDataRowIndex = findJuchuSheetLastPopulatedDataRowIndex(sJuchu);
                int firstDataRow = juchuFirstDataRowIndex0();

                for (int r = firstDataRow; r <= lastDataRowIndex; r++) {
                    Row row = sJuchu.getRow(r);
                    if (row == null) continue;
                    Cell reqCell = row.getCell(0); 
                    if (reqCell == null || reqCell.getCellType() == CellType.BLANK) continue;
                    
                    String reqNo = getCellValueAsString(reqCell).trim();
                    if (reqNo.isEmpty()) continue;
                    
                    Map<String, String> vals =
                            JuchuSheetColumnLayout.readDbValuesFromRow(
                                    row, juchuHeaderAliasRegistry, juchuFilePath);
                    dbRows.put(normalize_key(reqNo), vals);
                }
                wbJuchu.close();
                fis.close();
                
                File folder = new File(targetFolder);
                File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsm") && !name.startsWith("~$") && !name.equals("加工依頼書入力.xlsm"));
                
                List<Map<String, String>> rawRequests = new ArrayList<>();
                File parseCacheRoot = previewCacheDirectory();
                
                if (files != null) {
                    final int totalFiles = files.length;
                    for (int i = 0; i < totalFiles; i++) {
                        File file = files[i];
                        final String currentFileName = file.getName();
                        final int fileIdx = i + 1;
                        Optional<List<Map<String, String>>> cachedEntries =
                                RequestFormSourceCache.loadParseEntries(parseCacheRoot, file);
                        if (cachedEntries.isPresent()) {
                            rawRequests.addAll(cachedEntries.get());
                            Platform.runLater(
                                    () ->
                                            updateLoadingOverlayText(
                                                    String.format(
                                                            "原本キャッシュ使用 (%d / %d)\n%s",
                                                            fileIdx, totalFiles, currentFileName)));
                            continue;
                        }
                        Platform.runLater(
                                () ->
                                        updateLoadingOverlayText(
                                                String.format(
                                                        "原本ファイルを解析中 (%d / %d)\n%s",
                                                        fileIdx, totalFiles, currentFileName)));
                        try {
                            List<Map<String, String>> parsed = parseOriginalWorkbook(file);
                            RequestFormSourceCache.saveParseEntries(parseCacheRoot, file, parsed);
                            rawRequests.addAll(parsed);
                        } catch (Exception ex) {
                            System.err.println("Error reading " + file.getName() + ": " + ex.getMessage());
                        }
                    }
                }
                
                Set<String> processedKeys = new HashSet<>();
                
                for (Map<String, String> raw : rawRequests) {
                    String reqNo = raw.get("依頼Ｎｏ");
                    String normK = normalize_key(reqNo);
                    processedKeys.add(normK);
                    
                    if (dbRows.containsKey(normK)) {
                        Map<String, String> dbRow = dbRows.get(normK);
                        
                        List<String> diffs = new ArrayList<>();
                        if (!normalize_text(raw.get("品名")).equals(normalize_text(dbRow.get("品名")))) {
                            diffs.add("品名相違");
                        }
                        if (!normalize_text(raw.get("製品")).equals(normalize_text(dbRow.get("製品")))) {
                            diffs.add("製品コード相違");
                        }
                        if (normalize_numeric(raw.get("数量1")) != normalize_numeric(dbRow.get("数量1"))) {
                            diffs.add("数量1相違");
                        }
                        String ru = normalize_text(raw.get("ユーザー"));
                        String dbu = normalize_text(dbRow.get("ユーザー"));
                        if (!ru.equals(dbu) && !ru.contains(dbu) && !dbu.contains(ru)) {
                            diffs.add("ユーザー相違");
                        }
                        if (!normalize_date_val(raw.get("希望納期")).equals(normalize_date_val(dbRow.get("希望納期")))) {
                            diffs.add("希望納期相違");
                        }
                        if (!normalize_text(raw.get("原反")).equals(normalize_text(dbRow.get("原反")))) {
                            diffs.add("原反相違");
                        }
                        String r_p = normalize_text(raw.get("加工内容")).replace(",","").replace("、","");
                        String db_p = normalize_text(dbRow.get("加工内容")).replace(",","").replace("、","");
                        if (!r_p.equals(db_p)) {
                            diffs.add("加工内容相違");
                        }
                        if (!normalize_text(raw.get("契約Ｎｏ"))
                                .equals(normalize_text(dbRow.get("契約Ｎｏ")))) {
                            if (!normalize_text(raw.get("契約Ｎｏ")).isEmpty()
                                    || !normalize_text(dbRow.get("契約Ｎｏ")).isEmpty()) {
                                diffs.add("契約No相違");
                            }
                        }
                        
                        String status = diffs.isEmpty() ? "既存登録 (原本一致)" : "既存登録 (相違あり)";
                        String discrepancy = diffs.isEmpty() ? "原本と完全一致" : "相違詳細: " + String.join(", ", diffs);
                        
                        loadedRecords.add(new OrderRecord(
                            reqNo, status, raw.get("ユーザー"), raw.get("製品"), discrepancy, raw, dbRow
                        ));
                    } else {
                        loadedRecords.add(new OrderRecord(
                            reqNo,
                            "新規自動追加 (未登録)",
                            raw.get("ユーザー"),
                            raw.get("製品"),
                            "受注ファイル未入力のため自動追加",
                            raw,
                            RequestFormOriginalExtractor.buildDbDefaultsFromRaw(raw)));
                    }
                }
                
                for (Map.Entry<String, Map<String, String>> entry : dbRows.entrySet()) {
                    if (!processedKeys.contains(entry.getKey())) {
                        Map<String, String> dbRow = entry.getValue();
                        String reqNoDisplay =
                                firstNonBlank(dbRow.get("依頼No"), dbRow.get("依頼Ｎｏ"), entry.getKey());
                        loadedRecords.add(new OrderRecord(
                            reqNoDisplay, "既存登録 (原本未確認)", dbRow.get("ユーザー"), dbRow.get("製品"), "原本ファイル未検出（過去データ）", new HashMap<>(), dbRow
                        ));
                    }
                }
                
                success = true;
            } catch (Exception e) {
                statusMsg = e.getMessage();
                e.printStackTrace();
            }
            
            final boolean finalSuccess = success;
            final List<OrderRecord> finalLoaded = loadedRecords;
            final String finalStatusMsg = statusMsg;
            final List<String> finalHeaderWarnings = headerWarningsFinal;
            Platform.runLater(() -> {
                showLoadingOverlay(false, "");
                if (finalSuccess) {
                    orderRecords.clear();
                    orderRecords.addAll(finalLoaded);
                    applyRecordFilter();
                    int withOriginalCount =
                            (int) orderRecords.stream().filter(this::hasExistingFile).count();
                    int juchuOnlyCount =
                            (int)
                                    orderRecords.stream()
                                            .filter(this::isJuchuRowWithoutRequestFormOriginal)
                                            .count();
                    statusLabel.setText(
                            String.format(
                                    "読込完了: 全 %d 件 / 依頼書あり %d 件 / 原本なし・受注のみ %d 件",
                                    orderRecords.size(), withOriginalCount, juchuOnlyCount));
                    if (!finalHeaderWarnings.isEmpty()) {
                        statusLabel.setText(
                                statusLabel.getText()
                                        + " / 列定義警告 "
                                        + finalHeaderWarnings.size()
                                        + " 件");
                        File juchuForHeaders =
                                juchuFilePath != null && !juchuFilePath.isBlank()
                                        ? new File(juchuFilePath)
                                        : null;
                        if (juchuForHeaders != null && juchuForHeaders.isFile()) {
                            confirmJuchuHeaderWarnings(juchuForHeaders, finalHeaderWarnings);
                        } else {
                            confirmJuchuHeaderWarnings(null, finalHeaderWarnings);
                        }
                    }
                    
                    syncOriginalFileMonitorAfterReload();
                    enqueueBackgroundCacheTasks();
                } else {
                    statusLabel.setText("エラー: データの読み込みに失敗しました。" + finalStatusMsg);
                }
            });
        });
        
        loadThread.setDaemon(true);
        loadThread.start();
    }

    private List<Map<String, String>> parseOriginalWorkbook(File file) throws Exception {
        List<Map<String, String>> parsed = new ArrayList<>();
        try (FileInputStream fisRaw = new FileInputStream(file);
                Workbook wbRaw = PoiWorkbookOpener.open(fisRaw)) {
            for (int s = 0; s < wbRaw.getNumberOfSheets(); s++) {
                String sName = wbRaw.getSheetName(s);
                if (Pattern.matches("^[A-Z]+\\d+-\\d+$", sName)
                        || Pattern.matches("^[A-Z]\\d+-\\d+-\\d+$", sName)) {
                    Sheet rawSheet = wbRaw.getSheetAt(s);
                    parsed.add(RequestFormOriginalExtractor.buildRawMapFromSheet(file, sName, rawSheet));
                }
            }
        }
        return parsed;
    }

    private void enqueueBackgroundCacheTasks() {
        synchronized (cacheQueue) {
            cacheQueue.clear();
            for (OrderRecord record : orderRecords) {
                Map<String, String> raw = record.getRawValues();
                if (raw != null && !raw.isEmpty()) {
                    String fName = raw.get("原本ファイル名");
                    String sName = raw.get("原本シート名");
                    if (fName != null && sName != null) {
                        File f = new File(targetFolder + "\\" + fName);
                        if (f.exists()) {
                            File cacheRoot = previewCacheDirectory();
                            File cachedPdf =
                                    RequestFormSourceCache.pdfCacheFile(cacheRoot, fName, sName);
                            if (!RequestFormSourceCache.isPreviewCacheValid(cachedPdf, f)) {
                                RequestFormSourceCache.deletePreviewCache(cachedPdf);
                                boolean exists = false;
                                for (CacheTask t : cacheQueue) {
                                    if (t.outputFile.equals(cachedPdf)) {
                                        exists = true;
                                        break;
                                    }
                                }
                                if (!exists) {
                                    cacheQueue.add(new CacheTask(f, sName, cachedPdf));
                                }
                            }
                        }
                    }
                }
            }
        }
        startCachingThread();
    }

    private void startCachingThread() {
        if (cacheThread != null && cacheThread.isAlive()) {
            return;
        }
        
        cacheThread = new Thread(() -> {
            while (true) {
                CacheTask task = null;
                synchronized (cacheQueue) {
                    if (cacheQueue.isEmpty()) {
                        break;
                    }
                    task = cacheQueue.peek();
                }

                // Wait for user to be idle
                long timeSinceInteraction = System.currentTimeMillis() - lastInteractionTime;
                if (timeSinceInteraction < IDLE_THRESHOLD_MS) {
                    try {
                        Thread.sleep(500);
                        continue;
                    } catch (InterruptedException e) {
                        break;
                    }
                }

                if (isPlanningPipelineStageBusy()) {
                    try {
                        Thread.sleep(500);
                        continue;
                    } catch (InterruptedException e) {
                        break;
                    }
                }

                if (RequestFormSourceCache.isPreviewCacheValid(task.outputFile, task.excelFile)) {
                    synchronized (cacheQueue) {
                        cacheQueue.poll();
                    }
                    continue;
                }
                RequestFormSourceCache.deletePreviewCache(task.outputFile);

                try {
                    RequestFormSheetPreviewRenderer.generatePreviewPdf(
                            task.excelFile, task.sheetName, task.outputFile);
                    RequestFormSourceCache.writePreviewMeta(task.outputFile, task.excelFile);
                    System.out.println(
                            "Background cached (pdf): " + task.outputFile.getName());
                } catch (Exception e) {
                    System.err.println("Background cache exception: " + e.getMessage());
                }

                synchronized (cacheQueue) {
                    if (!cacheQueue.isEmpty() && cacheQueue.peek() == task) {
                        cacheQueue.poll();
                    }
                }

                try {
                    Thread.sleep(BACKGROUND_PDF_CACHE_INTERVAL_MS);
                } catch (InterruptedException e) {
                    break;
                }
            }
        });
        cacheThread.setDaemon(true);
        cacheThread.setPriority(Thread.MIN_PRIORITY);
        cacheThread.start();
    }

    private boolean hasExistingFile(OrderRecord rec) {
        Map<String, String> raw = rec.getRawValues();
        if (raw == null || raw.isEmpty()) return false;
        String fName = raw.get("原本ファイル名");
        if (fName == null || fName.isBlank()) return false;
        return new File(targetFolder + "\\" + fName).exists();
    }

    /**
     * 選択中のラジオを再クリックしたときは選択解除し、依頼書あり一覧（既定）へ戻す。
     */
    private static void installRecordFilterRadioDeselectOnReselect(
            ToggleGroup group, RadioButton radio) {
        radio.addEventFilter(
                javafx.scene.input.MouseEvent.MOUSE_PRESSED,
                e -> {
                    if (radio.isSelected()) {
                        Platform.runLater(() -> group.selectToggle(null));
                    }
                });
    }

    private RecordListFilterMode resolveRecordListFilterMode() {
        if (rbAllRecordsFilter != null && rbAllRecordsFilter.isSelected()) {
            return RecordListFilterMode.ALL;
        }
        if (rbExistingOnlyFilter != null && rbExistingOnlyFilter.isSelected()) {
            return RecordListFilterMode.EXISTING_ONLY;
        }
        if (rbNewOnlyFilter != null && rbNewOnlyFilter.isSelected()) {
            return RecordListFilterMode.NEW_ONLY;
        }
        if (rbJuchuWithoutOriginalFilter != null && rbJuchuWithoutOriginalFilter.isSelected()) {
            return RecordListFilterMode.JUCHU_WITHOUT_ORIGINAL;
        }
        return RecordListFilterMode.WITH_ORIGINAL;
    }

    private void applyRecordFilter() {
        if (comboRecord == null) {
            return;
        }
        RecordListFilterMode mode = resolveRecordListFilterMode();
        ObservableList<OrderRecord> base = FXCollections.observableArrayList();
        for (OrderRecord rec : orderRecords) {
            if (recordIncludedInListFilter(rec, mode, this::hasExistingFile)) {
                base.add(rec);
            }
        }
        boolean newOnly = mode == RecordListFilterMode.NEW_ONLY;
        String query =
                txtRecordFilter != null && txtRecordFilter.getText() != null
                        ? txtRecordFilter.getText().strip()
                        : "";
        String q = query.toLowerCase(Locale.ROOT);
        ObservableList<OrderRecord> filtered = FXCollections.observableArrayList();
        for (OrderRecord rec : base) {
            if (!recordMatchesFilter(rec, q, newOnly)) {
                continue;
            }
            filtered.add(rec);
        }
        if (mode == RecordListFilterMode.JUCHU_WITHOUT_ORIGINAL && filtered.size() > 1) {
            FXCollections.sort(filtered, ReconciliationApp::compareRecordByInputDateDesc);
        }
        comboRecord.setItems(filtered);
    }

    static boolean recordIncludedInListFilter(
            OrderRecord rec,
            RecordListFilterMode mode,
            java.util.function.Predicate<OrderRecord> hasOriginalFile) {
        if (rec == null || mode == null) {
            return false;
        }
        return switch (mode) {
            case ALL -> true;
            case EXISTING_ONLY -> recordStatusContainsExisting(rec);
            case NEW_ONLY -> hasOriginalFile.test(rec);
            case JUCHU_WITHOUT_ORIGINAL ->
                    isJuchuRowWithoutRequestFormOriginal(rec, hasOriginalFile);
            case WITH_ORIGINAL -> hasOriginalFile.test(rec);
        };
    }

    static boolean recordStatusContainsExisting(OrderRecord rec) {
        if (rec == null) {
            return false;
        }
        String status = rec.getStatus();
        return status != null && status.contains("既存");
    }

    /**
     * 依頼書原本が無い（または未検出）が受注ファイル行として存在するレコード。
     * 「新規自動追加」は原本由来のため除外する。
     */
    boolean isJuchuRowWithoutRequestFormOriginal(OrderRecord rec) {
        return isJuchuRowWithoutRequestFormOriginal(rec, this::hasExistingFile);
    }

    static boolean isJuchuRowWithoutRequestFormOriginal(
            OrderRecord rec, java.util.function.Predicate<OrderRecord> hasOriginalFile) {
        if (rec == null || hasOriginalFile.test(rec)) {
            return false;
        }
        Map<String, String> db = rec.getDbValues();
        if (db == null || db.isEmpty()) {
            return false;
        }
        String status = rec.getStatus();
        if (status != null && status.contains("新規")) {
            return false;
        }
        return true;
    }

    static int compareRecordByInputDateDesc(OrderRecord a, OrderRecord b) {
        return resolveRecordInputDate(b).compareTo(resolveRecordInputDate(a));
    }

    static java.time.LocalDate resolveRecordInputDate(OrderRecord rec) {
        if (rec == null || rec.getDbValues() == null) {
            return java.time.LocalDate.MIN;
        }
        return parseInputDateForSort(rec.getDbValues().get("入力日"));
    }

    static java.time.LocalDate parseInputDateForSort(String val) {
        if (val == null || val.strip().isEmpty()) {
            return java.time.LocalDate.MIN;
        }
        String text = val.strip();
        java.util.List<String> fmts =
                java.util.Arrays.asList(
                        "yyyy-MM-dd HH:mm:ss",
                        "yyyy/MM/dd HH:mm:ss",
                        "yyyy-MM-dd",
                        "yyyy/MM/dd",
                        "yyyyMMdd");
        for (String fmt : fmts) {
            try {
                java.time.format.DateTimeFormatter dtf =
                        java.time.format.DateTimeFormatter.ofPattern(fmt);
                if (fmt.contains("HH")) {
                    return java.time.LocalDateTime.parse(text, dtf).toLocalDate();
                }
                return java.time.LocalDate.parse(text, dtf);
            } catch (Exception ignored) {
                // try next
            }
        }
        return java.time.LocalDate.MIN;
    }

    /** 依頼一覧の検索文字列・新規のみチェックに合致するか。 */
    static boolean recordMatchesFilter(OrderRecord rec, String queryLower, boolean newOnly) {
        if (rec == null) {
            return false;
        }
        if (newOnly) {
            String status = rec.getStatus();
            if (status == null || !status.contains("新規")) {
                return false;
            }
        }
        if (queryLower == null || queryLower.isEmpty()) {
            return true;
        }
        String reqNo = rec.getReqNo() != null ? rec.getReqNo().toLowerCase(Locale.ROOT) : "";
        String user = rec.getUser() != null ? rec.getUser().toLowerCase(Locale.ROOT) : "";
        return reqNo.contains(queryLower) || user.contains(queryLower);
    }

    // --- LOGIC: RENDER ORIGINAL SHEET VIEW AND FILL FORM ---
    private void loadRecordDetails(OrderRecord record) {
        this.selectedRecord = record;
        hideAllMasterCandidateCombos();
        isLoadingRecord = true;
        try {
        txtReqNo.setText(record.getReqNo());
        txtReqNo.setEditable(false);
        if (lblFormTitle != null) {
            lblFormTitle.setText("受注データ手入力・修正フォーム (左画面で選択中のデータを編集し、受注ファイルへ書き込みます)");
            lblFormTitle.setStyle("-fx-font-size: 13px; -fx-font-weight: bold;");
            lblFormTitle.getStyleClass().removeAll("section-title-form", "section-title-new-mode");
            lblFormTitle.getStyleClass().add("section-title-form");
        }
        
        Map<String, String> rawVals = record.getRawValues() != null ? record.getRawValues() : Map.of();
        Map<String, String> activeVals;
        if (record.getDbValues() == null || record.getDbValues().isEmpty()) {
            activeVals =
                    rawVals.isEmpty()
                            ? Map.of()
                            : RequestFormOriginalExtractor.buildDbDefaultsFromRaw(rawVals);
        } else if (record.getStatus() != null && record.getStatus().contains("未登録")) {
            activeVals = new LinkedHashMap<>(record.getDbValues());
            mergeJuchuDbFromRawDefaults(activeVals, rawVals);
            mergeJuchuContractNoFromRawWhenBlankOrDifferent(activeVals, rawVals);
        } else {
            activeVals = new LinkedHashMap<>(record.getDbValues());
            mergeJuchuContractNoFromRawWhenBlankOrDifferent(activeVals, rawVals);
        }
        
        newCmbFormUser.setValue(activeVals.getOrDefault("ユーザー", ""));
        newDpFormDeliv.setValue(parseLocalDate(activeVals.getOrDefault("希望納期", "")));
        newDpFormAdjustDeliv.setValue(parseLocalDate(activeVals.getOrDefault("調整納期", "")));
        newDpFormInputDate.setValue(java.time.LocalDate.now());
        txtProcess.setText(activeVals.getOrDefault("加工内容", ""));
        newTxtFormWage.setText(activeVals.getOrDefault("加工賃", ""));
        newTxtUketsukeNo.setText(activeVals.getOrDefault("受付Ｎｏ", ""));
        newCmbFormYoto.setValue(activeVals.getOrDefault("用途", ""));
        newCmbFormInputKbn.setValue(activeVals.getOrDefault("入力区分", ""));
        newCmbFormKakoKbn.setValue(activeVals.getOrDefault("加工区分", ""));
        applyWorkInstructionDefaultsToFormCombosIfBlank();
        refreshFormInputTantoLabel();
        newTxtFormTokki1.setText(activeVals.getOrDefault("特記事項1", ""));
        newTxtFormTokki2.setText(activeVals.getOrDefault("特記事項2", ""));
        newTxtFormTokki3.setText(activeVals.getOrDefault("特記事項3", ""));

        String[] hinmeiVals = activeVals.getOrDefault("品名", "").split("\\n", -1);
        String[] masterProdVals =
                activeVals.getOrDefault(COL_MASTER_BASE_SHOHIN_PRODUCT, "").split("\\n", -1);
        String[] prodVals = activeVals.getOrDefault("製品", "").split("\\n", -1);
        String[] qtyVals = activeVals.getOrDefault("数量1", "").split("\\n", -1);
        String[] gradeVals = activeVals.getOrDefault("梱-等1", "").split("\\n", -1);
        String[] colorVals = activeVals.getOrDefault("色1", "").split("\\n", -1);
        String[] catVals = activeVals.getOrDefault("区分1", "").split("\\n", -1);
        String[] edabanVals = activeVals.getOrDefault("枝番", "").split("\\n", -1);
        String[] contractVals =
                firstNonBlank(activeVals.get("契約Ｎｏ"), activeVals.get("契約No")).split("\\n", -1);
        String[] ecVals = activeVals.getOrDefault("ＥＣ面", "").split("\\n", -1);
        String[] trimVals = activeVals.getOrDefault("ﾄﾘﾐﾝｸﾞ", "").split("\\n", -1);
        
        int numProds = Math.max(1, Math.max(prodVals.length, hinmeiVals.length));
        productRowsContainer.getChildren().clear();
        productRows.clear();
        
        for (int i = 0; i < numProds; i++) {
            ProductRow pRow = addProductRow(null);
            if (i < hinmeiVals.length) pRow.txtHinmei.setText(hinmeiVals[i].trim());
            if (i < masterProdVals.length) pRow.txtItem.setText(masterProdVals[i].trim());
            if (i < qtyVals.length) pRow.txtQty.setText(qtyVals[i].trim());
            if (i < gradeVals.length) pRow.txtGrade.setText(gradeVals[i].trim());
            if (i < colorVals.length) pRow.txtColor.setText(colorVals[i].trim());
            if (i < catVals.length) pRow.txtCategory.setText(catVals[i].trim());
            if (i < edabanVals.length) pRow.txtEdaban.setText(edabanVals[i].trim());
            if (i < contractVals.length) pRow.txtKeiyakuNo.setText(contractVals[i].trim());
            if (i < ecVals.length && !ecVals[i].isBlank()) pRow.cmbEcSide.setValue(ecVals[i].trim());
            if (i < trimVals.length && !trimVals[i].isBlank()) pRow.cmbTrimming.setValue(trimVals[i].trim());
            
            if (i < prodVals.length && !prodVals[i].isBlank()) {
                pRow.txtSeihinmei.setText(prodVals[i].trim());
                String[] parts = JuchuSheetColumnLayout.parseSpecName(prodVals[i]);
                pRow.txtPart.setText(parts[0]);
                pRow.txtType.setText(parts[1]);
                pRow.txtWidth.setText(parts[2]);
                pRow.txtLength.setText(parts[3]);
            } else {
                updateProductRowSpecDisplay(pRow);
            }
        }
        
        String rawHinmeiJoined =
                firstNonBlank(activeVals.get("品名1"), activeVals.get("原反品名"));
        String[] rawHinmeiVals = rawHinmeiJoined.isBlank() ? new String[0] : rawHinmeiJoined.split("\\n", -1);
        String[] rawMasterVals =
                activeVals.getOrDefault(COL_MASTER_BASE_SHOHIN_RAW, "").split("\\n", -1);
        String[] rawValsArr = activeVals.getOrDefault("原反", "").split("\\n", -1);
        String[] rawQtyVals = activeVals.getOrDefault("原反数量", "").split("\\n", -1);
        String[] rawGradeVals = activeVals.getOrDefault("原反梱-等", "").split("\\n", -1);
        String[] rawColorVals = activeVals.getOrDefault("原反色", "").split("\\n", -1);
        String[] rawCatVals = activeVals.getOrDefault("原反区分", "").split("\\n", -1);
        String[] wariVals = activeVals.getOrDefault("割数", "").split("\\n", -1);
        String[] feedVals = activeVals.getOrDefault("投入場所", "").split("\\n", -1);
        String[] storageVals = activeVals.getOrDefault("在庫場所", "").split("\\n", -1);
        String[] inputDateVals = activeVals.getOrDefault("投入日", "").split("\\n", -1);
        
        int numRaws = Math.max(1, Math.max(rawValsArr.length, rawHinmeiVals.length));
        rawRowsContainer.getChildren().clear();
        rawRows.clear();
        
        for (int i = 0; i < numRaws; i++) {
            RawMaterialRow rRow = addRawRow(null);
            if (i < rawHinmeiVals.length) rRow.txtHinmei.setText(rawHinmeiVals[i].trim());
            if (i < rawMasterVals.length) rRow.txtItem.setText(rawMasterVals[i].trim());
            if (i < rawQtyVals.length) rRow.txtQty.setText(rawQtyVals[i].trim());
            if (i < rawGradeVals.length) rRow.txtGrade.setText(rawGradeVals[i].trim());
            if (i < rawColorVals.length) rRow.txtColor.setText(rawColorVals[i].trim());
            if (i < rawCatVals.length) rRow.txtCategory.setText(rawCatVals[i].trim());
            if (i < wariVals.length && !wariVals[i].isBlank()) rRow.cmbWariSu.setValue(wariVals[i].trim());
            if (i < feedVals.length && !feedVals[i].isBlank()) rRow.cmbFeedLoc.setValue(feedVals[i].trim());
            if (i < storageVals.length && !storageVals[i].isBlank()) rRow.cmbStorageLoc.setValue(storageVals[i].trim());
            if (i < inputDateVals.length && !inputDateVals[i].isBlank()) {
                rRow.dpInputDate.setValue(parseLocalDate(inputDateVals[i].trim()));
            }
            
            if (i < rawValsArr.length && !rawValsArr[i].isBlank()) {
                rRow.txtGenpanmei.setText(rawValsArr[i].trim());
                String[] parts = JuchuSheetColumnLayout.parseSpecName(rawValsArr[i]);
                rRow.txtPart.setText(parts[0]);
                rRow.txtType.setText(parts[1]);
                rRow.txtWidth.setText(parts[2]);
                rRow.txtLength.setText(parts[3]);
            } else {
                updateRawRowSpecDisplay(rRow);
            }
            rRow.rollCountManualOverride = false;
            updateRawRowRollCountDisplay(rRow);
        }
        
        discrepancyLabel.setText(record.getDiscrepancy());
        discrepancyLabel.getStyleClass().removeAll("discrepancy-label-info", "discrepancy-label-matched", "discrepancy-label-discrepancy", "discrepancy-label-unregistered");
        
        if (record.getStatus().contains("相違あり")) {
            discrepancyLabel.getStyleClass().add("discrepancy-label-discrepancy");
        } else if (record.getStatus().contains("未登録")) {
            discrepancyLabel.getStyleClass().add("discrepancy-label-unregistered");
        } else if (record.getStatus().contains("一致")) {
            discrepancyLabel.getStyleClass().add("discrepancy-label-matched");
        } else {
            discrepancyLabel.getStyleClass().add("discrepancy-label-info");
        }
        
        renderOriginalSheetInGrid(record);
        } finally {
            isLoadingRecord = false;
        }
        refreshAllRowCandidates();
    }

    private void confirmAndClearPreviewCache() {
        File cacheRoot = previewCacheDirectory();
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("プレビューキャッシュのクリア");
        alert.setHeaderText("preview_cache を削除しますか？");
        alert.setContentText(
                "次をすべて削除します。\n"
                        + "・PDF プレビュー（preview_cache/pdf）\n"
                        + "・原本解析キャッシュ（preview_cache/parse）\n\n"
                        + "保存先: "
                        + cacheRoot.getAbsolutePath()
                        + "\n\n"
                        + "表示中の依頼書プレビューは再生成されます。"
                        + "照合データの再解析が必要なときは、読込を実行してください。");
        ButtonType ok = new ButtonType("クリア", ButtonBar.ButtonData.OK_DONE);
        ButtonType cancel = new ButtonType("キャンセル", ButtonBar.ButtonData.CANCEL_CLOSE);
        alert.getButtonTypes().setAll(ok, cancel);
        Optional<ButtonType> choice = alert.showAndWait();
        if (choice.isEmpty() || choice.get() != ok) {
            return;
        }

        synchronized (cacheQueue) {
            cacheQueue.clear();
        }

        RequestFormSourceCache.ClearDiskCacheResult result =
                RequestFormSourceCache.clearAllDiskCache(cacheRoot);

        StringBuilder msg = new StringBuilder();
        msg.append("プレビューキャッシュをクリアしました。\n");
        msg.append("PDF: ").append(result.pdfFilesDeleted()).append(" 件\n");
        msg.append("解析(parse): ").append(result.parseFilesDeleted()).append(" 件");
        if (result.deleteFailures() > 0) {
            msg.append("\n削除できなかったファイル: ").append(result.deleteFailures()).append(" 件");
        }
        if (result.totalDeleted() == 0) {
            msg.append("\n（削除対象のファイルはありませんでした）");
        }
        showAlert("プレビューキャッシュ", msg.toString());

        if (selectedRecord != null) {
            renderOriginalSheetInGrid(selectedRecord);
        } else {
            clearOriginalSheetPreview();
        }
        enqueueBackgroundCacheTasks();
    }

    /** 依頼レコード未選択・新規追加時: プレビュー領域を空にする。 */
    private void clearOriginalSheetPreview() {
        if (sheetGrid != null) {
            sheetGrid.getChildren().clear();
        }
        currentPreviewOriginalFile = null;
        refreshPreviewFileHeader();
    }

    private void resetDiscrepancyLabelForNewEntry() {
        if (discrepancyLabel == null) {
            return;
        }
        discrepancyLabel.setText("依頼を選択するか、新規の依頼Noを入力して登録してください。");
        discrepancyLabel
                .getStyleClass()
                .removeAll(
                        "discrepancy-label-info",
                        "discrepancy-label-matched",
                        "discrepancy-label-discrepancy",
                        "discrepancy-label-unregistered");
        discrepancyLabel.getStyleClass().add("discrepancy-label-info");
    }

    private void openOriginalExcel() {
        if (selectedRecord == null) return;
        Map<String, String> raw = selectedRecord.getRawValues();
        if (raw.isEmpty()) return;
        String fName = raw.get("原本ファイル名");
        File f = new File(targetFolder + "\\" + fName);
        if (f.exists()) {
            try {
                java.awt.Desktop.getDesktop().open(f);
            } catch (Exception e) {
                showAlert("エラー", "ファイルを開けませんでした: " + e.getMessage());
            }
        }
    }

    private void renderOriginalSheetInGrid(OrderRecord record) {
        sheetGrid.getChildren().clear();
        
        Map<String, String> raw = record.getRawValues();
        if (raw.isEmpty()) {
            currentPreviewOriginalFile = null;
            refreshPreviewFileHeader();
            Label lblEmpty = new Label("このレコードに関連する原本ファイルはありません（過去に手入力されたデータです）");
            lblEmpty.getStyleClass().add("excel-grid-label-empty");
            sheetGrid.add(lblEmpty, 0, 0);
            return;
        }
        
        String fName = raw.get("原本ファイル名");
        String sName = raw.get("原本シート名");
        
        File f = new File(targetFolder + "\\" + fName);
        currentPreviewOriginalFile = f.isFile() ? f : null;
        refreshPreviewFileHeader();
        if (!f.exists()) {
            Label lblEmpty = new Label("原本ファイルが見つかりません: " + fName);
            lblEmpty.getStyleClass().add("excel-grid-label-error");
            sheetGrid.add(lblEmpty, 0, 0);
            return;
        }
        
        // Cache directory
        File cacheRoot = previewCacheDirectory();
        File cachedPdf = RequestFormSourceCache.pdfCacheFile(cacheRoot, fName, sName);

        if (RequestFormSourceCache.isPreviewCacheValid(cachedPdf, f)) {
            acknowledgePreviewForCurrentOriginalFile();
            displayPreviewPdf(cachedPdf);
        } else if (isPlanningPipelineStageBusy()) {
            Label lblPaused = new Label(PDF_SUPPRESSED_PIPELINE_MSG);
            lblPaused.getStyleClass().add("excel-grid-label-info");
            sheetGrid.add(lblPaused, 0, 0);
        } else {
            RequestFormSourceCache.deletePreviewCache(cachedPdf);
            Label lblLoading = new Label("原本の PDF プレビューを生成しています... (しばらくお待ちください)");
            lblLoading.getStyleClass().add("excel-grid-label-info");
            sheetGrid.add(lblLoading, 0, 0);

            Thread t =
                    new Thread(
                            () -> {
                                try {
                                    RequestFormSheetPreviewRenderer.generatePreviewPdf(
                                            f, sName, cachedPdf);
                                    RequestFormSourceCache.writePreviewMeta(cachedPdf, f);
                                    Platform.runLater(
                                            () -> {
                                                if (RequestFormSourceCache.isPreviewCacheValid(
                                                        cachedPdf, f)) {
                                                    acknowledgePreviewForCurrentOriginalFile();
                                                    displayPreviewPdf(cachedPdf);
                                                } else {
                                                    sheetGrid.getChildren().clear();
                                                    Label lblErr =
                                                            new Label("PDF プレビューの生成に失敗しました。");
                                                    lblErr.getStyleClass().add("excel-grid-label-error");
                                                    sheetGrid.add(lblErr, 0, 0);
                                                }
                                            });
                                } catch (Exception e) {
                                    Platform.runLater(
                                            () -> {
                                                sheetGrid.getChildren().clear();
                                                Label lblErr =
                                                        new Label("エラーが発生しました: " + e.getMessage());
                                                lblErr.getStyleClass().add("excel-grid-label-error");
                                                sheetGrid.add(lblErr, 0, 0);
                                            });
                                }
                            });
            t.setDaemon(true);
            t.start();
        }
    }

    private void acknowledgePreviewForCurrentOriginalFile() {
        if (currentPreviewOriginalFile != null && currentPreviewOriginalFile.isFile()) {
            originalUpdateMonitor.markPreviewAcknowledged(currentPreviewOriginalFile);
            refreshPreviewFileHeader();
        }
    }

    private void refreshPreviewFileHeader() {
        if (previewFileNameLabel == null || previewFileNameRow == null) {
            return;
        }
        if (currentPreviewOriginalFile == null || !currentPreviewOriginalFile.isFile()) {
            previewFileNameLabel.setText("");
            previewFileNameRow.setManaged(false);
            previewFileNameRow.setVisible(false);
            return;
        }
        previewFileNameRow.setManaged(true);
        previewFileNameRow.setVisible(true);
        previewFileNameLabel.setText(currentPreviewOriginalFile.getName());
        if (previewFileUpdateBadgeHost == null) {
            return;
        }
        previewFileUpdateBadgeHost.getChildren().clear();
        boolean showBadge = originalUpdateMonitor.isUpdated(currentPreviewOriginalFile);
        previewFileUpdateBadgeHost.setManaged(showBadge);
        previewFileUpdateBadgeHost.setVisible(showBadge);
        if (showBadge) {
            RequestFormPreviewBadgeConfig cfg = previewBadgeConfigSupplier.get();
            previewFileUpdateBadgeHost
                    .getChildren()
                    .add(
                            PersonBadgeNodeFactory.createBadge(
                                    cfg.label(), cfg.style(), 1.0, 12.0));
        }
    }

    /** 依頼書タブ表示中のみ: 転記ロック（10秒）と原本更新（30秒）を監視。 */
    private void startEmbeddedTabPolling() {
        stopEmbeddedTabPolling();
        if (embeddedTabPollStatusLabel != null) {
            embeddedTabPollStatusLabel.setManaged(true);
            embeddedTabPollStatusLabel.setVisible(true);
            embeddedTabPollStatusLabel.setText("監視中（転記ロック 10秒 / 原本 30秒）");
        }
        juchuLockPollTimeline =
                new javafx.animation.Timeline(
                        new javafx.animation.KeyFrame(
                                javafx.util.Duration.millis(JUCHU_LOCK_POLL_INTERVAL_MS),
                                e -> runJuchuLockPollTick()));
        juchuLockPollTimeline.setCycleCount(javafx.animation.Timeline.INDEFINITE);
        juchuLockPollTimeline.play();

        originalFilePollTimeline =
                new javafx.animation.Timeline(
                        new javafx.animation.KeyFrame(
                                javafx.util.Duration.millis(ORIGINAL_FILE_POLL_INTERVAL_MS),
                                e -> runOriginalFilePollTick()));
        originalFilePollTimeline.setCycleCount(javafx.animation.Timeline.INDEFINITE);
        originalFilePollTimeline.play();

        runJuchuLockPollTick();
        runOriginalFilePollTick();
    }

    private void stopEmbeddedTabPolling() {
        if (juchuLockPollTimeline != null) {
            juchuLockPollTimeline.stop();
            juchuLockPollTimeline = null;
        }
        if (originalFilePollTimeline != null) {
            originalFilePollTimeline.stop();
            originalFilePollTimeline = null;
        }
        if (pollStatusHighlightPause != null) {
            pollStatusHighlightPause.stop();
            pollStatusHighlightPause = null;
        }
        if (embeddedTabPollStatusLabel != null) {
            embeddedTabPollStatusLabel.getStyleClass().remove("request-form-poll-active");
            embeddedTabPollStatusLabel.setManaged(false);
            embeddedTabPollStatusLabel.setVisible(false);
        }
    }

    private void runJuchuLockPollTick() {
        updateTransferButtonState();
        flashPollStatus("転記ロック");
    }

    private void runOriginalFilePollTick() {
        pollOriginalFilesInWorkspace();
        flashPollStatus("原本更新");
    }

    private void flashPollStatus(String aspect) {
        if (embeddedTabPollStatusLabel == null) {
            return;
        }
        String time =
                java.time.LocalTime.now()
                        .format(java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss"));
        embeddedTabPollStatusLabel.setText("確認 " + time + "（" + aspect + "）");
        if (!embeddedTabPollStatusLabel.getStyleClass().contains("request-form-poll-active")) {
            embeddedTabPollStatusLabel.getStyleClass().add("request-form-poll-active");
        }
        if (pollStatusHighlightPause != null) {
            pollStatusHighlightPause.stop();
        }
        pollStatusHighlightPause =
                new javafx.animation.PauseTransition(javafx.util.Duration.seconds(2.5));
        pollStatusHighlightPause.setOnFinished(
                e -> {
                    if (embeddedTabPollStatusLabel == null) {
                        return;
                    }
                    embeddedTabPollStatusLabel.getStyleClass().remove("request-form-poll-active");
                    embeddedTabPollStatusLabel.setText("監視中（転記ロック 10秒 / 原本 30秒）");
                });
        pollStatusHighlightPause.play();
    }

    private void pollOriginalFilesInWorkspace() {
        File folder = new File(targetFolder);
        if (!folder.isDirectory()) {
            return;
        }
        File[] files = listOriginalWorkbooks(folder);
        if (files == null) {
            return;
        }
        for (File file : files) {
            originalUpdateMonitor.ensureTracked(file);
            originalUpdateMonitor.poll(file);
        }
        Platform.runLater(this::refreshPreviewFileHeader);
    }

    private void syncOriginalFileMonitorAfterReload() {
        File folder = new File(targetFolder);
        File[] files = listOriginalWorkbooks(folder);
        if (files == null) {
            return;
        }
        for (File file : files) {
            originalUpdateMonitor.ensureTracked(file);
        }
        pollOriginalFilesInWorkspace();
    }

    private static File[] listOriginalWorkbooks(File folder) {
        return folder.listFiles(
                (dir, name) ->
                        name.endsWith(".xlsm")
                                && !name.startsWith("~$")
                                && !name.equals("加工依頼書入力.xlsm"));
    }
    
    private void displayPreviewPdf(File pdfFile) {
        sheetGrid.getChildren().clear();
        try {
            java.awt.image.BufferedImage buffered =
                    RequestFormSheetPreviewPdfRenderer.renderFirstPageImage(pdfFile);
            javafx.scene.image.Image image = javafx.embed.swing.SwingFXUtils.toFXImage(buffered, null);
            if (image == null || image.getWidth() <= 0 || image.getHeight() <= 0) {
                throw new IllegalStateException("PDF プレビューのラスタライズに失敗しました");
            }
            InteractiveImageViewer viewer = new InteractiveImageViewer(image);

            viewer.prefWidthProperty().bind(sheetScrollPane.widthProperty().subtract(24));
            viewer.maxWidthProperty().bind(sheetScrollPane.widthProperty().subtract(24));

            sheetGrid.add(viewer, 0, 0);
            viewer.resetZoomToFit();
        } catch (Exception e) {
            Label lblErr = new Label("PDF プレビューの表示に失敗しました: " + e.getMessage());
            lblErr.getStyleClass().add("excel-grid-label-error");
            sheetGrid.add(lblErr, 0, 0);
        }
    }

    private void saveLocalForm() {
        if (rejectGuestMutation("手修正の一時保存")) {
            return;
        }
        if (selectedRecord == null) {
            showAlert("エラー", "修正するレコードを選択してください。");
            return;
        }
        
        Map<String, String> db = selectedRecord.getDbValues();
        db.clear();
        db.putAll(buildJuchuDbValuesFromForm());
        
        selectedRecord.setStatus(STATUS_LOCAL_SAVE_PENDING);
        syncOrderRecordSummaryFromDb(selectedRecord, db);
        
        // ComboBox表示のリフレッシュ
        comboRecord.setItems(null);
        comboRecord.setItems(orderRecords);
        comboRecord.getSelectionModel().select(selectedRecord);
        
        statusLabel.setText("修正をローカル保存しました（受注ファイルに反映させるには転記ボタンを押してください）");
    }

    /** 転記・一時保存後に {@link OrderRecord} の db と一覧表示用サマリをフォーム内容へ揃える。 */
    private void syncOrderRecordDbFromCurrentForm(OrderRecord record) {
        if (record == null) {
            return;
        }
        Map<String, String> db = record.getDbValues();
        if (db == null) {
            return;
        }
        Map<String, String> fromForm = buildJuchuDbValuesFromForm();
        db.clear();
        db.putAll(fromForm);
        syncOrderRecordSummaryFromDb(record, fromForm);
    }

    private static void syncOrderRecordSummaryFromDb(
            OrderRecord record, Map<String, String> db) {
        if (record == null || db == null) {
            return;
        }
        record.setUser(db.getOrDefault("ユーザー", ""));
        String product = db.getOrDefault("製品", "");
        int nl = product.indexOf('\n');
        if (nl >= 0) {
            product = product.substring(0, nl);
        }
        record.setProduct(product.trim());
    }

    private void transferToExcel() {
        if (rejectGuestMutation("受注ファイルへの自動転記")) {
            return;
        }
        if (juchuTransferInProgress) {
            return;
        }

        String reqNo = txtReqNo.getText().trim();
        if (reqNo.isEmpty()) {
            showAlert("エラー", "依頼Ｎｏを入力してください。");
            return;
        }

        if (isJuchuFileLockedForWrite()) {
            showJuchuWriteBlockedAlert();
            updateTransferButtonState();
            return;
        }

        File file = new File(juchuFilePath);
        if (!file.isFile()) {
            showAlert("エラー", "受注ファイルが見つかりません。");
            return;
        }

        List<String> headerWarnings = readJuchuHeaderWarnings(file);
        if (!headerWarnings.isEmpty()) {
            statusLabel.setText("列定義警告: " + headerWarnings.size() + " 件（詳細はダイアログ）");
        }
        if (!confirmJuchuHeaderWarnings(file, headerWarnings)) {
            return;
        }

        JuchuTransferFormData formData = captureJuchuTransferFormData(reqNo);
        OrderRecord recordRef = selectedRecord;

        beginJuchuTransfer(
                "受注ファイルへ転記しています…\n(1/5) 受注ファイルを開いています…\n依頼No: " + reqNo);
        statusLabel.setText("受注ファイル更新中...");

        Thread transferThread =
                new Thread(
                        () -> {
                            Exception error = null;
                            JuchuTransferUndoState undoState = null;
                            try {
                                undoState =
                                        performSingleJuchuTransfer(
                                                file, formData, this::updateLoadingOverlayText);
                            } catch (Exception ex) {
                                error = ex;
                            }

                            final Exception finalError = error;
                            final JuchuTransferUndoState finalUndo = undoState;
                            Platform.runLater(
                                    () -> {
                                        try {
                                            if (finalError != null) {
                                                statusLabel.setText("エラー: " + finalError.getMessage());
                                                showAlert(
                                                        "エラー",
                                                        "ファイルの書き込み中にエラーが発生しました:\n"
                                                                + finalError.getMessage());
                                                finalError.printStackTrace();
                                                return;
                                            }

                                            statusLabel.setText("受注ファイルの更新が完了しました。");
                                            showAlert("成功", "受注ファイルにデータが一括転記・保存されました！");
                                            rememberLastJuchuTransferUndo(finalUndo, recordRef);

                                            if (recordRef != null) {
                                                recordRef.setStatus("一致 (転記完了)");
                                                recordRef.setDiscrepancy("一致 (受注ファイルへ転記完了)");
                                                syncOrderRecordDbFromCurrentForm(recordRef);

                                                ObservableList<OrderRecord> currentItems =
                                                        comboRecord.getItems();
                                                comboRecord.setItems(null);
                                                comboRecord.setItems(currentItems);
                                                comboRecord.getSelectionModel().select(recordRef);

                                                discrepancyLabel.setText("一致 (受注ファイルへ転記完了)");
                                                discrepancyLabel
                                                        .getStyleClass()
                                                        .removeAll(
                                                                "discrepancy-label-info",
                                                                "discrepancy-label-matched",
                                                                "discrepancy-label-discrepancy",
                                                                "discrepancy-label-unregistered");
                                                discrepancyLabel
                                                        .getStyleClass()
                                                        .add("discrepancy-label-matched");
                                            }
                                        } finally {
                                            endJuchuTransfer();
                                        }
                                    });
                        },
                        "request-form-juchu-transfer");
        transferThread.setDaemon(true);
        transferThread.start();
    }

    /** フォーム入力値のスナップショット（転記はバックグラウンドスレッドで実行）。 */
    private record JuchuTransferFormData(
            String reqNo, Map<String, String> dbValues, Map<String, String> rawValuesForNewRow) {}

    /** 直前の単票自動転記を Excel 上で戻すための行スナップショット。 */
    private record JuchuTransferUndoState(
            String reqNo,
            String juchuFilePath,
            boolean insertedNewRow,
            int rowIndex0,
            Map<String, String> priorDbValues,
            String priorRecordStatus,
            String priorRecordDiscrepancy) {}

    private JuchuTransferFormData captureJuchuTransferFormData(String reqNo) {
        Map<String, String> rawValues =
                selectedRecord != null
                        ? new HashMap<>(selectedRecord.getRawValues())
                        : Map.of();
        return new JuchuTransferFormData(reqNo, buildJuchuDbValuesFromForm(), rawValues);
    }

    private List<String> readJuchuHeaderWarnings(File file) {
        try (FileInputStream fis = new FileInputStream(file);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
            if (sheet == null) {
                return List.of("受注ﾌｧｲﾙ シートが見つかりません。");
            }
            Row hRow = sheet.getRow(juchuHeaderRowIndexFor(file.getAbsolutePath()));
            return JuchuSheetColumnLayout.validateHeaders(
                    hRow, juchuHeaderAliasRegistry, file.getAbsolutePath());
        } catch (Exception ex) {
            return List.of("受注ファイルの見出し検証に失敗: " + ex.getMessage());
        }
    }

    private List<JuchuHeaderMismatch> readJuchuHeaderMismatches(File file) {
        try {
            return JuchuSheetHeaderRepairWizard.readMismatches(file, juchuHeaderAliasRegistry);
        } catch (Exception ex) {
            return List.of();
        }
    }

    private JuchuTransferUndoState performSingleJuchuTransfer(
            File file, JuchuTransferFormData form, Consumer<String> progress) throws Exception {
        progress.accept(
                "受注ファイルへ転記しています…\n(1/5) 受注ファイルを開いています…\n依頼No: " + form.reqNo());

        JUCHU_TRANSFER_REPLACE_FORMULA.set(Boolean.TRUE);
        try (Workbook wb = PoiWorkbookOpener.open(file)) {
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");

            Map<String, Integer> colMap = buildJuchuColumnMap(sheet, file.getAbsolutePath());

            progress.accept(
                    "受注ファイルへ転記しています…\n(2/5) 依頼No を検索しています…\n依頼No: " + form.reqNo());

            String normKey = normalize_key(form.reqNo());
            int lastDataRowIndex = findJuchuSheetLastPopulatedDataRowIndex(sheet);
            Row targetRow = null;
            int targetRowIndex = -1;
            int matchedRowIndex = findJuchuDataRowIndexByReqNo(sheet, normKey, lastDataRowIndex);
            if (matchedRowIndex >= 0) {
                targetRow = sheet.getRow(matchedRowIndex);
                targetRowIndex = matchedRowIndex + 1;
            }

            boolean isNewRow = targetRow == null;
            JuchuTransferUndoState undoState =
                    captureJuchuTransferUndoState(
                            file,
                            form.reqNo(),
                            targetRow,
                            matchedRowIndex,
                            isNewRow,
                            lastDataRowIndex);

            if (isNewRow) {
                progress.accept(
                        "受注ファイルへ転記しています…\n(2/5) 最終行の次に新規行を追加しています…\n依頼No: "
                                + form.reqNo());
                int destRowIdx =
                        insertNewJuchuDataRowPreservingFormulas(sheet, lastDataRowIndex, progress);
                targetRow = sheet.getRow(destRowIdx);
                targetRowIndex = destRowIdx + 1;
                applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);
            }

            progress.accept(
                    "受注ファイルへ転記しています…\n(3/5) セルへ転記しています…\n依頼No: "
                            + form.reqNo()
                            + (isNewRow ? "（新規行・末尾追加）" : "（既存行）"));

            setJuchuSheetReqNoIfIncluded(wb, sheet, targetRow, form.reqNo());

            Map<String, String> mergedDb = new LinkedHashMap<>(form.dbValues());
            Map<String, String> raw = form.rawValuesForNewRow();
            if (raw != null && !raw.isEmpty()) {
                if (isNewRow) {
                    mergeJuchuDbFromRawDefaults(mergedDb, raw);
                } else {
                    mergeJuchuContractNoFromRawWhenBlankOrDifferent(mergedDb, raw);
                }
            }

            writeJuchuRowFromValues(
                    targetRow,
                    mergedDb,
                    isNewRow,
                    mergedDb.get("入力区分"),
                    mergedDb.get("加工区分"),
                    mergedDb.get("入力担当"),
                    mergedDb.get("特記事項1"),
                    mergedDb.get("特記事項2"),
                    mergedDb.get("特記事項3"),
                    isNewRow,
                    true);

            if (isNewRow) {
                progress.accept(
                        "受注ファイルへ転記しています…\n(4/5) 数式列は保全したまま転記完了…\n依頼No: "
                                + form.reqNo());
            } else {
                progress.accept(
                        "受注ファイルへ転記しています…\n(4/5) 既存行の更新を確定しています…\n依頼No: "
                                + form.reqNo());
            }

            progress.accept(
                    "受注ファイルへ転記しています…\n(5/5) ファイルを保存しています…\n依頼No: "
                            + form.reqNo());
            saveJuchuWorkbookWithLocalBackup(file, wb, "single-transfer");
            return undoState;
        } finally {
            JUCHU_TRANSFER_REPLACE_FORMULA.remove();
        }
    }

    private JuchuTransferUndoState captureJuchuTransferUndoState(
            File file,
            String reqNo,
            Row targetRow,
            int matchedRowIndex,
            boolean willInsertNewRow,
            int lastDataRowIndex0) {
        String path = file.getAbsolutePath();
        int firstDataRow = juchuFirstDataRowIndexFor(path);
        if (willInsertNewRow) {
            int destRowIdx =
                    lastDataRowIndex0 >= firstDataRow ? lastDataRowIndex0 + 1 : firstDataRow;
            return new JuchuTransferUndoState(
                    reqNo, path, true, destRowIdx, Map.of(), null, null);
        }
        Map<String, String> prior =
                targetRow != null
                        ? new LinkedHashMap<>(
                                JuchuSheetColumnLayout.readDbValuesFromRow(
                                        targetRow, juchuHeaderAliasRegistry, path))
                        : Map.of();
        return new JuchuTransferUndoState(
                prior.getOrDefault("依頼No", prior.getOrDefault("依頼Ｎｏ", "")),
                path,
                false,
                matchedRowIndex,
                prior,
                null,
                null);
    }

    private void undoLastJuchuTransfer() {
        if (rejectGuestMutation("直前の自動転記の取り消し")) {
            return;
        }
        if (juchuTransferInProgress || lastJuchuTransferUndo == null) {
            return;
        }
        if (resolveTransferBlockedReason() != null) {
            showAlert("エラー", resolveTransferBlockedReason());
            updateTransferButtonState();
            return;
        }
        JuchuTransferUndoState undo = lastJuchuTransferUndo;
        File file = new File(juchuFilePath);
        if (!file.isFile()) {
            showAlert("エラー", "受注ファイルが見つかりません。");
            return;
        }
        if (!file.getAbsolutePath().equals(undo.juchuFilePath())) {
            showAlert(
                    "エラー",
                    "受注ファイルが転記時と異なります。転記時: "
                            + undo.juchuFilePath()
                            + "\n現在: "
                            + file.getAbsolutePath());
            return;
        }

        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("自動転記の取り消し");
        confirm.setHeaderText("直前の自動転記を取り消しますか？");
        confirm.setContentText(
                "依頼No: "
                        + undo.reqNo()
                        + (undo.insertedNewRow()
                                ? "\n受注ファイルで新規挿入した行を削除します。"
                                : "\n受注ファイルの該当行を転記前の内容に戻します。"));
        if (confirm.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) {
            return;
        }

        OrderRecord recordRef = selectedRecord;
        beginJuchuTransfer(
                "直前の自動転記を取り消しています…\n依頼No: " + undo.reqNo());
        statusLabel.setText("直前の自動転記を取り消し中...");

        Thread undoThread =
                new Thread(
                        () -> {
                            Exception error = null;
                            try {
                                performJuchuTransferUndo(file, undo, this::updateLoadingOverlayText);
                            } catch (Exception ex) {
                                error = ex;
                            }
                            final Exception finalError = error;
                            Platform.runLater(
                                    () -> {
                                        try {
                                            if (finalError != null) {
                                                statusLabel.setText(
                                                        "取り消しエラー: " + finalError.getMessage());
                                                showAlert(
                                                        "エラー",
                                                        "自動転記の取り消しに失敗しました:\n"
                                                                + finalError.getMessage());
                                                finalError.printStackTrace();
                                                return;
                                            }
                                            clearLastJuchuTransferUndo();
                                            statusLabel.setText("直前の自動転記を取り消しました。");
                                            showAlert("完了", "直前の自動転記を取り消しました。");
                                            if (recordRef != null
                                                    && undo.reqNo() != null
                                                    && !undo.reqNo().isBlank()
                                                    && normalize_key(recordRef.getReqNo())
                                                            .equals(normalize_key(undo.reqNo()))) {
                                                if (undo.priorRecordStatus() != null) {
                                                    recordRef.setStatus(undo.priorRecordStatus());
                                                }
                                                if (undo.priorRecordDiscrepancy() != null) {
                                                    recordRef.setDiscrepancy(
                                                            undo.priorRecordDiscrepancy());
                                                    discrepancyLabel.setText(
                                                            undo.priorRecordDiscrepancy());
                                                }
                                                refreshComboRecordItems();
                                                comboRecord.getSelectionModel().select(recordRef);
                                            }
                                            requestReloadData(
                                                    "取り消し後、一覧を再読込します。", null);
                                        } finally {
                                            endJuchuTransfer();
                                        }
                                    });
                        },
                        "request-form-juchu-undo");
        undoThread.setDaemon(true);
        undoThread.start();
    }

    private void performJuchuTransferUndo(
            File file, JuchuTransferUndoState undo, Consumer<String> progress) throws Exception {
        progress.accept("自動転記の取り消し…\n(1/3) 受注ファイルを開いています…");
        JUCHU_TRANSFER_REPLACE_FORMULA.set(Boolean.TRUE);
        try (Workbook wb = PoiWorkbookOpener.open(file)) {
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
            if (sheet == null) {
                throw new IllegalStateException("受注ﾌｧｲﾙ シートが見つかりません。");
            }
            if (undo.insertedNewRow()) {
                progress.accept("自動転記の取り消し…\n(2/3) 新規挿入行を削除しています…");
                removeJuchuDataRowAt(sheet, undo.rowIndex0());
            } else {
                progress.accept("自動転記の取り消し…\n(2/3) 転記前の内容へ戻しています…");
                Row row = sheet.getRow(undo.rowIndex0());
                if (row == null) {
                    throw new IllegalStateException(
                            "取り消し対象行が見つかりません（行 "
                                    + (undo.rowIndex0() + 1)
                                    + "）。");
                }
                Map<String, String> prior = undo.priorDbValues();
                setJuchuSheetReqNoIfIncluded(wb, sheet, row, undo.reqNo());
                writeJuchuRowFromValues(
                        row,
                        prior,
                        false,
                        prior.getOrDefault("入力区分", ""),
                        prior.getOrDefault("加工区分", ""),
                        prior.getOrDefault("入力担当", ""),
                        prior.getOrDefault("特記事項1", ""),
                        prior.getOrDefault("特記事項2", ""),
                        prior.getOrDefault("特記事項3", ""),
                        false,
                        false);
            }
            progress.accept("自動転記の取り消し…\n(3/3) ファイルを保存しています…");
            saveJuchuWorkbookWithLocalBackup(file, wb, "undo-transfer");
        } finally {
            JUCHU_TRANSFER_REPLACE_FORMULA.remove();
        }
    }

    private void removeJuchuDataRowAt(Sheet sheet, int rowIndex0) {
        int last = findJuchuSheetLastPopulatedDataRowIndex(sheet);
        int firstDataRow = juchuFirstDataRowIndex0();
        if (rowIndex0 < firstDataRow || rowIndex0 > last) {
            return;
        }
        if (rowIndex0 < last) {
            sheet.shiftRows(rowIndex0 + 1, last, -1, true, true);
        } else {
            Row row = sheet.getRow(rowIndex0);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    /** 製品行の「商品」欄（{@link ProductRow#txtItem}）を改行連結。 */
    private static String joinMasterBaseShohinFromProductRows(List<ProductRow> rows) {
        List<String> lines = new ArrayList<>();
        for (ProductRow r : rows) {
            String t = r.txtItem.getText().trim();
            if (!t.isEmpty()) {
                lines.add(t);
            }
        }
        return String.join("\n", lines);
    }

    private String joinMasterBaseShohinFromProductRows() {
        return joinMasterBaseShohinFromProductRows(productRows);
    }

    /** 原反行の「商品」欄（{@link RawMaterialRow#txtItem}）を改行連結。 */
    private static String joinMasterBaseShohinFromRawRows(List<RawMaterialRow> rows) {
        List<String> lines = new ArrayList<>();
        for (RawMaterialRow r : rows) {
            String t = r.txtItem.getText().trim();
            if (!t.isEmpty()) {
                lines.add(t);
            }
        }
        return String.join("\n", lines);
    }

    private String joinMasterBaseShohinFromRawRows() {
        return joinMasterBaseShohinFromRawRows(rawRows);
    }

    /**
     * 受注ﾌｧｲﾙの実データ末尾の次行へ新規行を追加する（全行シフトなし）。データが無いときのみ先頭データ行を使う。
     *
     * @param lastPopulatedDataRowIndex {@link #findJuchuSheetLastPopulatedDataRowIndex} の結果（実データ末尾）
     * @param progress 進捗表示（{@code null} 可）
     * @return 挿入した行の 0-based index
     */
    private int insertNewJuchuDataRowPreservingFormulas(
            Sheet sheet, int lastPopulatedDataRowIndex, Consumer<String> progress) {
        final int firstDataRow = juchuFirstDataRowIndex0();
        final int destRowIdx;
        final int templateRowIdx;
        if (lastPopulatedDataRowIndex >= firstDataRow) {
            destRowIdx = lastPopulatedDataRowIndex + 1;
            templateRowIdx = lastPopulatedDataRowIndex;
        } else {
            destRowIdx = firstDataRow;
            templateRowIdx = resolveJuchuFormulaTemplateRowIndex(sheet, destRowIdx);
        }
        if (progress != null) {
            progress.accept(
                    "受注ファイルへ転記しています…\n(2/5) テンプレート行から数式・書式を複製しています…");
        }
        if (sheet.getRow(destRowIdx) == null) {
            sheet.createRow(destRowIdx);
        }
        copyJuchuTemplateRowFormulasAndStyles(sheet, templateRowIdx, destRowIdx);
        return destRowIdx;
    }

    /** データ行が無いとき、数式セルが最も多い行をテンプレートとして選ぶ（挿入先自身は除外）。 */
    private static int resolveJuchuFormulaTemplateRowIndex(Sheet sheet, int destRowIdx) {
        int best = destRowIdx + 1;
        int bestScore = -1;
        for (int r = destRowIdx + 1; r <= destRowIdx + 2; r++) {
            int score = countJuchuFormulaCells(sheet.getRow(r));
            if (score > bestScore) {
                bestScore = score;
                best = r;
            }
        }
        return best;
    }

    private static int countJuchuFormulaCells(Row row) {
        if (row == null) {
            return 0;
        }
        int count = 0;
        for (int c = 0; c < row.getLastCellNum(); c++) {
            if (isJuchuFormulaCell(row.getCell(c))) {
                count++;
            }
        }
        return count;
    }

    /** 受注ﾌｧｲﾙ A 列（依頼No）に実データがある行か（数式の 0 や空文字は除外）。 */
    private boolean juchuRowHasReqNo(Row row) {
        if (row == null) {
            return false;
        }
        Cell cell = row.getCell(0);
        if (cell == null) {
            return false;
        }
        return switch (cell.getCellType()) {
            case STRING -> !cell.getStringCellValue().strip().isEmpty();
            case NUMERIC -> true;
            case FORMULA -> {
                try {
                    yield !cell.getStringCellValue().strip().isEmpty();
                } catch (Exception ex) {
                    try {
                        yield cell.getNumericCellValue() != 0.0d;
                    } catch (Exception ignored) {
                        yield false;
                    }
                }
            }
            default -> false;
        };
    }

    /**
     * 受注ﾌｧｲﾙの実データ末尾行（0-based）。後方走査で A 列に依頼No がある最終行を返す。
     */
    private int findJuchuSheetLastPopulatedDataRowIndex(Sheet sheet) {
        int first = juchuFirstDataRowIndex0();
        int poiLast = sheet.getLastRowNum();
        if (poiLast < first) {
            return first - 1;
        }
        int scanLimit = Math.min(poiLast, first + JUCHU_SHEET_MAX_SCAN_ROWS);
        for (int r = scanLimit; r >= first; r--) {
            if (juchuRowHasReqNo(sheet.getRow(r))) {
                return r;
            }
        }
        return first - 1;
    }

    /** 依頼No に一致するデータ行 index（0-based）。見つからなければ {@code -1}。 */
    private int findJuchuDataRowIndexByReqNo(Sheet sheet, String normKey, int lastDataRowIndex) {
        int firstDataRow = juchuFirstDataRowIndex0();
        if (lastDataRowIndex < firstDataRow) {
            return -1;
        }
        for (int r = firstDataRow; r <= lastDataRowIndex; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(0);
            if (cell != null && normalize_key(getCellValueAsString(cell)).equals(normKey)) {
                return r;
            }
        }
        return -1;
    }

    /** テンプレート行の数式（行参照は POI が補正）と書式のみ複製。 */
    private static void copyJuchuTemplateRowFormulasAndStyles(Sheet sheet, int srcRowIdx, int destRowIdx) {
        Row srcRow = sheet.getRow(srcRowIdx);
        if (srcRow == null) {
            if (sheet.getRow(destRowIdx) == null) {
                sheet.createRow(destRowIdx);
            }
            return;
        }
        CellCopyPolicy policy =
                new CellCopyPolicy.Builder()
                        .cellFormula(true)
                        .cellStyle(true)
                        .cellValue(false)
                        .mergedRegions(true)
                        .rowHeight(true)
                        .build();
        if (sheet instanceof XSSFSheet xssfSheet) {
            xssfSheet.copyRows(srcRowIdx, srcRowIdx, destRowIdx, policy);
            return;
        }
        Row destRow = sheet.getRow(destRowIdx);
        if (destRow == null) {
            destRow = sheet.createRow(destRowIdx);
        }
        for (int c = 0; c < srcRow.getLastCellNum(); c++) {
            Cell srcCell = srcRow.getCell(c);
            if (srcCell == null) {
                continue;
            }
            Cell destCell = destRow.getCell(c);
            if (destCell == null) {
                destCell = destRow.createCell(c);
            }
            destCell.setCellStyle(srcCell.getCellStyle());
            if (srcCell.getCellType() == CellType.FORMULA) {
                destCell.setCellFormula(srcCell.getCellFormula());
            }
        }
    }

    private static boolean isJuchuFormulaCell(Cell cell) {
        return cell != null && cell.getCellType() == CellType.FORMULA;
    }

    /** 依頼No（A列）のみ値として書き込む（他列の数式セルは触らない）。 */
    private void setJuchuSheetReqNo(Workbook wb, Sheet sheet, Row targetRow, String reqNo) {
        if (targetRow == null || reqNo == null) {
            return;
        }
        Cell cellReqNo = targetRow.getCell(0);
        if (cellReqNo == null) {
            cellReqNo = targetRow.createCell(0);
        }
        cellReqNo.setCellValue(reqNo);
        CellStyle yellowStyle = wb.createCellStyle();
        int firstDataRow = juchuFirstDataRowIndex0();
        Row refRow = sheet.getRow(firstDataRow + 1);
        if (refRow == null) {
            refRow = sheet.getRow(firstDataRow);
        }
        if (refRow != null && refRow.getCell(0) != null) {
            yellowStyle.cloneStyleFrom(refRow.getCell(0).getCellStyle());
        }
        yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellReqNo.setCellStyle(yellowStyle);
    }

    private static boolean juchuTransferReplaceFormulaCells() {
        return Boolean.TRUE.equals(JUCHU_TRANSFER_REPLACE_FORMULA.get());
    }

    /** 通常は数式セルを触らない。自動転記中は数式を値セルへ置換してフォーム内容を反映する。 */
    private static Cell writableJuchuCell(Row row, int col) {
        Cell existing = row.getCell(col);
        if (isJuchuFormulaCell(existing)) {
            if (!juchuTransferReplaceFormulaCells()) {
                return null;
            }
            if (existing instanceof org.apache.poi.xssf.usermodel.XSSFCell xssfCell) {
                xssfCell.setBlank();
                return xssfCell;
            }
            if (existing != null) {
                row.removeCell(existing);
            }
            return row.createCell(col);
        }
        return existing != null ? existing : row.createCell(col);
    }

    private static void setJuchuSheetString(
            Row row, Map<String, Integer> colMap, String header, String value) {
        Integer col = colMap.get(header);
        if (col == null) {
            return;
        }
        Cell cell = writableJuchuCell(row, col);
        if (cell == null) {
            return;
        }
        cell.setCellValue(value != null ? value : "");
    }

    private static void setJuchuSheetNumericOrString(
            Row row, Map<String, Integer> colMap, String header, String text) {
        if (text == null || text.isBlank()) {
            setJuchuSheetString(row, colMap, header, "");
            return;
        }
        try {
            Integer col = colMap.get(header);
            if (col == null) {
                return;
            }
            Cell cell = writableJuchuCell(row, col);
            if (cell == null) {
                return;
            }
            cell.setCellValue(Double.parseDouble(text.trim()));
        } catch (Exception e) {
            setJuchuSheetString(row, colMap, header, text);
        }
    }

    private static void setJuchuSheetDate(
            Row row, Map<String, Integer> colMap, String header, Date value) {
        Integer col = colMap.get(header);
        if (col == null || value == null) {
            return;
        }
        Cell cell = writableJuchuCell(row, col);
        if (cell == null) {
            return;
        }
        cell.setCellValue(value);
    }

    /** テンプレートに無い列向けの既定数式（既に数式があれば触らない）。 */
    private static void applyDefaultJuchuFormulasIfMissing(
            Row row, Map<String, Integer> colMap, int excelRow1Based) {
        setJuchuFormulaIfMissing(row, colMap, "月数", "MONTH(AF" + excelRow1Based + ")");
        setJuchuFormulaIfMissing(row, colMap, "受注金額", "AI" + excelRow1Based + "*AH" + excelRow1Based);
        setJuchuFormulaIfMissing(row, colMap, "受注数", "M" + excelRow1Based);
        setJuchuFormulaIfMissing(row, colMap, "単価", "AH" + excelRow1Based);
    }

    private static void setJuchuFormulaIfMissing(
            Row row, Map<String, Integer> colMap, String header, String formula) {
        Integer col = colMap.get(header);
        if (col == null || formula == null || formula.isBlank()) {
            return;
        }
        Cell cell = row.getCell(col);
        if (isJuchuFormulaCell(cell)) {
            return;
        }
        if (cell == null) {
            cell = row.createCell(col);
        }
        cell.setCellFormula(formula);
    }

    private static void setJuchuSheetCellIfPresent(
            Row row, Map<String, Integer> colMap, String columnHeader, String value) {
        setJuchuSheetString(row, colMap, columnHeader, value);
    }

    private boolean isJuchuColumnExcluded(JuchuSheetColumnLayout.Col col) {
        return col != null
                && juchuFilePath != null
                && !juchuFilePath.isBlank()
                && juchuHeaderAliasRegistry.isExcludedFromTransfer(juchuFilePath, col);
    }

    private void setJuchuCellIfIncluded(
            Row row, JuchuSheetColumnLayout.Col col, String value) {
        if (!isJuchuColumnExcluded(col)) {
            setJuchuCellByLayout(row, col, value);
        }
    }

    private void setJuchuNumericOrStringIfIncluded(
            Row row, JuchuSheetColumnLayout.Col col, String text) {
        if (!isJuchuColumnExcluded(col)) {
            setJuchuNumericOrStringByLayout(row, col, text);
        }
    }

    private void setJuchuDateOrStringIfIncluded(
            Row row, JuchuSheetColumnLayout.Col col, String rawValue) {
        if (!isJuchuColumnExcluded(col)) {
            setJuchuDateOrStringByLayout(row, col, rawValue);
        }
    }

    private void applyJuchuNyuryokuBiFromDbIfIncluded(
            Row targetRow, Map<String, String> db, boolean fallbackToTodayIfBlank) {
        if (!isJuchuColumnExcluded(JuchuSheetColumnLayout.Col.NYURYOKU_BI)) {
            applyJuchuNyuryokuBiFromDb(targetRow, db, fallbackToTodayIfBlank);
        }
    }

    private void setJuchuSheetReqNoIfIncluded(
            Workbook wb, Sheet sheet, Row targetRow, String reqNo) {
        if (!isJuchuColumnExcluded(JuchuSheetColumnLayout.Col.IRAI_NO)) {
            setJuchuSheetReqNo(wb, sheet, targetRow, reqNo);
        }
    }

    private static void setJuchuCellByLayout(Row row, JuchuSheetColumnLayout.Col col, String value) {
        if (row == null || col == null) {
            return;
        }
        Cell cell = writableJuchuCell(row, col.columnIndex());
        if (cell == null) {
            return;
        }
        cell.setCellValue(value != null ? value : "");
    }

    private static void setJuchuNumericOrStringByLayout(
            Row row, JuchuSheetColumnLayout.Col col, String text) {
        if (row == null || col == null) {
            return;
        }
        if (text == null || text.isBlank()) {
            setJuchuCellByLayout(row, col, "");
            return;
        }
        try {
            Cell cell = writableJuchuCell(row, col.columnIndex());
            if (cell == null) {
                return;
            }
            cell.setCellValue(Double.parseDouble(text.trim()));
        } catch (Exception e) {
            setJuchuCellByLayout(row, col, text);
        }
    }

    /** 受注ファイルの「入力日」列。転記時は db の値で上書き（空のときのみ新規行で本日）。 */
    private static void applyJuchuNyuryokuBiFromDb(
            Row targetRow, Map<String, String> db, boolean fallbackToTodayIfBlank) {
        if (targetRow == null) {
            return;
        }
        String nyuryokuBi = db != null ? db.getOrDefault("入力日", "") : "";
        if (nyuryokuBi != null && !nyuryokuBi.isBlank()) {
            setJuchuDateOrStringByLayout(targetRow, JuchuSheetColumnLayout.Col.NYURYOKU_BI, nyuryokuBi);
        } else if (fallbackToTodayIfBlank) {
            setJuchuDateByLayout(targetRow, JuchuSheetColumnLayout.Col.NYURYOKU_BI, new Date());
        }
    }

    private static void setJuchuDateByLayout(Row row, JuchuSheetColumnLayout.Col col, Date value) {
        if (row == null || col == null || value == null) {
            return;
        }
        Cell cell = writableJuchuCell(row, col.columnIndex());
        if (cell == null) {
            return;
        }
        cell.setCellValue(value);
    }

    private static void setJuchuDateOrStringByLayout(
            Row row, JuchuSheetColumnLayout.Col col, String rawValue) {
        if (rawValue == null || rawValue.isBlank()) {
            setJuchuCellByLayout(row, col, "");
            return;
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            setJuchuDateByLayout(row, col, sdf.parse(rawValue.strip()));
        } catch (Exception ex) {
            setJuchuCellByLayout(row, col, rawValue.strip());
        }
    }

    private static void updateProductRowSpecDisplay(ProductRow pRow) {
        if (pRow == null || pRow.txtSeihinmei == null) {
            return;
        }
        pRow.txtSeihinmei.setText(
                JuchuSheetColumnLayout.buildSpecName(
                        pRow.txtPart.getText(),
                        pRow.txtType.getText(),
                        pRow.txtWidth.getText(),
                        pRow.txtLength.getText()));
    }

    private static void updateRawRowSpecDisplay(RawMaterialRow rRow) {
        if (rRow == null || rRow.txtGenpanmei == null) {
            return;
        }
        rRow.txtGenpanmei.setText(
                JuchuSheetColumnLayout.buildSpecName(
                        rRow.txtPart.getText(),
                        rRow.txtType.getText(),
                        rRow.txtWidth.getText(),
                        rRow.txtLength.getText()));
        updateRawRowRollCountDisplay(rRow);
    }

    private static void updateRawRowRollCountDisplay(RawMaterialRow rRow) {
        if (rRow == null || rRow.txtRollCount == null || rRow.rollCountManualOverride) {
            return;
        }
        java.util.OptionalInt rollCount =
                JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength(
                        rRow.txtQty.getText(), rRow.txtLength.getText());
        rRow.suppressRollCountAutoListener = true;
        try {
            rRow.txtRollCount.setText(rollCount.isPresent() ? String.valueOf(rollCount.getAsInt()) : "");
        } finally {
            rRow.suppressRollCountAutoListener = false;
        }
    }

    private static void bindRawRowRollCountManualEdit(RawMaterialRow rRow) {
        if (rRow == null || rRow.txtRollCount == null) {
            return;
        }
        rRow.txtRollCount.textProperty().addListener((obs, oldV, newV) -> {
            if (rRow.suppressRollCountAutoListener) {
                return;
            }
            if (newV == null || newV.isBlank()) {
                rRow.rollCountManualOverride = false;
                updateRawRowRollCountDisplay(rRow);
                return;
            }
            rRow.rollCountManualOverride = true;
        });
    }

    private static String firstNonBlank(String... values) {
        for (String value : values) {
            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }
        return "";
    }

    private static String resolveWorkFieldValue(String formValue, String defaultValue) {
        if (formValue != null && !formValue.isBlank()) {
            return formValue.trim();
        }
        return defaultValue != null ? defaultValue : "";
    }

    private void openJuchuExcelExternally() {
        if (rejectGuestMutation("受注 Excel を開く")) {
            return;
        }
        File currentFile = new File(juchuFilePath);
        if (!currentFile.isFile()) {
            showAlert("エラー", "指定された受注ファイルが見つかりません:\n" + juchuFilePath);
            return;
        }
        try {
            if (java.awt.Desktop.isDesktopSupported()) {
                java.awt.Desktop.getDesktop().open(currentFile);
            } else {
                new ProcessBuilder("cmd", "/c", "start", "", juchuFilePath).start();
            }
            scheduleTransferButtonStateRefresh();
        } catch (Exception ex) {
            showAlert("エラー", "Excelファイルを開けませんでした: " + ex.getMessage());
        }
    }

    /** 受注ファイルのローカル世代バックアップフォルダをエクスプローラー等で開く。 */
    private void openJuchuBackupFolderExternally() {
        try {
            Path dir;
            if (juchuFilePath != null && !juchuFilePath.isBlank()) {
                dir =
                        RequestFormJuchuFileBackupStore.resolveSourceBackupsRoot(
                                uiEnvSnapshot, Path.of(juchuFilePath.trim()));
            } else {
                dir = RequestFormJuchuFileBackupStore.resolveBackupsRoot(uiEnvSnapshot);
            }
            Files.createDirectories(dir);
            if (!java.awt.Desktop.isDesktopSupported()
                    || !java.awt.Desktop.getDesktop().isSupported(java.awt.Desktop.Action.OPEN)) {
                showAlert("エラー", "この環境ではフォルダを開けません。");
                return;
            }
            java.awt.Desktop.getDesktop().open(dir.toFile());
        } catch (Exception ex) {
            showAlert("エラー", "バックアップフォルダを開けませんでした: " + ex.getMessage());
        }
    }

    private void refreshJuchuBackupList() {
        if (juchuBackupListView == null) {
            return;
        }
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            juchuBackupListView.getItems().clear();
            if (juchuBackupHintLabel != null) {
                juchuBackupHintLabel.setText("受注ファイルが未設定のため、バックアップ一覧を表示できません。");
            }
            return;
        }
        Path target = Path.of(juchuFilePath.trim());
        List<RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry> entries =
                RequestFormJuchuFileBackupStore.loadIndexForSource(uiEnvSnapshot, target);
        juchuBackupListView.getItems().setAll(entries);
        if (juchuBackupHintLabel != null) {
            Path root = RequestFormJuchuFileBackupStore.resolveSourceBackupsRoot(uiEnvSnapshot, target);
            juchuBackupHintLabel.setText(
                    "保存先: "
                            + root
                            + "　保持上限 "
                            + RequestFormJuchuFileBackupStore.MAX_GENERATIONS_PER_SOURCE
                            + " 世代（15 分間隔で自動退避）");
        }
    }

    private void restoreSelectedJuchuBackupFromList() {
        if (rejectGuestMutation("受注ファイルのバックアップ復元")) {
            return;
        }
        if (juchuBackupListView == null) {
            return;
        }
        RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry sel =
                juchuBackupListView.getSelectionModel().getSelectedItem();
        if (sel == null) {
            showAlert("復元", "復元するバックアップを一覧から選んでください。");
            return;
        }
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            showAlert("エラー", "受注ファイルが未設定です。");
            return;
        }
        if (isJuchuFileLockedForWrite()) {
            showJuchuWriteBlockedAlert();
            return;
        }
        if (juchuTransferInProgress) {
            showAlert("エラー", "受注ファイルへの転記処理を実行中です。完了までお待ちください。");
            return;
        }

        File targetFile = new File(juchuFilePath);
        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("復元の確認");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "選択したバックアップ（"
                        + sel.displayLabel()
                        + "）で受注ファイルを上書きします。\n"
                        + "復元前に現行ファイルは自動で退避されます。\n\n"
                        + "対象: "
                        + targetFile.getAbsolutePath()
                        + "\n\n続行しますか？");
        if (hostWindow != null) {
            confirm.initOwner(hostWindow);
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }

        statusLabel.setText("バックアップから受注ファイルを復元しています…");
        Thread restoreThread =
                new Thread(
                        () -> {
                            Exception error = null;
                            try {
                                RequestFormJuchuFileBackupStore.restoreToSourceWorkbook(
                                        sel, uiEnvSnapshot, targetFile.toPath());
                            } catch (Exception ex) {
                                error = ex;
                            }
                            final Exception finalError = error;
                            Platform.runLater(
                                    () -> {
                                        if (finalError != null) {
                                            statusLabel.setText(
                                                    "エラー: " + finalError.getMessage());
                                            showAlert(
                                                    "エラー",
                                                    "バックアップからの復元に失敗しました:\n"
                                                            + finalError.getMessage());
                                            finalError.printStackTrace();
                                            return;
                                        }
                                        statusLabel.setText("バックアップから受注ファイルを復元しました。");
                                        refreshJuchuBackupList();
                                        showAlert(
                                                "成功",
                                                "バックアップから受注ファイルを復元しました。\nデータを再読込します。");
                                        requestReloadData(
                                                "バックアップ復元後、データを再読込します。", null);
                                    });
                        },
                        "request-form-juchu-restore");
        restoreThread.setDaemon(true);
        restoreThread.start();
    }

    private void openJuchuColumnDefinitionWizard() {
        if (rejectGuestMutation("列定義ウィザード")) {
            return;
        }
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            showAlert("エラー", "受注ファイルが未設定です。");
            return;
        }
        File file = new File(juchuFilePath);
        if (!file.isFile()) {
            showAlert("エラー", "受注ファイルが見つかりません:\n" + juchuFilePath);
            return;
        }
        juchuHeaderAliasRegistry.reloadFromDisk();
        JuchuSheetHeaderRepairWizard.showManage(hostWindow, file, juchuHeaderAliasRegistry);
    }

    private boolean confirmJuchuHeaderWarnings(File juchuFile, List<String> warnings) {
        if (warnings == null || warnings.isEmpty()) {
            return true;
        }
        if (juchuFile != null && juchuFile.isFile()) {
            List<JuchuHeaderMismatch> mismatches = readJuchuHeaderMismatches(juchuFile);
            if (!mismatches.isEmpty()) {
                JuchuSheetHeaderRepairWizard.Result wizardResult =
                        JuchuSheetHeaderRepairWizard.showTransferPrompt(
                                hostWindow,
                                juchuFile,
                                mismatches,
                                juchuHeaderAliasRegistry);
                return switch (wizardResult) {
                    case CANCEL -> false;
                    case CONTINUE -> true;
                    case FIXED -> true;
                };
            }
        }
        return showSimpleJuchuHeaderWarningDialog(warnings);
    }

    private boolean showSimpleJuchuHeaderWarningDialog(List<String> warnings) {
        Alert alert = new Alert(Alert.AlertType.WARNING);
        if (hostWindow != null) {
            alert.initOwner(hostWindow);
        }
        alert.setTitle("受注シート列定義の警告");
        alert.setHeaderText(
                "見出し行（行"
                        + juchuHeaderAliasRegistry.headerRowOneBasedFor(juchuFilePath)
                        + "）と列位置の定義が一致しない列があります。");
        TextArea area = new TextArea(String.join("\n", warnings));
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefRowCount(Math.min(12, warnings.size() + 1));
        alert.getDialogPane().setContent(area);
        ButtonType proceed = new ButtonType("続行", ButtonBar.ButtonData.OK_DONE);
        ButtonType cancel = new ButtonType("中止", ButtonBar.ButtonData.CANCEL_CLOSE);
        alert.getButtonTypes().setAll(proceed, cancel);
        return alert.showAndWait().orElse(cancel) == proceed;
    }

    private void writeJuchuRowFromValues(
            Row targetRow,
            Map<String, String> db,
            boolean isNewRow,
            String inputKbn,
            String kakoKbn,
            String inputTanto,
            String tokki1,
            String tokki2,
            String tokki3,
            boolean setInputDateNow,
            boolean useSessionInputTanto) {
        if (targetRow == null || db == null) {
            return;
        }
        setJuchuCellIfIncluded(
                targetRow,
                JuchuSheetColumnLayout.Col.NYURYOKU_KBN,
                resolveWorkFieldValue(inputKbn, isNewRow ? defaultInputKbnForNewRow() : ""));
        setJuchuCellIfIncluded(
                targetRow,
                JuchuSheetColumnLayout.Col.KAKO_KBN,
                resolveWorkFieldValue(kakoKbn, isNewRow ? defaultKakoKbnForNewRow() : ""));
        String resolvedInputTanto;
        if (useSessionInputTanto) {
            resolvedInputTanto = sessionInputTanto();
            if (resolvedInputTanto.isBlank()) {
                resolvedInputTanto =
                        resolveWorkFieldValue(inputTanto, isNewRow ? "自動転記" : "");
            }
        } else {
            resolvedInputTanto = resolveWorkFieldValue(inputTanto, isNewRow ? "自動転記" : "");
        }
        setJuchuCellIfIncluded(
                targetRow,
                JuchuSheetColumnLayout.Col.NYURYOKU_TANTO,
                resolvedInputTanto);
        applyJuchuNyuryokuBiFromDbIfIncluded(targetRow, db, setInputDateNow);

        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.HINMEI, db.getOrDefault("品名", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.SEIHIN, db.getOrDefault("製品", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KON_TO_1, db.getOrDefault("梱-等1", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.IRO_1, db.getOrDefault("色1", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KUBUN_1, db.getOrDefault("区分1", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.EDABAN, db.getOrDefault("枝番", ""));
        setJuchuNumericOrStringIfIncluded(targetRow, JuchuSheetColumnLayout.Col.SURYO_1, db.getOrDefault("数量1", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.EC_MEN, db.getOrDefault("ＥＣ面", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.TRIMMING, db.getOrDefault("ﾄﾘﾐﾝｸﾞ", ""));

        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.WARISU, db.getOrDefault("割数", ""));
        setJuchuCellIfIncluded(
                targetRow,
                JuchuSheetColumnLayout.Col.HINMEI_1,
                firstNonBlank(db.get("品名1"), db.get("原反品名")));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.GENPAN, db.getOrDefault("原反", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KON_TO, db.getOrDefault("原反梱-等", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.IRO, db.getOrDefault("原反色", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KUBUN, db.getOrDefault("原反区分", ""));
        setJuchuNumericOrStringIfIncluded(targetRow, JuchuSheetColumnLayout.Col.SURYO, db.getOrDefault("原反数量", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.ZAIKO_BASHO, db.getOrDefault("在庫場所", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.TONYU_BASHO, db.getOrDefault("投入場所", ""));
        setJuchuDateOrStringIfIncluded(targetRow, JuchuSheetColumnLayout.Col.TONYU_BI, db.getOrDefault("投入日", ""));

        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.UKETSUKE_NO, db.getOrDefault("受付Ｎｏ", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KAKO_NAIYO, db.getOrDefault("加工内容", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.TOKKI_1, tokki1 != null ? tokki1 : db.getOrDefault("特記事項1", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.TOKKI_2, tokki2 != null ? tokki2 : db.getOrDefault("特記事項2", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.TOKKI_3, tokki3 != null ? tokki3 : db.getOrDefault("特記事項3", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.YOTO, db.getOrDefault("用途", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.USER, db.getOrDefault("ユーザー", ""));
        setJuchuDateOrStringIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KIBO_NOKI, db.getOrDefault("希望納期", ""));
        setJuchuDateOrStringIfIncluded(targetRow, JuchuSheetColumnLayout.Col.CHOSEI_NOKI, db.getOrDefault("調整納期", ""));
        setJuchuNumericOrStringIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KAKOCHIN, db.getOrDefault("加工賃", ""));
        setJuchuCellIfIncluded(targetRow, JuchuSheetColumnLayout.Col.KEIYAKU_NO, db.getOrDefault("契約Ｎｏ", ""));
        setJuchuNumericOrStringIfIncluded(
                targetRow, JuchuSheetColumnLayout.Col.GENPAN_ROLL_SU, db.getOrDefault("原反ロール数", ""));

        setJuchuCellIfIncluded(
                targetRow,
                JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT,
                db.getOrDefault(COL_MASTER_BASE_SHOHIN_PRODUCT, ""));
        setJuchuCellIfIncluded(
                targetRow,
                JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_RAW,
                db.getOrDefault(COL_MASTER_BASE_SHOHIN_RAW, ""));
    }

    private Map<String, String> buildJuchuDbValuesFromForm() {
        Map<String, String> db = new LinkedHashMap<>();
        db.put("ユーザー", newCmbFormUser.getValue() != null ? newCmbFormUser.getValue().trim() : "");
        if (newDpFormDeliv.getValue() != null) {
            db.put("希望納期", newDpFormDeliv.getValue().toString());
        } else {
            db.put("希望納期", "");
        }
        if (newDpFormAdjustDeliv.getValue() != null) {
            db.put("調整納期", newDpFormAdjustDeliv.getValue().toString());
        } else {
            db.put("調整納期", "");
        }
        db.put("加工内容", txtProcess.getText().trim());
        db.put("加工賃", newTxtFormWage.getText().trim());
        db.put("受付Ｎｏ", newTxtUketsukeNo.getText().trim());
        db.put("用途", newCmbFormYoto.getValue() != null ? newCmbFormYoto.getValue().trim() : "");
        db.put(
                "入力区分",
                newCmbFormInputKbn.getValue() != null ? newCmbFormInputKbn.getValue().trim() : "");
        db.put(
                "加工区分",
                newCmbFormKakoKbn.getValue() != null ? newCmbFormKakoKbn.getValue().trim() : "");
        db.put("入力担当", sessionInputTanto());
        if (newDpFormInputDate.getValue() != null) {
            db.put("入力日", newDpFormInputDate.getValue().toString());
        } else {
            db.put("入力日", java.time.LocalDate.now().toString());
        }
        db.put("特記事項1", newTxtFormTokki1.getText().trim());
        db.put("特記事項2", newTxtFormTokki2.getText().trim());
        db.put("特記事項3", newTxtFormTokki3.getText().trim());

        List<String> prodHinmeis = new ArrayList<>();
        List<String> prodSpecs = new ArrayList<>();
        List<String> prodGrades = new ArrayList<>();
        List<String> prodColors = new ArrayList<>();
        List<String> prodCats = new ArrayList<>();
        List<String> prodQtys = new ArrayList<>();
        List<String> prodEdabans = new ArrayList<>();
        List<String> prodContracts = new ArrayList<>();
        List<String> prodEcs = new ArrayList<>();
        List<String> prodTrims = new ArrayList<>();
        List<String> prodMaster = new ArrayList<>();

        for (ProductRow pRow : productRows) {
            prodHinmeis.add(pRow.txtHinmei.getText().trim());
            prodSpecs.add(
                    firstNonBlank(
                            pRow.txtSeihinmei.getText(),
                            JuchuSheetColumnLayout.buildSpecName(
                                    pRow.txtPart.getText(),
                                    pRow.txtType.getText(),
                                    pRow.txtWidth.getText(),
                                    pRow.txtLength.getText())));
            prodMaster.add(pRow.txtItem.getText().trim());
            prodGrades.add(pRow.txtGrade.getText().trim());
            prodColors.add(pRow.txtColor.getText().trim());
            prodCats.add(pRow.txtCategory.getText().trim());
            prodQtys.add(pRow.txtQty.getText().trim());
            prodEdabans.add(pRow.txtEdaban.getText().trim());
            prodContracts.add(pRow.txtKeiyakuNo.getText().trim());
            prodEcs.add(pRow.cmbEcSide.getValue() != null ? pRow.cmbEcSide.getValue() : "");
            prodTrims.add(pRow.cmbTrimming.getValue() != null ? pRow.cmbTrimming.getValue() : "");
        }

        db.put("品名", String.join("\n", prodHinmeis));
        db.put("製品", String.join("\n", prodSpecs));
        db.put(COL_MASTER_BASE_SHOHIN_PRODUCT, String.join("\n", prodMaster));
        db.put("梱-等1", String.join("\n", prodGrades));
        db.put("色1", String.join("\n", prodColors));
        db.put("区分1", String.join("\n", prodCats));
        db.put("数量1", String.join("\n", prodQtys));
        db.put("枝番", String.join("\n", prodEdabans));
        db.put("契約Ｎｏ", String.join("\n", prodContracts));
        db.put("ＥＣ面", String.join("\n", prodEcs));
        db.put("ﾄﾘﾐﾝｸﾞ", String.join("\n", prodTrims));

        List<String> rawHinmeis = new ArrayList<>();
        List<String> rawSpecs = new ArrayList<>();
        List<String> rawGrades = new ArrayList<>();
        List<String> rawColors = new ArrayList<>();
        List<String> rawCats = new ArrayList<>();
        List<String> rawQtys = new ArrayList<>();
        List<String> wariVals = new ArrayList<>();
        List<String> feedVals = new ArrayList<>();
        List<String> storageVals = new ArrayList<>();
        List<String> inputDateVals = new ArrayList<>();
        List<String> rollCountVals = new ArrayList<>();
        List<String> rawMaster = new ArrayList<>();

        for (RawMaterialRow rRow : rawRows) {
            rawHinmeis.add(rRow.txtHinmei.getText().trim());
            rawSpecs.add(
                    firstNonBlank(
                            rRow.txtGenpanmei.getText(),
                            JuchuSheetColumnLayout.buildSpecName(
                                    rRow.txtPart.getText(),
                                    rRow.txtType.getText(),
                                    rRow.txtWidth.getText(),
                                    rRow.txtLength.getText())));
            rawMaster.add(rRow.txtItem.getText().trim());
            rawGrades.add(rRow.txtGrade.getText().trim());
            rawColors.add(rRow.txtColor.getText().trim());
            rawCats.add(rRow.txtCategory.getText().trim());
            rawQtys.add(rRow.txtQty.getText().trim());
            wariVals.add(rRow.cmbWariSu.getValue() != null ? rRow.cmbWariSu.getValue() : "1");
            feedVals.add(rRow.cmbFeedLoc.getValue() != null ? rRow.cmbFeedLoc.getValue() : "");
            storageVals.add(rRow.cmbStorageLoc.getValue() != null ? rRow.cmbStorageLoc.getValue() : "");
            inputDateVals.add(rRow.dpInputDate.getValue() != null ? rRow.dpInputDate.getValue().toString() : "");
            rollCountVals.add(rRow.txtRollCount.getText().trim());
        }

        String rawHinmeiJoined = String.join("\n", rawHinmeis);
        db.put("品名1", rawHinmeiJoined);
        db.put("原反品名", rawHinmeiJoined);
        db.put("原反", String.join("\n", rawSpecs));
        db.put(COL_MASTER_BASE_SHOHIN_RAW, String.join("\n", rawMaster));
        db.put("原反梱-等", String.join("\n", rawGrades));
        db.put("原反色", String.join("\n", rawColors));
        db.put("原反区分", String.join("\n", rawCats));
        db.put("原反数量", String.join("\n", rawQtys));
        db.put("割数", String.join("\n", wariVals));
        db.put("投入場所", String.join("\n", feedVals));
        db.put("在庫場所", String.join("\n", storageVals));
        db.put("投入日", String.join("\n", inputDateVals));
        db.put("原反ロール数", String.join("\n", rollCountVals));
        return db;
    }

    private void showAlert(String title, String content) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(content);
        alert.showAndWait();
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    return sdf.format(cell.getDateCellValue());
                } else {
                    double val = cell.getNumericCellValue();
                    if (val == (long) val) {
                        return String.format("%d", (long) val);
                    } else {
                        return String.format("%s", val);
                    }
                }
            case FORMULA:
                try {
                    double val = cell.getNumericCellValue();
                    if (val == (long) val) {
                        return String.format("%d", (long) val);
                    } else {
                        return String.format("%s", val);
                    }
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
            default:
                return "";
        }
    }

    private String normalize_key(String val) {
        if (val == null) return "";
        String text = val.strip().toUpperCase();
        text = java.text.Normalizer.normalize(text, java.text.Normalizer.Form.NFKC);
        text = text.replaceAll("\\s+", "");
        text = text.replace("－", "-").replace("ー", "-").replace("―", "-").replace("‐", "-");
        return text;
    }

    private String normalize_text(String val) {
        if (val == null) return "";
        String text = val.strip();
        text = java.text.Normalizer.normalize(text, java.text.Normalizer.Form.NFKC);
        text = text.replaceAll("\\s+", "");
        text = text.replace("－", "-").replace("ー", "-").replace("―", "-").replace("‐", "-");
        return text.toUpperCase(java.util.Locale.ROOT);
    }

    private void copyToClipboard(String text, Button sourceButton) {
        if (text == null || text.trim().isEmpty()) {
            return;
        }
        javafx.scene.input.Clipboard clipboard = javafx.scene.input.Clipboard.getSystemClipboard();
        javafx.scene.input.ClipboardContent content = new javafx.scene.input.ClipboardContent();
        content.putString(text.trim());
        clipboard.setContent(content);
        
        javafx.stage.Popup popup = new javafx.stage.Popup();
        Label label = new Label("コピーしました！");
        label.setStyle(
            "-fx-background-color: #0F172A;" +
            "-fx-text-fill: #38BDF8;" +
            "-fx-border-color: #38BDF8;" +
            "-fx-border-width: 1px;" +
            "-fx-border-radius: 4px;" +
            "-fx-background-radius: 4px;" +
            "-fx-padding: 6px 12px;" +
            "-fx-font-size: 11px;" +
            "-fx-font-weight: bold;" +
            "-fx-effect: dropshadow(three-pass-box, rgba(56, 189, 248, 0.4), 8, 0, 0, 0);"
        );
        popup.getContent().add(label);
        
        javafx.geometry.Bounds bounds = sourceButton.localToScreen(sourceButton.getBoundsInLocal());
        if (bounds != null) {
            popup.show(sourceButton, bounds.getMinX() - 10, bounds.getMaxY() + 5);
        } else {
            popup.show(sourceButton.getScene().getWindow());
        }
        
        javafx.animation.Timeline timeline = new javafx.animation.Timeline(
            new javafx.animation.KeyFrame(javafx.util.Duration.seconds(2), e -> popup.hide())
        );
        timeline.play();
    }

    private double normalize_numeric(String val) {
        if (val == null || val.isEmpty()) return 0.0;
        String text = val.strip();
        Pattern p = Pattern.compile("[-+]?\\d*\\.\\d+|\\d+");
        var m = p.matcher(text);
        if (m.find()) {
            return Double.parseDouble(m.group());
        }
        return 0.0;
    }

    private String normalize_date_val(String val) {
        if (val == null) return "";
        String text = val.strip();
        List<String> fmts = Arrays.asList("yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss", "yyyy-MM-dd", "yyyy/MM/dd");
        for (String fmt : fmts) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                Date d = sdf.parse(text);
                SimpleDateFormat outSdf = new SimpleDateFormat("yyyy-MM-dd");
                return outSdf.format(d);
            } catch (Exception e) {
                // continue
            }
        }
        return normalize_text(val);
    }

    private static void applyMainShellAlignedStyles(Parent root) {
        if (root == null) {
            return;
        }
        if (!root.getStyleClass().contains("pm-request-form-reconciliation-root")) {
            root.getStyleClass().add("pm-request-form-reconciliation-root");
        }
        var url =
                ReconciliationApp.class.getResource(
                        "/jp/co/pm/ai/desktop/reconciliation/request-form-reconciliation.css");
        if (url == null) {
            System.err.println("Request form stylesheet not found on classpath");
            return;
        }
        String external = url.toExternalForm();
        if (!root.getStylesheets().contains(external)) {
            root.getStylesheets().add(external);
        }
        root.getStylesheets().removeIf(s -> s.contains("/reconciliation/style_"));
    }

    private void loadSettings() {
        RequestFormInputSettingsStore.load(uiEnvSnapshot)
                .ifPresent(
                        settings -> {
                            RequestFormInputSettingsStore.ReconciliationPaths paths =
                                    settings.paths();
                            if (paths != null) {
                                String folder = paths.targetFolder();
                                if (folder != null
                                        && !folder.isBlank()
                                        && new File(folder).isDirectory()) {
                                    targetFolder = folder;
                                }
                                String juchuPath = paths.juchuFilePath();
                                if (juchuPath != null
                                        && !juchuPath.isBlank()
                                        && new File(juchuPath).isFile()) {
                                    juchuFilePath = juchuPath;
                                }
                            }
                            if (settings.comboChoices() != null
                                    && !settings.comboChoices().isEmpty()) {
                                applyComboChoices(settings.comboChoices().mergedWithDefaults());
                            }
                        });
    }

    private void saveSettings() {
        if (!FactoryOperatorUserStore.sessionMayMutateRequestFormInput()) {
            return;
        }
        comboChoicesState = snapshotComboChoices();
        RequestFormInputSettingsStore.save(
                uiEnvSnapshot, comboChoicesState, targetFolder, juchuFilePath);
    }

    /** サマリ Excel 同フォルダの {@link AppPaths#REQUEST_FORM_INPUT_SETTINGS_JSON_FILENAME} から候補を再読込。 */
    private void reloadComboChoicesFromSummarySettings() {
        RequestFormComboChoices fromSummary =
                RequestFormInputSettingsStore.loadComboChoices(
                        uiEnvSnapshot, GlobalInitSettingTarget.load());
        if (fromSummary != null && !fromSummary.isEmpty()) {
            applyComboChoices(fromSummary);
        }
    }

    public static class InteractiveImageViewer extends Pane {
        private final ImageView imageView;
        private final Scale scaleTransform;
        private final Translate translateTransform;
        private double currentScale = 1.0;
        private double currentTranslateX = 0;
        private double currentTranslateY = 0;
        private double lastMouseX;
        private double lastMouseY;
        private boolean firstLayout = true;
        private final javafx.scene.image.Image image;

        public InteractiveImageViewer(javafx.scene.image.Image img) {
            this.image = img;
            this.imageView = new ImageView(img);
            this.imageView.setSmooth(true);
            
            scaleTransform = new Scale(1, 1, 0, 0);
            translateTransform = new Translate(0, 0);
            imageView.getTransforms().setAll(translateTransform, scaleTransform);
            
            getChildren().add(imageView);

            getStyleClass().add("request-form-image-viewer");

            image.widthProperty()
                    .addListener(
                            (obs, oldVal, newVal) -> {
                                if (newVal.doubleValue() > 0 && image.getHeight() > 0) {
                                    resetZoomToFit();
                                }
                            });
            image.heightProperty()
                    .addListener(
                            (obs, oldVal, newVal) -> {
                                if (newVal.doubleValue() > 0 && image.getWidth() > 0) {
                                    resetZoomToFit();
                                }
                            });
            
            // 縦方向は ScrollPane に任せる（Pane 自体はクリップしない）
            
            // Mouse pressed: drag init & double click reset
            setOnMousePressed(e -> {
                if (e.getButton() == javafx.scene.input.MouseButton.PRIMARY) {
                    if (e.getClickCount() == 2) {
                        resetZoomToFit();
                    } else {
                        lastMouseX = e.getX();
                        lastMouseY = e.getY();
                    }
                }
            });
            
            // Mouse dragged: pan image
            setOnMouseDragged(e -> {
                if (e.getButton() == javafx.scene.input.MouseButton.PRIMARY) {
                    double dx = e.getX() - lastMouseX;
                    double dy = e.getY() - lastMouseY;
                    currentTranslateX += dx;
                    currentTranslateY += dy;
                    translateTransform.setX(currentTranslateX);
                    translateTransform.setY(currentTranslateY);
                    lastMouseX = e.getX();
                    lastMouseY = e.getY();
                }
            });
            
            // Scroll: Zoom towards mouse pointer
            setOnScroll(e -> {
                double mouseX = e.getX();
                double mouseY = e.getY();
                
                double zoomFactor = (e.getDeltaY() > 0) ? 1.15 : 0.85;
                double newScale = currentScale * zoomFactor;
                
                // Constraints
                if (newScale < 0.05) newScale = 0.05;
                if (newScale > 15.0) newScale = 15.0;
                
                double actualFactor = newScale / currentScale;
                currentScale = newScale;
                
                currentTranslateX = mouseX - (mouseX - currentTranslateX) * actualFactor;
                currentTranslateY = mouseY - (mouseY - currentTranslateY) * actualFactor;
                
                scaleTransform.setX(currentScale);
                scaleTransform.setY(currentScale);
                translateTransform.setX(currentTranslateX);
                translateTransform.setY(currentTranslateY);
                updateContentSize();

                e.consume();
            });
            
            widthProperty()
                    .addListener(
                            (obs, oldVal, newVal) -> {
                                if (newVal.doubleValue() > 0 && firstLayout) {
                                    resetZoomToFit();
                                    firstLayout = false;
                                }
                            });
        }
        
        public void resetZoomToFit() {
            double containerW = getWidth();
            double imageW = image.getWidth();
            double imageH = image.getHeight();

            if (imageW <= 0 || imageH <= 0 || containerW <= 0) {
                return;
            }

            double scaleX = containerW / imageW;
            currentScale = scaleX;

            // 上端（宛先・タイトル行）を必ず見える位置から表示
            currentTranslateX = 0;
            currentTranslateY = 0;

            applyTransform();
            updateContentSize();
        }

        private void applyTransform() {
            scaleTransform.setX(currentScale);
            scaleTransform.setY(currentScale);
            translateTransform.setX(currentTranslateX);
            translateTransform.setY(currentTranslateY);
        }

        private void updateContentSize() {
            double contentH = Math.max(1.0, image.getHeight() * currentScale);
            setMinHeight(contentH);
            setPrefHeight(contentH);
        }
    }

    private void showLoadingOverlay(boolean show, String text) {
        if (loadingOverlay != null) {
            Platform.runLater(() -> {
                loadingOverlayLabel.setText(text);
                loadingOverlay.setVisible(show);
            });
        }
    }

    private void beginJuchuTransfer(String overlayText) {
        juchuTransferInProgress = true;
        showLoadingOverlay(true, overlayText);
        updateTransferButtonState();
    }

    private void endJuchuTransfer() {
        juchuTransferInProgress = false;
        showLoadingOverlay(false, "");
        updateTransferButtonState();
    }

    /**
     * 受注ファイル保存の直前にローカル世代バックアップを試行し、POI で上書き保存する。
     * バックアップは転記スレッド上で同期的に行う（15 分以内の連続書き込みではスキップ）。
     */
    private void saveJuchuWorkbookWithLocalBackup(File file, Workbook wb, String backupReason)
            throws Exception {
        try {
            RequestFormJuchuFileBackupStore.maybeBackupBeforeWrite(
                            file.toPath(), uiEnvSnapshot, backupReason)
                    .ifPresent(
                            entry ->
                                    System.out.println(
                                            "[request-form-juchu-backup] 退避: "
                                                    + entry.displayLabel()
                                                    + " → "
                                                    + entry.resolveArchivePath(
                                                            RequestFormJuchuFileBackupStore
                                                                    .resolveSourceBackupsRoot(
                                                                            uiEnvSnapshot,
                                                                            file.toPath()))));
        } catch (Exception ex) {
            System.err.println(
                    "[request-form-juchu-backup] 退避に失敗しましたが保存は続行します: "
                            + ex.getMessage());
        }
        PoiWorkbookFileWriter.writeReplacing(file.toPath(), wb, uiEnvSnapshot);
    }

    private void updateLoadingOverlayText(String text) {
        if (loadingOverlayLabel != null) {
            Platform.runLater(() -> {
                loadingOverlayLabel.setText(text);
            });
        }
    }

    private TextField addCompactFormField(GridPane grid, String labelText, int col, int row) {
        Label label = new Label(labelText);
        styleFormLabel(label);
        TextField tf = new TextField();
        tf.setStyle("-fx-font-size: 11px;");
        grid.add(label, col * 2, row);
        addFormField(grid, tf, col * 2 + 1, row);
        return tf;
    }

    private static void configureTwoColumnFormGrid(GridPane grid) {
        grid.setHgap(6);
        if (!grid.getColumnConstraints().isEmpty()) {
            return;
        }
        grid.getColumnConstraints()
                .addAll(
                        formLabelColumn(),
                        expandingFormFieldColumn(),
                        formLabelColumn(),
                        expandingFormFieldColumn());
    }

    private static ColumnConstraints formLabelColumn() {
        ColumnConstraints cc = new ColumnConstraints();
        cc.setMinWidth(FORM_LABEL_COLUMN_MIN_WIDTH);
        cc.setPrefWidth(FORM_LABEL_COLUMN_MIN_WIDTH);
        cc.setMaxWidth(Region.USE_PREF_SIZE);
        cc.setHgrow(Priority.NEVER);
        cc.setHalignment(javafx.geometry.HPos.RIGHT);
        return cc;
    }

    private static ColumnConstraints expandingFormFieldColumn() {
        ColumnConstraints cc = new ColumnConstraints();
        cc.setHgrow(Priority.ALWAYS);
        cc.setMinWidth(48);
        cc.setFillWidth(true);
        return cc;
    }

    private static void styleFormLabel(Label label) {
        label.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        label.getStyleClass().add("request-form-field-label");
        label.setWrapText(true);
        label.setTextOverrun(javafx.scene.control.OverrunStyle.CLIP);
        GridPane.setHalignment(label, javafx.geometry.HPos.RIGHT);
        GridPane.setFillWidth(label, true);
    }

    private static void addFormField(GridPane grid, Node field, int col, int row) {
        addFormField(grid, field, col, row, 1, 1);
    }

    private static void addFormField(
            GridPane grid, Node field, int col, int row, int colSpan, int rowSpan) {
        if (field instanceof Region region) {
            region.setMaxWidth(Double.MAX_VALUE);
        }
        GridPane.setHgrow(field, Priority.ALWAYS);
        GridPane.setFillWidth(field, true);
        grid.add(field, col, row, colSpan, rowSpan);
    }

    /** 横並び操作ボタン: ラベル全文表示（均等縮小・省略記号を避ける）。 */
    private static void configureSideActionButton(Button button) {
        button.setStyle("-fx-font-weight: bold; -fx-cursor: hand; -fx-padding: 8px 12px;");
        button.setMinWidth(Region.USE_COMPUTED_SIZE);
        button.setPrefWidth(Region.USE_COMPUTED_SIZE);
        button.setMaxWidth(Region.USE_COMPUTED_SIZE);
        button.setWrapText(false);
        button.getStyleClass().add("request-form-side-action-btn");
    }

    private static void configureSplitFieldRow(HBox box, Region primary, Region secondary) {
        box.setAlignment(Pos.CENTER_LEFT);
        box.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(primary, Priority.ALWAYS);
        primary.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(secondary, Priority.SOMETIMES);
        secondary.setMaxWidth(Double.MAX_VALUE);
    }

    /** 製品行の「商品」欄（商品コード）のみクリアする。品番・タイプ・寸法・候補コンボは触らない。 */
    private void clearProductRowShohin(ProductRow pRow) {
        isLoadingRecord = true;
        try {
            pRow.txtItem.clear();
            clearSelectedMasterCandidate(pRow.cmbSearch, pRow.selectedCandidatePane);
        } finally {
            isLoadingRecord = false;
        }
    }

    /** 原反行の「商品」欄（商品コード）のみクリアする。品番・タイプ・寸法・候補コンボは触らない。 */
    private void clearRawMaterialRowShohin(RawMaterialRow rRow) {
        isLoadingRecord = true;
        try {
            rRow.txtItem.clear();
            clearSelectedMasterCandidate(rRow.cmbSearch, rRow.selectedCandidatePane);
        } finally {
            isLoadingRecord = false;
        }
    }

    private static Button createShohinFieldButton(String text, String styleClass, String tooltip, Runnable action) {
        Button btn = new Button(text);
        btn.setStyle("-fx-font-size: 11px; -fx-padding: 2px 6px; -fx-cursor: hand;");
        btn.getStyleClass().add(styleClass);
        if (tooltip != null && !tooltip.isEmpty()) {
            btn.setTooltip(new Tooltip(tooltip));
        }
        btn.setOnAction(evt -> action.run());
        return btn;
    }

    private void parseAndSetProductFields(String productStr) {
        if (productStr == null || productStr.isEmpty()) {
            newTxtProdPart.setText("");
            newTxtProdType.setText("");
            newTxtProdWidth.setText("");
            newTxtProdLength.setText("");
            return;
        }
        String[] parts = productStr.split("-");
        if (parts.length >= 3) {
            newTxtProdPart.setText(parts[0]);
            newTxtProdType.setText(parts[1]);
            String dims = parts[2];
            String[] wL = dims.split("X");
            if (wL.length >= 2) {
                newTxtProdWidth.setText(wL[0]);
                newTxtProdLength.setText(wL[1]);
            } else {
                newTxtProdWidth.setText(dims);
                newTxtProdLength.setText("");
            }
        } else {
            newTxtProdPart.setText(productStr);
            newTxtProdType.setText("");
            newTxtProdWidth.setText("");
            newTxtProdLength.setText("");
        }
    }

    private void parseAndSetRawFields(String rawStr) {
        if (rawStr == null || rawStr.isEmpty()) {
            newTxtRawPart.setText("");
            newTxtRawType.setText("");
            newTxtRawWidth.setText("");
            newTxtRawLength.setText("");
            return;
        }
        String[] parts = rawStr.split("-");
        if (parts.length >= 3) {
            newTxtRawPart.setText(parts[0]);
            newTxtRawType.setText(parts[1]);
            String dims = parts[2];
            String[] wL = dims.split("X");
            if (wL.length >= 2) {
                newTxtRawWidth.setText(wL[0]);
                newTxtRawLength.setText(wL[1]);
            } else {
                newTxtRawWidth.setText(dims);
                newTxtRawLength.setText("");
            }
        } else {
            newTxtRawPart.setText(rawStr);
            newTxtRawType.setText("");
            newTxtRawWidth.setText("");
            newTxtRawLength.setText("");
        }
    }

    /** 依頼書候補と同じメモリ上商品一覧（統合マスタ②由来）。 */
    public List<ProductInfo> snapshotMasterProductCatalog() {
        synchronized (this) {
            return List.copyOf(masterProductList);
        }
    }

    /** 後加工商品マスタ参照検索向け: 製品側・原反側の先頭フィルタ。 */
    public PostProcessingProductMasterSearch.MasterReferencePrefixFilters
            snapshotMasterCandidatePrefixFilters() {
        return new PostProcessingProductMasterSearch.MasterReferencePrefixFilters(
                List.copyOf(optMasterCandidatePrefixProduct),
                List.copyOf(optMasterCandidatePrefixRaw));
    }

    /**
     * {@code create_integrated_master.py} を実行する。
     *
     * @param statusLabel 進捗表示（任意）
     * @param disableWhileRunning 実行中に無効化するボタン（任意）
     * @param showResultDialog 完了時にダイアログを出すか
     */
    private void runIntegratedMasterGeneration(
            Label statusLabel, Button disableWhileRunning, boolean showResultDialog) {
        if (rejectGuestMutation("マスタリレーション統合結果の生成")) {
            return;
        }
        if (statusLabel != null) {
            statusLabel.setText("生成実行中...");
        }
        if (disableWhileRunning != null) {
            disableWhileRunning.setDisable(true);
        }

        Thread thread =
                new Thread(
                        () -> {
                            try {
                                String scriptPath = resolveIntegratedMasterScript();
                                Path pythonExe = StagePythonExecutable.resolve(uiEnvSnapshot);
                                Map<String, String> childEnv = new HashMap<>(uiEnvSnapshot);
                                childEnv.put("PM_AI_REQUEST_FORM_WORKSPACE", targetFolder);
                                childEnv.put(
                                        AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, targetFolder);
                                childEnv.put(
                                        AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                                        aladdinMasterDirectory().getAbsolutePath());
                                ProcessBuilder pb =
                                        new ProcessBuilder(pythonExe.toString(), scriptPath);
                                pb.directory(new File(targetFolder));
                                pb.redirectErrorStream(true);
                                PythonProcessRunner.mergeUiEnvIntoProcess(pb, childEnv, null);
                                Process process = pb.start();
                                StringBuilder childOut = new StringBuilder();
                                try (BufferedReader br =
                                        new BufferedReader(
                                                new InputStreamReader(
                                                        process.getInputStream(),
                                                        StandardCharsets.UTF_8))) {
                                    String line;
                                    while ((line = br.readLine()) != null) {
                                        if (childOut.length() > 0) {
                                            childOut.append('\n');
                                        }
                                        childOut.append(line);
                                    }
                                }
                                int exitCode = process.waitFor();
                                String outputTail = tailOfChildOutput(childOut.toString(), 1200);

                                Platform.runLater(
                                        () -> {
                                            if (disableWhileRunning != null) {
                                                disableWhileRunning.setDisable(false);
                                            }
                                            if (exitCode == 0) {
                                                PostProcessingProductMasterReferenceCache
                                                        .invalidate();
                                                reloadMasterProductListFromDisk();
                                                if (statusLabel != null) {
                                                    statusLabel.setText(
                                                            "成功！「マスタリレーション統合結果.xlsx」が生成されました。");
                                                }
                                                if (showResultDialog) {
                                                    Alert alert =
                                                            new Alert(Alert.AlertType.INFORMATION);
                                                    alert.setTitle("統合成功");
                                                    alert.setHeaderText(null);
                                                    alert.setContentText(
                                                            "マスタリレーションの統合が完了し、「マスタリレーション統合結果.xlsx」を更新しました！");
                                                    alert.showAndWait();
                                                }
                                            } else {
                                                if (statusLabel != null) {
                                                    statusLabel.setText(
                                                            "失敗 (終了コード: " + exitCode + ")");
                                                }
                                                if (showResultDialog) {
                                                    Alert alert = new Alert(Alert.AlertType.ERROR);
                                                    alert.setTitle("統合失敗");
                                                    alert.setHeaderText(null);
                                                    String detail =
                                                            "統合処理中にエラーが発生しました。"
                                                                    + "\nPython: "
                                                                    + pythonExe
                                                                    + "\nマスタフォルダ: "
                                                                    + aladdinMasterDirectory()
                                                                            .getAbsolutePath();
                                                    if (!outputTail.isBlank()) {
                                                        detail +=
                                                                "\n\n--- スクリプト出力（末尾） ---\n"
                                                                        + outputTail;
                                                    } else {
                                                        detail +=
                                                                "\n\nPythonスクリプト、3つのマスタ xlsx、"
                                                                        + "環境変数 PM_AI_PYTHON を確認してください。";
                                                    }
                                                    alert.setContentText(detail);
                                                    alert.showAndWait();
                                                }
                                            }
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            if (disableWhileRunning != null) {
                                                disableWhileRunning.setDisable(false);
                                            }
                                            if (statusLabel != null) {
                                                statusLabel.setText("エラー: " + ex.getMessage());
                                            }
                                            if (showResultDialog) {
                                                Alert alert = new Alert(Alert.AlertType.ERROR);
                                                alert.setTitle("実行エラー");
                                                alert.setHeaderText(null);
                                                alert.setContentText(
                                                        "プロセスの実行中に例外が発生しました: "
                                                                + ex.getMessage());
                                                alert.showAndWait();
                                            }
                                        });
                            }
                        },
                        "integrated-master-gen");
        thread.setDaemon(true);
        thread.start();
    }

    // --- STYLISH DYNAMIC MULTIPLE ROWS HELPERS & CLASSES ---
    
    public static class ProductRow {
        public TextField txtHinmei;
        public TextField txtSeihinmei;
        public TextField txtItem;
        public TextField txtPart;
        public ComboBox<String> cmbSearch;
        /** 候補確定後のマスタ行表示（フィルタ一致色付き）。 */
        public HBox selectedCandidatePane;
        public TextField txtType;
        public TextField txtWidth;
        public TextField txtLength;
        public TextField txtQty;
        public TextField txtGrade;
        public TextField txtColor;
        public TextField txtCategory;
        public TextField txtEdaban;
        public TextField txtKeiyakuNo;
        public ComboBox<String> cmbEcSide;
        public ComboBox<String> cmbTrimming;
        public Button btnDelete;
        public GridPane grid;
    }
    
    public static class RawMaterialRow {
        public TextField txtHinmei;
        public TextField txtGenpanmei;
        public TextField txtItem;
        public TextField txtPart;
        public ComboBox<String> cmbSearch;
        /** 候補確定後のマスタ行表示（フィルタ一致色付き）。 */
        public HBox selectedCandidatePane;
        public TextField txtType;
        public TextField txtWidth;
        public TextField txtLength;
        public TextField txtQty;
        public TextField txtGrade;
        public TextField txtColor;
        public TextField txtCategory;
        public ComboBox<String> cmbWariSu;
        public ComboBox<String> cmbFeedLoc;
        public ComboBox<String> cmbStorageLoc;
        public DatePicker dpInputDate;
        public TextField txtRollCount;
        /** 原反ロール数を手入力したときは数量・長さ変更での自動上書きを抑止する。 */
        public boolean rollCountManualOverride;
        /** {@link #updateRawRowRollCountDisplay} による setText 中は手入力フラグを立てない。 */
        public boolean suppressRollCountAutoListener;
        public Button btnDelete;
        public GridPane grid;
    }

    private java.time.LocalDate parseLocalDate(String val) {
        if (val == null || val.strip().isEmpty()) return null;
        String text = val.strip();
        java.util.List<String> fmts = java.util.Arrays.asList("yyyy-MM-dd", "yyyy/MM/dd", "yyyyMMdd");
        for (String fmt : fmts) {
            try {
                java.time.format.DateTimeFormatter dtf = java.time.format.DateTimeFormatter.ofPattern(fmt);
                return java.time.LocalDate.parse(text, dtf);
            } catch (Exception e) {
                // try next
            }
        }
        try {
            String dateOnly = text.split(" ")[0];
            return java.time.LocalDate.parse(dateOnly, java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        } catch (Exception e) {}
        try {
            String dateOnly = text.split(" ")[0];
            return java.time.LocalDate.parse(dateOnly, java.time.format.DateTimeFormatter.ofPattern("yyyy/MM/dd"));
        } catch (Exception e) {}
        return null;
    }

    private ProductRow addProductRow(Map<String, String> initialValues) {
        ProductRow pRow = new ProductRow();
        pRow.grid = new GridPane();
        pRow.grid.setVgap(6);
        pRow.grid.setMaxWidth(Double.MAX_VALUE);
        pRow.grid.getStyleClass().add("request-form-prod-row");
        configureTwoColumnFormGrid(pRow.grid);

        Label lblHinmei = new Label("品名:");
        styleFormLabel(lblHinmei);
        pRow.txtHinmei = new TextField();
        pRow.txtHinmei.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblHinmei, 0, 0);
        addFormField(pRow.grid, pRow.txtHinmei, 1, 0);

        Label lblSeihinmei = new Label("製品名:");
        styleFormLabel(lblSeihinmei);
        pRow.txtSeihinmei = new TextField();
        pRow.txtSeihinmei.setStyle("-fx-font-size: 11px;");
        pRow.txtSeihinmei.setEditable(false);
        pRow.txtSeihinmei.setFocusTraversable(false);
        pRow.grid.add(lblSeihinmei, 2, 0);
        addFormField(pRow.grid, pRow.txtSeihinmei, 3, 0);

        Label lblItem = new Label("商品:");
        styleFormLabel(lblItem);
        pRow.txtItem = new TextField();
        pRow.txtItem.setStyle("-fx-font-size: 11px;");
        Button btnClear =
                createShohinFieldButton("✕", "btn-clear", "商品をクリア", () -> clearProductRowShohin(pRow));
        Button btnCopy = new Button("📋");
        btnCopy.setStyle("-fx-font-size: 11px; -fx-padding: 2px 6px; -fx-cursor: hand;");
        btnCopy.getStyleClass().add("btn-copy");
        btnCopy.setTooltip(new Tooltip("商品をコピー"));
        btnCopy.setOnAction(evt -> copyToClipboard(pRow.txtItem.getText(), btnCopy));
        HBox itemActions = new HBox(3, btnClear, btnCopy);
        HBox.setHgrow(itemActions, Priority.NEVER);
        HBox itemBox = new HBox(5, pRow.txtItem, itemActions);
        configureSplitFieldRow(itemBox, pRow.txtItem, itemActions);
        pRow.grid.add(lblItem, 0, 1);
        addFormField(pRow.grid, itemBox, 1, 1, 3, 1);

        Label lblPart = new Label("品番:");
        styleFormLabel(lblPart);
        pRow.txtPart = new TextField();
        pRow.txtPart.setStyle("-fx-font-size: 11px;");
        pRow.cmbSearch = new ComboBox<>();
        pRow.cmbSearch.getStyleClass().add("request-form-master-candidate-combo");
        pRow.cmbSearch.setStyle("-fx-font-size: 11px;");
        pRow.cmbSearch.setPromptText("候補");
        pRow.cmbSearch.setMinWidth(96);
        pRow.cmbSearch.setPrefWidth(128);
        wireCandidateComboBox(pRow.cmbSearch, () -> updateRowProdCandidates(pRow, true));
        HBox partBox = new HBox(5, pRow.txtPart, pRow.cmbSearch);
        configureSplitFieldRow(partBox, pRow.txtPart, pRow.cmbSearch);
        pRow.selectedCandidatePane = createSelectedCandidatePane();
        VBox partColumn = new VBox(4, partBox, pRow.selectedCandidatePane);
        partColumn.setMaxWidth(Double.MAX_VALUE);
        pRow.grid.add(lblPart, 0, 2);
        addFormField(pRow.grid, partColumn, 1, 2, 3, 1);

        Label lblType = new Label("タイプ:");
        styleFormLabel(lblType);
        pRow.txtType = new TextField();
        pRow.txtType.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblType, 0, 3);
        addFormField(pRow.grid, pRow.txtType, 1, 3);

        Label lblWidth = new Label("幅(mm):");
        styleFormLabel(lblWidth);
        pRow.txtWidth = new TextField();
        pRow.txtWidth.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblWidth, 2, 3);
        addFormField(pRow.grid, pRow.txtWidth, 3, 3);

        Label lblLength = new Label("長さ(m):");
        styleFormLabel(lblLength);
        pRow.txtLength = new TextField();
        pRow.txtLength.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblLength, 0, 4);
        addFormField(pRow.grid, pRow.txtLength, 1, 4);

        Label lblQty = new Label("数量:");
        styleFormLabel(lblQty);
        pRow.txtQty = new TextField();
        pRow.txtQty.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblQty, 2, 4);
        addFormField(pRow.grid, pRow.txtQty, 3, 4);

        Label lblGrade = new Label("梱-等:");
        styleFormLabel(lblGrade);
        pRow.txtGrade = new TextField();
        pRow.txtGrade.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblGrade, 0, 5);
        addFormField(pRow.grid, pRow.txtGrade, 1, 5);

        Label lblColor = new Label("色:");
        styleFormLabel(lblColor);
        pRow.txtColor = new TextField();
        pRow.txtColor.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblColor, 2, 5);
        addFormField(pRow.grid, pRow.txtColor, 3, 5);

        Label lblCategory = new Label("区分:");
        styleFormLabel(lblCategory);
        pRow.txtCategory = new TextField();
        pRow.txtCategory.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblCategory, 0, 6);
        addFormField(pRow.grid, pRow.txtCategory, 1, 6);

        Label lblEdaban = new Label("枝番:");
        styleFormLabel(lblEdaban);
        pRow.txtEdaban = new TextField();
        pRow.txtEdaban.setStyle("-fx-font-size: 11px;");
        pRow.txtEdaban.getStyleClass().add("paper-edaban-input");
        pRow.grid.add(lblEdaban, 2, 6);
        addFormField(pRow.grid, pRow.txtEdaban, 3, 6);

        Label lblKeiyakuNo = new Label("契約Ｎｏ:");
        styleFormLabel(lblKeiyakuNo);
        pRow.txtKeiyakuNo = new TextField();
        pRow.txtKeiyakuNo.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblKeiyakuNo, 0, 7);
        addFormField(pRow.grid, pRow.txtKeiyakuNo, 1, 7, 3, 1);

        Label lblEc = new Label("EC面:");
        styleFormLabel(lblEc);
        pRow.cmbEcSide = new ComboBox<>(optEcSide);
        pRow.cmbEcSide.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblEc, 0, 8);
        addFormField(pRow.grid, pRow.cmbEcSide, 1, 8);

        Label lblTrimming = new Label("トリミング:");
        styleFormLabel(lblTrimming);
        pRow.cmbTrimming = new ComboBox<>(optTrimming);
        pRow.cmbTrimming.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblTrimming, 2, 8);
        addFormField(pRow.grid, pRow.cmbTrimming, 3, 8);

        Button btnDelete = new Button("削除");
        btnDelete.getStyleClass().add("btn-settings-del");
        btnDelete.setStyle("-fx-font-size: 11px; -fx-padding: 3px 8px; -fx-cursor: hand;");
        btnDelete.setOnAction(evt -> {
            productRows.remove(pRow);
            productRowsContainer.getChildren().remove(pRow.grid);
            if (!productRows.isEmpty()) {
                newTxtProdItem = productRows.get(0).txtItem;
                newTxtProdPart = productRows.get(0).txtPart;
                newTxtProdType = productRows.get(0).txtType;
                newTxtProdWidth = productRows.get(0).txtWidth;
                newTxtProdLength = productRows.get(0).txtLength;
                newTxtProdGrade = productRows.get(0).txtGrade;
                newTxtProdColor = productRows.get(0).txtColor;
                newTxtProdCategory = productRows.get(0).txtCategory;
                newTxtProdEdaban = productRows.get(0).txtEdaban;
                newTxtProdQty = productRows.get(0).txtQty;
                newCmbProdEcSide = productRows.get(0).cmbEcSide;
                newCmbProdTrimming = productRows.get(0).cmbTrimming;
            }
        });
        pRow.grid.add(btnDelete, 3, 9);
        GridPane.setHalignment(btnDelete, javafx.geometry.HPos.RIGHT);

        Runnable updateSpec = () -> updateProductRowSpecDisplay(pRow);
        pRow.txtHinmei.textProperty().addListener((obs, oldV, newV) -> updateRowProdCandidates(pRow, false));
        pRow.txtPart.textProperty().addListener((obs, oldV, newV) -> {
            updateRowProdCandidates(pRow, false);
            updateSpec.run();
        });
        pRow.txtItem.textProperty().addListener((obs, oldV, newV) -> {
            updateRowProdCandidates(pRow, false);
            syncSelectedCandidatePaneWithItemCode(pRow.cmbSearch, pRow.selectedCandidatePane, newV);
        });
        pRow.txtType.textProperty().addListener((obs, oldV, newV) -> {
            updateRowProdCandidates(pRow, false);
            updateSpec.run();
        });
        pRow.txtWidth.textProperty().addListener((obs, oldV, newV) -> updateSpec.run());
        pRow.txtLength.textProperty().addListener((obs, oldV, newV) -> {
            updateRowProdCandidates(pRow, false);
            updateSpec.run();
        });
        
        pRow.cmbSearch.setOnAction(evt -> applyMasterProductCandidateSelection(pRow.cmbSearch.getValue(), pRow));

        productRows.add(pRow);
        productRowsContainer.getChildren().add(pRow.grid);

        if (productRows.size() == 1) {
            newTxtProdItem = pRow.txtItem;
            newTxtProdPart = pRow.txtPart;
            newTxtProdType = pRow.txtType;
            newTxtProdWidth = pRow.txtWidth;
            newTxtProdLength = pRow.txtLength;
            newTxtProdGrade = pRow.txtGrade;
            newTxtProdColor = pRow.txtColor;
            newTxtProdCategory = pRow.txtCategory;
            newTxtProdEdaban = pRow.txtEdaban;
            newTxtProdQty = pRow.txtQty;
            newCmbProdEcSide = pRow.cmbEcSide;
            newCmbProdTrimming = pRow.cmbTrimming;
        }
        return pRow;
    }

    private void refreshAllRowCandidates() {
        for (ProductRow pRow : productRows) {
            updateRowProdCandidates(pRow, false, false);
        }
        for (RawMaterialRow rRow : rawRows) {
            updateRowRawCandidates(rRow, false, false);
        }
    }

    private void hideAllMasterCandidateCombos() {
        for (ProductRow pRow : productRows) {
            if (pRow.cmbSearch != null) {
                pRow.cmbSearch.hide();
            }
        }
        for (RawMaterialRow rRow : rawRows) {
            if (rRow.cmbSearch != null) {
                rRow.cmbSearch.hide();
            }
        }
    }

    private static final String MASTER_CANDIDATE_FILTER_KEYWORDS_PROP = "filterKeywords";
    private static final String MASTER_CANDIDATE_SELECTED_LABEL_PROP = "selectedCandidateLabel";
    private static final String MASTER_CANDIDATE_ALL_ITEMS_PROP = "allCandidateItems";
    private static final String MASTER_CANDIDATE_SUPPRESS_EDITOR_FILTER_PROP = "suppressCandidateEditorFilter";

    private static void wireCandidateComboBox(ComboBox<String> combo, Runnable refreshOnOpen) {
        combo.setEditable(true);
        combo.getEditor()
                .textProperty()
                .addListener(
                        (obs, oldV, newV) -> {
                            if (Boolean.TRUE.equals(
                                    combo.getProperties().get(MASTER_CANDIDATE_SUPPRESS_EDITOR_FILTER_PROP))) {
                                return;
                            }
                            applyMasterCandidateComboEditorFilter(combo, newV);
                        });
        combo.setOnShowing(
                e -> {
                    if (refreshOnOpen != null) {
                        refreshOnOpen.run();
                    }
                });
        combo.setButtonCell(
                new ListCell<>() {
                    @Override
                    protected void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (empty || item == null) {
                            setText(null);
                            setGraphic(null);
                            return;
                        }
                        setText(null);
                        setGraphic(
                                RequestFormMasterCandidateLabelHighlighter.buildGraphic(
                                        item, masterCandidateFilterKeywords(combo)));
                    }
                });
        combo.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    setGraphic(null);
                                    return;
                                }
                                setText(null);
                                setGraphic(
                                        RequestFormMasterCandidateLabelHighlighter.buildGraphic(
                                                item, masterCandidateFilterKeywords(combo)));
                            }
                        });
    }

    @SuppressWarnings("unchecked")
    private static java.util.List<String> masterCandidateAllItems(ComboBox<String> combo) {
        Object stored = combo.getProperties().get(MASTER_CANDIDATE_ALL_ITEMS_PROP);
        if (stored instanceof java.util.List<?> list) {
            return (java.util.List<String>) list;
        }
        return java.util.List.of();
    }

    private static void setMasterCandidateAllItems(ComboBox<String> combo, java.util.List<String> allItems) {
        combo.getProperties()
                .put(
                        MASTER_CANDIDATE_ALL_ITEMS_PROP,
                        allItems != null ? java.util.List.copyOf(allItems) : java.util.List.of());
        applyMasterCandidateComboEditorFilter(combo, combo.getEditor().getText());
    }

    private static void applyMasterCandidateComboEditorFilter(ComboBox<String> combo, String editorText) {
        java.util.List<String> all = masterCandidateAllItems(combo);
        String query = RequestFormMasterProductCandidateMatcher.normalize(editorText);
        java.util.List<String> visible;
        if (query.isEmpty()) {
            visible = all;
        } else {
            visible = new java.util.ArrayList<>();
            for (String item : all) {
                if (RequestFormMasterProductCandidateMatcher.normalize(item).contains(query)) {
                    visible.add(item);
                }
            }
        }
        runWithSuppressedCandidateEditorFilter(
                combo, () -> combo.setItems(javafx.collections.FXCollections.observableArrayList(visible)));
    }

    private static void clearMasterCandidateEditor(ComboBox<String> combo) {
        runWithSuppressedCandidateEditorFilter(combo, () -> combo.getEditor().clear());
    }

    /**
     * 候補確定でエディタに入った完全ラベル（{@code " | "} を含む）は、次回ドロップダウンを開いたときの
     * 絞り込み条件になってしまい「それ以外の候補が表示されない」原因になる。確定ラベルのときだけ消す
     * （ユーザーが途中まで打ったクエリは {@code " | "} を含まないため残す）。
     */
    private static void clearConfirmedCandidateEditorText(ComboBox<String> combo) {
        if (combo == null) {
            return;
        }
        String text = combo.getEditor().getText();
        if (text != null && text.contains(" | ")) {
            clearMasterCandidateEditor(combo);
        }
    }

    private static void runWithSuppressedCandidateEditorFilter(ComboBox<String> combo, Runnable action) {
        combo.getProperties().put(MASTER_CANDIDATE_SUPPRESS_EDITOR_FILTER_PROP, Boolean.TRUE);
        try {
            action.run();
        } finally {
            combo.getProperties().remove(MASTER_CANDIDATE_SUPPRESS_EDITOR_FILTER_PROP);
        }
    }

    private static void setMasterCandidateFilterKeywords(
            ComboBox<String> combo, String kwItem, String kwPart, String kwType, String kwLength, String kwHinmei) {
        java.util.List<String> keywords = new java.util.ArrayList<>(5);
        if (kwItem != null && !kwItem.isEmpty()) {
            keywords.add(kwItem);
        }
        if (kwPart != null && !kwPart.isEmpty()) {
            keywords.add(kwPart);
        }
        if (kwType != null && !kwType.isEmpty()) {
            keywords.add(kwType);
        }
        if (kwLength != null && !kwLength.isEmpty()) {
            keywords.add(kwLength);
        }
        if (kwHinmei != null && !kwHinmei.isEmpty()) {
            keywords.add(kwHinmei);
        }
        combo.getProperties().put(MASTER_CANDIDATE_FILTER_KEYWORDS_PROP, keywords);
    }

    @SuppressWarnings("unchecked")
    private static java.util.List<String> masterCandidateFilterKeywords(ComboBox<String> combo) {
        Object value = combo.getProperties().get(MASTER_CANDIDATE_FILTER_KEYWORDS_PROP);
        if (value instanceof java.util.List<?> list) {
            return (java.util.List<String>) list;
        }
        return java.util.List.of();
    }

    private static String shohinCodeFromMasterCandidateLabel(String candidateLabel) {
        if (candidateLabel == null || !candidateLabel.contains(" | ")) {
            return "";
        }
        return candidateLabel.split(" \\| ", 2)[0].trim();
    }

    private ProductInfo findMasterProductByShohinCode(String shohinCode) {
        if (shohinCode == null || shohinCode.isBlank()) {
            return null;
        }
        String code = RequestFormMasterProductCandidateMatcher.normalize(shohinCode);
        for (ProductInfo p : masterProductList) {
            if (RequestFormMasterProductCandidateMatcher.normalize(p.getShohinCode()).equals(code)) {
                return p;
            }
        }
        return null;
    }

    private static HBox createSelectedCandidatePane() {
        HBox pane = new HBox();
        pane.getStyleClass().add("request-form-master-candidate-selected");
        pane.setAlignment(Pos.CENTER_LEFT);
        pane.setManaged(false);
        pane.setVisible(false);
        pane.setMaxWidth(Double.MAX_VALUE);
        Tooltip.install(pane, new Tooltip("マスタ候補で確定した内容（フィルタ一致部分は黄色）"));
        return pane;
    }

    private void refreshSelectedCandidateDisplay(ComboBox<String> combo, HBox selectedPane) {
        if (selectedPane == null) {
            return;
        }
        selectedPane.getChildren().clear();
        Object stored = combo != null ? combo.getProperties().get(MASTER_CANDIDATE_SELECTED_LABEL_PROP) : null;
        String label = stored instanceof String s ? s : null;
        if (label == null || label.isBlank()) {
            selectedPane.setManaged(false);
            selectedPane.setVisible(false);
            return;
        }
        selectedPane.setManaged(true);
        selectedPane.setVisible(true);
        selectedPane
                .getChildren()
                .add(
                        RequestFormMasterCandidateLabelHighlighter.buildGraphic(
                                label, masterCandidateFilterKeywords(combo)));
    }

    private static void clearSelectedMasterCandidate(ComboBox<String> combo, HBox selectedPane) {
        if (combo != null) {
            combo.getProperties().remove(MASTER_CANDIDATE_SELECTED_LABEL_PROP);
        }
        if (selectedPane != null) {
            selectedPane.getChildren().clear();
            selectedPane.setManaged(false);
            selectedPane.setVisible(false);
        }
    }

    private void syncSelectedCandidatePaneWithItemCode(
            ComboBox<String> combo, HBox selectedPane, String itemCodeText) {
        if (combo == null || selectedPane == null) {
            return;
        }
        Object stored = combo.getProperties().get(MASTER_CANDIDATE_SELECTED_LABEL_PROP);
        if (!(stored instanceof String label) || label.isBlank()) {
            return;
        }
        String selectedCode = shohinCodeFromMasterCandidateLabel(label);
        String current = normalize_text(itemCodeText);
        if (!RequestFormMasterProductCandidateMatcher.normalize(selectedCode)
                .equals(RequestFormMasterProductCandidateMatcher.normalize(current))) {
            clearSelectedMasterCandidate(combo, selectedPane);
        }
    }

    /**
     * マスタ候補の選択確定: 「商品」欄へ商品コードを入れ、確定ラベルを色付きで表示する。
     */
    private void confirmMasterCandidateSelection(
            String candidateLabel, TextField txtItem, ComboBox<String> combo, HBox selectedPane) {
        applyMasterCandidateShohinCodeOnly(candidateLabel, txtItem);
        if (combo == null) {
            return;
        }
        if (candidateLabel != null && candidateLabel.contains(" | ")) {
            combo.getProperties().put(MASTER_CANDIDATE_SELECTED_LABEL_PROP, candidateLabel);
        } else {
            combo.getProperties().remove(MASTER_CANDIDATE_SELECTED_LABEL_PROP);
        }
        refreshSelectedCandidateDisplay(combo, selectedPane);
    }

    /** マスタ候補コンボの選択確定時、当該行の「商品」欄（商品コード）のみを反映する。 */
    private void applyMasterProductCandidateSelection(String candidateLabel, ProductRow pRow) {
        if (pRow == null) {
            return;
        }
        confirmMasterCandidateSelection(
                candidateLabel, pRow.txtItem, pRow.cmbSearch, pRow.selectedCandidatePane);
    }

    private void applyMasterCandidateShohinCodeOnly(String candidateLabel, TextField txtItem) {
        if (txtItem == null) {
            return;
        }
        ProductInfo product = findMasterProductByShohinCode(shohinCodeFromMasterCandidateLabel(candidateLabel));
        if (product == null) {
            return;
        }
        isLoadingRecord = true;
        try {
            txtItem.setText(product.getShohinCode());
        } finally {
            isLoadingRecord = false;
        }
    }

    private void updateRowProdCandidates(ProductRow pRow, boolean fromDropdownOpen) {
        updateRowProdCandidates(pRow, fromDropdownOpen, true);
    }

    private void updateRowProdCandidates(ProductRow pRow, boolean fromDropdownOpen, boolean autoOpenPopup) {
        if (isLoadingRecord && !fromDropdownOpen) {
            return;
        }
        if (fromDropdownOpen) {
            clearConfirmedCandidateEditorText(pRow.cmbSearch);
        }
        // 「商品」欄（商品コード）は候補選択の出力先のため、候補フィルタの条件には使わない
        String kwItem = "";
        String kwPart = normalize_text(pRow.txtPart.getText());
        String kwType = normalize_text(pRow.txtType.getText());
        String kwLength = normalize_text(pRow.txtLength.getText());
        String kwHinmei = normalize_text(pRow.txtHinmei.getText());
        setMasterCandidateFilterKeywords(pRow.cmbSearch, kwItem, kwPart, kwType, kwLength, kwHinmei);
        refreshSelectedCandidateDisplay(pRow.cmbSearch, pRow.selectedCandidatePane);

        java.util.List<String> filtered;
        if (kwItem.isEmpty()
                && kwPart.isEmpty()
                && kwType.isEmpty()
                && kwLength.isEmpty()
                && kwHinmei.isEmpty()) {
            if (!fromDropdownOpen) {
                setMasterCandidateAllItems(pRow.cmbSearch, java.util.List.of());
                return;
            }
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            filteredMasterCatalogForProduct(), "", "", "", "", "", 50);
        } else {
            if (!fromDropdownOpen) {
                clearMasterCandidateEditor(pRow.cmbSearch);
            }
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            filteredMasterCatalogForProduct(),
                            kwItem,
                            kwPart,
                            kwType,
                            kwLength,
                            kwHinmei,
                            50);
        }
        setMasterCandidateAllItems(pRow.cmbSearch, filtered);
        if (autoOpenPopup && !fromDropdownOpen && !filtered.isEmpty() && !pRow.cmbSearch.isShowing()) {
            pRow.cmbSearch.show();
        }
    }

    private void updateRowRawCandidates(RawMaterialRow rRow, boolean fromDropdownOpen) {
        updateRowRawCandidates(rRow, fromDropdownOpen, true);
    }

    private void updateRowRawCandidates(RawMaterialRow rRow, boolean fromDropdownOpen, boolean autoOpenPopup) {
        if (isLoadingRecord && !fromDropdownOpen) {
            return;
        }
        if (fromDropdownOpen) {
            clearConfirmedCandidateEditorText(rRow.cmbSearch);
        }
        // 「商品」欄（商品コード）は候補選択の出力先のため、候補フィルタの条件には使わない
        String kwItem = "";
        String kwPart = normalize_text(rRow.txtPart.getText());
        String kwType = normalize_text(rRow.txtType.getText());
        String kwLength = normalize_text(rRow.txtLength.getText());
        String kwHinmei = normalize_text(rRow.txtHinmei.getText());
        setMasterCandidateFilterKeywords(rRow.cmbSearch, kwItem, kwPart, kwType, kwLength, kwHinmei);
        refreshSelectedCandidateDisplay(rRow.cmbSearch, rRow.selectedCandidatePane);

        java.util.List<String> filtered;
        if (kwItem.isEmpty()
                && kwPart.isEmpty()
                && kwType.isEmpty()
                && kwLength.isEmpty()
                && kwHinmei.isEmpty()) {
            if (!fromDropdownOpen) {
                setMasterCandidateAllItems(rRow.cmbSearch, java.util.List.of());
                return;
            }
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            filteredMasterCatalogForRaw(), "", "", "", "", "", 50);
        } else {
            if (!fromDropdownOpen) {
                clearMasterCandidateEditor(rRow.cmbSearch);
            }
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            filteredMasterCatalogForRaw(),
                            kwItem,
                            kwPart,
                            kwType,
                            kwLength,
                            kwHinmei,
                            50);
        }
        setMasterCandidateAllItems(rRow.cmbSearch, filtered);
        if (autoOpenPopup && !fromDropdownOpen && !filtered.isEmpty() && !rRow.cmbSearch.isShowing()) {
            rRow.cmbSearch.show();
        }
    }

    private RawMaterialRow addRawRow(Map<String, String> initialValues) {
        RawMaterialRow rRow = new RawMaterialRow();
        rRow.grid = new GridPane();
        rRow.grid.setVgap(6);
        rRow.grid.setMaxWidth(Double.MAX_VALUE);
        rRow.grid.getStyleClass().add("request-form-prod-row");
        configureTwoColumnFormGrid(rRow.grid);

        Label lblHinmei = new Label("品名:");
        styleFormLabel(lblHinmei);
        rRow.txtHinmei = new TextField();
        rRow.txtHinmei.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblHinmei, 0, 0);
        addFormField(rRow.grid, rRow.txtHinmei, 1, 0);

        Label lblGenpanmei = new Label("原反名:");
        styleFormLabel(lblGenpanmei);
        rRow.txtGenpanmei = new TextField();
        rRow.txtGenpanmei.setStyle("-fx-font-size: 11px;");
        rRow.txtGenpanmei.setEditable(false);
        rRow.txtGenpanmei.setFocusTraversable(false);
        rRow.grid.add(lblGenpanmei, 2, 0);
        addFormField(rRow.grid, rRow.txtGenpanmei, 3, 0);

        Label lblItem = new Label("商品:");
        styleFormLabel(lblItem);
        rRow.txtItem = new TextField();
        rRow.txtItem.setStyle("-fx-font-size: 11px;");
        Button btnClear =
                createShohinFieldButton("✕", "btn-clear", "商品をクリア", () -> clearRawMaterialRowShohin(rRow));
        Button btnCopy = new Button("📋");
        btnCopy.setStyle("-fx-font-size: 11px; -fx-padding: 2px 6px; -fx-cursor: hand;");
        btnCopy.getStyleClass().add("btn-copy");
        btnCopy.setTooltip(new Tooltip("商品をコピー"));
        btnCopy.setOnAction(evt -> copyToClipboard(rRow.txtItem.getText(), btnCopy));
        HBox itemActions = new HBox(3, btnClear, btnCopy);
        HBox.setHgrow(itemActions, Priority.NEVER);
        HBox itemBox = new HBox(5, rRow.txtItem, itemActions);
        configureSplitFieldRow(itemBox, rRow.txtItem, itemActions);
        rRow.grid.add(lblItem, 0, 1);
        addFormField(rRow.grid, itemBox, 1, 1, 3, 1);

        Label lblPart = new Label("品番:");
        styleFormLabel(lblPart);
        rRow.txtPart = new TextField();
        rRow.txtPart.setStyle("-fx-font-size: 11px;");
        rRow.cmbSearch = new ComboBox<>();
        rRow.cmbSearch.getStyleClass().add("request-form-master-candidate-combo");
        rRow.cmbSearch.setStyle("-fx-font-size: 11px;");
        rRow.cmbSearch.setPromptText("候補");
        rRow.cmbSearch.setMinWidth(96);
        rRow.cmbSearch.setPrefWidth(128);
        wireCandidateComboBox(rRow.cmbSearch, () -> updateRowRawCandidates(rRow, true));
        HBox partBox = new HBox(5, rRow.txtPart, rRow.cmbSearch);
        configureSplitFieldRow(partBox, rRow.txtPart, rRow.cmbSearch);
        rRow.selectedCandidatePane = createSelectedCandidatePane();
        VBox partColumn = new VBox(4, partBox, rRow.selectedCandidatePane);
        partColumn.setMaxWidth(Double.MAX_VALUE);
        rRow.grid.add(lblPart, 0, 2);
        addFormField(rRow.grid, partColumn, 1, 2, 3, 1);

        Label lblType = new Label("タイプ:");
        styleFormLabel(lblType);
        rRow.txtType = new TextField();
        rRow.txtType.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblType, 0, 3);
        addFormField(rRow.grid, rRow.txtType, 1, 3);

        Label lblWidth = new Label("幅(mm):");
        styleFormLabel(lblWidth);
        rRow.txtWidth = new TextField();
        rRow.txtWidth.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblWidth, 2, 3);
        addFormField(rRow.grid, rRow.txtWidth, 3, 3);

        Label lblLength = new Label("長さ(m):");
        styleFormLabel(lblLength);
        rRow.txtLength = new TextField();
        rRow.txtLength.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblLength, 0, 4);
        addFormField(rRow.grid, rRow.txtLength, 1, 4);

        Label lblQty = new Label("数量:");
        styleFormLabel(lblQty);
        rRow.txtQty = new TextField();
        rRow.txtQty.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblQty, 2, 4);
        addFormField(rRow.grid, rRow.txtQty, 3, 4);

        Label lblGrade = new Label("梱-等:");
        styleFormLabel(lblGrade);
        rRow.txtGrade = new TextField();
        rRow.txtGrade.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblGrade, 0, 5);
        addFormField(rRow.grid, rRow.txtGrade, 1, 5);

        Label lblColor = new Label("色:");
        styleFormLabel(lblColor);
        rRow.txtColor = new TextField();
        rRow.txtColor.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblColor, 2, 5);
        addFormField(rRow.grid, rRow.txtColor, 3, 5);

        Label lblCategory = new Label("区分:");
        styleFormLabel(lblCategory);
        rRow.txtCategory = new TextField();
        rRow.txtCategory.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblCategory, 0, 6);
        addFormField(rRow.grid, rRow.txtCategory, 1, 6);

        Label lblWariSu = new Label("割数:");
        styleFormLabel(lblWariSu);
        rRow.cmbWariSu = new ComboBox<>(optWariSu);
        rRow.cmbWariSu.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblWariSu, 2, 6);
        addFormField(rRow.grid, rRow.cmbWariSu, 3, 6);

        Label lblFeed = new Label("投入場所:");
        styleFormLabel(lblFeed);
        rRow.cmbFeedLoc = new ComboBox<>(optFeedLoc);
        rRow.cmbFeedLoc.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblFeed, 0, 7);
        addFormField(rRow.grid, rRow.cmbFeedLoc, 1, 7);

        Label lblStorage = new Label("在庫場所:");
        styleFormLabel(lblStorage);
        rRow.cmbStorageLoc = new ComboBox<>(optStorageLoc);
        rRow.cmbStorageLoc.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblStorage, 2, 7);
        addFormField(rRow.grid, rRow.cmbStorageLoc, 3, 7);

        Label lblInputDate = new Label("投入日:");
        styleFormLabel(lblInputDate);
        rRow.dpInputDate = new DatePicker();
        rRow.dpInputDate.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblInputDate, 0, 8);
        addFormField(rRow.grid, rRow.dpInputDate, 1, 8);

        Label lblRollCount = new Label("原反ロール数:");
        styleFormLabel(lblRollCount);
        rRow.txtRollCount = new TextField();
        rRow.txtRollCount.setStyle("-fx-font-size: 11px;");
        rRow.rollCountManualOverride = false;
        rRow.suppressRollCountAutoListener = false;
        bindRawRowRollCountManualEdit(rRow);
        rRow.grid.add(lblRollCount, 2, 8);
        addFormField(rRow.grid, rRow.txtRollCount, 3, 8);

        Button btnDelete = new Button("削除");
        btnDelete.getStyleClass().add("btn-settings-del");
        btnDelete.setStyle("-fx-font-size: 11px; -fx-padding: 3px 8px; -fx-cursor: hand;");
        btnDelete.setOnAction(evt -> {
            rawRows.remove(rRow);
            rawRowsContainer.getChildren().remove(rRow.grid);
            if (!rawRows.isEmpty()) {
                newTxtRawItem = rawRows.get(0).txtItem;
                newTxtRawPart = rawRows.get(0).txtPart;
                newTxtRawType = rawRows.get(0).txtType;
                newTxtRawWidth = rawRows.get(0).txtWidth;
                newTxtRawLength = rawRows.get(0).txtLength;
                newTxtRawGrade = rawRows.get(0).txtGrade;
                newTxtRawColor = rawRows.get(0).txtColor;
                newTxtRawCategory = rawRows.get(0).txtCategory;
                newTxtRawQty = rawRows.get(0).txtQty;
                newCmbWariSu = rawRows.get(0).cmbWariSu;
                newCmbRawFeedLoc = rawRows.get(0).cmbFeedLoc;
                newCmbRawStorageLoc = rawRows.get(0).cmbStorageLoc;
            }
        });
        rRow.grid.add(btnDelete, 3, 9);
        GridPane.setHalignment(btnDelete, javafx.geometry.HPos.RIGHT);

        Runnable updateSpec = () -> updateRawRowSpecDisplay(rRow);
        rRow.txtHinmei.textProperty().addListener((obs, oldV, newV) -> updateRowRawCandidates(rRow, false));
        rRow.txtPart.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            updateSpec.run();
        });
        rRow.txtItem.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            syncSelectedCandidatePaneWithItemCode(rRow.cmbSearch, rRow.selectedCandidatePane, newV);
        });
        rRow.txtType.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            updateSpec.run();
        });
        rRow.txtWidth.textProperty().addListener((obs, oldV, newV) -> updateSpec.run());
        rRow.txtLength.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            updateSpec.run();
        });
        rRow.txtQty.textProperty().addListener((obs, oldV, newV) -> updateRawRowRollCountDisplay(rRow));
        
        rRow.cmbSearch.setOnAction(
                evt ->
                        confirmMasterCandidateSelection(
                                rRow.cmbSearch.getValue(), rRow.txtItem, rRow.cmbSearch, rRow.selectedCandidatePane));

        rawRows.add(rRow);
        rawRowsContainer.getChildren().add(rRow.grid);

        if (rawRows.size() == 1) {
            newTxtRawItem = rRow.txtItem;
            newTxtRawPart = rRow.txtPart;
            newTxtRawType = rRow.txtType;
            newTxtRawWidth = rRow.txtWidth;
            newTxtRawLength = rRow.txtLength;
            newTxtRawGrade = rRow.txtGrade;
            newTxtRawColor = rRow.txtColor;
            newTxtRawCategory = rRow.txtCategory;
            newTxtRawQty = rRow.txtQty;
            newCmbWariSu = rRow.cmbWariSu;
            newCmbRawFeedLoc = rRow.cmbFeedLoc;
            newCmbRawStorageLoc = rRow.cmbStorageLoc;
        }
        updateRawRowRollCountDisplay(rRow);
        return rRow;
    }
}
