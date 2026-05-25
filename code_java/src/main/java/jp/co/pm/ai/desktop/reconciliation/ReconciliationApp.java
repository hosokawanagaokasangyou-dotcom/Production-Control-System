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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

public class ReconciliationApp {

    private static final String ALADDIN_MASTER_DIR_NAME = AppPaths.ALADDIN_MASTER_DIR_LEAF_NAME;
    private static final String INTEGRATED_MASTER_FILE_NAME = "マスタリレーション統合結果.xlsx";
    private static final String COL_MASTER_BASE_SHOHIN_PRODUCT = "masterBase商品(製品)";
    private static final String COL_MASTER_BASE_SHOHIN_RAW = "masterBase商品(原反)";
    /** 受注ﾌｧｲﾙ: ヘッダ行（0-based）。 */
    private static final int JUCHU_SHEET_HEADER_ROW_INDEX = 2;
    /** 受注ﾌｧｲﾙ: 見出し行直下の先頭データ行（0-based。新規行は常にここへ挿入）。 */
    private static final int JUCHU_SHEET_FIRST_DATA_ROW_INDEX = 3;
    /** 受注ﾌｧｲﾙ: POI lastRowNum が書式だけで膨らんだときの最大走査行数。 */
    private static final int JUCHU_SHEET_MAX_SCAN_ROWS = 20_000;
    private static final Path SETTINGS_FILE =
            Path.of(System.getProperty("user.home"), ".pm-ai-desktop", "request-form-reconciliation.properties");
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
    private static final long ORIGINAL_FILE_POLL_INTERVAL_MS = 5L * 60L * 1000L;
    /** {@link #saveLocalForm()} 後・受注ファイル未転記のステータス。 */
    private static final String STATUS_LOCAL_SAVE_PENDING = "手修正済み (未保存)";

    private Window hostWindow;
    private Path repoRootHint;
    private Path aladdinMasterDir;
    private Map<String, String> uiEnvSnapshot = Map.of();
    private Consumer<String> originalDirChangeHandler;
    private Consumer<String> juchuFileChangeHandler;
    private TextField txtJuchuPathDisplay;
    private Button btnTransfer;
    private Button btnBulkTransferPending;
    private Label transferBlockedReasonLabel;
    private String targetFolder = "";
    private String juchuFilePath;
    private boolean isLoadingRecord = false;
    /** 受注ファイル転記中（UI スレッドをブロックしないようバックグラウンド実行）。 */
    private volatile boolean juchuTransferInProgress = false;

    private ComboBox<OrderRecord> comboRecord;
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
    private javafx.animation.Timeline originalFilePollTimeline;
    private Supplier<RequestFormPreviewBadgeConfig> previewBadgeConfigSupplier =
            RequestFormPreviewBadgeConfig::defaults;
    
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
    // private TextField txtContract; // unified with newTxtFormContractNo
    
    private OrderRecord selectedRecord;

    // Master Product list for cache search
    
    // Dynamic multiple rows VBox and list properties
    private VBox productRowsContainer;
    private VBox rawRowsContainer;
    private final List<ProductRow> productRows = new ArrayList<>();
    private final List<RawMaterialRow> rawRows = new ArrayList<>();
private final List<ProductInfo> masterProductList = new ArrayList<>();

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
    private ComboBox<String> newCmbFormInputTanto;
         // 用途
    private ComboBox<String> newCmbFormUser;         // ユーザー
    private ComboBox<String> newCmbInputKbn;         // 入力区分 (header)
    private ComboBox<String> newCmbKakoKbn;          // 加工区分 (header)
    private ComboBox<String> newCmbInputTanto;       // 入力担当 (header)
    private ComboBox<String> newCmbWariSu;           // 割数 (product row)

    private TextField newTxtFormTokki1;
    private TextField newTxtFormTokki2;
    private TextField newTxtFormTokki3;
    private DatePicker newDpFormDeliv;
    private DatePicker newDpFormAdjustDeliv;
    private DatePicker newDpFormInputDate;
    private TextField newTxtFormWage;
    private TextField newTxtFormContractNo;
    
    private TextField newTxtUketsukeNo;
    private TextField newTxtIraiNo;

    // Settings data - combo option lists (editable by user)
    private final ObservableList<String> optInputKbn    = FXCollections.observableArrayList("通常入力", "例外入力");
    private final ObservableList<String> optKakoKbn     = FXCollections.observableArrayList("後加工", "TPI");
    private final ObservableList<String> optInputTanto  = FXCollections.observableArrayList("古家", "図司", "砂田");
    private final ObservableList<String> optWariSu      = FXCollections.observableArrayList("1","2","3","5","6","7","8","9","10");
    private final ObservableList<String> optEcSide      = FXCollections.observableArrayList("Ｈ面","Ｑ面","両面","ｽﾗｲｽ面","ｽｷﾝ面","Ｑ面/-","Ｈ面/-","Ｑ面/Ｑ面/-","H面/H面/-");
    private final ObservableList<String> optTrimming    = FXCollections.observableArrayList("有","無","-");
    private final ObservableList<String> optFeedLoc     = FXCollections.observableArrayList("EC","SEC","ｽﾗｲｽ","ｽﾘｯﾄ","ｴﾝﾎﾞｽ","検反","融着");
    private final ObservableList<String> optStorageLoc  = FXCollections.observableArrayList("滋賀","湖南","滋賀/湖南","湖南/中央","山田","中山","中央湖東","湖南/滋賀","奥田");
    private final ObservableList<String> optYoto        = FXCollections.observableArrayList("W（自動車）","B（輸出）","Y（工材）","V（TPI）","A（TPI）","JR（屋根）","P（TPI）");
    private final ObservableList<String> optUser        = FXCollections.observableArrayList("自動転記","ｵｶﾓﾄ","ﾀﾂﾀ","共和ﾚｻﾞｰ","Scientex","共和興","ｻｶｲﾅｺﾞﾔ","ﾀﾞｲｳﾚ","在ｴﾙ","U4059","U5001","張家港","ｲｽﾞﾐ","盟和","高山産業","中央物産");

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
        // Load saved settings（UI 個人設定。環境変数はこの後に上書き）
        loadSettings();
        applyRepoRootAsWorkspaceIfPresent(repoRootHint);
        configureFromUiEnv(uiEnv);
        ensureJuchuPathDefault();

        // Load master product list in cache on startup
        loadMasterProductList();

        // --- TOP MENU BAR ---
        BorderPane root = new BorderPane();
        root.setPadding(new Insets(0)); // Flush against edges
        
        HBox topBar = new HBox(15);
        topBar.getStyleClass().add("top-bar");
        topBar.setAlignment(Pos.CENTER_LEFT);
        
        Label titleLabel = new Label("湖南工場 統合管理データベース (JavaFX版)");
        titleLabel.getStyleClass().add("top-title");
        
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

        topBar.getChildren().addAll(titleLabel, btnSelectFolder, btnReload, statusLabel);
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
        
        HBox filterBox = new HBox(5);
        filterBox.setAlignment(Pos.CENTER_LEFT);
        Label lblSearch = new Label("検索・絞り込み:");
        lblSearch.setStyle("-fx-font-weight: bold;");
        TextField txtSearch = new TextField();
        txtSearch.setPromptText("依頼No / ユーザー...");
        txtSearch.setPrefWidth(160 * UI_WIDTH_SCALE);
        txtSearch.textProperty().addListener((obs, oldVal, newVal) -> filterTable(newVal));
        filterBox.getChildren().addAll(lblSearch, txtSearch);
        
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
        
        leftContainer.getChildren().addAll(lblLeftTitle, filterBox, comboRecord);
        
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
        
        Label lblContract = new Label("契約Ｎｏ:");
        styleFormLabel(lblContract);
        newTxtFormContractNo = new TextField();
        newTxtFormContractNo.setStyle("-fx-font-size: 11px;");
        newTxtFormContractNo.setEditable(false);
        newTxtFormContractNo.setFocusTraversable(false);
        basicGrid.add(lblContract, 2, 2);
        addFormField(basicGrid, newTxtFormContractNo, 3, 2);

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
        newCmbFormInputKbn = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optInputKbn));
        newCmbFormInputKbn.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblInputKbn, 0, 0);
        addFormField(workGrid, newCmbFormInputKbn, 1, 0);

        Label lblKakoKbn = new Label("加工区分:");
        styleFormLabel(lblKakoKbn);
        newCmbFormKakoKbn = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optKakoKbn));
        newCmbFormKakoKbn.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblKakoKbn, 2, 0);
        addFormField(workGrid, newCmbFormKakoKbn, 3, 0);

        Label lblInputTanto = new Label("入力担当:");
        styleFormLabel(lblInputTanto);
        newCmbFormInputTanto = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optInputTanto));
        newCmbFormInputTanto.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblInputTanto, 0, 1);
        addFormField(workGrid, newCmbFormInputTanto, 1, 1);

        Label lblYoto = new Label("用途:");
        styleFormLabel(lblYoto);
        newCmbFormYoto = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optYoto));
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
        
        Button btnSaveLocal = new Button("手修正を一時保存");
        btnSaveLocal.setMaxWidth(Double.MAX_VALUE);
        btnSaveLocal.setStyle("-fx-font-weight: bold; -fx-cursor: hand; -fx-padding: 8px;");
        btnSaveLocal.getStyleClass().add("btn-save-local");
        btnSaveLocal.setOnAction(e -> saveLocalForm());
        
        btnTransfer = new Button("受注ファイルへ自動転記・更新");
        btnTransfer.setMaxWidth(Double.MAX_VALUE);
        btnTransfer.setStyle("-fx-font-weight: bold; -fx-cursor: hand; -fx-padding: 8px;");
        btnTransfer.getStyleClass().add("btn-transfer");
        btnTransfer.setOnAction(e -> transferToExcel());

        btnBulkTransferPending = new Button("一時保存分一括転記");
        btnBulkTransferPending.setMaxWidth(Double.MAX_VALUE);
        btnBulkTransferPending.setStyle("-fx-font-weight: bold; -fx-cursor: hand; -fx-padding: 8px;");
        btnBulkTransferPending.getStyleClass().add("btn-transfer");
        btnBulkTransferPending.setOnAction(e -> transferAllPendingLocalSaves());
        updateTransferButtonState();

        Button btnOpenJuchu = new Button("受注エクセルを開く");
        btnOpenJuchu.setMaxWidth(Double.MAX_VALUE);
        btnOpenJuchu.setStyle("-fx-font-weight: bold; -fx-cursor: hand; -fx-padding: 8px;");
        btnOpenJuchu.getStyleClass().add("btn-reload");
        btnOpenJuchu.setOnAction(evt -> {
            File currentFile = new File(juchuFilePath);
            if (!currentFile.exists()) {
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
        });
        
        HBox sideBtns = new HBox(10, btnSaveLocal, btnTransfer, btnBulkTransferPending, btnOpenJuchu);
        sideBtns.setAlignment(Pos.CENTER);
        HBox.setHgrow(btnSaveLocal, Priority.ALWAYS);
        HBox.setHgrow(btnTransfer, Priority.ALWAYS);
        HBox.setHgrow(btnBulkTransferPending, Priority.ALWAYS);
        HBox.setHgrow(btnOpenJuchu, Priority.ALWAYS);

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
        
        // Tab 3: Settings - edit ComboBox option lists
        Tab tabSettings = createSettingsTab();
        
        tabPane.getTabs().addAll(tabVerification, tabSettings);
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
        startOriginalFilePolling();

        // Load initial data
        reloadData();
        return mainStackPane;
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

    // --- TAB 3: SETTINGS (Edit ComboBox option lists) ---
    private Tab createSettingsTab() {
        Tab tab = new Tab("\u2699\ufe0f \u8a2d\u5b9a");
        tab.setClosable(false);

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

        GridPane editorsGrid = new GridPane();
        editorsGrid.setHgap(12);
        editorsGrid.setVgap(12);
        editorsGrid.setAlignment(Pos.TOP_LEFT);
        VBox[] editors =
                new VBox[] {
                    buildListEditor("\u5165\u529b\u533a\u5206", optInputKbn),
                    buildListEditor("\u52a0\u5de5\u533a\u5206", optKakoKbn),
                    buildListEditor("\u5165\u529b\u62c5\u5f53", optInputTanto),
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
        
        Button btnOpenJuchuFile = new Button("Excelで開く");
        btnOpenJuchuFile.setStyle("-fx-font-size: 11px; -fx-padding: 6px 12px;");
        btnOpenJuchuFile.getStyleClass().add("btn-transfer");
        
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
        
        btnOpenJuchuFile.setOnAction(evt -> {
            File currentFile = new File(juchuFilePath);
            if (!currentFile.exists()) {
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
        });
        
        HBox juchuBtnBox = new HBox(10, btnSelectJuchuFile, btnOpenJuchuFile);
        juchuBtnBox.setAlignment(Pos.CENTER_LEFT);
        
        juchuCard.getChildren().addAll(lblJuchuCardTitle, lblJuchuDesc, txtJuchuPathDisplay, juchuBtnBox);
        root.getChildren().add(juchuCard);

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
        
        Label lblToolStatus = new Label("待機中...");
        lblToolStatus.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: inherit;");
        
        btnRunTool.setOnAction(e -> {
            lblToolStatus.setText("生成実行中...");
            btnRunTool.setDisable(true);
            
            Thread thread = new Thread(() -> {
                try {
                    String scriptPath = resolveIntegratedMasterScript();
                    ProcessBuilder pb = new ProcessBuilder("python", scriptPath);
                    pb.directory(new File(targetFolder));
                    pb.environment().put("PM_AI_REQUEST_FORM_WORKSPACE", targetFolder);
                    pb.environment()
                            .put(
                                    AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                                    targetFolder);
                    pb.environment()
                            .put(
                                    AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                                    aladdinMasterDirectory().getAbsolutePath());
                    Process process = pb.start();
                    int exitCode = process.waitFor();
                    
                    javafx.application.Platform.runLater(() -> {
                        btnRunTool.setDisable(false);
                        if (exitCode == 0) {
                            loadMasterProductList();
                            lblToolStatus.setText("成功！「マスタリレーション統合結果.xlsx」が生成されました。");
                            Alert alert = new Alert(Alert.AlertType.INFORMATION);
                            alert.setTitle("統合成功");
                            alert.setHeaderText(null);
                            alert.setContentText("マスタリレーションの統合が完了し、「マスタリレーション統合結果.xlsx」を更新しました！");
                            alert.showAndWait();
                        } else {
                            lblToolStatus.setText("失敗 (終了コード: " + exitCode + ")");
                            Alert alert = new Alert(Alert.AlertType.ERROR);
                            alert.setTitle("統合失敗");
                            alert.setHeaderText(null);
                            alert.setContentText("統合処理中にエラーが発生しました。Pythonスクリプトおよびマスタファイルを確認してください。");
                            alert.showAndWait();
                        }
                    });
                } catch (Exception ex) {
                    javafx.application.Platform.runLater(() -> {
                        btnRunTool.setDisable(false);
                        lblToolStatus.setText("エラー: " + ex.getMessage());
                        Alert alert = new Alert(Alert.AlertType.ERROR);
                        alert.setTitle("実行エラー");
                        alert.setHeaderText(null);
                        alert.setContentText("プロセスの実行中に例外が発生しました: " + ex.getMessage());
                        alert.showAndWait();
                    });
                }
            });
            thread.setDaemon(true);
            thread.start();
        });
        
        toolCard.getChildren().addAll(lblToolTitle, lblToolDesc, btnRunTool, lblToolStatus);
        root.getChildren().add(toolCard);

        sp.setContent(root);
        tab.setContent(sp);
        return tab;
    }

    /**
     * Build a compact editor card for one ObservableList<String>.
     * Changes here update the ComboBoxes live (shared reference).
     */
    private VBox buildListEditor(String fieldName, ObservableList<String> items) {
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
        btnAdd.setOnAction(
                e -> {
                    String v = tfNew.getText().trim();
                    if (!v.isEmpty() && !items.contains(v)) {
                        items.add(v);
                        listView.setPrefHeight(Math.min(Math.max(items.size(), 2) * 26 + 2, 140));
                        tfNew.clear();
                    }
                });
        tfNew.setOnAction(btnAdd.getOnAction());

        Button btnDel = new Button("\u524a\u9664");
        btnDel.getStyleClass().add("btn-settings-del");
        btnDel.setMinWidth(Region.USE_PREF_SIZE);
        btnDel.setOnAction(
                e -> {
                    String sel = listView.getSelectionModel().getSelectedItem();
                    if (sel != null) {
                        items.remove(sel);
                        listView.setPrefHeight(Math.min(Math.max(items.size(), 2) * 26 + 2, 140));
                    }
                });

        addRow.getChildren().addAll(tfNew, btnAdd, btnDel);
        card.getChildren().addAll(lbl, listView, addRow);
        return card;
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
        if (!optYoto.isEmpty()) newCmbFormYoto.setValue(optYoto.get(0));
        if (!optUser.isEmpty()) newCmbFormUser.setValue(optUser.get(0));
        if (!optInputKbn.isEmpty()) newCmbFormInputKbn.setValue(optInputKbn.get(0));
        if (!optKakoKbn.isEmpty()) newCmbFormKakoKbn.setValue(optKakoKbn.get(0));
        if (!optInputTanto.isEmpty()) newCmbFormInputTanto.setValue(optInputTanto.get(0));
        newDpFormDeliv.setValue(null);
        newDpFormAdjustDeliv.setValue(null);
        newDpFormInputDate.setValue(null);
        newTxtFormWage.setText("9");
        newTxtFormContractNo.setText("");
        
        if (!optInputKbn.isEmpty()) newCmbInputKbn.setValue(optInputKbn.get(0));
        if (!optKakoKbn.isEmpty()) newCmbKakoKbn.setValue(optKakoKbn.get(0));
        if (!optInputTanto.isEmpty()) newCmbInputTanto.setValue(optInputTanto.get(0));
        newTxtUketsukeNo.setText("");
        newTxtIraiNo.setText("");
    }

    /** 環境変数タブの依頼書入力向けパスを反映する（タブ再選択時も可）。 */
    public void configureFromUiEnv(Map<String, String> uiEnv) {
        uiEnvSnapshot = uiEnv != null ? Map.copyOf(uiEnv) : Map.of();
        RequestFormSheetPreviewPdfRenderer.applyCjkMetricsScaleFromUi(uiEnvSnapshot);
        aladdinMasterDir = AppPaths.resolveAladdinMasterDir(uiEnvSnapshot);
        applyRequestFormOriginalDirFromUiEnv();
        applyJuchuFilePathFromUiEnv();
        refreshJuchuPathDisplay();
        updateTransferButtonState();
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

    private boolean isJuchuFileLockedForWrite() {
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            return true;
        }
        File file = new File(juchuFilePath);
        if (!file.isFile()) {
            return true;
        }
        File lockFile = excelLockFileFor(file);
        return lockFile != null && lockFile.isFile();
    }

    /**
     * 自動転記ボタンを無効化する理由。転記可能なら {@code null}。
     * JavaFX では disable 中のボタンに Tooltip が表示されないため、{@link #updateTransferButtonState()} でラベルにも出す。
     */
    private String resolveTransferBlockedReason() {
        if (juchuTransferInProgress) {
            return "受注ファイルへの転記処理を実行中です。完了までお待ちください。";
        }
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            return "受注ファイルが未設定です。設定タブまたは環境変数 PM_AI_REQUEST_FORM_JUCHU_FILE を指定してください。";
        }
        File juchuFile = new File(juchuFilePath);
        if (!juchuFile.isFile()) {
            return "受注ファイルが見つかりません: " + juchuFilePath;
        }
        File lockFile = excelLockFileFor(juchuFile);
        if (lockFile != null && lockFile.isFile()) {
            return "受注ファイルが Excel で使用中（読み取り専用）です。保存して閉じてから再試行してください。"
                    + " ロック: "
                    + lockFile.getName();
        }
        return null;
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
            File juchuFile = new File(juchuFilePath);
            File lockFile = excelLockFileFor(juchuFile);
            String lockHint =
                    lockFile != null && lockFile.isFile()
                            ? "\nロックファイル: " + lockFile.getAbsolutePath()
                            : "";
            showAlert(
                    "エラー",
                    "受注ファイルが読み取り専用（使用中）のため転記できません。"
                            + " Excel で開いている場合は保存して閉じてから再試行してください。"
                            + lockHint);
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
        if (!confirmJuchuHeaderWarnings(headerWarnings)) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }

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
                                try (FileInputStream fis = new FileInputStream(file);
                                        Workbook wb = WorkbookFactory.create(fis)) {
                                    Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
                                    Map<String, Integer> colMap = buildJuchuColumnMap(sheet);
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
                                    try (FileOutputStream fos = new FileOutputStream(file)) {
                                        wb.write(fos);
                                    }
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
        transferAllPendingLocalSavesAsync(true, null);
    }

    private static Map<String, Integer> buildJuchuColumnMap(Sheet sheet) {
        Row hRow = sheet.getRow(JUCHU_SHEET_HEADER_ROW_INDEX);
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
            targetRowIndex = insertNewJuchuDataRowPreservingFormulas(sheet, lastDataRowIndex, null);
            targetRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
            applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);
        }

        setJuchuSheetReqNo(wb, sheet, targetRow, reqNo);

        Map<String, String> mergedDb = new LinkedHashMap<>(db);
        if (isNewRow) {
            Map<String, String> raw = record.getRawValues();
            if (raw != null && !raw.isEmpty()) {
                mergeJuchuDbFromRawDefaults(mergedDb, raw);
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
                isNewRow);

        if (isNewRow) {
            Map<String, String> raw = record.getRawValues();
            if (raw != null && !raw.isEmpty()) {
                setJuchuNumericOrStringByLayout(
                        targetRow,
                        JuchuSheetColumnLayout.Col.KAKOCHIN,
                        firstNonBlank(mergedDb.get("加工賃"), raw.get("加工賃")));
                setJuchuNumericOrStringByLayout(
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
        filterTable("");
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
        loadMasterProductList();
    }

    private void loadMasterProductList() {
        masterProductList.clear();
        File masterFile = integratedMasterFile();
        if (!masterFile.exists()) {
            System.err.println("Integrated master file not found: " + masterFile.getAbsolutePath());
            return;
        }
        
        try (FileInputStream fis = new FileInputStream(masterFile);
             Workbook wb = WorkbookFactory.create(fis)) {
            Sheet sheet = wb.getSheet("②商品別・工程展開リスト");
            if (sheet == null) {
                System.err.println("Sheet ②商品別・工程展開リスト not found!");
                return;
            }
            
            Row hRow = sheet.getRow(0);
            Map<String, Integer> colMap = new HashMap<>();
            for (int c = 0; c < hRow.getLastCellNum(); c++) {
                Cell cell = hRow.getCell(c);
                if (cell != null) {
                    colMap.put(cell.getStringCellValue().trim(), c);
                }
            }
            
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                
                Cell codeCell = row.getCell(colMap.getOrDefault("商品コード", 0));
                String shohinCode = getCellValueAsString(codeCell);
                if (shohinCode.isEmpty()) continue;
                
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
                        if (stepVal != null && !stepVal.trim().isEmpty() && !stepVal.equalsIgnoreCase("None")) {
                            steps.add(stepVal.trim());
                        }
                    }
                }
                String kakoNaiyo = String.join(",", steps);
                
                masterProductList.add(new ProductInfo(
                    shohinCode, seihinCode, shohinName1, shohinName2,
                    unitName, quantityPerCase, selfKakoKbn, foamName,
                    foamPartNo, foamWidth, foamLength, foamColor, foamThickness,
                    kakoNaiyo
                ));
            }
            System.out.println("Loaded " + masterProductList.size() + " products into lookup cache from Integrated Master.");
        } catch (Exception e) {
            System.err.println("Error loading integrated product master: " + e.getMessage());
            e.printStackTrace();
        }
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
            
            FileInputStream fis = new FileInputStream(file);
            Workbook wb = WorkbookFactory.create(fis);
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
                    showAlert("エラー", "この依頼Ｎｏ（" + reqNo + "）は既に存在します。別の番号を指定してください。");
                    wb.close();
                    fis.close();
                    return;
            }
            
            int targetRowIndex = insertNewJuchuDataRowPreservingFormulas(sheet, lastDataRowIndex, null);
            Row targetRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
            applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);

            setJuchuSheetReqNo(wb, sheet, targetRow, reqNo);

            Map<String, String> db = buildJuchuDbValuesFromForm();
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
                    true);
            
            fis.close();
            FileOutputStream fos = new FileOutputStream(file);
            wb.write(fos);
            fos.close();
            wb.close();
            
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
                Workbook wbJuchu = WorkbookFactory.create(fis);
                Sheet sJuchu = wbJuchu.getSheet("受注ﾌｧｲﾙ");
                Row hRow = sJuchu.getRow(JUCHU_SHEET_HEADER_ROW_INDEX);
                List<String> headerWarnings = JuchuSheetColumnLayout.validateHeaders(hRow);
                headerWarningsFinal = headerWarnings;

                Map<String, Map<String, String>> dbRows = new HashMap<>();
                int lastDataRowIndex = findJuchuSheetLastPopulatedDataRowIndex(sJuchu);
                
                for (int r = JUCHU_SHEET_FIRST_DATA_ROW_INDEX; r <= lastDataRowIndex; r++) {
                    Row row = sJuchu.getRow(r);
                    if (row == null) continue;
                    Cell reqCell = row.getCell(0); 
                    if (reqCell == null || reqCell.getCellType() == CellType.BLANK) continue;
                    
                    String reqNo = getCellValueAsString(reqCell).trim();
                    if (reqNo.isEmpty()) continue;
                    
                    Map<String, String> vals = JuchuSheetColumnLayout.readDbValuesFromRow(row);
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
                    filterTable("");  // ファイルが存在するレコードのみ表示
                    int visibleCount = (int) orderRecords.stream().filter(r -> hasExistingFile(r)).count();
                    statusLabel.setText(String.format("読込完了: 全 %d 件中 %d 件 (依頼書あり)", orderRecords.size(), visibleCount));
                    if (!finalHeaderWarnings.isEmpty()) {
                        statusLabel.setText(
                                statusLabel.getText()
                                        + " / 列定義警告 "
                                        + finalHeaderWarnings.size()
                                        + " 件");
                        confirmJuchuHeaderWarnings(finalHeaderWarnings);
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
                Workbook wbRaw = WorkbookFactory.create(fisRaw)) {
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

    private void filterTable(String query) {
        ObservableList<OrderRecord> base = FXCollections.observableArrayList();
        for (OrderRecord rec : orderRecords) {
            if (hasExistingFile(rec)) base.add(rec);
        }
        
        if (query == null || query.isEmpty()) {
            comboRecord.setItems(base);
            return;
        }
        
        ObservableList<OrderRecord> filtered = FXCollections.observableArrayList();
        String q = query.toLowerCase();
        for (OrderRecord rec : base) {
            if (rec.getReqNo().toLowerCase().contains(q) || rec.getUser().toLowerCase().contains(q)) {
                filtered.add(rec);
            }
        }
        comboRecord.setItems(filtered);
    }

    // --- LOGIC: RENDER ORIGINAL SHEET VIEW AND FILL FORM ---
    private void loadRecordDetails(OrderRecord record) {
        this.selectedRecord = record;
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
        } else {
            activeVals = record.getDbValues();
        }
        
        newCmbFormUser.setValue(activeVals.getOrDefault("ユーザー", ""));
        newDpFormDeliv.setValue(parseLocalDate(activeVals.getOrDefault("希望納期", "")));
        newDpFormAdjustDeliv.setValue(parseLocalDate(activeVals.getOrDefault("調整納期", "")));
        newDpFormInputDate.setValue(parseLocalDate(activeVals.getOrDefault("入力日", "")));
        txtProcess.setText(activeVals.getOrDefault("加工内容", ""));
        newTxtFormWage.setText(activeVals.getOrDefault("加工賃", ""));
        newTxtFormContractNo.setText(
                firstNonBlank(
                        rawVals.get("契約Ｎｏ"),
                        activeVals.get("契約Ｎｏ"),
                        activeVals.get("契約No")));
        newCmbFormYoto.setValue(activeVals.getOrDefault("用途", ""));
        newCmbFormInputKbn.setValue(activeVals.getOrDefault("入力区分", ""));
        newCmbFormKakoKbn.setValue(activeVals.getOrDefault("加工区分", ""));
        newCmbFormInputTanto.setValue(activeVals.getOrDefault("入力担当", ""));
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
            sheetGrid.getChildren().clear();
            currentPreviewOriginalFile = null;
            refreshPreviewFileHeader();
        }
        enqueueBackgroundCacheTasks();
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

    private void startOriginalFilePolling() {
        originalUpdateMonitor.setOnUpdatedKeysChanged(
                keys -> Platform.runLater(this::refreshPreviewFileHeader));
        if (originalFilePollTimeline != null) {
            originalFilePollTimeline.stop();
        }
        originalFilePollTimeline =
                new javafx.animation.Timeline(
                        new javafx.animation.KeyFrame(
                                javafx.util.Duration.millis(ORIGINAL_FILE_POLL_INTERVAL_MS),
                                e -> pollOriginalFilesInWorkspace()));
        originalFilePollTimeline.setCycleCount(javafx.animation.Timeline.INDEFINITE);
        originalFilePollTimeline.play();
        pollOriginalFilesInWorkspace();
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
        if (selectedRecord == null) {
            showAlert("エラー", "修正するレコードを選択してください。");
            return;
        }
        
        Map<String, String> db = selectedRecord.getDbValues();
        db.clear();
        db.putAll(buildJuchuDbValuesFromForm());
        
        selectedRecord.setStatus(STATUS_LOCAL_SAVE_PENDING);
        
        // ComboBox表示のリフレッシュ
        comboRecord.setItems(null);
        comboRecord.setItems(orderRecords);
        comboRecord.getSelectionModel().select(selectedRecord);
        
        statusLabel.setText("修正をローカル保存しました（受注ファイルに反映させるには転記ボタンを押してください）");
    }

    private void transferToExcel() {
        if (juchuTransferInProgress) {
            return;
        }

        String reqNo = txtReqNo.getText().trim();
        if (reqNo.isEmpty()) {
            showAlert("エラー", "依頼Ｎｏを入力してください。");
            return;
        }

        if (isJuchuFileLockedForWrite()) {
            File juchuFile = new File(juchuFilePath);
            File lockFile = excelLockFileFor(juchuFile);
            String lockHint =
                    lockFile != null && lockFile.isFile()
                            ? "\nロックファイル: " + lockFile.getAbsolutePath()
                            : "";
            showAlert(
                    "エラー",
                    "受注ファイルが読み取り専用（使用中）のため転記できません。"
                            + " Excel で開いている場合は保存して閉じてから再試行してください。"
                            + lockHint);
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
        if (!confirmJuchuHeaderWarnings(headerWarnings)) {
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
                            try {
                                performSingleJuchuTransfer(file, formData, this::updateLoadingOverlayText);
                            } catch (Exception ex) {
                                error = ex;
                            }

                            final Exception finalError = error;
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

                                            if (recordRef != null) {
                                                recordRef.setStatus("一致 (転記完了)");
                                                recordRef.setDiscrepancy("一致 (受注ファイルへ転記完了)");

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

    private JuchuTransferFormData captureJuchuTransferFormData(String reqNo) {
        Map<String, String> rawValues =
                selectedRecord != null
                        ? new HashMap<>(selectedRecord.getRawValues())
                        : Map.of();
        return new JuchuTransferFormData(reqNo, buildJuchuDbValuesFromForm(), rawValues);
    }

    private List<String> readJuchuHeaderWarnings(File file) {
        try (FileInputStream fis = new FileInputStream(file);
                Workbook wb = WorkbookFactory.create(fis)) {
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
            if (sheet == null) {
                return List.of("受注ﾌｧｲﾙ シートが見つかりません。");
            }
            Row hRow = sheet.getRow(JUCHU_SHEET_HEADER_ROW_INDEX);
            return JuchuSheetColumnLayout.validateHeaders(hRow);
        } catch (Exception ex) {
            return List.of("受注ファイルの見出し検証に失敗: " + ex.getMessage());
        }
    }

    private void performSingleJuchuTransfer(
            File file, JuchuTransferFormData form, Consumer<String> progress) throws Exception {
        progress.accept(
                "受注ファイルへ転記しています…\n(1/5) 受注ファイルを開いています…\n依頼No: " + form.reqNo());

        try (FileInputStream fis = new FileInputStream(file);
                Workbook wb = WorkbookFactory.create(fis)) {
            wb.setForceFormulaRecalculation(false);
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");

            Map<String, Integer> colMap = buildJuchuColumnMap(sheet);

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

            if (isNewRow) {
                progress.accept(
                        "受注ファイルへ転記しています…\n(2/5) 見出し直下(先頭)に新規行を挿入しています…\n依頼No: "
                                + form.reqNo());
                targetRowIndex = insertNewJuchuDataRowPreservingFormulas(sheet, lastDataRowIndex, progress);
                targetRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
                applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);
            }

            progress.accept(
                    "受注ファイルへ転記しています…\n(3/5) セルへ転記しています…\n依頼No: "
                            + form.reqNo()
                            + (isNewRow ? "（新規行・先頭）" : "（既存行）"));

            setJuchuSheetReqNo(wb, sheet, targetRow, form.reqNo());

            Map<String, String> mergedDb = new LinkedHashMap<>(form.dbValues());
            if (isNewRow && form.rawValuesForNewRow() != null && !form.rawValuesForNewRow().isEmpty()) {
                mergeJuchuDbFromRawDefaults(mergedDb, form.rawValuesForNewRow());
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
                    isNewRow);

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
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
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
     * 受注ﾌｧｲﾙの見出し行直下（先頭データ行）へ行を挿入し、テンプレート行から数式・書式を複製する（値は複製しない）。
     *
     * @param lastPopulatedDataRowIndex {@link #findJuchuSheetLastPopulatedDataRowIndex} の結果（実データ末尾）
     * @param progress 進捗表示（{@code null} 可）
     * @return Excel 上の 1 始まり行番号（数式の行参照用）
     */
    private int insertNewJuchuDataRowPreservingFormulas(
            Sheet sheet, int lastPopulatedDataRowIndex, Consumer<String> progress) {
        final int destRowIdx = JUCHU_SHEET_FIRST_DATA_ROW_INDEX;
        int templateRowIdx;
        if (lastPopulatedDataRowIndex >= destRowIdx) {
            int shiftCount = lastPopulatedDataRowIndex - destRowIdx + 1;
            if (progress != null) {
                progress.accept(
                        String.format(
                                "受注ファイルへ転記しています…\n(2/5) 見出し直下へ挿入のため既存 %d 行を下へずらしています…\n（行数が多いと 1～2 分かかることがあります）",
                                shiftCount));
            }
            sheet.shiftRows(destRowIdx, lastPopulatedDataRowIndex, 1, true, true);
            templateRowIdx = destRowIdx + 1;
        } else {
            if (progress != null) {
                progress.accept(
                        "受注ファイルへ転記しています…\n(2/5) 見出し直下(先頭)に新規行を確保しています…");
            }
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
        return destRowIdx + 1;
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
        int first = JUCHU_SHEET_FIRST_DATA_ROW_INDEX;
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
        if (lastDataRowIndex < JUCHU_SHEET_FIRST_DATA_ROW_INDEX) {
            return -1;
        }
        for (int r = JUCHU_SHEET_FIRST_DATA_ROW_INDEX; r <= lastDataRowIndex; r++) {
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
    private static void setJuchuSheetReqNo(Workbook wb, Sheet sheet, Row targetRow, String reqNo) {
        if (targetRow == null || reqNo == null) {
            return;
        }
        Cell cellReqNo = targetRow.getCell(0);
        if (cellReqNo == null) {
            cellReqNo = targetRow.createCell(0);
        }
        cellReqNo.setCellValue(reqNo);
        CellStyle yellowStyle = wb.createCellStyle();
        Row refRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX + 1);
        if (refRow == null) {
            refRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
        }
        if (refRow != null && refRow.getCell(0) != null) {
            yellowStyle.cloneStyleFrom(refRow.getCell(0).getCellStyle());
        }
        yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellReqNo.setCellStyle(yellowStyle);
    }

    /** 数式セルは上書きしない。 */
    private static Cell writableJuchuCell(Row row, int col) {
        Cell existing = row.getCell(col);
        if (isJuchuFormulaCell(existing)) {
            return null;
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

    private boolean confirmJuchuHeaderWarnings(List<String> warnings) {
        if (warnings == null || warnings.isEmpty()) {
            return true;
        }
        Alert alert = new Alert(Alert.AlertType.WARNING);
        alert.setTitle("受注シート列定義の警告");
        alert.setHeaderText("見出し行（行3）と列位置の定義が一致しない列があります。");
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
            boolean setInputDateNow) {
        if (targetRow == null || db == null) {
            return;
        }
        setJuchuCellByLayout(
                targetRow,
                JuchuSheetColumnLayout.Col.NYURYOKU_KBN,
                resolveWorkFieldValue(inputKbn, isNewRow ? "通常入力" : ""));
        setJuchuCellByLayout(
                targetRow,
                JuchuSheetColumnLayout.Col.KAKO_KBN,
                resolveWorkFieldValue(kakoKbn, isNewRow ? "後加工" : ""));
        setJuchuCellByLayout(
                targetRow,
                JuchuSheetColumnLayout.Col.NYURYOKU_TANTO,
                resolveWorkFieldValue(inputTanto, isNewRow ? "自動転記" : ""));
        if (setInputDateNow) {
            setJuchuDateByLayout(targetRow, JuchuSheetColumnLayout.Col.NYURYOKU_BI, new Date());
        }

        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.HINMEI, db.getOrDefault("品名", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.SEIHIN, db.getOrDefault("製品", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.KON_TO_1, db.getOrDefault("梱-等1", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.IRO_1, db.getOrDefault("色1", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.KUBUN_1, db.getOrDefault("区分1", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.EDABAN, db.getOrDefault("枝番", ""));
        setJuchuNumericOrStringByLayout(targetRow, JuchuSheetColumnLayout.Col.SURYO_1, db.getOrDefault("数量1", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.EC_MEN, db.getOrDefault("ＥＣ面", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.TRIMMING, db.getOrDefault("ﾄﾘﾐﾝｸﾞ", ""));

        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.WARISU, db.getOrDefault("割数", ""));
        setJuchuCellByLayout(
                targetRow,
                JuchuSheetColumnLayout.Col.HINMEI_1,
                firstNonBlank(db.get("品名1"), db.get("原反品名")));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.GENPAN, db.getOrDefault("原反", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.KON_TO, db.getOrDefault("原反梱-等", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.IRO, db.getOrDefault("原反色", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.KUBUN, db.getOrDefault("原反区分", ""));
        setJuchuNumericOrStringByLayout(targetRow, JuchuSheetColumnLayout.Col.SURYO, db.getOrDefault("原反数量", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.ZAIKO_BASHO, db.getOrDefault("在庫場所", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.TONYU_BASHO, db.getOrDefault("投入場所", ""));

        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.KAKO_NAIYO, db.getOrDefault("加工内容", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.TOKKI_1, tokki1 != null ? tokki1 : db.getOrDefault("特記事項1", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.TOKKI_2, tokki2 != null ? tokki2 : db.getOrDefault("特記事項2", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.TOKKI_3, tokki3 != null ? tokki3 : db.getOrDefault("特記事項3", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.YOTO, db.getOrDefault("用途", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.USER, db.getOrDefault("ユーザー", ""));
        setJuchuDateOrStringByLayout(targetRow, JuchuSheetColumnLayout.Col.KIBO_NOKI, db.getOrDefault("希望納期", ""));
        setJuchuCellByLayout(targetRow, JuchuSheetColumnLayout.Col.CHOSEI_NOKI, db.getOrDefault("調整納期", ""));
        setJuchuNumericOrStringByLayout(targetRow, JuchuSheetColumnLayout.Col.KAKOCHIN, db.getOrDefault("加工賃", ""));

        setJuchuCellByLayout(
                targetRow,
                JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT,
                db.getOrDefault(COL_MASTER_BASE_SHOHIN_PRODUCT, ""));
        setJuchuCellByLayout(
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
        db.put("用途", newCmbFormYoto.getValue() != null ? newCmbFormYoto.getValue().trim() : "");
        db.put(
                "入力区分",
                newCmbFormInputKbn.getValue() != null ? newCmbFormInputKbn.getValue().trim() : "");
        db.put(
                "加工区分",
                newCmbFormKakoKbn.getValue() != null ? newCmbFormKakoKbn.getValue().trim() : "");
        db.put(
                "入力担当",
                newCmbFormInputTanto.getValue() != null ? newCmbFormInputTanto.getValue().trim() : "");
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
        return text.toUpperCase();
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
        Properties props = new Properties();
        Path settingsPath = SETTINGS_FILE;
        if (Files.isRegularFile(settingsPath)) {
            try (FileInputStream fis = new FileInputStream(settingsPath.toFile())) {
                props.load(fis);
                String folder = props.getProperty("targetFolder");
                if (folder != null && new File(folder).exists()) {
                    targetFolder = folder;
                }
                String juchuPath = props.getProperty("juchuFilePath");
                if (juchuPath != null && new File(juchuPath).exists()) {
                    juchuFilePath = juchuPath;
                }
            } catch (Exception e) {
                System.err.println("Could not load settings: " + e.getMessage());
            }
        }
    }

    private void saveSettings() {
        Properties props = new Properties();
        props.setProperty("targetFolder", targetFolder);
        props.setProperty("juchuFilePath", juchuFilePath != null ? juchuFilePath : "");
        try {
            Files.createDirectories(SETTINGS_FILE.getParent());
            try (FileOutputStream fos = new FileOutputStream(SETTINGS_FILE.toFile())) {
                props.store(fos, "Request form reconciliation settings");
            }
        } catch (Exception e) {
            System.err.println("Could not save settings: " + e.getMessage());
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

    private static void configureSplitFieldRow(HBox box, Region primary, Region secondary) {
        box.setAlignment(Pos.CENTER_LEFT);
        box.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(primary, Priority.ALWAYS);
        primary.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(secondary, Priority.SOMETIMES);
        secondary.setMaxWidth(Double.MAX_VALUE);
    }

    /** 製品行の「商品」欄とマスタ候補から自動転記された関連欄をクリアする。 */
    private void clearProductRowShohin(ProductRow pRow) {
        isLoadingRecord = true;
        try {
            pRow.txtItem.clear();
            pRow.txtPart.clear();
            pRow.txtType.clear();
            pRow.txtWidth.clear();
            pRow.txtLength.clear();
            pRow.cmbSearch.getSelectionModel().clearSelection();
            pRow.cmbSearch.setValue(null);
            pRow.cmbSearch.setItems(FXCollections.emptyObservableList());
            pRow.cmbSearch.hide();
        } finally {
            isLoadingRecord = false;
        }
    }

    /** 原反行の「商品」欄とマスタ候補から自動転記された関連欄をクリアする。 */
    private void clearRawMaterialRowShohin(RawMaterialRow rRow) {
        isLoadingRecord = true;
        try {
            rRow.txtItem.clear();
            rRow.txtPart.clear();
            rRow.txtType.clear();
            rRow.txtWidth.clear();
            rRow.txtLength.clear();
            rRow.cmbSearch.getSelectionModel().clearSelection();
            rRow.cmbSearch.setValue(null);
            rRow.cmbSearch.setItems(FXCollections.emptyObservableList());
            rRow.cmbSearch.hide();
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

    // --- STYLISH DYNAMIC MULTIPLE ROWS HELPERS & CLASSES ---
    
    public static class ProductRow {
        public TextField txtHinmei;
        public TextField txtSeihinmei;
        public TextField txtItem;
        public TextField txtPart;
        public ComboBox<String> cmbSearch;
        public TextField txtType;
        public TextField txtWidth;
        public TextField txtLength;
        public TextField txtQty;
        public TextField txtGrade;
        public TextField txtColor;
        public TextField txtCategory;
        public TextField txtEdaban;
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
        pRow.txtHinmei.setEditable(false);
        pRow.txtHinmei.setFocusTraversable(false);
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
        pRow.cmbSearch.setStyle("-fx-font-size: 11px;");
        pRow.cmbSearch.setPromptText("候補");
        pRow.cmbSearch.setMinWidth(96);
        pRow.cmbSearch.setPrefWidth(128);
        wireCandidateComboBox(pRow.cmbSearch, () -> updateRowProdCandidates(pRow, true));
        HBox partBox = new HBox(5, pRow.txtPart, pRow.cmbSearch);
        configureSplitFieldRow(partBox, pRow.txtPart, pRow.cmbSearch);
        pRow.grid.add(lblPart, 0, 2);
        addFormField(pRow.grid, partBox, 1, 2, 3, 1);

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

        Label lblEc = new Label("EC面:");
        styleFormLabel(lblEc);
        pRow.cmbEcSide = new ComboBox<>(FXCollections.observableArrayList(optEcSide));
        pRow.cmbEcSide.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblEc, 0, 7);
        addFormField(pRow.grid, pRow.cmbEcSide, 1, 7);

        Label lblTrimming = new Label("トリミング:");
        styleFormLabel(lblTrimming);
        pRow.cmbTrimming = new ComboBox<>(FXCollections.observableArrayList(optTrimming));
        pRow.cmbTrimming.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblTrimming, 2, 7);
        addFormField(pRow.grid, pRow.cmbTrimming, 3, 7);

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
        pRow.grid.add(btnDelete, 3, 8);
        GridPane.setHalignment(btnDelete, javafx.geometry.HPos.RIGHT);

        Runnable updateSpec = () -> updateProductRowSpecDisplay(pRow);
        pRow.txtPart.textProperty().addListener((obs, oldV, newV) -> {
            updateRowProdCandidates(pRow, false);
            updateSpec.run();
        });
        pRow.txtItem.textProperty().addListener((obs, oldV, newV) -> updateRowProdCandidates(pRow, false));
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
            updateRowProdCandidates(pRow, false);
        }
        for (RawMaterialRow rRow : rawRows) {
            updateRowRawCandidates(rRow, false);
        }
    }

    private static final String MASTER_CANDIDATE_FILTER_KEYWORDS_PROP = "filterKeywords";

    private static void wireCandidateComboBox(ComboBox<String> combo, Runnable refreshOnOpen) {
        combo.setOnShowing(
                e -> {
                    if (refreshOnOpen != null) {
                        refreshOnOpen.run();
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
                                                item, masterCandidateFilterKeywords(combo), combo));
                            }
                        });
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
        String code = shohinCode.trim();
        for (ProductInfo p : masterProductList) {
            if (p.getShohinCode().equals(code)) {
                return p;
            }
        }
        return null;
    }

    /** マスタ候補コンボの選択を製品行とフォーム共通の加工内容へ反映する。 */
    private void applyMasterProductCandidateSelection(String candidateLabel, ProductRow pRow) {
        ProductInfo product = findMasterProductByShohinCode(shohinCodeFromMasterCandidateLabel(candidateLabel));
        if (product == null || pRow == null) {
            return;
        }
        isLoadingRecord = true;
        try {
            pRow.txtItem.setText(product.getShohinCode());
            pRow.txtPart.setText(product.getFoamPartNo());
            String[] nameParts = product.getShohinName1().split("-");
            if (nameParts.length >= 2) {
                pRow.txtType.setText(nameParts[1]);
            } else if (!product.getShohinName1().isBlank()) {
                pRow.txtType.setText(product.getShohinName1());
            }
            pRow.txtWidth.setText(product.getFoamWidth());
            pRow.txtLength.setText(product.getFoamLength());
            updateProductRowSpecDisplay(pRow);
            if (txtProcess != null) {
                txtProcess.setText(product.getKakoNaiyo());
            }
        } finally {
            isLoadingRecord = false;
        }
    }

    private void updateRowProdCandidates(ProductRow pRow, boolean fromDropdownOpen) {
        if (isLoadingRecord && !fromDropdownOpen) {
            return;
        }
        String kwItem = normalize_text(pRow.txtItem.getText());
        String kwPart = normalize_text(pRow.txtPart.getText());
        String kwType = normalize_text(pRow.txtType.getText());
        String kwLength = normalize_text(pRow.txtLength.getText());
        String kwHinmei = normalize_text(pRow.txtHinmei.getText());
        setMasterCandidateFilterKeywords(pRow.cmbSearch, kwItem, kwPart, kwType, kwLength, kwHinmei);

        java.util.List<String> filtered;
        if (kwItem.isEmpty()
                && kwPart.isEmpty()
                && kwType.isEmpty()
                && kwLength.isEmpty()
                && kwHinmei.isEmpty()) {
            if (!fromDropdownOpen) {
                pRow.cmbSearch.setItems(javafx.collections.FXCollections.emptyObservableList());
                return;
            }
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            masterProductList, "", "", "", "", "", 50);
        } else {
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            masterProductList, kwItem, kwPart, kwType, kwLength, kwHinmei, 50);
        }
        pRow.cmbSearch.setItems(javafx.collections.FXCollections.observableArrayList(filtered));
        if (!fromDropdownOpen && !filtered.isEmpty() && !pRow.cmbSearch.isShowing()) {
            pRow.cmbSearch.show();
        }
    }

    private void updateRowRawCandidates(RawMaterialRow rRow, boolean fromDropdownOpen) {
        if (isLoadingRecord && !fromDropdownOpen) {
            return;
        }
        String kwItem = normalize_text(rRow.txtItem.getText());
        String kwPart = normalize_text(rRow.txtPart.getText());
        String kwType = normalize_text(rRow.txtType.getText());
        String kwLength = normalize_text(rRow.txtLength.getText());
        String kwHinmei = normalize_text(rRow.txtHinmei.getText());
        setMasterCandidateFilterKeywords(rRow.cmbSearch, kwItem, kwPart, kwType, kwLength, kwHinmei);

        java.util.List<String> filtered;
        if (kwItem.isEmpty()
                && kwPart.isEmpty()
                && kwType.isEmpty()
                && kwLength.isEmpty()
                && kwHinmei.isEmpty()) {
            if (!fromDropdownOpen) {
                rRow.cmbSearch.setItems(javafx.collections.FXCollections.emptyObservableList());
                return;
            }
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            masterProductList, "", "", "", "", "", 50);
        } else {
            filtered =
                    RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                            masterProductList, kwItem, kwPart, kwType, kwLength, kwHinmei, 50);
        }
        rRow.cmbSearch.setItems(javafx.collections.FXCollections.observableArrayList(filtered));
        if (!fromDropdownOpen && !filtered.isEmpty() && !rRow.cmbSearch.isShowing()) {
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
        rRow.txtHinmei.setEditable(false);
        rRow.txtHinmei.setFocusTraversable(false);
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
        rRow.cmbSearch.setStyle("-fx-font-size: 11px;");
        rRow.cmbSearch.setPromptText("候補");
        rRow.cmbSearch.setMinWidth(96);
        rRow.cmbSearch.setPrefWidth(128);
        wireCandidateComboBox(rRow.cmbSearch, () -> updateRowRawCandidates(rRow, true));
        HBox partBox = new HBox(5, rRow.txtPart, rRow.cmbSearch);
        configureSplitFieldRow(partBox, rRow.txtPart, rRow.cmbSearch);
        rRow.grid.add(lblPart, 0, 2);
        addFormField(rRow.grid, partBox, 1, 2, 3, 1);

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
        rRow.cmbWariSu = new ComboBox<>(FXCollections.observableArrayList(optWariSu));
        rRow.cmbWariSu.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblWariSu, 2, 6);
        addFormField(rRow.grid, rRow.cmbWariSu, 3, 6);

        Label lblFeed = new Label("投入場所:");
        styleFormLabel(lblFeed);
        rRow.cmbFeedLoc = new ComboBox<>(FXCollections.observableArrayList(optFeedLoc));
        rRow.cmbFeedLoc.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblFeed, 0, 7);
        addFormField(rRow.grid, rRow.cmbFeedLoc, 1, 7);

        Label lblStorage = new Label("在庫場所:");
        styleFormLabel(lblStorage);
        rRow.cmbStorageLoc = new ComboBox<>(FXCollections.observableArrayList(optStorageLoc));
        rRow.cmbStorageLoc.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblStorage, 2, 7);
        addFormField(rRow.grid, rRow.cmbStorageLoc, 3, 7);

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
        rRow.grid.add(btnDelete, 3, 8);
        GridPane.setHalignment(btnDelete, javafx.geometry.HPos.RIGHT);

        Runnable updateSpec = () -> updateRawRowSpecDisplay(rRow);
        rRow.txtPart.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            updateSpec.run();
        });
        rRow.txtItem.textProperty().addListener((obs, oldV, newV) -> updateRowRawCandidates(rRow, false));
        rRow.txtType.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            updateSpec.run();
        });
        rRow.txtWidth.textProperty().addListener((obs, oldV, newV) -> updateSpec.run());
        rRow.txtLength.textProperty().addListener((obs, oldV, newV) -> {
            updateRowRawCandidates(rRow, false);
            updateSpec.run();
        });
        
        rRow.cmbSearch.setOnAction(evt -> {
            String sel = rRow.cmbSearch.getValue();
            if (sel != null && sel.contains(" | ")) {
                String code = sel.split(" \\| ")[0].trim();
                for (ProductInfo p : masterProductList) {
                    if (p.getShohinCode().equals(code)) {
                        isLoadingRecord = true;
                        try {
                            rRow.txtItem.setText(p.getShohinCode());
                            rRow.txtPart.setText(p.getFoamPartNo());
                            String[] nameParts = p.getShohinName1().split("-");
                            if (nameParts.length >= 2) rRow.txtType.setText(nameParts[1]);
                            rRow.txtWidth.setText(p.getFoamWidth());
                            rRow.txtLength.setText(p.getFoamLength());
                            updateRawRowSpecDisplay(rRow);
                        } finally {
                            isLoadingRecord = false;
                        }
                        break;
                    }
                }
            }
        });

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
        return rRow;
    }
}
