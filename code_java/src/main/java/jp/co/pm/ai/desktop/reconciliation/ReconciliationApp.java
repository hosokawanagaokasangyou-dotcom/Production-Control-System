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
    /** 受注ﾌｧｲﾙ: 先頭データ行（0-based）。 */
    private static final int JUCHU_SHEET_FIRST_DATA_ROW_INDEX = 3;
    private static final Path SETTINGS_FILE =
            Path.of(System.getProperty("user.home"), ".pm-ai-desktop", "request-form-reconciliation.properties");
    private static final double SETTINGS_CARD_WIDTH = 300.0;
    private static final long ORIGINAL_FILE_POLL_INTERVAL_MS = 5L * 60L * 1000L;
    /** {@link #saveLocalForm()} 後・受注ファイル未転記のステータス。 */
    private static final String STATUS_LOCAL_SAVE_PENDING = "手修正済み (未保存)";

    private Window hostWindow;
    private Path repoRootHint;
    private Path aladdinMasterDir;
    private Map<String, String> uiEnvSnapshot = Map.of();
    private Consumer<String> originalDirChangeHandler;
    private TextField txtJuchuPathDisplay;
    private Button btnTransfer;
    private Button btnBulkTransferPending;
    private String targetFolder = "";
    private String juchuFilePath;
    private boolean isLoadingRecord = false;

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

    private TextField newTxtFormKakoNaiyo;
    private TextField newTxtFormTokki1;
    private TextField newTxtFormTokki2;
    private TextField newTxtFormTokki3;
    private DatePicker newDpFormDeliv;
    private DatePicker newDpFormAdjustDeliv;
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
        leftScroll.setStyle("-fx-background-color: transparent; -fx-border-color: transparent;");
        leftScroll.setMinWidth(380);
        leftScroll.setPrefWidth(480);
        leftScroll.setMaxWidth(580);
        
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
        txtSearch.setPrefWidth(160);
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
        newTxtFormContractNo.setTextFormatter(new TextFormatter<>(change -> {
            String text = change.getText();
            if (text.matches("^[\\x20-\\x7E]*$")) {
                return change;
            }
            return null;
        }));
        basicGrid.add(lblContract, 2, 2);
        addFormField(basicGrid, newTxtFormContractNo, 3, 2);
        
        // Row 3
        Label lblProcess = new Label("加工内容:");
        styleFormLabel(lblProcess);
        txtProcess = new TextField();
        txtProcess.setStyle("-fx-font-size: 11px;");
        basicGrid.add(lblProcess, 0, 3);
        addFormField(basicGrid, txtProcess, 1, 3, 3, 1);
        
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
        
        Label lblKakoNaiyo = new Label("加工内容(詳細):");
        lblKakoNaiyo.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newTxtFormKakoNaiyo = new TextField();
        newTxtFormKakoNaiyo.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblKakoNaiyo, 0, 0);
        workGrid.add(newTxtFormKakoNaiyo, 1, 0, 3, 1);
        
        Label lblTokki1 = new Label("特記事項1:");
        lblTokki1.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newTxtFormTokki1 = new TextField();
        newTxtFormTokki1.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblTokki1, 0, 1);
        workGrid.add(newTxtFormTokki1, 1, 1, 3, 1);
        
        Label lblTokki2 = new Label("特記事項2:");
        lblTokki2.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newTxtFormTokki2 = new TextField();
        newTxtFormTokki2.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblTokki2, 0, 2);
        workGrid.add(newTxtFormTokki2, 1, 2, 3, 1);
        
        Label lblTokki3 = new Label("特記事項3:");
        lblTokki3.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newTxtFormTokki3 = new TextField();
        newTxtFormTokki3.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblTokki3, 0, 3);
        workGrid.add(newTxtFormTokki3, 1, 3, 3, 1);
        
        Label lblYoto = new Label("用途:");
        lblYoto.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newCmbFormYoto = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optYoto));
        newCmbFormYoto.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblYoto, 0, 4);
        workGrid.add(newCmbFormYoto, 1, 4);
        
        Label lblInputKbn = new Label("入力区分:");
        lblInputKbn.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newCmbFormInputKbn = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optInputKbn));
        newCmbFormInputKbn.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblInputKbn, 2, 4);
        workGrid.add(newCmbFormInputKbn, 3, 4);

        Label lblKakoKbn = new Label("加工区分:");
        lblKakoKbn.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newCmbFormKakoKbn = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optKakoKbn));
        newCmbFormKakoKbn.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblKakoKbn, 0, 5);
        workGrid.add(newCmbFormKakoKbn, 1, 5);
        
        Label lblInputTanto = new Label("入力担当:");
        lblInputTanto.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        newCmbFormInputTanto = new ComboBox<>(javafx.collections.FXCollections.observableArrayList(optInputTanto));
        newCmbFormInputTanto.setStyle("-fx-font-size: 11px;");
        workGrid.add(lblInputTanto, 2, 5);
        workGrid.add(newCmbFormInputTanto, 3, 5);
        
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
        
        btnContainer.getChildren().addAll(btnNewRecord, sideBtns);
        leftContainer.getChildren().add(btnContainer);
        
        // 2. RIGHT PANE: Visual Sheet Viewer & Discrepancies
        VBox rightPane = new VBox(10);
        rightPane.getStyleClass().add("pane-container");
        
        HBox viewerHeaderBox = new HBox(15);
        viewerHeaderBox.setAlignment(Pos.CENTER_LEFT);
        Label lblViewer = new Label("原本加工依頼書プレビュー (画像キャッシュ)");
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
        viewerHeaderBox.getChildren().addAll(lblViewer, previewFileNameRow, btnOpenExcel);
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
        splitPane.setDividerPositions(0.30); // Left takes 30%, Right takes 70%

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
                juchuFilePath = chosenFile.getAbsolutePath();
                txtJuchuPathDisplay.setText(juchuFilePath);
                File parentDir = chosenFile.getParentFile();
                if (parentDir != null) {
                    applySelectedOriginalDir(parentDir.getAbsolutePath());
                }
                saveSettings();
                statusLabel.setText(
                        "受注ファイルと依頼書原本フォルダを更新しました: "
                                + juchuFilePath);
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
        
        newTxtFormKakoNaiyo.setText("スライス");
        newTxtFormTokki1.setText("");
        newTxtFormTokki2.setText("");
        newTxtFormTokki3.setText("");
        if (!optYoto.isEmpty()) newCmbFormYoto.setValue(optYoto.get(0));
        if (!optUser.isEmpty()) newCmbFormUser.setValue(optUser.get(0));
        newDpFormDeliv.setValue(null);
        newDpFormAdjustDeliv.setValue(null);
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

    private void updateTransferButtonState() {
        boolean blocked = false;
        String blockedTooltip = null;
        if (juchuFilePath == null || juchuFilePath.isBlank()) {
            blocked = true;
            blockedTooltip = "受注ファイルが未設定のため転記できません。";
        } else {
            File juchuFile = new File(juchuFilePath);
            if (!juchuFile.isFile()) {
                blocked = true;
                blockedTooltip = "受注ファイルが見つからないため転記できません。";
            } else {
                File lockFile = excelLockFileFor(juchuFile);
                if (lockFile != null && lockFile.isFile()) {
                    blocked = true;
                    blockedTooltip =
                            "受注ファイルが Excel で使用中（読み取り専用）のため転記できません。"
                                    + " ロックファイル: "
                                    + lockFile.getName();
                }
            }
        }
        if (btnTransfer != null) {
            btnTransfer.setDisable(blocked);
            btnTransfer.setTooltip(blocked ? new Tooltip(blockedTooltip) : null);
        }
        if (btnBulkTransferPending != null) {
            btnBulkTransferPending.setDisable(blocked);
            btnBulkTransferPending.setTooltip(
                    blocked
                            ? new Tooltip(blockedTooltip)
                            : new Tooltip("一時保存済みで未転記の全レコードを受注ファイルへ書き込みます。"));
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
            if (!transferAllPendingLocalSaves(false)) {
                return;
            }
        }
        reloadData();
        if (afterReload != null) {
            afterReload.run();
        }
    }

    /** @param showSuccessDialog {@code true} のとき完了ダイアログを表示 */
    private boolean transferAllPendingLocalSaves(boolean showSuccessDialog) {
        List<OrderRecord> pending = collectPendingLocalSaveRecords();
        if (pending.isEmpty()) {
            if (showSuccessDialog) {
                showAlert("情報", "一時保存済みで未転記のデータはありません。");
            }
            return true;
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
            return false;
        }

        statusLabel.setText("一時保存分を受注ファイルへ一括転記中...");
        List<String> errors = new ArrayList<>();
        int transferred = 0;
        try {
            File file = new File(juchuFilePath);
            if (!file.isFile()) {
                showAlert("エラー", "受注ファイルが見つかりません。");
                return false;
            }

            try (FileInputStream fis = new FileInputStream(file);
                    Workbook wb = WorkbookFactory.create(fis)) {
                Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
                Map<String, Integer> colMap = buildJuchuColumnMap(sheet);
                for (OrderRecord record : pending) {
                    try {
                        transferRecordFromDbValues(wb, sheet, colMap, record);
                        record.setStatus("一致 (転記完了)");
                        record.setDiscrepancy("一致 (受注ファイルへ転記完了)");
                        transferred++;
                    } catch (Exception ex) {
                        errors.add(record.getReqNo() + ": " + ex.getMessage());
                    }
                }
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    wb.write(fos);
                }
            }

            refreshComboRecordItems();
            if (selectedRecord != null && isPendingLocalSave(selectedRecord)) {
                loadRecordDetails(selectedRecord);
            } else if (selectedRecord != null) {
                discrepancyLabel.setText(selectedRecord.getDiscrepancy());
            }
            updateTransferButtonState();

            if (!errors.isEmpty()) {
                statusLabel.setText(
                        "一括転記: 成功 " + transferred + " 件 / 失敗 " + errors.size() + " 件");
                showAlert(
                        "一括転記（一部失敗）",
                        "成功: "
                                + transferred
                                + " 件\n失敗: "
                                + errors.size()
                                + " 件\n\n"
                                + String.join("\n", errors));
                return false;
            }

            statusLabel.setText("一時保存 " + transferred + " 件の一括転記が完了しました。");
            if (showSuccessDialog) {
                showAlert("成功", "一時保存 " + transferred + " 件を受注ファイルへ転記しました。");
            }
            return true;
        } catch (Exception ex) {
            statusLabel.setText("エラー: " + ex.getMessage());
            showAlert("エラー", "一括転記中にエラーが発生しました:\n" + ex.getMessage());
            ex.printStackTrace();
            updateTransferButtonState();
            return false;
        }
    }

    private void transferAllPendingLocalSaves() {
        transferAllPendingLocalSaves(true);
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
        Row targetRow = null;
        int targetRowIndex = -1;
        for (int r = JUCHU_SHEET_FIRST_DATA_ROW_INDEX; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            Cell c = row.getCell(0);
            if (c != null && normalize_key(getCellValueAsString(c)).equals(normKey)) {
                targetRow = row;
                targetRowIndex = r + 1;
                break;
            }
        }

        boolean isNewRow = targetRow == null;
        if (isNewRow) {
            targetRowIndex = insertNewJuchuDataRowPreservingFormulas(sheet);
            targetRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
        }

        Cell cellReqNo = targetRow.createCell(0);
        cellReqNo.setCellValue(reqNo);

        CellStyle yellowStyle = wb.createCellStyle();
        Row refRow = sheet.getRow(4);
        if (refRow == null) {
            refRow = sheet.getRow(3);
        }
        if (refRow != null && refRow.getCell(0) != null) {
            yellowStyle.cloneStyleFrom(refRow.getCell(0).getCellStyle());
        }
        yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellReqNo.setCellStyle(yellowStyle);

        if (isNewRow) {
            setJuchuSheetString(targetRow, colMap, "入力区分", "通常入力");
            setJuchuSheetString(targetRow, colMap, "加工区分", "後加工");
        }
        setJuchuSheetString(targetRow, colMap, "入力担当", "自動転記");
        setJuchuSheetDate(targetRow, colMap, "入力日", new Date());

        setJuchuSheetString(targetRow, colMap, "品名", db.getOrDefault("品名", ""));
        setJuchuSheetString(targetRow, colMap, "製品", db.getOrDefault("製品", ""));
        setJuchuSheetString(targetRow, colMap, "梱-等1", db.getOrDefault("梱-等1", ""));
        setJuchuSheetString(targetRow, colMap, "色1", db.getOrDefault("色1", ""));
        setJuchuSheetString(targetRow, colMap, "区分1", db.getOrDefault("区分1", ""));
        setJuchuSheetString(targetRow, colMap, "数量1", db.getOrDefault("数量1", ""));
        setJuchuSheetString(targetRow, colMap, "枝番", db.getOrDefault("枝番", ""));
        setJuchuSheetString(targetRow, colMap, "ＥＣ面", db.getOrDefault("ＥＣ面", ""));
        setJuchuSheetString(targetRow, colMap, "ﾄﾘﾐﾝｸﾞ", db.getOrDefault("ﾄﾘﾐﾝｸﾞ", ""));
        setJuchuSheetString(targetRow, colMap, "ユーザー", db.getOrDefault("ユーザー", ""));

        setJuchuSheetDateOrStringFromDb(targetRow, colMap, "希望納期", db.getOrDefault("希望納期", ""));
        setJuchuSheetString(targetRow, colMap, "調整納期", db.getOrDefault("調整納期", ""));

        setJuchuSheetString(targetRow, colMap, "原反", db.getOrDefault("原反", ""));
        setJuchuSheetString(targetRow, colMap, "品名1", db.getOrDefault("原反品名", ""));
        setJuchuSheetString(targetRow, colMap, "梱-等", db.getOrDefault("原反梱-等", ""));
        setJuchuSheetString(targetRow, colMap, "色", db.getOrDefault("原反色", ""));
        setJuchuSheetString(targetRow, colMap, "区分", db.getOrDefault("原反区分", ""));
        setJuchuSheetString(targetRow, colMap, "数量", db.getOrDefault("原反数量", ""));
        setJuchuSheetString(targetRow, colMap, "割数", db.getOrDefault("割数", ""));
        setJuchuSheetString(targetRow, colMap, "投入場所", db.getOrDefault("投入場所", ""));
        setJuchuSheetString(targetRow, colMap, "在庫場所", db.getOrDefault("在庫場所", ""));

        setJuchuSheetCellIfPresent(
                targetRow, colMap, COL_MASTER_BASE_SHOHIN_PRODUCT, db.getOrDefault(COL_MASTER_BASE_SHOHIN_PRODUCT, ""));
        setJuchuSheetCellIfPresent(
                targetRow, colMap, COL_MASTER_BASE_SHOHIN_RAW, db.getOrDefault(COL_MASTER_BASE_SHOHIN_RAW, ""));

        setJuchuSheetString(targetRow, colMap, "加工内容", db.getOrDefault("加工内容", ""));
        setJuchuSheetNumericOrString(targetRow, colMap, "加工賃", db.getOrDefault("加工賃", ""));
        setJuchuSheetString(targetRow, colMap, "契約Ｎｏ", db.getOrDefault("契約Ｎｏ", ""));
        setJuchuSheetString(targetRow, colMap, "用途", db.getOrDefault("用途", ""));

        if (isNewRow) {
            applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);
            Map<String, String> raw = record.getRawValues();
            if (raw != null && !raw.isEmpty()) {
                setJuchuSheetString(targetRow, colMap, "品名1", raw.getOrDefault("原反品名", ""));
                setJuchuSheetString(targetRow, colMap, "梱-等1", raw.getOrDefault("梱-等1", ""));
                setJuchuSheetString(targetRow, colMap, "色1", raw.getOrDefault("色1", ""));
                setJuchuSheetString(targetRow, colMap, "区分1", raw.getOrDefault("区分1", ""));
                setJuchuSheetString(targetRow, colMap, "梱-等", raw.getOrDefault("原反梱-等", ""));
                setJuchuSheetString(targetRow, colMap, "色", raw.getOrDefault("原反色", ""));
                setJuchuSheetString(targetRow, colMap, "区分", raw.getOrDefault("原反区分", ""));
                setJuchuSheetString(targetRow, colMap, "在庫場所", raw.getOrDefault("在庫場所", ""));
                setJuchuSheetNumericOrString(targetRow, colMap, "数量", raw.getOrDefault("原反数量", ""));
                setJuchuSheetNumericOrString(targetRow, colMap, "加工賃", raw.getOrDefault("加工賃", ""));
                setJuchuSheetNumericOrString(
                        targetRow, colMap, "原反ロール数", raw.getOrDefault("原反ロール数", ""));
            }
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
            for (int r = 3; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell c = row.getCell(0);
                if (c != null && normalize_key(getCellValueAsString(c)).equals(normKey)) {
                    showAlert("エラー", "この依頼Ｎｏ（" + reqNo + "）は既に存在します。別の番号を指定してください。");
                    wb.close();
                    fis.close();
                    return;
                }
            }
            
            int targetRowIndex = insertNewJuchuDataRowPreservingFormulas(sheet);
            Row targetRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
            
            String inputKbnVal = newCmbInputKbn.getValue() != null ? newCmbInputKbn.getValue().trim() : "通常入力";
            String kakoKbnVal  = newCmbKakoKbn.getValue()  != null ? newCmbKakoKbn.getValue().trim()  : "後加工";
            String inputTanto  = newCmbInputTanto.getValue() != null ? newCmbInputTanto.getValue().trim() : "";
            String wariSuVal   = newCmbWariSu.getValue() != null ? newCmbWariSu.getValue().trim() : "1";

            Cell cellReqNo = targetRow.createCell(0);
            cellReqNo.setCellValue(reqNo);
            
            // Set yellow background fill to mark as automatically transferred
            CellStyle yellowStyle = wb.createCellStyle();
            Row refRow = sheet.getRow(4);
            if (refRow == null) refRow = sheet.getRow(3);
            if (refRow != null && refRow.getCell(0) != null) {
                yellowStyle.cloneStyleFrom(refRow.getCell(0).getCellStyle());
            }
            yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellReqNo.setCellStyle(yellowStyle);
            setJuchuSheetString(targetRow, colMap, "入力区分", inputKbnVal);
            setJuchuSheetString(targetRow, colMap, "加工区分", kakoKbnVal);
            setJuchuSheetNumericOrString(targetRow, colMap, "割数", wariSuVal);
            setJuchuSheetString(targetRow, colMap, "入力担当", inputTanto);
            setJuchuSheetDate(targetRow, colMap, "入力日", new Date());
            setJuchuSheetString(targetRow, colMap, "品名", newTxtProdItem.getText().trim());
            String prodSpec = newTxtProdPart.getText().trim() + "-" + newTxtProdType.getText().trim() + "-" + newTxtProdWidth.getText().trim() + "X" + newTxtProdLength.getText().trim();
            setJuchuSheetString(targetRow, colMap, "製品", prodSpec);
            setJuchuSheetString(targetRow, colMap, "梱-等1", newTxtProdGrade.getText().trim());
            setJuchuSheetString(targetRow, colMap, "色1", newTxtProdColor.getText().trim());
            setJuchuSheetString(targetRow, colMap, "区分1", newTxtProdCategory.getText().trim());
            setJuchuSheetNumericOrString(targetRow, colMap, "枝番", newTxtProdEdaban.getText().trim());
            setJuchuSheetNumericOrString(targetRow, colMap, "数量1", newTxtProdQty.getText().trim());
            String ecSideVal  = newCmbProdEcSide.getValue()   != null ? newCmbProdEcSide.getValue().trim()  : "";
            String trimmingVal = newCmbProdTrimming.getValue() != null ? newCmbProdTrimming.getValue().trim() : "";
            setJuchuSheetString(targetRow, colMap, "ＥＣ面", ecSideVal);
            setJuchuSheetString(targetRow, colMap, "ﾄﾘﾐﾝｸﾞ", trimmingVal);
            setJuchuSheetString(targetRow, colMap, "品名1", newTxtRawItem.getText().trim());
            String rawSpec = newTxtRawPart.getText().trim() + "-" + newTxtRawType.getText().trim() + "-" + newTxtRawWidth.getText().trim() + "X" + newTxtRawLength.getText().trim();
            setJuchuSheetString(targetRow, colMap, "原反", rawSpec);
            setJuchuSheetString(targetRow, colMap, "梱-等", newTxtRawGrade.getText().trim());
            setJuchuSheetString(targetRow, colMap, "色", newTxtRawColor.getText().trim());
            setJuchuSheetString(targetRow, colMap, "区分", newTxtRawCategory.getText().trim());
            setJuchuSheetNumericOrString(targetRow, colMap, "数量", newTxtRawQty.getText().trim());
            String storageLoc = newCmbRawStorageLoc.getValue() != null ? newCmbRawStorageLoc.getValue().trim() : "";
            String feedLoc    = newCmbRawFeedLoc.getValue()    != null ? newCmbRawFeedLoc.getValue().trim()    : "";
            setJuchuSheetString(targetRow, colMap, "在庫場所", storageLoc);
            setJuchuSheetString(targetRow, colMap, "投入場所", feedLoc);
            setJuchuSheetString(targetRow, colMap, "加工内容", newTxtFormKakoNaiyo.getText().trim());
            setJuchuSheetString(targetRow, colMap, "特記事項1", newTxtFormTokki1.getText().trim());
            setJuchuSheetString(targetRow, colMap, "特記事項2", newTxtFormTokki2.getText().trim());
            setJuchuSheetString(targetRow, colMap, "特記事項3", newTxtFormTokki3.getText().trim());
            String yotoVal = newCmbFormYoto.getValue() != null ? newCmbFormYoto.getValue().trim() : "";
            String userVal = newCmbFormUser.getValue() != null ? newCmbFormUser.getValue().trim() : "";
            setJuchuSheetString(targetRow, colMap, "用途", yotoVal);
            setJuchuSheetString(targetRow, colMap, "ユーザー", userVal);
            if (newDpFormDeliv.getValue() != null) {
                try {
                    Calendar cVal = Calendar.getInstance();
                    cVal.set(newDpFormDeliv.getValue().getYear(), newDpFormDeliv.getValue().getMonthValue() - 1, newDpFormDeliv.getValue().getDayOfMonth());
                    setJuchuSheetDate(targetRow, colMap, "希望納期", cVal.getTime());
                } catch (Exception e) {
                    // skip
                }
            }
            setJuchuSheetString(
                    targetRow,
                    colMap,
                    "調整納期",
                    newDpFormAdjustDeliv.getValue() != null ? newDpFormAdjustDeliv.getValue().toString() : "");
            setJuchuSheetNumericOrString(targetRow, colMap, "加工賃", newTxtFormWage.getText().trim());
            setJuchuSheetString(targetRow, colMap, "契約Ｎｏ", newTxtFormContractNo.getText().trim());
            applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);
            
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
                
                Row hRow = sJuchu.getRow(2);
                Map<String, Integer> colMap = new HashMap<>();
                for (int c = 0; c < hRow.getLastCellNum(); c++) {
                    Cell cell = hRow.getCell(c);
                    if (cell != null) {
                        colMap.put(cell.getStringCellValue().trim(), c);
                    }
                }
                
                Map<String, Map<String, String>> dbRows = new HashMap<>();
                
                for (int r = 3; r <= sJuchu.getLastRowNum(); r++) {
                    Row row = sJuchu.getRow(r);
                    if (row == null) continue;
                    Cell reqCell = row.getCell(0); 
                    if (reqCell == null || reqCell.getCellType() == CellType.BLANK) continue;
                    
                    String reqNo = getCellValueAsString(reqCell).trim();
                    if (reqNo.isEmpty()) continue;
                    
                    Map<String, String> vals = new HashMap<>();
                    for (Map.Entry<String, Integer> entry : colMap.entrySet()) {
                        Cell c = row.getCell(entry.getValue());
                        vals.put(entry.getKey(), getCellValueAsString(c));
                    }
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
                        if (!normalize_text(raw.get("契約Ｎｏ")).equals(normalize_text(dbRow.get("契約Ｎｏ")))) {
                            if (!raw.get("契約Ｎｏ").isEmpty() || !dbRow.get("契約Ｎｏ").isEmpty()) {
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
                            reqNo, "新規自動追加 (未登録)", raw.get("ユーザー"), raw.get("製品"), "受注ファイル未入力のため自動追加", raw, new HashMap<>()
                        ));
                    }
                }
                
                for (Map.Entry<String, Map<String, String>> entry : dbRows.entrySet()) {
                    if (!processedKeys.contains(entry.getKey())) {
                        Map<String, String> dbRow = entry.getValue();
                        loadedRecords.add(new OrderRecord(
                            dbRow.get("依頼Ｎｏ"), "既存登録 (原本未確認)", dbRow.get("ユーザー"), dbRow.get("製品"), "原本ファイル未検出（過去データ）", new HashMap<>(), dbRow
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
            Platform.runLater(() -> {
                showLoadingOverlay(false, "");
                if (finalSuccess) {
                    orderRecords.clear();
                    orderRecords.addAll(finalLoaded);
                    filterTable("");  // ファイルが存在するレコードのみ表示
                    int visibleCount = (int) orderRecords.stream().filter(r -> hasExistingFile(r)).count();
                    statusLabel.setText(String.format("読込完了: 全 %d 件中 %d 件 (依頼書あり)", orderRecords.size(), visibleCount));
                    
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
                    parsed.add(buildRawMapFromSheet(file, sName, rawSheet));
                }
            }
        }
        return parsed;
    }

    private Map<String, String> buildRawMapFromSheet(File file, String sName, Sheet rawSheet) {
        String reqNo = getCellValueAsString(rawSheet.getRow(4).getCell(17));
        if (reqNo.isEmpty()) {
            reqNo = sName;
        }

        String itemCode = getCellValueAsString(rawSheet.getRow(9).getCell(1));
        String partNo = getCellValueAsString(rawSheet.getRow(9).getCell(5));
        String itemType = getCellValueAsString(rawSheet.getRow(9).getCell(10));
        String itemWidth = getCellValueAsString(rawSheet.getRow(9).getCell(15));
        String itemLen = getCellValueAsString(rawSheet.getRow(9).getCell(20));
        String grade = getCellValueAsString(rawSheet.getRow(9).getCell(23));
        String color = getCellValueAsString(rawSheet.getRow(9).getCell(26));
        String category = getCellValueAsString(rawSheet.getRow(9).getCell(28));

        String qty = getCellValueAsString(rawSheet.getRow(9).getCell(30));
        String ecSide = getCellValueAsString(rawSheet.getRow(9).getCell(35));
        String trimming = getCellValueAsString(rawSheet.getRow(9).getCell(38));

        String rawItem = getCellValueAsString(rawSheet.getRow(22).getCell(7));
        String rawPart = getCellValueAsString(rawSheet.getRow(22).getCell(10));
        String rawType = getCellValueAsString(rawSheet.getRow(22).getCell(13));
        String rawWidth = getCellValueAsString(rawSheet.getRow(22).getCell(16));
        String rawLen = getCellValueAsString(rawSheet.getRow(22).getCell(19));
        String rawGrade = getCellValueAsString(rawSheet.getRow(22).getCell(21));
        String rawColor = getCellValueAsString(rawSheet.getRow(22).getCell(24));
        String rawCategory = getCellValueAsString(rawSheet.getRow(22).getCell(26));

        String constructedProd = partNo + "-" + itemType + "-" + itemWidth + "X" + itemLen;
        String constructedRaw = rawPart + "-" + rawType + "-" + rawWidth + "X" + rawLen;

        List<String> steps = new ArrayList<>();
        for (int r = 12; r <= 16; r++) {
            Row rRow = rawSheet.getRow(r);
            if (rRow != null) {
                Cell c = rRow.getCell(8);
                if (c != null && c.getCellType() != CellType.BLANK) {
                    steps.add(getCellValueAsString(c).trim());
                }
            }
        }
        String processStr = String.join(", ", steps);

        String useVal = getCellValueAsString(rawSheet.getRow(17).getCell(4));
        String userVal = getCellValueAsString(rawSheet.getRow(18).getCell(4));
        String delivery = getCellValueAsString(rawSheet.getRow(19).getCell(8));
        String contract = getCellValueAsString(rawSheet.getRow(20).getCell(4));
        String rawQty = getCellValueAsString(rawSheet.getRow(22).getCell(28));
        String location = getCellValueAsString(rawSheet.getRow(22).getCell(31));
        String inputDay = getCellValueAsString(rawSheet.getRow(22).getCell(38));

        String issuer = getCellValueAsString(rawSheet.getRow(25).getCell(2));
        String rawRoll = getCellValueAsString(rawSheet.getRow(19).getCell(26));
        String charge = getCellValueAsString(rawSheet.getRow(19).getCell(30));

        Map<String, String> rawMap = new HashMap<>();
        rawMap.put("依頼Ｎｏ", reqNo.trim());
        rawMap.put("原本ファイル名", file.getName());
        rawMap.put("原本シート名", sName);
        rawMap.put("品名", itemCode);
        rawMap.put("製品", constructedProd);
        rawMap.put("梱-等1", grade);
        rawMap.put("色1", color);
        rawMap.put("区分1", category);
        rawMap.put("数量1", qty);
        rawMap.put("ＥＣ面", ecSide);
        rawMap.put("ﾄﾘﾐﾝｸﾞ", trimming);
        rawMap.put("加工内容", processStr);
        rawMap.put("用途", useVal);
        rawMap.put("ユーザー", userVal);
        rawMap.put("希望納期", delivery);
        rawMap.put("契約Ｎｏ", contract);
        rawMap.put("原反品名", rawItem);
        rawMap.put("原反", constructedRaw);
        rawMap.put("原反梱-等", rawGrade);
        rawMap.put("原反色", rawColor);
        rawMap.put("原反区分", rawCategory);
        rawMap.put("原反数量", rawQty);
        rawMap.put("在庫場所", location);
        rawMap.put("投入日", inputDay);
        rawMap.put("発行者", issuer);
        rawMap.put("原反ロール数", rawRoll);
        rawMap.put("加工賃", charge);
        return rawMap;
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
                            File cachedImg =
                                    RequestFormSourceCache.pngCacheFile(cacheRoot, fName, sName);
                            if (!RequestFormSourceCache.isPreviewCacheValid(cachedImg, f)) {
                                RequestFormSourceCache.deletePreviewCache(cachedImg);
                                boolean exists = false;
                                for (CacheTask t : cacheQueue) {
                                    if (t.outputFile.equals(cachedImg)) {
                                        exists = true;
                                        break;
                                    }
                                }
                                if (!exists) {
                                    cacheQueue.add(new CacheTask(f, sName, cachedImg));
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
                    RequestFormSheetPreviewRenderer.generatePreviewPng(
                            task.excelFile, task.sheetName, task.outputFile);
                    RequestFormSourceCache.writePreviewMeta(task.outputFile, task.excelFile);
                    System.out.println("Background cached: " + task.outputFile.getName());
                } catch (Exception e) {
                    System.err.println("Background cache exception: " + e.getMessage());
                }

                synchronized (cacheQueue) {
                    if (!cacheQueue.isEmpty() && cacheQueue.peek() == task) {
                        cacheQueue.poll();
                    }
                }

                try {
                    Thread.sleep(200);
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
        
        Map<String, String> activeVals = record.getDbValues().isEmpty() ? record.getRawValues() : record.getDbValues();
        
        newCmbFormUser.setValue(activeVals.getOrDefault("ユーザー", ""));
        newDpFormDeliv.setValue(parseLocalDate(activeVals.getOrDefault("希望納期", "")));
        newDpFormAdjustDeliv.setValue(parseLocalDate(activeVals.getOrDefault("調整納期", "")));
        txtProcess.setText(activeVals.getOrDefault("加工内容", ""));
        newTxtFormWage.setText(activeVals.getOrDefault("加工賃", ""));
        newTxtFormContractNo.setText(activeVals.getOrDefault("契約Ｎｏ", ""));
        newCmbFormYoto.setValue(activeVals.getOrDefault("用途", ""));
        
        // Load dynamic product rows
        String[] itemVals = activeVals.getOrDefault("品名", "").split("\\n");
        String[] prodVals = activeVals.getOrDefault("製品", "").split("\\n");
        String[] qtyVals = activeVals.getOrDefault("数量1", "").split("\\n");
        String[] gradeVals = activeVals.getOrDefault("梱-等1", "").split("\\n");
        String[] colorVals = activeVals.getOrDefault("色1", "").split("\\n");
        String[] catVals = activeVals.getOrDefault("区分1", "").split("\\n");
        String[] edabanVals = activeVals.getOrDefault("枝番", "").split("\\n");
        String[] ecVals = activeVals.getOrDefault("ＥＣ面", "").split("\\n");
        String[] trimVals = activeVals.getOrDefault("ﾄﾘﾐﾝｸﾞ", "").split("\\n");
        
        int numProds = Math.max(1, prodVals.length);
        productRowsContainer.getChildren().clear();
        productRows.clear();
        
        for (int i = 0; i < numProds; i++) {
            ProductRow pRow = addProductRow(null);
            // Set item code empty as requested by user to allow clean search
            pRow.txtItem.setText("");
            if (i < qtyVals.length) pRow.txtQty.setText(qtyVals[i]);
            if (i < gradeVals.length) pRow.txtGrade.setText(gradeVals[i]);
            if (i < colorVals.length) pRow.txtColor.setText(colorVals[i]);
            if (i < catVals.length) pRow.txtCategory.setText(catVals[i]);
            if (i < edabanVals.length) pRow.txtEdaban.setText(edabanVals[i]);
            if (i < ecVals.length) pRow.cmbEcSide.setValue(ecVals[i]);
            if (i < trimVals.length) pRow.cmbTrimming.setValue(trimVals[i]);
            
            if (i < prodVals.length) {
                String pStr = prodVals[i];
                String[] parts = pStr.split("-");
                if (parts.length >= 3) {
                    pRow.txtPart.setText(parts[0]);
                    pRow.txtType.setText(parts[1]);
                    String dims = parts[2];
                    String[] wL = dims.split("X");
                    if (wL.length >= 2) {
                        pRow.txtWidth.setText(wL[0]);
                        pRow.txtLength.setText(wL[1]);
                    } else {
                        pRow.txtWidth.setText(dims);
                        pRow.txtLength.setText("");
                    }
                } else {
                    pRow.txtPart.setText(pStr);
                }
            }
        }
        
        // Load dynamic raw material rows
        String[] rawItemVals = activeVals.getOrDefault("原反品名", "").split("\\n");
        String[] rawVals = activeVals.getOrDefault("原反", "").split("\\n");
        String[] rawQtyVals = activeVals.getOrDefault("原反数量", "").split("\\n");
        String[] rawGradeVals = activeVals.getOrDefault("原反梱-等", "").split("\\n");
        String[] rawColorVals = activeVals.getOrDefault("原反色", "").split("\\n");
        String[] rawCatVals = activeVals.getOrDefault("原反区分", "").split("\\n");
        String[] wariVals = activeVals.getOrDefault("割数", "").split("\\n");
        String[] feedVals = activeVals.getOrDefault("投入場所", "").split("\\n");
        String[] storageVals = activeVals.getOrDefault("在庫場所", "").split("\\n");
        
        int numRaws = Math.max(1, rawVals.length);
        rawRowsContainer.getChildren().clear();
        rawRows.clear();
        
        for (int i = 0; i < numRaws; i++) {
            RawMaterialRow rRow = addRawRow(null);
            // Set raw item empty to allow clean search
            rRow.txtItem.setText("");
            if (i < rawQtyVals.length) rRow.txtQty.setText(rawQtyVals[i]);
            if (i < rawGradeVals.length) rRow.txtGrade.setText(rawGradeVals[i]);
            if (i < rawColorVals.length) rRow.txtColor.setText(rawColorVals[i]);
            if (i < rawCatVals.length) rRow.txtCategory.setText(rawCatVals[i]);
            if (i < wariVals.length) rRow.cmbWariSu.setValue(wariVals[i]);
            if (i < feedVals.length) rRow.cmbFeedLoc.setValue(feedVals[i]);
            if (i < storageVals.length) rRow.cmbStorageLoc.setValue(storageVals[i]);
            
            if (i < rawVals.length) {
                String rStr = rawVals[i];
                String[] parts = rStr.split("-");
                if (parts.length >= 3) {
                    rRow.txtPart.setText(parts[0]);
                    rRow.txtType.setText(parts[1]);
                    String dims = parts[2];
                    String[] wL = dims.split("X");
                    if (wL.length >= 2) {
                        rRow.txtWidth.setText(wL[0]);
                        rRow.txtLength.setText(wL[1]);
                    } else {
                        rRow.txtWidth.setText(dims);
                        rRow.txtLength.setText("");
                    }
                } else {
                    rRow.txtPart.setText(rStr);
                }
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
        File cachedImg = RequestFormSourceCache.pngCacheFile(cacheRoot, fName, sName);

        if (RequestFormSourceCache.isPreviewCacheValid(cachedImg, f)) {
            acknowledgePreviewForCurrentOriginalFile();
            displayPreviewImage(cachedImg);
        } else {
            RequestFormSourceCache.deletePreviewCache(cachedImg);
            // Generate async
            Label lblLoading = new Label("原本の画像を抽出しています... (しばらくお待ちください)");
            lblLoading.getStyleClass().add("excel-grid-label-info");
            sheetGrid.add(lblLoading, 0, 0);
            
            Thread t = new Thread(() -> {
                try {
                    RequestFormSheetPreviewRenderer.generatePreviewPng(f, sName, cachedImg);
                    RequestFormSourceCache.writePreviewMeta(cachedImg, f);
                    Platform.runLater(() -> {
                        if (RequestFormSourceCache.isPreviewCacheValid(cachedImg, f)) {
                            acknowledgePreviewForCurrentOriginalFile();
                            displayPreviewImage(cachedImg);
                        } else {
                            sheetGrid.getChildren().clear();
                            Label lblErr = new Label("画像の生成に失敗しました。");
                            lblErr.getStyleClass().add("excel-grid-label-error");
                            sheetGrid.add(lblErr, 0, 0);
                        }
                    });
                } catch (Exception e) {
                    Platform.runLater(() -> {
                        sheetGrid.getChildren().clear();
                        Label lblErr = new Label("エラーが発生しました: " + e.getMessage());
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
    
    private void displayPreviewImage(File imgFile) {
        sheetGrid.getChildren().clear();
        try {
            javafx.scene.image.Image image =
                    new javafx.scene.image.Image(imgFile.toURI().toString(), false);
            if (image.isError()) {
                Throwable ex = image.getException();
                throw new IllegalStateException(
                        ex != null ? ex.getMessage() : "画像の読み込みに失敗しました");
            }
            if (image.getWidth() <= 0 || image.getHeight() <= 0) {
                throw new IllegalStateException("画像サイズが 0 です");
            }
            InteractiveImageViewer viewer = new InteractiveImageViewer(image);

            viewer.prefWidthProperty().bind(sheetScrollPane.widthProperty().subtract(24));
            viewer.maxWidthProperty().bind(sheetScrollPane.widthProperty().subtract(24));
            
            sheetGrid.add(viewer, 0, 0);
            viewer.resetZoomToFit();
        } catch (Exception e) {
            Label lblErr = new Label("画像の読み込みに失敗しました: " + e.getMessage());
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
        
        // Collect basicGrid fields
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
        db.put("契約Ｎｏ", newTxtFormContractNo.getText().trim());
        db.put("用途", newCmbFormYoto.getValue() != null ? newCmbFormYoto.getValue().trim() : "");
        
        // Collect dynamic product rows
        List<String> prodItems = new ArrayList<>();
        List<String> prodSpecs = new ArrayList<>();
        List<String> prodGrades = new ArrayList<>();
        List<String> prodColors = new ArrayList<>();
        List<String> prodCats = new ArrayList<>();
        List<String> prodQtys = new ArrayList<>();
        List<String> prodEdabans = new ArrayList<>();
        List<String> prodEcs = new ArrayList<>();
        List<String> prodTrims = new ArrayList<>();
        
        for (ProductRow pRow : productRows) {
            prodItems.add(pRow.txtItem.getText().trim());
            String spec = pRow.txtPart.getText().trim() + "-" + pRow.txtType.getText().trim() + "-" + pRow.txtWidth.getText().trim() + "X" + pRow.txtLength.getText().trim();
            prodSpecs.add(spec);
            prodGrades.add(pRow.txtGrade.getText().trim());
            prodColors.add(pRow.txtColor.getText().trim());
            prodCats.add(pRow.txtCategory.getText().trim());
            prodQtys.add(pRow.txtQty.getText().trim());
            prodEdabans.add(pRow.txtEdaban.getText().trim());
            prodEcs.add(pRow.cmbEcSide.getValue() != null ? pRow.cmbEcSide.getValue() : "");
            prodTrims.add(pRow.cmbTrimming.getValue() != null ? pRow.cmbTrimming.getValue() : "");
        }
        
        db.put("品名", String.join("\n", prodItems));
        db.put("製品", String.join("\n", prodSpecs));
        db.put("梱-等1", String.join("\n", prodGrades));
        db.put("色1", String.join("\n", prodColors));
        db.put("区分1", String.join("\n", prodCats));
        db.put("数量1", String.join("\n", prodQtys));
        db.put("枝番", String.join("\n", prodEdabans));
        db.put("ＥＣ面", String.join("\n", prodEcs));
        db.put("ﾄﾘﾐﾝｸﾞ", String.join("\n", prodTrims));
        db.put(COL_MASTER_BASE_SHOHIN_PRODUCT, joinMasterBaseShohinFromProductRows());
        
        // Collect dynamic raw material rows
        List<String> rawItems = new ArrayList<>();
        List<String> rawSpecs = new ArrayList<>();
        List<String> rawGrades = new ArrayList<>();
        List<String> rawColors = new ArrayList<>();
        List<String> rawCats = new ArrayList<>();
        List<String> rawQtys = new ArrayList<>();
        List<String> wariVals = new ArrayList<>();
        List<String> feedVals = new ArrayList<>();
        List<String> storageVals = new ArrayList<>();
        
        for (RawMaterialRow rRow : rawRows) {
            rawItems.add(rRow.txtItem.getText().trim());
            String spec = rRow.txtPart.getText().trim() + "-" + rRow.txtType.getText().trim() + "-" + rRow.txtWidth.getText().trim() + "X" + rRow.txtLength.getText().trim();
            rawSpecs.add(spec);
            rawGrades.add(rRow.txtGrade.getText().trim());
            rawColors.add(rRow.txtColor.getText().trim());
            rawCats.add(rRow.txtCategory.getText().trim());
            rawQtys.add(rRow.txtQty.getText().trim());
            wariVals.add(rRow.cmbWariSu.getValue() != null ? rRow.cmbWariSu.getValue() : "1");
            feedVals.add(rRow.cmbFeedLoc.getValue() != null ? rRow.cmbFeedLoc.getValue() : "");
            storageVals.add(rRow.cmbStorageLoc.getValue() != null ? rRow.cmbStorageLoc.getValue() : "");
        }
        
        db.put("原反品名", String.join("\n", rawItems));
        db.put("原反", String.join("\n", rawSpecs));
        db.put("原反梱-等", String.join("\n", rawGrades));
        db.put("原反色", String.join("\n", rawColors));
        db.put("原反区分", String.join("\n", rawCats));
        db.put("原反数量", String.join("\n", rawQtys));
        db.put("割数", String.join("\n", wariVals));
        db.put("投入場所", String.join("\n", feedVals));
        db.put("在庫場所", String.join("\n", storageVals));
        db.put(COL_MASTER_BASE_SHOHIN_RAW, joinMasterBaseShohinFromRawRows());
        
        selectedRecord.setStatus(STATUS_LOCAL_SAVE_PENDING);
        
        // ComboBox表示のリフレッシュ
        comboRecord.setItems(null);
        comboRecord.setItems(orderRecords);
        comboRecord.getSelectionModel().select(selectedRecord);
        
        statusLabel.setText("修正をローカル保存しました（受注ファイルに反映させるには転記ボタンを押してください）");
    }

    private void transferToExcel() {
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
        
        statusLabel.setText("受注ファイル更新中...");
        
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
            Row targetRow = null;
            int targetRowIndex = -1;
            
            for (int r = 3; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell c = row.getCell(0);
                if (c != null && normalize_key(getCellValueAsString(c)).equals(normKey)) {
                    targetRow = row;
                    targetRowIndex = r + 1; 
                    break;
                }
            }
            
            boolean isNewRow = (targetRow == null);
            
            if (isNewRow) {
                targetRowIndex = insertNewJuchuDataRowPreservingFormulas(sheet);
                targetRow = sheet.getRow(JUCHU_SHEET_FIRST_DATA_ROW_INDEX);
            }
            
            Cell cellReqNo = targetRow.createCell(0);
            cellReqNo.setCellValue(reqNo);
            
            // Set yellow background fill to mark as automatically transferred
            CellStyle yellowStyle = wb.createCellStyle();
            Row refRow = sheet.getRow(4);
            if (refRow == null) refRow = sheet.getRow(3);
            if (refRow != null && refRow.getCell(0) != null) {
                yellowStyle.cloneStyleFrom(refRow.getCell(0).getCellStyle());
            }
            yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellReqNo.setCellStyle(yellowStyle);
            
            if (isNewRow) {
                setJuchuSheetString(targetRow, colMap, "入力区分", "通常入力");
                setJuchuSheetString(targetRow, colMap, "加工区分", "後加工");
            }
            setJuchuSheetString(targetRow, colMap, "入力担当", "自動転記");
            setJuchuSheetDate(targetRow, colMap, "入力日", new Date());
            
            // Gather multiple product and raw material lines
            List<String> prodItems = new ArrayList<>();
            List<String> prodSpecs = new ArrayList<>();
            List<String> prodGrades = new ArrayList<>();
            List<String> prodColors = new ArrayList<>();
            List<String> prodCats = new ArrayList<>();
            List<String> prodQtys = new ArrayList<>();
            List<String> prodEdabans = new ArrayList<>();
            List<String> prodEcs = new ArrayList<>();
            List<String> prodTrims = new ArrayList<>();
            
            for (ProductRow pRow : productRows) {
                prodItems.add(pRow.txtItem.getText().trim());
                prodSpecs.add(pRow.txtPart.getText().trim() + "-" + pRow.txtType.getText().trim() + "-" + pRow.txtWidth.getText().trim() + "X" + pRow.txtLength.getText().trim());
                prodGrades.add(pRow.txtGrade.getText().trim());
                prodColors.add(pRow.txtColor.getText().trim());
                prodCats.add(pRow.txtCategory.getText().trim());
                prodQtys.add(pRow.txtQty.getText().trim());
                prodEdabans.add(pRow.txtEdaban.getText().trim());
                prodEcs.add(pRow.cmbEcSide.getValue() != null ? pRow.cmbEcSide.getValue() : "");
                prodTrims.add(pRow.cmbTrimming.getValue() != null ? pRow.cmbTrimming.getValue() : "");
            }
            
            List<String> rawItems = new ArrayList<>();
            List<String> rawSpecs = new ArrayList<>();
            List<String> rawGrades = new ArrayList<>();
            List<String> rawColors = new ArrayList<>();
            List<String> rawCats = new ArrayList<>();
            List<String> rawQtys = new ArrayList<>();
            List<String> wariVals = new ArrayList<>();
            List<String> feedVals = new ArrayList<>();
            List<String> storageVals = new ArrayList<>();
            
            for (RawMaterialRow rRow : rawRows) {
                rawItems.add(rRow.txtItem.getText().trim());
                rawSpecs.add(rRow.txtPart.getText().trim() + "-" + rRow.txtType.getText().trim() + "-" + rRow.txtWidth.getText().trim() + "X" + rRow.txtLength.getText().trim());
                rawGrades.add(rRow.txtGrade.getText().trim());
                rawColors.add(rRow.txtColor.getText().trim());
                rawCats.add(rRow.txtCategory.getText().trim());
                rawQtys.add(rRow.txtQty.getText().trim());
                wariVals.add(rRow.cmbWariSu.getValue() != null ? rRow.cmbWariSu.getValue() : "1");
                feedVals.add(rRow.cmbFeedLoc.getValue() != null ? rRow.cmbFeedLoc.getValue() : "");
                storageVals.add(rRow.cmbStorageLoc.getValue() != null ? rRow.cmbStorageLoc.getValue() : "");
            }
            
            setJuchuSheetString(targetRow, colMap, "品名", String.join("\n", prodItems));
            setJuchuSheetString(targetRow, colMap, "製品", String.join("\n", prodSpecs));
            setJuchuSheetString(targetRow, colMap, "梱-等1", String.join("\n", prodGrades));
            setJuchuSheetString(targetRow, colMap, "色1", String.join("\n", prodColors));
            setJuchuSheetString(targetRow, colMap, "区分1", String.join("\n", prodCats));
            setJuchuSheetString(targetRow, colMap, "数量1", String.join("\n", prodQtys));
            setJuchuSheetString(targetRow, colMap, "枝番", String.join("\n", prodEdabans));
            setJuchuSheetString(targetRow, colMap, "ＥＣ面", String.join("\n", prodEcs));
            setJuchuSheetString(targetRow, colMap, "ﾄﾘﾐﾝｸﾞ", String.join("\n", prodTrims));
            setJuchuSheetString(
                    targetRow,
                    colMap,
                    "ユーザー",
                    newCmbFormUser.getValue() != null ? newCmbFormUser.getValue().trim() : "");
            if (newDpFormDeliv.getValue() != null) {
                try {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    setJuchuSheetDate(targetRow, colMap, "希望納期", sdf.parse(newDpFormDeliv.getValue().toString()));
                } catch (Exception ex) {
                    setJuchuSheetString(targetRow, colMap, "希望納期", newDpFormDeliv.getValue().toString());
                }
            } else {
                setJuchuSheetString(targetRow, colMap, "希望納期", "");
            }
            setJuchuSheetString(
                    targetRow,
                    colMap,
                    "調整納期",
                    newDpFormAdjustDeliv.getValue() != null ? newDpFormAdjustDeliv.getValue().toString() : "");
            setJuchuSheetString(targetRow, colMap, "原反", String.join("\n", rawSpecs));
            setJuchuSheetString(targetRow, colMap, "品名1", String.join("\n", rawItems));
            setJuchuSheetString(targetRow, colMap, "梱-等", String.join("\n", rawGrades));
            setJuchuSheetString(targetRow, colMap, "色", String.join("\n", rawColors));
            setJuchuSheetString(targetRow, colMap, "区分", String.join("\n", rawCats));
            setJuchuSheetString(targetRow, colMap, "数量", String.join("\n", rawQtys));
            setJuchuSheetString(targetRow, colMap, "割数", String.join("\n", wariVals));
            setJuchuSheetString(targetRow, colMap, "投入場所", String.join("\n", feedVals));
            setJuchuSheetString(targetRow, colMap, "在庫場所", String.join("\n", storageVals));

            setJuchuSheetCellIfPresent(
                    targetRow, colMap, COL_MASTER_BASE_SHOHIN_PRODUCT, joinMasterBaseShohinFromProductRows());
            setJuchuSheetCellIfPresent(
                    targetRow, colMap, COL_MASTER_BASE_SHOHIN_RAW, joinMasterBaseShohinFromRawRows());
            
            setJuchuSheetString(targetRow, colMap, "加工内容", txtProcess.getText().trim());
            setJuchuSheetNumericOrString(targetRow, colMap, "加工賃", newTxtFormWage.getText().trim());
            setJuchuSheetString(targetRow, colMap, "契約Ｎｏ", newTxtFormContractNo.getText().trim());
            setJuchuSheetString(
                    targetRow,
                    colMap,
                    "用途",
                    newCmbFormYoto.getValue() != null ? newCmbFormYoto.getValue().trim() : "");
            if (isNewRow) {
                applyDefaultJuchuFormulasIfMissing(targetRow, colMap, targetRowIndex);
                Map<String, String> raw = selectedRecord != null ? selectedRecord.getRawValues() : new HashMap<>();
                if (!raw.isEmpty()) {
                    setJuchuSheetString(targetRow, colMap, "品名1", raw.get("原反品名"));
                    setJuchuSheetString(targetRow, colMap, "梱-等1", raw.get("梱-等1"));
                    setJuchuSheetString(targetRow, colMap, "色1", raw.get("色1"));
                    setJuchuSheetString(targetRow, colMap, "区分1", raw.get("区分1"));
                    setJuchuSheetString(targetRow, colMap, "梱-等", raw.get("原反梱-等"));
                    setJuchuSheetString(targetRow, colMap, "色", raw.get("原反色"));
                    setJuchuSheetString(targetRow, colMap, "区分", raw.get("原反区分"));
                    setJuchuSheetString(targetRow, colMap, "在庫場所", raw.get("在庫場所"));
                    setJuchuSheetNumericOrString(targetRow, colMap, "数量", raw.get("原反数量"));
                    setJuchuSheetNumericOrString(targetRow, colMap, "加工賃", raw.get("加工賃"));
                    setJuchuSheetNumericOrString(targetRow, colMap, "原反ロール数", raw.get("原反ロール数"));
                }
            }
            
            fis.close();
            FileOutputStream fos = new FileOutputStream(file);
            wb.write(fos);
            fos.close();
            wb.close();
            
            statusLabel.setText("受注ファイルの更新が完了しました。");
            showAlert("成功", "受注ファイルにデータが一括転記・保存されました！");
            
            if (selectedRecord != null) {
                selectedRecord.setStatus("一致 (転記完了)");
                selectedRecord.setDiscrepancy("一致 (受注ファイルへ転記完了)");
                
                // Refresh the comboRecord items display snappily
                ObservableList<OrderRecord> currentItems = comboRecord.getItems();
                comboRecord.setItems(null);
                comboRecord.setItems(currentItems);
                comboRecord.getSelectionModel().select(selectedRecord);
                
                // Update discrepancy label in UI
                discrepancyLabel.setText("一致 (受注ファイルへ転記完了)");
                discrepancyLabel.getStyleClass().removeAll("discrepancy-label-info", "discrepancy-label-matched", "discrepancy-label-discrepancy", "discrepancy-label-unregistered");
                discrepancyLabel.getStyleClass().add("discrepancy-label-matched");
            }
            updateTransferButtonState();
            
        } catch (Exception ex) {
            statusLabel.setText("エラー: " + ex.getMessage());
            showAlert("エラー", "ファイルの書き込み中にエラーが発生しました:\n" + ex.getMessage());
            ex.printStackTrace();
            updateTransferButtonState();
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
     * 受注ﾌｧｲﾙ先頭へ行を挿入し、直下のテンプレート行から数式・書式を複製する（値は複製しない）。
     *
     * @return Excel 上の 1 始まり行番号（数式の行参照用）
     */
    private static int insertNewJuchuDataRowPreservingFormulas(Sheet sheet) {
        int lastRowIdx = sheet.getLastRowNum();
        if (lastRowIdx >= JUCHU_SHEET_FIRST_DATA_ROW_INDEX) {
            sheet.shiftRows(JUCHU_SHEET_FIRST_DATA_ROW_INDEX, lastRowIdx, 1, true, true);
        }
        int destRowIdx = JUCHU_SHEET_FIRST_DATA_ROW_INDEX;
        int templateRowIdx = destRowIdx + 1;
        copyJuchuTemplateRowFormulasAndStyles(sheet, templateRowIdx, destRowIdx);
        return destRowIdx + 1;
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
                        compactFormLabelColumn(),
                        expandingFormFieldColumn(),
                        compactFormLabelColumn(),
                        expandingFormFieldColumn());
    }

    private static ColumnConstraints compactFormLabelColumn() {
        ColumnConstraints cc = new ColumnConstraints();
        cc.setMinWidth(Region.USE_COMPUTED_SIZE);
        cc.setPrefWidth(Region.USE_COMPUTED_SIZE);
        cc.setMaxWidth(Region.USE_PREF_SIZE);
        cc.setHalignment(javafx.geometry.HPos.RIGHT);
        return cc;
    }

    private static ColumnConstraints expandingFormFieldColumn() {
        ColumnConstraints cc = new ColumnConstraints();
        cc.setHgrow(Priority.ALWAYS);
        cc.setMinWidth(72);
        cc.setFillWidth(true);
        return cc;
    }

    private static void styleFormLabel(Label label) {
        label.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
        GridPane.setHalignment(label, javafx.geometry.HPos.RIGHT);
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

        Label lblItem = new Label("商品:");
        styleFormLabel(lblItem);
        pRow.txtItem = new TextField();
        pRow.txtItem.setStyle("-fx-font-size: 11px;");
        Button btnCopy = new Button("📋");
        btnCopy.setStyle("-fx-font-size: 11px; -fx-padding: 2px 6px; -fx-cursor: hand;");
        btnCopy.getStyleClass().add("btn-copy");
        btnCopy.setOnAction(evt -> copyToClipboard(pRow.txtItem.getText(), btnCopy));
        HBox itemBox = new HBox(5, pRow.txtItem, btnCopy);
        configureSplitFieldRow(itemBox, pRow.txtItem, btnCopy);
        pRow.grid.add(lblItem, 0, 0);
        addFormField(pRow.grid, itemBox, 1, 0);

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
        pRow.grid.add(lblPart, 2, 0);
        addFormField(pRow.grid, partBox, 3, 0);

        Label lblType = new Label("タイプ:");
        styleFormLabel(lblType);
        pRow.txtType = new TextField();
        pRow.txtType.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblType, 0, 1);
        addFormField(pRow.grid, pRow.txtType, 1, 1);

        Label lblWidth = new Label("幅(mm):");
        styleFormLabel(lblWidth);
        pRow.txtWidth = new TextField();
        pRow.txtWidth.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblWidth, 2, 1);
        addFormField(pRow.grid, pRow.txtWidth, 3, 1);

        Label lblLength = new Label("長さ(m):");
        styleFormLabel(lblLength);
        pRow.txtLength = new TextField();
        pRow.txtLength.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblLength, 0, 2);
        addFormField(pRow.grid, pRow.txtLength, 1, 2);

        Label lblQty = new Label("数量:");
        styleFormLabel(lblQty);
        pRow.txtQty = new TextField();
        pRow.txtQty.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblQty, 2, 2);
        addFormField(pRow.grid, pRow.txtQty, 3, 2);

        Label lblGrade = new Label("梱-等:");
        styleFormLabel(lblGrade);
        pRow.txtGrade = new TextField();
        pRow.txtGrade.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblGrade, 0, 3);
        addFormField(pRow.grid, pRow.txtGrade, 1, 3);

        Label lblColor = new Label("色:");
        styleFormLabel(lblColor);
        pRow.txtColor = new TextField();
        pRow.txtColor.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblColor, 2, 3);
        addFormField(pRow.grid, pRow.txtColor, 3, 3);

        Label lblCategory = new Label("区分:");
        styleFormLabel(lblCategory);
        pRow.txtCategory = new TextField();
        pRow.txtCategory.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblCategory, 0, 4);
        addFormField(pRow.grid, pRow.txtCategory, 1, 4);

        Label lblEdaban = new Label("枝番:");
        styleFormLabel(lblEdaban);
        pRow.txtEdaban = new TextField();
        pRow.txtEdaban.setStyle("-fx-font-size: 11px;");
        pRow.txtEdaban.getStyleClass().add("paper-edaban-input");
        pRow.grid.add(lblEdaban, 2, 4);
        addFormField(pRow.grid, pRow.txtEdaban, 3, 4);

        Label lblEc = new Label("EC面:");
        styleFormLabel(lblEc);
        pRow.cmbEcSide = new ComboBox<>(FXCollections.observableArrayList(optEcSide));
        pRow.cmbEcSide.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblEc, 0, 5);
        addFormField(pRow.grid, pRow.cmbEcSide, 1, 5);

        Label lblTrimming = new Label("トリミング:");
        styleFormLabel(lblTrimming);
        pRow.cmbTrimming = new ComboBox<>(FXCollections.observableArrayList(optTrimming));
        pRow.cmbTrimming.setStyle("-fx-font-size: 11px;");
        pRow.grid.add(lblTrimming, 2, 5);
        addFormField(pRow.grid, pRow.cmbTrimming, 3, 5);

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
        pRow.grid.add(btnDelete, 3, 6);
        GridPane.setHalignment(btnDelete, javafx.geometry.HPos.RIGHT);

        pRow.txtPart.textProperty().addListener((obs, oldV, newV) -> updateRowProdCandidates(pRow, false));
        pRow.txtItem.textProperty().addListener((obs, oldV, newV) -> updateRowProdCandidates(pRow, false));
        pRow.txtType.textProperty().addListener((obs, oldV, newV) -> updateRowProdCandidates(pRow, false));
        pRow.txtLength.textProperty().addListener((obs, oldV, newV) -> updateRowProdCandidates(pRow, false));
        
        pRow.cmbSearch.setOnAction(evt -> {
            String sel = pRow.cmbSearch.getValue();
            if (sel != null && sel.contains(" | ")) {
                String code = sel.split(" \\| ")[0].trim();
                for (ProductInfo p : masterProductList) {
                    if (p.getShohinCode().equals(code)) {
                        isLoadingRecord = true;
                        try {
                            pRow.txtItem.setText(p.getShohinCode());
                            pRow.txtPart.setText(p.getFoamPartNo());
                            String[] nameParts = p.getShohinName1().split("-");
                            if (nameParts.length >= 2) pRow.txtType.setText(nameParts[1]);
                            pRow.txtWidth.setText(p.getFoamWidth());
                            pRow.txtLength.setText(p.getFoamLength());
                            newTxtFormKakoNaiyo.setText(p.getKakoNaiyo());
                        } finally {
                            isLoadingRecord = false;
                        }
                        break;
                    }
                }
            }
        });

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

    private static void wireCandidateComboBox(ComboBox<String> combo, Runnable refreshOnOpen) {
        combo.setOnShowing(
                e -> {
                    if (refreshOnOpen != null) {
                        refreshOnOpen.run();
                    }
                });
    }

    private void updateRowProdCandidates(ProductRow pRow, boolean fromDropdownOpen) {
        if (isLoadingRecord && !fromDropdownOpen) {
            return;
        }
        String kwItem = normalize_text(pRow.txtItem.getText());
        String kwPart = normalize_text(pRow.txtPart.getText());
        String kwType = normalize_text(pRow.txtType.getText());
        String kwLength = normalize_text(pRow.txtLength.getText());

        java.util.List<String> filtered;
        if (kwItem.isEmpty() && kwPart.isEmpty() && kwType.isEmpty() && kwLength.isEmpty()) {
            if (!fromDropdownOpen) {
                pRow.cmbSearch.setItems(javafx.collections.FXCollections.emptyObservableList());
                return;
            }
            filtered = buildMasterProductCandidateLabels(null, null, null, null, 50);
        } else {
            filtered = buildMasterProductCandidateLabels(kwItem, kwPart, kwType, kwLength, 50);
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

        java.util.List<String> filtered;
        if (kwItem.isEmpty() && kwPart.isEmpty() && kwType.isEmpty() && kwLength.isEmpty()) {
            if (!fromDropdownOpen) {
                rRow.cmbSearch.setItems(javafx.collections.FXCollections.emptyObservableList());
                return;
            }
            filtered = buildMasterProductCandidateLabels(null, null, null, null, 50);
        } else {
            filtered = buildMasterProductCandidateLabels(kwItem, kwPart, kwType, kwLength, 50);
        }
        rRow.cmbSearch.setItems(javafx.collections.FXCollections.observableArrayList(filtered));
        if (!fromDropdownOpen && !filtered.isEmpty() && !rRow.cmbSearch.isShowing()) {
            rRow.cmbSearch.show();
        }
    }

    private java.util.List<String> buildMasterProductCandidateLabels(
            String kwItem, String kwPart, String kwType, String kwLength, int limit) {
        java.util.List<String> filtered = new java.util.ArrayList<>();
        boolean filterActive =
                kwItem != null && !kwItem.isEmpty()
                        || kwPart != null && !kwPart.isEmpty()
                        || kwType != null && !kwType.isEmpty()
                        || kwLength != null && !kwLength.isEmpty();
        String normItem = kwItem != null ? kwItem : "";
        String normPart = kwPart != null ? kwPart : "";
        String normType = kwType != null ? kwType : "";
        String normLength = kwLength != null ? kwLength : "";

        for (ProductInfo p : masterProductList) {
            if (filterActive) {
                boolean matchItem =
                        normItem.isEmpty()
                                || (p.getShohinCode() != null
                                        && normalize_text(p.getShohinCode()).contains(normItem))
                                || (p.getShohinName1() != null
                                        && normalize_text(p.getShohinName1()).contains(normItem))
                                || (p.getSeihinCode() != null
                                        && normalize_text(p.getSeihinCode()).contains(normItem))
                                || (p.getFoamName() != null
                                        && normalize_text(p.getFoamName()).contains(normItem));
                boolean matchPart =
                        normPart.isEmpty()
                                || (p.getFoamPartNo() != null
                                        && normalize_text(p.getFoamPartNo()).contains(normPart))
                                || normalize_text(p.getShohinCode()).contains(normPart);
                boolean matchType =
                        normType.isEmpty()
                                || (p.getFoamName() != null
                                        && normalize_text(p.getFoamName()).contains(normType))
                                || (p.getShohinName1() != null
                                        && normalize_text(p.getShohinName1()).contains(normType));
                String pLength = normalize_text(p.getFoamLength()).replaceAll("\\.0$", "");
                boolean matchLength = normLength.isEmpty() || pLength.contains(normLength);
                if (!(matchItem && matchPart && matchType && matchLength)) {
                    continue;
                }
            }
            filtered.add(formatMasterProductCandidateLabel(p));
            if (filtered.size() >= limit) {
                break;
            }
        }
        return filtered;
    }

    private static String formatMasterProductCandidateLabel(ProductInfo p) {
        String pLength = p.getFoamLength() != null ? p.getFoamLength().replaceAll("\\.0$", "") : "";
        String pWidth = p.getFoamWidth() != null ? p.getFoamWidth().replaceAll("\\.0$", "") : "";
        String dims = (pWidth.isEmpty() ? "?" : pWidth) + "×" + (pLength.isEmpty() ? "?" : pLength);
        return p.getShohinCode()
                + " | "
                + p.getFoamPartNo()
                + " | "
                + p.getFoamName()
                + " | "
                + dims;
    }

    private RawMaterialRow addRawRow(Map<String, String> initialValues) {
        RawMaterialRow rRow = new RawMaterialRow();
        rRow.grid = new GridPane();
        rRow.grid.setVgap(6);
        rRow.grid.setMaxWidth(Double.MAX_VALUE);
        rRow.grid.getStyleClass().add("request-form-prod-row");
        configureTwoColumnFormGrid(rRow.grid);

        Label lblItem = new Label("商品:");
        styleFormLabel(lblItem);
        rRow.txtItem = new TextField();
        rRow.txtItem.setStyle("-fx-font-size: 11px;");
        Button btnCopy = new Button("📋");
        btnCopy.setStyle("-fx-font-size: 11px; -fx-padding: 2px 6px; -fx-cursor: hand;");
        btnCopy.getStyleClass().add("btn-copy");
        btnCopy.setOnAction(evt -> copyToClipboard(rRow.txtItem.getText(), btnCopy));
        HBox itemBox = new HBox(5, rRow.txtItem, btnCopy);
        configureSplitFieldRow(itemBox, rRow.txtItem, btnCopy);
        rRow.grid.add(lblItem, 0, 0);
        addFormField(rRow.grid, itemBox, 1, 0);

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
        rRow.grid.add(lblPart, 2, 0);
        addFormField(rRow.grid, partBox, 3, 0);

        Label lblType = new Label("タイプ:");
        styleFormLabel(lblType);
        rRow.txtType = new TextField();
        rRow.txtType.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblType, 0, 1);
        addFormField(rRow.grid, rRow.txtType, 1, 1);

        Label lblWidth = new Label("幅(mm):");
        styleFormLabel(lblWidth);
        rRow.txtWidth = new TextField();
        rRow.txtWidth.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblWidth, 2, 1);
        addFormField(rRow.grid, rRow.txtWidth, 3, 1);

        Label lblLength = new Label("長さ(m):");
        styleFormLabel(lblLength);
        rRow.txtLength = new TextField();
        rRow.txtLength.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblLength, 0, 2);
        addFormField(rRow.grid, rRow.txtLength, 1, 2);

        Label lblQty = new Label("数量:");
        styleFormLabel(lblQty);
        rRow.txtQty = new TextField();
        rRow.txtQty.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblQty, 2, 2);
        addFormField(rRow.grid, rRow.txtQty, 3, 2);

        Label lblGrade = new Label("梱-等:");
        styleFormLabel(lblGrade);
        rRow.txtGrade = new TextField();
        rRow.txtGrade.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblGrade, 0, 3);
        addFormField(rRow.grid, rRow.txtGrade, 1, 3);

        Label lblColor = new Label("色:");
        styleFormLabel(lblColor);
        rRow.txtColor = new TextField();
        rRow.txtColor.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblColor, 2, 3);
        addFormField(rRow.grid, rRow.txtColor, 3, 3);

        Label lblCategory = new Label("区分:");
        styleFormLabel(lblCategory);
        rRow.txtCategory = new TextField();
        rRow.txtCategory.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblCategory, 0, 4);
        addFormField(rRow.grid, rRow.txtCategory, 1, 4);

        Label lblWariSu = new Label("割数:");
        styleFormLabel(lblWariSu);
        rRow.cmbWariSu = new ComboBox<>(FXCollections.observableArrayList(optWariSu));
        rRow.cmbWariSu.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblWariSu, 2, 4);
        addFormField(rRow.grid, rRow.cmbWariSu, 3, 4);

        Label lblFeed = new Label("投入場所:");
        styleFormLabel(lblFeed);
        rRow.cmbFeedLoc = new ComboBox<>(FXCollections.observableArrayList(optFeedLoc));
        rRow.cmbFeedLoc.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblFeed, 0, 5);
        addFormField(rRow.grid, rRow.cmbFeedLoc, 1, 5);

        Label lblStorage = new Label("在庫場所:");
        styleFormLabel(lblStorage);
        rRow.cmbStorageLoc = new ComboBox<>(FXCollections.observableArrayList(optStorageLoc));
        rRow.cmbStorageLoc.setStyle("-fx-font-size: 11px;");
        rRow.grid.add(lblStorage, 2, 5);
        addFormField(rRow.grid, rRow.cmbStorageLoc, 3, 5);

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
        rRow.grid.add(btnDelete, 3, 6);
        GridPane.setHalignment(btnDelete, javafx.geometry.HPos.RIGHT);

        rRow.txtPart.textProperty().addListener((obs, oldV, newV) -> updateRowRawCandidates(rRow, false));
        rRow.txtItem.textProperty().addListener((obs, oldV, newV) -> updateRowRawCandidates(rRow, false));
        rRow.txtType.textProperty().addListener((obs, oldV, newV) -> updateRowRawCandidates(rRow, false));
        rRow.txtLength.textProperty().addListener((obs, oldV, newV) -> updateRowRawCandidates(rRow, false));
        
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
