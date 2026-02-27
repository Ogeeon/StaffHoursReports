package staffhoursreports;

import com.typesafe.config.Config;
import com.typesafe.config.ConfigFactory;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.control.*;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.sql.*;
import java.time.LocalDate;
import java.util.Date;
import java.util.List;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.prefs.Preferences;

public class RootPaneView implements Initializable {
    private static final String ZI_SRC_FN_KEY = "zi_src_fn";
    private static final String ZI_DST_DIR_KEY = "zi_dst_dir";
    private static final String MSG_EMPTY_FIELD = "Строка %d: пустое поле '%s', пропускаем запись";
    private static final String MSG_EMPTY_DATE  = "Строка %d: пустое или некорректное поле 'Дата', пропускаем запись";
    private static final String REPORT_FONT_NAME = "Times New Roman";
    private static final String HOURS_SQL =
        "select Users.FIO usrFIO, Tasks.taskName, hrt.totals, Tasks.CreationDate, Tasks.extRefNum" +
        ", Requesters.Organization, Requesters.FIO reqFIO, Users.category " +
        "from (" +
            "select User_id, Task_id, sum(hr) totals " +
            "from (" +
                "select User_id, Task_id, hr, dat " +
                "from TSRecordViews " +
                "unpivot ((hr, dat) for day in (" +
                    "(hr1, dat1) as '1', (hr2, dat2) as '2', (hr3, dat3) as '3', " +
                    "(hr4, dat4) as '4', (hr5, dat5) as '5', (hr6, dat6) as '6', " +
                    "(hr7, dat7) as '7'" +
                "))" +
            ") " +
            "where ? <= dat and dat <= ? and hr > 0 " +
            "group by User_id, Task_id" +
        ") hrt " +
        "left join Tasks on (Tasks.id=hrt.Task_id) " +
        "left join Users on (Users.id=hrt.User_id) " +
        "left join Requesters on (Requesters.id = tasks.requester_id) " +
        "ORDER BY 1, 2";
    private record SavedTask(
        Integer taskId, String extRefNum, String executor,
        Integer executorId, Integer requesterId, String requester,
        String taskName, Date creationDate
    ) {
        @Override
        public String toString() {
            return "[" + taskId() + ", " + extRefNum() + ", " + requesterId() + "]";
        }
    }
    private record RowData(
            String taskName, String executorFI, Date creationDate,
            String extRefNum, String requester, String organization
    ) {}
    private record ReportRecord(
        String userFIO, String taskName, int totals,
        LocalDate creationDate, String extRefNum,
        String requesterOrg, String requesterName, String userCategory
    ) {}
    Connection connection = null;
    private final FileChooser fileChooser = new FileChooser();
    private final Preferences preferences = Preferences.userNodeForPackage(RootPaneView.class);
    private Config costsConfig;

    @FXML
    private BorderPane topPane;
    @FXML
    private Accordion accordion;
    @FXML
    private TitledPane tpImport;
    @FXML
    private TextField tfInputFileName;
    @FXML
    private ProgressBar pbImport;
    @FXML
    private HBox hbProgress;
    @FXML
    private Label lblOutputDir;
    @FXML
    private TextField tfOutputFileName;
    @FXML
    private DatePicker dtpckStart;
    @FXML
    private DatePicker dtpckEnd;

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        accordion.setExpandedPane(tpImport);
        fileChooser.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("Файлы Excel", "*.xlsx"));
        tfInputFileName.setText(preferences.get(ZI_SRC_FN_KEY, "i:\\УИТ\\ОП\\_Регламентная отчетность\\Отчет по ЗИ\\ЗИ_2020-КНПЗ.xlsx"));
        lblOutputDir.setText(preferences.get(ZI_DST_DIR_KEY, "i:\\УИТ\\ОП\\_Регламентная отчетность\\Отчет по ЗИ"));
        LocalDate currDate = LocalDate.now();
        if (currDate.getDayOfMonth() < 15)
            currDate = currDate.minusMonths(1);
        LocalDate dateTo = LocalDate.of(currDate.getYear(), currDate.getMonth(), 14);
        LocalDate dateFrom = dateTo.minusMonths(1).plusDays(1);
        dtpckStart.setValue(dateFrom);
        dtpckEnd.setValue(dateTo);
        tfOutputFileName.setText(getReportName(dateFrom, dateTo));
        costsConfig = ConfigFactory.load().getConfig("costs");
    }

    private String getCellAsString(Cell cell) {
        return ExcelUtils.getCellAsString(cell);
    }

    private Date getCellAsDate(Cell cell) {
        return ExcelUtils.getCellAsDate(cell);
    }

    public boolean connect() {
        try {
            Locale.setDefault(Locale.ENGLISH);
            String host = System.getenv("STAFFHOURS_DB_HOST");
            String service = System.getenv("STAFFHOURS_DB_SERVICE");
            String login = System.getenv("STAFFHOURS_DB_USER");
            String password = System.getenv("STAFFHOURS_DB_PASSWORD");
            String connStr = String.format("jdbc:oracle:thin:@%s:1521:%s", host, service);
            connection = DriverManager.getConnection(connStr, login, password);
        } catch (SQLException e) {
            Utils.showErrorAndStack(e);
            return false;
        }
        return true;
    }

    @FXML
    private void handleBrowseSrcClick() {
        Stage root = ((Stage) topPane.getScene().getWindow());
        fileChooser.setTitle("Открыть отчёт по ЗИ");
        String srcFn = preferences.get(ZI_SRC_FN_KEY, "");
        if (!srcFn.isEmpty()) {
            fileChooser.setInitialDirectory((new File(srcFn)).getParentFile());
        }
        File source = fileChooser.showOpenDialog(root);
        if (source != null) {
            tfInputFileName.setText(source.getPath());
            preferences.put(ZI_SRC_FN_KEY, source.getPath());
        }
    }

    @FXML
    private void handleImportClick() {
        if (tfInputFileName.getText().isEmpty()) {
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setHeaderText("Не выбран файл.");
            alert.setTitle("Ошибка");
            alert.showAndWait();
            return;
        }
        Task<List<String>> task = new Task<>() {
            @Override
            protected List<String> call() throws ImportException {
                return performImport(this::updateProgress);
            }
        };
        task.progressProperty().addListener(
            (obs, oldProgress, newProgress) -> pbImport.setProgress(newProgress.doubleValue()));
        task.setOnSucceeded(e -> onImportSuccess(task.getValue()));
        task.setOnFailed(e -> onImportFailed(task.getException()));
        new Thread(task).start();
    }

    private List<String> performImport(BiConsumer<Long, Long> progressUpdater) throws ImportException {
        List<String> messages = new ArrayList<>();
        try {
            File file = new File(tfInputFileName.getText());
            try (FileInputStream fis = new FileInputStream(file);
                 XSSFWorkbook wb = new XSSFWorkbook(fis)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                int totalRows = sheet.getLastRowNum();
                if (totalRows < 2)
                    return new ArrayList<>();
                else {
                    pbImport.setProgress(0);
                    hbProgress.setVisible(true);
                }
                for (Row row : sheet) {
                    processSheetRow(row, totalRows, messages, progressUpdater);
                }
            }
        } catch (IOException e) {
            messages.add("Критическая ошибка ввода-вывода: " + e.getMessage());
            throw new ImportException("IO error during import", e);
        } catch (Exception e) {
            messages.add("Критическая ошибка: " + e.getMessage());
            throw new ImportException("Unexpected error during import", e);
        }
        return messages;
    }

    private void onImportSuccess(List<String> messages) {
        hbProgress.setVisible(false);
        if (messages != null && !messages.isEmpty()) {
            StringBuilder sb = new StringBuilder();
            sb.append("Выполненные действия:\n\n");
            for (String msg : messages) {
                sb.append("• ").append(msg).append("\n");
            }
            Utils.showInfo(sb.toString());
        } else {
            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.setHeaderText("Обработка файла завершена.");
            alert.setTitle("Готово");
            alert.showAndWait();
        }
    }

    private void onImportFailed(Throwable exception) {
        hbProgress.setVisible(false);
        if (exception != null) {
            Utils.showErrorAndStack((Exception) exception);
        } else {
            Utils.showError("Произошла неизвестная ошибка при обработке файла.");
        }
    }

    private void processSheetRow(Row row, int totalRows, List<String> messages,
            BiConsumer<Long, Long> progressUpdater) {
        try {
            RowData data = parseRowData(row, messages);
            if (data == null) return;

            SavedTask st = getTask(data.extRefNum());
            if (st != null) {
                updateExistingTask(data, st, row.getRowNum(), messages);
            } else {
                int rid = getRequesterID(data.requester(), data.organization(), messages);
                if (rid != 0) {
                    insertTask(data.taskName(), data.executorFI(), data.creationDate(),
                            data.extRefNum(), rid, messages);
                }
            }
            progressUpdater.accept((long) row.getRowNum(), (long) totalRows);
        } catch (SQLException e) {
            messages.add("Ошибка при обработке строки " + (row.getRowNum() + 1) + ": " + e.getMessage());
        }
    }

    private RowData parseRowData(Row row, List<String> messages) {
        String taskName = getCellAsString(row.getCell(0));
        if (taskName == null || taskName.equals("Тема")) return null;
        String executorFI = getCellAsString(row.getCell(1));
        if (executorFI == null || executorFI.isEmpty()) {
            messages.add(String.format(MSG_EMPTY_FIELD, row.getRowNum() + 1, "Исполнитель"));
            return null;
        }
        Date creationDate = getCellAsDate(row.getCell(2));
        if (creationDate == null) {
            messages.add(String.format(MSG_EMPTY_DATE, row.getRowNum() + 1));
            return null;
        }
        String extRefNum = getCellAsString(row.getCell(3));
        if (extRefNum == null || extRefNum.isEmpty()) {
            messages.add(String.format(MSG_EMPTY_FIELD, row.getRowNum() + 1, "№ обращения"));
            return null;
        }
        String requester = getCellAsString(row.getCell(4));
        if (requester == null || requester.isEmpty()) {
            messages.add(String.format(MSG_EMPTY_FIELD, row.getRowNum() + 1, "Заявитель"));
            return null;
        }
        String organization = getCellAsString(row.getCell(5));
        if (organization == null || organization.isEmpty()) {
            messages.add(String.format(MSG_EMPTY_FIELD, row.getRowNum() + 1, "Организация"));
            return null;
        }
        return new RowData(taskName, executorFI, creationDate, extRefNum, requester, organization);
    }

    private void updateExistingTask(RowData data, SavedTask st, int rowIdx,
            List<String> messages) throws SQLException {

        // 1. Detect
        boolean taskNameChanged     = hasTaskNameChanged(data.taskName(), st);
        boolean creationDateChanged = st.creationDate() != null
                && !st.creationDate().equals(data.creationDate());
        boolean executorChanged     = !st.executor().equals(data.executorFI());
        int newRequesterId          = getRequesterID(data.requester(), data.organization(), messages);
        boolean requesterChanged    = st.requesterId() != newRequesterId;

        // 2. Validate
        int newExecutorId = 0;
        if (executorChanged) {
            newExecutorId = getExecutorID(data.executorFI(), messages);
            if (newExecutorId == 0) {
                messages.add("Строка " + rowIdx + ": не удалось определить исполнителя: "
                        + data.executorFI() + ", пропускаем запись");
                return;
            }
        }
        if (requesterChanged && newRequesterId == 0) return;  // getRequesterID already logged

        // 3. Execute + 4. Report
        StringBuilder updates = new StringBuilder();

        if (taskNameChanged) {
            updateTaskNameField(st.taskId(), data.taskName());
            appendSep(updates);
            updates.append("название: '").append(st.taskName()).append("' -> '").append(data.taskName()).append("'");
        }
        if (creationDateChanged) {
            updateCreationDateField(st.taskId(), data.creationDate());
            appendSep(updates);
            String oldDate = Utils.localizeDate(Utils.toLocalDate(st.creationDate()), Locale.forLanguageTag("ru"));
            String newDate = Utils.localizeDate(Utils.toLocalDate(data.creationDate()), Locale.forLanguageTag("ru"));
            updates.append("дата: ").append(oldDate).append(" -> ").append(newDate);
        }
        if (executorChanged) {
            updateExecutorField(st.taskId(), newExecutorId);
            appendSep(updates);
            updates.append("исполнитель: ").append(st.executor()).append(" -> ").append(data.executorFI());
        }
        if (requesterChanged) {
            updateRequesterField(st.taskId(), newRequesterId);
            appendSep(updates);
            updates.append("заявитель: ").append(st.requester()).append(" -> ").append(data.requester());
        }

        if (!updates.isEmpty())
            messages.add("Обновлена задача " + data.extRefNum() + ": " + updates);
    }

    private static void appendSep(StringBuilder sb) {
        if (!sb.isEmpty()) sb.append(", ");
    }

    private void updateExecutorField(int taskId, int executorId) throws SQLException {
        try (PreparedStatement pstmt = connection.prepareStatement(
                "UPDATE Tasks SET executor_id=? WHERE id=?")) {
            pstmt.setInt(1, executorId);
            pstmt.setInt(2, taskId);
            pstmt.executeUpdate();
        }
    }

    private void updateRequesterField(int taskId, int requesterId) throws SQLException {
        try (PreparedStatement pstmt = connection.prepareStatement(
                "UPDATE Tasks SET requester_id=? WHERE id=?")) {
            pstmt.setInt(1, requesterId);
            pstmt.setInt(2, taskId);
            pstmt.executeUpdate();
        }
    }

    private void updateTaskNameField(int taskId, String taskName) throws SQLException {
        try (PreparedStatement pstmt = connection.prepareStatement(
                "UPDATE Tasks SET taskName=? WHERE id=?")) {
            pstmt.setString(1, taskName);
            pstmt.setInt(2, taskId);
            pstmt.executeUpdate();
        }
    }

    private void updateCreationDateField(int taskId, Date creationDate) throws SQLException {
        try (PreparedStatement pstmt = connection.prepareStatement(
                "UPDATE Tasks SET creationDate=? WHERE id=?")) {
            pstmt.setDate(1, new java.sql.Date(creationDate.getTime()));
            pstmt.setInt(2, taskId);
            pstmt.executeUpdate();
        }
    }

    private SavedTask getTask(String extRefNum) throws SQLException {
        SavedTask st = null;
        try (PreparedStatement pstmt = connection.prepareStatement(
                "SELECT t.id, t.executor_id, t.requester_id, u.fio, r.fio reqfio, t.taskName, t.creationDate " +
                "FROM Tasks t " +
                "INNER JOIN Users u ON t.executor_id = u.id " +
                "LEFT JOIN Requesters r ON r.id = t.requester_id " +
                "WHERE extRefNum = ?")) {
            pstmt.setString(1, extRefNum);
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    st = new SavedTask(rs.getInt("id"), extRefNum, rs.getString("fio"), rs.getInt("executor_id"),
                                    rs.getInt("requester_id"), rs.getString("reqfio"), rs.getString("taskName"), rs.getDate("creationDate"));
                }
            }
        }
        return st;
    }

    private void insertTask(String taskName, String executorFI, Date creationDate, String extRefNum, int requesterId, List<String> messages) throws SQLException {
        int executorId = getExecutorID(executorFI, messages);
        try (PreparedStatement pstmt = connection.prepareStatement("INSERT INTO Tasks (taskName, creationDate, extRefNum, executor_id, requester_id)" +
                " VALUES (?, ?, ?, ?, ?)", Statement.RETURN_GENERATED_KEYS)) {
            pstmt.setString(1, taskName);
            pstmt.setDate(2, new java.sql.Date(creationDate.getTime()));
            pstmt.setString(3, extRefNum);
            pstmt.setInt(4, executorId);
            pstmt.setInt(5, requesterId);
            int r = pstmt.executeUpdate();
            if (r > 0) {
                messages.add("Добавлена задача: [" + extRefNum + " - " + taskName + "]");
            }
        }
    }

    private int getExecutorID(String name, List<String> messages) throws SQLException {
        try (PreparedStatement pstmt = connection.prepareStatement("SELECT id FROM Users WHERE FIO = ?")) {
            pstmt.setString(1, name);
            try (ResultSet rs = pstmt.executeQuery()) {
                int id = 0;
                while (rs.next()) {
                    id = rs.getInt("id");
                }
                if (id == 0) {
                    messages.add("Не найден сотрудник: " + name);
                }
                return id;
            }
        }
    }

    private int getRequesterID(String name, String organization, List<String> messages) throws SQLException {
        // First try to find existing requester
        int id = 0;
        String query = "SELECT id FROM Requesters WHERE FIO = ? AND Organization = ?";
        try (PreparedStatement pstmt = connection.prepareStatement(query)) {
            pstmt.setString(1, name);
            pstmt.setString(2, organization);
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    id = rs.getInt("id");
                }
            }
        } catch (SQLException e) {
            messages.add("Ошибка поиска заявителя [" + name + ", " + organization + "]: " + e.getMessage());
            throw e;
        }

        if (id == 0) {
            // Insert new requester
            String insert = "INSERT INTO Requesters (FIO, Organization) VALUES (?, ?)";
            try (PreparedStatement insertStmt = connection.prepareStatement(insert, new String[]{"id"})) {
                insertStmt.setString(1, name);
                insertStmt.setString(2, organization);
                insertStmt.executeUpdate();
                try (ResultSet ins = insertStmt.getGeneratedKeys()) {
                    if (ins.next()) {
                        int lastInsertedId = ins.getInt(1);
                        messages.add("Добавлен заявитель: [" + name + ", " + organization + "]");
                        id = lastInsertedId;
                    }
                }
            } catch (SQLException e) {
                messages.add("Ошибка добавления заявителя [" + name + ", " + organization + "]: " + e.getMessage());
                throw e;
            }
        }
        return id;
    }

    @FXML
    private void handleBrowseDstClick() {
        Stage root = ((Stage) topPane.getScene().getWindow());
        fileChooser.setTitle("Сохранить отчёт о трудозатратах");
        String dstDir = preferences.get(ZI_DST_DIR_KEY, "");
        if (!dstDir.isEmpty()) {
            fileChooser.setInitialDirectory((new File(dstDir)));
        }
        File dest = fileChooser.showSaveDialog(root);
        if (dest != null) {
            lblOutputDir.setText(dest.getParentFile().getPath());
            preferences.put(ZI_DST_DIR_KEY, dest.getParentFile().getPath());
            tfOutputFileName.setText(dest.getName());
        }
    }

    @FXML
    private void handleDateChange() {
        tfOutputFileName.setText(getReportName(dtpckStart.getValue(), dtpckEnd.getValue()));
    }

    private static String getReportName(LocalDate dt1, LocalDate dt2) {
        return Utils.getReportName(dt1, dt2);
    }

    @FXML
    private void handleGenerateReportClick() {
        File dstFile = new File(lblOutputDir.getText(), tfOutputFileName.getText());
        if (!confirmFileOverwrite(dstFile)) {
            return;
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Отчет");
        setupSheetColumns(sheet);
        
        putCaption(workbook, sheet);
        int lastRowNum = processRecords(workbook, sheet);
        putFooter(workbook, sheet, lastRowNum);
        
        saveAndOpenWorkbook(workbook);
    }

    private void putCaption(XSSFWorkbook workbook, Sheet sheet) {
        // --- Preamble rows ---
        XSSFFont preambleFont = workbook.createFont();
        preambleFont.setFontName(REPORT_FONT_NAME);
        preambleFont.setFontHeightInPoints((short) 11);
        preambleFont.setBold(true);
        CellStyle preambleStyle = workbook.createCellStyle();
        preambleStyle.setFont(preambleFont);
        preambleStyle.setAlignment(HorizontalAlignment.RIGHT);

        Row row1 = sheet.createRow(1);
        Cell preambleCell = row1.createCell(11);
        preambleCell.setCellValue("Приложение № 1 к акту 1094-02");
        preambleCell.setCellStyle(preambleStyle);

        Row row2 = sheet.createRow(2);
        preambleCell = row2.createCell(11);
        preambleCell.setCellValue("к Договору №  1С0925/01094Д   ");
        preambleCell.setCellStyle(preambleStyle);

        Row row3 = sheet.createRow(3);
        preambleCell = row3.createCell(11);
        preambleCell.setCellValue("от «___» ____________ 20___ г.");
        preambleCell.setCellStyle(preambleStyle);

        Row row6 = sheet.createRow(6);
        Cell titleCell = row6.createCell(3);
        titleCell.setCellValue("Отчет об оказанных услугах");
        CellStyle preambleStyleLeft = workbook.createCellStyle();
        preambleStyleLeft.setFont(preambleFont);
        preambleStyleLeft.setAlignment(HorizontalAlignment.LEFT);
        titleCell.setCellStyle(preambleStyleLeft);

        Row row8 = sheet.createRow(8);
        titleCell = row8.createCell(5);
        Locale ru = Locale.forLanguageTag("ru");
        String startFmt = Utils.localizeDate(dtpckStart.getValue(), ru);
        String endFmt   = Utils.localizeDate(dtpckEnd.getValue(), ru);
        // localizeDate returns e.g. "27 февраля 2026 г." — wrap day in guillemets
        startFmt = startFmt.replaceFirst("^(\\d+)", "«$1»");
        endFmt   = endFmt.replaceFirst("^(\\d+)", "«$1»");
        titleCell.setCellValue("Отчет об оказанных услугах за период с " + startFmt + " по " + endFmt);
        CellStyle preambleStyleCentered = workbook.createCellStyle();
        preambleStyleCentered.setFont(preambleFont);
        preambleStyleCentered.setAlignment(HorizontalAlignment.CENTER);
        titleCell.setCellStyle(preambleStyleCentered);

        Row row9 = sheet.createRow(9);
        titleCell = row9.createCell(5);
        titleCell.setCellValue("к Договору № _____ от «___» ____________ 20___ г.");
        titleCell.setCellStyle(preambleStyleCentered);

        XSSFFont headerFont = workbook.createFont();
        headerFont.setFontName(REPORT_FONT_NAME);
        headerFont.setFontHeightInPoints((short) 8);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setWrapText(true);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        // Row 17 (POI 16) — primary header
        Row hdr0 = sheet.createRow(16);
        hdr0.setHeightInPoints(45.0f);
        Cell cell;

        cell = hdr0.createCell(0);
        cell.setCellValue("Номер запроса на изменение");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(1);
        cell.setCellValue("Краткое описание запроса на изменение");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(2);
        cell.setCellValue("Результат работ");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(3);
        cell.setCellValue("Наименование и код информационного ресурса");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(4);
        cell.setCellValue("Наименование и код информационной системы");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(5);
        cell.setCellValue("Предприятие");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(6);
        cell.setCellValue("Категория Специалиста Исполнителя");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(7);
        cell.setCellValue("Согласованная ставка специалиста Исполнителя");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(8);
        cell.setCellValue("Плановая");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(9);
        cell.setCellValue("Фактическая трудоемкость");
        cell.setCellStyle(headerStyle);

        cell = hdr0.createCell(10);
        cell.setCellValue("Стоимость /работ (рубли)");
        cell.setCellStyle(headerStyle);

        // Row 18 (POI 17) — sub-headers level 1
        Row hdr1 = sheet.createRow(17);
        hdr1.setHeightInPoints(15.0f);

        cell = hdr1.createCell(7);
        cell.setCellValue("(рубли в час)");
        cell.setCellStyle(headerStyle);

        cell = hdr1.createCell(8);
        cell.setCellValue("трудоемкость");
        cell.setCellStyle(headerStyle);

        cell = hdr1.createCell(9);
        cell.setCellValue("(часы)");
        cell.setCellStyle(headerStyle);

        cell = hdr1.createCell(10);
        cell.setCellValue("без НДС");
        cell.setCellStyle(headerStyle);

        // Row 19 (POI 18) — sub-headers level 2
        Row hdr2 = sheet.createRow(18);
        hdr1.setHeightInPoints(15.0f);

        cell = hdr2.createCell(7);
        cell.setCellValue("без НДС");
        cell.setCellStyle(headerStyle);

        cell = hdr2.createCell(8);
        cell.setCellValue("(часы)");
        cell.setCellStyle(headerStyle);

        // Empty cells at the end of the table
        cell = hdr2.createCell(9);
        cell.setCellStyle(headerStyle);
        cell = hdr2.createCell(10);
        cell.setCellStyle(headerStyle);

        // --- Merged regions ---
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 0, 0)); // A17:A19
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 1, 1)); // B17:B19
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 2, 2)); // C17:C19
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 3, 3)); // D17:D19
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 4, 4)); // E17:E19
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 5, 5)); // F17:F19
        mergeWithBorders(sheet, new CellRangeAddress(16, 18, 6, 6)); // G17:G19
    }

    private void putFooter(XSSFWorkbook workbook, Sheet sheet, int startRow) {
        XSSFFont footerFont = workbook.createFont();
        footerFont.setFontName(REPORT_FONT_NAME);
        footerFont.setFontHeightInPoints((short) 8);
        CellStyle footerStyle = workbook.createCellStyle();
        footerStyle.setFont(footerFont);
        footerStyle.setWrapText(true);
        footerStyle.setAlignment(HorizontalAlignment.RIGHT);
        footerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        footerStyle.setBorderTop(BorderStyle.THIN);
        footerStyle.setBorderBottom(BorderStyle.THIN);
        footerStyle.setBorderLeft(BorderStyle.THIN);
        footerStyle.setBorderRight(BorderStyle.THIN);

        CellStyle footerNumStyle = workbook.createCellStyle();
        footerNumStyle.cloneStyleFrom(footerStyle);
        footerNumStyle.setAlignment(HorizontalAlignment.CENTER);
        footerNumStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        footerNumStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

        String[] labels = {"Итого без НДС:", "НДС:", "Итого с НДС:"};
        for (int i = 0; i < labels.length; i++) {
            int rowNum = startRow + i;
            Row row = sheet.createRow(rowNum);
            for (int col : new int[]{0, 8, 9}) {
                row.createCell(col).setCellStyle(footerStyle);
            }
            Cell labelCell = row.createCell(1);
            labelCell.setCellValue(labels[i]);
            labelCell.setCellStyle(footerStyle);
            mergeWithBorders(sheet, new CellRangeAddress(rowNum, rowNum, 1, 7));
            Cell sumCell = row.createCell(10);
            sumCell.setCellStyle(footerNumStyle);
            if (i == 0) {
                sumCell.setCellFormula("SUM(K20:K" + startRow + ")");
            }
        }
    }

    private static void mergeWithBorders(Sheet sheet, CellRangeAddress region) {
        sheet.addMergedRegion(region);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    private boolean confirmFileOverwrite(File dstFile) {
        if (!dstFile.exists()) {
            return true;
        }
        
        Alert alert = new Alert(Alert.AlertType.WARNING);
        alert.setTitle("Подтверждение");
        alert.setHeaderText("Такой файл уже существует. Перезаписать его?");
        ButtonType btnYes = new ButtonType("Да");
        ButtonType btnNo = new ButtonType("Нет");
        alert.getButtonTypes().setAll(btnYes, btnNo);
        ((Button) alert.getDialogPane().lookupButton(btnYes)).setDefaultButton(false);
        ((Button) alert.getDialogPane().lookupButton(btnNo)).setDefaultButton(true);
        Optional<ButtonType> result = alert.showAndWait();
        return !(result.isPresent() && result.get() == btnNo);
    }

    private void setupSheetColumns(Sheet sheet) {
        // (target + 0.714) * 256 rounded to an integer
        sheet.setColumnWidth(0, 3218);  // A: 11.86
        sheet.setColumnWidth(1, 6144);  // B: 23.29
        sheet.setColumnWidth(2, 8265);  // C: 31.57
        sheet.setColumnWidth(3, 2817);  // D: 10.29
        sheet.setColumnWidth(4, 3152);  // E: 11.6
        sheet.setColumnWidth(5, 3840);  // F: 14.29
        sheet.setColumnWidth(6, 2707);  // G: 9.86
        sheet.setColumnWidth(7, 2925);  // H: 10.71
        sheet.setColumnWidth(8, 2742);  // I: 10
        sheet.setColumnWidth(9, 2669);  // J: 9.71
        sheet.setColumnWidth(10, 4681); // K: 17.57
        sheet.setColumnWidth(11, 2340); // L: 8.43
    }

    private int processRecords(XSSFWorkbook workbook, Sheet sheet) {
        List<ReportRecord> records = loadRecords(dtpckStart.getValue(), dtpckEnd.getValue());
        int rowNum = 19;
        
        for (ReportRecord r : records) {
            if (shouldSkipRecord(r)) {
                continue;
            }
            putReportRecord(workbook, sheet, r, rowNum++);
        }
        return rowNum;
    }

    private boolean shouldSkipRecord(ReportRecord repRec) {
        return repRec.taskName().startsWith("SAP") || repRec.taskName().startsWith("САП");
    }

    private void saveAndOpenWorkbook(XSSFWorkbook workbook) {
        String fileLocation = lblOutputDir.getText() + "\\" + tfOutputFileName.getText();
        try {
            FileOutputStream outputStream = new FileOutputStream(fileLocation);
            workbook.write(outputStream);
            workbook.close();
            Desktop.getDesktop().open(new File(fileLocation));
        } catch (IOException e) {
            Utils.showErrorAndStack(e);
        }
    }

    private List<ReportRecord> loadRecords(LocalDate dtStart, LocalDate dtEnd) {
        List<ReportRecord> result = new ArrayList<>();
        java.sql.Date sqlStart = java.sql.Date.valueOf(dtStart);
        java.sql.Date sqlEnd   = java.sql.Date.valueOf(dtEnd);
        try (PreparedStatement pstmt = connection.prepareStatement(HOURS_SQL)) {
            pstmt.setDate(1, sqlStart);
            pstmt.setDate(2, sqlEnd);
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    result.add(new ReportRecord(
                        rs.getString("usrFIO"),
                        rs.getString("taskName"),
                        rs.getInt("totals"),
                        Utils.toLocalDate(rs.getTimestamp("CreationDate")),
                        rs.getString("extRefNum"),
                        rs.getString("Organization"),
                        rs.getString("reqFIO"),
                        rs.getString("category")
                    ));
                }
            }
        } catch (SQLException e) {
            Utils.showErrorAndStack(e);
        }
        return result;
    }

    private void putReportRecord(XSSFWorkbook workbook, Sheet sheet, ReportRecord line, int rowNum) {
        XSSFFont tblFont = workbook.createFont();
        tblFont.setFontName(REPORT_FONT_NAME);
        tblFont.setFontHeightInPoints((short) 8);
        CellStyle tblStyle = workbook.createCellStyle();
        tblStyle.setFont(tblFont);
        tblStyle.setWrapText(true);
        tblStyle.setAlignment(HorizontalAlignment.CENTER);
        tblStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        tblStyle.setBorderTop(BorderStyle.THIN);
        tblStyle.setBorderBottom(BorderStyle.THIN);
        tblStyle.setBorderLeft(BorderStyle.THIN);
        tblStyle.setBorderRight(BorderStyle.THIN);

        Row row = sheet.createRow(rowNum);
        Cell cell = row.createCell(0);
        cell.setCellValue(line.extRefNum());
        cell.setCellStyle(tblStyle);

        cell = row.createCell(1);
        cell.setCellValue(line.taskName());
        cell.setCellStyle(tblStyle);

        cell = row.createCell(2);
        // TODO: сюда результат работ
        cell.setCellValue("");
        cell.setCellStyle(tblStyle);

        cell = row.createCell(3);
        // "Наименование и код информационного ресурса"
        cell.setCellValue("");
        cell.setCellStyle(tblStyle);

        cell = row.createCell(4);
        if (line.taskName().startsWith("EAM") || line.taskName().startsWith("ЕАМ")) {
            cell.setCellValue("ЛУС \"Галактика EAM\"");
        } else {
            cell.setCellValue("ЛУС \"Галактика ERP\"");
        }
        cell.setCellStyle(tblStyle);

        cell = row.createCell(5);
        cell.setCellValue(line.requesterOrg());
        cell.setCellStyle(tblStyle);

        cell = row.createCell(6);
        cell.setCellValue(line.userCategory());
        cell.setCellStyle(tblStyle);

        cell = row.createCell(7);
        cell.setCellValue(getRate(line));
        cell.setCellStyle(tblStyle);

        cell = row.createCell(8);
        cell.setCellValue(line.totals());
        cell.setCellStyle(tblStyle);

        cell = row.createCell(9);
        cell.setCellValue(line.totals());
        cell.setCellStyle(tblStyle);

        CellStyle tblNumStyle = workbook.createCellStyle();
        tblNumStyle.cloneStyleFrom(tblStyle);
        tblNumStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
        cell = row.createCell(10);
        cell.setCellValue(getRate(line) * line.totals());
        cell.setCellStyle(tblNumStyle);
    }

    /** Fallback chains: if a category is absent from config, try the next key. */
    private static final Map<String, String[]> CATEGORY_FALLBACKS = Map.of(
        "k1", new String[]{"k1", "k2", "k3"},
        "k2", new String[]{"k2", "k3"},
        "k3", new String[]{"k3"},
        "k4", new String[]{"k4"},
        "k5", new String[]{"k5", "k4"}
    );

    /** Returns the config sub-key ("erp" or "eam") for a record. */
    private static String getSystemKey(ReportRecord r) {
        return (r.taskName().startsWith("EAM") || r.taskName().startsWith("ЕАМ"))
                ? "eam" : "erp";
    }

    /**
     * Looks up the hourly rate for a record, applying category fallbacks.
     * DB stores Cyrillic "К4"…"К1"; config uses Latin "k4"…"k1".
     * Returns 0.0 if no matching key is found.
     */
    private double getRate(ReportRecord r) {
        if (r.userCategory() == null) return 0.0;
        String baseKey = r.userCategory().toLowerCase()
                          .replace("к", "k");   // Cyrillic К -> Latin k
        String systemKey = getSystemKey(r);
        String[] fallbacks = CATEGORY_FALLBACKS.getOrDefault(baseKey, new String[]{baseKey});
        for (String key : fallbacks) {
            if (costsConfig.hasPath(systemKey + "." + key)) {
                return costsConfig.getDouble(systemKey + "." + key);
            }
        }
        return 0.0;
    }

    private boolean hasTaskNameChanged(String taskName, SavedTask st) {
        return taskName != null && !taskName.isEmpty() && st.taskName() != null && !st.taskName().equals(taskName);
    }

}
