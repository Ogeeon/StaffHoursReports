package staffhoursreports;

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
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;
import java.util.Date;
import java.util.List;
import java.util.*;
import java.util.prefs.Preferences;

public class RootPaneView implements Initializable {
    private static final String ZI_SRC_FN_KEY = "zi_src_fn";
    private static final String ZI_DST_DIR_KEY = "zi_dst_dir";
    private static final String HOURS_SQL = buildHoursSql();
    private record Response(
        Integer taskId, String extRefNum, String executor,
        Integer executorId, Integer requesterId,
        String taskName, Date creationDate
    ) {
        @Override
        public String toString() {
            return "[" + taskId() + ", " + extRefNum() + ", " + requesterId() + "]";
        }
    }
    private record ReportRecord(
        String userFIO, String taskName, int totals,
        LocalDate creationDate, String extRefNum,
        String requesterOrg, String requesterName, String userCategory
    ) {}
    Connection connection = null;
    private final FileChooser fileChooser = new FileChooser();
    private final Preferences preferences = Preferences.userNodeForPackage(RootPaneView.class);
    private Map<String, Double> rates;
    double costERP = 0;
    double costEAM = 0;

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
        if (currDate.getDayOfMonth() < 16)
            currDate = currDate.minusMonths(1);
        LocalDate dateTo = LocalDate.of(currDate.getYear(), currDate.getMonth(), 15);
        LocalDate dateFrom = dateTo.minusMonths(1).plusDays(1);
        dtpckStart.setValue(dateFrom);
        dtpckEnd.setValue(dateTo);
        tfOutputFileName.setText(getReportName(dateFrom, dateTo));
        costERP = ConfigFactory.load().getDouble("costs.erp");
        costEAM = ConfigFactory.load().getDouble("costs.eam");
        rates = new HashMap<>();
        rates.put("К4", ConfigFactory.load().getDouble("costs.k4"));
        rates.put("К3", ConfigFactory.load().getDouble("costs.k3"));
        rates.put("К2", ConfigFactory.load().getDouble("costs.k2"));
        rates.put("К1", ConfigFactory.load().getDouble("costs.k1"));
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
        Task<List<String>> task = new Task<List<String>>() {
            @Override
            protected List<String> call() {
                List<String> messages = new ArrayList<>();
                DecimalFormat decimalFormat = (DecimalFormat) NumberFormat.getIntegerInstance();
                decimalFormat.setMinimumFractionDigits(0);
                String taskName;
                String executorFI;
                Date creationDate;
                String extRefNum;
                String organization;
                String requester;
                try {
                    File file = new File(tfInputFileName.getText());
                    FileInputStream fis = new FileInputStream(file);
                    XSSFWorkbook wb = new XSSFWorkbook(fis);
                    XSSFSheet sheet = wb.getSheetAt(0);
                    int totalRows = sheet.getLastRowNum();
                    if (totalRows < 2)
                        return null;
                    else {
                        pbImport.setProgress(0);
                        hbProgress.setVisible(true);
                    }
                    for (Row row : sheet) {
                        try {
                            taskName = getCellAsString(row.getCell(0));
                            if (taskName == null || taskName.equals("Тема")) continue; // Заголовок таблицы пропускаем
                            executorFI = getCellAsString(row.getCell(1));
                            if (executorFI == null || executorFI.isEmpty()) {
                                messages.add("Строка " + (row.getRowNum() + 1) + ": пустое поле 'Исполнитель', пропуск");
                                continue;
                            }
                            creationDate = getCellAsDate(row.getCell(2));
                            if (creationDate == null) {
                                messages.add("Строка " + (row.getRowNum() + 1) + ": пустое или некорректное поле 'Дата', пропуск");
                                continue;
                            }
                            extRefNum = getCellAsString(row.getCell(3));
                            if (extRefNum == null || extRefNum.isEmpty()) {
                                messages.add("Строка " + (row.getRowNum() + 1) + ": пустое поле '№ обращения', пропуск");
                                continue;
                            }
                            requester = getCellAsString(row.getCell(4));
                            if (requester == null || requester.isEmpty()) {
                                messages.add("Строка " + (row.getRowNum() + 1) + ": пустое поле 'Заявитель', пропуск");
                                continue;
                            }
                            organization = getCellAsString(row.getCell(5));
                            if (organization == null || organization.isEmpty()) {
                                messages.add("Строка " + (row.getRowNum() + 1) + ": пустое поле 'Организация', пропуск");
                                continue;
                            }
                            Response r = getTask(extRefNum);
                            if (r != null) {
                                boolean updated = false;
                                StringBuilder updates = new StringBuilder();

                                // Check for task name change
                                if (taskName != null && !taskName.isEmpty() && r.taskName() != null && !r.taskName().equals(taskName)) {
                                    updates.append("название: '").append(r.taskName()).append("' -> '").append(taskName).append("'");
                                    updated = true;
                                }

                                // Check for creation date change
                                if (r.creationDate() != null && !r.creationDate().equals(creationDate)) {
                                    if (updated) updates.append(", ");
                                    String oldDate = Utils.toLocalDate(r.creationDate()).format(DateTimeFormatter.ofLocalizedDate(FormatStyle.MEDIUM));
                                    String newDate = Utils.toLocalDate(creationDate).format(DateTimeFormatter.ofLocalizedDate(FormatStyle.MEDIUM));
                                    updates.append("дата: ").append(oldDate).append(" -> ").append(newDate);
                                    updated = true;
                                }

                                // Check for executor change
                                if (!r.executor().equals(executorFI)) {
                                    int executorId = getExecutorID(executorFI, messages);
                                    if (executorId == 0) {
                                        messages.add("Строка " + (row.getRowNum() + 1) + ": не удалось определить исполнителя: " + executorFI + ", пропуск записи");
                                        continue;
                                    }
                                    try (PreparedStatement pstmt = connection.prepareStatement("UPDATE Tasks SET executor_id=? WHERE id=?")) {
                                        pstmt.setInt(1, executorId);
                                        pstmt.setInt(2, r.taskId());
                                        int ur = pstmt.executeUpdate();
                                        if (ur == 1) {
                                            if (updated) updates.append(", ");
                                            updates.append("исполнитель: ").append(executorFI);
                                            updated = true;
                                        }
                                    }
                                }

                                // Check for requester
                                if (r.requesterId() == 0) {
                                    int rid = getRequesterID(requester, organization, messages);
                                    if (rid == 0) {
                                        if (updated) {
                                            messages.add("Обновлена задача " + extRefNum + ": " + updates.toString());
                                        }
                                        updateProgress(row.getRowNum(), totalRows);
                                        continue;
                                    }
                                    try (PreparedStatement pstmt = connection.prepareStatement("UPDATE Tasks SET requester_id=? WHERE id=?")) {
                                        pstmt.setInt(1, rid);
                                        pstmt.setInt(2, r.taskId());
                                        int ur = pstmt.executeUpdate();
                                        if (ur == 1) {
                                            if (updated) updates.append(", ");
                                            updates.append("id заявителя: ").append(rid);
                                            updated = true;
                                        }
                                    }
                                }

                                // Update task name and/or date if they changed
                                if (taskName != null && !taskName.isEmpty() && r.taskName() != null && !r.taskName().equals(taskName)) {
                                    try (PreparedStatement pstmt = connection.prepareStatement("UPDATE Tasks SET taskName=? WHERE id=?")) {
                                        pstmt.setString(1, taskName);
                                        pstmt.setInt(2, r.taskId());
                                        pstmt.executeUpdate();
                                    }
                                }
                                if (r.creationDate() != null && !r.creationDate().equals(creationDate)) {
                                    try (PreparedStatement pstmt = connection.prepareStatement("UPDATE Tasks SET creationDate=? WHERE id=?")) {
                                        pstmt.setDate(1, new java.sql.Date(creationDate.getTime()));
                                        pstmt.setInt(2, r.taskId());
                                        pstmt.executeUpdate();
                                    }
                                }

                                if (updated) {
                                    messages.add("Обновлена задача " + extRefNum + ": " + updates.toString());
                                }
                            } else {
                                int rid = getRequesterID(requester, organization, messages);
                                if (rid == 0) {
                                    updateProgress(row.getRowNum(), totalRows);
                                    continue;
                                }
                                insertTask(taskName, executorFI, creationDate, extRefNum, rid, messages);
                            }
                            updateProgress(row.getRowNum(), totalRows);
                        } catch (Exception e) {
                            messages.add("Ошибка при обработке строки " + (row.getRowNum() + 1) + ": " + e.getMessage());
                        }
                    }
                }
                catch(Exception e) {
                    messages.add("Критическая ошибка: " + e.getMessage());
                    throw new RuntimeException(e);
                }
                return messages;
            }
        };

        task.progressProperty().addListener((obs, oldProgress, newProgress) -> pbImport.setProgress(newProgress.doubleValue()));
        task.setOnSucceeded(e -> {
            hbProgress.setVisible(false);
            List<String> messages = task.getValue();
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
        });
        task.setOnFailed(e -> {
            hbProgress.setVisible(false);
            Throwable exception = task.getException();
            if (exception != null) {
                Utils.showErrorAndStack((Exception) exception);
            } else {
                Utils.showError("Произошла неизвестная ошибка при обработке файла.");
            }
        });
        Thread th = new Thread(task);
        th.start();
    }

    private Response getTask(String extRefNum) throws SQLException {
        Response r = null;
        try (PreparedStatement pstmt = connection.prepareStatement(
                "SELECT t.id, t.executor_id, t.requester_id, u.fio, t.taskName, t.creationDate " +
                "FROM Tasks t " +
                "INNER JOIN Users u ON t.executor_id = u.id " +
                "WHERE extRefNum = ?")) {
            pstmt.setString(1, extRefNum);
            try (ResultSet rs = pstmt.executeQuery()) {
                while (rs.next()) {
                    r = new Response(rs.getInt("id"), extRefNum, rs.getString("fio"), rs.getInt("executor_id"),
                                    rs.getInt("requester_id"), rs.getString("taskName"), rs.getDate("creationDate"));
                }
            }
        }
        return r;
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
        if (dstFile.exists()) {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.setTitle("Подтверждение");
            alert.setHeaderText("Такой файл уже существует. Перезаписать его?");
            ButtonType btnYes = new ButtonType("Да");
            ButtonType btnNo = new ButtonType("Нет");
            alert.getButtonTypes().setAll(btnYes, btnNo);
            ((Button) alert.getDialogPane().lookupButton(btnYes)).setDefaultButton(false);
            ((Button) alert.getDialogPane().lookupButton(btnNo)).setDefaultButton(true);
            Optional<ButtonType> result = alert.showAndWait();
            if (result.isPresent() && result.get() == btnNo)
                return;
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Отчёт по заявкам");
        sheet.setColumnWidth(0, 6600);
        sheet.setColumnWidth(1, 3800);
        sheet.setColumnWidth(2, 10560);
        sheet.setColumnWidth(3, 3800);
        sheet.setColumnWidth(4, 3800);
        sheet.setColumnWidth(5, 3800);
        sheet.setColumnWidth(6, 11880);
        sheet.setColumnWidth(7, 9240);
        putCaption(workbook, sheet);
        List<ReportRecord> records = loadRecords(dtpckStart.getValue(), dtpckEnd.getValue(), false);
        int rowNum = 2;
        int hoursERP = 0;
        int hoursEAM = 0;
        for (ReportRecord r: records) {
            if (r.taskName().startsWith("SAP") || r.taskName().startsWith("САП"))
                continue;
            if (r.taskName().startsWith("EAM") || r.taskName().startsWith("ЕАМ"))
                hoursEAM += r.totals();
            else
                hoursERP += r.totals();
            putReportRecord(workbook, sheet, r, rowNum++);
        }
        rowNum += 2;
        putKNZPTotals(workbook, sheet, hoursERP, hoursEAM, costERP, costEAM, rowNum++);

        records = loadRecords(dtpckStart.getValue(), dtpckEnd.getValue(), true);
        if (!records.isEmpty()) {
            rowNum += 4;
            Map<String, Integer> hrsByCat = new HashMap<>();
            for (ReportRecord r: records) {
                if (hrsByCat.containsKey(r.userCategory())) {
                    hrsByCat.put(r.userCategory(), hrsByCat.get(r.userCategory()) + r.totals());
                } else {
                    hrsByCat.put(r.userCategory(), r.totals());
                }
            }
            rowNum = putRNTTotals(workbook, sheet, hrsByCat, rowNum) + 2;
            for (ReportRecord r: records) {
                putReportRecord(workbook, sheet, r, rowNum++);
            }
        }
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

    private void putCaption(XSSFWorkbook workbook, Sheet sheet) {
        Row caption = sheet.createRow(0);
        CellStyle captionStyle = workbook.createCellStyle();

        XSSFFont captionFont = workbook.createFont();
        captionFont.setFontHeightInPoints((short) 19);
        captionFont.setBold(true);
        captionStyle.setFont(captionFont);

        Cell headerCell = caption.createCell(0);
        headerCell.setCellValue(tfOutputFileName.getText().substring(0, tfOutputFileName.getText().length() - 5));
        headerCell.setCellStyle(captionStyle);

        CellStyle headerStyle = workbook.createCellStyle();

        XSSFFont headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 11);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        Row header = sheet.createRow(1);
        Cell cell = header.createCell(0);
        cell.setCellValue("Исполнитель");
        cell.setCellStyle(headerStyle);

        cell = header.createCell(1);
        cell.setCellValue("№ обращения");
        cell.setCellStyle(headerStyle);

        cell = header.createCell(2);
        cell.setCellValue("Наименование");
        cell.setCellStyle(headerStyle);

        cell = header.createCell(3);
        cell.setCellValue("Трудозатраты");
        cell.setCellStyle(headerStyle);

        cell = header.createCell(4);
        cell.setCellValue("Дата создания");
        cell.setCellStyle(headerStyle);

        cell = header.createCell(5);
        cell.setCellValue("Организация");
        cell.setCellStyle(headerStyle);

        cell = header.createCell(6);
        cell.setCellValue("Пользователь");
        cell.setCellStyle(headerStyle);
    }

    private List<ReportRecord> loadRecords(LocalDate dtStart, LocalDate dtEnd, boolean showRNT) {
        List<ReportRecord> result = new ArrayList<>();
        String sql = String.format(HOURS_SQL, showRNT ? "" : "NOT ");
        java.sql.Date sqlStart = java.sql.Date.valueOf(dtStart);
        java.sql.Date sqlEnd   = java.sql.Date.valueOf(dtEnd);
        try (PreparedStatement pstmt = connection.prepareStatement(sql)) {
            for (int i = 1; i <= 7; i++) {
                pstmt.setDate(2 * i - 1, sqlStart);
                pstmt.setDate(2 * i,     sqlEnd);
            }
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

    private static String buildHoursSql() {
        StringBuilder sb = new StringBuilder();
        sb.append("select Users.FIO usrFIO, Tasks.taskName, hrt.totals, Tasks.CreationDate, Tasks.extRefNum")
          .append(", Requesters.Organization, Requesters.FIO reqFIO, Users.category ")
          .append("from (")
          .append("select User_id, Task_id, sum(sumhr) totals ")
          .append("from (");
        for (int i = 1; i <= 7; i++) {
            String h = "hr" + i;
            String d = "dat" + i;
            sb.append("select User_id, Task_id, sum(").append(h).append(") sumhr ")
              .append("from TSRecordViews where ")
              .append("? <= ").append(d)
              .append(" AND ").append(d).append(" <= ? ")
              .append("AND ").append(h).append(" > 0 ")
              .append("group by User_id, Task_id");
            if (i < 7)
                sb.append(" union all\n");
        }
        sb.append(") hrs ")
          .append("group by hrs.User_id, hrs.Task_id")
          .append(") hrt ")
          .append("left join Tasks on (Tasks.id=hrt.Task_id) ")
          .append("left join Users on (Users.id=hrt.User_id) ")
          .append("left join Requesters on (Requesters.id = tasks.requester_id) ")
          .append("WHERE Tasks.taskName %sLIKE 'РН-Транс%%' ")
          .append("ORDER BY 1, 2");
        return sb.toString();
    }

    private void putReportRecord(XSSFWorkbook workbook, Sheet sheet, ReportRecord line, int rowNum) {
        CreationHelper createHelper = workbook.getCreationHelper();
        Row row = sheet.createRow(rowNum);
        Cell cell = row.createCell(0);
        cell.setCellValue(line.userFIO());

        cell = row.createCell(1);
        cell.setCellValue(line.extRefNum());

        cell = row.createCell(2);
        cell.setCellValue(line.taskName());

        cell = row.createCell(3);
        cell.setCellValue(line.totals());

        cell = row.createCell(4);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
        cell.setCellStyle(cellStyle);
        cell.setCellValue(Utils.fromLocalDate(line.creationDate()));

        cell = row.createCell(5);
        cell.setCellValue(line.requesterOrg());

        cell = row.createCell(6);
        cell.setCellValue(line.requesterName());
    }

    private void putKNZPTotals(XSSFWorkbook workbook, Sheet sheet, int totalERP, int totalEAM, double costERP, double costEAM, int rowNum) {
        CreationHelper createHelper = workbook.getCreationHelper();
        XSSFRow xssfRowrow = ((XSSFSheet) sheet).createRow(rowNum);
        XSSFCell xssfCellcell = xssfRowrow.createCell(0);
        XSSFRichTextString rt = new XSSFRichTextString("Отчет по заявкам КНПЗ во вложении.");
        XSSFFont font1 = workbook.createFont();
        font1.setBold(true);
        rt.applyFont(17, 21, font1);
        xssfCellcell.setCellValue(rt);

        Row row = sheet.createRow(++rowNum);
        Cell cell = row.createCell(0);
        cell.setCellValue("ERP");
        cell = row.createCell(1);
        cell.setCellValue(costERP);
        cell = row.createCell(2);
        cell.setCellValue(totalERP);
        cell = row.createCell(3);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("# ##0.00"));
        cell.setCellStyle(cellStyle);
        String frm = "B%d*C%d";
        cell.setCellFormula(String.format(frm, rowNum+1, rowNum+1));

        row = sheet.createRow(++rowNum);
        cell = row.createCell(0);
        cell.setCellValue("EAM");
        cell = row.createCell(1);
        cell.setCellValue(costEAM);
        cell = row.createCell(2);
        cell.setCellValue(totalEAM);
        cell = row.createCell(3);
        cell.setCellStyle(cellStyle);
        cell.setCellFormula(String.format(frm, rowNum+1, rowNum+1));

        row = sheet.createRow(++rowNum);
        cell = row.createCell(3);
        cell.setCellStyle(cellStyle);
        String frm2 = "D%d+D%d";
        cell.setCellFormula(String.format(frm2, rowNum-1, rowNum));
    }

    private int putRNTTotals(XSSFWorkbook workbook, Sheet sheet, Map<String, Integer> hrsByCat, int rowNum) {
        CreationHelper createHelper = workbook.getCreationHelper();
        XSSFRow xssfRowrow = ((XSSFSheet) sheet).createRow(rowNum);
        XSSFCell xssfCellcell = xssfRowrow.createCell(0);
        XSSFRichTextString rt = new XSSFRichTextString("Было выполнено дополнительных работ по РН-Транс");
        XSSFFont font1 = workbook.createFont();
        font1.setBold(true);
        rt.applyFont(39, 47, font1);
        xssfCellcell.setCellValue(rt);

        Row row;
        Cell cell;
        Set<String> keys = hrsByCat.keySet();
        int startRN = rowNum;
        for (String key: keys) {
            row = sheet.createRow(++rowNum);
            cell = row.createCell(0);
            cell.setCellValue(key);
            cell = row.createCell(1);
            cell.setCellValue(rates.get(key));
            cell = row.createCell(2);
            cell.setCellValue(hrsByCat.get(key));
            cell = row.createCell(3);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("# ##0.00"));
            cell.setCellStyle(cellStyle);
            String frm = "B%d*C%d";
            cell.setCellFormula(String.format(frm, rowNum+1, rowNum+1));
        }
        row = sheet.createRow(++rowNum);
        cell = row.createCell(3);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("# ##0.00"));
        cell.setCellStyle(cellStyle);
        String frm = "SUM(D%d:D%d)";
        cell.setCellFormula(String.format(frm, startRN + 2, rowNum));
        return rowNum;
    }
}
