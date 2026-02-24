package staffhoursreports;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;
import java.util.Date;
import java.util.Locale;

public class Utils {

    private Utils() {
        // Utility class - prevent instantiation
    }

    public static void showErrorAndStack(Exception e) {
        Alert alert = new Alert(AlertType.ERROR);
        alert.setTitle("Error");
        alert.setHeaderText("Exception occurred.");
        alert.setContentText(e.getMessage());
            
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        e.printStackTrace(pw);
        String exceptionText = sw.toString();

        Label label = new Label("The exception stacktrace was:");

        TextArea textArea = new TextArea(exceptionText);
        textArea.setEditable(false);
        textArea.setWrapText(true);

        textArea.setMaxWidth(Double.MAX_VALUE);
        textArea.setMaxHeight(Double.MAX_VALUE);
        GridPane.setVgrow(textArea, Priority.ALWAYS);
        GridPane.setHgrow(textArea, Priority.ALWAYS);

        GridPane expContent = new GridPane();
        expContent.setMaxWidth(Double.MAX_VALUE);
        expContent.add(label, 0, 0);
        expContent.add(textArea, 0, 1);

        alert.getDialogPane().setExpandableContent(expContent);

        alert.showAndWait();
    }
    
    public static void showError(String errorText) {
        Alert alert = new Alert(AlertType.ERROR);
        alert.setTitle("Ошибка");
        alert.setHeaderText(errorText);
        alert.showAndWait();
    }

    public static void showInfo(String s) {
        Alert alert = new Alert(AlertType.INFORMATION);
        alert.setTitle("Информация");
        alert.setHeaderText("Обработка файла завершена.");

        TextArea textArea = new TextArea(s);
        textArea.setEditable(false);
        textArea.setWrapText(false);

        textArea.setMaxWidth(Double.MAX_VALUE);
        textArea.setMaxHeight(Double.MAX_VALUE);
        GridPane.setVgrow(textArea, Priority.ALWAYS);
        GridPane.setHgrow(textArea, Priority.ALWAYS);

        GridPane expContent = new GridPane();
        expContent.setMaxWidth(Double.MAX_VALUE);
        expContent.add(textArea, 0, 1);

        alert.getDialogPane().setExpandableContent(expContent);
        alert.getDialogPane().setPrefWidth(800);
        alert.getDialogPane().setMinWidth(600);

        alert.showAndWait();
    }

    public static LocalDate toLocalDate(Date input) {
        if (input == null) return null;
        if (input instanceof java.sql.Date sqlDate) return sqlDate.toLocalDate();
        return input.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }
    
    public static Date fromLocalDate(LocalDate input) {
        return input == null ? null : Date.from(input.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    public static String localizeDate(LocalDate date, Locale locale) {
        DateTimeFormatter formatter = DateTimeFormatter.ofLocalizedDate(FormatStyle.LONG).withLocale(locale);
        return formatter.format(date);
    }

    public static String getReportName(LocalDate dt1, LocalDate dt2) {
        return "Отчет по ЗИ за период с " + localizeDate(dt1, Locale.forLanguageTag("ru")) +
                " по " +
                localizeDate(dt2, Locale.forLanguageTag("ru")) +
                ".xlsx";
    }

}
