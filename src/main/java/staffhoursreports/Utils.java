package staffhoursreports;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.TemporalAdjusters;
import java.time.temporal.WeekFields;
import java.util.Date;
import java.util.Locale;
import java.util.Optional;

public class Utils {
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

        // Localize the "Show Details" hyperlink
        alert.getDialogPane().expandedProperty().addListener((observable, wasExpanded, isExpanded) -> {
            Hyperlink detailsLink = (Hyperlink) alert.getDialogPane().lookup(".hyperlink");
            if (detailsLink != null) {
                detailsLink.setText(isExpanded ? "Скрыть детали" : "Показать детали");
            }
        });

        // Set initial hyperlink text before showing
        alert.setOnShown(event -> {
            Hyperlink detailsLink = (Hyperlink) alert.getDialogPane().lookup(".hyperlink");
            if (detailsLink != null && detailsLink.getText().contains("Details")) {
                detailsLink.setText("Показать детали");
            }
        });

        alert.showAndWait();
    }

    public static LocalDate toLocalDate(Date input) {
        return input == null ? null : input.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }
    
    public static Date fromLocalDate(LocalDate input) {
        return input == null ? null : Date.from(input.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

}
