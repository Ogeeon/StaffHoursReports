package staffhoursreports;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.FileInputStream;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        FXMLLoader loader = new FXMLLoader(getClass().getResource("/RootPane.fxml"));
        Scene scene = new Scene(loader.load());
//        scene.getStylesheets().add("styles.css");
        try {
            primaryStage.getIcons().add(new Image(getClass().getResource("/StaffHoursReports.png").toExternalForm()));
        } catch (Exception e) {
            System.out.println("Failed to load application icon: " + e.getMessage());
        }
        primaryStage.setScene(scene);
        Platform.runLater(() -> primaryStage.setTitle("Админка для трудозатрат"));
        primaryStage.show();
        if (!((RootPaneView) loader.getController()).connect())
            Platform.exit();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
