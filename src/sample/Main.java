package sample;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        Parent root = FXMLLoader.load(getClass().getResource("sample.fxml"));
        primaryStage.setTitle("20 CVO");
        Scene scene = new Scene(root, 640, 390);
        primaryStage.setScene(scene);
        primaryStage.setMaxHeight(390);
        primaryStage.setMaxWidth(640);
        primaryStage.setMinHeight(390);
        primaryStage.setMinWidth(640);
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
