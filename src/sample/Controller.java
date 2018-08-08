package sample;

import com.google.common.base.Strings;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.io.File;
import java.io.IOException;
import java.util.Optional;

public class Controller {

    FileManager fileManager = new FileManager();

    public Controller() { }

    @FXML
    private void initialize(){}

    @FXML
    private TextField previousPath;

    @FXML
    private TextField currentPath;

    @FXML
    private TextField resultPath;

    @FXML
    private void selectPrevious() {
        Optional<File> file = getFile();
        if(file.isPresent())
            previousPath.setText(file.get().getAbsolutePath());
    }

    @FXML
    private void selectCurrent() {
        Optional<File> file = getFile();
        if(file.isPresent() && fileManager.isNewPeriodValidator(file.get().getAbsolutePath()))
            currentPath.setText(file.get().getAbsolutePath());
        else
            showAlertEx("Укажите файл за прошедший месяц !!!");
    }

    @FXML
    private void selectResultPath() {
        Optional<File> file = getDirectory();
        if(file.isPresent())
            resultPath.setText(file.get().getAbsolutePath());
    }

    @FXML
    private void save() {
        try {
             fileManager.combineFile(Optional.ofNullable(Strings.emptyToNull(previousPath.getText())), Optional.ofNullable(Strings.emptyToNull(currentPath.getText())), Optional.ofNullable(Strings.emptyToNull(resultPath.getText())));
            showAlert("Файл сохранен", "Файл сохранен по пути: " + resultPath.getText() + "\\" +fileManager.resultFile, Alert.AlertType.INFORMATION);
        } catch (IOException e) {
            showAlertEx("Ошибка чтения/записи файлов, попробуйте закрыть все открытые копии форм 20-CVO !!!");
        } catch (FileManagerException e) {
            showAlertEx(e.getMessage());
        } catch (InvalidFormatException e) {
            showAlertEx("Неизвестный формат данных в форме 20-CVO, необходимы изменения данной программы");
        }
    }

    private void showAlertEx(String msg) {
        showAlert("Что-то пошло не так :(", msg, Alert.AlertType.WARNING);
    }


    private void showAlert(String title, String msg, Alert.AlertType type) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setContentText(msg);
        alert.showAndWait();
    }

    private Optional<File> getFile() {
        FileChooser chooser = new FileChooser();
        chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls"));
        chooser.setTitle("Открыть файл");
        return Optional.ofNullable(chooser.showOpenDialog(new Stage()));
    }

    private Optional<File> getDirectory() {
        DirectoryChooser chooser = new DirectoryChooser();
        chooser.setTitle("Выбрать папку");
        return Optional.ofNullable(chooser.showDialog(new Stage()));
    }
}
