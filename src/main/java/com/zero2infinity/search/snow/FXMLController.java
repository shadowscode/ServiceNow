package com.zero2infinity.search.snow;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Iterator;
import java.util.ResourceBundle;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.concurrent.Service;
import javafx.concurrent.Task;
import javafx.concurrent.WorkerStateEvent;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.TextField;
import javafx.scene.control.ToggleGroup;
import javafx.scene.input.Clipboard;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Duration;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.controlsfx.control.Notifications;

public class FXMLController implements Initializable {

    String home = System.getProperty("user.home");
    //String XLSX_FILE_PATH = home + "/Desktop/SNOW/Matrix.xlsx";
    String XLSX_FILE_PATH = home + "\\Desktop\\SNOW\\Matrix.xlsx";

    @FXML
    Button btn_search;

    @FXML
    Button btn_filepicker;

    @FXML
    TextField txtfld_appname;

    @FXML
    TextField txtfld_filepath;

    @FXML
    Label label_errorbox;

    @FXML
    RadioButton rbtnonshore;

    @FXML
    RadioButton rbtnoffshore;

    @FXML
    CheckBox cbclipboard;

    ToggleGroup rbgroup;

    Service<Void> service;
    String sample, inputString;
    String Group = null;
    String ApplicationName = null;
    String PocName = null;
    String PocContact = null;
    int poccol;
    int poccontactcol;
    Notifications notifications;

    com.sun.glass.ui.ClipboardAssistance clipboardAssistance;

    @FXML
    private void handleButtonAction(ActionEvent event) throws IOException {
        if (event.getSource() == btn_search) {
            prepareSearch();
        } else if (event.getSource() == btn_filepicker) {
            FileChooser fileChooser = new FileChooser();
            File selectedFile = fileChooser.showOpenDialog(null);
            if (selectedFile != null) {
                txtfld_filepath.setText(selectedFile.getAbsolutePath());
                XLSX_FILE_PATH = selectedFile.getAbsolutePath();
            }
        }
    }

    public void prepareSearch() {
        inputString = txtfld_appname.getText().trim();
        if (rbgroup.getSelectedToggle().equals(this.rbtnoffshore)) {
            poccol = 1;
            poccontactcol = 2;
        } else if (rbgroup.getSelectedToggle().equals(this.rbtnonshore)) {
            poccol = 3;
            poccontactcol = 4;
        }

        if (inputString.isEmpty()) {
            label_errorbox.setText("Enter Text");
        } else {
            service = new Service<Void>() {
                @Override
                protected Task<Void> createTask() {
                    return new Task<Void>() {
                        @Override
                        protected Void call() throws Exception {
                            readExcel(inputString);
                            return null;
                        }
                    };
                }
            };
            service.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
                @Override
                public void handle(WorkerStateEvent event) {
                    final Stage stage;
                    if (PocName == null) {
                        label_errorbox.setText("Result Not Found Try Again !");
                    } else if (PocName != null) {
                        notifications = Notifications.create()
                                .title(Group)
                                .text("Group : " + '\t' + Group + '\n'
                                        + "Application Name : " + '\t' + ApplicationName + '\n'
                                        + "POC Name : " + '\t' + PocName + '\n'
                                        + "POC Contact : " + '\t' + PocContact + '\n')
                                .darkStyle()
                                .hideAfter(Duration.seconds(20))
                                .onAction(new EventHandler<ActionEvent>() {
                                    @Override
                                    public void handle(ActionEvent event) {
                                        notifications.hideAfter(Duration.ZERO);
                                    }
                                });
                        notifications.show();
                        PocName = null;
                        label_errorbox.setText("Log");
                    }
                }
            });
            service.restart();
        }
    }

    public void readExcel(String Appname) throws IOException, InvalidFormatException {

        final String searchappname = Appname;
        Workbook workbook;
        workbook = WorkbookFactory.create(new File(XLSX_FILE_PATH));
        int numberofsheets = workbook.getNumberOfSheets();
        for (int j = 0; j < numberofsheets; j++) {
            Sheet sheet = workbook.getSheetAt(j);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0);
                if (searchappname.toLowerCase().equals(cell.getStringCellValue().toLowerCase())) {
                    //if (Pattern.compile(Pattern.quote(cell.getStringCellValue()), Pattern.CASE_INSENSITIVE).matcher(searchappname).find()) {
                    this.Group = workbook.getSheetName(j);
                    this.ApplicationName = cell.getStringCellValue();
                    Cell cell1 = row.getCell(poccol);
                    this.PocName = cell1.getStringCellValue();
                    Cell cell2 = row.getCell(poccontactcol);
                    this.PocContact = cell2.getStringCellValue();
                    break;
                }
            }
        }
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // TODO
        Clipboard copy = Clipboard.getSystemClipboard();

        txtfld_filepath.setText(XLSX_FILE_PATH);
        txtfld_appname.setText("");
        cbclipboard.setSelected(false);
        txtfld_appname.setOnKeyPressed(new EventHandler<KeyEvent>() {
            @Override
            public void handle(KeyEvent event) {
                if (event.getCode() == KeyCode.ENTER) {
                    prepareSearch();
                }
            }
        });

        rbgroup = new ToggleGroup();
        rbtnoffshore.setToggleGroup(rbgroup);

        String getrbv = "off";

        if (getrbv.equals("off")) {
            rbtnoffshore.setSelected(true);
        } else if (getrbv.equals("on")) {
            rbtnonshore.setSelected(true);
        }
        rbtnonshore.setToggleGroup(rbgroup);

        cbclipboard.selectedProperty().addListener(new ChangeListener<Boolean>() {
            @Override
            public void changed(ObservableValue<? extends Boolean> observable, Boolean oldValue, Boolean newValue) {
                if (cbclipboard.isSelected()) {
                    txtfld_appname.setText("");
                    txtfld_appname.setText(Clipboard.getSystemClipboard().getString().trim());
                    clipboardAssistance = new com.sun.glass.ui.ClipboardAssistance(com.sun.glass.ui.Clipboard.SYSTEM) {
                        @Override
                        public void contentChanged() {
                            txtfld_appname.setText("");
                            if (Clipboard.getSystemClipboard().hasString()) {
                                txtfld_appname.setText(Clipboard.getSystemClipboard().getString().trim());
                                prepareSearch();
                            }
                        }
                    };
                } else if (!cbclipboard.isSelected()) {
                    txtfld_appname.setText("");
                    clipboardAssistance.close();
                }
            }
        });
    }
}
