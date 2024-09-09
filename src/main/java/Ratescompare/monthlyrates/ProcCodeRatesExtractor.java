package Ratescompare.monthlyrates;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ProcCodeRatesUI extends Application {

    private ComboBox<String> columnComboBox;

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Proc Code Rates Comparison");

        // GridPane layout for form elements
        GridPane grid = new GridPane();
        grid.setPadding(new Insets(15));
        grid.setHgap(10);
        grid.setVgap(10);

        // Label for selecting folder
        Label folderLabel = new Label("Select Folder:");
        grid.add(folderLabel, 0, 0);

        // Button to select folder
        Button folderButton = new Button("Browse...");
        grid.add(folderButton, 1, 0);

        // Label for selecting column
        Label columnLabel = new Label("Select Column Header:");
        grid.add(columnLabel, 0, 1);

        // ComboBox for column headers
        columnComboBox = new ComboBox<>();
        columnComboBox.setEditable(true); // Allow typing and auto-completion
        grid.add(columnComboBox, 1, 1);

        // Button to start comparison
        Button compareButton = new Button("Compare Rates");
        grid.add(compareButton, 1, 2);

        // Center the button and ComboBox
        GridPane.setMargin(folderButton, new Insets(10, 0, 10, 0));
        GridPane.setMargin(compareButton, new Insets(10, 0, 10, 0));

        // FileChooser to select folder
        folderButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Select Excel File");
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                List<String> headers = getColumnHeaders(selectedFile);
                columnComboBox.getItems().setAll(headers); // Load headers into ComboBox
            }
        });

        // Handle the comparison button action
        compareButton.setOnAction(e -> {
            String selectedColumn = columnComboBox.getValue();
            if (selectedColumn != null && !selectedColumn.isEmpty()) {
                // Perform rate comparison using the selected column header
                System.out.println("Selected Column: " + selectedColumn);
                // Add your comparison logic here...
            } else {
                Alert alert = new Alert(Alert.AlertType.ERROR, "Please select a valid column header.");
                alert.showAndWait();
            }
        });

        // Set up the scene and stage
        Scene scene = new Scene(grid, 400, 200);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    /**
     * Reads the column headers from the first row of the Excel file.
     * @param file the Excel file to read.
     * @return a List of column headers.
     */
    private List<String> getColumnHeaders(File file) {
        List<String> headers = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the first sheet
            Row headerRow = sheet.getRow(0); // Assuming first row is the header

            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if (cell.getCellType() == CellType.STRING) {
                        headers.add(cell.getStringCellValue().trim());
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return headers;
    }
}
