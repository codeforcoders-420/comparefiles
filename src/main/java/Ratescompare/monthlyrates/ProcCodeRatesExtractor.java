package Ratescompare.monthlyrates;

import javafx.application.Application;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.control.Label;
import javafx.scene.layout.VBox;
import javafx.scene.layout.HBox;
import javafx.geometry.Pos;
import javafx.geometry.Insets;

import java.io.File;

public class ProcCodeRatesExtractor extends Application {

    private File selectedDirectory;
    private TextField columnInputField;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Select Excel Files Directory and Column");

        // Label for column input
        Label columnLabel = new Label("Enter Column Name (e.g., A, B, C):");

        // Text field to input the column name
        columnInputField = new TextField();
        columnInputField.setPromptText("Enter Column Name");

        // Button to open directory chooser
        Button selectFolderButton = new Button("Select Folder Containing Excel Files");
        selectFolderButton.setOnAction(event -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            directoryChooser.setTitle("Select Folder");
            File directory = directoryChooser.showDialog(primaryStage);

            if (directory != null) {
                selectedDirectory = directory;
                System.out.println("Selected Directory: " + selectedDirectory.getAbsolutePath());
                
                // Retrieve column name from text field
                String columnName = columnInputField.getText().trim().toUpperCase();
                if (!columnName.isEmpty() && isValidColumnName(columnName)) {
                    // Call the processing method with user input
                    processExcelFiles(selectedDirectory, columnName);
                } else {
                    System.out.println("Invalid Column Name! Please enter a valid column name (A-Z).");
                }
            }
        });

        // Set button and text field alignment
        HBox inputBox = new HBox(10, columnLabel, columnInputField);  // Label and TextField in HBox
        inputBox.setAlignment(Pos.CENTER);

        HBox buttonBox = new HBox(selectFolderButton);
        buttonBox.setAlignment(Pos.CENTER);

        // VBox for layout
        VBox vbox = new VBox(20, inputBox, buttonBox);
        vbox.setAlignment(Pos.CENTER);
        vbox.setPadding(new Insets(50, 50, 50, 50)); // Padding around the VBox

        // Scene
        Scene scene = new Scene(vbox, 500, 300);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private boolean isValidColumnName(String columnName) {
        // Check if the entered column name is valid (A-Z)
        return columnName.matches("[A-Z]");
    }

    private void processExcelFiles(File folder, String columnName) {
        // Get list of files in the selected folder
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx") || name.endsWith(".xls"));
        if (files != null && files.length > 0) {
            // Processing logic here (use your existing logic to process Excel files)
            System.out.println("Processing files in folder: " + folder.getAbsolutePath() + " for Column: " + columnName);

            // Define output folder (example: creating an "output" folder inside the selected directory)
            File outputFolder = new File(folder, "output");
            if (!outputFolder.exists()) {
                outputFolder.mkdir();  // Create output folder if it does not exist
            }
            System.out.println("Output files will be saved to: " + outputFolder.getAbsolutePath());
            
            // Call your processing function and save results in the output folder, using the columnName for comparison
        } else {
            System.out.println("No Excel files found in the selected folder.");
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
