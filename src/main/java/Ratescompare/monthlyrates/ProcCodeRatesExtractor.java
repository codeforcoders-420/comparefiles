package Ratescompare.monthlyrates;

import javafx.application.Application;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Hello world!
 */
public class ProcCodeRatesExtractor extends Application {

    private Map<String, List<Double>> procModRates = new LinkedHashMap<>();

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel Rates Comparator");

        DirectoryChooser directoryChooser = new DirectoryChooser();
        Button selectFolderButton = new Button("Select Folder Containing Excel Files");

        selectFolderButton.setOnAction(e -> {
            File selectedDirectory = directoryChooser.showDialog(primaryStage);
            if (selectedDirectory != null) {
                processExcelFiles(selectedDirectory);
            }
        });

        VBox layout = new VBox(20, selectFolderButton);
        Scene scene = new Scene(layout, 400, 200);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void processExcelFiles(File folder) {
        File[] excelFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));

        if (excelFiles != null) {
            Arrays.sort(excelFiles, Comparator.comparing(File::getName)); // Sort files by name (month order)

            for (File file : excelFiles) {
                readExcelFile(file);
            }

            // Write the extracted data to the output Excel file
            String outputFilePath = "Output_Rates.xlsx";
            writeOutputExcelFile(outputFilePath, procModRates);
            System.out.println("Output file created successfully: " + outputFilePath);
        } else {
            System.out.println("No Excel files found in the specified folder.");
        }
    }

    private void readExcelFile(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) { // Skip header row
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell procCodeCell = row.getCell(0); // Proc Code in Column A
                    Cell modCell = row.getCell(1); // Mod in Column B
                    Cell mod2Cell = row.getCell(2); // Mod 2 in Column C
                    Cell rateCell = row.getCell(6); // Rate from Column D

                    if (procCodeCell != null && modCell != null && mod2Cell != null && rateCell != null) {
                        String key = procCodeCell.getStringCellValue() + "-" + modCell.getStringCellValue() + "-" + mod2Cell.getStringCellValue();
                        double rate = rateCell.getNumericCellValue();

                        procModRates.computeIfAbsent(key, k -> new ArrayList<>(Collections.nCopies(12, 0.0)));
                        procModRates.get(key).set(getMonthIndex(file.getName()), rate);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private int getMonthIndex(String fileName) {
        // Get month index based on file name (0 for Jan, 1 for Feb, etc.)
        List<String> months = Arrays.asList("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
        for (int i = 0; i < months.size(); i++) {
            if (fileName.toLowerCase().contains(months.get(i).toLowerCase())) {
                return i;
            }
        }
        return 0; // Default to January if not found
    }

    private void writeOutputExcelFile(String outputFilePath, Map<String, List<Double>> procModRates) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Rates");

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Proc-Mod-Mod2");
            for (int i = 0; i < 12; i++) {
                headerRow.createCell(i + 1).setCellValue(getMonthName(i));
            }

            // Write data rows
            int rowIndex = 1;
            for (Map.Entry<String, List<Double>> entry : procModRates.entrySet()) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(entry.getKey());
                for (int i = 0; i < entry.getValue().size(); i++) {
                    row.createCell(i + 1).setCellValue(entry.getValue().get(i));
                }
            }

            // Write to the output file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String getMonthName(int monthIndex) {
        // Get month name based on index (0 for Jan, 1 for Feb, etc.)
        String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
        return months[monthIndex];
    }
}