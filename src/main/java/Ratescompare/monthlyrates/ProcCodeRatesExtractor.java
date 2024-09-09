package Ratescompare.monthlyrates;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.GridPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ProcCodeRatesUI extends Application {

    private ComboBox<String> columnComboBox;
    private File selectedFolder;

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

        // DirectoryChooser to select folder
        folderButton.setOnAction(e -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            directoryChooser.setTitle("Select Folder Containing Excel Files");
            selectedFolder = directoryChooser.showDialog(primaryStage);
            if (selectedFolder != null) {
                File[] excelFiles = selectedFolder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));
                if (excelFiles != null && excelFiles.length > 0) {
                    List<String> headers = getColumnHeaders(excelFiles[0]);
                    columnComboBox.getItems().setAll(headers); // Load headers into ComboBox
                }
            }
        });

        // Handle the comparison button action
        compareButton.setOnAction(e -> {
            String selectedColumn = columnComboBox.getValue();
            if (selectedFolder != null && selectedColumn != null && !selectedColumn.isEmpty()) {
                try {
                    compareRatesInFolder(selectedFolder, selectedColumn);
                    Alert alert = new Alert(Alert.AlertType.INFORMATION, "Comparison complete. Check the output folder.");
                    alert.showAndWait();
                } catch (IOException ex) {
                    Alert alert = new Alert(Alert.AlertType.ERROR, "Error during comparison: " + ex.getMessage());
                    alert.showAndWait();
                    ex.printStackTrace();
                }
            } else {
                Alert alert = new Alert(Alert.AlertType.ERROR, "Please select a folder and a valid column header.");
                alert.showAndWait();
            }
        });

        // Set up the scene and stage
        Scene scene = new Scene(grid, 500, 250);
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
             Workbook workbook = WorkbookFactory.create(fis)) {

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

    /**
     * Compares rates in all Excel files in the selected folder.
     * @param folder the folder containing Excel files.
     * @param selectedColumn the column header to compare.
     * @throws IOException if an I/O error occurs.
     */
    private void compareRatesInFolder(File folder, String selectedColumn) throws IOException {
        File[] excelFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));

        if (excelFiles == null || excelFiles.length == 0) {
            System.out.println("No Excel files found in the folder.");
            return;
        }

        // Map to store rates with Proc + Mod + Mod2 as key and month-year as subkey
        Map<String, Map<String, Double>> ratesMap = new HashMap<>();

        // Pattern to extract month and year from file names
        Pattern pattern = Pattern.compile("([a-zA-Z]+)-?(\\d{4})");

        for (File file : excelFiles) {
            Matcher matcher = pattern.matcher(file.getName());
            if (matcher.find()) {
                String monthYear = matcher.group(1) + "-" + matcher.group(2);
                System.out.println("Processing file: " + file.getName() + " for " + monthYear);
                extractRatesFromExcel(file, selectedColumn, monthYear, ratesMap);
            }
        }

        // Write output to an Excel file
        writeOutputToExcel(ratesMap);
    }

    /**
     * Extracts rates from a given Excel file.
     * @param file the Excel file to read.
     * @param selectedColumn the column header to compare.
     * @param monthYear the month-year derived from the file name.
     * @param ratesMap the map to store extracted rates.
     */
    private void extractRatesFromExcel(File file, String selectedColumn, String monthYear, Map<String, Map<String, Double>> ratesMap) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return;

            // Get the index of the selected column
            int columnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equalsIgnoreCase(selectedColumn)) {
                    columnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndex == -1) return; // Column not found

            // Read rates from the column
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell keyCell1 = row.getCell(0); // Proc
                Cell keyCell2 = row.getCell(1); // Mod
                Cell keyCell3 = row.getCell(2); // Mod2
                Cell rateCell = row.getCell(columnIndex);

                if (keyCell1 != null && keyCell2 != null && keyCell3 != null && rateCell != null && rateCell.getCellType() == CellType.NUMERIC) {
                    String key = keyCell1.getStringCellValue().trim() + "+" + keyCell2.getStringCellValue().trim() + "+" + keyCell3.getStringCellValue().trim();
                    double rate = rateCell.getNumericCellValue();

                    ratesMap.putIfAbsent(key, new HashMap<>());
                    ratesMap.get(key).put(monthYear, rate);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Writes the extracted rates to an output Excel file.
     * @param ratesMap the map containing extracted rates.
     * @throws IOException if an I/O error occurs.
     */
    private void writeOutputToExcel(Map<String, Map<String, Double>> ratesMap) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Rates Comparison");

        int rowNum = 0;
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Proc+Mod+Mod2");

        Set<String> allMonthsYears = new TreeSet<>();
        ratesMap.values().forEach(map -> allMonthsYears.addAll(map.keySet()));
        List<String> sortedMonthsYears = new ArrayList<>(allMonthsYears);

        for (int i = 0; i < sortedMonthsYears.size(); i++) {
            headerRow.createCell(i + 1).setCellValue(sortedMonthsYears.get(i));
        }

        for (Map.Entry<String, Map<String, Double>> entry : ratesMap.entrySet()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(entry.getKey());
            Map<String, Double> monthYearRates = entry.getValue();

            for (int i = 0; i < sortedMonthsYears.size(); i++) {
                String monthYear = sortedMonthsYears.get(i);
                Double rate = monthYearRates.get(monthYear);
                if (rate != null) {
                    row.createCell(i + 1).setCellValue(rate);
                }
            }
        }

        try (FileOutputStream fos = new FileOutputStream("ComparisonReport.xlsx")) {
            workbook.write(fos);
            System.out.println("Comparison report generated: ComparisonReport.xlsx");
        }

        workbook.close();
    }
}
