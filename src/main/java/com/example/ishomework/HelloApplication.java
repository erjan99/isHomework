package com.example.ishomework;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class HelloApplication extends Application {

    private final Map<Integer, Map<Integer, Double>> yearlyMonthlySales = new HashMap<>();
    private ComboBox<Integer> yearComboBox;
    private Label statusLabel;
    private LineChart<String, Number> chart;

    // Column indexes based on your file
    private static final int DATE_COLUMN = 5;
    private static final int TOTAL_COLUMN = 4;

    private static final DateTimeFormatter DATE_FORMATTER =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Анализатор продаж");

        Button loadButton = new Button("Загрузить  файл");
        yearComboBox = new ComboBox<>();
        yearComboBox.setPromptText("Выберите год");
        Button analyzeButton = new Button("Анализировать");
        statusLabel = new Label("Загрузите файл  ");



        CategoryAxis xAxis = new CategoryAxis();
        xAxis.setLabel("Месяц");
        NumberAxis yAxis = new NumberAxis();
        yAxis.setLabel("Прибыль ");

        chart = new LineChart<>(xAxis, yAxis);
        chart.setTitle("График продаж по месяцам");



        loadButton.setOnAction(e -> loadExcelFile(primaryStage));
        analyzeButton.setOnAction(e -> analyzeData());

        VBox vbox = new VBox(10,
                new HBox(10, loadButton, yearComboBox, analyzeButton),
                statusLabel,
                chart);

        Scene scene = new Scene(vbox, 800, 600);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void loadExcelFile(Stage primaryStage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Открыть файл ");
        fileChooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));

        File file = fileChooser.showOpenDialog(primaryStage);
        if (file != null) {
            try (FileInputStream fis = new FileInputStream(file);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                yearlyMonthlySales.clear();
                Set<Integer> years = new HashSet<>();

                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();

                if (rowIterator.hasNext()) rowIterator.next();

                int processedRows = 0;
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    try {
                        Cell dateCell = row.getCell(DATE_COLUMN);
                        if (dateCell == null) continue;
                        LocalDateTime saleDateTime;
                        if (dateCell.getCellType() == CellType.NUMERIC) {
                            saleDateTime = dateCell.getDateCellValue().toInstant()
                                    .atZone(ZoneId.systemDefault()).toLocalDateTime();
                        } else {
                            String dateString = dateCell.getStringCellValue().trim();
                            saleDateTime = LocalDateTime.parse(dateString, DATE_FORMATTER);
                        }
                        Cell totalCell = row.getCell(TOTAL_COLUMN);
                        double totalSale;
                        if (totalCell.getCellType() == CellType.FORMULA) {
                            totalSale = totalCell.getNumericCellValue(); // Get calculated value
                        } else if (totalCell.getCellType() == CellType.NUMERIC) {
                            totalSale = totalCell.getNumericCellValue();
                        }else {
                            totalSale = Double.parseDouble(totalCell.getStringCellValue());}

                        LocalDate saleDate = saleDateTime.toLocalDate();
                        int year = saleDate.getYear();
                        int month = saleDate.getMonthValue();

                        yearlyMonthlySales
                                .computeIfAbsent(year, k -> new HashMap<>())
                                .merge(month, totalSale, Double::sum);

                        years.add(year);
                        processedRows++;
                    } catch (Exception e) {
                        System.err.println("Ошибка обработки строки " + (row.getRowNum()+1) + ": " + e.getMessage());
                    }
                }

                yearComboBox.getItems().setAll(years.stream().sorted().collect(Collectors.toList()));
                statusLabel.setText(String.format("Файл загружен. ",
                        processedRows, years.size()));

            } catch (Exception e) {
                statusLabel.setText("Ошибка загрузки файла: " + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    private void analyzeData() {
        Integer selectedYear = yearComboBox.getValue();
        if (selectedYear == null) {
            statusLabel.setText("Выберите год для анализа");
            return;
        }


        Map<Integer, Double> monthlySales = yearlyMonthlySales.get(selectedYear);
        if (monthlySales == null || monthlySales.isEmpty()) {
            statusLabel.setText("Нет данных для выбранного года");
            return;
        }

        XYChart.Series<String, Number> series = new XYChart.Series<>();

        for (int month = 1; month <= 12; month++) {
            String monthName = getMonthName(month);
            double sales = monthlySales.getOrDefault(month, 0.0);
            series.getData().add(new XYChart.Data<>(monthName, sales));
        }



        chart.getData().clear();
        chart.getData().add(series);

    }

    private String getMonthName(int month) {
        return switch (month) {
            case 1 -> "Январь";
            case 2 -> "Февраль";
            case 3 -> "Март";
            case 4 -> "Апрель";
            case 5 -> "Май";
            case 6 -> "Июнь";
            case 7 -> "Июль";
            case 8 -> "Август";
            case 9 -> "Сентябрь";
            case 10 -> "Октябрь";
            case 11 -> "Ноябрь";
            case 12 -> "Декабрь";
            default -> "";
        };
    }



    public static void main(String[] args) {
        launch(args);
    }
}