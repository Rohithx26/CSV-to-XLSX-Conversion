package CSVtoXLSX;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AdvancedCSVToXLSXConverter {
    public static void main(String[] args) {
        String csvFilePath = "C:\\Users\\Rohit\\OneDrive - Digital Software Inc\\Exception_File\\Exceptions.csv";
        String xlsxFilePath = "C:\\Users\\Rohit\\OneDrive - Digital Software Inc\\Recordings\\Exceptions.xlsx";
        convertCSVToXLSX(csvFilePath, xlsxFilePath);
    }

    public static void convertCSVToXLSX(String csvPath, String xlsxPath) {
        try (
            Reader reader = new InputStreamReader(new FileInputStream(csvPath), StandardCharsets.UTF_8);
            @SuppressWarnings("deprecation")
			CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT)
        ) {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Data");

            int rowNumber = 0;
            for (CSVRecord csvRecord : csvParser) {
                Row row = sheet.createRow(rowNumber++);
                for (int i = 0; i < csvRecord.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(csvRecord.get(i)); // Always as string
                }
            }

            // Auto-size columns
            if (sheet.getRow(0) != null) {
                for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
                    sheet.autoSizeColumn(i);
                }
            }

            // Optional: Style header row
            if (sheet.getRow(0) != null) {
                CellStyle headerStyle = workbook.createCellStyle();
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerStyle.setFont(headerFont);
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                Row headerRow = sheet.getRow(0);
                for (Cell cell : headerRow) {
                    cell.setCellStyle(headerStyle);
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream(xlsxPath)) {
                workbook.write(outputStream);
            }
            workbook.close();

            System.out.println("Conversion completed successfully!");
        } catch (IOException e) {
            System.err.println("Error: " + e.getMessage());
        }
    }
}