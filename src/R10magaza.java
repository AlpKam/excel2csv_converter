import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.CSVWriter;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class R10magaza {
    private static final List<String> HEADERS = Arrays.asList(
        "Store Code", "Store Name", "Region", "Format", "System Name",
        "IP Address", "Total Register", "Normal Register", "Self Checkout",
        "Cafe Register", "Other Register", "Local Server", "Station"
        /* burdaki isimler temsilidir */
);
    );

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.print("Enter the full path to the input Excel file (.xlsx): ");
        String inputPath = scanner.nextLine().replace("\"", "").trim();

        System.out.print("Enter the full path for the output CSV file (.csv): ");
        String outputPath = scanner.nextLine().replace("\"", "").trim();

        File inputFile = new File(inputPath);
        File outputFile = new File(outputPath);

        if (outputFile.isDirectory()) {
            String fileName = inputFile.getName().replaceFirst("[.][^.]+$", "") + ".csv";
            outputFile = new File(outputFile, fileName);
        }

        try (
                FileInputStream fis = new FileInputStream(inputFile);
                Workbook workbook = new XSSFWorkbook(fis);
                CSVWriter writer = new CSVWriter(
                        new OutputStreamWriter(new FileOutputStream(outputFile), StandardCharsets.UTF_8),
                        ',', '"', // her değeri çift tırnak içinde yaz
                        CSVWriter.DEFAULT_ESCAPE_CHARACTER,
                        CSVWriter.DEFAULT_LINE_END
                )
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            writer.writeNext(HEADERS.toArray(new String[0]));

            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String[] csvRow = new String[HEADERS.size()];
                for (int j = 0; j < HEADERS.size(); j++) {
                    Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    String value = (cell != null) ? getCleanedCellValue(cell, evaluator) : "";
                    csvRow[j] = value;
                }
                writer.writeNext(csvRow);
            }

            System.out.println("Data has been successfully exported to " + outputFile.getAbsolutePath());

        } catch (Exception e) {
            System.err.println("Failed to export the CSV file. Error: " + e.getMessage());
        }
    }

    private static String getCleanedCellValue(Cell cell, FormulaEvaluator evaluator) {
        String value = "";

        try {
            if (cell.getCellType() == CellType.FORMULA) {
                CellValue cellValue = evaluator.evaluate(cell);
                if (cellValue == null) return "";
                switch (cellValue.getCellType()) {
                    case STRING:
                        value = cellValue.getStringValue();
                        break;
                    case NUMERIC:
                        double num = cellValue.getNumberValue();
                        value = (num == Math.floor(num)) ? String.valueOf((int) num) : String.valueOf(num);
                        break;
                    case BOOLEAN:
                        value = String.valueOf(cellValue.getBooleanValue());
                        break;
                    default:
                        value = "";
                }
            } else {
                switch (cell.getCellType()) {
                    case STRING:
                        value = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        double num = cell.getNumericCellValue();
                        value = (num == Math.floor(num)) ? String.valueOf((int) num) : String.valueOf(num);
                        break;
                    case BOOLEAN:
                        value = String.valueOf(cell.getBooleanCellValue());
                        break;
                    default:
                        value = "";
                }
            }
        } catch (Exception e) {
            value = "";
        }

        // Görünmez boşlukları ve özel karakterleri temizle
        value = value
                .replaceAll("[\\u00A0\\u2007\\u202F\\t\\r\\n]", "") // non-breaking space, tab, vb.
                .replaceAll("_+(\\d+)_+", "$1") // _6_ gibi yapılar → 6
                .trim();

        return value.isEmpty() ? "" : value;
    }
}


