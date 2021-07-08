import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Iterator;

public class TextFileWriter {

    public static void main(String[] args){
        int dataSize = 1024 * 1024;

        Runtime runtime = Runtime.getRuntime();
        System.out.println("Memoria máximo: " + runtime.maxMemory() / dataSize + "MB");
        System.out.println("Memoria total: " + runtime.totalMemory() / dataSize + "MB");
        System.out.println("Memoria livre: " + runtime.freeMemory() / dataSize + "MB");
        System.out.println("Memoria usada: " + (runtime.totalMemory() - runtime.freeMemory()) / dataSize + "MB");

        StringBuilder sb = new StringBuilder();
        try {

            String excelFilePath = "/Users/mariojorgepaixaodealbuquerque/PJ/teste.xls";


            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

            Workbook workbook = new HSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            FileWriter file = new FileWriter("/Users/mariojorgepaixaodealbuquerque/PJ/fim.txt");
            Iterator<Row> iterator = firstSheet.iterator();

            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue());
                            sb.append(cell.getStringCellValue() + ", ");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue());
                            sb.append(String.valueOf(cell.getBooleanCellValue() + ", "));
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            sb.append(String.valueOf(cell.getNumericCellValue() + ", "));
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula());
                            sb.append(String.valueOf(cell.getCellFormula() + ", "));
                            break;
                        default :
                            System.out.print(cell.getStringCellValue());
                            sb.append(cell.getStringCellValue() + ", ");
                    }
                    System.out.print(" - ");
                }
                System.out.println();
                sb.append("\n");
            }
            workbook.close();
            inputStream.close();
            Path path = Paths.get("/Users/mariojorgepaixaodealbuquerque/PJ/fim.txt");
            Files.write(path, Arrays.asList(sb.toString()));

        }catch(Exception e) {
            e.printStackTrace();
        }
    }

}
