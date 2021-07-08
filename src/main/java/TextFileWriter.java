import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.ls.LSOutput;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Array;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class TextFileWriter {

    static int vi =5;
    static int vf =8;
    static StringBuilder sb = new StringBuilder();
    static Iterator<Row> iterator;
    static FileInputStream inputStream;
    static Workbook workbook;

    public static void main(String[] args){
        int dataSize = 1024 * 1024;
        Runtime runtime = Runtime.getRuntime();

        try {

            String excelFilePath = "/Users/mariojorgepaixaodealbuquerque/PJ/teste.xls";

            inputStream = new FileInputStream(new File(excelFilePath));

            workbook = new HSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            FileWriter file = new FileWriter("/Users/mariojorgepaixaodealbuquerque/PJ/fim.txt");
            iterator = firstSheet.iterator();
            System.out.println(firstSheet.getLastRowNum());

            new Thread(t1).start();
            new Thread(t2).start();
            new Thread(t3).start();
            new Thread(t4).start();


            System.out.println("Memoria máximo: " + runtime.maxMemory() / dataSize + "MB");
            System.out.println("Memoria total: " + runtime.totalMemory() / dataSize + "MB");
            System.out.println("Memoria livre: " + runtime.freeMemory() / dataSize + "MB");
            System.out.println("Memoria usada: " + (runtime.totalMemory() - runtime.freeMemory()) / dataSize + "MB");

        }catch(Exception e) {
            e.printStackTrace();
        }
    }

    public static void extracted(int valorInicial, int valorFinal) throws IOException {
        for(int i= valorInicial; i< valorFinal; i++){
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println(cell.getStringCellValue());
                sb.append(cell.getStringCellValue() );
            }
            sb.append("\n");
        }

        workbook.close();
        inputStream.close();
        Path path = Paths.get("/Users/mariojorgepaixaodealbuquerque/PJ/fim.txt");
        Files.write(path, Arrays.asList(sb.toString()));
    }

    public static Runnable t1 = new Runnable() {
        @Override
        public void run() {
            try {
                extracted(0, 5);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    };

    public static Runnable t2 = new Runnable() {
        @Override
        public void run() {
            try {
                extracted(6, 10);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    };

    public static Runnable t3 = new Runnable() {
        @Override
        public void run() {
            try {
                extracted(11, 15);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    };

    public static Runnable t4 = new Runnable() {
        @Override
        public void run() {
            try {
                extracted(16, 23);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    };
}
