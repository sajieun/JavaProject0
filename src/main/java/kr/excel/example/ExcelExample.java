package kr.excel.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelExample {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("example.xlsx"));
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row: sheet){
                for (Cell cell : row){
                    System.out.print(cell.toString()+"\t");
                }
                System.out.println();
            }
            file.close();
            System.out.println("엑셀에서 데이터 읽어오기 성공");
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}
