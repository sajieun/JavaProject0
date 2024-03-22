package kr.excel.example;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class ExcelExampleReview {
    public static void main(String[] args) {
        try {
            FileInputStream fileInputStream = new FileInputStream("example.xlsx");

            Workbook workbook = WorkbookFactory.create(fileInputStream);

            Sheet sheet = workbook.getSheetAt(0);

            SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd");

            for (Row row : sheet){
                for (Cell cell : row){
                    if (cell.getCellType() == CellType.NUMERIC){
                        if (DateUtil.isCellDateFormatted(cell)){
                            System.out.println(sdf.format(cell.getNumericCellValue())+"\t");
                        }else {
                            if (cell.getNumericCellValue() == (int) cell.getNumericCellValue()){
                                System.out.println((int) cell.getNumericCellValue()+"\t");
                            }else {
                                System.out.println(cell.getNumericCellValue()+"\t");
                            }
                        }
                    }else {
                        System.out.println(cell + "\t");
                    }
                }
                System.out.println();
            }
            fileInputStream.close();
            System.out.println("엑셀에서 데이터 읽어오기 성공");
        }catch (IOException e){
            e.printStackTrace();
        }
    }
}
