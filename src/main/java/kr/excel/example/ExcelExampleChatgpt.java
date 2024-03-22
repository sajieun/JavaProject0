package kr.excel.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class ExcelExampleChatgpt {
    public static void main(String[] args) {
        try {
            // 1. Excel 파일을 읽어들이기 위한 FileInputStream을 생성
            FileInputStream file = new FileInputStream(new File("example.xlsx"));

            // 2. 엑셀 파일로부터 Workbook 객체 생성
            Workbook workbook = WorkbookFactory.create(file);

            // 3. 첫 번째 시트를 선택
            Sheet sheet = workbook.getSheetAt(0);

            // 4. 날짜 형식을 변환하기 위한 SimpleDateFormat 객체 생성
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

            // 5. 각 행을 순회하면서 셀의 내용을 출력
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // 6. 셀의 타입이 숫자인 경우
                    if (cell.getCellType() == CellType.NUMERIC) {
                        // 7. 날짜 형식인 경우
                        if (DateUtil.isCellDateFormatted(cell)) {
                            // 8. 날짜를 지정된 형식으로 변환하여 출력
                            System.out.print(sdf.format(cell.getDateCellValue()) + "\t");
                        } else {
                            // 9. 소수점이 있는 경우는 그대로 출력
                            if (cell.getNumericCellValue() == (int) cell.getNumericCellValue()) {
                                System.out.print((int) cell.getNumericCellValue() + "\t");
                            } else {
                                System.out.print(cell.getNumericCellValue() + "\t");
                            }
                        }
                    } else {
                        // 10. 숫자가 아닌 경우는 그대로 출력
                        System.out.print(cell + "\t");
                    }
                }
                System.out.println(); // 11. 행의 끝을 나타내기 위한 개행
            }

            // 12. 파일을 닫음
            file.close();

            // 13. 처리 완료 메시지 출력
            System.out.println("엑셀에서 데이터 읽어오기 성공");
        } catch (IOException e) {
            // 14. IOException 발생 시 예외 메시지 출력
            e.printStackTrace();
        }
    }
}
