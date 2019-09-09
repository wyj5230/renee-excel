import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelService
{
    public static Map<String, Sample> getSampleXls(String path) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(new File(path));
        Map<String, Sample> sampleMap = new HashMap<>();
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        int total = firstSheet.getLastRowNum();
        System.out.println("对照表共有" + total + "行");
        for (int i = 1; i <= total - 1; i++)
        {
            Row currentRow = firstSheet.getRow(i);
            sampleMap.put(currentRow.getCell(0).getStringCellValue(),
                    new Sample(currentRow.getCell(0).getStringCellValue(),
                            currentRow.getCell(0).getStringCellValue(),
                            currentRow.getCell(0).getStringCellValue()));
        }
        System.out.println("对照表(去重)共有" + sampleMap.size() + "行");
        workbook.close();
        inputStream.close();
        return sampleMap;
    }

    public static void getActualXls(String path, Map<String, Sample> sampleMap) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(new File(path));
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        int total = firstSheet.getLastRowNum();
        System.out.println("原表共有" + total + "行");
        Row second = firstSheet.getRow(1);
        moveRight(second);
        Cell nine = second.createCell(9);
        Cell ten = second.createCell(10);
        nine.setCellValue("yearSeason");
        ten.setCellValue("event");
        second.getLastCellNum();
//        for (int i = 2; i <= total - 1; i++)
//        {
//            Row currentRow = firstSheet.getRow(i);
//            Sample sample = sampleMap.get(currentRow.getCell(0).getStringCellValue());
//            if (sample != null){
//                Cell newNine = currentRow.createCell(9);
//                Cell newTen = currentRow.createCell(10);
//            }
//        }
        FileOutputStream out =
                new FileOutputStream(new File("./new.xls"));
        workbook.write(out);
        out.close();
        workbook.close();
        inputStream.close();
    }

    private static void moveRight(Row row)
    {
        for (int i = 39; i >= 9; i--)
        {
            Cell right = row.createCell(i + 2);
            Cell left = row.getCell(i);
            if (left.getCellType() == CellType.BOOLEAN)
            {
                right.setCellValue(left.getBooleanCellValue());
            }
            else if (left.getCellType() == CellType.NUMERIC)
            {
                right.setCellValue(left.getNumericCellValue());
            }
            else
            {
                right.setCellValue(left.getStringCellValue());
            }
            right.setCellStyle(left.getCellStyle());
        }
    }
}
