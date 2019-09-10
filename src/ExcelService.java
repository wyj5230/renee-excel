import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

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
                            currentRow.getCell(1).getStringCellValue(),
                            currentRow.getCell(2).getStringCellValue()));
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
        Cell lastleft = second.createCell(40);
        Cell lastright = second.createCell(41);
        lastleft.setCellValue("yearSeason");
        lastright.setCellValue("event");
        List<Cloth> clothList = new ArrayList<>();
        Map<String, Cloth> clothMap = new HashMap<>();
        for (int i = 2; i <= total; i++)
        {
            Row currentRow = firstSheet.getRow(i);
            String id = currentRow.getCell(1).getStringCellValue();
            Sample sample = sampleMap.get(id);
            if (sample != null)
            {
                Cell newNine = currentRow.createCell(40);
                newNine.setCellValue(sample.getYearSeason());
                Cell newTen = currentRow.createCell(41);
                newTen.setCellValue(sample.getEvent());
                validateAndCreateCloth(currentRow, "2018PE", "1", "FALSE", "TRUE", "0");
            }
            else
            {
                System.out.println("未能在对照表中发现货品：");
            }
        }
        FileOutputStream out =
                new FileOutputStream(new File("./newActual.xls"));
        workbook.write(out);
        out.close();
        workbook.close();
        inputStream.close();
    }

    public static void validateAndCreateCloth(Row row, String yearSeasonFilter, String originalFilter,
                                              String cnFilter, String styleTytleFilter, String stockFilter)
    {
        String yearSeason = row.getCell(40).getStringCellValue();
        double original = row.getCell(12).getNumericCellValue();
        String cnr = row.getCell(15).getStringCellValue();
        String styleTytle = row.getCell(21).getStringCellValue();
        double stock = row.getCell(31).getNumericCellValue();
        System.out.println("yearSeason:" + yearSeason + ", original:" + original + "cnr:" + ", " + cnr + ", " +
                "styleTytle:" + styleTytle + ", stock:" + stock);
    }


}
