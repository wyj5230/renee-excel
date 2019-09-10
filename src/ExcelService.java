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
                Cloth cloth = validateAndCreateCloth(currentRow, "2018PE", "1", "FALSE", "TRUE", "0");
                clothList.add(cloth);
                clothMap.put(cloth.getStyleItem(), cloth);
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
        System.out.println("成功生成新的原表：newActual.xls");
        workbook.close();
        inputStream.close();
        System.out.println("生成新表，sheet1共有：" + clothList.size() + "条记录, sheet2（去重后）共有：" + clothMap.size() + "条记录");
    }

    public static Cloth validateAndCreateCloth(Row row, String yearSeasonFilter, String originalFilter,
                                               String cnFilter, String styleTytleFilter, String stockFilter)
    {
        String yearSeason = row.getCell(40).getStringCellValue();
        int original = (int) row.getCell(12).getNumericCellValue();
        String cnr = row.getCell(15).getStringCellValue();
        String styleTytle = row.getCell(21).getStringCellValue();
        int stock = (int) row.getCell(31).getNumericCellValue();
        System.out.println("yearSeason:" + yearSeason + ", original:" + original + "cnr:" + ", " + cnr + ", " +
                "styleTytle:" + styleTytle + ", stock:" + stock);
        String styleItem = row.getCell(0) != null ? row.getCell(0).getStringCellValue() : null;
        String styleItemColor = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : null;
        String shelfItem = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : null;
        String title = row.getCell(3) != null ? row.getCell(3).getStringCellValue() : null;
        String brandDesc = row.getCell(4) != null ? row.getCell(4).getStringCellValue() : null;
        String brandCode = row.getCell(5) != null ? row.getCell(5).getStringCellValue() : null;
        int size = 0;
        if (row.getCell(6) != null)
        {
            size = (int) row.getCell(6).getNumericCellValue();
        }
        String gender = row.getCell(7) != null ? row.getCell(7).getStringCellValue() : null;
        String lowestCategories1 = row.getCell(17) != null ? row.getCell(17).getStringCellValue() : null;
        String lowestCategories2 = row.getCell(18) != null ? row.getCell(18).getStringCellValue() : null;
        String lowestCategories3 = row.getCell(19) != null ? row.getCell(19).getStringCellValue() : null;
        String lowestCategories4 = row.getCell(20) != null ? row.getCell(20).getStringCellValue() : null;
        return new Cloth(styleItem, styleItemColor, shelfItem, title, brandDesc, brandCode, size, gender,
                lowestCategories1, lowestCategories2, lowestCategories3, lowestCategories4);
    }


}
