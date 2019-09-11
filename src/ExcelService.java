import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class ExcelService {
    public static Map<String, Sample> getSampleXls(String path) throws Exception {
        FileInputStream inputStream = new FileInputStream(new File(path));
        Map<String, Sample> sampleMap = new HashMap<>();
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        int total = firstSheet.getLastRowNum();
        System.out.println("对照表共有" + total + "行");
        for (int i = 1; i <= total - 1; i++) {
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

    public static void getActualXls(String path, Map<String, Sample> sampleMap) throws Exception {
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
        for (int i = 2; i <= total; i++) {
            Row currentRow = firstSheet.getRow(i);
            String id = currentRow.getCell(1).getStringCellValue();
            Sample sample = sampleMap.get(id);
            if (sample != null) {
                Cell newNine = currentRow.createCell(40);
                newNine.setCellValue(sample.getYearSeason());
                Cell newTen = currentRow.createCell(41);
                newTen.setCellValue(sample.getEvent());
                Cloth cloth = validateAndCreateCloth(currentRow,
                        "2019AI", "FALSE", "TRUE");
                if (cloth != null) {
                    clothList.add(cloth);
                    clothMap.put(cloth.getStyleItem(), cloth);
                }
            } else {
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
        createNewExcel(clothList, clothMap);
    }

    public static Cloth validateAndCreateCloth(Row row, String yearSeasonFilter,
                                               String cnFilter, String styleTytleFilter) {
        String yearSeason = row.getCell(40).getStringCellValue();
        if (!yearSeason.equals(yearSeasonFilter)) {
            return null;
        }
        int original = (int) row.getCell(12).getNumericCellValue();
        if (original <= 0) {
            return null;
        }
        String cnr = row.getCell(15).getStringCellValue();
        if (!cnr.equals(cnFilter)) {
            return null;
        }
        String styleTytle = row.getCell(21).getStringCellValue();
        if (!styleTytle.equals(styleTytleFilter)) {
            return null;
        }
        int price = (int) row.getCell(25).getNumericCellValue();
        if (price <= 0) {
            return null;
        }
        int stock = (int) row.getCell(31).getNumericCellValue();
        if (stock <= 0) {
            return null;
        }
        System.out.println("yearSeason:" + yearSeason + ", original:" + original + "， cnr:" + cnr + ", price：" + price +
                "，styleTytle:" + styleTytle + ", stock:" + stock);
        String styleItem = row.getCell(0) != null ? row.getCell(0).getStringCellValue() : null;
        String styleItemColor = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : null;
        String shelfItem = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : null;
        String title = row.getCell(3) != null ? row.getCell(3).getStringCellValue() : null;
        String brandDesc = row.getCell(4) != null ? row.getCell(4).getStringCellValue() : null;
        String brandCode = row.getCell(5) != null ? row.getCell(5).getStringCellValue() : null;
        Cell sizeCell = row.getCell(6);
        String size = null;
        if (sizeCell != null) {
            sizeCell.setCellType(CellType.STRING);
            size = sizeCell.getStringCellValue();
        }
        String gender = row.getCell(7) != null ? row.getCell(7).getStringCellValue() : null;
        String lowestCategories1 = row.getCell(17) != null ? row.getCell(17).getStringCellValue() : null;
        String lowestCategories2 = row.getCell(18) != null ? row.getCell(18).getStringCellValue() : null;
        String lowestCategories3 = row.getCell(19) != null ? row.getCell(19).getStringCellValue() : null;
        String lowestCategories4 = row.getCell(20) != null ? row.getCell(20).getStringCellValue() : null;
        return new Cloth(styleItem, styleItemColor, shelfItem, title, brandDesc, brandCode, size, gender,
                lowestCategories1, lowestCategories2, lowestCategories3, lowestCategories4);
    }

    public static void createNewExcel(List<Cloth> clothList, Map<String, Cloth> clothMap) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet1");
        createTtitleRow(sheet1);
        createRowFromClothList(clothList, sheet1);


        Sheet sheet2 = workbook.createSheet("Sheet2");
        createTtitleRow(sheet2);
        createRowFromClothMap(clothMap, sheet2);

        String fileName = getFileName();
        FileOutputStream out =
                new FileOutputStream(new File("./" + fileName + ".xls"));
        workbook.write(out);
        out.close();
        System.out.println("成功生成新表：" + fileName + ".xls");
        System.out.println("sheet1共有：" + clothList.size() + "条记录, sheet2（去重后）共有：" + clothMap.size() + "条记录");
        workbook.close();
    }

    private static String getFileName() {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
        Date date = new Date();
        return "Flag_Online_CN_" + formatter.format(date);
    }

    private static void createRowFromClothList(List<Cloth> clothList, Sheet sheet1) {
        for (int i = 1; i <= clothList.size(); i++) {
            Cloth cloth = clothList.get(i - 1);
            createRowFromCloth(sheet1, i, cloth);
        }
    }

    private static void createRowFromClothMap(Map<String, Cloth> clothMap, Sheet sheet1) {
        int i = 1;
        for (Cloth cloth : clothMap.values()) {
            createRowFromCloth(sheet1, i, cloth);
            i++;
        }
    }

    private static void createRowFromCloth(Sheet sheet1, int i, Cloth cloth) {
        Row currentRow = sheet1.createRow(i);
        Cell cell0 = currentRow.createCell(0);
        cell0.setCellValue(cloth.getStyleItem());
        Cell cell1 = currentRow.createCell(1);
        cell1.setCellValue(cloth.getStyleItemColor());
        Cell cell2 = currentRow.createCell(2);
        cell2.setCellValue(cloth.getShelfItem());
        Cell cell3 = currentRow.createCell(3);
        cell3.setCellValue(cloth.getTitle());
        Cell cell4 = currentRow.createCell(4);
        cell4.setCellValue(cloth.getBrandDesc());
        Cell cell5 = currentRow.createCell(5);
        cell5.setCellValue(cloth.getBrandCode());
        Cell cell6 = currentRow.createCell(6);
        cell6.setCellValue(cloth.getSize());
        Cell cell7 = currentRow.createCell(7);
        cell7.setCellValue(cloth.getGender());

        Cell cell8 = currentRow.createCell(8);
        cell8.setCellValue(cloth.getLowestCategories1());
        Cell cell9 = currentRow.createCell(9);
        cell9.setCellValue(cloth.getLowestCategories2());
        Cell cell10 = currentRow.createCell(10);
        cell10.setCellValue(cloth.getLowestCategories3());
        Cell cell11 = currentRow.createCell(11);
        cell11.setCellValue(cloth.getLowestCategories4());
    }

    private static void createTtitleRow(Sheet sheet1) {
        Row sheet1TitleRow = sheet1.createRow(0);
        Cell cell0 = sheet1TitleRow.createCell(0);
        cell0.setCellValue("Style-Item");
        Cell cell1 = sheet1TitleRow.createCell(1);
        cell1.setCellValue("Style-Item_Color");
        Cell cell2 = sheet1TitleRow.createCell(2);
        cell2.setCellValue("Shelf-Item");
        Cell cell3 = sheet1TitleRow.createCell(3);
        cell3.setCellValue("Title");
        Cell cell4 = sheet1TitleRow.createCell(4);
        cell4.setCellValue("Brand Descr.");
        Cell cell5 = sheet1TitleRow.createCell(5);
        cell5.setCellValue("Brand Code");
        Cell cell6 = sheet1TitleRow.createCell(6);
        cell6.setCellValue("Size");
        Cell cell7 = sheet1TitleRow.createCell(7);
        cell7.setCellValue("Gender");

        Cell cell8 = sheet1TitleRow.createCell(8);
        cell8.setCellValue("1");
        Cell cell9 = sheet1TitleRow.createCell(9);
        cell9.setCellValue("2");
        Cell cell10 = sheet1TitleRow.createCell(10);
        cell10.setCellValue("3");
        Cell cell11 = sheet1TitleRow.createCell(11);
        cell11.setCellValue("4");
    }
}
