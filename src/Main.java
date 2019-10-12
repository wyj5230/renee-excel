import java.io.Console;
import java.util.Map;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws Exception {
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入yearSeason：(例如'2019AI')");
        String yearSeasonFilter = sc.next();
        System.out.println(yearSeasonFilter);
        System.out.println("请输入origin：('>0' 或者 '=0')");
        String originFilter = sc.next();
        System.out.println(originFilter);
        System.out.println("请输入styleTytle：('TRUE' 或者 'FALSE')");
        String styleTytleFilter = sc.next();
        System.out.println(styleTytleFilter);
        Map<String, Sample> sampleMap = ExcelService.getSampleXls("./sample.xls");
        ExcelService.getActualXls("./actual.xls",sampleMap, yearSeasonFilter, originFilter, styleTytleFilter);
        sc.next();
    }
}
