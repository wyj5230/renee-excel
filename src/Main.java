import java.io.Console;
import java.util.Map;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws Exception {
        Map<String, Sample> sampleMap = ExcelService.getSampleXls("./sample.xls");
        ExcelService.getActualXls("./actual.xls",sampleMap);
        Scanner sc = new Scanner(System.in);
        sc.next();
    }
}
