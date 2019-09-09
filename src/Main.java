import java.util.Map;

public class Main {

    public static void main(String[] args) throws Exception {
        Map<String, Sample> sampleMap = ExcelService.getSampleXls("./123.xls");
        ExcelService.getActualXls("./abc.xls",sampleMap);
    }
}
