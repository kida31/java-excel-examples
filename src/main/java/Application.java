import excel.ApacheExcel;
import excel.FontHighlight;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.IOException;
import java.util.List;

class Application {
    public static void main(String[]args){
        System.out.println("Hello");

        try (var obj = new ApacheExcel()) {
            obj.setHeader(List.of("Name", "Count", "Price", "Description"));
            obj.appendRow(List.of("R2D2", 1, "Invaluable", "A cute robot"), FontHighlight.UNIQUE_IN_A);
            obj.appendRow(List.of("C3PO", 2, 0, "Helpful"), FontHighlight.UNIQUE_IN_B);
            obj.appendRow(List.of("Some", "Common", "Data"), FontHighlight.COMMON);
            obj.saveAs("./generatedExample.xlsx");
            System.out.println("Complete");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}