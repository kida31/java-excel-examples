package excel;

import java.io.IOException;
import java.util.List;

public interface ExcelApi extends AutoCloseable {
    ExcelApi appendRow(List<Object> values, FontHighlight highlight);
    ExcelApi appendRow(List<Object> values);
    ExcelApi appendRows(List<List<Object>> values);
    ExcelApi setHeader(List<Object> values);
    void saveAs(String fileLocation) throws IOException;
}
