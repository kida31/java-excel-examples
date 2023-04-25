package excel;

import java.io.IOException;
import java.util.List;

public class JExcel implements ExcelApi{

    public JExcel() {

    }

    @Override
    public ExcelApi appendRow(List<Object> values, FontHighlight highlight) {
        return null;
    }

    @Override
    public ExcelApi appendRow(List<Object> values) {
        return null;
    }

    @Override
    public ExcelApi appendRows(List<List<Object>> values) {
        return null;
    }

    @Override
    public ExcelApi setHeader(List<Object> values) {
        return null;
    }

    @Override
    public void saveAs(String fileLocation) throws IOException {

    }

    @Override
    public void close() throws Exception {

    }
}
