package excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class ApacheExcelStreaming implements ExcelApi {
    private final SXSSFWorkbook workbook;
    private final SXSSFSheet sheet;
    private final HashMap<FontHighlight, CellStyle> styles;

    public ApacheExcelStreaming() {
        this.workbook = new SXSSFWorkbook();
        this.sheet = workbook.createSheet();

        this.sheet.setColumnWidth(0, 6000);
        this.sheet.setColumnWidth(1, 4000);

        this.styles = new HashMap();
        this.styles.put(FontHighlight.HEADER, createHeaderStyle());
        this.styles.put(FontHighlight.COMMON, createSimpleStyle(IndexedColors.BLACK));
        this.styles.put(FontHighlight.UNIQUE_IN_A, createSimpleStyle(IndexedColors.RED));
        this.styles.put(FontHighlight.UNIQUE_IN_B, createSimpleStyle(IndexedColors.BLUE));
    }

    private CellStyle createSimpleStyle(IndexedColors color) {
        var font = workbook.createFont();
        font.setColor(color.getIndex());
        var style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }

    private CellStyle createHeaderStyle() {
        var font = workbook.createFont();
        font.setBold(true);

        var style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        return style;
    }

    @Override
    public void close() throws IOException {
        this.workbook.close();
    }

    @Override
    public ExcelApi setHeader(List<Object> values) {
        var headerRow = this.sheet.createRow(0);

        for (int i = 0; i < values.size(); i++) {
            var cell = headerRow.createCell(i);
            cell.setCellValue(values.get(i).toString());
            cell.setCellStyle(styles.get(FontHighlight.HEADER));
        }

        return this;
    }

    @Override
    public ExcelApi appendRow(List<Object> values, FontHighlight highlight) {
        var row = this.sheet.createRow(this.sheet.getLastRowNum() + 1);

        for (int i = 0; i < values.size(); i++) {
            var cell = row.createCell(i);
            cell.setCellValue(values.get(i).toString());
            cell.setCellStyle(styles.get(highlight));
        }

        return this;
    }

    @Override
    public ExcelApi appendRow(List<Object> values) {
        return this.appendRow(values, FontHighlight.COMMON);
    }

    @Override
    public ExcelApi appendRows(List<List<Object>> values) {
        values.forEach(this::appendRow);
        return this;
    }

    @Override
    public void saveAs(String fileLocation) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(fileLocation)) {
            workbook.write(outputStream);
        }
    }
}
