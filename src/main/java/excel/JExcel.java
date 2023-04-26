package excel;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.*;

import java.io.*;
import java.util.HashMap;
import java.util.List;

public class JExcel implements ExcelApi {
    private final WritableWorkbook workbook;
    private final WritableSheet sheet;
    private final HashMap<FontHighlight, WritableCellFormat> styles;

    public JExcel() throws IOException, WriteException {
        this.workbook = Workbook.createWorkbook(new File("test.xls"));
        this.sheet = workbook.createSheet("Sheet1", 0);

        this.sheet.setColumnView(0, 20);
        this.sheet.setColumnView(1, 15);

        this.styles = new HashMap<>();
        this.styles.put(FontHighlight.HEADER, createHeaderStyle());
        this.styles.put(FontHighlight.COMMON, createSimpleStyle(Colour.BLACK));
        this.styles.put(FontHighlight.UNIQUE_IN_A, createSimpleStyle(Colour.RED));
        this.styles.put(FontHighlight.UNIQUE_IN_B, createSimpleStyle(Colour.BLUE));
    }

    private WritableCellFormat createSimpleStyle(Colour color) throws WriteException {
        var font = new jxl.write.WritableFont(jxl.write.WritableFont.ARIAL);
        font.setColour(color);
        return new WritableCellFormat(font);
    }

    private WritableCellFormat createHeaderStyle() throws WriteException {
        var font = new jxl.write.WritableFont(jxl.write.WritableFont.ARIAL, 10, jxl.write.WritableFont.BOLD);
        var style = new WritableCellFormat(font);
        style.setBackground(Colour.GRAY_25);
        return style;
    }

    @Override
    public void close() throws Exception {
        workbook.close();
    }

    @Override
    public ExcelApi setHeader(List<Object> values) {
        for (int i = 0; i < values.size(); i++) {
            Label label = new Label(i, 0, values.get(i).toString());
            try {
                sheet.addCell(label);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        return this;
    }

    @Override
    public ExcelApi appendRow(List<Object> values, FontHighlight highlight) {
        int row = sheet.getRows();
        for (int i = 0; i < values.size(); i++) {
            Label label = new Label(i, row, values.get(i).toString());
            try {
                sheet.addCell(label);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        return this;
    }

    @Override
    public ExcelApi appendRow(List<Object> values) {
        return this.appendRow(values, FontHighlight.COMMON);
    }

    @Override
    public ExcelApi appendRows(List<List<Object>> values) {
        for (List<Object> row : values) {
            this.appendRow(row);
        }
        return this;
    }

    @Override
    public void saveAs(String fileLocation) throws IOException, WriteException {
        workbook.setOutputFile(new File(fileLocation));
        workbook.write();
    }
}