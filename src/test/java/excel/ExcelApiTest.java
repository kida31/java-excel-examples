package excel;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTimeout;

public class ExcelApiTest {
    protected static final int ROW_COUNT = 100_000;
    protected static final int COLUMN_COUNT = 6;

    static Stream<ExcelApi> implementations() {
        return Stream.of(new ApacheExcel());
    }

    @ParameterizedTest
    @MethodSource("implementations")
    @DisplayName("Performance test: writing big file with default highlight")
    void testWriteBigFileWithDefaultHighlight(ExcelApi excelApi) throws Exception {
        List<Object> rowValues = createRowValues(COLUMN_COUNT);
        List<List<Object>> rows = createRows(rowValues, ROW_COUNT);

        assertTimeout(
                java.time.Duration.ofSeconds(30),
                () -> excelApi.setHeader(rowValues).appendRows(rows).saveAs("big-file-default-highlight.xlsx")
        );

        excelApi.close();
    }

    private static List<Object> createRowValues(int columnCount) {
        List<Object> rowValues = new ArrayList<>();
        Random random = new Random();
        for (int i = 0; i < columnCount; i++) {
            rowValues.add(random.nextDouble());
        }
        return rowValues;
    }

    private static List<List<Object>> createRows(List<Object> rowValues, int rowCount) {
        List<List<Object>> rows = new ArrayList<>();
        for (int i = 0; i < rowCount; i++) {
            rows.add(rowValues);
        }
        return rows;
    }
}
