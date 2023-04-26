package excel;

import jxl.write.WriteException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTimeout;

public class ExcelApiTest {
    protected static final int ROW_COUNT = 10_000;
    protected static final int COLUMN_COUNT = 6;

    static Stream<ExcelApi> implementations() throws WriteException, IOException {
        return Stream.of(new JExcel(), new ApacheExcel(), new ApacheExcelStreaming());
    }

    @ParameterizedTest
    @MethodSource("implementations")
    @DisplayName("Performance test: writing big file with default highlight")
    void testWriteBigFileWithDefaultHighlight(ExcelApi excelApi) throws Exception {
        var header = createRowValues(COLUMN_COUNT);
        List<List<Object>> rows = createRows(ROW_COUNT, COLUMN_COUNT);

        var sh = excelApi.getClass().getSimpleName().toLowerCase();
        assertTimeout(
                java.time.Duration.ofSeconds(30),
                () -> excelApi.setHeader(header).appendRows(rows).saveAs("big-file-default-highlight-" + sh + ".xlsx")
        );
    }

    private static List<Object> createRowValues(int columnCount, Random random) {
        List<Object> rowValues = new ArrayList<>();
        for (int i = 0; i < columnCount; i++) {
            rowValues.add(random.nextDouble());
        }
        return rowValues;
    }

    private static List<Object> createRowValues(int columnCount) {
        return createRowValues(columnCount, new Random());
    }

    private static List<List<Object>> createRows(int rowCount, int columnCount) {
        Random random = new Random();
        List<List<Object>> rows = new ArrayList<>();
        for (int i = 0; i < rowCount; i++) {
            rows.add(createRowValues(columnCount, random));
        }
        return rows;
    }
}
