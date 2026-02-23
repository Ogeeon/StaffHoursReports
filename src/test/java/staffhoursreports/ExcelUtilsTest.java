package staffhoursreports;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.time.LocalDate;
import java.time.Month;
import java.util.Calendar;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.*;

class ExcelUtilsTest {

    private Workbook workbook;
    private Sheet sheet;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test");
    }

    @AfterEach
    void tearDown() throws IOException {
        workbook.close();
    }

    private Cell createCell(int row, int col) {
        Row r = sheet.createRow(row);
        return r.createCell(col);
    }

    // --- getCellAsString tests ---

    @Test
    void getCellAsString_nullCell_returnsEmpty() {
        assertEquals("", ExcelUtils.getCellAsString(null));
    }

    @Test
    void getCellAsString_stringCell_returnsString() {
        Cell cell = createCell(0, 0);
        cell.setCellValue("  Hello World  ");
        assertEquals("Hello World", ExcelUtils.getCellAsString(cell));
    }

    @Test
    void getCellAsString_wholeNumericCell_returnsIntegerString() {
        Cell cell = createCell(1, 0);
        cell.setCellValue(42.0);
        assertEquals("42", ExcelUtils.getCellAsString(cell));
    }

    @Test
    void getCellAsString_fractionalNumericCell_returnsDecimalString() {
        Cell cell = createCell(2, 0);
        cell.setCellValue(3.14);
        String result = ExcelUtils.getCellAsString(cell);
        assertEquals("3.14", result);
    }

    @Test
    void getCellAsString_blankCell_returnsEmpty() {
        Cell cell = createCell(3, 0);
        cell.setBlank();
        assertEquals("", ExcelUtils.getCellAsString(cell));
    }

    @Test
    void getCellAsString_booleanTrueCell_returnsTrue() {
        Cell cell = createCell(4, 0);
        cell.setCellValue(true);
        assertEquals("true", ExcelUtils.getCellAsString(cell));
    }

    @Test
    void getCellAsString_booleanFalseCell_returnsFalse() {
        Cell cell = createCell(5, 0);
        cell.setCellValue(false);
        assertEquals("false", ExcelUtils.getCellAsString(cell));
    }

    // --- getCellAsDate tests ---

    @Test
    void getCellAsDate_nullCell_returnsNull() {
        assertNull(ExcelUtils.getCellAsDate(null));
    }

    @Test
    void getCellAsDate_dateFormattedNumericCell_returnsDate() {
        Cell cell = createCell(6, 0);
        // Create a date-formatted cell
        CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
        cell.setCellStyle(dateStyle);

        Calendar cal = Calendar.getInstance();
        cal.set(2024, Calendar.MARCH, 15, 0, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        Date expectedDate = cal.getTime();
        cell.setCellValue(expectedDate);

        Date result = ExcelUtils.getCellAsDate(cell);
        assertNotNull(result);
        // Compare only year/month/day
        Calendar resultCal = Calendar.getInstance();
        resultCal.setTime(result);
        assertEquals(2024, resultCal.get(Calendar.YEAR));
        assertEquals(Calendar.MARCH, resultCal.get(Calendar.MONTH));
        assertEquals(15, resultCal.get(Calendar.DAY_OF_MONTH));
    }

    @Test
    void getCellAsDate_stringCellWithIsoDate_returnsDate() {
        Cell cell = createCell(7, 0);
        cell.setCellValue("2024-01-15");

        Date result = ExcelUtils.getCellAsDate(cell);
        assertNotNull(result);
        // java.sql.Date returned by Date.valueOf() supports toLocalDate() directly
        LocalDate localResult = ((java.sql.Date) result).toLocalDate();
        assertEquals(LocalDate.of(2024, Month.JANUARY, 15), localResult);
    }

    @Test
    void getCellAsDate_stringCellWithInvalidDate_returnsNull() {
        Cell cell = createCell(8, 0);
        cell.setCellValue("not-a-date");
        assertNull(ExcelUtils.getCellAsDate(cell));
    }

    @Test
    void getCellAsDate_nonDateNumericCell_returnsNull() {
        Cell cell = createCell(9, 0);
        cell.setCellValue(12345.0);
        // No date format applied — isCellDateFormatted should return false
        assertNull(ExcelUtils.getCellAsDate(cell));
    }
}
