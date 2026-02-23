package staffhoursreports;

import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.time.Month;
import java.util.Date;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

class UtilsTest {

    @Test
    void toLocalDate_convertsCorrectly() {
        // 2024-03-15 as a java.util.Date
        LocalDate expected = LocalDate.of(2024, Month.MARCH, 15);
        Date input = Utils.fromLocalDate(expected);
        LocalDate result = Utils.toLocalDate(input);
        assertEquals(expected, result);
    }

    @Test
    void toLocalDate_nullReturnsNull() {
        assertNull(Utils.toLocalDate(null));
    }

    @Test
    void fromLocalDate_convertsCorrectly() {
        LocalDate input = LocalDate.of(2024, Month.JUNE, 1);
        Date result = Utils.fromLocalDate(input);
        assertNotNull(result);
        // Round-trip check
        assertEquals(input, Utils.toLocalDate(result));
    }

    @Test
    void fromLocalDate_nullReturnsNull() {
        assertNull(Utils.fromLocalDate(null));
    }

    @Test
    void roundTrip_fromLocalDateAndBack() {
        LocalDate original = LocalDate.of(2023, Month.DECEMBER, 31);
        Date converted = Utils.fromLocalDate(original);
        LocalDate restored = Utils.toLocalDate(converted);
        assertEquals(original, restored);
    }

    @Test
    void localizeDate_russianLocale() {
        LocalDate date = LocalDate.of(2024, Month.JANUARY, 15);
        String result = Utils.localizeDate(date, Locale.forLanguageTag("ru"));
        // Russian long date should contain the month name "января" and year
        assertTrue(result.contains("2024"), "Result should contain the year: " + result);
        assertTrue(result.contains("15"), "Result should contain the day: " + result);
    }

    @Test
    void localizeDate_englishLocale() {
        LocalDate date = LocalDate.of(2024, Month.MARCH, 5);
        String result = Utils.localizeDate(date, Locale.ENGLISH);
        assertTrue(result.contains("2024"), "Result should contain the year: " + result);
        assertTrue(result.contains("March"), "Result should contain 'March': " + result);
    }

    @Test
    void getReportName_containsDateRange() {
        LocalDate from = LocalDate.of(2024, Month.JANUARY, 16);
        LocalDate to = LocalDate.of(2024, Month.FEBRUARY, 15);
        String result = Utils.getReportName(from, to);
        assertTrue(result.startsWith("Отчет по ЗИ за период с "), "Should start with expected prefix: " + result);
        assertTrue(result.endsWith(".xlsx"), "Should end with .xlsx: " + result);
        assertTrue(result.contains(" по "), "Should contain ' по ': " + result);
    }
}
