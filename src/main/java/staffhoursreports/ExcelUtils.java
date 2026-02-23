package staffhoursreports;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.util.Date;

public class ExcelUtils {

    private ExcelUtils() {
        // Utility class - prevent instantiation
    }

    /**
     * Безопасно читает значение ячейки как строку, обрабатывая разные типы данных
     */
    public static String getCellAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // Убираем дробную часть, если число целое
                    double numValue = cell.getNumericCellValue();
                    if (numValue == (long) numValue) {
                        return String.valueOf((long) numValue);
                    } else {
                        return String.valueOf(numValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue().trim();
                } catch (Exception e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                return "";
        }
    }

    /**
     * Безопасно читает значение ячейки как дату
     */
    public static Date getCellAsDate(Cell cell) {
        if (cell == null) {
            return null;
        }
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue();
        }
        if (cell.getCellType() == CellType.STRING) {
            String strValue = cell.getStringCellValue().trim();
            if (!strValue.isEmpty()) {
                // Попытка распарсить строку как дату
                try {
                    return java.sql.Date.valueOf(strValue);
                } catch (IllegalArgumentException e) {
                    // Не удалось распарсить как дату
                    return null;
                }
            }
        }
        return null;
    }
}
