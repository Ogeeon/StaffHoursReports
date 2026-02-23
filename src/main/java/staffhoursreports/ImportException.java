package staffhoursreports;

/**
 * Custom exception for import operations in StaffHoursReports.
 * Used to handle errors during Excel file processing and database operations.
 */
public class ImportException extends Exception {
    public ImportException(String message, Throwable cause) {
        super(message, cause);
    }
}
