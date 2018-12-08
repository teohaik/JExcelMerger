package gr.teohaik.jexcelmerger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.time.LocalDate;

import java.util.ArrayList;

import java.util.List;
import java.util.Locale;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeReaderEpistrofes {

    static List<String> extremeCases = new ArrayList();

    static DataFormatter formater = new DataFormatter();

    public static Employee readEmployee(String FILE_NAME) throws FileNotFoundException, IOException, InvalidFormatException {

        //FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        OPCPackage pkg = OPCPackage.open(new File(FILE_NAME));

        Employee emp = new Employee();
        DateFormat df = DateFormat.getDateInstance(DateFormat.SHORT, Locale.forLanguageTag("el-GR"));
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            emp.aem = getStringValueFrom(6, 1, datatypeSheet);
            emp.surname = getStringValueFrom(7, 1, datatypeSheet);

            emp.name = getStringValueFrom(8, 1, datatypeSheet);
            emp.fatherOrHusbandName = getStringValueFrom(9, 1, datatypeSheet);
            emp.birthdate = getStringValueFrom(11, 1, datatypeSheet);
            emp.afm = getStringValueFrom(12, 1, datatypeSheet);
            emp.amka = getStringValueFrom(13, 1, datatypeSheet);
            emp.address = getStringValueFrom(14, 1, datatypeSheet);

            emp.telephone1 = getStringValueFrom(15, 1, datatypeSheet);
            emp.telephone1 =emp.telephone1.replaceAll(",", "-");
            emp.bankAccount = getStringValueFrom(16, 1, datatypeSheet);
            emp.bank = getStringValueFrom(17, 1, datatypeSheet);

            int leaveDay = (int) getNumericValueFrom(22, 1, datatypeSheet);
            int leaveMonth = (int) getNumericValueFrom(22, 2, datatypeSheet);
            int leaveYear = (int) getNumericValueFrom(22, 3, datatypeSheet);

            LocalDate leaveDate = LocalDate.of(leaveYear, leaveMonth, leaveDay);

            int hireDay = (int) getNumericValueFrom(24, 1, datatypeSheet);
            int hireMonth = (int) getNumericValueFrom(24, 2, datatypeSheet);
            int hireYear = (int) getNumericValueFrom(24, 3, datatypeSheet);

            LocalDate hireDate = LocalDate.of(hireYear, hireMonth, hireDay);

            if (leaveDate != null) {
                emp.setLeaveDate(leaveDate.toString());
            }

            if (hireDate != null) {
                emp.setHireDate(hireDate.toString());
            }

            int monthsInsured = (int) getNumericValueFrom(30, 13, datatypeSheet);
            emp.setMonthsInsured(monthsInsured);

            emp.strikeDays = (int) getNumericValueFrom(29, 1, datatypeSheet);
            emp.strikeMonths = (int) getNumericValueFrom(29, 2, datatypeSheet);
            emp.strikeYears = (int) getNumericValueFrom(29, 3, datatypeSheet);

            emp.parentalLeaveDays = (int) getNumericValueFrom(30, 1, datatypeSheet);
            emp.parentalLeaveMonths = (int) getNumericValueFrom(30, 2, datatypeSheet);
            emp.parentalLeaveYears = (int) getNumericValueFrom(30, 3, datatypeSheet);

            emp.unpaidLeaveDays = (int) getNumericValueFrom(31, 1, datatypeSheet);
            emp.unpaidLeaveMonths = (int) getNumericValueFrom(31, 2, datatypeSheet);
            emp.unpaidLeaveYears = (int) getNumericValueFrom(31, 3, datatypeSheet);

            emp.unjustifiedAbsenceDays = (int) getNumericValueFrom(32, 1, datatypeSheet);
            emp.unjustifiedAbsenceMonths = (int) getNumericValueFrom(32, 2, datatypeSheet);
            emp.unjustifiedAbsenceYears = (int) getNumericValueFrom(32, 3, datatypeSheet);

            emp.pauseDays = (int) getNumericValueFrom(33, 1, datatypeSheet);
            emp.pauseMonths = (int) getNumericValueFrom(33, 2, datatypeSheet);
            emp.pauseYears = (int) getNumericValueFrom(33, 3, datatypeSheet);

            emp.praxiAponomis = getStringValueFrom(41, 1, datatypeSheet);

            for (int i = 40; i <= 46; i++) {
                if (emp.average3Years > 0 && emp.average5Years > 0) {
                    break;
                }

                Row currentRow = datatypeSheet.getRow(i);
                for (int c = 0; c <= 4; c++) {
                    Cell currentCell = currentRow.getCell(c);
                    String cv = currentCell.getStringCellValue();

                    if (cv.contains("ΤΡΙΕΤΙΑΣ")) {
                        double value = 0;
                        int k = c;
                        do {

                            try {
                                value = getNumericValueFrom(currentRow.getRowNum(), k, datatypeSheet);
                            } catch (Exception e) {
                            }

                            k++;
                        } while (value <= 0 && k < 30);
                        emp.setAverage3Years(value);
                        break;
                    } else if (cv.contains("ΠΕΝΤΑΕΤΙΑΣ")) {
                        double value = 0;
                        int k = c;
                        do {
                            try {
                                value = getNumericValueFrom(currentRow.getRowNum(), k, datatypeSheet);
                            } catch (Exception e) {
                            }
                            k++;
                        } while (value <= 0 && k < 30);
                        emp.setAverage5Years(value);
                        break;
                    }
                }

            }
            if (emp.average3Years == 0 && emp.average5Years == 0) {
                extremeCases.add(emp.aem);
            };

        } catch (Exception ex) {
            extremeCases.add(emp.aem);
            ex.printStackTrace();
            return emp;
        }

        return emp;
    }

    public static void printExtremeCases() {
        System.out.println("=============================================");
        System.out.println("Extreme cases");
        System.out.println("=============================================");
        for (String s : extremeCases) {
            System.out.println(s);
        }
        System.out.println("=============================================");
    }

    public static LocalDate getDateFromRow(Row leaveDateRow) {
        Cell dayC = leaveDateRow.getCell(1);
        int day = (int) dayC.getNumericCellValue();

        Cell monthC = leaveDateRow.getCell(2);
        int month = (int) monthC.getNumericCellValue();

        Cell yearC = leaveDateRow.getCell(3);
        int year = (int) yearC.getNumericCellValue();

        try {
            return LocalDate.of(year, month, day);
        } catch (java.time.DateTimeException dte) {
            return null;
        }
    }

    public static Cell getNextCell(Cell currentCell, Row row) {
        int columnIndex = currentCell.getColumnIndex();
        return row.getCell(columnIndex + 1);
    }

    public static String getNextCellValue(Cell currentCell, Row row) {
        int columnIndex = currentCell.getColumnIndex();
        return row.getCell(columnIndex + 1).getStringCellValue();
    }

    public static double getNumericValueFrom(int rowNum, int colNum, Sheet sheet) {
        try {
            Row row = sheet.getRow(rowNum);
            Cell cell = row.getCell(colNum);
            return cell.getNumericCellValue();
        } catch (Exception e) {
            return 0;
        }
    }

    public static String getStringValueFrom(int rowNum, int colNum, Sheet sheet) {
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(colNum);
        return formater.formatCellValue(cell);
    }

}
