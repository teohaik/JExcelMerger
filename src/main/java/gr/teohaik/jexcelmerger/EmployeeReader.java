package gr.teohaik.jexcelmerger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeReader {

    static List<String> extremeCases = new ArrayList();

    static DataFormatter formater = new DataFormatter();

    public static Employee readEmployee(String FILE_NAME) throws FileNotFoundException, IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Employee emp = new Employee();
        DateFormat df = DateFormat.getDateInstance(DateFormat.SHORT, Locale.forLanguageTag("el-GR"));

        DataFormatter formater = new DataFormatter();

        for (int i = 6; i <= 17; i++) {
            Row currentRow = datatypeSheet.getRow(i);
            Cell currentCell = currentRow.getCell(0);
            String cv = currentCell.getStringCellValue();

            if (cv.contains("ΑΜ ΑΤΕ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setAem(formater.formatCellValue(nextCell));
            } else if (cv.contains("ΕΠΩΝΥΜΟ")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setSurname(formater.formatCellValue(nextCell));
            } else if (cv.equals("ΟΝΟΜΑ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setName(formater.formatCellValue(nextCell));
            } else if (cv.contains("ΟΝΟΜΑ ΠΑΤΡΟΣ/ΣΥΖΥΓΟΥ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setFatherOrHusbandName(formater.formatCellValue(nextCell));
            } else if (cv.equals("ΗΜΕΡ.ΓΕΝΝΗΣΗΣ")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                Date dateCellValue = nextCell.getDateCellValue();
                try {

                    String formattedDate = df.format(dateCellValue);
                    emp.setBirthdate(formattedDate);
                } catch (Exception e) {
                    emp.setBirthdate("");

                }
            } else if (cv.contains("ΑΦΜ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setAfm(formater.formatCellValue(nextCell));
            } else if (cv.contains("ΑΜΚΑ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setAmka(formater.formatCellValue(nextCell));
            } else if (cv.contains("ΔΙΕΥΘΥΝΣΗ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setAddress(formater.formatCellValue(nextCell));
            } else if (cv.contains("ΤΗΛΕΦΩΝΑ ΕΠΙΚΟΙΝΩΝΙΑΣ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                String tel1 = formater.formatCellValue(nextCell);

                if (tel1.contains(",")) {
                    String[] tels = tel1.split(",");
                    emp.setTelephone1(tels[0]);
                    emp.setTelephone2(tels[1]);
                } else {
                    emp.setTelephone1(tel1);
                }

                int columnIndex = currentCell.getColumnIndex();
                Cell nextCell2 = currentRow.getCell(columnIndex + 4);
                String nextCell2Value = formater.formatCellValue(nextCell2);
                if (nextCell2Value != null && !nextCell2Value.isEmpty()) {
                    emp.setTelephone2(formater.formatCellValue(nextCell2));
                }

            } else if (cv.contains("ΤΡΑΠΕΖΙΚΟΣ ΛΟΓ/ΣΜΟΣ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setBankAccount(formater.formatCellValue(nextCell));
            } else if (cv.contains("ΤΡΑΠΕΖΑ:")) {
                Cell nextCell = getNextCell(currentCell, currentRow);
                emp.setBank(formater.formatCellValue(nextCell));
            }
        }

        try {
            Row leaveDateRow = datatypeSheet.getRow(22);

            LocalDate leaveDate = getDateFromRow(leaveDateRow);

            Row hireDateRow = datatypeSheet.getRow(24);

            LocalDate hireDate = getDateFromRow(hireDateRow);

            if (leaveDate != null) {
                emp.setLeaveDate(leaveDate.toString());
            }

            if (hireDate != null) {
                emp.setHireDate(hireDate.toString());
            }

            Row monthsInsuredRow = datatypeSheet.getRow(36);

            Cell monthsInsuredCell = monthsInsuredRow.getCell(1);

            int monthsInsured = (int) monthsInsuredCell.getNumericCellValue();

            emp.setMonthsInsured(monthsInsured);

            double average3Years = getNumericValueFrom(40, 1, datatypeSheet);
            double average5Years = getNumericValueFrom(41, 1, datatypeSheet);
            emp.setAverage3Years(average3Years);
            emp.setAverage5Years(average5Years);

            int strikeD = (int) getNumericValueFrom(29, 1, datatypeSheet);
            int strikeM = (int) getNumericValueFrom(29, 2, datatypeSheet);
            int strikeY = (int) getNumericValueFrom(29, 3, datatypeSheet);

            emp.setStrikeDays(strikeD);
            emp.setStrikeMonths(strikeM);
            emp.setStrikeYears(strikeY);

            double efapax = getNumericValueFrom(48, 1, datatypeSheet);

            emp.setEfapax(efapax);

            double efapaxTaken = getNumericValueFrom(55, 6, datatypeSheet);

            double threshold = 0.7 * efapax;

            if (efapaxTaken < threshold) {
                efapaxTaken = getNumericValueFrom(56, 6, datatypeSheet);
            }

            emp.setEfapaxTaken(efapaxTaken);

            //System.out.println("efapaxTaken= " + efapaxTaken);
        } catch (Exception ex) {
            extremeCases.add(emp.aem);
            return emp;
        }
        return emp;
    }

    public static Employee readEmployee2(String FILE_NAME) throws FileNotFoundException, IOException {

        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));

        Employee emp = new Employee();
        DateFormat df = DateFormat.getDateInstance(DateFormat.SHORT, Locale.forLanguageTag("el-GR"));
        try {
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);

            emp.aem = getStringValueFrom(1, 12, datatypeSheet);
            emp.surname = getStringValueFrom(6, 2, datatypeSheet);

            emp.name = getStringValueFrom(6, 10, datatypeSheet);
            emp.fatherOrHusbandName = getStringValueFrom(7, 2, datatypeSheet);
            emp.birthdate = getStringValueFrom(7, 7, datatypeSheet);
            String ad1 = getStringValueFrom(8, 2, datatypeSheet);
            String ad2 = getStringValueFrom(8, 7, datatypeSheet);
            String ad3 = getStringValueFrom(8, 10, datatypeSheet);
            String ad4 = getStringValueFrom(8, 13, datatypeSheet);
            emp.address = ad1 + " " + ad2 + ", " + ad3 + " " + ad4;
            String tels = getStringValueFrom(9, 4, datatypeSheet);
            if (tels.contains("-")) {
                String[] telss = tels.split("-");
                emp.telephone1 = telss[0];
                emp.telephone2 = telss[1];
            } else {
                emp.telephone1 = tels;
            }
            emp.afm = getStringValueFrom(10, 1, datatypeSheet);

            int leaveDay = (int) getNumericValueFrom(11, 11, datatypeSheet);
            int leaveMonth = (int) getNumericValueFrom(11, 12, datatypeSheet);
            int leaveYear = (int) getNumericValueFrom(11, 13, datatypeSheet);

            LocalDate leaveDate = LocalDate.of(leaveYear, leaveMonth, leaveDay);

            int hireDay = (int) getNumericValueFrom(12, 11, datatypeSheet);
            int hireMonth = (int) getNumericValueFrom(12, 12, datatypeSheet);
            int hireYear = (int) getNumericValueFrom(12, 13, datatypeSheet);

            LocalDate hireDate = LocalDate.of(hireYear, hireMonth, hireDay);

            if (leaveDate != null) {
                emp.setLeaveDate(leaveDate.toString());
            }

            if (hireDate != null) {
                emp.setHireDate(hireDate.toString());
            }

            int monthsInsured = (int) getNumericValueFrom(30, 13, datatypeSheet);
            emp.setMonthsInsured(monthsInsured);
//
//            int strikeD = (int) getNumericValueFrom(29, 1, datatypeSheet);
//            int strikeM = (int) getNumericValueFrom(29, 2, datatypeSheet);
//            int strikeY = (int) getNumericValueFrom(29, 3, datatypeSheet);
//
//            emp.setStrikeDays(strikeD);
//            emp.setStrikeMonths(strikeM);
//            emp.setStrikeYears(strikeY);
//
//            double efapax = getNumericValueFrom(48, 1, datatypeSheet);
//
//            emp.setEfapax(efapax);
//
//            double efapaxTaken = getNumericValueFrom(55, 6, datatypeSheet);
//
//            double threshold = 0.7 * efapax;
//
//            if (efapaxTaken < threshold) {
//                efapaxTaken = getNumericValueFrom(56, 6, datatypeSheet);
//            }
//
//            emp.setEfapaxTaken(efapaxTaken);

            //System.out.println("efapaxTaken= " + efapaxTaken);
            for (int i = 34; i <= 46; i++) {
                Row currentRow = datatypeSheet.getRow(i);
                for (int c = 4; c <= 14; c++) {
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

        } catch (Exception ex) {
            extremeCases.add(emp.aem);
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
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(colNum);
        return cell.getNumericCellValue();
    }

    public static String getStringValueFrom(int rowNum, int colNum, Sheet sheet) {
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(colNum);
        return formater.formatCellValue(cell);
    }

}
