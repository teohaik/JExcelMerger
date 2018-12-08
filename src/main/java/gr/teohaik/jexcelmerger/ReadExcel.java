package gr.teohaik.jexcelmerger;

/**
 *
 * @author Theodore Chaikalis
 */
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {

    private static final String FILE_NAME = "files/1800/86878.xlsx";

    //private static final String fileDir = "files/1800";
    private static final String FILE_DIR = "files/epistrofes";

    private static final String DELIMITER = ";";

    public static void main(String[] args) throws NoSuchFieldException, IllegalArgumentException, IllegalAccessException {

        List<Employee> employees = new ArrayList();

        try {

            //Employee emp = readEmployee(FILE_NAME);
            //       System.out.println(emp);
            File dir = new File(FILE_DIR);
            File[] directoryListing = dir.listFiles();
            if (directoryListing != null) {
                int i = 1;
                for (File child : directoryListing) {
                    if (!child.getAbsolutePath().endsWith("xlsx")) {
                        continue;
                    }
                    System.out.println(child.getName());
                    Employee emp = EmployeeReaderEpistrofes.readEmployee(child.getAbsolutePath());
                    employees.add(emp);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {

            Field[] fields = Employee.class.getDeclaredFields();

            try (FileWriter fw = new FileWriter("out.csv", true);
                    BufferedWriter bw = new BufferedWriter(fw);
                    PrintWriter out = new PrintWriter(bw)) {

                for (Field f : fields) {
                    System.out.print(f.getName() + DELIMITER);
                    out.print(f.getName() + DELIMITER);
                }
                System.out.println();
                out.println();

                String praxiAponomis;

                for (Employee e : employees) {
                    for (Field f : fields) {
                        Object value = e.getClass().getDeclaredField(f.getName()).get(e);
                        String outputValue = "null";

                        if (value != null) {
                            outputValue = String.valueOf(value);
                            outputValue = outputValue.trim();
                            outputValue = outputValue.replaceAll("\\n", "");
                            outputValue = outputValue.replaceAll("\\r", "");
                        }

                        System.out.print(outputValue + DELIMITER);
                        out.print(outputValue + DELIMITER);
                    }
                    System.out.println();
                    out.println();

                }

            } catch (IOException e) {
                //exception handling left as an exercise for the reader
            }
        }

        EmployeeReaderEpistrofes.printExtremeCases();
    }

}
