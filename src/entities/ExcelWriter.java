package entities;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

/** A classe ExcelWriter contém um vetor com as colunas,
 *  Uma lista de Employee e as instancias dos employees
 *  Seguidos de sua adição na lista.
 *  Detalhe para o Calendar.getInstance() que mostra
 *  como instanciar uma data pré-estabelecida.
 */

public class ExcelWriter {
    private static String[] columns = {"Name", "Email", "Date Of Birth", "Salary"};
    private static List<Employee> employees =  new ArrayList<>();

    public static String[] getColumns() {
        return columns;
    }

    public static void setColumns(String[] columns) {
        ExcelWriter.columns = columns;
    }

    public static List<Employee> getEmployees() {
        return employees;
    }

    public static void setEmployees(List<Employee> employees) {
        ExcelWriter.employees = employees;
    }

    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Antonio Carlos", "antonio@example.com",
                dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Alberto Silva", "alberto@example.com",
                dateOfBirth.getTime(), 1500000.0));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("João Paulo", "joao@example.com",
                dateOfBirth.getTime(), 1800000.0));

    }
}
