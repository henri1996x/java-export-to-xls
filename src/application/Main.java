package application;

import entities.Employee;
import entities.ExcelWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Na classe principal é onde está o "motor" da aplicação.
 * Cada passo está comentado com seu objetivo.
 */

public class Main {
    public static void main(String[] args) {

        /** Um objeto Workbook é criado. */
        Workbook workbook = new XSSFWorkbook();

        /** CreationHelper nos ajuda a criar instâncias de várias coisas, como DataFormat,
         Hiperlink, RichTextString etc; Em um formato independente (HSSF, XSSF)
         */
        CreationHelper createHelper = workbook.getCreationHelper();

        /** Criação da planilha */
        Sheet sheet = workbook.createSheet("Employee");

        /** Criação da fonte para as colunas */
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        /** Criação da celula com a fonte */
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        /** Criação da linha */
        Row headerRow = sheet.createRow(0);

        /** Criação da celula */
        for (int i = 0; i < ExcelWriter.getColumns().length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(ExcelWriter.getColumns()[i]);
            cell.setCellStyle(headerCellStyle);
        }

        /** Criação do estilo de célula para formatar a data */
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));


        /** Criação de outras linhas com dados dos employees */
        int rowNum = 1;
        for (Employee employee : ExcelWriter.getEmployees()) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(employee.getName());

            row.createCell(1)
                    .setCellValue(employee.getEmail());

            Cell dateOfBirthCell = row.createCell(2);
            dateOfBirthCell.setCellValue(employee.getDateOfBirth());
            dateOfBirthCell.setCellStyle(dateCellStyle);

            row.createCell(3)
                    .setCellValue(employee.getSalary());
        }

        /** Redimensiona todas as colunas para caber no tamanho do conteúdo */
        for (int i = 0; i < ExcelWriter.getColumns().length; i++) {
            sheet.autoSizeColumn(i);
        }

        /** Escreve a saída em um arquivo */
        try {
            FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Arquivo exportado com sucesso!");
        } catch (IOException e){
            System.out.println(e.getMessage());
        }
        finally {
            try {
                /** Fecha o workbook */
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
}
