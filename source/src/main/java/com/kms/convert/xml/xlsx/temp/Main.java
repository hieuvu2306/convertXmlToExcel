package com.kms.convert.xml.xlsx.temp;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.kms.convert.xml.xlsx.model.TestCase;
import com.kms.convert.xml.xlsx.model.TestSuite;
import com.kms.convert.xml.xlsx.service.Converter;

public class Main {

    public static void writeXLSXFile() throws IOException {

        String excelFileName = "D:/Test.xlsx";

        String sheetName = "TestCases";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        Map<String, Object[]> header = new TreeMap<String, Object[]>();
        Map<String, Employee> data = new TreeMap<String, Employee>();
        List<Employee> employees = new ArrayList<Employee>();
        employees.add(new Employee("Hoang", 23));
        employees.add(new Employee("Nguyen", 24));
        employees.add(new Employee("Hieu", 23));
        header.put("header", new Object[] { "name", "age" });
        for (int i = 0; i < employees.size(); i++) {
            data.put(String.valueOf(i), employees.get(i));
        }
        int headerRowNum = 0;
        Row headerRow = sheet.createRow(headerRowNum++);
        Object[] headerColumnNames = header.get("header");
        int headerCellNum = 0;
        for (Object obj : headerColumnNames) {
            Cell cell = headerRow.createCell(headerCellNum++);
            cell.setCellValue((String) obj);
        }

        Set<String> keyset = data.keySet();
        int dataRowNum = 1;
        for (String key : keyset) {
            Row row = sheet.createRow(dataRowNum++);
            Employee employee = data.get(key);
            for (int i = 0; i < headerColumnNames.length; i++) {
                if (i == 0) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(employee.getName());
                } else {
                    Cell cell2 = row.createCell(i);
                    cell2.setCellValue(employee.getAge());
                }
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        // write this workbook to an Outputstream.
        wb.write(fileOut);
        wb.close();
        fileOut.flush();
        fileOut.close();
    }

    public static void writeTestSuiteToXLSXFile(List<TestSuite> testSuiteSamples) throws IOException {

        String excelFileName = "D:/Report.xlsx";

        String sheetName = "TestCases";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        Map<String, Object[]> header = new TreeMap<String, Object[]>();
        Map<Long, Object[]> data = new TreeMap<Long, Object[]>();
        header.put("header",
                new Object[] { "Internal Id", "Name", "Node Order", "Extermal Id", "Version", "Summary",
                        "Preconditions", "Execution Type", "Importance", "Test Suite Name", "Step Number",
                        "Step Action", "Expected Results" });

        int countTestCase = 0;
        int countTestCaseHaveMoreThanOneTestStep = 0;
        Long key = 1L;
        for (int i = 0; i < testSuiteSamples.size(); i++) {
            TestSuite testSuiteSimple = testSuiteSamples.get(i);
            for (int j = 0; j < testSuiteSimple.getTestCases().size(); j++) {
                TestCase testCase = testSuiteSimple.getTestCases().get(j);
                int stepInTestCaseCount = testCase.getSteps().size();
                if (stepInTestCaseCount >= 1) {
                    for (int k = 0; k < stepInTestCaseCount; k++) {
                        data.put(key++,
                                new Object[] { testCase.getInternalId().toString(), testCase.getName(),
                                        testCase.getNodeOrder().toString(), testCase.getExternalId().toString(),
                                        testCase.getVersion().toString(), testCase.getSummary(),
                                        testCase.getPreconditions(), testCase.getExecutionType().toString(),
                                        testCase.getImportance().toString(), testSuiteSimple.getName(),
                                        testCase.getSteps().get(k).getStepNumber().toString(),
                                        testCase.getSteps().get(k).getActions(),
                                        testCase.getSteps().get(k).getExpectedResults() });
                    }
                } else {
                    data.put(key++,
                            new Object[] { testCase.getInternalId().toString(), testCase.getName(),
                                    testCase.getNodeOrder().toString(), testCase.getExternalId().toString(),
                                    testCase.getVersion().toString(), testCase.getSummary(),
                                    testCase.getPreconditions(), testCase.getExecutionType().toString(),
                                    testCase.getImportance().toString(), testSuiteSimple.getName(), "", "", "" });
                }
                if (testCase.getSteps().size() > 1) {
                    countTestCaseHaveMoreThanOneTestStep += 1;
                }
                countTestCase += 1;
            }
        }
        System.out.println("count testCase have more than one test step" + countTestCaseHaveMoreThanOneTestStep);
        System.out.println(" count test case" + countTestCase);
        System.out.println(data.size() + " count test case data map");

        int headerRowNum = 0;
        Row headerRow = sheet.createRow(headerRowNum++);
        Object[] headerColumnNames = header.get("header");
        int headerCellNum = 0;
        for (Object obj : headerColumnNames) {
            Cell cell = headerRow.createCell(headerCellNum++);
            cell.setCellValue((String) obj);
        }

        Set<Long> keyset = data.keySet();
        int dataRowNum = 1;
        for (Long keys : keyset) {
            Row row = sheet.createRow(dataRowNum++);
            Object[] objArr = data.get(keys);
            int dataCellNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(dataCellNum++);
                cell.setCellValue((String) obj);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        // write this workbook to an Outputstream.
        wb.write(fileOut);
        wb.close();
        fileOut.flush();
        fileOut.close();
    }

    public static void main(String[] args) {
        Converter converter = new Converter();
        TestSuite testSuite = converter.convertXmlToTestSuite("D:\\report.xml");
        List<TestSuite> testSuiteSimples = new ArrayList<TestSuite>();
        testSuiteSimples = converter.getTestSuiteSimples(testSuite);
        try {
            // writeXLSXFile();
            writeTestSuiteToXLSXFile(testSuiteSimples);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
