package com.kms.convert.xml.xlsx.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.kms.convert.xml.xlsx.model.TestCase;
import com.kms.convert.xml.xlsx.model.TestSuite;

public class Converter {

    private static final String TESTCASES_SHEET_NAME = "TestCases";
    private static final int HEADER_ROW_NUMBER = 0;
    private static final int HEADER_CELL_NUMBER = 0;
    private static final int DATA_ROW_NUMBER = 1;
    private static final int DATA_CELL_NUMBER = 0;

    public void writeToXLSXFile(String xmlFilePath, String reportFilePath) throws IOException {

        TestSuite testSuite = convertXmlToTestSuite(xmlFilePath);
        List<TestSuite> testSuiteSamples = getTestSuiteSimples(testSuite);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(TESTCASES_SHEET_NAME);

        Map<String, Object[]> header = new TreeMap<String, Object[]>();
        Map<Long, Object[]> data = new TreeMap<Long, Object[]>();
        header.put("header",
                new Object[] { "Internal Id", "Name", "Node Order", "Extermal Id", "Version", "Summary",
                        "Preconditions", "Execution Type", "Importance", "Test Suite Name", "Step Number",
                        "Step Action", "Expected Results" });

        Long dataKey = 1L;
        for (int i = 0; i < testSuiteSamples.size(); i++) {
            TestSuite testSuiteSimple = testSuiteSamples.get(i);
            for (int j = 0; j < testSuiteSimple.getTestCases().size(); j++) {
                TestCase testCase = testSuiteSimple.getTestCases().get(j);
                int stepInTestCaseCount = testCase.getSteps().size();
                if (stepInTestCaseCount >= 1) {
                    for (int k = 0; k < stepInTestCaseCount; k++) {
                        data.put(dataKey++,
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
                    data.put(dataKey++,
                            new Object[] { testCase.getInternalId().toString(), testCase.getName(),
                                    testCase.getNodeOrder().toString(), testCase.getExternalId().toString(),
                                    testCase.getVersion().toString(), testCase.getSummary(),
                                    testCase.getPreconditions(), testCase.getExecutionType().toString(),
                                    testCase.getImportance().toString(), testSuiteSimple.getName(), "", "", "" });
                }
            }
        }
        int headerRowNum = HEADER_ROW_NUMBER;
        Row headerRow = sheet.createRow(headerRowNum++);
        Object[] headerColumnNames = header.get("header");
        int headerCellNum = HEADER_CELL_NUMBER;
        for (Object obj : headerColumnNames) {
            Cell cell = headerRow.createCell(headerCellNum++);
            cell.setCellValue((String) obj);
        }

        Set<Long> keyset = data.keySet();
        int dataRowNum = DATA_ROW_NUMBER;
        for (Long keys : keyset) {
            Row row = sheet.createRow(dataRowNum++);
            Object[] objArr = data.get(keys);
            int dataCellNum = DATA_CELL_NUMBER;
            for (Object obj : objArr) {
                Cell cell = row.createCell(dataCellNum++);
                cell.setCellValue((String) obj);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(reportFilePath);

        workbook.write(fileOut);
        workbook.close();
        fileOut.flush();
        fileOut.close();
    }

    /**
     * This method to get testSuite simples have only testCases and name
     */
    public List<TestSuite> getTestSuiteSimples(TestSuite testSuite) {
        List<TestSuite> testSuiteSimples = new ArrayList<TestSuite>();
        if (testSuite.getTestSuites().isEmpty()) {
            testSuiteSimples.add(testSuite);
            return testSuiteSimples;
        }
        for (TestSuite testSuiteTemp : testSuite.getTestSuites()) {
            testSuiteSimples.addAll(getTestSuiteSimples(testSuiteTemp));
        }
        return testSuiteSimples;
    }

    public TestSuite convertXmlToTestSuite(String filePath) {
        File xmlFile = new File(filePath);
        JAXBContext jaxbContext;
        TestSuite testSuite = new TestSuite();
        try {
            jaxbContext = JAXBContext.newInstance(TestSuite.class);
            Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
            testSuite = (TestSuite) jaxbUnmarshaller.unmarshal(xmlFile);
        } catch (JAXBException e) {
            e.printStackTrace();
        }
        return testSuite;
    }
}
