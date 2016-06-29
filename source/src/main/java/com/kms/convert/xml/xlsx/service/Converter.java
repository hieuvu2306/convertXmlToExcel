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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.kms.convert.xml.xlsx.model.TestCase;
import com.kms.convert.xml.xlsx.model.TestSuite;

public class Converter {

    private static final Logger LOG = LoggerFactory.getLogger(Converter.class);
    private static final String EXCEL_EXTENSION_NAME = ".xlsx";
    private static final String TESTCASES_SHEET_NAME = "Test Cases";
    private static final int HEADER_ROW_NUMBER = 0;
    private static final int HEADER_CELL_NUMBER = 0;
    private static final int DATA_ROW_NUMBER = 1;
    private static final int DATA_CELL_NUMBER = 0;
    private static final int REPORT_FILE_NAME_POSITION = 0;
    int rowNumber = 0;
    int cellNumber = 0;

    public void writeToXLSXFile(String xmlFilePath, String reportFilePath) {
        File file = new File(xmlFilePath);
        String excelFileName = "/" + file.getName().split(".xml")[REPORT_FILE_NAME_POSITION];

        TestSuite testSuite = convertXmlToTestSuite(xmlFilePath);
        List<TestSuite> testSuites = getTestSuites(testSuite);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet testCasesSheet = workbook.createSheet(TESTCASES_SHEET_NAME);

        createTestCasesSheet(testSuites, testCasesSheet);

        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(reportFilePath + excelFileName + EXCEL_EXTENSION_NAME);
            workbook.write(fileOut);
            LOG.info("Convert from xml file to excel file successfull!");
            workbook.close();
            fileOut.flush();
            fileOut.close();
        } catch (IOException e) {
            LOG.error("Wrong excel file path, error message : " + e.getMessage());
        }

    }

    private void createTestCasesSheet(List<TestSuite> testSuites, XSSFSheet testCasesSheet) {
        Map<String, Object[]> header = new TreeMap<String, Object[]>();
        header.put("header",
                new Object[] { "Internal Id", "Name", "Node Order", "Extermal Id", "Version", "Summary",
                        "Preconditions", "Execution Type", "Importance", "Key Words", "Test Suite Name", "Step Number",
                        "Step Action", "Expected Results" });

        Map<Long, Object[]> testCasesData = new TreeMap<Long, Object[]>();
        testCasesData = createTestCasesDataMap(testSuites);
        int headerRowNum = HEADER_ROW_NUMBER;
        Row headerRow = testCasesSheet.createRow(headerRowNum++);
        Object[] headerColumnNames = header.get("header");
        int headerCellNum = HEADER_CELL_NUMBER;
        for (Object obj : headerColumnNames) {
            Cell cell = headerRow.createCell(headerCellNum++);
            cell.setCellValue((String) obj);
        }

        Set<Long> keyset = testCasesData.keySet();
        int dataRowNum = DATA_ROW_NUMBER;
        for (Long keys : keyset) {
            Row row = testCasesSheet.createRow(dataRowNum++);
            Object[] objArr = testCasesData.get(keys);
            int dataCellNum = DATA_CELL_NUMBER;
            for (Object obj : objArr) {
                Cell cell = row.createCell(dataCellNum++);
                cell.setCellValue((String) obj);
            }
        }
    }

    private Map<Long, Object[]> createTestCasesDataMap(List<TestSuite> testSuites) {
        Map<Long, Object[]> data = new TreeMap<Long, Object[]>();
        Long dataKey = 1L;
        for (int i = 0; i < testSuites.size(); i++) {
            TestSuite testSuite = testSuites.get(i);
            for (int j = 0; j < testSuite.getTestCases().size(); j++) {
                TestCase testCase = testSuite.getTestCases().get(j);
                int stepInTestCaseCount = testCase.getSteps().size();
                if (stepInTestCaseCount > 0) {
                    for (int k = 0; k < stepInTestCaseCount; k++) {
                        data.put(dataKey++,
                                new Object[] { testCase.getInternalId().toString(), testCase.getName(),
                                        testCase.getNodeOrder().toString(), testCase.getExternalId().toString(),
                                        testCase.getVersion().toString(), testCase.getSummary(),
                                        testCase.getPreconditions(), testCase.getExecutionType().toString(),
                                        testCase.getImportance().toString(), getKeyWords(testCase), testSuite.getPath(),
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
                                    testCase.getImportance().toString(), getKeyWords(testCase), testSuite.getPath(), "",
                                    "", "" });
                }
            }
        }
        return data;
    }

    private String getKeyWords(TestCase testCase) {
        int keyWordsCount = testCase.getKeyWords().size();
        String keyWords = "";
        if (keyWordsCount > 0) {
            for (int i = 0; i < keyWordsCount; i++) {
                keyWords += testCase.getKeyWords().get(i).getName() + ",";
            }
        }
        return keyWords;
    }

    private List<TestSuite> getTestSuites(TestSuite testSuite) {
        List<TestSuite> testSuitesWithPath = createTestSuitesWithPath(testSuite);
        List<TestSuite> testSuites = new ArrayList<>();
        for (TestSuite testSuiteTemp : testSuitesWithPath) {
            if ((!testSuiteTemp.getTestCases().isEmpty() && !testSuiteTemp.getTestSuites().isEmpty())) {
                testSuites.add(testSuiteTemp);
            }
            if (testSuiteTemp.getTestSuites().isEmpty()) {
                testSuites.add(testSuiteTemp);
            }
        }
        return testSuites;
    }

    private List<TestSuite> createTestSuitesWithPath(TestSuite testSuite) {
        List<TestSuite> testSuites = new ArrayList<>();
        testSuite.setPath(testSuite.getName());
        testSuites.add(testSuite);
        testSuites.addAll(getChild(testSuite));
        return testSuites;
    }

    private List<TestSuite> getChild(TestSuite testSuite) {
        List<TestSuite> testSuites = new ArrayList<>();
        int testSuitesChildCount = testSuite.getTestSuites().size();
        if (testSuitesChildCount > 0) {
            for (TestSuite testSuiteChild : testSuite.getTestSuites()) {
                testSuiteChild.setPath(testSuite.getPath() + "\\" + testSuiteChild.getName());
                testSuites.add(testSuiteChild);
                testSuites.addAll(getChild(testSuiteChild));
            }
        }
        return testSuites;
    }

    public TestSuite convertXmlToTestSuite(String xmlFilePath) {
        File xmlFile = new File(xmlFilePath);
        JAXBContext jaxbContext;
        TestSuite testSuite = new TestSuite();
        try {
            jaxbContext = JAXBContext.newInstance(TestSuite.class);
            Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
            testSuite = (TestSuite) jaxbUnmarshaller.unmarshal(xmlFile);
        } catch (JAXBException e) {
            LOG.error("Can't parse from XML file to Object , error message : " + e.getMessage());
        }
        return testSuite;
    }
}
