package com.goodoldpepega.esrconstructor5.test;

import com.goodoldpepega.esrconstructor5.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Multiple functional test. Also these test can be used as examples
 * If includeTestsThatWillCrashWorkbookCreation is true, file will not be created
 */
public class FuncTest {
   private static ESRConstructor5 esrConstructor5;
   private static HSSFCellStyle customCellStyle;
   private static boolean includeTestsThatWillCrashWorkbookCreation = true; // include tests that should crash workbook creation (for example tests that should exceed limits)
   private static File excelFile = new File("D:/Test/funcTestReport.xls");

    public static void main(String[] args) throws IOException {
        startTest(true);
    }

    private static void startTest(boolean createExcelFileOnDisc) throws IOException {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        initCustomStyles(hssfWorkbook);
        esrConstructor5 = new ESRConstructor5(hssfWorkbook);
        tryToPutReports(esrConstructor5, includeTestsThatWillCrashWorkbookCreation);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        hssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    private static void initCustomStyles(HSSFWorkbook hssfWorkbook){
        customCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont customFont = hssfWorkbook.createFont();
        customFont.setBoldweight((short)700);
        customCellStyle.setFont(customFont);
        customCellStyle.setAlignment((short)2);
        customCellStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        customCellStyle.setFillPattern((short)1);
        customCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        customCellStyle.setWrapText(true);
        customCellStyle.setBorderBottom((short)2);
        customCellStyle.setBorderTop((short)2);
        customCellStyle.setBorderLeft((short)2);
        customCellStyle.setBorderRight((short)2);
    }

    /**
     *
     * @param esrConstructor5
     * @param includeTestsThatWillCrashWorkbookCreation if true, file would not be created
     */
    private static void tryToPutReports(ESRConstructor5 esrConstructor5, boolean includeTestsThatWillCrashWorkbookCreation){
        try{
            esrConstructor5.writeToExcel(testReport1(), "testReport1", false);
            System.out.println("testReport1 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport2(), "testReport2", false);
            System.out.println("testReport2 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport3(), "testReport3", false);
            System.out.println("testReport3 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport4(), "testReport4", false);
            System.out.println("testReport4 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport5(), "testReport5", false);
            System.out.println("testReport5 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport6(), "testReport6", false);
            System.out.println("testReport6 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport7(), "testReport7", false);
            System.out.println("testReport7 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport8(), "testReport8", false);
            System.out.println("testReport8 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport9(), "testReport9", false);
            System.out.println("testReport9 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport10(), "testReport10", false);
            System.out.println("testReport10 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport11(), "testReport11", false);
            System.out.println("testReport11 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport12(), "testReport12", false);
            System.out.println("testReport12 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport1(), "testReport13", true);
            System.out.println("testReport13 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport2(), "testReport14", true);
            System.out.println("testReport14 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport3(), "testReport15", true);
            System.out.println("testReport15 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport4(), "testReport16", true);
            System.out.println("testReport16 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport5(), "testReport17", true);
            System.out.println("testReport17 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport6(), "testReport18", true);
            System.out.println("testReport18 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport7(), "testReport19", true);
            System.out.println("testReport19 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport8(), "testReport20", true);
            System.out.println("testReport20 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport9(), "testReport21", true);
            System.out.println("testReport21 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport10(), "testReport22", true);
            System.out.println("testReport22 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport11(), "testReport23", true);
            System.out.println("testReport23 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport12(), "testReport24", true);
            System.out.println("testReport24 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport13(), "testReport25", true);
            System.out.println("testReport25 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport14(), "testReport26", true);
            System.out.println("testReport26 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport15(), "testReport27", true);
            System.out.println("testReport27 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport16(), "testReport28", true);
            System.out.println("testReport28 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport17(), "testReport29", true);
            System.out.println("testReport29 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport18(), "testReport30", true);
            System.out.println("testReport30 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport19(), "testReport31", true);
            System.out.println("testReport31 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport20(), "testReport32", true);
            System.out.println("testReport32 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport21(), "testReport33", true);
            System.out.println("testReport33 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport22(), "testReport34", true);
            System.out.println("testReport34 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport23(), "testReport35", true);
            System.out.println("testReport35 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport24(), "testReport36", true);
            System.out.println("testReport36 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            esrConstructor5.writeToExcel(testReport25(), "testReport37", true);
            System.out.println("testReport37 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }
        try{
            System.out.println("testReport37 must throw an Exception!");
            esrConstructor5.writeToExcel(testReport26(), "testReport38", true);
            System.out.println("testReport37 OK");
        } catch (ESRException e) {
            e.printStackTrace();
        }

        if (includeTestsThatWillCrashWorkbookCreation){
            try{
                System.out.println("testMoreThanMaxSheetsReport must throw an Exception!");
                for (int i = 0; i < 300; i++){
                    esrConstructor5.writeToExcel(testMoreThanMaxSheetsReport(), "moreThanMaxSheets" + i, false);
                }
                System.out.println("testMoreThanMaxSheetsReport OK");
            } catch (ESRException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * Report with title and header,
     * without header merge,
     * without main content merge,
     * without a striped background
     */
    private static ESRReport testReport1() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 3; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with title and no header,
     * without main content merge,
     * without a striped background
     */
    private static ESRReport testReport2() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title 1");
        esrTitleBlock.addTitleRow("Title 2");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k, 3000));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with header and no title,
     * without header merge,
     * without main content merge,
     * without a striped background
     */
    private static ESRReport testReport3() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 3; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with no header and no title,
     * without main content merge,
     * without a striped background
     */
    private static ESRReport testReport4() throws ESRException {
        ESRReport esrReport = new ESRReport();

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k, 3000));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with title and header,
     * with header merge,
     * without main content merge,
     * without a striped background
     */
    private static ESRReport testReport5() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with header and no title,
     * with header merge,
     * without main content merge,
     * without a striped background
     */
    private static ESRReport testReport6() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with title and header,
     * without header merge,
     * with main content merge,
     * without a striped background
     */
    private static ESRReport testReport7() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 0, 2);
            esrMerge.addMergeGroup(1, 2, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }


    private static ESRReport testReport8() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 0, 2);
            esrMerge.addMergeGroup(1, 2, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with header and no title,
     * without header merge,
     * with main content merge,
     * without a striped background
     */
    private static ESRReport testReport9() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 0, 2);
            esrMerge.addMergeGroup(1, 2, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with no title and no header,
     * with main content merge,
     * without a striped background
     */
    private static ESRReport testReport10() throws ESRException {
        ESRReport esrReport = new ESRReport();

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 0, 2);
            esrMerge.addMergeGroup(1, 2, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with title and header,
     * with header merge,
     * with main content merge,
     * without a striped background
     */
    private static ESRReport testReport11() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 0, 2);
            esrMerge.addMergeGroup(1, 2, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with header and no title,
     * with header merge,
     * with main content merge,
     * without a striped background
     */
    private static ESRReport testReport12() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 0, 2);
            esrMerge.addMergeGroup(1, 2, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with title and header,
     * without header merge,
     * without main content merge,
     * with different color schemes
     */
    private static ESRReport testReport13() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        ESRBlock esrBlock = new ESRTableBlock();
        for (int j = 0; j < 3; j++){
            ESRRow esrRow = new ESRTableRow();
            ESRCell esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 0);
            esrCell.setStyle(esrConstructor5.getGreenCellStyle());
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(esrConstructor5.getYellowCellStyle());
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(esrConstructor5.getRedCellStyle());
            esrRow.addCell(esrCell);
            esrBlock.addRow(esrRow);
        }
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        for (int j = 0; j < 3; j++){
            ESRRow esrRow = new ESRTableRow();
            ESRCell esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 0);
            esrCell.setStyle(customCellStyle);
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(customCellStyle);
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(customCellStyle);
            esrRow.addCell(esrCell);
            esrBlock.addRow(esrRow);
        }
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    /**
     * Report with title and header,
     * with header merge,
     * with main content merge,
     * with different color schemes
     */
    private static ESRReport testReport14() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        ESRBlock esrBlock = new ESRTableBlock();
        for (int j = 0; j < 3; j++){
            ESRRow esrRow = new ESRTableRow();
            ESRCell esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 0);
            esrCell.setStyle(esrConstructor5.getGreenCellStyle());
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(esrConstructor5.getYellowCellStyle());
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(esrConstructor5.getRedCellStyle());
            esrRow.addCell(esrCell);
            esrBlock.addRow(esrRow);
        }
        esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        for (int j = 0; j < 3; j++){
            ESRRow esrRow = new ESRTableRow();
            ESRCell esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 0);
            esrCell.setStyle(customCellStyle);
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(customCellStyle);
            esrRow.addCell(esrCell);
            esrCell = new ESRTableCell("Cell_" + 0 + "." + j + "." + 1);
            esrCell.setStyle(customCellStyle);
            esrRow.addCell(esrCell);
            esrBlock.addRow(esrRow);
        }
        esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);

        return esrReport;
    }

    /**
     * Report with max rows on one sheet
     * with title and header
     * without main content merge
     */
    private static ESRReport testReport15() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 3; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 65534; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            ESRRow esrRow = new ESRTableRow();
            for (int j = 0; j < 3; j++){
                esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j));
            }
            esrBlock.addRow(esrRow);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with max rows on one sheet
     * with title and header
     * with main content merge
     */
    private static ESRReport testReport16() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 3; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 32767; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 2; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j +"." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 1, 2);
            esrMerge.addMergeGroup(0, 1, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with title and header
     * without main content merge
     */
    private static ESRReport testReport17() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 3; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 75000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            ESRRow esrRow = new ESRTableRow();
            for (int j = 0; j < 3; j++){
                esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j));
            }
            esrBlock.addRow(esrRow);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with title and no header
     * without main content merge
     */
    private static ESRReport testReport18() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        for (int i = 0; i < 75000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            ESRRow esrRow = new ESRTableRow();
            for (int j = 0; j < 3; j++){
                esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j));
            }
            esrBlock.addRow(esrRow);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with header and no title
     * without main content merge
     */
    private static ESRReport testReport19() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 3; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 75000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            ESRRow esrRow = new ESRTableRow();
            for (int j = 0; j < 3; j++){
                esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j));
            }
            esrBlock.addRow(esrRow);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with no header and no title
     * without main content merge
     */
    private static ESRReport testReport20() throws ESRException {
        ESRReport esrReport = new ESRReport();

        for (int i = 0; i < 75000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            ESRRow esrRow = new ESRTableRow();
            for (int j = 0; j < 3; j++){
                esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j));
            }
            esrBlock.addRow(esrRow);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with title and header
     * with main content merge
     * with header merge
     */
    private static ESRReport testReport21() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 40000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 2; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j +"." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 1, 2);
            esrMerge.addMergeGroup(0, 1, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with title and no header
     * with main content merge
     */
    private static ESRReport testReport22() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        for (int i = 0; i < 40000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 2; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j +"." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 1, 2);
            esrMerge.addMergeGroup(0, 1, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with header and no title
     * with main content merge
     * with header merge
     */
    private static ESRReport testReport23() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        for (int i = 0; i < 3; i++){
            ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
            for (int j = 0; j < 3; j++){
                esrHeaderRow.addCell(new ESRHeaderCell("Header " + i + "." + j, 3000));
            }
            esrHeaderBlock.addRow(esrHeaderRow);
        }
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 2);
        esrMerge.addMergeGroup(1, 2, 0, 0);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 40000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 2; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j +"." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 1, 2);
            esrMerge.addMergeGroup(0, 1, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max rows on one sheet
     * with no header and no title
     * with main content merge
     */
    private static ESRReport testReport24() throws ESRException {
        ESRReport esrReport = new ESRReport();

        for (int i = 0; i < 40000; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 2; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 3; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j +"." + k));
                }
                esrBlock.addRow(esrRow);
            }
            ESRMerge esrMerge = new ESRMerge();
            esrMerge.addMergeGroup(0, 0, 1, 2);
            esrMerge.addMergeGroup(0, 1, 0, 0);
            esrBlock.setMerge(esrMerge);
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with max columns
     */
    private static ESRReport testReport25() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 255; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 255; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report with more than max columns
     */
    private static ESRReport testReport26() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Title");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        for (int i = 0; i < 300; i++){
            esrHeaderRow.addCell(new ESRHeaderCell("Header " + i, 3000));
        }
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        for (int i = 0; i < 3; i++){
            ESRBlock esrBlock = new ESRTableBlock();
            for (int j = 0; j < 3; j++){
                ESRRow esrRow = new ESRTableRow();
                for (int k = 0; k < 300; k++){
                    esrRow.addCell(new ESRTableCell("Cell_" + i + "." + j + "." + k));
                }
                esrBlock.addRow(esrRow);
            }
            esrReport.addBlock(esrBlock);
        }
        return esrReport;
    }

    /**
     * Report to test more than max sheets
     */
    private static ESRReport testMoreThanMaxSheetsReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

}
