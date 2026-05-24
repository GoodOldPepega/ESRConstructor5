package com.goodoldpepega.esrconstructor5.examples;

import com.goodoldpepega.esrconstructor5.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class Example1 {
    public static void main(String[] args) throws Exception {
        createReport();
    }

    public static void createReport() throws Exception {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        ESRConstructor5 esrConstructor5 = new ESRConstructor5(hssfWorkbook);
        esrConstructor5.writeToExcel(createFirstReport(), "First Report", false);
        esrConstructor5.writeToExcel(createSecondReport(), "Second Report", false);
        esrConstructor5.writeToExcel(createThirdReport(), "Third Report", false);
        esrConstructor5.writeToExcel(createForthReport(), "Forth Report", true);
        esrConstructor5.writeToExcel(createFifthReport(), "Fifth Report", false);
        esrConstructor5.writeToExcel(createSixthReport(), "Sixth Report", false);
        esrConstructor5.writeToExcel(createSeventhReport(), "Seventh Report", false);
        esrConstructor5.writeToExcel(createEightReport(), "EightReport Report", false);
        esrConstructor5.writeToExcel(createNinthReport(), "NinthReport Report", false);
        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:/Test/testReport.xls"));
        hssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    /**
     * Report with single header row and without any merge
     * @return
     * @throws ESRException
     */
    private static ESRReport createFirstReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A title for first report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column3", 5000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 1.1"));
        esrRow.addCell(new ESRTableCell("Cell 1.2"));
        esrRow.addCell(new ESRTableCell("Cell 1.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 2.1"));
        esrRow.addCell(new ESRTableCell("Cell 2.2"));
        esrRow.addCell(new ESRTableCell("Cell 2.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1"));
        esrRow.addCell(new ESRTableCell("Cell 3.2"));
        esrRow.addCell(new ESRTableCell("Cell 3.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    /**
     * Report with using single row blocks
     * @return
     * @throws ESRException
     */
    private static ESRReport createSecondReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A title for second report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column3", 5000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRSingleRowBlock esrBlock = new ESRSingleRowBlock();
        esrBlock.addCell("Cell 1.1");
        esrBlock.addCell("Cell 1.2");
        esrBlock.addCell("Cell 1.3");
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRSingleRowBlock();
        esrBlock.addCell("Cell 2.1");
        esrBlock.addCell("Cell 2.2");
        esrBlock.addCell("Cell 2.3");
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRSingleRowBlock();
        esrBlock.addCell("Cell 3.1");
        esrBlock.addCell("Cell 3.2");
        esrBlock.addCell("Cell 3.3");
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    /**
     * Report with double rows in header and merged cells
     * @return
     * @throws ESRException
     */
    private static ESRReport createThirdReport() throws ESRException{
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A first row of title for third report");
        esrTitleBlock.addTitleRow("A second row of title for third report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1.1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2.1", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column3.1", 5000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column2.1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2.2", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2.3", 5000));
        esrHeaderBlock.addRow(esrHeaderRow);
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0,0,0,1);
        esrHeaderBlock.setMerge(esrMerge);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 1.1"));
        esrRow.addCell(new ESRTableCell("Cell 1.2"));
        esrRow.addCell(new ESRTableCell("Cell 1.3"));
        esrBlock.addRow(esrRow);
        esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 1, 2);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 2.1"));
        esrRow.addCell(new ESRTableCell("Cell 2.2"));
        esrRow.addCell(new ESRTableCell("Cell 2.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1"));
        esrRow.addCell(new ESRTableCell("Cell 3.2"));
        esrRow.addCell(new ESRTableCell("Cell 3.3"));
        esrBlock.addRow(esrRow);
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1.1"));
        esrRow.addCell(new ESRTableCell("Cell 3.2.2"));
        esrRow.addCell(new ESRTableCell("Cell 3.3.3"));
        esrBlock.addRow(esrRow);
        esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 1, 0, 0);
        esrMerge.addMergeGroup(0, 1, 2, 2);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }


    /**
     * Report with single column
     * @return
     * @throws ESRException
     */
    private static ESRReport createForthReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A title for forth report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1", 7000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 1.1"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 2.1"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }


    /**
     * Report with different cell types
     * @return
     * @throws ESRException
     */
    private static ESRReport createFifthReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A title for fifth report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column3", 5000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column4", 6000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column5", 7000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        ESRTableCell esrTableCell = new ESRTableCell("Cell 1.1");
        esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(101);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(true);
        esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell("SUM(B4:B5)");
        esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell("Blank cell");
        esrTableCell.setType(ESRCell.CELL_TYPE_BLANK);
        esrRow.addCell(esrTableCell);
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrTableCell = new ESRTableCell("Cell 2.1");
        esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(102);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(false);
        esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell("B4+B5");
        esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell("Blank cell");
        esrTableCell.setType(ESRCell.CELL_TYPE_BLANK);
        esrRow.addCell(esrTableCell);
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrTableCell = new ESRTableCell("Cell 3.1");
        esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(103);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(true);
        esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell("B4*B5");
        esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell("Blank cell");
        esrTableCell.setType(ESRCell.CELL_TYPE_BLANK);
        esrRow.addCell(esrTableCell);
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    private static ESRReport createSixthReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A title for sixth report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column3", 5000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        ESRTableCell esrTableCell = new ESRTableCell(100);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(90);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(110);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrBlock.addRow(esrRow);
        ESRMerge esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 1);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrTableCell = new ESRTableCell(101);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(91);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(111);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrBlock.addRow(esrRow);
        esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 1);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrTableCell = new ESRTableCell(102);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(92);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(112);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrRow.addCell(esrTableCell);
        esrBlock.addRow(esrRow);
        esrMerge = new ESRMerge();
        esrMerge.addMergeGroup(0, 0, 0, 1);
        esrBlock.setMerge(esrMerge);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    /**
     * Report without title block
     * @return
     * @throws ESRException
     */
    private static ESRReport createSeventhReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a header:
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column2", 4000));
        esrHeaderRow.addCell(new ESRHeaderCell("Column3", 5000));
        esrHeaderBlock.addRow(esrHeaderRow);
        esrReport.setEsrHeaderBlock(esrHeaderBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 1.1"));
        esrRow.addCell(new ESRTableCell("Cell 1.2"));
        esrRow.addCell(new ESRTableCell("Cell 1.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 2.1"));
        esrRow.addCell(new ESRTableCell("Cell 2.2"));
        esrRow.addCell(new ESRTableCell("Cell 2.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1"));
        esrRow.addCell(new ESRTableCell("Cell 3.2"));
        esrRow.addCell(new ESRTableCell("Cell 3.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    /**
     * Report without header block
     * @return
     * @throws ESRException
     */
    private static ESRReport createEightReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create a title:
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("A title for eight report");
        esrReport.setEsrTitleBlock(esrTitleBlock);

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 1.1", 3000));
        esrRow.addCell(new ESRTableCell("Cell 1.2", 4000));
        esrRow.addCell(new ESRTableCell("Cell 1.3", 5000));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 2.1"));
        esrRow.addCell(new ESRTableCell("Cell 2.2"));
        esrRow.addCell(new ESRTableCell("Cell 2.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1"));
        esrRow.addCell(new ESRTableCell("Cell 3.2"));
        esrRow.addCell(new ESRTableCell("Cell 3.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }

    /**
     * Report without title and header blocks
     * @return
     * @throws ESRException
     */
    private static ESRReport createNinthReport() throws ESRException {
        ESRReport esrReport = new ESRReport();

        //Create main content
        ESRBlock esrBlock = new ESRTableBlock();
        ESRRow esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 1.1", 3000));
        esrRow.addCell(new ESRTableCell("Cell 1.2", 4000));
        esrRow.addCell(new ESRTableCell("Cell 1.3", 5000));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 2.1"));
        esrRow.addCell(new ESRTableCell("Cell 2.2"));
        esrRow.addCell(new ESRTableCell("Cell 2.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        esrBlock = new ESRTableBlock();
        esrRow = new ESRTableRow();
        esrRow.addCell(new ESRTableCell("Cell 3.1"));
        esrRow.addCell(new ESRTableCell("Cell 3.2"));
        esrRow.addCell(new ESRTableCell("Cell 3.3"));
        esrBlock.addRow(esrRow);
        esrReport.addBlock(esrBlock);
        return esrReport;
    }
}
