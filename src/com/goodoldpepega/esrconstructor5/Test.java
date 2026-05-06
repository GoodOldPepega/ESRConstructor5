package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class Test {
    public static void main(String[] args) throws Exception {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        ESRConstructor5 esrConstructor5 = new ESRConstructor5(hssfWorkbook);
        ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
        esrTitleBlock.addTitleRow("Test String");
        ESRReport esrReport = new ESRReport();
        esrReport.setEsrTitleBlock(esrTitleBlock);
        ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
        ESRHeaderRow esrHeaderRow = new ESRHeaderRow();
        esrHeaderRow.addCell(new ESRHeaderCell("Test1", 8000));
        esrHeaderRow.addCell(new ESRHeaderCell("Test2", 3000));
        esrHeaderRow.addCell(new ESRHeaderCell("Test3", 15000));
        esrHeaderBlock.addRow(esrHeaderRow);
//        esrReport.setEsrHeaderBlock(esrHeaderBlock);
        ESRTableBlock esrTableBlock = new ESRTableBlock();
        ESRTableRow esrTableRow = new ESRTableRow();
        esrTableRow.addCell(new ESRTableCell("Test1", 8000));
        esrTableRow.addCell(new ESRTableCell("Test2", esrConstructor5.getGreenCellStyle(), 3000));
        esrTableRow.addCell(new ESRTableCell("Test3"));
        esrTableBlock.addRow(esrTableRow);
        esrTableRow = new ESRTableRow();
        esrTableRow.addCell(new ESRTableCell("Test1", 8000));
        esrTableRow.addCell(new ESRTableCell("Test2", 20000));
        esrTableRow.addCell(new ESRTableCell("Test3", 15000));
        esrTableBlock.addRow(esrTableRow);
        esrReport.addBlock(esrTableBlock);
        esrTableBlock = new ESRTableBlock();
        esrTableRow = new ESRTableRow();
        ESRTableCell esrTableCell = new ESRTableCell(1);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrTableRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(2);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrTableRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(3);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrTableRow.addCell(esrTableCell);
        esrTableBlock.addRow(esrTableRow);
        esrTableRow = new ESRTableRow();
        esrTableCell = new ESRTableCell(1);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrTableRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(2);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrTableRow.addCell(esrTableCell);
        esrTableCell = new ESRTableCell(3);
        esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
        esrTableRow.addCell(esrTableCell);
        esrTableBlock.addRow(esrTableRow);
        esrReport.addBlock(esrTableBlock);
        ESRSingleRowBlock esrSingleRowBlock = new ESRSingleRowBlock();
        esrSingleRowBlock.addCell("Test1");
        esrSingleRowBlock.addCell("Test2");
        esrSingleRowBlock.addCell("Test3");
        esrTableRow = new ESRTableRow();
        esrTableRow.addCell(new ESRTableCell("Test4", 8000));
        esrTableRow.addCell(new ESRTableCell("Test5", esrConstructor5.getGreenCellStyle(), 3000));
        esrTableRow.addCell(new ESRTableCell("Test6"));
//        esrSingleRowBlock.addRow(esrTableRow);
        esrReport.addBlock(esrSingleRowBlock);
        esrConstructor5.writeToExcel(esrReport, "Test", true);
        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:/1/test.xls"));
        hssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
