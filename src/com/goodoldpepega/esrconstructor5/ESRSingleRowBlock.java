package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 04.03.2023
 *
 * Single row block.
 * Block that has only one row.
 * Notice that if You will add another row, the row in block will be rewritten.
 */


public class ESRSingleRowBlock implements ESRBlock {

    private List<ESRRow> listOfRows = new ArrayList<ESRRow>();
    private ESRRow row;

    /**
     * Get cells of single row block
     * @return
     */
    public List<ESRCell> getListOfCells() {
        return row.getCells();
    }

    /**
     * Add cell to single row block
     * @param value Text value of cell
     * @throws ESRException
     */
    public void addCell(String value) throws ESRException {
        if (row == null) {
            row = new ESRTableRow();
            listOfRows.add(row);
        }
        ESRCell esrCell = new ESRTableCell(value);
        row.addCell(esrCell);
    }

    /**
     * Add cell to single row block
     * @param value Text value of cell
     * @param width Width of cell
     * @throws ESRException
     */
    public void addCell(String value, int width) throws ESRException {
        if (row == null) {
            row = new ESRTableRow();
            listOfRows.add(row);
        }
        ESRCell esrCell = new ESRTableCell(value, width);
        row.addCell(esrCell);
    }

    /**
     * Add cell to single row block
     * @param value Text value of cell
     * @param style Style of cell
     * @throws ESRException
     */
    public void addCell(String value, HSSFCellStyle style) throws ESRException {
        if (row == null) {
            row = new ESRTableRow();
            listOfRows.add(row);
        }
        ESRCell esrCell = new ESRTableCell(value, style);
        row.addCell(esrCell);
    }

    /**
     * Add cell to single row block
     * @param value Text value of cell
     * @param width Width of cell
     * @param style Style of cell
     * @throws ESRException
     */
    public void addCell(String value, int width, HSSFCellStyle style) throws ESRException {
        if (row == null) {
            row = new ESRTableRow();
            listOfRows.add(row);
        }
        ESRCell esrCell = new ESRTableCell(value, style, width);
        row.addCell(esrCell);
    }

    /**
     * Get rows.
     * Always returns one row
     * @return
     */
    public List<ESRRow> getRows() {
        return listOfRows;
    }

    /**
     * Be careful, object ESRSingleRowBlock can have only one ESRRow, so
     * this method will rewrite existing ESRRow.
     * @param row
     */
    public void addRow(ESRRow row) {
        listOfRows = new ArrayList<ESRRow>();
        listOfRows.add(row);
    }

    /**
     * ESRSingleRowBlock can not have any merge. Method will always return null.
     * @return NULL
     */
    public ESRMerge getMerge() {
        return null;
    }

    /**
     * ESRSingleRowBlock can not have any merge. Method will do nothing.
     * @param esrMerge
     */
    public void setMerge(ESRMerge esrMerge) {
    }
}
