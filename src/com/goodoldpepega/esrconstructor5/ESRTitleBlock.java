package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 04.03.2023
 *
 * Block of title
 */


public class ESRTitleBlock {
    private List<String> listOfTitleRows = new ArrayList<String>();
    private HSSFCellStyle hssfCellStyle;

    /**
     * Get rows of title
     * @return
     */
    public List<String> geTitleRows() {
        return listOfTitleRows;
    }

    /**
     * Add row to title
     * @param titleRow
     */
    public void addTitleRow(String titleRow) {
        listOfTitleRows.add(titleRow);
    }

    /**
     * Get style of title cell
     * @return
     */
    public HSSFCellStyle getStyle() {
        return hssfCellStyle;
    }

    /**
     * Set style of title cell
     * @param hssfCellStyle
     */
    public void setStyle(HSSFCellStyle hssfCellStyle) {
        this.hssfCellStyle = hssfCellStyle;
    }
}
