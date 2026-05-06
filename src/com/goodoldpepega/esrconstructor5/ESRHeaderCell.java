package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Header cell class.
 * Has standard style that You can override
 */

public class ESRHeaderCell implements ESRCell {
    private Object value;
    private int width;
    private HSSFCellStyle hssfCellStyle;

    /**
     * Create new header cell
     * @param value Text value of cell
     * @param width width of column in sheet
     */
    public ESRHeaderCell(Object value, int width) {
        this.value = value;
        this.width = width;
    }

    /**
     * Get value of cell
     * @return
     */
    public Object getValue() {
        return value;
    }

    /**
     * Set value of cell
     * @param object
     */
    public void setValue(Object object) {
        this.value = object;
    }

    /**
     * Get style of cell
     * @return
     */
    public HSSFCellStyle getStyle() {
        return hssfCellStyle;
    }

    /**
     * Set style of cell
     * @param hssfCellStyle
     */
    public void setStyle(HSSFCellStyle hssfCellStyle) {
        this.hssfCellStyle = hssfCellStyle;
    }

    /**
     * Get width of column of cell
     * @return
     */
    public int getWidth() {
        return width;
    }

    /**
     * Set width of cell column
     * @param width
     */
    public void setWidth(int width) {
        this.width = width;
    }
}
