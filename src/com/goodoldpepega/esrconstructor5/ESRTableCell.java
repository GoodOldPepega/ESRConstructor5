package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Table cell
 */


public class ESRTableCell implements ESRCell {
    private Object value;
    private String hyperlink;
    private int cellType = 1;
    private HSSFCellStyle hssfCellStyle;
    private int width;

    /**
     * Cell with default style
     * @param value
     */
    public ESRTableCell(Object value) {
        this.value = value;
    }

    /**
     * Cell with default style and custom width. Pay attention that first of all width implements from header cells.
     * If there is no header block, width implements from first row with assigned width
     * @param value
     * @param width
     */
    public ESRTableCell(Object value, int width) {
        this.value = value;
        this.width = width;
    }

    /**
     * Cell with custom style and custom width. Pay attention that first of all width implements from header cells.
     * If there is no header block, width implements from first row with assigned width
     * @param value
     * @param hssfCellStyle
     */
    public ESRTableCell(Object value, HSSFCellStyle hssfCellStyle) {
        this.value = value;
        this.hssfCellStyle = hssfCellStyle;
    }

    /**
     * Cell with custom style and custom width. Pay attention that first of all width implements from header cells.
     * If there is no header block, width implements from first row with assigned width
     * @param value
     * @param hssfCellStyle
     * @param width
     */
    public ESRTableCell(Object value, HSSFCellStyle hssfCellStyle, int width) {
        this.value = value;
        this.hssfCellStyle = hssfCellStyle;
        this.width = width;
    }

    /**
     * Get cell value
     * @return
     */
    public Object getValue() {
        return value;
    }

    /**
     * Set cell value
     * @param object
     */
    public void setValue(Object object) {
        this.value = object;
    }

    /**
     * Get cell hyperlink
     * @return
     */
    public String getHyperlink() {
        return hyperlink;
    }

    /**
     * Set cell hyperling
     * @param url
     */
    public void setHyperlink(String url) {
        this.hyperlink = url;
    }

    /**
     * Get cell type
     * @return
     */
    public int getType() {
        return cellType;
    }

    /**
     * Set cell type
     * 0 - numeric type;
     * 1 - string type;
     * 2 - formula type;
     * 3 - blank type;
     * 4 - boolean type;
     * 5 - error type.
     * @param type
     */
    public void setType(int type) {
        this.cellType = type;
    }

    /**
     * Get cell style
     * @return
     */
    public HSSFCellStyle getStyle() {
        return hssfCellStyle;
    }

    /**
     * Set cell style
     * @param hssfCellStyle
     */
    public void setStyle(HSSFCellStyle hssfCellStyle) {
        this.hssfCellStyle = hssfCellStyle;
    }

    /**
     * Get cell width. Pay attention that first of all width implements from header cells.
     * If there is no header block, width implements from first row with assigned width
     * @return
     */
    public int getWidth() {
        return width;
    }

    /**
     * Set cell width. Pay attention that first of all width implements from header cells.
     * If there is no header block, width implements from first row with assigned width
     * @param width
     */
    public void setWidth(int width) {
        this.width = width;
    }
}
