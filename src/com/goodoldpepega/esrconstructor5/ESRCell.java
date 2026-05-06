package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

/**
 * Report cell interface
 * Types of cells:
 * 0 - numeric type;
 * 1 - string type;
 * 2 - formula type;
 * 3 - blank type;
 * 4 - boolean type;
 * 5 - error type.
 */

public interface ESRCell {

    int CELL_TYPE_NUMERIC = 0;
    int CELL_TYPE_STRING = 1;
    int CELL_TYPE_FORMULA = 2;
    int CELL_TYPE_BLANK = 3;
    int CELL_TYPE_BOOLEAN = 4;
    int CELL_TYPE_ERROR = 5;

    /**
     * Get cell value
     * @return
     */
    Object getValue();

    /**
     * Set cell value
     * @param object
     */
    void setValue(Object object);

    /**
     * Get cell style
     * @return
     */
    HSSFCellStyle getStyle();

    /**
     * Set cell style
     * @param hssfCellStyle
     */
    void setStyle(HSSFCellStyle hssfCellStyle);
}
