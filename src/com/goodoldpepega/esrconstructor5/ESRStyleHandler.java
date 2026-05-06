package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Created by Anton Nikitin
 * 01.03.2023
 *
 * Handler of cells styles
 */


public class ESRStyleHandler {

    private HSSFWorkbook hssfWorkbook;
    private HSSFCellStyle headerCellStyle;
    private HSSFCellStyle titleCellStyle;
    private HSSFCellStyle defaultCellStyle;
    private HSSFCellStyle greenCellStyle;
    private HSSFCellStyle yellowCellStyle;
    private HSSFCellStyle redCellStyle;
    private HSSFCellStyle strippedCellStyle1;
    private HSSFCellStyle strippedCellStyle2;

    /**
     * Style of workbook
     * @param hssfWorkbook HSSFWorkbook
     */
    public ESRStyleHandler(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
        initCellStyles();

    }

    /**
     * Initialize all standard cell styles
     */
    private void initCellStyles() {
        HSSFDataFormat dataFormat = hssfWorkbook.createDataFormat();
        HSSFPalette palette = hssfWorkbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.INDIGO.index, (byte)90,(byte)200,(byte)15);
        palette.setColorAtIndex(HSSFColor.LAVENDER.index, (byte)250,(byte)230,(byte)40);
        palette.setColorAtIndex(HSSFColor.LIME.index, (byte)250,(byte)30,(byte)5);
        palette.setColorAtIndex(HSSFColor.VIOLET.index, (byte)255,(byte)250,(byte)230);
        palette.setColorAtIndex(HSSFColor.OLIVE_GREEN.index, (byte)255,(byte)235,(byte)210);
        HSSFColor customGreenColor = palette.getColor(HSSFColor.INDIGO.index);
        HSSFColor customYellowColor = palette.getColor(HSSFColor.LAVENDER.index);
        HSSFColor customRedColor = palette.getColor(HSSFColor.LIME.index);
        HSSFColor customStripped1Color = palette.getColor(HSSFColor.VIOLET.index);
        HSSFColor customStripped2Color = palette.getColor(HSSFColor.OLIVE_GREEN.index);

        headerCellStyle = hssfWorkbook.createCellStyle();
        titleCellStyle = hssfWorkbook.createCellStyle();
        defaultCellStyle = hssfWorkbook.createCellStyle();
        greenCellStyle = hssfWorkbook.createCellStyle();
        yellowCellStyle = hssfWorkbook.createCellStyle();
        redCellStyle = hssfWorkbook.createCellStyle();
        strippedCellStyle1 = hssfWorkbook.createCellStyle();
        strippedCellStyle2 = hssfWorkbook.createCellStyle();

        HSSFCellStyle[] defaultStylesArray = new HSSFCellStyle[]{headerCellStyle, titleCellStyle, defaultCellStyle,
                greenCellStyle, yellowCellStyle, redCellStyle, strippedCellStyle1,
                strippedCellStyle2};

        for (HSSFCellStyle hssfCellStyle : defaultStylesArray) {
            hssfCellStyle.setWrapText(true);
            hssfCellStyle.setBorderBottom((short)2);
            hssfCellStyle.setBorderTop((short)2);
            hssfCellStyle.setBorderLeft((short)2);
            hssfCellStyle.setBorderRight((short)2);

        }

        HSSFFont hdrFont = hssfWorkbook.createFont();
        hdrFont.setBoldweight((short)700);
        headerCellStyle.setFont(hdrFont);
        headerCellStyle.setAlignment((short)2);
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerCellStyle.setFillPattern((short)1);
        headerCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);


        titleCellStyle.setDataFormat(dataFormat.getFormat("@"));
        titleCellStyle.setVerticalAlignment((short)1);
        titleCellStyle.setFont(hdrFont);


        defaultCellStyle.setDataFormat(dataFormat.getFormat("@"));
        defaultCellStyle.setVerticalAlignment((short)1);


        greenCellStyle.setDataFormat(dataFormat.getFormat("@"));
        greenCellStyle.setVerticalAlignment((short)1);
        greenCellStyle.setFillForegroundColor(customGreenColor.getIndex());
        greenCellStyle.setFillPattern((short)1);


        yellowCellStyle.setDataFormat(dataFormat.getFormat("@"));
        yellowCellStyle.setVerticalAlignment((short)1);
        yellowCellStyle.setFillForegroundColor(customYellowColor.getIndex());
        yellowCellStyle.setFillPattern((short)1);

        redCellStyle.setDataFormat(dataFormat.getFormat("@"));
        redCellStyle.setVerticalAlignment((short)1);
        redCellStyle.setFillForegroundColor(customRedColor.getIndex());
        redCellStyle.setFillPattern((short)1);


        strippedCellStyle1.setDataFormat(dataFormat.getFormat("@"));
        strippedCellStyle1.setVerticalAlignment((short)1);
        strippedCellStyle1.setFillForegroundColor(customStripped1Color.getIndex());
        strippedCellStyle1.setFillPattern((short)1);

        strippedCellStyle2.setDataFormat(dataFormat.getFormat("@"));
        strippedCellStyle2.setVerticalAlignment((short)1);
        strippedCellStyle2.setFillForegroundColor(customStripped2Color.getIndex());
        strippedCellStyle2.setFillPattern((short)1);
    }

    HSSFCellStyle getHeaderCellStyle() {
        return headerCellStyle;
    }

    void setHeaderCellStyle(HSSFCellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
    }

    HSSFCellStyle getTitleCellStyle() {
        return titleCellStyle;
    }

    void setTitleCellStyle(HSSFCellStyle titleCellStyle) {
        this.titleCellStyle = titleCellStyle;
    }

    HSSFCellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    void setDefaultCellStyle(HSSFCellStyle defaultCellStyle) {
        this.defaultCellStyle = defaultCellStyle;
    }

    HSSFCellStyle getGreenCellStyle() {
        return greenCellStyle;
    }

    void setGreenCellStyle(HSSFCellStyle greenCellStyle) {
        this.greenCellStyle = greenCellStyle;
    }

    HSSFCellStyle getYellowCellStyle() {
        return yellowCellStyle;
    }

    void setYellowCellStyle(HSSFCellStyle yellowCellStyle) {
        this.yellowCellStyle = yellowCellStyle;
    }

    HSSFCellStyle getRedCellStyle() {
        return redCellStyle;
    }

    void setRedCellStyle(HSSFCellStyle redCellStyle) {
        this.redCellStyle = redCellStyle;
    }

    HSSFCellStyle getStrippedCellStyle1() {
        return strippedCellStyle1;
    }

    void setStrippedCellStyle1(HSSFCellStyle strippedCellStyle1) {
        this.strippedCellStyle1 = strippedCellStyle1;
    }

    HSSFCellStyle getStrippedCellStyle2() {
        return strippedCellStyle2;
    }

    void setStrippedCellStyle2(HSSFCellStyle strippedCellStyle2) {
        this.strippedCellStyle2 = strippedCellStyle2;
    }
}
