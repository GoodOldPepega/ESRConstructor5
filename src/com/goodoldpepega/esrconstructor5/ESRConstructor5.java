package com.goodoldpepega.esrconstructor5;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Main wizard that creates a report.
 * You can write multiple reports to one xls file. Every report will be placed on new sheet.
 *
 * Notice that when You set cell color with custom pallet, You should avoid INDIGO, LAVENDER, LIME, VIOLET, OLIVE_GREEN indexes.
 */

public class ESRConstructor5 {

    private HSSFWorkbook workbook;
    private ESRStyleHandler esrStyleHandler;
    private HSSFCellStyle currentStripeColor;
    private int rowIndex;
    private int stripeColorNumber;
    private boolean hasHeader;
    private boolean widthComplete;
    private Map<Integer, Integer> mapOfWidth;
    private static final int LAST_ROW_INDEX = 65535;

    public ESRConstructor5(HSSFWorkbook hssfWorkbook) {
        this.workbook = hssfWorkbook;
        esrStyleHandler = new ESRStyleHandler(hssfWorkbook);
        resetStripeColor();

    }

    private void reset(){
        rowIndex = 0;
        stripeColorNumber = 0;
        hasHeader = true;
        widthComplete = false;
    }

    private void resetStripeColor(){
        currentStripeColor = esrStyleHandler.getStrippedCellStyle1();
    }

    /**
     * Write report to xls file.
     * You can write multiple reports to one xls file
     * @param report ESRReport object
     * @param sheetName Name of sheet. Name must be less then 28 symbols, otherwise it will be cut.
     * @param needStrippedColourScheme if true - even and odd blocks will have different colors.
     * @throws ESRException
     */
    public void writeToExcel(ESRReport report, String sheetName, boolean needStrippedColourScheme) throws ESRException {
        reset();
        ESRTitleBlock titleBlock = report.getEsrTitleBlock();
        ESRHeaderBlock headerBlock = report.getEsrHeaderBlock();
        List<ESRBlock> listOfValueBlocks = report.getEsrBlocksList();
        int sheetIndex = 1;
        Sheet sheet = createSheet(handleSheetName(sheetName, sheetIndex));
        int titleRowsMergeQuantity = 1;
        int headerRowIndex = 0;
        if (headerBlock != null){
            titleRowsMergeQuantity = headerBlock.getMaxHeaderRowLength();
            if (titleBlock != null){
                headerRowIndex = titleBlock.getTitleRows().size();
            }
            createHeader(headerBlock, sheet, headerRowIndex);
        } else {
            hasHeader = false;
            mapOfWidth = new HashMap<Integer, Integer>();
            if (listOfValueBlocks.size() > 0) {
                for (ESRBlock rowsBlock : listOfValueBlocks) {
                    List<ESRRow> rows = rowsBlock.getRows();
                    if (rows.size() > 0) {
                        ESRRow esrRow = rows.get(0);
                        titleRowsMergeQuantity = esrRow.getCells().size();
                    }
                }
            }
        }
        if (titleBlock != null){
            if (headerBlock == null){
                rowIndex = titleBlock.getTitleRows().size();
            }
            createTitleRows(titleBlock, sheet, titleRowsMergeQuantity);
        }
        for (int i = 0; i < listOfValueBlocks.size(); i++) {
            ESRBlock rowsBlock = listOfValueBlocks.get(i);
            if (rowIndex + rowsBlock.getRows().size() >= LAST_ROW_INDEX) {
                if (sheetIndex == 1) {
                    workbook.setSheetName(workbook.getSheetIndex(sheet), sheetName + "_" + sheetIndex);
                }
                sheet = createSheet(handleSheetName(sheetName, ++sheetIndex));
                if (titleBlock != null){
                    if (headerBlock == null){
                        rowIndex = titleBlock.getTitleRows().size();
                    }
                    createTitleRows(titleBlock, sheet, titleRowsMergeQuantity);
                }
                if (headerBlock != null) {
                    createHeader(headerBlock, sheet, headerRowIndex);
                }
                if (titleBlock == null && headerBlock == null){
                    rowIndex = 0;
                }
                resetStripeColor();
            }
            writeRowsBlock(rowsBlock, sheet, needStrippedColourScheme);
            toggleStripeColors();
        }
        if (!hasHeader) {
            for (Integer columnIndex : mapOfWidth.keySet()) {
                sheet.setColumnWidth(columnIndex, mapOfWidth.get(columnIndex));
            }
        }
    }

    private String handleSheetName(String name, int sheetIndex){
        if (sheetIndex > 1){
            name += "_" + sheetIndex;
        }
        if (name.length() > 28){
            name = name.substring(0, 28);
        }
        return name;
    }

    private Sheet createSheet (String sheetName) throws ESRException {
        if (workbook.getNumberOfSheets() == 255) {
            throw new ESRException(ESRConstructor5_old.class.getName(),
                    "EXCEL 2003 does not allows add more than 255 sheets");
        }
        return workbook.createSheet(sheetName);
    }

    private void createHeader(ESRHeaderBlock headerBlock, Sheet sheet, int headerRowIndex) {
        this.rowIndex = headerRowIndex;
        int firstRowIndex = rowIndex;
        List headerRows = headerBlock.getRows();
        for (int k = 0; k < headerRows.size(); k++) {
            ESRHeaderRow esrHeaderRow = (ESRHeaderRow)headerRows.get(k);
            Row row = sheet.createRow(rowIndex);
            List listOfCells = esrHeaderRow.getCells();
            for (int i = 0; i < listOfCells.size(); i++) {
                ESRHeaderCell esrHeaderCell = (ESRHeaderCell) listOfCells.get(i);
                sheet.setColumnWidth(i, esrHeaderCell.getWidth());
                Cell cell = row.createCell(i);
                if (esrHeaderCell.getStyle() == null) {
                    esrHeaderCell.setStyle(esrStyleHandler.getHeaderCellStyle());
                }
                cell.setCellStyle(esrHeaderCell.getStyle());
                cell.setCellValue(String.valueOf(esrHeaderCell.getValue()));
            }
            rowIndex++;

            if (headerBlock.getMerge() != null) {
                ESRMerge blockMerge = headerBlock.getMerge();
                List<int[]> listOfMerges = blockMerge.getListOfMerges();
                for (int[] mergeArray : listOfMerges) {
                    try{
                        sheet.addMergedRegion(new CellRangeAddress(firstRowIndex + mergeArray[0],
                                firstRowIndex + mergeArray[1], mergeArray[2], mergeArray[3]));
                    } catch (Exception e) {
                        System.out.println("ESRConstructor: Merge " + mergeArray[0] + mergeArray[1] + mergeArray[2] +
                                mergeArray[3] + " exceeds existing cells range in header block");
                    }
                }
            }
        }
    }

    private void createTitleRows(ESRTitleBlock titleBlock, Sheet sheet, int mergedColumnsNumber) {
        List<String> listOfTitleRows = titleBlock.getTitleRows();
        int titleRowIndex = 0;
        for (String titleString : listOfTitleRows) {
            Row row = sheet.createRow(titleRowIndex);
            for (int i = 0; i < mergedColumnsNumber; i++) {
                Cell cell = row.createCell(i);
                if (titleBlock.getStyle() == null) {
                    titleBlock.setStyle(esrStyleHandler.getTitleCellStyle());
                }
                cell.setCellStyle(titleBlock.getStyle());
                if (i == 0) {
                    cell.setCellValue(titleString);
                }
            }
            int newMergedColumnsNumber = mergedColumnsNumber;
            if (newMergedColumnsNumber > 0) {
                newMergedColumnsNumber--;
            }
            sheet.addMergedRegion(new CellRangeAddress(titleRowIndex, titleRowIndex, 0,
                    newMergedColumnsNumber));
            titleRowIndex++;
        }
    }

    private void writeRowsBlock(ESRBlock rowsBlock, Sheet sheet, boolean needStrippedColourScheme) {
        int firstRowIndex = rowIndex;
        boolean hasMerge = rowsBlock.getMerge() != null;
        List<ESRRow> listOfRows = rowsBlock.getRows();
        for (int i = 0; i < listOfRows.size(); i++) {
            ESRTableRow esrTableRow = (ESRTableRow) listOfRows.get(i);
            Row row = sheet.createRow(rowIndex++);
            writeRow(esrTableRow, row, needStrippedColourScheme, hasMerge);
        }
        if (rowsBlock.getMerge() != null) {
            ESRMerge esrMerge = rowsBlock.getMerge();
            for (int[] mergeArray : esrMerge.getListOfMerges()) {
                try{
                    sheet.addMergedRegion(new CellRangeAddress(firstRowIndex + mergeArray[0],
                            firstRowIndex + mergeArray[1], mergeArray[2], mergeArray[3]));
                }
                catch (Exception e) {
                    System.out.println("ESRConstructor: Merge " + mergeArray[0] + mergeArray[1] + mergeArray[2] +
                            mergeArray[3] + " exceeds existing cells range in table block`");
                }
            }
        }
    }

    private void writeRow(ESRTableRow srRow, Row row, boolean needStrippedColourScheme, boolean hasMerge) {
        List<ESRCell> esrCells = srRow.getCells();
        int maxLinesQuantity = 0;
        int esrCellsSize = esrCells.size();
        for (int i = 0; i < esrCellsSize; i++) {
            ESRTableCell esrTableCell = (ESRTableCell) esrCells.get(i);
            if (!hasHeader && !widthComplete) {
                int width = esrTableCell.getWidth();
                if (width > 0) {
                    if (mapOfWidth.get(i) == null) {
                        mapOfWidth.put(i, width);
                        if (mapOfWidth.size() == esrCellsSize) {
                            widthComplete = true;
                        }
                    }
                }
            }
            Cell cell = row.createCell(i);
            HSSFCellStyle cellStyle;
            if (esrTableCell.getStyle() != null) {
                cellStyle = esrTableCell.getStyle();
            } else if (needStrippedColourScheme) {
                cellStyle = currentStripeColor;
            } else {
                cellStyle = esrStyleHandler.getDefaultCellStyle();
            }
            Object cellValueObject = esrTableCell.getValue();
            //!!! Здесь можно добавить объединение значений при мерже
            if (hasMerge) {
                int colWidth = row.getSheet().getColumnWidth(i);
                int valueLength = String.valueOf(cellValueObject).length();
                int newLineQuantity = String.valueOf(cellValueObject).split("\n").length;
                if (newLineQuantity > 0) {
                    newLineQuantity = newLineQuantity - 1;
                }
                int linesQuantity = (valueLength * 256 / colWidth) + newLineQuantity;
                if (maxLinesQuantity < linesQuantity) {
                    maxLinesQuantity = linesQuantity;
                }
            }
            setValueToCell(esrTableCell, cell);
            cell.setCellStyle(cellStyle);
            setHyperlink(esrTableCell, cell);
        }
        if (maxLinesQuantity == 0) {
            maxLinesQuantity = 1;
        }
        if (hasMerge) {
            row.setHeightInPoints(maxLinesQuantity * 10 * 1.7F);
        }

    }

    private void setValueToCell(ESRTableCell esrTableCell, Cell cell){
        Object cellValueObject = esrTableCell.getValue();
        String valueS = String.valueOf(cellValueObject);
        if (esrTableCell.getType() == Cell.CELL_TYPE_NUMERIC && ESRHelper.checkIfDouble(valueS)) {
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            cell.setCellValue(Double.parseDouble(valueS));
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_FORMULA) {
            cell.setCellType(Cell.CELL_TYPE_FORMULA);
            cell.setCellFormula(valueS);
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_BLANK) {
            cell.setCellType(Cell.CELL_TYPE_BLANK);
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_BOOLEAN) {
            if ("true".equals(valueS) || "false".equals(valueS)) {
                cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
                cell.setCellValue(Boolean.parseBoolean(valueS));
            } else {
                cell.setCellValue(valueS);
            }
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_ERROR) {
            cell.setCellType(Cell.CELL_TYPE_ERROR);
        } else {
            cell.setCellValue(valueS);
        }
    }

    private void setHyperlink(ESRTableCell esrTableCell, Cell cell){
        if (esrTableCell.getHyperlink() != null) {
            String url = esrTableCell.getHyperlink();
            Hyperlink link = (workbook).getCreationHelper().createHyperlink(HSSFHyperlink.LINK_URL);

            link.setAddress(url);
            cell.setHyperlink(link);
        }
    }

    private void toggleStripeColors(){
        stripeColorNumber++;
        if (stripeColorNumber % 2 == 0) {
            resetStripeColor();
        } else {
            currentStripeColor = esrStyleHandler.getStrippedCellStyle2();
        }
    }

    /**
     * Get default cell style
     * @return
     */
    public HSSFCellStyle getDefaultCellStyle() {
        return esrStyleHandler.getDefaultCellStyle();
    }

    /**
     * Set new default cell style
     * @param hssfCellStyle New default cell style
     */
    public void setDefaultCellStyle(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setDefaultCellStyle(hssfCellStyle);
    }

    /**
     * Get green cell style
     * @return
     */
    public HSSFCellStyle getGreenCellStyle() {
        return esrStyleHandler.getGreenCellStyle();
    }

    /**
     * Set new green cell style
     * @param hssfCellStyle New default cell style
     */
    public void setGreenCellStyle(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setGreenCellStyle(hssfCellStyle);
    }

    /**
     * Set new RGB color of green cell style
     * @param red
     * @param green
     * @param blue
     */
    public void setGreenCellStyleColor(int red, int green, int blue) {
        HSSFCellStyle hssfCellStyle = esrStyleHandler.getGreenCellStyle();
        HSSFPalette palette = workbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.INDIGO.index, (byte)red,(byte)green,(byte)blue);
        hssfCellStyle.setFillForegroundColor(hssfCellStyle.getIndex());
    }

    /**
     * Get yellow cell style
     * @return
     */
    public HSSFCellStyle getYellowCellStyle() {
        return esrStyleHandler.getYellowCellStyle();
    }

    /**
     * Set new yellow cell style
     * @param hssfCellStyle New default cell style
     */
    public void setYellowCellStyle(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setYellowCellStyle(hssfCellStyle);
    }

    /**
     * Set new RGB color of yellow cell style
     * @param red
     * @param green
     * @param blue
     */
    public void setYellowCellStyleColor(int red, int green, int blue) {
        HSSFCellStyle hssfCellStyle = esrStyleHandler.getYellowCellStyle();
        HSSFPalette palette = workbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.LAVENDER.index, (byte)red,(byte)green,(byte)blue);
        hssfCellStyle.setFillForegroundColor(hssfCellStyle.getIndex());
    }

    /**
     * Get red cell style
     * @return
     */
    public HSSFCellStyle getRedCellStyle() {
        return esrStyleHandler.getRedCellStyle();
    }

    /**
     * Set new red cell style
     * @param hssfCellStyle New default cell style
     */
    public void setRedCellStyle(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setRedCellStyle(hssfCellStyle);
    }

    /**
     * Set new RGB color of red cell style
     * @param red
     * @param green
     * @param blue
     */
    public void setRedCellStyleColor(int red, int green, int blue) {
        HSSFCellStyle hssfCellStyle = esrStyleHandler.getRedCellStyle();
        HSSFPalette palette = workbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.LIME.index, (byte)red,(byte)green,(byte)blue);
        hssfCellStyle.setFillForegroundColor(hssfCellStyle.getIndex());
    }

    /**
     * Get header cell style
     * @return
     */
    public HSSFCellStyle getHeaderCellStyle() {
        return esrStyleHandler.getHeaderCellStyle();
    }

    /**
     * Set new header cell style
     * @param hssfCellStyle New default cell style
     */
    public void setHeaderCellStyle(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setHeaderCellStyle(hssfCellStyle);
    }

    /**
     * Get title cell style
     * @return
     */
    public HSSFCellStyle getTitleCellStyle() {
        return esrStyleHandler.getTitleCellStyle();
    }

    /**
     * Set new title cell style
     * @param hssfCellStyle New default cell style
     */
    public void setTitleCellStyle(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setTitleCellStyle(hssfCellStyle);
    }

    /**
     * Set new stripped cell style 1
     * @param hssfCellStyle New default cell style
     */
    public void setStrippedCellStyle1(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setStrippedCellStyle1(hssfCellStyle);
    }

    /**
     * Set new RGB color of stripped cell style 1
     * @param red
     * @param green
     * @param blue
     */
    public void setStrippedCellStyle1Color(int red, int green, int blue) {
        HSSFCellStyle hssfCellStyle = esrStyleHandler.getStrippedCellStyle1();
        HSSFPalette palette = workbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.VIOLET.index, (byte)red,(byte)green,(byte)blue);
        hssfCellStyle.setFillForegroundColor(hssfCellStyle.getIndex());
    }

    /**
     * Set new strinpped cell style 2
     * @param hssfCellStyle New default cell style
     */
    public void setStrippedCellStyle2(HSSFCellStyle hssfCellStyle) {
        esrStyleHandler.setStrippedCellStyle2(hssfCellStyle);
    }

    /**
     * Set new RGB color of stripped cell style 2
     * @param red
     * @param green
     * @param blue
     */
    public void setStrippedCellStyle2Color(int red, int green, int blue) {
        HSSFCellStyle hssfCellStyle = esrStyleHandler.getStrippedCellStyle2();
        HSSFPalette palette = workbook.getCustomPalette();
        palette.setColorAtIndex(HSSFColor.OLIVE_GREEN.index, (byte)red,(byte)green,(byte)blue);
        hssfCellStyle.setFillForegroundColor(hssfCellStyle.getIndex());
    }

}
