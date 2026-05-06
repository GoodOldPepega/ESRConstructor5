package com.goodoldpepega.esrconstructor5;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Row of header cells
 */

public class ESRHeaderRow implements ESRRow {

    private List<ESRCell> cellsList = new ArrayList<ESRCell>();

    /**
     * Get header cells of header row
     * @return
     */
    public List<ESRCell> getCells() {
        return cellsList;
    }

    /**
     * Add header cell to header row.
     * Notice that EXCEL 2003 allows only 256 columns in one sheet.
     * @param cell Header cell
     * @throws ESRException
     */
    public void addCell(ESRCell cell) throws ESRException {
        if (!(cell instanceof ESRHeaderCell)) {
            throw new ESRException(ESRHeaderCell.class.getName(),
                    "You can only add ESRHeaderCell to ESRHeaderRow.");
        }
        if (cellsList.size() == 256) {
            throw new ESRException(ESRHeaderCell.class.getName(),
                    "EXCEL 2003 does not allows add more than 256 columns");
        }
        cellsList.add(cell);
    }
}
