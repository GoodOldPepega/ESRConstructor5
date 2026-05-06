package com.goodoldpepega.esrconstructor5;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 */


public class ESRTableRow implements ESRRow {
    private List<ESRCell> cellsList = new ArrayList<ESRCell>();

    /**
     * Get cells of row
     * @return
     */
    public List<ESRCell> getCells() {
        return cellsList;
    }

    /**
     * Add cell to row.
     * Notice that row in EXCEL 2003 can have only 256 columns
     * @param cell Cell
     * @throws ESRException
     */
    public void addCell(ESRCell cell) throws ESRException {
        if (!(cell instanceof ESRTableCell)) {
            throw new ESRException(ESRTableRow.class.getName(),
                    "You can add only add ESRTableCell to ESRTableRow.");
        }
        if (cellsList.size() == 256) {
            throw new ESRException(ESRTableRow.class.getName(),
                    "EXCEL 2003 does not allows add more than 256 columns");
        }
        cellsList.add(cell);
    }
}
