package com.goodoldpepega.esrconstructor5;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Class of header rows block.
 */


public class ESRHeaderBlock implements ESRBlock {
    private List<ESRRow> rows = new ArrayList<ESRRow>();
    private ESRMerge merge;
    private int maxHeaderRowLength = 0;

    /**
     * Get rows of header block
     * @return
     */
    public List<ESRRow> getRows() {
        return rows;
    }

    /**
     * Add header row to header block
     * @param row
     * @throws ESRException
     */
    public void addRow(ESRRow row) throws ESRException {
        if (!(row instanceof ESRHeaderRow)) {
            throw new ESRException(ESRHeaderRow.class.getName(),
                    "You can add only add ESRHeaderCell to ESRHeaderRow.");
        }
        if (maxHeaderRowLength < row.getCells().size()) {
            maxHeaderRowLength = row.getCells().size();
        }
        rows.add(row);
    }

    /**
     * Get info about merged cells in header block
     * @return
     */
    public ESRMerge getMerge() {
        return merge;
    }

    /**
     * Set info about merged cells in header block
     * @param esrMerge
     */
    public void setMerge(ESRMerge esrMerge) {
        this.merge = esrMerge;
    }

    /**
     * Get max quantity of cells of header rows in header block
     * @return
     */
    int getMaxHeaderRowLength() {
        return maxHeaderRowLength;
    }
}
