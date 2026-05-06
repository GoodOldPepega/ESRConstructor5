package com.goodoldpepega.esrconstructor5;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Block of table rows
 */


public class ESRTableBlock implements ESRBlock{
    private List<ESRRow> rows = new ArrayList<ESRRow>();
    private ESRMerge merge;

    /**
     * Get rows of block
     * @return
     */
    public List<ESRRow> getRows() {
        return rows;
    }

    /**
     * Add rows to block
     * @param row
     */
    public void addRow(ESRRow row) {
        rows.add(row);
    }

    /**
     * Get cells merge regions
     * @return
     */
    public ESRMerge getMerge() {
        return merge;
    }

    /**
     * Set cells merge region
     * @param esrMerge
     */
    public void setMerge(ESRMerge esrMerge) {
        this.merge = esrMerge;
    }
}
