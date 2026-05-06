package com.goodoldpepega.esrconstructor5;

import java.util.List;

/**
 * Report block interface
 */

public interface ESRBlock {

    /**
     * Get rows of block in report
     * @return
     */
    List<ESRRow> getRows();

    /**
     * Add rows to block in report
     * @param row
     * @throws ESRException
     */
    void addRow(ESRRow row) throws ESRException;

    /**
     * Get merged regions info
     * @return
     */
    ESRMerge getMerge();

    /**
     * Set merged regions info
     * @param esrMerge
     */
    void setMerge(ESRMerge esrMerge);
}
