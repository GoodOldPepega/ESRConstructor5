package com.goodoldpepega.esrconstructor5;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 *
 * Report has multiple blocks of rows.
 * Also report has title block and header rows block.
 */

public class ESRReport {
    private ESRHeaderBlock esrHeaderBlock;
    private List<ESRBlock> esrBlocksList = new ArrayList<ESRBlock>();
    private ESRTitleBlock esrTitleBlock;

    /**
     * Get header block
     * @return
     */
    public ESRHeaderBlock getEsrHeaderBlock() {
        return esrHeaderBlock;
    }

    /**
     * Set header block
     * @param esrHeaderBlock
     */
    public void setEsrHeaderBlock(ESRHeaderBlock esrHeaderBlock) {
        this.esrHeaderBlock = esrHeaderBlock;
    }

    /**
     * Get table blocks
     * @return
     */
    public List<ESRBlock> getEsrBlocksList() {
        return esrBlocksList;
    }

    /**
     * Add block to table blocks list
     * @param esrBlock
     */
    public void addBlock(ESRBlock esrBlock) {
        esrBlocksList.add(esrBlock);
    }

    /**
     * Get title block
     * @return
     */
    public ESRTitleBlock getEsrTitleBlock() {
        return esrTitleBlock;
    }

    /**
     * Set title block
     * @param esrTitleBlock
     */
    public void setEsrTitleBlock(ESRTitleBlock esrTitleBlock) {
        this.esrTitleBlock = esrTitleBlock;
    }
}
