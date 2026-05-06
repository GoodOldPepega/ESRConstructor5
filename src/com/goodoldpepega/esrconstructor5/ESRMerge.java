package com.goodoldpepega.esrconstructor5;

import java.util.ArrayList;
import java.util.List;

/**
 * Info about merged cells in rows block
 * Created by Anton Nikitin
 * 27.02.2023
 */

public class ESRMerge {
    private List<int[]> listOfMerges = new ArrayList<int[]>();

    /**
     * Get list of merge info
     * @return
     */
    public List<int[]> getListOfMerges() {
        return listOfMerges;
    }

    /**
     * Add merge info
     * Notice that every rows block starts indexing from 0. So, for example, if second rows block in report starts from
     * 20th row of EXCEL document, first row of this block will have index of 0.
     * @param firstRow index of first row in merge
     * @param lastRow index of last row in merge
     * @param firstColumn index of first column in merge
     * @param lastColumn index of last column in merge
     */
    public void addMergeGroup(int firstRow, int lastRow, int firstColumn, int lastColumn){
        listOfMerges.add(new int[]{firstRow, lastRow, firstColumn, lastColumn});
    }
}
