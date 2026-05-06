package com.goodoldpepega.esrconstructor5;

/**
 * Created by Anton Nikitin
 * 04.03.2023
 */


class ESRException extends Exception {

    /**
     *
     * @param classNameStr
     * @param exceptionDetails
     */
    ESRException (String classNameStr, String exceptionDetails) {
        super(classNameStr + " : " + exceptionDetails);
    }
}
