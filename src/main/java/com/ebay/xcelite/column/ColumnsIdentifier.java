package com.ebay.xcelite.column;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.util.Date;

/**
 * FIXME.
 *
 * @author wanglei
 * @since 14-11-8 下午11:38
 */
public class ColumnsIdentifier {

    public static boolean isNumericByFieldType(Class<?> type) {
        return type == Double.class || type == double.class ||
               type == Integer.class || type == int.class ||
               type == Long.class || type == long.class ||
               type == Float.class || type == float.class ||
               type == Short.class || type == short.class;
    }

    public static boolean isStringByFeildType(Class<?> type) {
        return type == String.class;
    }

    public static boolean isBooleanByFieldType(Class<?> type) {
        return type == Boolean.class || type == boolean.class;
    }

    public static boolean isDateByFieldType(Class<?> type) {
        return type == Date.class;
    }

    public static boolean isDateByCell(Cell cell) {
        return DateUtil.isCellDateFormatted(cell);
    }
}
