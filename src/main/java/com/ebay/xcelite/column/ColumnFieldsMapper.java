package com.ebay.xcelite.column;

import com.google.common.collect.Maps;

import java.lang.reflect.Field;
import java.util.Map;
import java.util.Set;

/**
 * FIXME.
 *
 * @author wanglei
 * @since 14-11-7 上午1:30
 */
public class ColumnFieldsMapper {

    private final Map<String, Field> columnFieldsMap;

    public ColumnFieldsMapper(Set<Field> columnFields) {
        columnFieldsMap = Maps.newHashMap();
        for (Field field : columnFields) {
            columnFieldsMap.put(field.getName(), field);
        }
    }

    public Field getColumnField(String fieldName) {
        return columnFieldsMap.get(fieldName);
    }

}
