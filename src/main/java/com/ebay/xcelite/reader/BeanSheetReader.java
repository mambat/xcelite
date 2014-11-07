/*
  Copyright [2013-2014] eBay Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/
package com.ebay.xcelite.reader;

import com.ebay.xcelite.annotate.NoConverterClass;
import com.ebay.xcelite.annotations.AnyColumn;
import com.ebay.xcelite.column.Col;
import com.ebay.xcelite.column.ColumnFieldsMapper;
import com.ebay.xcelite.column.ColumnsExtractor;
import com.ebay.xcelite.column.ColumnsMapper;
import com.ebay.xcelite.converters.ColumnValueConverter;
import com.ebay.xcelite.exceptions.XceliteException;
import com.ebay.xcelite.sheet.XceliteSheet;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Class description...
 *
 * @author kharel (kharel@ebay.com)
 * @creation_date Sep 9, 2013
 */
public class BeanSheetReader<T> extends SheetReaderAbs<T> {

    private final LinkedHashSet<Col> columns;
    private final Col anyColumn;
    private final Field anyColumnField;
    private final ColumnsMapper mapper;
    private final ColumnFieldsMapper fieldsMapper;
    private final Class<T> type;
    private LinkedHashSet<String> header;
    private Iterator<Row> rowIterator;
    private Set<String> ignoreCols;

    public BeanSheetReader(XceliteSheet sheet, Class<T> type) {
        super(sheet, false);
        this.type = type;
        ColumnsExtractor extractor = new ColumnsExtractor(type);
        extractor.extract();

        columns = extractor.getColumns();
        mapper = new ColumnsMapper(columns);

        fieldsMapper = new ColumnFieldsMapper(extractor.getColumnFields());

        anyColumn = extractor.getAnyColumn();
        anyColumnField = extractor.getAnyColumnField();

        if (anyColumnField != null) {
            AnyColumn anyColumnAnno = anyColumnField.getAnnotation(AnyColumn.class);
            ignoreCols = Sets.newHashSet(anyColumnAnno.ignoreCols());
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public Collection<T> read() {
        buildHeader();
        List<T> data = Lists.newArrayList();
        try {
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (isBlankRow(row)) continue;
                T object = type.newInstance();

                int i = 0;
                for (String columnName : header) {
                    Cell cell = row.getCell(i, Row.RETURN_BLANK_AS_NULL);
                    Col col = mapper.getColumn(columnName);
                    if (col == null) {
                        if (anyColumn != null && !isColumnInIgnoreList(columnName)) {
                            writeToAnyColumnField(object, cell, columnName);
                        }
                    } else {
                        Field field = fieldsMapper.getColumnField(col.getFieldName());
                        writeToField(field, object, cell, col);
                    }
                    i++;
                }

                boolean keepObject = true;
                for (RowPostProcessor<T> rowPostProcessor : rowPostProcessors) {
                    keepObject = rowPostProcessor.process(object);
                    if (!keepObject) break;
                }

                if (keepObject) {
                    data.add(object);
                }
            }
        } catch (InstantiationException e) {
            throw new RuntimeException(e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        return data;
    }

    private boolean isColumnInIgnoreList(String columnName) {
        return ignoreCols.contains(columnName);
    }

    @SuppressWarnings("unchecked")
    private void writeToAnyColumnField(T object, Cell cell, String columnName) {
        try {
            anyColumnField.setAccessible(true);
            Object value = readValueFromCell(cell);
            if (value == null) return;

            AnyColumn annotation = anyColumnField.getAnnotation(AnyColumn.class);
            if (anyColumnField.get(object) == null) {
                Map<String, Object> map = (Map<String, Object>) annotation.as().newInstance();
                anyColumnField.set(object, map);
            }

            Map<String, Object> map = (Map<String, Object>) anyColumnField.get(object);
            if (annotation.converter() != NoConverterClass.class) {
                ColumnValueConverter<Object, ?> converter = (ColumnValueConverter<Object, ?>) annotation.converter()
                        .newInstance();
                value = converter.deserialize(value);
            }
            map.put(columnName, value);
        } catch (IllegalArgumentException e) {
            throw new RuntimeException(e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        } catch (InstantiationException e) {
            throw new RuntimeException(e);
        }
    }

    @SuppressWarnings("unchecked")
    private void writeToField(Field field, T object, Cell cell, Col column) {
        try {
            Object cellValue = readValueFromCell(cell);
            if (cellValue != null) {
                if (column.getConverter() != null) {
                    ColumnValueConverter<Object, ?> converter = (ColumnValueConverter<Object, ?>) column.getConverter()
                            .newInstance();
                    cellValue = converter.deserialize(cellValue);
                } else {
                    cellValue = convertToFieldType(cellValue, field.getType());
                }
            }
            field.setAccessible(true);
            field.set(object, cellValue);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        } catch (InstantiationException e) {
            throw new RuntimeException(e);
        }
    }

    private Object convertToFieldType(Object cellValue, Class<?> fieldType) {
        String value = String.valueOf(cellValue);
        if (fieldType == Double.class || fieldType == double.class) {
            return Double.valueOf(value);
        }
        if (fieldType == Integer.class || fieldType == int.class) {
            return Double.valueOf(value).intValue();
        }
        if (fieldType == Short.class || fieldType == short.class) {
            return Double.valueOf(value).shortValue();
        }
        if (fieldType == Long.class || fieldType == long.class) {
            return Double.valueOf(value).longValue();
        }
        if (fieldType == Float.class || fieldType == float.class) {
            return Double.valueOf(value).floatValue();
        }
        if (fieldType == Character.class || fieldType == char.class) {
            return value.charAt(0);
        }
        if (fieldType == Date.class) {
            return DateUtil.getJavaDate(Double.valueOf(value));
        }
        return value;
    }

    private void buildHeader() {
        header = Sets.newLinkedHashSet();
        rowIterator = sheet.getNativeSheet().rowIterator();
        Row row = rowIterator.next();
        if (row == null) {
            throw new XceliteException("First row in sheet is empty. First row must contain header");
        }
        Iterator<Cell> itr = row.cellIterator();
        while (itr.hasNext()) {
            Cell cell = itr.next();
            header.add(cell.getStringCellValue());
        }
    }
}
