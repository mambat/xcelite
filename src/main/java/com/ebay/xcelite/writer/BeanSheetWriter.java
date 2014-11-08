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
package com.ebay.xcelite.writer;

import com.ebay.xcelite.annotate.NoConverterClass;
import com.ebay.xcelite.column.Col;
import com.ebay.xcelite.column.ColumnFieldsMapper;
import com.ebay.xcelite.column.ColumnsExtractor;
import com.ebay.xcelite.converters.ColumnValueConverter;
import com.ebay.xcelite.sheet.XceliteSheet;
import com.ebay.xcelite.styles.CellStylesBank;
import com.google.common.collect.Sets;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.*;

public class BeanSheetWriter<T> extends SheetWriterAbs<T> {

    private final LinkedHashSet<Col> columns;
    private final ColumnFieldsMapper fieldsMapper;
    private final Col anyColumn;
    private Row headerRow;
    private int rowIndex = 0;

    public BeanSheetWriter(XceliteSheet sheet, Class<T> type) {
        super(sheet, true);
        ColumnsExtractor extractor = new ColumnsExtractor(type);
        extractor.extract();
        columns = extractor.getColumns();
        fieldsMapper = new ColumnFieldsMapper(extractor.getColumnFields());
        anyColumn = extractor.getAnyColumn();
    }

    @Override
    public void write(Collection<T> data) {
        if (writeHeader) writeHeader();

        writeData(data);

        autoSizeColumn();
    }

    private void autoSizeColumn() {
        Sheet nativeSheet = sheet.getNativeSheet();
        Row row = nativeSheet.getRow(0);

        for (int i = 0; i < row.getLastCellNum(); i++) {
            nativeSheet.autoSizeColumn(i);
        }

    }

    @SuppressWarnings("unchecked")
    private void writeData(Collection<T> data) {
        try {
            Set<Col> columnsToAdd = Sets.newTreeSet();

            if (anyColumn != null) {
                for (T t : data)
                    appendAnyColumns(t, columnsToAdd);
            }

            addColumns(columnsToAdd, true);
            for (T t : data) {
                Row row = sheet.getNativeSheet().createRow(rowIndex);
                int i = 0;
                for (Col col : columns) {
                    Field field = fieldsMapper.getColumnField(col.getFieldName());
                    Object fieldValueObj = null;
                    if (col.isAnyColumn()) {
                        Map<String, Object> anyColumnMap = (Map<String, Object>) field.get(t);
                        fieldValueObj = anyColumnMap.get(col.getName());
                    } else {
                        fieldValueObj = field.get(t);
                    }
                    Cell cell = row.createCell(i);
                    cell.setCellStyle(CellStylesBank.get(sheet.getNativeSheet().getWorkbook()).getNormalStyle());
                    writeToCell(cell, col, fieldValueObj);
                    i++;
                }
                rowIndex++;
            }
        } catch (SecurityException e) {
            throw new RuntimeException(e);
        } catch (IllegalArgumentException e) {
            throw new RuntimeException(e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    @SuppressWarnings("unchecked")
    private void writeToCell(Cell cell, Col col, Object fieldValueObj) {
        if (fieldValueObj == null) {
            cell.setCellValue((String) null);
            return;
        }
        if (col.getConverter() != null) {
            try {
                ColumnValueConverter<?, Object> converter = (ColumnValueConverter<?, Object>) col.getConverter().newInstance();
                fieldValueObj = converter.serialize(fieldValueObj);
            } catch (InstantiationException e) {
                throw new RuntimeException(e);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }
        if (col.getDataFormat() != null) {
            cell.setCellStyle(CellStylesBank.get(sheet.getNativeSheet().getWorkbook()).getCustomDataFormatStyle(
                    col.getDataFormat()));
        }

        if (col.getType() == Date.class) {
            if (col.getDataFormat() == null) {
                cell.setCellStyle(CellStylesBank.get(sheet.getNativeSheet().getWorkbook()).getDateStyle());
            }
        }

        writeToCell(cell, fieldValueObj, col.getType());
    }

    private void writeHeader() {
        headerRow = sheet.getNativeSheet().createRow(rowIndex);
        rowIndex++;
        addColumns(columns, false);
    }

    @SuppressWarnings("unchecked")
    private void appendAnyColumns(T t, Set<Col> columnToAdd) {
        try {
            Field anyColumnField = fieldsMapper.getColumnField(anyColumn.getFieldName());
            Map<String, Object> fieldValueObj = (Map<String, Object>) anyColumnField.get(t);

            for (Map.Entry<String, Object> entry : fieldValueObj.entrySet()) {
                Col column = new Col(entry.getKey(), anyColumnField.getName());
                column.setType(entry.getValue() == null ? String.class : entry.getValue().getClass());
                column.setAnyColumn(true);
                if (anyColumn.getConverter() != NoConverterClass.class) {
                    column.setConverter(anyColumn.getConverter());
                }
                columnToAdd.add(column);
            }
        } catch (SecurityException e) {
            throw new RuntimeException(e);
        } catch (IllegalArgumentException e) {
            throw new RuntimeException(e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    private void addColumns(Set<Col> columnsToAdd, boolean append) {
        int i = (headerRow == null || headerRow.getLastCellNum() == -1) ? 0 : headerRow.getLastCellNum();
        for (Col column : columnsToAdd) {
            if (append && columns.contains(column))
                continue;
            if (writeHeader) {
                Cell cell = headerRow.createCell(i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellStyle(CellStylesBank.get(sheet.getNativeSheet().getWorkbook()).getHeaderStyle());
                cell.setCellValue(column.getName());
                i++;
            }
            columns.add(column);
        }
    }
}
