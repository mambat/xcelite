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

import com.ebay.xcelite.sheet.XceliteSheet;
import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.Iterator;
import java.util.List;

import static com.ebay.xcelite.column.ColumnsIdentifier.*;

/**
 * Class description...
 *
 * @author kharel (kharel@ebay.com)
 * @creation_date Nov 11, 2013
 */
public abstract class SheetReaderAbs<T> implements SheetReader<T> {

    protected final XceliteSheet sheet;
    protected final List<RowPostProcessor<T>> rowPostProcessors;
    protected boolean skipHeader;

    public SheetReaderAbs(XceliteSheet sheet, boolean skipHeader) {
        this.sheet = sheet;
        this.skipHeader = skipHeader;
        rowPostProcessors = Lists.newArrayList();
    }

    protected boolean isBlankRow(Row row) {
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Object value = readValueFromCell(cellIterator.next());

            if (value != null && !String.valueOf(value).isEmpty())
                return false;
        }
        return true;
    }

    protected Object readValueFromCell(Cell cell) {
        return readValueFromCell(cell, null);
    }

    protected Object readValueFromCell(Cell cell, Class<?> type) {
        if (cell == null) return null;

        Object cellValue = null;

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (isStringByFeildType(type)) {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cellValue = cell.getStringCellValue();
                } else {
                    cellValue = cell.getNumericCellValue();
                }
                break;
            default:
                cellValue = cell.getStringCellValue();
        }
        return cellValue;
    }

    @Override
    public void skipHeaderRow(boolean skipHeaderRow) {
        this.skipHeader = skipHeaderRow;
    }

    @Override
    public XceliteSheet getSheet() {
        return sheet;
    }

    @Override
    public void addRowPostProcessor(RowPostProcessor<T> rowPostProcessor) {
        rowPostProcessors.add(rowPostProcessor);
    }

    @Override
    public void removeRowPostProcessor(RowPostProcessor<T> rowPostProcessor) {
        rowPostProcessors.remove(rowPostProcessor);
    }
}
