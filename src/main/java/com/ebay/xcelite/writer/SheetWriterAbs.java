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

import com.ebay.xcelite.sheet.XceliteSheet;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Date;

import static com.ebay.xcelite.column.ColumnsIdentifier.*;

/**
 * Class description...
 *
 * @author kharel (kharel@ebay.com)
 * @creation_date Nov 10, 2013
 */
public abstract class SheetWriterAbs<T> implements SheetWriter<T> {

    protected XceliteSheet sheet;
    protected boolean writeHeader;

    public SheetWriterAbs(XceliteSheet sheet, boolean writeHeader) {
        this.sheet = sheet;
        this.writeHeader = writeHeader;
    }

    protected void writeToCell(Cell cell, Object fieldValueObj, Class<?> dataType) {
        Class<?> type = fieldValueObj.getClass();
        if (dataType != null) {
            type = dataType;
        }
        if (isDateByFieldType(type)) {
            cell.setCellValue((Date) fieldValueObj);
        } else if (isBooleanByFieldType(type)) {
            cell.setCellValue((Boolean) fieldValueObj);
        } else if (isNumericByFieldType(type)) {
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            cell.setCellValue(Double.valueOf(fieldValueObj.toString()));
        } else {
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(fieldValueObj.toString());
        }
    }

    @Override
    public void generateHeaderRow(boolean generateHeaderRow) {
        this.writeHeader = generateHeaderRow;
    }

    @Override
    public XceliteSheet getSheet() {
        return sheet;
    }
}
