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
package com.ebay.xcelite.styles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;

public final class CellStyles {

    private final String DEFAULT_DATE_FORMAT = "ddd mmm dd hh:mm:ss yyy";

    private final Workbook wb;
    private CellStyle headerStyle;
    private CellStyle normalStyle;
    private CellStyle dateStyle;

    public CellStyles(Workbook wb) {
        this.wb = wb;
        initStyles();
    }

    private void initStyles() {
        createHeaderStyle();
        createDateFormatStyle();
        createNormalStyle();
    }

    private void addAlignmentStyle(CellStyle style) {
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
    }

    private void addColorStyle(CellStyle style) {
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
    }

    private void addBorderStyle(CellStyle style) {
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
    }

    private void addFontStyle(CellStyle style, short boldweight) {
        Font font = wb.createFont();
        font.setFontName("宋体");
        font.setBoldweight(boldweight);
        style.setFont(font);
    }

    private void createNormalStyle() {
        normalStyle = wb.createCellStyle();
        addFontStyle(normalStyle, Font.BOLDWEIGHT_NORMAL);
        addAlignmentStyle(normalStyle);
        addBorderStyle(normalStyle);
    }

    private void createHeaderStyle() {
        headerStyle = wb.createCellStyle();
        addFontStyle(headerStyle, Font.BOLDWEIGHT_BOLD);
        addAlignmentStyle(headerStyle);
        addBorderStyle(headerStyle);
        addColorStyle(headerStyle);
    }

    private void createDateFormatStyle() {
        dateStyle = wb.createCellStyle();
        DataFormat df = wb.createDataFormat();
        dateStyle.setDataFormat(df.getFormat(DEFAULT_DATE_FORMAT));

        addAlignmentStyle(dateStyle);
        addBorderStyle(dateStyle);
        addColorStyle(dateStyle);
    }

    public CellStyle getHeaderStyle() {
        return headerStyle;
    }

    public CellStyle getDateStyle() {
        return dateStyle;
    }

    public CellStyle getCustomDataFormatStyle(String dataFormat) {
        CellStyle cellStyle = wb.createCellStyle();
        DataFormat df = wb.createDataFormat();
        cellStyle.setDataFormat(df.getFormat(dataFormat));
        return cellStyle;
    }

    public CellStyle getNormalStyle() {
        return normalStyle;
    }

    public Workbook getWorkbook() {
        return wb;
    }
}
