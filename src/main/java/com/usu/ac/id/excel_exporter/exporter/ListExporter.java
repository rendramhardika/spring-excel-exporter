package com.usu.ac.id.excel_exporter.exporter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ListExporter {
    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;
    private List<HashMap<String, Object>> dataList;

    public ListExporter(List<HashMap<String, Object>> dataList) {
        this.dataList = dataList;
        workbook = new SXSSFWorkbook();
    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.trackColumnForAutoSizing(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).doubleValue());
        } else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }

    private void writeHeaderLine(){
        sheet = workbook.createSheet("Billings");

        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(Short.parseShort("14"));
        style.setFont(font);
        int columnCount = 0;

        createCell(row, columnCount++, "Full Name", style);
        createCell(row, columnCount++, "Age", style);
        createCell(row, columnCount, "Major", style);
    }

    private void writeDataLines(){
        final int[] rowCount = {1};
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints(Short.parseShort("14"));
        style.setFont(font);

        dataList.forEach(data->{
            Row row = sheet.createRow(rowCount[0]++);
            int columnCount = 0;

            createCell(row, columnCount++, data.get("name").toString(), style);
            createCell(row, columnCount++, data.get("age"), style);
            createCell(row, columnCount++, data.get("major").toString(), style);
        });
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeDataLines();

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();

        outputStream.close();
    }
}
