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

public class BillingExporter {
    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;
    private JSONArray billings;


    public BillingExporter(JSONArray billings) {
        this.billings = billings;
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

        createCell(row, columnCount++, "Transaction ID", style);
        createCell(row, columnCount++, "Virtual Account", style);
        createCell(row, columnCount++, "Name", style);
        createCell(row, columnCount++, "Total Amount", style);
        createCell(row, columnCount++, "Created Date", style);
        createCell(row, columnCount++, "Expired Date", style);
        createCell(row, columnCount++, "Status", style);
        createCell(row, columnCount++, "Payment Date", style);
    }

    private void writeDataLines(){
        final int[] rowCount = {1};
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints(Short.parseShort("14"));
        style.setFont(font);

        for (int i=0; i<=billings.length()-1; i++){
            JSONObject bill = billings.getJSONObject(i);
            JSONArray itemList = bill.getJSONArray("items");

            Row row = sheet.createRow(rowCount[0]++);
            int columnCount = 0;

            createCell(row, columnCount++, bill.get("transaction_id").toString(), style);
            createCell(row, columnCount++, bill.get("virtual_account").toString(), style);
            createCell(row, columnCount++, bill.get("customer_name").toString(), style);
            createCell(row, columnCount++, new BigDecimal(bill.get("total_amount").toString()), style);
            createCell(row, columnCount++, bill.get("datetime_created_iso8601").toString(), style);
            createCell(row, columnCount++, bill.get("datetime_expired_iso8601").toString(), style);
            createCell(row, columnCount++, bill.get("billing_status"), style);
            createCell(row, columnCount, bill.get("datetime_payment_iso8601"), style);

        }
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
