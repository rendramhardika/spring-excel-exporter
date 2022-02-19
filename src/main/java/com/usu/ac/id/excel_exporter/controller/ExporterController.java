package com.usu.ac.id.excel_exporter.controller;

import com.usu.ac.id.excel_exporter.data.DataList;
import com.usu.ac.id.excel_exporter.exporter.BillingExporter;
import com.usu.ac.id.excel_exporter.exporter.ListExporter;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.client.RestTemplate;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

@RestController
@CrossOrigin(origins = "*")
public class ExporterController {

    @GetMapping("/billings_api_to_excel")
    public void generateExcelFromApiBilling(HttpServletResponse response) throws IOException {
        RestTemplate restTemplate = new RestTemplate();
        HttpHeaders headers = new HttpHeaders();
        headers.setAccept(Arrays.asList(MediaType.APPLICATION_JSON));
        HttpEntity<String> entity = new HttpEntity<>(headers);
        String request = restTemplate.exchange("https://tagihan.usu.ac.id/api/billing?page=1", HttpMethod.GET, entity, String.class).getBody();

        JSONObject result = new JSONObject(request);
        JSONObject data = result.getJSONObject("data");

        JSONArray billings = data.getJSONArray("billings");

        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=Billings.xlsx";
        response.setHeader(headerKey, headerValue);

        BillingExporter billingExporter = new BillingExporter(billings);
        billingExporter.export(response);
    }

    @GetMapping("/datalist_to_excel")
    public void generateExcelFromDataList(HttpServletResponse response) throws IOException {

        List<HashMap<String, Object>> dataList = new ArrayList<>();
        dataList.addAll(DataList.getDataList());

        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=Billings.xlsx";
        response.setHeader(headerKey, headerValue);

        ListExporter listExporter = new ListExporter(dataList);
        listExporter.export(response);

    }


}
