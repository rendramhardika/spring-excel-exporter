package com.usu.ac.id.excel_exporter.data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class DataList {
    public static List<HashMap<String, Object>> getDataList(){
        List<HashMap<String, Object>> res = new ArrayList<>();

        HashMap<String, Object> data1 = new HashMap<>();
        data1.put("name", "Rendra Mahardika");
        data1.put("age", 25);
        data1.put("major", "Data Science anda Artificial Intelligence");
        res.add(data1);

        HashMap<String, Object> data2 = new HashMap<>();
        data2.put("name", "Fiqih Fatwa");
        data2.put("age", 25);
        data2.put("major", "Information Technology");
        res.add(data2);

        HashMap<String, Object> data3 = new HashMap<>();
        data3.put("name", "Ghalib Abdillah");
        data3.put("age", 22);
        data3.put("major", "Computer Science");
        res.add(data3);

        return res;
    }
}
