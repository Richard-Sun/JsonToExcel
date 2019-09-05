package core;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class JsonToExcel {

    public static void jsonToExcel(String path, JSONArray array, int startRow, int startColumn) throws IOException {

        boolean initFlag = false;
        Workbook wb;
        if (path.contains(".xlsx")) {
            wb = new XSSFWorkbook();
        } else {
            wb = new HSSFWorkbook();
        }
        Sheet st1 = wb.createSheet();
        FileOutputStream out = new FileOutputStream(path);
        List<String> titles = null;

        for (int i = 0; i < array.size(); i++) {
            JSONObject ob = array.getJSONObject(i);
            if (!initFlag) {
                titles = init(st1, ob, startRow, startColumn);
                initFlag = true;
            }
            fillRow(st1, titles, ob, startRow + i + 1, startColumn);
        }
        adjustColumn(st1, titles, startColumn);
        wb.write(out);
        out.close();
    }

    private static List<String> init(Sheet st, JSONObject object, int startRow, int startColumn) {
        List<String> seq = new ArrayList<>();
        Row row1 = st.createRow(startRow);
        Map<String, Object> map = object.getInnerMap();
        int i = startColumn;
        for (String column : map.keySet()) {
            seq.add(column);
            Cell cell = row1.createCell(i);
            cell.setCellValue(column);
            i++;
        }
        return seq;
    }

    public static void fillRow(Sheet st, List<String> titles, JSONObject ob, int startRow, int startColumn) {
        Row row = st.createRow(startRow);
        for (int i = 0; i < titles.size(); i++) {
            Cell cell = row.createCell(startColumn + i);
            cell.setCellValue(ob.getString(titles.get(i)));
        }
    }

    private static void adjustColumn(Sheet st1, List<String> titles, int startColumn) {
        for (int i = 0; i < titles.size(); i++) {
            st1.autoSizeColumn(i + startColumn);
        }
    }

}
