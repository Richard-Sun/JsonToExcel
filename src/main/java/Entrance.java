import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import core.ImageToExcel;
import core.JsonToExcel;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import other.Constant;

import java.io.FileInputStream;
import java.io.IOException;

public class Entrance {
    public static void main(String[] args) throws IOException {
        /*byte[] buf = new byte[20 * Constant.MB];
        FileInputStream in = new FileInputStream("text.txt");
        in.read(buf);
        JSONArray array = JSON.parseArray(new String(buf));
        JsonToExcel.jsonToExcel("1.xlsx", array, 0, 0);*/

        ImageToExcel.imageToExcel("2.xlsx","2.jpg",0,0);
    }


}
