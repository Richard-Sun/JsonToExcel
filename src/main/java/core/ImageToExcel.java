package core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ImageToExcel {

    public static void imageToExcel(String excelPath, String picPath, int startRow, int startColumn) throws IOException {

        XSSFCellStyle[][][] csTable = new XSSFCellStyle[32][32][32];
        XSSFWorkbook wb = new XSSFWorkbook();
        OutputStream out = new FileOutputStream(excelPath);
        Sheet st = wb.createSheet();
        BufferedImage image = ImageIO.read(new File(picPath));

        for (int r = 0; r < 16; r++) {
            for (int g = 0; g < 16; g++) {
                for (int b = 0; b < 16; b++) {
                    csTable[r][g][b] = wb.createCellStyle();
                    int red = r * 16 + 4;
                    int green = g * 16 + 4;
                    int blue = b * 16 + 4;
                    csTable[r][g][b].setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    XSSFColor color = new XSSFColor(new java.awt.Color(red, green, blue));
                    csTable[r][g][b].setFillForegroundColor(color);
                }
            }
        }

        int width = image.getWidth();
        int height = image.getHeight();

        for (int i = 0; i < height; i++) {
            Row row = st.createRow(startRow + i);
            row.setHeight((short) 50);
            for (int j = 0; j < width; j++) {
                Cell cell = row.createCell(startColumn + j);
                int rgb = image.getRGB(j, i);
                int red = (rgb & 16711680) >> 16;
                int green = (rgb & 65280) >> 8;
                int blue = (rgb & 255);
                red = red / 16;
                green = green / 16;
                blue = blue / 16;
                cell.setCellStyle(csTable[red][green][blue]);
            }
        }
        for (int i = 0; i < width; i++) {
            st.setColumnWidth(i + startColumn, 100);
        }
        wb.write(out);
        out.close();
    }
}
