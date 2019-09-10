package core;

import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGImageEncoder;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ScaleImage {

    public static void scaleImage(String path, int width, int height) throws IOException {
        BufferedImage img = ImageIO.read(new File(path));
        BufferedImage newImg = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        newImg.getGraphics().drawImage(img.getScaledInstance(width, height, Image.SCALE_AREA_AVERAGING), 0, 0, null);
        FileOutputStream out = new FileOutputStream(getNewPath(path, "scale"));
        JPEGImageEncoder encoder = JPEGCodec.createJPEGEncoder(out);
        encoder.encode(newImg);
        out.close();
    }



    private static String getNewPath(String path, String opreation) {
        File file = new File(path);
        String name = file.getName();
        String[] names = name.split("\\.");
        String newName = names[0] + "_" + opreation + "." + names[1];
        return newName;
    }
}
