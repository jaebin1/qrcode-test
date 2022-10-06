package com.example.demo;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

@RestController
public class TestController {

    @GetMapping("/")
    public void qrCreate(){
        QRCodeWriter q = new QRCodeWriter();
        String path = "C:\\test";

        try {
            String text = "http://google.com";
            text = new String(text.getBytes("UTF-8"), "ISO-8859-1");
            BitMatrix bitMatrix = q.encode(text, BarcodeFormat.QR_CODE,200,200);

//            File imgFile = new File(path+"\\qrcode.png");
//            FileOutputStream qrImg = new FileOutputStream(imgFile);
            ByteArrayOutputStream pngOutputStream = new ByteArrayOutputStream();
            MatrixToImageWriter.writeToStream(bitMatrix, "png", pngOutputStream);
            byte[] bytes = pngOutputStream.toByteArray();
            // 엑셀
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("sample");

            //파일 읽기
//            InputStream inputStream = new FileInputStream(path+"\\qrcode.png");
//            byte[] bytes = IOUtils.toByteArray(inputStream);

            //Adds a picture to the workbook
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
//            inputStream.close();
            pngOutputStream.close();

            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = wb.getCreationHelper();

            //Creates the top-level drawing patriarch.
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();
            //set top-left corner for the image
            anchor.setCol1(0);
            anchor.setRow1(0);

            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //Reset the image to the original size
            pict.resize();

            //Write the Excel file
            FileOutputStream fileOut = null;
            fileOut = new FileOutputStream("C:\\test\\qrcode.xlsx");
            wb.write(fileOut);
            fileOut.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

