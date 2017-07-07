package com.wuage.file.converter.controller;

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import com.sun.xml.internal.messaging.saaj.util.ByteOutputStream;
import com.wuage.file.converter.ConverterImpl;

@Controller
@RequestMapping(value = "/file")
public class ConverterController {

    private ConverterImpl converter = new ConverterImpl("127.0.0.1:8100");

    @RequestMapping(value = "/converter", method = { RequestMethod.GET, RequestMethod.POST })
    public void converter(@RequestParam("url") String url, HttpServletResponse response) {

        OutputStream os = null;
        ByteOutputStream outputStream = null;
        response.setContentType("application/pdf");
        try {
            os = new BufferedOutputStream(response.getOutputStream());
            response.setContentType("application/pdf");
            // TODO 模拟
            String filePath = "D:/测试数据/29186947397435984.xlsx";
            InputStream inputStream = new FileInputStream(filePath);
            outputStream = converter.converter2Pdf(inputStream, "xlsx", inputStream.available());
            os.write(outputStream.getBytes());
            os.flush();
        } catch (Exception e) {
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (Exception e) {
                }
            }
            if (os != null) {
                try {
                    os.close();
                } catch (Exception e) {
                }
            }
        }
    }

}
