package com.hdfc.reportgeneration.service;

import com.hdfc.reportgeneration.helper.ExcelHelper;
import org.apache.commons.compress.utils.FileNameUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;

import java.util.HashMap;

import java.util.Map;


@Service
public class ExcelUploadService {


    public ResponseEntity<ByteArrayResource> readExcel(MultipartFile file) {
        Map<String,Object> response=new HashMap<>();
        String fileName=FileNameUtils.getExtension(file.getOriginalFilename());
        ExcelHelper excelHelper=new ExcelHelper();
        try {
            XSSFWorkbook xssfWorkbook = excelHelper.excelUpload(file.getInputStream(), fileName);
            ByteArrayOutputStream stream = new ByteArrayOutputStream();
            HttpHeaders header = new HttpHeaders();
            header.setContentType(new MediaType("application", "force-download"));
            header.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ProductTemplate.xlsx");
            xssfWorkbook.write(stream);
            xssfWorkbook.close();
            return new ResponseEntity<>(new ByteArrayResource(stream.toByteArray()),
                    header, HttpStatus.CREATED);
        } catch (Exception e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }
}
