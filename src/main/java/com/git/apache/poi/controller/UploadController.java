package com.git.apache.poi.controller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import com.git.apache.poi.service.UploadService;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

@RestController
public class UploadController {

    @Autowired
    UploadService uploadService;

    @RequestMapping("/readAllCellValue")
    public List<Map<String, Object>> ReadColumn(@RequestParam("file")MultipartFile file){
        return uploadService.readColumn(file);
    }

    @RequestMapping("/readCellValue")
    public List<String> ReadCellValueWithList(@RequestParam("file")MultipartFile file){
        return uploadService.readCellValueWithList(file);
    }


    @PostMapping("/readWithStream")
    public List<Map<String, String>> ReadWithStream(@RequestParam("file")MultipartFile file) throws Exception{
        return uploadService.readWithStream(file);
    }
}
