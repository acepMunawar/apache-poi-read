package com.git.apache.poi.service;

import com.git.apache.poi.util.UploadUtil;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.Supplier;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

@Service
public class uploadServiceExperiment {
    @Autowired
    UploadUtil uploadUtil;
    public void upload(MultipartFile file) throws Exception{
        Path tempDir = Files.createTempDirectory("");
        File tempFile = tempDir.resolve(file.getOriginalFilename()).toFile();
        file.transferTo(tempFile);
        Workbook workbook = WorkbookFactory.create(tempFile);
        Sheet sheet = workbook.getSheetAt(0);

//        get Row with util
        Supplier<Stream<Row>> rowStreamSupplier = uploadUtil.getRowStreamSupplier(sheet);
//        get Row
//        Stream<Row> rowStream = StreamSupport.stream(sheet.spliterator(), false);

//        get Row
//        Row headerRow = rowStream.findFirst().get();
        Row headerRow = rowStreamSupplier.get().findFirst().get();

//        get cell
        List<String> headerCells = StreamSupport.stream(headerRow.spliterator(), false)
                .map(Cell::getStringCellValue).collect(Collectors.toList());
        System.out.println(headerCells);


        rowStreamSupplier.get().forEach(row -> {
//            given a row, get cell stream from it ?
            Stream<Cell> cellStream = StreamSupport.stream(row.spliterator(), false);

            List<String> cellVals = cellStream.map(cell -> {
                String cellVal = cell.getStringCellValue();
                return cellVal;
            }).collect(Collectors.toList());
//            read file excel
            System.out.println(cellVals);
        });
    }
}
