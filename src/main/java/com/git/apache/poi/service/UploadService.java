package com.git.apache.poi.service;

import com.git.apache.poi.util.UploadUtil;
import org.apache.poi.ss.formula.atp.Switch;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.function.Supplier;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static java.util.stream.Collectors.toMap;

@Service
public class UploadService{
    @Autowired
    UploadUtil uploadUtil;

public List<Map<String, Object>> readColumn(MultipartFile file) {
    List<Map<String,Object>> listData = new ArrayList<>();
    try {
        Path tempDir = Files.createTempDirectory("");
        File tempFile = tempDir.resolve(file.getOriginalFilename()).toFile();
        file.transferTo(tempFile);
        Workbook workbook = WorkbookFactory.create(tempFile);
        Sheet sheet = workbook.getSheetAt(0);

        Map<String, Integer> columns = new HashMap<>();
        sheet.getRow(0).forEach(cell ->{

            columns.put(cell.getStringCellValue(), cell.getColumnIndex());
        });

        for(Row row:sheet){
            if(!row.getCell(columns.get("FirstName")).getStringCellValue().equalsIgnoreCase("FirstName") ||
                    !row.getCell(columns.get("LastName")).getStringCellValue().equalsIgnoreCase("LastName") ||
                    !row.getCell(columns.get("Level")).getStringCellValue().equalsIgnoreCase("Level")){
                Map<String,Object> data=new HashMap<>();
                data.put("FistName",row.getCell(columns.get("FirstName")).getStringCellValue());
                data.put("LastName",row.getCell(columns.get("LastName")).getStringCellValue());
                data.put("Level",row.getCell(columns.get("Level")).getNumericCellValue());

                listData.add(data);
            }


        }
    } catch (IOException ex) {
        ex.printStackTrace();

    }
    return listData;
}

    public List<String> readCellValueWithList(MultipartFile file){
    List<String> listData = new LinkedList<>();
        try{
            Path tempDir = Files.createTempDirectory("");
            File tempFile = tempDir.resolve(file.getOriginalFilename()).toFile();
            file.transferTo(tempFile);
            Workbook workbook = WorkbookFactory.create(tempFile);
            Sheet sheet = workbook.getSheetAt(0);
            CellReference cellReference = new CellReference("B1");
            Row row =sheet.getRow(cellReference.getRow());
            Cell cell= row.getCell(cellReference.getCol());
            listData.add(String.valueOf(cell));
        }catch(IOException ex){
            ex.printStackTrace();
        }
        return listData;
    }

    public List<Map<String, String>> readWithStream(MultipartFile file) throws Exception{
        Path tempDir = Files.createTempDirectory("");
        File tempFile = tempDir.resolve(file.getOriginalFilename()).toFile();
        file.transferTo(tempFile);
        Workbook workbook = WorkbookFactory.create(tempFile);
        Sheet sheet = workbook.getSheetAt(0);

        Supplier<Stream<Row>> rowStreamSupplier = uploadUtil.getRowStreamSupplier(sheet);

        Row headerRow = rowStreamSupplier.get().findFirst().get();

        List<String> headerCells = uploadUtil.getStream(headerRow)
                .map(Cell::getStringCellValue)
                .collect(Collectors.toList());
        System.out.println(headerCells);


        int colCount = headerCells.size();

        return rowStreamSupplier.get()
                    .skip(1)
                    .map(row -> {
                    List<String> cellList = uploadUtil.getStream(row)
                            .map(Cell::getStringCellValue)
                            .collect(Collectors.toList());
            return uploadUtil.cellIteratorSupplier(colCount)
                    .get()
                    .collect(toMap(
                        headerCells::get, cellList::get
                    ));
        }).collect(Collectors.toList());
    }
}