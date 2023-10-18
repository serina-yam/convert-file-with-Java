package com.example.controller;

import com.example.Status;
import com.example.models.ApiResponse;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.List;

@RestController
public class ExcelMergeController {
    private static final String BASE_FILE_PATH = "/path/to/base_file.xlsx"; // ベースファイルのパス

    @PostMapping("/mergeExcel")
    public ApiResponse mergeExcelFiles(@RequestParam("filesToMerge") List<MultipartFile> filesToMerge) {
        try {
            // ベースファイルを読み込む
            Workbook baseWorkbook = new Workbook();
            baseWorkbook.loadFromFile(BASE_FILE_PATH);

            // ベースファイルに新しいシートを追加
            Worksheet baseSheet = baseWorkbook.getWorksheets().get(0);

            // マージするファイルを順番に読み込み、ベースファイルに追加
            for (MultipartFile file : filesToMerge) {
                File tempFile = File.createTempFile("temp_", ".xlsx");
                file.transferTo(tempFile);

                Workbook mergeWorkbook = new Workbook();
                mergeWorkbook.loadFromFile(tempFile.getAbsolutePath());
                baseSheet.copyFrom(mergeWorkbook.getWorksheets().get(0));
            }

            // マージ後のファイルを保存
            String mergedFilePath = "/path/to/merged_file.xlsx"; // 保存先のパス
            baseWorkbook.saveToFile(mergedFilePath);

            // 一時ファイルを削除
            for (MultipartFile file : filesToMerge) {
                File tempFile = new File(file.getOriginalFilename());
                tempFile.delete();
            }

            return new ApiResponse(Status.OK.getValue(), "Excelファイルが正常に結合され、保存されました。", BASE_FILE_PATH);
        } catch (IOException e) {
            e.printStackTrace();
            return new ApiResponse(Status.NG.getValue(), "Excelファイルの結合中にエラーが発生しました。", null);
        }
    }
}
