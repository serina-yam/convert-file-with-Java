package com.example.controller;
import org.springframework.web.multipart.MultipartFile;

import com.example.Status;
import com.example.model.ApiResponse;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

/**
 * POIを使ってxlsmファイルをxlsxファイルに変換する.
 */
@RestController
public class FileConversionPdfboxController {

    @PostMapping("/convertXlsmToXlsx")
    public ApiResponse convertXLSToPDF(@RequestParam("xlsmFile") MultipartFile xlsmFileMulti) {
        try {
            // 一時ディレクトリを作成
            File tempDir = new File("tempDir");
            tempDir.mkdirs();

            // xlsmファイルからマクロを削除し、xlsxに変換
            File xlsmFile = new File (xlsmFileMulti.getOriginalFilename());
            File xlsxFile = new File(tempDir, "temp.xlsx");
            removeMacrosAndConvertToXLSX(xlsmFile, xlsxFile);

            return new ApiResponse(Status.OK.getValue(), "PDF変換処理は正常終了しました。", xlsxFile.toPath().toString());
        } catch (Exception e) {
            return new ApiResponse(Status.NG.getValue(), "PDF変換処理は異常終了しました。", null);
        }
    }

    /**
     * xlsmファイルからマクロの情報を削除する.
     * @param xlsmFile
     * @param xlsxFile
     * @throws IOException
     * @throws InvalidFormatException 
     */
    private void removeMacrosAndConvertToXLSX(File xlsmFileBase, File xlsxFileBase) throws IOException, InvalidFormatException {
    	
    	  XSSFWorkbook workbook = (XSSFWorkbook)WorkbookFactory.create(new FileInputStream(xlsmFileBase));

    	  OPCPackage opcpackage = workbook.getPackage();

    	  // vbaProject.bin 部分を取得してパッケージから削除
    	  PackagePart vbapart = opcpackage.getPartsByName(Pattern.compile("/xl/vbaProject.bin")).get(0);
    	  opcpackage.removePart(vbapart);

    	  // パッケージから削除された vbaProject.bin 部分との関係を取得および削除
    	  PackagePart wbpart = workbook.getPackagePart();
    	  PackageRelationshipCollection wbrelcollection = wbpart.getRelationshipsByType("http://schemas.microsoft.com/office/2006/relationships/vbaProject");
    	  for (PackageRelationship relship : wbrelcollection) {
    	   wbpart.removeRelationship(relship.getId());
    	  }

    	  // コンテンツタイプをXLSXに設定
    	  workbook.setWorkbookType(XSSFWorkbookType.XLSX);
    	  workbook.write(new FileOutputStream(xlsxFileBase));
    	  workbook.close();
    }
}
