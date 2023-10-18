package com.example.controller;

import java.io.File;
import java.io.IOException;

import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.Status;
import com.example.models.ApiResponse;

/**
 * JODConverterを使ってxlsxファイルをpdfファイルに変換する.
 */
@RestController
public class ExcelToPdfConverter {

        /**
         * Main.
         * @param args
        * @throws OfficeException
        * @throws IOException
        */
        @PostMapping("/convertXlsxToPdf")
        public ApiResponse excute(@RequestParam("xlsxFile") MultipartFile xlsxFileMulti) {

                String officeHome = "/Applications/LibreOffice.app/Contents"; // LibreOfficeのインストールパスを指定
                String pdfFilePath = "/Users/user-name/Downloads/output.pdf"; // 出力したいpdfファイルのパスを指定

                OfficeManager officeManager = LocalOfficeManager.builder()
                        .officeHome(officeHome)
                        .build();

                try {
                        convertXlsxToPdf(officeManager, xlsxFileMulti, pdfFilePath);
                        return new ApiResponse(Status.OK.getValue(), "PDF変換処理は正常終了しました。", pdfFilePath);
                } catch (OfficeException | IOException e) {
                        e.printStackTrace();
                        return new ApiResponse(Status.NG.getValue(), "PDF変換処理は異常しました。", null);
                }
        }

        /**
         * xlsxファイルをpdfファイルに変換.
         * @param officeManager
         * @param xlsxFileMulti
         * @param pdfFilePath
         * @return
         * @throws OfficeException
         * @throws IOException
         */
        private ApiResponse convertXlsxToPdf(OfficeManager officeManager, MultipartFile xlsxFileMulti,
                        String pdfFilePath) throws OfficeException, IOException {

                DocumentConverter converter = LocalConverter.make(officeManager);

                try {
                        officeManager.start();

                        File excelFile = convertMultipartFileToFile(xlsxFileMulti, "/Users/yamamotoserina/Desktop");
                        File pdfFile = new File(pdfFilePath);
                        // 変換処理
                        converter.convert(excelFile).to(pdfFile).execute();
                } finally {
                        officeManager.stop();
                }

                return null;
        }

        /**
         * MultipartFileからFileへの変換.
         * @param multipartFile
         * @param uploadDirectory
         * @return
         * @throws IOException
         */
        public File convertMultipartFileToFile(MultipartFile multipartFile, String uploadDirectory) throws IOException {
                String originalFileName = multipartFile.getOriginalFilename();
                String filePath = uploadDirectory + File.separator + originalFileName;

                File file = new File(filePath);

                // MultipartFileをFileにコピー
                multipartFile.transferTo(file);

                return file;
        }
}
