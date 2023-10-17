package com.example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ExcelToPdfConverter {

        /**
         * Main.
         * @param args
        * @throws OfficeException
        * @throws IOException
        */
        @PostMapping("/convertXlsxToPdf")
        public void excute(@RequestParam("xlsxFile") MultipartFile xlsxFileMulti) throws OfficeException, IOException {
                String officeHome = "/Applications/LibreOffice.app"; // LibreOfficeのインストールパスを指定
                String pdfFilePath = "/Users/yamamotoserina/Downloads/output.pdf";


                OfficeManager officeManager = LocalOfficeManager.builder()
                        .officeHome(officeHome)
                        .build();
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
