package com.example;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;

@RestController
public class FileConversionPdfboxController {

    @PostMapping("/convertXLSToPDF")
    public ResponseEntity<String> convertXLSToPDF(@RequestParam("xlsmFiles") MultipartFile[] xlsmFiles) {
        try {
            // 一時ディレクトリを作成
            File tempDir = new File("tempDir");
            tempDir.mkdirs();

            File xlsxFile = new File(tempDir, "temp.xlsx");

            for (MultipartFile xlsmFile : xlsmFiles) {
                // xlsmファイルからマクロを削除し、xlsxに変換
                removeMacrosAndConvertToXLSX(xlsmFile, xlsxFile);
            }

            // xlsxファイルからPDFを生成
            File pdfFile = new File(tempDir, "output.pdf");
            createPDFFromXLSX(xlsxFile, pdfFile);

            // 生成されたPDFファイルを保存またはクライアントに返す

            return ResponseEntity.ok("PDFファイルが生成されました。");
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("エラーが発生しました: " + e.getMessage());
        }
    }

    /**
     * xlsxファイルからpdfファイルを作成.
     * @param xlsxFile
     * @param pdfFile
     * @throws IOException
     */
    private void createPDFFromXLSX(File xlsxFile, File pdfFile) throws IOException {
        FileInputStream inputStream = new FileInputStream(xlsxFile);
        Workbook workbook = WorkbookFactory.create(inputStream);

        PDDocument document = new PDDocument();
        for (Sheet sheet : workbook) {
            PDPage page = new PDPage();
            document.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            float margin = 50;
            float yPosition = page.getMediaBox().getHeight() - margin;
            float rowHeight = 20f;

            for (Row row : sheet) {
                float xPosition = margin;
                for (Cell cell : row) {
                    String value = cell.toString();
                    contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA), 12);
                    contentStream.beginText();
                    contentStream.newLineAtOffset(xPosition, yPosition);
                    contentStream.showText(value);
                    contentStream.endText();
                    xPosition += 100; // セルの幅に応じて調整
                }
                yPosition -= rowHeight;
            }

            contentStream.close();
        }

        document.save(pdfFile);
        document.close();
        inputStream.close();
    }

    /**
     * xlsmファイルからマクロの情報を削除する.
     * @param xlsmFile
     * @param xlsxFile
     * @throws IOException
     */
    private void removeMacrosAndConvertToXLSX(MultipartFile xlsmFile, File xlsxFile) throws IOException {
        InputStream inputStream = xlsmFile.getInputStream();
        Workbook workbook = WorkbookFactory.create(inputStream);

        if (workbook instanceof XSSFWorkbook) {

            String originalFileName = xlsmFile.getOriginalFilename();
            if (originalFileName != null && originalFileName.toLowerCase().endsWith(".xlsm")) {
                // マクロファイルの場合のみマクロデータを取り除く
                XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
                removeMacroSheets(xssfWorkbook);
                removeMacroRelations(xssfWorkbook);
            }

            workbook.write(new FileOutputStream(xlsxFile));
        }

        workbook.close();
        inputStream.close();
    }

    /**
     * マクロシートを削除する.
     * @param workbook
     */
    private void removeMacroSheets(XSSFWorkbook workbook) {
        // Remove sheets related to macros (e.g., VBA modules)
        int sheetCount = workbook.getNumberOfSheets();
        for (int i = sheetCount - 1; i >= 0; i--) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            if (sheet.getPackagePart().getContentType().equals("application/vnd.ms-excel.vbaProject.sheet")) {
                workbook.removeSheetAt(i);
            }
        }
    }
    
    /**
     * マクロ関係のデータを削除する.
     * @param workbook
     */
    private void removeMacroRelations(XSSFWorkbook workbook) {
        // Remove relations related to macros
        for (POIXMLDocumentPart part : workbook.getRelations()) {
            if (part.getPackagePart().getContentType().equals("application/vnd.ms-office.vbaProject")) {
                workbook.getRelations().remove(part);
                workbook.getPackage().removePart(part.getPackagePart());
            }
        }
    }


}
