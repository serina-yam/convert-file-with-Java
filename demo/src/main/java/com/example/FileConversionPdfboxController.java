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
import java.util.regex.Pattern;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;

@RestController
public class FileConversionPdfboxController {

    @PostMapping("/convertXLSToPDFWithPdfbox")
    public ResponseEntity<String> convertXLSToPDF(@RequestParam("xlsmFiles") MultipartFile xlsmFileMulti) {
        try {
            // 一時ディレクトリを作成
            File tempDir = new File("tempDir");
            tempDir.mkdirs();

            File xlsxFile = new File(tempDir, "temp.xlsx");

            // xlsmファイルからマクロを削除し、xlsxに変換
            File xlsmFile = new File (xlsmFileMulti.getOriginalFilename());
            removeMacrosAndConvertToXLSX(xlsmFile, xlsxFile);

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
     * xlsmファイルからマクロの情報を削除する.
     * @param xlsmFile
     * @param xlsxFile
     * @throws IOException
     * @throws InvalidFormatException 
     */
    private void removeMacrosAndConvertToXLSX(File xlsmFileBase, File xlsxFileBase) throws IOException, InvalidFormatException {
    	
    	  XSSFWorkbook workbook = (XSSFWorkbook)WorkbookFactory.create(new FileInputStream(xlsmFileBase));

    	  OPCPackage opcpackage = workbook.getPackage();

    	  // vbaProject.bin 部分を取得してパッケージから削除します
    	  PackagePart vbapart = opcpackage.getPartsByName(Pattern.compile("/xl/vbaProject.bin")).get(0);
    	  opcpackage.removePart(vbapart);

    	  // パッケージから削除された vbaProject.bin 部分との関係を取得および削除します
    	  PackagePart wbpart = workbook.getPackagePart();
    	  PackageRelationshipCollection wbrelcollection = wbpart.getRelationshipsByType("http://schemas.microsoft.com/office/2006/relationships/vbaProject");
    	  for (PackageRelationship relship : wbrelcollection) {
    	   wbpart.removeRelationship(relship.getId());
    	  }

    	  // コンテンツタイプをXLSXに設定します
    	  workbook.setWorkbookType(XSSFWorkbookType.XLSX);
    	  workbook.write(new FileOutputStream(xlsxFileBase));
    	  workbook.close();
    }

    /**
     * xlsxファイルからpdfファイルを作成.
     * @param xlsxFile
     * @param pdfFile
     * @throws IOException
     */
    private void createPDFFromXLSX(File xlsxFile, File pdfFile) throws IOException {

        // [PDFBOXを利用して日本語を出力したい。](https://teratail.com/questions/301141)
        // try (TrueTypeCollection collection = new TrueTypeCollection(xlsxFile)) {
        try (PDDocument document = new PDDocument()) {

            PDPage page = new PDPage();
            document.addPage(page);
            
            PDFont font = new PDType1Font(Standard14Fonts.FontName.HELVETICA);

            // コンテンツを作成し、テキストを出力
            try (PDPageContentStream content = new PDPageContentStream(document, page)) {
                content.beginText();
                content.setFont(font, 12);
                content.showText("HelloWorld");
                content.endText();
                font = null;
            }

            document.save(pdfFile);
            document.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }

        // FileInputStream inputStream = new FileInputStream(xlsxFile);
        // Workbook workbook = WorkbookFactory.create(inputStream);

        // PDDocument document = new PDDocument();
        // for (Sheet sheet : workbook) {
        //     PDPage page = new PDPage();
        //     document.addPage(page);

        //     PDPageContentStream contentStream = new PDPageContentStream(document, page);

        //     float margin = 50;
        //     float yPosition = page.getMediaBox().getHeight() - margin;
        //     float rowHeight = 20f;

        //     // PDFont font = PDType0Font.load(pdfDocument, new File("mplus-1p-regular.ttf"));

        //     for (Row row : sheet) {
        //         float xPosition = margin;
        //         for (Cell cell : row) {
        //             String value = cell.toString();
        //             contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA), 12);
        //             contentStream.beginText();
        //             contentStream.newLineAtOffset(xPosition, yPosition);
        //             contentStream.showText(value);
        //             contentStream.endText();
        //             xPosition += 100; // セルの幅に応じて調整
        //         }
        //         yPosition -= rowHeight;
        //     }

        //     contentStream.close();
        // }

        // document.save(pdfFile);
        // document.close();
        // inputStream.close();
    }




}
