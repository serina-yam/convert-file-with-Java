package com.example;
import org.springframework.web.multipart.MultipartFile;

import com.itextpdf.text.pdf.PdfWriter;

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
import java.util.ArrayList;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;

import java.util.List;

@RestController
public class FileConversionItextController {

    @PostMapping("/convertXLSToPDF1")
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
        FileInputStream fileInputStream = new FileInputStream(xlsxFile);

        try (Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            Document document = new Document();
            try {
                PdfWriter.getInstance(document, new FileOutputStream(pdfFile));
            } catch (DocumentException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            document.open();

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);

                for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    List<String> rowList = getRow(row);
                    try {
                        addPDFText(document, rowList);
                    } catch (DocumentException e) {
                        // TODO Auto-generated catch block
                        e.printStackTrace();
                    }
                }
            }

            document.close();
        }
    }

   private static List<String> getRow(Row row) {
        List<String> list = new ArrayList<>();

        for (Cell cell : row) {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case STRING:
                    list.add(cell.getStringCellValue());
                    break;
                case NUMERIC:
                    list.add(String.valueOf(cell.getNumericCellValue()));
                    break;
                case BOOLEAN:
                    list.add(String.valueOf(cell.getBooleanCellValue()));
                    break;
                case FORMULA:
                    list.add(cell.getCellFormula().toString());
                    break;
                case BLANK:
                case ERROR:
                case _NONE:
                    list.add("");
                    break;
            }
        }

        return list;
    }

    private static void addPDFText(Document document, List<String> textList) throws DocumentException {
        for (String text : textList) {
            document.add(new Paragraph(text));
        }
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
