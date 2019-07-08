package com.dankirberger.demo.poiexcelimageissue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelGenerator {
    private ClientAnchor.AnchorType imageAnchorType;

    public ExcelGenerator(ClientAnchor.AnchorType imageAnchorType) {
        this.imageAnchorType = imageAnchorType;
    }

    public void generateExampleWorksheet() {
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Example");
        writeHeader(sheet);
        addFirstRow(workbook, sheet);
        addSecondRow(sheet);
        addThirdRow(workbook, sheet);

        sheet.setAutoFilter(new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(), 0, 2));

        write(workbook);
    }

    private void writeHeader(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        Cell imageCell = headerRow.createCell(0);
        imageCell.setCellValue("Image");
        Cell nameHeader = headerRow.createCell(1);
        nameHeader.setCellValue("Text");
        Cell typeHeader = headerRow.createCell(2);
        typeHeader.setCellValue("Type");
    }

    private void addFirstRow(Workbook workbook, Sheet sheet) {
        Row row = sheet.createRow(1);
        Cell image = row.createCell(0);
        writeImage(workbook, image, "one.png");
        Cell name = row.createCell(1);
        name.setCellValue("One");
        Cell type = row.createCell(2);
        type.setCellValue("Picture and Text");
    }

    private void addSecondRow(Sheet sheet) {
        Row row = sheet.createRow(2);
        Cell name = row.createCell(1);
        name.setCellValue("Two");
        Cell type = row.createCell(2);
        type.setCellValue("Text only");
    }

    private void addThirdRow(Workbook workbook, Sheet sheet) {
        Row row = sheet.createRow(3);
        Cell image = row.createCell(0);
        writeImage(workbook, image, "three.png");
        Cell name = row.createCell(1);
        name.setCellValue("Three");
        Cell type = row.createCell(2);
        type.setCellValue("Picture and Text");
    }

    private void writeImage(Workbook workbook, Cell cell, String fileName) {

        byte[] imageBytes;
        try {
            imageBytes = IOUtils.toByteArray(this.getClass().getClassLoader().getResourceAsStream(fileName));
        } catch (IOException e) {
            throw new RuntimeException("Failed to load image", e);
        }
        int pictureId = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        Drawing drawing = cell.getSheet().createDrawingPatriarch();
        XSSFClientAnchor anchor = new XSSFClientAnchor();
        anchor.setAnchorType(imageAnchorType);
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());
        Picture picture = drawing.createPicture(anchor, pictureId);
        picture.resize(1, 1);
    }

    private void write(Workbook workbook) {
        String destination = System.getProperty("user.dir") + "/example-" + imageAnchorType + ".xlsx";
        try {
            workbook.write(new FileOutputStream(destination));
            System.out.println("Wrote file to " + destination);
        } catch (IOException e) {
            throw new RuntimeException("Failed to write workbook to " + destination, e);
        }

    }
}
